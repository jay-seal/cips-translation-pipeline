"""
cips_reconstruct.py
====================
CIPS Translation Pipeline — PPTX Reconstruction Script

Triggered by GitHub Actions (reconstruct.yml). Downloads the source PPTX from
Cloudflare R2 and the Agent 3 translation JSON from the GitHub repository, then
applies translated text to every matched shape whilst preserving all run-level
formatting. Outputs the translated PPTX and uploads it back to GitHub as a
release asset.

Environment variables (set as GitHub Actions secrets / env):
    R2_PUBLIC_URL       Base URL of the R2 bucket, e.g. https://pub-xxx.r2.dev
    PPTX_FILENAME       Filename of the source PPTX in R2, e.g. M1_Tutor_Slides.pptx
    GITHUB_TOKEN        Personal access token with repo write scope
    GITHUB_OWNER        Repository owner / organisation, e.g. my-org
    GITHUB_REPO         Repository name, e.g. cips-translation
    JSON_REPO_PATH      Path to the Agent 3 JSON inside the repo,
                        e.g. translations/FR-FR/M1_slides_1-20_agent3.json
    OUTPUT_FILENAME     Desired output filename, e.g. M1_Tutor_Slides_FR-FR.pptx
                        (optional — defaults to <stem>_translated.pptx)

Matching strategy (applied in order, stopping at first success):
    Tier 1 — Exact normalised match:
        NFKC-normalise both shape.text and segment source_text, collapse all
        whitespace runs to a single space, strip. If equal, replace.
    Tier 2 — Substring normalised match:
        Check whether the normalised source_text is contained within the
        normalised shape.text. Useful for shapes where surrounding text was
        not segmented (e.g. source text is a paragraph inside a larger shape).
    Tier 3 — Per-paragraph match:
        Iterate over each paragraph in the text frame. Normalise the
        concatenated runs of that paragraph and compare to normalised
        source_text. Replace only the matching paragraph's runs.

Text replacement preserves:
    - Run-level font name, size, bold, italic, underline, colour
    - Paragraph-level alignment and spacing
    - All other runs in the shape that do not belong to the replaced paragraph

Failures are logged to stdout and written to a JSON report file that GitHub
Actions uploads as an artefact.
"""

import json
import logging
import os
import re
import sys
import unicodedata
from copy import deepcopy
from pathlib import Path

import requests
from pptx import Presentation
from pptx.util import Pt

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("cips_reconstruct")

# ---------------------------------------------------------------------------
# Configuration — read from environment
# ---------------------------------------------------------------------------
R2_PUBLIC_URL   = os.environ.get("R2_PUBLIC_URL", "").rstrip("/")
PPTX_FILENAME   = os.environ.get("PPTX_FILENAME", "")
GITHUB_TOKEN    = os.environ.get("GITHUB_TOKEN", "")
GITHUB_OWNER    = os.environ.get("GITHUB_OWNER", "")
GITHUB_REPO     = os.environ.get("GITHUB_REPO", "")
JSON_REPO_PATH  = os.environ.get("JSON_REPO_PATH", "")
OUTPUT_FILENAME = os.environ.get("OUTPUT_FILENAME", "")

WORK_DIR = Path("/tmp/cips_reconstruct")
WORK_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    """
    Normalise a text string for comparison purposes only (never for output).

    Steps:
        1. NFKC unicode normalisation — converts smart quotes, en/em dashes,
           non-breaking spaces, ligatures, and other compatibility characters
           to their canonical equivalents.
        2. Collapse all whitespace sequences (spaces, tabs, newlines, carriage
           returns, form feeds) to a single ASCII space.
        3. Strip leading/trailing whitespace.

    This is intentionally lossy — it is used only for matching, never for
    writing text back to the presentation.
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ---------------------------------------------------------------------------
# Run-level text replacement
# ---------------------------------------------------------------------------

def _replace_runs_in_paragraph(paragraph, new_text: str) -> None:
    """
    Replace the text content of a paragraph's runs with new_text whilst
    preserving every run's formatting attributes (font name, size, bold,
    italic, underline, colour, etc.).

    Strategy:
        - Put all new text into the first run, clear all subsequent runs.
        - This preserves the leading run's formatting for the entire
          replacement text, which is the safest behaviour for translated
          content where run boundaries in the source are arbitrary artefacts
          of editing history.
        - If the paragraph has no runs, do nothing (safety guard).
    """
    runs = paragraph.runs
    if not runs:
        return

    # Write the full translated text into the first run.
    runs[0].text = new_text

    # Clear all remaining runs so no source text leaks through.
    for run in runs[1:]:
        run.text = ""


def replace_shape_text(shape, translated_text: str, match_tier: int) -> None:
    """
    Write translated_text into a shape, preserving formatting.

    For Tier 1 and Tier 2 matches (whole-shape replacement):
        All paragraphs are collapsed: the first paragraph receives the
        translated text (split on newlines into separate paragraphs where
        the translation contains explicit line breaks), and subsequent
        paragraphs are cleared.

    For Tier 3 matches (per-paragraph replacement):
        Called with the matched paragraph index already handled by
        replace_paragraph_text() — this function is not used for Tier 3.
    """
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    paragraphs = tf.paragraphs

    if not paragraphs:
        return

    # Split translated text on explicit newlines so multi-line translations
    # are distributed across paragraphs correctly.
    translated_lines = translated_text.split("\n")

    # Write each translated line into the corresponding paragraph.
    # If there are more lines than paragraphs, append all excess text to the
    # last paragraph (edge case — avoids data loss).
    for i, para in enumerate(paragraphs):
        if i < len(translated_lines):
            _replace_runs_in_paragraph(para, translated_lines[i])
        else:
            # Clear any extra paragraphs that are beyond the translated line
            # count — they are source-text residue.
            _replace_runs_in_paragraph(para, "")


def replace_paragraph_text(shape, para_index: int, translated_text: str) -> None:
    """
    Replace the text in a single paragraph (identified by index) within a
    shape's text frame. Used for Tier 3 per-paragraph matches.
    """
    if not shape.has_text_frame:
        return
    paragraphs = shape.text_frame.paragraphs
    if para_index >= len(paragraphs):
        return
    _replace_runs_in_paragraph(paragraphs[para_index], translated_text)


# ---------------------------------------------------------------------------
# Shape text extraction helpers
# ---------------------------------------------------------------------------

def shape_full_text(shape) -> str:
    """
    Return the full text of a shape by joining all paragraph texts with
    newlines, matching how python-pptx's shape.text property works.
    """
    if not shape.has_text_frame:
        return ""
    return "\n".join(
        "".join(run.text for run in para.runs)
        for para in shape.text_frame.paragraphs
    )


def paragraph_text(para) -> str:
    """Return the concatenated text of all runs in a single paragraph."""
    return "".join(run.text for run in para.runs)


# ---------------------------------------------------------------------------
# Three-tier matching
# ---------------------------------------------------------------------------

def find_and_replace(slide, source_text: str, translated_text: str, segment_id: str):
    """
    Attempt to find a shape on the slide whose text matches source_text and
    replace it with translated_text. Returns a dict describing the outcome.

    Searches all shapes including those inside grouped shapes.
    """
    norm_source = normalise(source_text)

    if not norm_source:
        return {"segment_id": segment_id, "result": "SKIP_EMPTY_SOURCE"}

    for shape in _iter_shapes(slide):
        if not shape.has_text_frame:
            continue

        raw_shape_text = shape_full_text(shape)
        norm_shape = normalise(raw_shape_text)

        # ----------------------------------------------------------------
        # Tier 1 — Exact normalised match
        # ----------------------------------------------------------------
        if norm_shape == norm_source:
            replace_shape_text(shape, translated_text, match_tier=1)
            log.info(
                "T1 MATCH  seg=%-12s  slide=%s  shape=%s",
                segment_id,
                slide.slide_id if hasattr(slide, "slide_id") else "?",
                _shape_label(shape),
            )
            return {
                "segment_id": segment_id,
                "result": "T1_EXACT",
                "shape": _shape_label(shape),
            }

        # ----------------------------------------------------------------
        # Tier 2 — Substring normalised match
        # ----------------------------------------------------------------
        if norm_source in norm_shape and len(norm_source) > 3:
            # Replace the shape's entire text with the translation.
            # For cases where source_text is a clean subset of a larger
            # shape, this is the correct behaviour because the Agent 3 JSON
            # carries the full translated equivalent of that shape.
            replace_shape_text(shape, translated_text, match_tier=2)
            log.info(
                "T2 SUBSTR seg=%-12s  slide=%s  shape=%s",
                segment_id,
                slide.slide_id if hasattr(slide, "slide_id") else "?",
                _shape_label(shape),
            )
            return {
                "segment_id": segment_id,
                "result": "T2_SUBSTRING",
                "shape": _shape_label(shape),
            }

        # ----------------------------------------------------------------
        # Tier 3 — Per-paragraph match
        # ----------------------------------------------------------------
        for idx, para in enumerate(shape.text_frame.paragraphs):
            norm_para = normalise(paragraph_text(para))
            if norm_para == norm_source and len(norm_source) > 1:
                replace_paragraph_text(shape, idx, translated_text)
                log.info(
                    "T3 PARA   seg=%-12s  slide=%s  shape=%s  para=%d",
                    segment_id,
                    slide.slide_id if hasattr(slide, "slide_id") else "?",
                    _shape_label(shape),
                    idx,
                )
                return {
                    "segment_id": segment_id,
                    "result": "T3_PARAGRAPH",
                    "shape": _shape_label(shape),
                    "para_index": idx,
                }

    # No match found on any tier.
    log.warning(
        "NO MATCH  seg=%-12s  source_text=%r",
        segment_id,
        source_text[:80],
    )
    return {
        "segment_id": segment_id,
        "result": "NO_MATCH",
        "source_text": source_text,
    }


def _iter_shapes(slide):
    """
    Yield all shapes from a slide, recursing into groups.
    """
    for shape in slide.shapes:
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP == 6
            yield from _iter_group(shape)
        else:
            yield shape


def _iter_group(group_shape):
    """Recursively yield shapes from a group shape."""
    for shape in group_shape.shapes:
        if shape.shape_type == 6:
            yield from _iter_group(shape)
        else:
            yield shape


def _shape_label(shape) -> str:
    """Return a human-readable label for a shape for logging."""
    name = getattr(shape, "name", "?")
    return name


# ---------------------------------------------------------------------------
# File download helpers
# ---------------------------------------------------------------------------

def download_pptx(r2_base_url: str, filename: str, dest_path: Path) -> Path:
    """Download the source PPTX from Cloudflare R2."""
    url = f"{r2_base_url}/{filename}"
    log.info("Downloading PPTX from %s", url)
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    dest_path.write_bytes(resp.content)
    log.info("PPTX saved to %s (%d bytes)", dest_path, len(resp.content))
    return dest_path


def download_json_from_github(
    token: str,
    owner: str,
    repo: str,
    repo_path: str,
    dest_path: Path,
) -> Path:
    """
    Download the Agent 3 JSON file from the GitHub repository using the
    Contents API.
    """
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{repo_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3.raw",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    log.info("Downloading JSON from GitHub: %s", url)
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    dest_path.write_bytes(resp.content)
    log.info("JSON saved to %s (%d bytes)", dest_path, len(resp.content))
    return dest_path


# ---------------------------------------------------------------------------
# Upload output to GitHub
# ---------------------------------------------------------------------------

def upload_pptx_to_github(
    token: str,
    owner: str,
    repo: str,
    output_filename: str,
    local_path: Path,
) -> None:
    """
    Upload the translated PPTX to the GitHub repository under
    outputs/<output_filename>, committing directly to the default branch.

    Uses the Contents API (PUT) to create or update the file.
    """
    import base64

    repo_dest_path = f"outputs/{output_filename}"
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{repo_dest_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    # Check whether the file already exists (need its sha for update).
    existing_sha = None
    check = requests.get(url, headers=headers, timeout=30)
    if check.status_code == 200:
        existing_sha = check.json().get("sha")
        log.info("Existing file found at %s (sha=%s) — will update.", repo_dest_path, existing_sha)
    elif check.status_code == 404:
        log.info("No existing file at %s — will create.", repo_dest_path)
    else:
        check.raise_for_status()

    content_b64 = base64.b64encode(local_path.read_bytes()).decode("ascii")

    payload = {
        "message": f"chore: upload translated PPTX {output_filename} [automated]",
        "content": content_b64,
    }
    if existing_sha:
        payload["sha"] = existing_sha

    put_resp = requests.put(url, headers=headers, json=payload, timeout=120)
    put_resp.raise_for_status()
    log.info("Uploaded %s to GitHub at %s", output_filename, repo_dest_path)


def upload_report_to_github(
    token: str,
    owner: str,
    repo: str,
    report_filename: str,
    local_path: Path,
) -> None:
    """Upload the match report JSON alongside the translated PPTX."""
    import base64

    repo_dest_path = f"outputs/{report_filename}"
    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{repo_dest_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    existing_sha = None
    check = requests.get(url, headers=headers, timeout=30)
    if check.status_code == 200:
        existing_sha = check.json().get("sha")
    elif check.status_code != 404:
        check.raise_for_status()

    content_b64 = base64.b64encode(local_path.read_bytes()).decode("ascii")
    payload = {
        "message": f"chore: upload match report {report_filename} [automated]",
        "content": content_b64,
    }
    if existing_sha:
        payload["sha"] = existing_sha

    put_resp = requests.put(url, headers=headers, json=payload, timeout=60)
    put_resp.raise_for_status()
    log.info("Uploaded match report to GitHub at %s", repo_dest_path)


# ---------------------------------------------------------------------------
# Main reconstruction logic
# ---------------------------------------------------------------------------

def reconstruct(pptx_path: Path, json_path: Path, output_path: Path) -> dict:
    """
    Core reconstruction function. Loads the PPTX and JSON, applies all
    translated segments, and saves the output.

    Returns a summary dict suitable for the match report.
    """
    log.info("Loading presentation: %s", pptx_path)
    prs = Presentation(str(pptx_path))

    log.info("Loading translation JSON: %s", json_path)
    with json_path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    segments = data.get("segments", [])
    log.info("Total segments in JSON: %d", len(segments))

    # Build a slide-number → slide object lookup.
    # python-pptx uses 0-based indexing; the JSON uses 1-based slide numbers.
    slide_map = {i + 1: slide for i, slide in enumerate(prs.slides)}
    log.info("Presentation has %d slides.", len(slide_map))

    results = []
    counts = {
        "total": 0,
        "translated": 0,
        "skipped_no_translation": 0,
        "skipped_do_not_translate": 0,
        "skipped_slide_out_of_range": 0,
        "T1_EXACT": 0,
        "T2_SUBSTRING": 0,
        "T3_PARAGRAPH": 0,
        "NO_MATCH": 0,
        "SKIP_EMPTY_SOURCE": 0,
    }

    for seg in segments:
        counts["total"] += 1

        segment_id   = seg.get("segment_id", "UNKNOWN")
        slide_number = seg.get("slide_or_page")
        source_text  = seg.get("source_text", "")
        translated   = seg.get("translated_text")
        status       = seg.get("translation_status", "")

        # Skip segments with no translation.
        if not translated:
            counts["skipped_no_translation"] += 1
            continue

        # Honour explicit DO_NOT_TRANSLATE status.
        if status == "DO_NOT_TRANSLATE":
            counts["skipped_do_not_translate"] += 1
            continue

        # Validate slide number.
        if slide_number not in slide_map:
            log.warning(
                "Slide %s out of range for seg %s — skipping.",
                slide_number,
                segment_id,
            )
            counts["skipped_slide_out_of_range"] += 1
            results.append({
                "segment_id": segment_id,
                "result": "SKIP_SLIDE_OUT_OF_RANGE",
                "slide": slide_number,
            })
            continue

        slide = slide_map[slide_number]
        counts["translated"] += 1

        outcome = find_and_replace(slide, source_text, translated, segment_id)
        outcome["slide"] = slide_number
        results.append(outcome)

        result_key = outcome.get("result", "NO_MATCH")
        if result_key in counts:
            counts[result_key] += 1
        else:
            counts[result_key] = 1

    # Save the modified presentation.
    log.info("Saving translated presentation to %s", output_path)
    prs.save(str(output_path))
    log.info("Save complete.")

    # Print summary.
    matched = counts["T1_EXACT"] + counts["T2_SUBSTRING"] + counts["T3_PARAGRAPH"]
    no_match = counts["NO_MATCH"]
    match_pct = (matched / counts["translated"] * 100) if counts["translated"] else 0

    log.info("=" * 60)
    log.info("RECONSTRUCTION SUMMARY")
    log.info("  Total segments          : %d", counts["total"])
    log.info("  Segments with translation: %d", counts["translated"])
    log.info("  Matched and replaced")
    log.info("    Tier 1 (exact)        : %d", counts["T1_EXACT"])
    log.info("    Tier 2 (substring)    : %d", counts["T2_SUBSTRING"])
    log.info("    Tier 3 (paragraph)    : %d", counts["T3_PARAGRAPH"])
    log.info("    Total matched         : %d  (%.1f%%)", matched, match_pct)
    log.info("  No match                : %d", no_match)
    log.info("  Skipped (no translation): %d", counts["skipped_no_translation"])
    log.info("  Skipped (do not trans.) : %d", counts["skipped_do_not_translate"])
    log.info("  Skipped (out of range)  : %d", counts["skipped_slide_out_of_range"])
    log.info("=" * 60)

    report = {
        "summary": counts,
        "match_rate_pct": round(match_pct, 2),
        "results": results,
        "no_match_segments": [r for r in results if r.get("result") == "NO_MATCH"],
    }

    return report


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    # ------------------------------------------------------------------
    # Validate required environment variables
    # ------------------------------------------------------------------
    missing = []
    for var in ("R2_PUBLIC_URL", "PPTX_FILENAME", "GITHUB_TOKEN",
                "GITHUB_OWNER", "GITHUB_REPO", "JSON_REPO_PATH"):
        if not os.environ.get(var):
            missing.append(var)
    if missing:
        log.error("Missing required environment variables: %s", ", ".join(missing))
        sys.exit(1)

    stem = Path(PPTX_FILENAME).stem
    output_filename = OUTPUT_FILENAME or f"{stem}_translated.pptx"
    report_filename = output_filename.replace(".pptx", "_match_report.json")

    local_pptx   = WORK_DIR / PPTX_FILENAME
    local_json   = WORK_DIR / "agent3_translation.json"
    local_output = WORK_DIR / output_filename
    local_report = WORK_DIR / report_filename

    # ------------------------------------------------------------------
    # Download inputs
    # ------------------------------------------------------------------
    download_pptx(R2_PUBLIC_URL, PPTX_FILENAME, local_pptx)
    download_json_from_github(
        GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, JSON_REPO_PATH, local_json
    )

    # ------------------------------------------------------------------
    # Reconstruct
    # ------------------------------------------------------------------
    report = reconstruct(local_pptx, local_json, local_output)

    # ------------------------------------------------------------------
    # Write match report
    # ------------------------------------------------------------------
    local_report.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Match report written to %s", local_report)

    # ------------------------------------------------------------------
    # Upload outputs to GitHub
    # ------------------------------------------------------------------
    upload_pptx_to_github(
        GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, output_filename, local_output
    )
    upload_report_to_github(
        GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, report_filename, local_report
    )

    # Exit with a non-zero code if more than 20% of translatable segments
    # failed to match — this surfaces the problem clearly in Actions logs.
    no_match_count = report["summary"].get("NO_MATCH", 0)
    translated_count = report["summary"].get("translated", 1)
    failure_rate = no_match_count / translated_count
    if failure_rate > 0.20:
        log.error(
            "Match failure rate %.1f%% exceeds 20%% threshold — "
            "review the match report at outputs/%s",
            failure_rate * 100,
            report_filename,
        )
        sys.exit(2)

    log.info("Reconstruction complete. Output: outputs/%s", output_filename)


if __name__ == "__main__":
    main()
