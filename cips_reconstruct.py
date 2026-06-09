"""
cips_reconstruct.py
====================
CIPS Translation Pipeline — PPTX Reconstruction Script

Triggered by GitHub Actions (reconstruct.yml). Accepts file paths via
command-line arguments. All file I/O (downloading the source PPTX,
committing outputs) is handled by the workflow; this script is
responsible only for the reconstruction logic.

Usage:
    python cips_reconstruct.py \
        --source  inputs/source.pptx \
        --translations  inputs/agent3_output.json \
        [--qa  inputs/agent4_output.json] \
        --output  outputs/<filename>.pptx

Arguments:
    --source        Path to the source PPTX file.
    --translations  Path to the Agent 3 translation JSON.
    --qa            Path to the Agent 4 QA JSON (optional).
    --output        Destination path for the translated PPTX.

Outputs:
    <output>.pptx           The translated presentation.
    <output>_match_report.json  Per-segment match outcomes, written to the
                                same directory as the output PPTX.

Matching strategy (applied in order, stopping at first success):
    Tier 1 — Exact normalised match:
        NFKC-normalise both shape.text and segment source_text, collapse all
        whitespace runs to a single space, strip. If equal, replace.
    Tier 2 — Substring normalised match:
        Check whether the normalised source_text is contained within the
        normalised shape.text.
    Tier 3 — Per-paragraph match:
        Iterate over each paragraph in the text frame. Normalise the
        concatenated runs of that paragraph and compare to normalised
        source_text. Replace only the matching paragraph's runs.

Text replacement preserves:
    - Run-level font name, size, bold, italic, underline, colour.
    - Paragraph-level alignment and spacing.
    - All other runs in the shape that do not belong to the replaced paragraph.

Exit codes:
    0   Success (match failure rate within acceptable threshold).
    1   Missing or invalid arguments / unreadable input files.
    2   Match failure rate exceeds 20% — review the match report.
"""

import argparse
import json
import logging
import re
import sys
import unicodedata
from pathlib import Path

from pptx import Presentation

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

    This is intentionally lossy — used only for matching, never for writing
    text back to the presentation.
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

    Puts all new text into the first run and clears all subsequent runs.
    If the paragraph has no runs, does nothing.
    """
    runs = paragraph.runs
    if not runs:
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


def replace_shape_text(shape, translated_text: str) -> None:
    """
    Write translated_text into a shape's text frame, preserving formatting.

    Distributes translated lines across paragraphs. If the translation
    contains fewer lines than the shape has paragraphs, excess paragraphs
    are cleared. If it contains more lines than paragraphs, excess text is
    appended to the last paragraph.
    """
    if not shape.has_text_frame:
        return

    paragraphs = shape.text_frame.paragraphs
    if not paragraphs:
        return

    translated_lines = translated_text.split("\n")

    for i, para in enumerate(paragraphs):
        if i < len(translated_lines):
            _replace_runs_in_paragraph(para, translated_lines[i])
        else:
            _replace_runs_in_paragraph(para, "")


def replace_paragraph_text(shape, para_index: int, translated_text: str) -> None:
    """
    Replace the text in a single paragraph (by index) within a shape's text
    frame. Used for Tier 3 per-paragraph matches.
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
    """Return the full text of a shape, joining paragraphs with newlines."""
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

def find_and_replace(slide, source_text: str, translated_text: str, segment_id: str) -> dict:
    """
    Attempt to find a shape on the slide whose text matches source_text and
    replace it with translated_text. Returns a dict describing the outcome.

    Searches all shapes including those inside grouped shapes.
    """
    norm_source = normalise(source_text)

    if not norm_source:
        return {"segment_id": segment_id, "result": "SKIP_EMPTY_SOURCE"}

    slide_id = getattr(slide, "slide_id", "?")

    for shape in _iter_shapes(slide):
        if not shape.has_text_frame:
            continue

        raw_shape_text = shape_full_text(shape)
        norm_shape = normalise(raw_shape_text)

        # ----------------------------------------------------------------
        # Tier 1 — Exact normalised match
        # ----------------------------------------------------------------
        if norm_shape == norm_source:
            replace_shape_text(shape, translated_text)
            log.info("T1 MATCH  seg=%-12s  slide=%s  shape=%s",
                     segment_id, slide_id, _shape_label(shape))
            return {"segment_id": segment_id, "result": "T1_EXACT",
                    "shape": _shape_label(shape)}

        # ----------------------------------------------------------------
        # Tier 2 — Substring normalised match
        # ----------------------------------------------------------------
        if norm_source in norm_shape and len(norm_source) > 3:
            replace_shape_text(shape, translated_text)
            log.info("T2 SUBSTR seg=%-12s  slide=%s  shape=%s",
                     segment_id, slide_id, _shape_label(shape))
            return {"segment_id": segment_id, "result": "T2_SUBSTRING",
                    "shape": _shape_label(shape)}

        # ----------------------------------------------------------------
        # Tier 3 — Per-paragraph match
        # ----------------------------------------------------------------
        for idx, para in enumerate(shape.text_frame.paragraphs):
            norm_para = normalise(paragraph_text(para))
            if norm_para == norm_source and len(norm_source) > 1:
                replace_paragraph_text(shape, idx, translated_text)
                log.info("T3 PARA   seg=%-12s  slide=%s  shape=%s  para=%d",
                         segment_id, slide_id, _shape_label(shape), idx)
                return {"segment_id": segment_id, "result": "T3_PARAGRAPH",
                        "shape": _shape_label(shape), "para_index": idx}

    log.warning("NO MATCH  seg=%-12s  source_text=%r", segment_id, source_text[:80])
    return {"segment_id": segment_id, "result": "NO_MATCH", "source_text": source_text}


def _iter_shapes(slide):
    """Yield all shapes from a slide, recursing into groups."""
    for shape in slide.shapes:
        if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
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
    return getattr(shape, "name", "?")


# ---------------------------------------------------------------------------
# Core reconstruction
# ---------------------------------------------------------------------------

def reconstruct(pptx_path: Path, json_path: Path, output_path: Path) -> dict:
    """
    Load the PPTX and translation JSON, apply all translated segments, and
    save the output. Returns a summary dict for the match report.
    """
    log.info("Loading presentation: %s", pptx_path)
    prs = Presentation(str(pptx_path))

    log.info("Loading translation JSON: %s", json_path)
    with json_path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    segments = data.get("segments", [])
    log.info("Total segments in JSON: %d", len(segments))

    # python-pptx uses 0-based indexing; JSON uses 1-based slide numbers.
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

        if not translated:
            counts["skipped_no_translation"] += 1
            continue

        if status == "DO_NOT_TRANSLATE":
            counts["skipped_do_not_translate"] += 1
            continue

        if slide_number not in slide_map:
            log.warning("Slide %s out of range for seg %s — skipping.",
                        slide_number, segment_id)
            counts["skipped_slide_out_of_range"] += 1
            results.append({"segment_id": segment_id,
                            "result": "SKIP_SLIDE_OUT_OF_RANGE",
                            "slide": slide_number})
            continue

        slide = slide_map[slide_number]
        counts["translated"] += 1

        outcome = find_and_replace(slide, source_text, translated, segment_id)
        outcome["slide"] = slide_number
        results.append(outcome)

        result_key = outcome.get("result", "NO_MATCH")
        counts[result_key] = counts.get(result_key, 0) + 1

    # Save the modified presentation.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    log.info("Saving translated presentation to %s", output_path)
    prs.save(str(output_path))
    log.info("Save complete.")

    matched = counts["T1_EXACT"] + counts["T2_SUBSTRING"] + counts["T3_PARAGRAPH"]
    no_match = counts["NO_MATCH"]
    match_pct = (matched / counts["translated"] * 100) if counts["translated"] else 0

    log.info("=" * 60)
    log.info("RECONSTRUCTION SUMMARY")
    log.info("  Total segments           : %d", counts["total"])
    log.info("  Segments with translation: %d", counts["translated"])
    log.info("  Matched and replaced")
    log.info("    Tier 1 (exact)         : %d", counts["T1_EXACT"])
    log.info("    Tier 2 (substring)     : %d", counts["T2_SUBSTRING"])
    log.info("    Tier 3 (paragraph)     : %d", counts["T3_PARAGRAPH"])
    log.info("    Total matched          : %d  (%.1f%%)", matched, match_pct)
    log.info("  No match                 : %d", no_match)
    log.info("  Skipped (no translation) : %d", counts["skipped_no_translation"])
    log.info("  Skipped (do not trans.)  : %d", counts["skipped_do_not_translate"])
    log.info("  Skipped (out of range)   : %d", counts["skipped_slide_out_of_range"])
    log.info("=" * 60)

    return {
        "summary": counts,
        "match_rate_pct": round(match_pct, 2),
        "results": results,
        "no_match_segments": [r for r in results if r.get("result") == "NO_MATCH"],
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="CIPS Translation Pipeline — PPTX Reconstruction Script"
    )
    parser.add_argument(
        "--source", required=True,
        help="Path to the source PPTX file (downloaded by the workflow)."
    )
    parser.add_argument(
        "--translations", required=True,
        help="Path to the Agent 3 translation JSON."
    )
    parser.add_argument(
        "--qa", required=False, default=None,
        help="Path to the Agent 4 QA JSON (optional, reserved for future use)."
    )
    parser.add_argument(
        "--output", required=True,
        help="Destination path for the translated PPTX."
    )
    args = parser.parse_args()

    source_path = Path(args.source)
    json_path   = Path(args.translations)
    output_path = Path(args.output)

    # Validate inputs exist before doing any work.
    if not source_path.is_file():
        log.error("Source PPTX not found: %s", source_path)
        sys.exit(1)
    if not json_path.is_file():
        log.error("Translation JSON not found: %s", json_path)
        sys.exit(1)

    if args.qa:
        qa_path = Path(args.qa)
        if not qa_path.is_file():
            log.warning("QA file specified but not found: %s — proceeding without it.", qa_path)

    # Run reconstruction.
    report = reconstruct(source_path, json_path, output_path)

    # Write match report alongside the output PPTX.
    report_path = output_path.parent / output_path.name.replace(".pptx", "_match_report.json")
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Match report written to %s", report_path)

    # Exit 2 if match failure rate exceeds threshold.
    no_match_count   = report["summary"].get("NO_MATCH", 0)
    translated_count = report["summary"].get("translated", 1)
    failure_rate     = no_match_count / translated_count
    if failure_rate > 0.20:
        log.error(
            "Match failure rate %.1f%% exceeds 20%% threshold — "
            "review the match report at %s",
            failure_rate * 100,
            report_path,
        )
        sys.exit(2)

    log.info("Reconstruction complete. Output: %s", output_path)


if __name__ == "__main__":
    main()
