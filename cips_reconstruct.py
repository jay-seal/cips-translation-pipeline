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

Matching strategy — applied in order per segment, stopping at first success:

    Tier 1  Exact normalised match on the full shape text.
    Tier 2  Substring normalised match on the full shape text.
            Restricted to single-paragraph shapes to prevent a bullet
            item's source_text matching as a substring of a multi-bullet
            placeholder and incorrectly replacing the entire shape.
    Tier 3  Per-paragraph match — iterates every paragraph in the text
            frame and matches the normalised paragraph text against the
            normalised source_text. Replaces only the matched paragraph.
    Tier 4  Layout shape match — repeats Tiers 1–3 on the shapes defined
            in the slide's layout. Modifying a layout shape propagates the
            change to all slides using that layout, which is the correct
            behaviour for consistent template elements (headers, labels,
            footers). Already-translated layout paragraphs are skipped to
            avoid redundant processing across slides.

Normalisation (used for comparison only, never for output):
    1. NFKC unicode normalisation.
    2. Collapse all whitespace to a single ASCII space.
    3. Strip leading/trailing whitespace.
    4. Strip leading bullet/list characters (•, ❶–❿ etc.) — these are
       applied via paragraph formatting in PPTX and absent from run text,
       but Agent 1 may include them in source_text.

Text replacement:
    - Writes translated text into the first run of the matched paragraph
      and clears all subsequent runs, preserving run-level formatting.
    - Removes field elements (<a:fld>, used for auto-updating slide numbers
      and dates) from replaced paragraphs, since the translated text
      already contains the field's semantic content as a literal.
    - Handles paragraphs with no runs but with field elements by inserting
      a new plain run.

Shape traversal:
    - Group shapes are traversed recursively.
    - Table shapes are traversed cell by cell via _TableCellProxy.

Exit codes:
    0   Success (match failure rate within acceptable threshold).
    1   Missing or invalid input files.
    2   Match failure rate exceeds 20% — review the match report.
"""

import argparse
import json
import logging
import re
import sys
import unicodedata
from pathlib import Path

from lxml import etree
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
# XML namespace constants
# ---------------------------------------------------------------------------
_A_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_A_T   = f'{{{_A_NS}}}t'
_A_R   = f'{{{_A_NS}}}r'
_A_FLD = f'{{{_A_NS}}}fld'

# ---------------------------------------------------------------------------
# Bullet/list character stripping
# Characters rendered as list markers via paragraph formatting in PPTX.
# Agent 1 includes them in source_text as part of the rendered
# representation, but they do not appear in <a:r> run text.
# ---------------------------------------------------------------------------
_LEADING_BULLET_RE = re.compile(
    '['
    '\u2022'        # BULLET •
    '\u2023'        # TRIANGULAR BULLET
    '\u25E6'        # WHITE BULLET
    '\u2043'        # HYPHEN BULLET
    '\u2219'        # BULLET OPERATOR
    '\u25CF'        # BLACK CIRCLE
    '\u25AA'        # BLACK SMALL SQUARE
    '\u2776-\u277F' # DINGBAT NEGATIVE CIRCLED DIGITS ❶-❿
    '\u2780-\u2789' # DINGBAT CIRCLED SANS-SERIF DIGITS
    ']+' + r'\s*'
)


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    """
    Normalise a text string for comparison only (never used for output).

    1. NFKC unicode normalisation — converts smart quotes, en/em dashes,
       non-breaking spaces, and other compatibility characters.
    2. Collapse all whitespace sequences to a single ASCII space.
    3. Strip leading/trailing whitespace.
    4. Strip leading bullet/list characters.
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    text = _LEADING_BULLET_RE.sub('', text)
    return text


# ---------------------------------------------------------------------------
# Paragraph text helpers
# ---------------------------------------------------------------------------

def _para_full_text(para) -> str:
    """
    Return all text in a paragraph element, including text cached inside
    field elements (<a:fld>), which python-pptx's para.text omits.

    Footer shapes in PPTX commonly store the page number as an <a:fld>
    element. Without this function, shape_full_text() would return the
    footer text without the page number, causing Tier 1 to miss it.
    """
    return ''.join(elem.text or '' for elem in para._p.iter(_A_T))


def shape_full_text(shape) -> str:
    """
    Return the full text of a shape by joining paragraphs with newlines,
    including text inside field elements.
    """
    if not shape.has_text_frame:
        return ""
    return "\n".join(_para_full_text(p) for p in shape.text_frame.paragraphs)


def paragraph_text(para) -> str:
    """Return the full text of a single paragraph, including field elements."""
    return _para_full_text(para)


# ---------------------------------------------------------------------------
# Table cell proxy
# ---------------------------------------------------------------------------

class _TableCellProxy:
    """
    Wraps a python-pptx table cell to expose the same interface as a Shape
    so that table cells are processed identically to regular text frames.
    """
    __slots__ = ('has_text_frame', 'text_frame', 'name')

    def __init__(self, cell, row_idx: int, col_idx: int, parent_name: str):
        self.has_text_frame = True
        self.text_frame = cell.text_frame
        self.name = f"{parent_name}[r{row_idx}c{col_idx}]"


# ---------------------------------------------------------------------------
# Shape iteration helpers
# ---------------------------------------------------------------------------

def _iter_shapes(shape_collection):
    """
    Yield all shapes from a collection, recursing into groups and
    iterating table cells via _TableCellProxy.
    """
    for shape in shape_collection:
        st = shape.shape_type
        if st == 6:    # MSO_SHAPE_TYPE.GROUP
            yield from _iter_group(shape)
        elif st == 19:  # MSO_SHAPE_TYPE.TABLE
            yield from _iter_table(shape)
        else:
            yield shape


def _iter_group(group_shape):
    """Recursively yield shapes from a group, handling nested groups and tables."""
    for shape in group_shape.shapes:
        st = shape.shape_type
        if st == 6:
            yield from _iter_group(shape)
        elif st == 19:
            yield from _iter_table(shape)
        else:
            yield shape


def _iter_table(table_shape):
    """Yield a _TableCellProxy for each cell in a table shape."""
    for ri, row in enumerate(table_shape.table.rows):
        for ci, cell in enumerate(row.cells):
            yield _TableCellProxy(cell, ri, ci, table_shape.name)


def _shape_label(shape) -> str:
    return getattr(shape, 'name', '?')


# ---------------------------------------------------------------------------
# Run-level text replacement
# ---------------------------------------------------------------------------

def _replace_runs_in_paragraph(paragraph, new_text: str) -> None:
    """
    Replace the text of a paragraph whilst preserving run-level formatting.

    - Writes new_text into the first run and clears all subsequent runs.
    - Removes any field elements (<a:fld>) from the paragraph. Field
      elements hold auto-updating values (slide numbers, dates); the
      translated text already contains the field's semantic content as a
      literal, so the field must be removed to prevent duplication.
    - If the paragraph has no runs but has field elements (field-only
      paragraph), inserts a new plain run with new_text and removes fields.
    - If the paragraph has neither runs nor fields, does nothing.
    """
    p_elem = paragraph._p
    runs = paragraph.runs
    field_elements = p_elem.findall(_A_FLD)

    if runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
        for fld in field_elements:
            p_elem.remove(fld)
    elif field_elements:
        # Field-only paragraph — insert a plain run then remove fields.
        r_elem = etree.SubElement(p_elem, _A_R)
        t_elem = etree.SubElement(r_elem, _A_T)
        t_elem.text = new_text
        for fld in field_elements:
            p_elem.remove(fld)
    # else: no runs, no fields — nothing to replace.


def replace_shape_text(shape, translated_text: str) -> None:
    """
    Write translated_text into a shape's text frame, distributing lines
    across paragraphs and preserving run-level formatting.
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
    """Replace a single paragraph by index. Used for Tier 3 matches."""
    if not shape.has_text_frame:
        return
    paragraphs = shape.text_frame.paragraphs
    if para_index >= len(paragraphs):
        return
    _replace_runs_in_paragraph(paragraphs[para_index], translated_text)


# ---------------------------------------------------------------------------
# Single-shape matching (shared by slide and layout traversal)
# ---------------------------------------------------------------------------

def _match_shape(shape, norm_source: str, translated_text: str,
                 segment_id: str, slide_id, tier_prefix: str = ""):
    """
    Apply Tiers 1–3 to a single shape. Returns a result dict on match,
    or None if no tier matched.

    tier_prefix is "" for slide shapes and "LAYOUT_" for layout shapes.
    """
    if not shape.has_text_frame:
        return None

    norm_shape = normalise(shape_full_text(shape))
    label = _shape_label(shape)

    # Tier 1 — Exact normalised match.
    if norm_shape == norm_source:
        replace_shape_text(shape, translated_text)
        log.info("%-16s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T1_EXACT", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T1_EXACT", "shape": label}

    # Tier 2 — Substring normalised match, single-paragraph shapes only.
    # Restricting to shapes with exactly one non-empty paragraph prevents
    # this tier from firing on multi-bullet content placeholders, which
    # would replace the entire shape with only one bullet's translation
    # and destroy all other bullet text.
    non_empty_paras = [
        p for p in shape.text_frame.paragraphs
        if normalise(paragraph_text(p))
    ]
    if (len(norm_source) > 3
            and norm_source in norm_shape
            and len(non_empty_paras) == 1):
        replace_shape_text(shape, translated_text)
        log.info("%-16s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T2_SUBSTR", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T2_SUBSTRING", "shape": label}

    # Tier 3 — Per-paragraph match.
    for idx, para in enumerate(shape.text_frame.paragraphs):
        norm_para = normalise(paragraph_text(para))
        if norm_para == norm_source and len(norm_source) > 1:
            replace_paragraph_text(shape, idx, translated_text)
            log.info("%-16s seg=%-12s  slide=%s  shape=%s  para=%d",
                     f"{tier_prefix}T3_PARA", segment_id, slide_id, label, idx)
            return {"result": f"{tier_prefix}T3_PARAGRAPH", "shape": label,
                    "para_index": idx}

    return None


# ---------------------------------------------------------------------------
# Main find-and-replace orchestration
# ---------------------------------------------------------------------------

def find_and_replace(
    slide,
    source_text: str,
    translated_text: str,
    segment_id: str,
    modified_layout_ids: set,
) -> dict:
    """
    Attempt to match and replace source_text on a slide. Returns a dict
    describing the outcome.

    Processing order:
        1. Tiers 1–3 on shapes defined on the individual slide.
        2. Tier 4 (Tiers 1–3 repeated) on shapes defined on the slide layout.
           Layout-level modifications affect all slides using that layout,
           which is correct for consistent template elements. Paragraphs
           already translated in a previous slide pass are skipped via
           modified_layout_ids to avoid redundant log noise.
    """
    norm_source = normalise(source_text)
    if not norm_source:
        return {"segment_id": segment_id, "result": "SKIP_EMPTY_SOURCE"}

    slide_id = getattr(slide, 'slide_id', '?')

    # --- Tiers 1–3: slide-level shapes ---
    for shape in _iter_shapes(slide.shapes):
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id, tier_prefix="")
        if result:
            result["segment_id"] = segment_id
            return result

    # --- Tier 4: layout-level shapes ---
    try:
        layout_shapes = slide.slide_layout.shapes
    except AttributeError:
        layout_shapes = []

    for shape in _iter_shapes(layout_shapes):
        if not shape.has_text_frame:
            continue
        # Skip layout paragraphs already translated in a previous slide pass.
        para_ids = {id(p._p) for p in shape.text_frame.paragraphs}
        if para_ids & modified_layout_ids:
            continue
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id, tier_prefix="LAYOUT_")
        if result:
            # Record the translated paragraph IDs so subsequent slides skip them.
            for p in shape.text_frame.paragraphs:
                modified_layout_ids.add(id(p._p))
            result["segment_id"] = segment_id
            result["layout_note"] = (
                "Layout shape — change applies to all slides using this layout."
            )
            return result

    log.warning("NO MATCH  seg=%-12s  source_text=%r",
                segment_id, source_text[:80])
    return {"segment_id": segment_id, "result": "NO_MATCH",
            "source_text": source_text}


# ---------------------------------------------------------------------------
# Core reconstruction
# ---------------------------------------------------------------------------

def reconstruct(pptx_path: Path, json_path: Path, output_path: Path) -> dict:
    """
    Load the PPTX and translation JSON, apply all translated segments,
    and save the output. Returns a summary dict for the match report.
    """
    log.info("Loading presentation: %s", pptx_path)
    prs = Presentation(str(pptx_path))

    log.info("Loading translation JSON: %s", json_path)
    with json_path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    segments = data.get("segments", [])
    log.info("Total segments in JSON: %d", len(segments))

    # python-pptx uses 0-based indexing; the JSON uses 1-based slide numbers.
    slide_map = {i + 1: slide for i, slide in enumerate(prs.slides)}
    log.info("Presentation has %d slides.", len(slide_map))

    # Tracks layout paragraph element IDs already translated, to avoid
    # double-processing the same layout shape across multiple slides.
    modified_layout_ids: set = set()

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
        "LAYOUT_T1_EXACT": 0,
        "LAYOUT_T2_SUBSTRING": 0,
        "LAYOUT_T3_PARAGRAPH": 0,
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

        outcome = find_and_replace(
            slide, source_text, translated, segment_id, modified_layout_ids
        )
        outcome["slide"] = slide_number
        results.append(outcome)
        counts[outcome.get("result", "NO_MATCH")] = (
            counts.get(outcome.get("result", "NO_MATCH"), 0) + 1
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    log.info("Saving translated presentation to %s", output_path)
    prs.save(str(output_path))
    log.info("Save complete.")

    slide_matched = (counts["T1_EXACT"] + counts["T2_SUBSTRING"]
                     + counts["T3_PARAGRAPH"])
    layout_matched = (counts["LAYOUT_T1_EXACT"] + counts["LAYOUT_T2_SUBSTRING"]
                      + counts["LAYOUT_T3_PARAGRAPH"])
    total_matched = slide_matched + layout_matched
    no_match = counts["NO_MATCH"]
    match_pct = (
        (total_matched / counts["translated"] * 100)
        if counts["translated"] else 0
    )

    log.info("=" * 60)
    log.info("RECONSTRUCTION SUMMARY")
    log.info("  Total segments            : %d", counts["total"])
    log.info("  Segments with translation : %d", counts["translated"])
    log.info("  Slide-level matches")
    log.info("    Tier 1 (exact)          : %d", counts["T1_EXACT"])
    log.info("    Tier 2 (substring)      : %d", counts["T2_SUBSTRING"])
    log.info("    Tier 3 (paragraph)      : %d", counts["T3_PARAGRAPH"])
    log.info("  Layout-level matches (Tier 4)")
    log.info("    Tier 4 exact            : %d", counts["LAYOUT_T1_EXACT"])
    log.info("    Tier 4 substring        : %d", counts["LAYOUT_T2_SUBSTRING"])
    log.info("    Tier 4 paragraph        : %d", counts["LAYOUT_T3_PARAGRAPH"])
    log.info("  Total matched             : %d  (%.1f%%)",
             total_matched, match_pct)
    log.info("  No match                  : %d", no_match)
    log.info("  Skipped (no translation)  : %d", counts["skipped_no_translation"])
    log.info("  Skipped (do not trans.)   : %d", counts["skipped_do_not_translate"])
    log.info("  Skipped (out of range)    : %d", counts["skipped_slide_out_of_range"])
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

    if not source_path.is_file():
        log.error("Source PPTX not found: %s", source_path)
        sys.exit(1)
    if not json_path.is_file():
        log.error("Translation JSON not found: %s", json_path)
        sys.exit(1)
    if args.qa:
        qa_path = Path(args.qa)
        if not qa_path.is_file():
            log.warning(
                "QA file specified but not found: %s — proceeding without it.",
                qa_path
            )

    report = reconstruct(source_path, json_path, output_path)

    report_path = (
        output_path.parent
        / output_path.name.replace(".pptx", "_match_report.json")
    )
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Match report written to %s", report_path)

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
