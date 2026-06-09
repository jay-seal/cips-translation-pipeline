"""
cips_reconstruct.py
====================
CIPS Translation Pipeline — PPTX Reconstruction Script

Triggered by GitHub Actions (reconstruct.yml). Accepts file paths via
command-line arguments. All file I/O is handled by the workflow.

Usage:
    python cips_reconstruct.py \
        --source  inputs/source.pptx \
        --translations  inputs/agent3_output.json \
        [--qa  inputs/agent4_output.json] \
        --output  outputs/<filename>.pptx \
        [--failure-threshold 0.30]

Matching strategy — applied in order per segment, stopping at first success:

    Tier 1  Exact normalised match on the full shape text.
    Tier 2  Substring normalised match (single-paragraph shapes only).
    Tier 3  Per-paragraph match — replaces only the matched paragraph.
    Tier 4  Layout shapes — Tiers 1–3 on slide.slide_layout.shapes.
    Tier 5  Master shapes — Tiers 1–3 on slide_layout.slide_master.shapes.

    Once a layout or master shape is translated, all subsequent slides that
    reference the same text return LAYOUT_ALREADY_TRANSLATED or
    MASTER_ALREADY_TRANSLATED rather than NO_MATCH. These are not counted
    against the failure rate.

Normalisation (comparison only — never applied to output text):
    1. NFKC unicode normalisation.
    2. Collapse all whitespace to a single ASCII space.
    3. Strip leading/trailing whitespace.
    4. Strip leading bullet/list characters (•, ❶–❿ etc.).

Field element handling:
    Paragraphs with <a:fld> field elements (slide numbers, dates) are handled
    specially during replacement. When a paragraph contains both runs and
    fields, the field is preserved and the trailing page-number literal is
    stripped from the translated text to avoid duplication.

Known permanent failures:
    Text embedded in raster images or SmartArt is not accessible to
    python-pptx. These segments will always produce NO_MATCH. In this deck
    approximately 22% of segments fall into this category. Set
    --failure-threshold accordingly (default 0.30).

Exit codes:
    0   Success.
    1   Missing or invalid input files.
    2   Match failure rate exceeds --failure-threshold.
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

# Pattern for stripping trailing page-number literals from footer translations
# before writing to a paragraph that still has a slide-number field element.
# Matches "| 3", "| 15" etc. at the end of a string.
_TRAILING_PAGE_RE = re.compile(r'\s*\|\s*\d+\s*$')


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    """
    Normalise for comparison only. Never applied to output text.
    1. NFKC unicode normalisation.
    2. Collapse whitespace to single space.
    3. Strip edges.
    4. Strip leading bullet/list characters.
    """
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    text = _LEADING_BULLET_RE.sub('', text)
    return text


def _source_key(norm_source: str) -> str:
    """
    Generate a deduplication key for layout/master already-translated tracking.
    Strips trailing page-number patterns so that footer segments for different
    slides map to the same key once the master is translated once.
    E.g. "cips.org | tagline | 3" and "cips.org | tagline | 15" both map to
    "cips.org | tagline".
    """
    return _TRAILING_PAGE_RE.sub('', norm_source).strip()


# ---------------------------------------------------------------------------
# Paragraph text helpers
# ---------------------------------------------------------------------------

def _para_full_text(para) -> str:
    """Return all text in a paragraph, including <a:fld> field element text."""
    return ''.join(elem.text or '' for elem in para._p.iter(_A_T))


def shape_full_text(shape) -> str:
    """Return full shape text, joining paragraphs with newlines."""
    if not shape.has_text_frame:
        return ""
    return "\n".join(_para_full_text(p) for p in shape.text_frame.paragraphs)


def paragraph_text(para) -> str:
    """Return full text of a single paragraph, including field elements."""
    return _para_full_text(para)


# ---------------------------------------------------------------------------
# Table cell proxy
# ---------------------------------------------------------------------------

class _TableCellProxy:
    """Wraps a table cell to expose the same interface as a Shape."""
    __slots__ = ('has_text_frame', 'text_frame', 'name')

    def __init__(self, cell, row_idx: int, col_idx: int, parent_name: str):
        self.has_text_frame = True
        self.text_frame = cell.text_frame
        self.name = f"{parent_name}[r{row_idx}c{col_idx}]"


# ---------------------------------------------------------------------------
# Shape iteration
# ---------------------------------------------------------------------------

def _iter_shapes(shape_collection):
    """Yield all shapes, recursing into groups and expanding table cells."""
    for shape in shape_collection:
        st = shape.shape_type
        if st == 6:    # GROUP
            yield from _iter_group(shape)
        elif st == 19:  # TABLE
            yield from _iter_table(shape)
        else:
            yield shape


def _iter_group(group_shape):
    for shape in group_shape.shapes:
        st = shape.shape_type
        if st == 6:
            yield from _iter_group(shape)
        elif st == 19:
            yield from _iter_table(shape)
        else:
            yield shape


def _iter_table(table_shape):
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
    Replace paragraph text whilst preserving run-level formatting.

    Field element handling:
    - Runs + fields: Keep the field (e.g. a slide-number field that
      auto-updates). Strip the trailing page-number literal from new_text
      so the field continues to provide it dynamically, avoiding duplication.
    - Runs only: Write new_text and remove stray fields.
    - Fields only: Insert a new plain run with new_text and remove fields
      (accepts a static value — appropriate for one-off translated decks).
    - Neither: Do nothing.
    """
    p_elem = paragraph._p
    runs = paragraph.runs
    field_elements = p_elem.findall(_A_FLD)

    if runs and field_elements:
        # Keep fields. Strip the page-number literal from the translation
        # so the field element continues to supply it dynamically.
        text_for_runs = _TRAILING_PAGE_RE.sub('', new_text)
        runs[0].text = text_for_runs
        for run in runs[1:]:
            run.text = ""
        # Fields are deliberately NOT removed here.
    elif runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
        for fld in field_elements:
            p_elem.remove(fld)
    elif field_elements:
        r_elem = etree.SubElement(p_elem, _A_R)
        t_elem = etree.SubElement(r_elem, _A_T)
        t_elem.text = new_text
        for fld in field_elements:
            p_elem.remove(fld)
    # else: no runs, no fields — nothing to replace.


def replace_shape_text(shape, translated_text: str) -> None:
    """Write translated_text into a shape, distributing lines across paragraphs."""
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
# Single-shape matching helper
# ---------------------------------------------------------------------------

def _match_shape(shape, norm_source: str, translated_text: str,
                 segment_id: str, slide_id, tier_prefix: str = ""):
    """
    Apply Tiers 1–3 to a single shape. Returns a result dict on match, else None.
    Tier 2 is restricted to single-paragraph shapes to prevent a bullet item
    from matching as a substring of a multi-bullet placeholder.
    """
    if not shape.has_text_frame:
        return None

    norm_shape = normalise(shape_full_text(shape))
    label = _shape_label(shape)

    # Tier 1 — Exact normalised match.
    if norm_shape == norm_source:
        replace_shape_text(shape, translated_text)
        log.info("%-24s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T1_EXACT", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T1_EXACT", "shape": label}

    # Tier 2 — Substring match (single-paragraph shapes only).
    non_empty_paras = [
        p for p in shape.text_frame.paragraphs
        if normalise(paragraph_text(p))
    ]
    if (len(norm_source) > 3
            and norm_source in norm_shape
            and len(non_empty_paras) == 1):
        replace_shape_text(shape, translated_text)
        log.info("%-24s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T2_SUBSTR", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T2_SUBSTRING", "shape": label}

    # Tier 3 — Per-paragraph match.
    for idx, para in enumerate(shape.text_frame.paragraphs):
        norm_para = normalise(paragraph_text(para))
        if norm_para == norm_source and len(norm_source) > 1:
            replace_paragraph_text(shape, idx, translated_text)
            log.info("%-24s seg=%-12s  slide=%s  shape=%s  para=%d",
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
    state: dict,
) -> dict:
    """
    Apply Tiers 1–5 to find and replace source_text on a slide.

    state keys:
        modified_layout_ids     set of lxml element IDs of translated layout paragraphs
        modified_master_ids     set of lxml element IDs of translated master paragraphs
        matched_layout_keys     set of _source_key() values translated via layout
        matched_master_keys     set of _source_key() values translated via master
    """
    norm_source = normalise(source_text)
    if not norm_source:
        return {"segment_id": segment_id, "result": "SKIP_EMPTY_SOURCE"}

    slide_id = getattr(slide, 'slide_id', '?')
    src_key = _source_key(norm_source)

    # --- Fast path: already handled by a previous layout/master translation ---
    if src_key in state['matched_master_keys']:
        return {"segment_id": segment_id, "result": "MASTER_ALREADY_TRANSLATED"}
    if src_key in state['matched_layout_keys']:
        return {"segment_id": segment_id, "result": "LAYOUT_ALREADY_TRANSLATED"}

    # --- Tiers 1–3: individual slide shapes ---
    for shape in _iter_shapes(slide.shapes):
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id)
        if result:
            result["segment_id"] = segment_id
            return result

    # --- Tier 4: slide layout shapes ---
    try:
        layout = slide.slide_layout
        layout_shapes = layout.shapes
    except AttributeError:
        layout_shapes = []

    for shape in _iter_shapes(layout_shapes):
        if not shape.has_text_frame:
            continue
        para_ids = {id(p._p) for p in shape.text_frame.paragraphs}
        if para_ids & state['modified_layout_ids']:
            continue
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id, tier_prefix="LAYOUT_")
        if result:
            for p in shape.text_frame.paragraphs:
                state['modified_layout_ids'].add(id(p._p))
            state['matched_layout_keys'].add(src_key)
            result["segment_id"] = segment_id
            result["note"] = "Layout shape — applies to all slides using this layout."
            return result

    # --- Tier 5: slide master shapes ---
    try:
        master_shapes = slide.slide_layout.slide_master.shapes
    except AttributeError:
        master_shapes = []

    for shape in _iter_shapes(master_shapes):
        if not shape.has_text_frame:
            continue
        para_ids = {id(p._p) for p in shape.text_frame.paragraphs}
        if para_ids & state['modified_master_ids']:
            continue
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id, tier_prefix="MASTER_")
        if result:
            for p in shape.text_frame.paragraphs:
                state['modified_master_ids'].add(id(p._p))
            state['matched_master_keys'].add(src_key)
            result["segment_id"] = segment_id
            result["note"] = "Master shape — applies to all slides in the presentation."
            return result

    log.warning("NO MATCH  seg=%-12s  source_text=%r",
                segment_id, source_text[:80])
    return {"segment_id": segment_id, "result": "NO_MATCH",
            "source_text": source_text}


# ---------------------------------------------------------------------------
# Core reconstruction
# ---------------------------------------------------------------------------

def reconstruct(pptx_path: Path, json_path: Path, output_path: Path) -> dict:
    """Load, translate, and save the presentation. Returns a match report dict."""
    log.info("Loading presentation: %s", pptx_path)
    prs = Presentation(str(pptx_path))

    log.info("Loading translation JSON: %s", json_path)
    with json_path.open(encoding="utf-8") as fh:
        data = json.load(fh)

    segments = data.get("segments", [])
    log.info("Total segments in JSON: %d", len(segments))

    slide_map = {i + 1: slide for i, slide in enumerate(prs.slides)}
    log.info("Presentation has %d slides.", len(slide_map))

    state = {
        'modified_layout_ids':  set(),
        'modified_master_ids':  set(),
        'matched_layout_keys':  set(),
        'matched_master_keys':  set(),
    }

    results = []
    counts = {
        "total": 0, "translated": 0,
        "skipped_no_translation": 0, "skipped_do_not_translate": 0,
        "skipped_slide_out_of_range": 0,
        "T1_EXACT": 0, "T2_SUBSTRING": 0, "T3_PARAGRAPH": 0,
        "LAYOUT_T1_EXACT": 0, "LAYOUT_T2_SUBSTRING": 0, "LAYOUT_T3_PARAGRAPH": 0,
        "MASTER_T1_EXACT": 0, "MASTER_T2_SUBSTRING": 0, "MASTER_T3_PARAGRAPH": 0,
        "LAYOUT_ALREADY_TRANSLATED": 0, "MASTER_ALREADY_TRANSLATED": 0,
        "NO_MATCH": 0, "SKIP_EMPTY_SOURCE": 0,
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
            slide, source_text, translated, segment_id, state
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

    # Matched = all successful tiers (slide, layout, master)
    slide_matched  = counts["T1_EXACT"] + counts["T2_SUBSTRING"] + counts["T3_PARAGRAPH"]
    layout_matched = (counts["LAYOUT_T1_EXACT"] + counts["LAYOUT_T2_SUBSTRING"]
                      + counts["LAYOUT_T3_PARAGRAPH"])
    master_matched = (counts["MASTER_T1_EXACT"] + counts["MASTER_T2_SUBSTRING"]
                      + counts["MASTER_T3_PARAGRAPH"])
    already_handled = (counts["LAYOUT_ALREADY_TRANSLATED"]
                       + counts["MASTER_ALREADY_TRANSLATED"])
    total_handled  = slide_matched + layout_matched + master_matched + already_handled
    no_match = counts["NO_MATCH"]
    handle_pct = (total_handled / counts["translated"] * 100) if counts["translated"] else 0
    fail_rate  = no_match / counts["translated"] if counts["translated"] else 0

    log.info("=" * 60)
    log.info("RECONSTRUCTION SUMMARY")
    log.info("  Total segments              : %d", counts["total"])
    log.info("  Segments with translation   : %d", counts["translated"])
    log.info("  Slide-level matches")
    log.info("    Tier 1 (exact)            : %d", counts["T1_EXACT"])
    log.info("    Tier 2 (substring)        : %d", counts["T2_SUBSTRING"])
    log.info("    Tier 3 (paragraph)        : %d", counts["T3_PARAGRAPH"])
    log.info("  Layout-level matches (Tier 4)")
    log.info("    T4 exact                  : %d", counts["LAYOUT_T1_EXACT"])
    log.info("    T4 substring              : %d", counts["LAYOUT_T2_SUBSTRING"])
    log.info("    T4 paragraph              : %d", counts["LAYOUT_T3_PARAGRAPH"])
    log.info("  Master-level matches (Tier 5)")
    log.info("    T5 exact                  : %d", counts["MASTER_T1_EXACT"])
    log.info("    T5 substring              : %d", counts["MASTER_T2_SUBSTRING"])
    log.info("    T5 paragraph              : %d", counts["MASTER_T3_PARAGRAPH"])
    log.info("  Already handled (deduped)")
    log.info("    Layout already translated : %d", counts["LAYOUT_ALREADY_TRANSLATED"])
    log.info("    Master already translated : %d", counts["MASTER_ALREADY_TRANSLATED"])
    log.info("  Total handled               : %d  (%.1f%%)", total_handled, handle_pct)
    log.info("  No match (investigate)      : %d  (%.1f%%)", no_match, fail_rate * 100)
    log.info("  Skipped (no translation)    : %d", counts["skipped_no_translation"])
    log.info("  Skipped (do not trans.)     : %d", counts["skipped_do_not_translate"])
    log.info("  Skipped (out of range)      : %d", counts["skipped_slide_out_of_range"])
    log.info("=" * 60)

    return {
        "summary": counts,
        "handle_rate_pct": round(handle_pct, 2),
        "no_match_rate_pct": round(fail_rate * 100, 2),
        "results": results,
        "no_match_segments": [r for r in results if r.get("result") == "NO_MATCH"],
        "known_inaccessible_note": (
            "NO_MATCH segments that persist after all five tiers are almost always "
            "text embedded in raster images or SmartArt, which python-pptx cannot "
            "access. These require manual intervention or slide regeneration."
        ),
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="CIPS Translation Pipeline — PPTX Reconstruction Script"
    )
    parser.add_argument("--source", required=True,
                        help="Path to the source PPTX file.")
    parser.add_argument("--translations", required=True,
                        help="Path to the Agent 3 translation JSON.")
    parser.add_argument("--qa", required=False, default=None,
                        help="Path to the Agent 4 QA JSON (optional).")
    parser.add_argument("--output", required=True,
                        help="Destination path for the translated PPTX.")
    parser.add_argument(
        "--failure-threshold", type=float, default=0.30,
        help=(
            "Exit with code 2 if the NO_MATCH rate exceeds this fraction. "
            "Default 0.30 (30%%). Set higher for decks with significant "
            "embedded-image content that is structurally inaccessible."
        ),
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
            log.warning("QA file specified but not found: %s — proceeding without it.",
                        qa_path)

    report = reconstruct(source_path, json_path, output_path)

    report_path = (output_path.parent
                   / output_path.name.replace(".pptx", "_match_report.json"))
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Match report written to %s", report_path)

    no_match_rate = report["summary"].get("NO_MATCH", 0) / max(
        report["summary"].get("translated", 1), 1
    )
    if no_match_rate > args.failure_threshold:
        log.error(
            "NO_MATCH rate %.1f%% exceeds threshold %.0f%% — "
            "review %s for permanently inaccessible segments.",
            no_match_rate * 100,
            args.failure_threshold * 100,
            report_path,
        )
        sys.exit(2)

    log.info("Reconstruction complete. Output: %s", output_path)


if __name__ == "__main__":
    main()
