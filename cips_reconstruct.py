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

Outputs written to the same directory as --output:
    <name>.pptx                     Translated presentation.
    <name>_match_report.json        Full per-segment match results (machine-readable).
    <name>_manual_corrections.html  Human-readable list of unmatched segments for
                                    manual correction in PowerPoint.

Matching strategy (applied in order per segment, stopping at first success):

    Tier 1   Exact normalised match on the full shape text.
    Tier 1b  Page-number-tolerant match.
    Tier 2   Substring normalised match (single-paragraph shapes only).
    Tier 3   Per-paragraph match.
    Tier 3b  Contiguous paragraph range match — for multi-line segments.
    Tier 4   Layout shapes — Tiers 1-3b on slide.slide_layout.shapes.
    Tier 5   Master shapes — Tiers 1-3b on slide_layout.slide_master.shapes.
    Notes    Speaker notes — Tiers 1, 3, 3b on slide.notes_slide.

Normalisation (comparison only, never applied to output):
    NFKC unicode, smart quote flattening, whitespace collapse, bullet strip.

Newline decoding:
    LLM agents may double-escape newlines, storing the two-character sequence
    backslash+n instead of a real newline (U+000A). Both source_text and
    translated_text are decoded (backslash+n -> actual newline) before use.

Smart paragraph mapping (Tier 3b):
    When source_lines contains empty strings (spacer paragraphs), translated
    lines are mapped only to non-empty (content) paragraphs. Spacer paragraphs
    are left unchanged, preserving visual spacing and paragraph styles.
"""

import argparse
import html
import json
import logging
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree
from pptx import Presentation

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("cips_reconstruct")

_A_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
_A_T   = f'{{{_A_NS}}}t'
_A_R   = f'{{{_A_NS}}}r'
_A_FLD = f'{{{_A_NS}}}fld'

_LEADING_BULLET_RE = re.compile(
    r'[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25AA\u2776-\u277F\u2780-\u2789]+\s*'
)
_TRAILING_PAGE_RE = re.compile(r'\s*\|\s*\d+\s*$')

_ALWAYS_PROCESS_TYPES = frozenset({
    'slide_title', 'slide_subtitle', 'body_text', 'bullet_point',
    'heading', 'table_cell', 'text_box', 'caption', 'label',
})


# ---------------------------------------------------------------------------
# Normalisation
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace('\u2018', "'").replace('\u2019', "'")
    text = text.replace('\u201C', '"').replace('\u201D', '"')
    text = text.replace('\u2013', '-').replace('\u2014', '-')
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    text = _LEADING_BULLET_RE.sub('', text)
    return text


def _source_key(norm_source: str) -> str:
    return _TRAILING_PAGE_RE.sub('', norm_source).strip()


# ---------------------------------------------------------------------------
# Paragraph text helpers
# ---------------------------------------------------------------------------

def _para_full_text(para) -> str:
    return ''.join(elem.text or '' for elem in para._p.iter(_A_T))


def shape_full_text(shape) -> str:
    if not shape.has_text_frame:
        return ""
    return "\n".join(_para_full_text(p) for p in shape.text_frame.paragraphs)


def paragraph_text(para) -> str:
    return _para_full_text(para)


# ---------------------------------------------------------------------------
# Proxy classes
# ---------------------------------------------------------------------------

class _TableCellProxy:
    __slots__ = ('has_text_frame', 'text_frame', 'name')
    def __init__(self, cell, row_idx, col_idx, parent_name):
        self.has_text_frame = True
        self.text_frame = cell.text_frame
        self.name = f"{parent_name}[r{row_idx}c{col_idx}]"


class _NotesProxy:
    __slots__ = ('has_text_frame', 'text_frame', 'name')
    def __init__(self, text_frame, slide_id):
        self.has_text_frame = True
        self.text_frame = text_frame
        self.name = f"notes_slide_{slide_id}"


# ---------------------------------------------------------------------------
# Shape iteration
# ---------------------------------------------------------------------------

def _iter_shapes(shape_collection):
    for shape in shape_collection:
        st = shape.shape_type
        if st == 6:
            yield from _iter_group(shape)
        elif st == 19:
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
    All <a:fld> field elements are removed and text written as static runs.

    If the paragraph has no runs and no field elements (empty spacer
    paragraph), a new run is created to hold the translated text. This
    handles the case where translated content lands in a previously-empty
    paragraph. If new_text is empty, spacer paragraphs are left as-is.
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
        r_elem = etree.SubElement(p_elem, _A_R)
        t_elem = etree.SubElement(r_elem, _A_T)
        t_elem.text = new_text
        for fld in field_elements:
            p_elem.remove(fld)
    else:
        # No runs and no field elements. Only create a run if there is
        # actual text to write — spacer paragraphs with new_text="" stay empty.
        if new_text:
            r_elem = etree.SubElement(p_elem, _A_R)
            t_elem = etree.SubElement(r_elem, _A_T)
            t_elem.text = new_text


def replace_shape_text(shape, translated_text: str) -> None:
    """
    Write translated_text into a shape, distributing lines across paragraphs.

    When the shape contains a mix of content paragraphs and empty spacer
    paragraphs, translated lines are mapped only to the non-empty (content)
    paragraphs in order. Spacer paragraphs are left unchanged. This preserves
    visual spacing and paragraph styles when the PPTX uses empty paragraphs
    between content — a common pattern that would otherwise cause translated
    lines to land in the wrong paragraphs.

    When all paragraphs are non-empty (the common case), a direct
    line-to-paragraph mapping is used.
    """
    if not shape.has_text_frame:
        return
    paragraphs = shape.text_frame.paragraphs
    if not paragraphs:
        return
    translated_lines = translated_text.split("\n")

    para_texts = [_para_full_text(p).strip() for p in paragraphs]
    has_spacers = any(t == "" for t in para_texts) and any(t != "" for t in para_texts)

    if has_spacers:
        # Smart mapping: leave spacer paragraphs unchanged; distribute
        # translated lines only to non-empty (content) paragraphs in order.
        content_trans = iter(translated_lines)
        for i, para in enumerate(paragraphs):
            if not para_texts[i]:
                continue  # spacer — leave unchanged
            try:
                line = next(content_trans)
            except StopIteration:
                line = ""
            _replace_runs_in_paragraph(para, line)
    else:
        # No spacers — direct line-to-paragraph mapping.
        non_empty_paras = [p for p in paragraphs if _para_full_text(p).strip()]
        if len(non_empty_paras) > 1 and len(translated_lines) != len(non_empty_paras):
            log.warning(
                "LINE_COUNT_MISMATCH  shape=%-30s  "
                "shape_paragraphs=%d  translated_lines=%d",
                _shape_label(shape), len(non_empty_paras), len(translated_lines),
            )
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


def replace_paragraph_range(shape, start: int, count: int,
                             translated_text: str,
                             source_lines=None) -> None:
    """
    Replace count consecutive paragraphs starting at start with the lines
    from translated_text.

    When source_lines is provided (the normalised source lines from T3b
    matching) and contains empty strings, smart mapping is used: empty
    source lines are treated as structural spacer paragraphs and left
    unchanged. Non-empty source lines are mapped to translated content
    lines in order. This preserves paragraph styles and visual spacing
    when the PPTX shape contains empty paragraphs between content.

    Without source_lines (or when source has no empty lines), a direct
    line-to-paragraph mapping is used.
    """
    if not shape.has_text_frame:
        return
    paragraphs = shape.text_frame.paragraphs
    translated_lines = translated_text.split("\n")

    if source_lines and any(l == "" for l in source_lines):
        # Smart mapping: source has spacer paragraphs.
        # Map non-empty source lines to translated lines in order;
        # leave spacer paragraphs (empty source lines) unchanged.
        content_trans = iter(translated_lines)
        for i in range(count):
            idx = start + i
            if idx >= len(paragraphs):
                break
            src_line = source_lines[i] if i < len(source_lines) else ""
            if not src_line:
                # Spacer paragraph — leave unchanged.
                continue
            try:
                line = next(content_trans)
            except StopIteration:
                line = ""
            _replace_runs_in_paragraph(paragraphs[idx], line)
    else:
        # No spacers: direct line-to-paragraph mapping.
        for i in range(count):
            idx = start + i
            if idx >= len(paragraphs):
                break
            line = translated_lines[i] if i < len(translated_lines) else ""
            _replace_runs_in_paragraph(paragraphs[idx], line)


# ---------------------------------------------------------------------------
# Single-shape matching
# ---------------------------------------------------------------------------

def _match_shape(shape, norm_source: str, translated_text: str,
                 segment_id: str, slide_id, tier_prefix: str = "",
                 source_text_raw: str = ""):
    """
    Apply Tiers 1, 1b, 2, 3, and 3b to a single shape.

    Tier 3b: contiguous paragraph range match for multi-line segments.
    source_text_raw (the original un-normalised source text) is used to
    split on newlines and normalise each line individually — normalise()
    collapses newlines to spaces so norm_source cannot be used for this.
    """
    if not shape.has_text_frame:
        return None

    norm_shape = normalise(shape_full_text(shape))
    label = _shape_label(shape)

    # Tier 1
    if norm_shape == norm_source:
        replace_shape_text(shape, translated_text)
        log.info("%-24s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T1_EXACT", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T1_EXACT", "shape": label}

    # Tier 1b
    src_key   = _source_key(norm_source)
    shape_key = _source_key(norm_shape)
    if src_key and src_key != norm_source and src_key == shape_key:
        replace_shape_text(shape, translated_text)
        log.info("%-24s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T1b_FUZZY", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T1b_FUZZY", "shape": label}

    # Tier 2 (single-paragraph shapes only)
    non_empty_paras = [p for p in shape.text_frame.paragraphs
                       if normalise(paragraph_text(p))]
    if (len(norm_source) > 3
            and norm_source in norm_shape
            and len(non_empty_paras) == 1):
        replace_shape_text(shape, translated_text)
        log.info("%-24s seg=%-12s  slide=%s  shape=%s",
                 f"{tier_prefix}T2_SUBSTR", segment_id, slide_id, label)
        return {"result": f"{tier_prefix}T2_SUBSTRING", "shape": label}

    # Tier 3 — per-paragraph (single-line source)
    for idx, para in enumerate(shape.text_frame.paragraphs):
        norm_para = normalise(paragraph_text(para))
        if norm_para == norm_source and len(norm_source) > 1:
            replace_paragraph_text(shape, idx, translated_text)
            log.info("%-24s seg=%-12s  slide=%s  shape=%s  para=%d",
                     f"{tier_prefix}T3_PARA", segment_id, slide_id, label, idx)
            return {"result": f"{tier_prefix}T3_PARAGRAPH", "shape": label,
                    "para_index": idx}

    # Tier 3b — contiguous paragraph range (multi-line source)
    # NOTE: norm_source has newlines collapsed to spaces, so we must split
    # source_text_raw and normalise each line individually.
    if source_text_raw and '\n' in source_text_raw:
        raw_lines = source_text_raw.split('\n')
        norm_lines = [normalise(l) for l in raw_lines]
        while norm_lines and not norm_lines[-1]:
            norm_lines.pop()
        n = len(norm_lines)
        if n > 1:
            paragraphs = shape.text_frame.paragraphs
            para_texts = [normalise(paragraph_text(p)) for p in paragraphs]
            for start in range(len(paragraphs) - n + 1):
                if para_texts[start:start + n] == norm_lines:
                    replace_paragraph_range(
                        shape, start, n, translated_text,
                        source_lines=norm_lines,
                    )
                    log.info(
                        "%-24s seg=%-12s  slide=%s  shape=%s  paras=%d-%d",
                        f"{tier_prefix}T3b_RANGE", segment_id, slide_id,
                        label, start, start + n - 1,
                    )
                    return {
                        "result": f"{tier_prefix}T3b_PARA_RANGE",
                        "shape": label,
                        "para_start": start,
                        "para_end": start + n - 1,
                    }
    return None


# ---------------------------------------------------------------------------
# Speaker notes matching
# ---------------------------------------------------------------------------

def _match_notes(slide, source_text: str, translated_text: str,
                 segment_id: str) -> dict:
    try:
        notes_tf = slide.notes_slide.notes_text_frame
    except AttributeError:
        return {"segment_id": segment_id, "result": "NO_MATCH",
                "source_text": source_text}

    slide_id = getattr(slide, 'slide_id', '?')
    proxy = _NotesProxy(notes_tf, slide_id)
    norm_source = normalise(source_text)

    result = _match_shape(proxy, norm_source, translated_text,
                          segment_id, slide_id, tier_prefix="NOTES_",
                          source_text_raw=source_text)
    if result:
        result["segment_id"] = segment_id
        return result

    log.warning("NO MATCH (notes)  seg=%-12s  source_text=%r",
                segment_id, source_text[:80])
    return {"segment_id": segment_id, "result": "NO_MATCH",
            "source_text": source_text}


# ---------------------------------------------------------------------------
# Main find-and-replace
# ---------------------------------------------------------------------------

def find_and_replace(slide, source_text: str, translated_text: str,
                     segment_id: str, state: dict) -> dict:
    norm_source = normalise(source_text)
    if not norm_source:
        return {"segment_id": segment_id, "result": "SKIP_EMPTY_SOURCE"}

    slide_id = getattr(slide, 'slide_id', '?')
    src_key = _source_key(norm_source)

    if src_key in state['matched_master_keys']:
        return {"segment_id": segment_id, "result": "MASTER_ALREADY_TRANSLATED"}
    if src_key in state['matched_layout_keys']:
        return {"segment_id": segment_id, "result": "LAYOUT_ALREADY_TRANSLATED"}

    # Tiers 1-3b: slide shapes
    for shape in _iter_shapes(slide.shapes):
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id,
                              source_text_raw=source_text)
        if result:
            result["segment_id"] = segment_id
            return result

    # Tier 4: layout shapes
    try:
        layout_shapes = slide.slide_layout.shapes
    except AttributeError:
        layout_shapes = []

    for shape in _iter_shapes(layout_shapes):
        if not shape.has_text_frame:
            continue
        para_ids = {id(p._p) for p in shape.text_frame.paragraphs}
        if para_ids & state['modified_layout_ids']:
            continue
        result = _match_shape(shape, norm_source, translated_text,
                              segment_id, slide_id, tier_prefix="LAYOUT_",
                              source_text_raw=source_text)
        if result:
            for p in shape.text_frame.paragraphs:
                state['modified_layout_ids'].add(id(p._p))
            state['matched_layout_keys'].add(src_key)
            result["segment_id"] = segment_id
            result["note"] = "Layout shape."
            return result

    # Tier 5: master shapes
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
                              segment_id, slide_id, tier_prefix="MASTER_",
                              source_text_raw=source_text)
        if result:
            for p in shape.text_frame.paragraphs:
                state['modified_master_ids'].add(id(p._p))
            state['matched_master_keys'].add(src_key)
            result["segment_id"] = segment_id
            result["note"] = "Master shape."
            return result

    log.warning("NO MATCH  seg=%-12s  source_text=%r",
                segment_id, source_text[:80])
    return {"segment_id": segment_id, "result": "NO_MATCH",
            "source_text": source_text}


# ---------------------------------------------------------------------------
# HTML corrections report
# ---------------------------------------------------------------------------

def _write_html_report(report: dict, output_path: Path,
                       source_filename: str) -> None:
    no_match = report.get("no_match_segments", [])
    summary  = report.get("summary", {})
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    rows_html = ""
    for item in no_match:
        slide_num  = item.get("slide", "—")
        seg_id     = item.get("segment_id", "—")
        src        = item.get("source_text", "")
        translated = item.get("translated_text", "")
        rows_html += f"""
        <tr>
          <td class="slide">{html.escape(str(slide_num))}</td>
          <td class="seg">{html.escape(seg_id)}</td>
          <td class="src">{html.escape(src).replace(chr(10), '<br>')}</td>
          <td class="tgt">{html.escape(translated).replace(chr(10), '<br>')}</td>
        </tr>"""

    total      = summary.get("translated", 0)
    no_match_n = summary.get("NO_MATCH", 0)
    handled_n  = total - no_match_n

    page = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Manual Corrections — {html.escape(source_filename)}</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 14px; color: #1a1a2e; background: #f4f6fb; padding: 32px 24px; }}
    header {{ max-width: 1100px; margin: 0 auto 28px; }}
    h1 {{ font-size: 22px; font-weight: 700; margin-bottom: 6px; }}
    .meta {{ color: #555; font-size: 13px; margin-bottom: 18px; }}
    .stats {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 28px; }}
    .stat {{ background: #fff; border: 1px solid #dde3f0; border-radius: 8px;
             padding: 14px 20px; min-width: 140px; }}
    .stat .label {{ font-size: 11px; text-transform: uppercase; letter-spacing: .06em;
                   color: #777; margin-bottom: 4px; }}
    .stat .value {{ font-size: 26px; font-weight: 700; }}
    .stat.warn .value {{ color: #c0392b; }}
    .stat.ok   .value {{ color: #27ae60; }}
    .notice {{ background: #fff8e1; border-left: 4px solid #f39c12; border-radius: 4px;
               padding: 12px 16px; margin-bottom: 24px; max-width: 1100px;
               margin-left: auto; margin-right: auto; font-size: 13px; line-height: 1.6; }}
    .table-wrap {{ max-width: 1100px; margin: 0 auto; overflow-x: auto; }}
    table {{ width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px;
             overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.08); }}
    thead th {{ background: #1a1a2e; color: #fff; font-size: 12px; text-transform: uppercase;
                letter-spacing: .06em; padding: 12px 14px; text-align: left; }}
    tbody tr:nth-child(odd)  {{ background: #fff; }}
    tbody tr:nth-child(even) {{ background: #f9fafc; }}
    tbody tr:hover {{ background: #eef2ff; }}
    td {{ padding: 10px 14px; vertical-align: top; line-height: 1.5;
         border-bottom: 1px solid #eee; }}
    td.slide {{ font-weight: 700; width: 60px; text-align: center; }}
    td.seg   {{ font-family: monospace; font-size: 12px; color: #555; white-space: nowrap; }}
    td.src   {{ color: #333; }}
    td.tgt   {{ color: #1a6b3a; font-weight: 500; }}
    .empty   {{ text-align: center; padding: 40px; color: #888; font-size: 16px; }}
  </style>
</head>
<body>
  <header>
    <h1>Manual Corrections Required</h1>
    <p class="meta">Source: {html.escape(source_filename)} &nbsp;·&nbsp; Generated: {ts}</p>
    <div class="stats">
      <div class="stat ok"><div class="label">Auto-applied</div><div class="value">{handled_n}</div></div>
      <div class="stat warn"><div class="label">Needs manual fix</div><div class="value">{no_match_n}</div></div>
      <div class="stat"><div class="label">Total segments</div><div class="value">{total}</div></div>
    </div>
    <div class="notice">
      <strong>How to use this file:</strong> Each row is a segment the pipeline could not
      apply automatically. The most common cause is text embedded in a raster image or
      SmartArt diagram. Open the translated PPTX, go to the slide shown, find the English
      text, and replace it with the French. Speaker notes are accessible via
      View &rarr; Notes.
    </div>
  </header>
  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>Slide</th><th>Segment</th>
          <th>English (source)</th><th>French (target)</th>
        </tr>
      </thead>
      <tbody>
        {"<tr><td colspan='4' class='empty'>No manual corrections required.</td></tr>" if not no_match else rows_html}
      </tbody>
    </table>
  </div>
</body>
</html>"""

    output_path.write_text(page, encoding="utf-8")
    log.info("Manual corrections report: %s  (%d items)", output_path, no_match_n)


# ---------------------------------------------------------------------------
# Core reconstruction
# ---------------------------------------------------------------------------

def reconstruct(pptx_path: Path, json_path: Path, output_path: Path) -> dict:
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
        'modified_layout_ids': set(),
        'modified_master_ids': set(),
        'matched_layout_keys': set(),
        'matched_master_keys': set(),
    }

    results = []
    counts = {
        "total": 0, "translated": 0,
        "skipped_no_translation": 0, "skipped_do_not_translate": 0,
        "skipped_slide_out_of_range": 0, "skipped_not_text_box": 0,
        "T1_EXACT": 0, "T1b_FUZZY": 0, "T2_SUBSTRING": 0,
        "T3_PARAGRAPH": 0, "T3b_PARA_RANGE": 0,
        "LAYOUT_T1_EXACT": 0, "LAYOUT_T1b_FUZZY": 0,
        "LAYOUT_T2_SUBSTRING": 0, "LAYOUT_T3_PARAGRAPH": 0,
        "LAYOUT_T3b_PARA_RANGE": 0,
        "MASTER_T1_EXACT": 0, "MASTER_T1b_FUZZY": 0,
        "MASTER_T2_SUBSTRING": 0, "MASTER_T3_PARAGRAPH": 0,
        "MASTER_T3b_PARA_RANGE": 0,
        "LAYOUT_ALREADY_TRANSLATED": 0, "MASTER_ALREADY_TRANSLATED": 0,
        "NOTES_T1_EXACT": 0, "NOTES_T3_PARAGRAPH": 0,
        "NOTES_T3b_PARA_RANGE": 0,
        "NO_MATCH": 0, "SKIP_EMPTY_SOURCE": 0,
    }

    for seg in segments:
        counts["total"] += 1
        segment_id   = seg.get("segment_id", "UNKNOWN")
        slide_number = seg.get("slide_or_page")
        source_text  = seg.get("source_text", "")
        translated   = seg.get("translated_text")
        status       = seg.get("translation_status", "")
        element_type = seg.get("element_type", "")

        # Decode literal \n (backslash+n) to actual newlines.
        # LLM agents may double-escape newlines, producing the two-character
        # sequence backslash+n instead of a real newline (U+000A).
        if source_text:
            source_text = source_text.replace('\\n', '\x0a')
        if translated:
            translated = translated.replace('\\n', '\x0a')

        if not translated:
            counts["skipped_no_translation"] += 1
            continue
        if status == "DO_NOT_TRANSLATE":
            counts["skipped_do_not_translate"] += 1
            continue

        if seg.get("is_in_text_box") is False:
            if element_type not in _ALWAYS_PROCESS_TYPES:
                log.debug("SKIP_NOT_TEXT_BOX  seg=%-12s  type=%-16s",
                          segment_id, element_type)
                counts["skipped_not_text_box"] += 1
                results.append({"segment_id": segment_id,
                                "result": "SKIP_NOT_TEXT_BOX",
                                "slide": slide_number})
                continue

        if slide_number not in slide_map:
            log.warning("Slide %s out of range for seg %s.", slide_number, segment_id)
            counts["skipped_slide_out_of_range"] += 1
            results.append({"segment_id": segment_id,
                            "result": "SKIP_SLIDE_OUT_OF_RANGE",
                            "slide": slide_number})
            continue

        slide = slide_map[slide_number]
        counts["translated"] += 1

        if element_type == 'speaker_note':
            outcome = _match_notes(slide, source_text, translated, segment_id)
        else:
            outcome = find_and_replace(slide, source_text, translated,
                                       segment_id, state)

        outcome["slide"] = slide_number
        if outcome.get("result") == "NO_MATCH":
            outcome["translated_text"] = translated
        results.append(outcome)
        counts[outcome.get("result", "NO_MATCH")] = (
            counts.get(outcome.get("result", "NO_MATCH"), 0) + 1
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    log.info("Saved: %s", output_path)

    slide_matched  = (counts["T1_EXACT"] + counts["T1b_FUZZY"]
                      + counts["T2_SUBSTRING"] + counts["T3_PARAGRAPH"]
                      + counts["T3b_PARA_RANGE"])
    layout_matched = (counts["LAYOUT_T1_EXACT"] + counts["LAYOUT_T1b_FUZZY"]
                      + counts["LAYOUT_T2_SUBSTRING"] + counts["LAYOUT_T3_PARAGRAPH"]
                      + counts["LAYOUT_T3b_PARA_RANGE"])
    master_matched = (counts["MASTER_T1_EXACT"] + counts["MASTER_T1b_FUZZY"]
                      + counts["MASTER_T2_SUBSTRING"] + counts["MASTER_T3_PARAGRAPH"]
                      + counts["MASTER_T3b_PARA_RANGE"])
    notes_matched  = (counts["NOTES_T1_EXACT"] + counts["NOTES_T3_PARAGRAPH"]
                      + counts["NOTES_T3b_PARA_RANGE"])
    already        = (counts["LAYOUT_ALREADY_TRANSLATED"]
                      + counts["MASTER_ALREADY_TRANSLATED"])
    total_handled  = slide_matched + layout_matched + master_matched + notes_matched + already
    no_match       = counts["NO_MATCH"]
    handle_pct     = total_handled / counts["translated"] * 100 if counts["translated"] else 0
    fail_rate      = no_match / counts["translated"] if counts["translated"] else 0

    log.info("=" * 60)
    log.info("RECONSTRUCTION SUMMARY")
    log.info("  Total segments   : %d", counts["total"])
    log.info("  Translated       : %d", counts["translated"])
    log.info("  T1_EXACT         : %d", counts["T1_EXACT"])
    log.info("  T3_PARAGRAPH     : %d", counts["T3_PARAGRAPH"])
    log.info("  T3b_PARA_RANGE   : %d", counts["T3b_PARA_RANGE"])
    log.info("  NOTES matches    : %d", notes_matched)
    log.info("  Layout/Master    : %d", layout_matched + master_matched + already)
    log.info("  Total handled    : %d  (%.1f%%)", total_handled, handle_pct)
    log.info("  NO_MATCH         : %d  (%.1f%%)", no_match, fail_rate * 100)
    log.info("  Skipped not_tbox : %d", counts["skipped_not_text_box"])
    log.info("=" * 60)

    return {
        "summary": counts,
        "handle_rate_pct": round(handle_pct, 2),
        "no_match_rate_pct": round(fail_rate * 100, 2),
        "results": results,
        "no_match_segments": [r for r in results if r.get("result") == "NO_MATCH"],
        "known_inaccessible_note": (
            "Persistent NO_MATCH segments are almost always text embedded in "
            "raster images or SmartArt, which python-pptx cannot access."
        ),
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source",       required=True)
    parser.add_argument("--translations", required=True)
    parser.add_argument("--qa",           required=False, default=None)
    parser.add_argument("--output",       required=True)
    parser.add_argument("--failure-threshold", type=float, default=0.30)
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

    report = reconstruct(source_path, json_path, output_path)

    report_path = output_path.parent / output_path.name.replace(".pptx", "_match_report.json")
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("Match report: %s", report_path)

    html_path = output_path.parent / output_path.name.replace(".pptx", "_manual_corrections.html")
    _write_html_report(report, html_path, source_path.name)

    no_match_rate = (report["summary"].get("NO_MATCH", 0)
                     / max(report["summary"].get("translated", 1), 1))
    if no_match_rate > args.failure_threshold:
        log.error("NO_MATCH rate %.1f%% exceeds threshold %.0f%%",
                  no_match_rate * 100, args.failure_threshold * 100)
        sys.exit(2)

    log.info("Done. Output: %s", output_path)


if __name__ == "__main__":
    main()
