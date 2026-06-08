#!/usr/bin/env python3
"""
CIPS Translation Pipeline — PPTX Reconstruction Script
=======================================================
Applies translated text from the Agent 3 JSON output to the original PPTX,
producing a translated PPTX file with original formatting preserved.

Usage:
    python cips_reconstruct.py \\
        --source  "Original.pptx" \\
        --translations  "agent3_output.json" \\
        --qa  "agent4_output.json" \\
        --output  "Original_FR-FR.pptx"

Arguments:
    --source        Path to the original source PPTX file
    --translations  Path to the Agent 3 JSON output (translated segments)
    --qa            Path to the Agent 4 JSON output (QA issues) [optional]
    --output        Path for the translated output PPTX file
    --dry-run       Preview matches without writing output [optional flag]

Requirements:
    pip install python-pptx
"""

import argparse
import json
import sys
import unicodedata
import re
from pathlib import Path
from copy import deepcopy
from datetime import datetime

try:
    from pptx import Presentation
    from pptx.util import Pt
except ImportError:
    print("ERROR: python-pptx is not installed.")
    print("Run:  pip install python-pptx")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Text normalisation helpers
# ---------------------------------------------------------------------------

def normalise(text: str) -> str:
    """Normalise text for matching — strips whitespace, normalises line endings
    and unicode, collapses multiple spaces."""
    if not text:
        return ""
    text = unicodedata.normalize("NFC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r" +", " ", text)
    return text.strip()


def shape_full_text(shape) -> str:
    """Return the full text of a shape's text frame with paragraphs
    separated by newlines."""
    if not shape.has_text_frame:
        return ""
    return "\n".join(
        "".join(run.text for run in para.runs)
        for para in shape.text_frame.paragraphs
    )


# ---------------------------------------------------------------------------
# Format-preserving text replacement
# ---------------------------------------------------------------------------

def replace_in_shape(shape, source_text: str, translated_text: str) -> bool:
    """
    Replace text in a shape while preserving run-level formatting.

    Strategy:
    - If source and translated text have the same number of paragraphs,
      replace paragraph-by-paragraph, preserving run formatting in each.
    - If counts differ, consolidate into the existing paragraph structure,
      putting the full translated text into the first run of the first
      paragraph and clearing the rest. Formatting of that first run is kept.

    Returns True if replacement was made, False if no match found.
    """
    if not shape.has_text_frame:
        return False

    current = shape_full_text(shape)
    if normalise(current) != normalise(source_text):
        return False

    tf = shape.text_frame
    src_paras = source_text.split("\n")
    tgt_paras = translated_text.split("\n")
    pptx_paras = tf.paragraphs

    if len(tgt_paras) == len(pptx_paras):
        # Best case: replace para by para
        for para, new_text in zip(pptx_paras, tgt_paras):
            _replace_para(para, new_text)
    else:
        # Para count mismatch — consolidate into first paragraph.
        # Join with a space to avoid concatenation artefacts.
        full_text = " ".join(tgt_paras)
        _replace_para(pptx_paras[0], full_text)
        for para in pptx_paras[1:]:
            _replace_para(para, "")

    return True


def _replace_para(para, new_text: str):
    """Replace text in a single paragraph, preserving the first run's
    formatting and clearing subsequent runs."""
    runs = para.runs
    if not runs:
        return
    # Put all text in the first run, preserve its formatting
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""


def apply_font_size_reduction(shape, reduction_pt: float = 2.0, min_pt: float = 10.0):
    """Reduce font size of all runs in a shape by reduction_pt, down to min_pt."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None:
                current_pt = run.font.size.pt
                new_pt = max(min_pt, current_pt - reduction_pt)
                run.font.size = Pt(new_pt)


# ---------------------------------------------------------------------------
# QA correction helpers
# ---------------------------------------------------------------------------

def apply_capitalisation_fix(text: str) -> str:
    """Capitalise the first character of a string if it is lowercase."""
    if not text:
        return text
    return text[0].upper() + text[1:]


# ---------------------------------------------------------------------------
# Main reconstruction logic
# ---------------------------------------------------------------------------

def load_json(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_qa_corrections(qa_data: dict) -> dict:
    """
    Build a dict of segment_id → list of CRITICAL corrections from Agent 4.
    Only capitalisation and punctuation issues are auto-applied.
    """
    corrections = {}
    if not qa_data:
        return corrections
    for issue in qa_data.get("qa_issues", []):
        if issue.get("severity") == "CRITICAL" and issue.get("issue_type") in (
            "capitalisation", "punctuation"
        ):
            seg_id = issue["segment_id"]
            corrections.setdefault(seg_id, []).append(issue)
    return corrections


def reconstruct(
    source_path: str,
    translations_path: str,
    output_path: str,
    qa_path: str = None,
    dry_run: bool = False,
) -> dict:
    """
    Core reconstruction function.

    Returns a report dict with counts and unmatched segments.
    """
    print(f"\nCIPS Translation Pipeline — PPTX Reconstruction")
    print(f"{'='*52}")
    print(f"Source PPTX  : {source_path}")
    print(f"Translations : {translations_path}")
    print(f"QA file      : {qa_path or 'not provided'}")
    print(f"Output       : {output_path}")
    print(f"Dry run      : {dry_run}")
    print()

    # Load data
    translation_data = load_json(translations_path)
    qa_data = load_json(qa_path) if qa_path else None

    # Handle Agent 4 JSON structure — translations may be nested
    if "reviewed_segments_json" in translation_data:
        # This is Agent 4 output — extract the inner translation data
        translation_data = translation_data["reviewed_segments_json"]

    segments = translation_data.get("segments", [])
    metadata = translation_data.get("document_metadata", {})
    qa_corrections = build_qa_corrections(qa_data)

    # Check routing if QA data provided
    if qa_data:
        routing = qa_data.get("qa_summary", {}).get("routing", "")
        if routing == "SME_REVIEW_REQUIRED":
            print("ERROR: Routing status is SME_REVIEW_REQUIRED.")
            print("       Resolve all CRITICAL issues in the Review Canvas")
            print("       before running reconstruction.")
            sys.exit(1)

    # Load presentation
    prs = Presentation(source_path)
    total_slides = len(prs.slides)
    print(f"Presentation loaded: {total_slides} slides")
    print(f"Translation segments: {len(segments)}")
    print()

    # Counters
    applied = 0
    skipped_keep = 0
    skipped_image = 0
    unmatched = []
    expansion_flags = []
    qa_applied = []

    for seg in segments:
        seg_id = seg.get("segment_id", "?")
        status = seg.get("translation_status", "")
        source_text = seg.get("source_text", "")
        translated_text = seg.get("translated_text", "")
        slide_num = seg.get("slide_or_page", 0)
        reconstruction_note = seg.get("reconstruction_note", "") or ""
        expansion_note = seg.get("expansion_note") or ""
        element_type = seg.get("element_type", "")

        # Skip IMAGE_TEXT
        if "IMAGE_TEXT" in reconstruction_note:
            skipped_image += 1
            continue

        # Skip KEPT segments
        if status in ("KEPT", "KEPT — IMAGE TEXT"):
            skipped_keep += 1
            continue

        # Skip if no translation
        if not translated_text or translated_text == source_text:
            skipped_keep += 1
            continue

        # Apply QA CRITICAL corrections to translated_text before matching
        qa_pending = []
        if seg_id in qa_corrections:
            for correction in qa_corrections[seg_id]:
                issue_type = correction.get("issue_type")
                if issue_type == "capitalisation":
                    original = translated_text
                    translated_text = apply_capitalisation_fix(translated_text)
                    if translated_text != original:
                        qa_pending.append(
                            f"{seg_id} (slide {slide_num}): capitalisation corrected"
                        )

        # Get target slide (0-indexed)
        slide_index = slide_num - 1
        if slide_index < 0 or slide_index >= total_slides:
            unmatched.append(
                f"{seg_id} | slide {slide_num} | slide out of range"
            )
            continue

        slide = prs.slides[slide_index]

        # Try to find and replace text in shapes on this slide
        matched = False
        for shape in slide.shapes:
            if dry_run:
                if shape.has_text_frame:
                    current = shape_full_text(shape)
                    if normalise(current) == normalise(source_text):
                        print(
                            f"  MATCH  {seg_id} | slide {slide_num} | "
                            f"{element_type} | "
                            f"{repr(source_text[:40])} → "
                            f"{repr(translated_text[:40])}"
                        )
                        matched = True
                        break
            else:
                if replace_in_shape(shape, source_text, translated_text):
                    matched = True
                    applied += 1
                    qa_applied.extend(qa_pending)
                    if expansion_note and seg_id in (
                        issue["segment_id"]
                        for issue in (qa_data or {}).get("qa_issues", [])
                        if issue.get("severity") == "CRITICAL"
                        and issue.get("issue_type") == "expansion"
                    ):
                        apply_font_size_reduction(shape)
                        expansion_flags.append(
                            f"{seg_id} | slide {slide_num}: font reduced 2pt"
                        )
                    break

        if not matched:
            unmatched.append(
                f"{seg_id} | slide {slide_num} | {element_type} | "
                f"{repr(source_text[:50])}"
            )

    # Save output
    if not dry_run:
        prs.save(output_path)
        print(f"\nOutput saved: {output_path}")

    # Report
    report = {
        "source": source_path,
        "output": output_path,
        "locale": metadata.get("target_locale", "unknown"),
        "timestamp": datetime.now().isoformat(),
        "total_segments": len(segments),
        "applied": applied,
        "skipped_keep": skipped_keep,
        "skipped_image_text": skipped_image,
        "unmatched": unmatched,
        "qa_corrections_applied": qa_applied,
        "expansion_font_reductions": expansion_flags,
    }

    _print_report(report, dry_run)
    return report


def _print_report(report: dict, dry_run: bool):
    print()
    print("RECONSTRUCTION REPORT")
    print("=" * 52)
    if dry_run:
        print("MODE: DRY RUN — no file written")
    print(f"Source        : {report['source']}")
    print(f"Output        : {report['output']}")
    print(f"Locale        : {report['locale']}")
    print(f"Timestamp     : {report['timestamp']}")
    print()
    print(f"Total segments     : {report['total_segments']}")
    print(f"Replacements made  : {report['applied']}")
    print(f"Kept in English    : {report['skipped_keep']}")
    print(f"Image text skipped : {report['skipped_image_text']}")
    print(f"Unmatched          : {len(report['unmatched'])}")
    print()

    if report["qa_corrections_applied"]:
        print("QA CORRECTIONS APPLIED (capitalisation/punctuation):")
        for c in report["qa_corrections_applied"]:
            print(f"  ✓ {c}")
        print()

    if report["expansion_font_reductions"]:
        print("FONT SIZE REDUCTIONS (CRITICAL expansion):")
        for e in report["expansion_font_reductions"]:
            print(f"  ↓ {e}")
        print()

    if report["unmatched"]:
        print(f"UNMATCHED SEGMENTS ({len(report['unmatched'])}):")
        for u in report["unmatched"]:
            print(f"  ✗ {u}")
        print()
        print("NOTE: Unmatched segments may indicate:")
        print("  - Text split across multiple shapes in source PPTX")
        print("  - PDF-to-PPTX text extraction differences")
        print("  - Slides beyond the translated batch range")
    else:
        print("All segments matched successfully.")

    status = "COMPLETE ✓" if not report["unmatched"] else "COMPLETE WITH WARNINGS ⚠"
    print()
    print(f"STATUS: {status}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="CIPS Translation Pipeline — PPTX Reconstruction"
    )
    parser.add_argument(
        "--source", required=True,
        help="Path to the original source PPTX file"
    )
    parser.add_argument(
        "--translations", required=True,
        help="Path to the Agent 3 JSON output (translated segments)"
    )
    parser.add_argument(
        "--qa", default=None,
        help="Path to the Agent 4 JSON output (QA issues) — optional"
    )
    parser.add_argument(
        "--output", required=True,
        help="Path for the translated output PPTX file"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview matches without writing output"
    )
    args = parser.parse_args()

    # Validate inputs
    if not Path(args.source).exists():
        print(f"ERROR: Source file not found: {args.source}")
        sys.exit(1)
    if not Path(args.translations).exists():
        print(f"ERROR: Translations file not found: {args.translations}")
        sys.exit(1)
    if args.qa and not Path(args.qa).exists():
        print(f"ERROR: QA file not found: {args.qa}")
        sys.exit(1)

    reconstruct(
        source_path=args.source,
        translations_path=args.translations,
        output_path=args.output,
        qa_path=args.qa,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
