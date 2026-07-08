"""
cips_merge_batches.py
=====================
CIPS Translation Pipeline — Batch JSON Merge Script

Merges two or more Agent 3 translation JSON files (one per slide batch) into
a single JSON file suitable for a single cips_reconstruct.py run. The output
is structurally identical to a single-batch Agent 3 JSON, so reconstruct.py
requires no modification.

Usage:
    python cips_merge_batches.py \\
        --inputs inputs/agent3_m1_fr-fr_slides_001_020.json \\
                 inputs/agent3_m1_fr-fr_slides_021_040.json \\
                 inputs/agent3_m1_fr-fr_slides_041_060.json \\
        --output inputs/agent3_m1_fr-fr_merged.json

Validations performed before writing output:
    - All input files must exist and be valid JSON.
    - All batches must share the same locale_code value in translation_summary.
    - Overlapping slide ranges are detected and reported as hard errors.
    - Gaps in slide coverage are reported as warnings (not errors — blank
      slides and section dividers legitimately produce no segments).

Duplicate segment_id handling:
    Agent 1 restarts segment numbering from SEG-001 in each batch. Duplicate
    IDs across batches are therefore expected and are NOT treated as errors.
    After sorting all segments by slide then original segment number, the
    script reassigns clean sequential IDs (SEG-001, SEG-002, ...) across the
    full merged output. Each segment retains its original batch ID in the
    'original_segment_id' field for traceability.

The merged output recomputes translation_summary counts from the combined
segment list and records the source batch filenames in a 'merged_from' field
for traceability.

Exit codes:
    0   Success — merged JSON written.
    1   Invalid input (missing files, malformed JSON, wrong arguments).
    2   Validation failure (locale mismatch, overlapping slide ranges).
"""

import argparse
import json
import logging
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("cips_merge_batches")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_batch(path: Path) -> dict:
    if not path.is_file():
        log.error("Input file not found: %s", path)
        sys.exit(1)
    with path.open(encoding="utf-8") as fh:
        try:
            data = json.load(fh)
        except json.JSONDecodeError as exc:
            log.error("Invalid JSON in %s: %s", path, exc)
            sys.exit(1)
    if "segments" not in data:
        log.error(
            "File %s does not contain a 'segments' array — is it an Agent 3 output?",
            path,
        )
        sys.exit(1)
    return data


def _seg_sort_key(seg: dict):
    """
    Sort by slide number, then by numeric part of the original segment_id.
    Uses original_segment_id if already set (post-duplicate-detection),
    otherwise falls back to segment_id.
    """
    slide = seg.get("slide_or_page") or 0
    id_to_use = seg.get("original_segment_id") or seg.get("segment_id", "")
    digits = "".join(c for c in id_to_use if c.isdigit())
    num = int(digits) if digits else 0
    return (slide, num)


def _recompute_summary(segments: list, locale_code: str, merged_from: list) -> dict:
    """Build a fresh translation_summary from the merged segment list."""
    total          = len(segments)
    translated     = sum(1 for s in segments if s.get("translation_status") == "TRANSLATED")
    partial        = sum(1 for s in segments if s.get("translation_status") == "PARTIAL_TRANSLATED")
    kept           = sum(1 for s in segments if "KEPT" in (s.get("translation_status") or ""))
    image_text     = sum(1 for s in segments if s.get("translation_status") == "KEPT — IMAGE TEXT")
    flagged        = sum(1 for s in segments if "FLAGGED" in (s.get("translation_status") or ""))
    expansion_warn = sum(1 for s in segments if s.get("expansion_note"))
    return {
        "total_segments":           total,
        "translated_count":         translated,
        "partial_translated_count": partial,
        "kept_count":               kept,
        "image_text_count":         image_text,
        "flagged_count":            flagged,
        "expansion_warnings":       expansion_warn,
        "locale_code":              locale_code,
        "merged_from":              merged_from,
    }


# ---------------------------------------------------------------------------
# Core merge logic
# ---------------------------------------------------------------------------

def merge_batches(input_paths: list, output_path: Path) -> None:

    # -----------------------------------------------------------------------
    # 1. Load all batches
    # -----------------------------------------------------------------------
    batches = []
    for p in input_paths:
        data = _load_batch(p)
        n = len(data["segments"])
        log.info("Loaded %-60s  (%d segments)", str(p), n)
        batches.append((p, data))

    if len(batches) < 2:
        log.warning(
            "Only one input file provided. Merge is a no-op — "
            "consider passing the file directly to cips_reconstruct_pptx.py "
            "or cips_reconstruct_docx.py."
        )

    # -----------------------------------------------------------------------
    # 1b. Completeness guard: Agent 3 output count must equal Agent 1 input count
    #
    # A single Opal loop iteration can silently return fewer segments than it
    # was given (e.g. an LLM stops generating early but still emits valid JSON
    # with a summary matching its truncated output). Downstream everything
    # proceeds happily and a partially-translated document ships. This guard
    # catches that by comparing each Agent 3 batch against its Agent 1 source
    # of truth — the extraction, which is deterministic and complete.
    #
    # The Agent 1 sibling is found by swapping 'agent3_' -> 'agent1_' in the
    # filename, in the same job folder. If a sibling cannot be found, that is
    # itself treated as a hard error rather than skipped, so the guard can
    # never be silently bypassed.
    # -----------------------------------------------------------------------
    count_errors = []
    for p, data in batches:
        agent3_name = Path(p).name
        if "agent3_" not in agent3_name:
            count_errors.append(
                f"{agent3_name}: filename does not contain 'agent3_', cannot "
                f"locate its Agent 1 source to verify completeness."
            )
            continue

        agent1_path = Path(p).with_name(agent3_name.replace("agent3_", "agent1_", 1))
        if not agent1_path.is_file():
            count_errors.append(
                f"{agent3_name}: Agent 1 source {agent1_path.name} not found in "
                f"the job folder, cannot verify segment completeness."
            )
            continue

        try:
            with agent1_path.open(encoding="utf-8") as fh:
                agent1_data = json.load(fh)
            agent1_count = len(agent1_data.get("segments", []))
        except (json.JSONDecodeError, OSError) as e:
            count_errors.append(
                f"{agent3_name}: could not read Agent 1 source {agent1_path.name} "
                f"to verify completeness ({e})."
            )
            continue

        agent3_count = len(data.get("segments", []))
        if agent3_count != agent1_count:
            count_errors.append(
                f"{agent3_name}: translated {agent3_count} segment(s) but the "
                f"source batch {agent1_path.name} contains {agent1_count}. "
                f"{agent1_count - agent3_count} segment(s) were lost during "
                f"translation — this batch must be re-run."
            )

    if count_errors:
        for e in count_errors:
            log.error(e)
        log.error(
            "%d batch completeness error(s) — one or more batches returned "
            "fewer segments than they were given. Merge aborted to prevent a "
            "partially-translated document from being produced. Re-run the "
            "affected batch(es) through the Opal pipeline, then merge again.",
            len(count_errors),
        )
        sys.exit(2)

    log.info(
        "Completeness check passed: all %d batch(es) returned the same segment "
        "count as their Agent 1 source.", len(batches),
    )

    # -----------------------------------------------------------------------
    # 2. Validate locale consistency
    # -----------------------------------------------------------------------
    locale_codes = {}
    for p, data in batches:
        lc = data.get("translation_summary", {}).get("locale_code")
        if lc:
            locale_codes[lc] = locale_codes.get(lc, []) + [str(p)]

    if len(locale_codes) > 1:
        log.error(
            "Batches have conflicting locale_codes: %s. "
            "All batches must share the same locale.",
            {lc: files for lc, files in locale_codes.items()},
        )
        sys.exit(2)

    locale_code = next(iter(locale_codes)) if locale_codes else None
    log.info("Locale: %s", locale_code or "(not set in any batch)")

    # -----------------------------------------------------------------------
    # 3. Collect all segments, preserving original IDs for traceability
    #
    # Agent 1 restarts from SEG-001 in each batch — duplicate IDs across
    # batches are expected. We collect every segment, preserve its original
    # ID, and renumber the full set sequentially after sorting.
    # -----------------------------------------------------------------------
    all_segments = []
    seen_ids: dict = {}
    has_duplicates = False

    for p, data in batches:
        for seg in data["segments"]:
            seg_id = seg.get("segment_id", "UNKNOWN")
            if seg_id in seen_ids:
                has_duplicates = True
            else:
                seen_ids[seg_id] = str(p)
            # Preserve the original batch-local ID before any renumbering.
            seg["original_segment_id"] = seg_id
            all_segments.append(seg)

    if has_duplicates:
        log.warning(
            "Duplicate segment_ids detected across batches (expected — "
            "Agent 1 restarts from SEG-001 per batch). All segments will "
            "be renumbered sequentially in the merged output. "
            "Original IDs are preserved in 'original_segment_id'."
        )

    # -----------------------------------------------------------------------
    # 4. Validate slide range overlap
    # -----------------------------------------------------------------------
    batch_slide_sets = []
    for p, data in batches:
        slides = {
            seg.get("slide_or_page")
            for seg in data["segments"]
            if isinstance(seg.get("slide_or_page"), int)
        }
        batch_slide_sets.append((str(p), slides))

    overlap_errors = []
    for i in range(len(batch_slide_sets)):
        for j in range(i + 1, len(batch_slide_sets)):
            name_i, slides_i = batch_slide_sets[i]
            name_j, slides_j = batch_slide_sets[j]
            overlap = slides_i & slides_j
            if overlap:
                overlap_errors.append(
                    f"Slide overlap between {name_i} and {name_j}: "
                    f"slides {sorted(overlap)}"
                )

    if overlap_errors:
        for e in overlap_errors:
            log.error(e)
        log.error(
            "%d slide overlap error(s) — each slide must appear in "
            "exactly one batch. Merge aborted.",
            len(overlap_errors),
        )
        sys.exit(2)

    # -----------------------------------------------------------------------
    # 5. Warn on coverage gaps
    # -----------------------------------------------------------------------
    all_slide_nums = sorted(
        {seg.get("slide_or_page") for seg in all_segments
         if isinstance(seg.get("slide_or_page"), int)}
    )
    if all_slide_nums:
        min_s, max_s = all_slide_nums[0], all_slide_nums[-1]
        slide_set = set(all_slide_nums)
        gaps = [s for s in range(min_s, max_s + 1) if s not in slide_set]
        if gaps:
            display = gaps[:20]
            tail = f" ... ({len(gaps) - 20} more)" if len(gaps) > 20 else ""
            log.warning(
                "Slides with no segments (may be blank/section dividers): %s%s",
                display, tail,
            )
        log.info(
            "Slide coverage: %d–%d  (%d slides represented across %d batches)",
            min_s, max_s, len(all_slide_nums), len(batches),
        )

    # -----------------------------------------------------------------------
    # 6. Sort and reassign sequential segment IDs
    # -----------------------------------------------------------------------
    all_segments.sort(key=_seg_sort_key)

    for i, seg in enumerate(all_segments, start=1):
        seg["segment_id"] = f"SEG-{i:03d}"

    log.info(
        "Segments renumbered SEG-001 to SEG-%03d across %d batches.",
        len(all_segments), len(batches),
    )

    # -----------------------------------------------------------------------
    # 7. Recompute summary
    # -----------------------------------------------------------------------
    merged_from = [str(p) for p, _ in batches]
    summary = _recompute_summary(all_segments, locale_code, merged_from)

    # Use document_metadata from the first batch; it describes the source file.
    document_metadata = batches[0][1].get("document_metadata", {})

    merged_output = {
        "translation_summary": summary,
        "document_metadata":   document_metadata,
        "segments":            all_segments,
    }

    # -----------------------------------------------------------------------
    # 8. Write output
    # -----------------------------------------------------------------------
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(merged_output, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # -----------------------------------------------------------------------
    # 9. Summary log
    # -----------------------------------------------------------------------
    log.info("=" * 60)
    log.info("MERGE COMPLETE")
    log.info("  Input batches          : %d", len(batches))
    log.info("  Total segments         : %d", summary["total_segments"])
    log.info("  Translated             : %d", summary["translated_count"])
    log.info("  Partial translated     : %d", summary["partial_translated_count"])
    log.info("  Kept (incl. image)     : %d  (image text: %d)",
             summary["kept_count"], summary["image_text_count"])
    log.info("  Flagged                : %d", summary["flagged_count"])
    log.info("  Expansion warnings     : %d", summary["expansion_warnings"])
    log.info("  Output                 : %s", output_path)
    log.info("=" * 60)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="CIPS Translation Pipeline — Batch JSON Merge Script"
    )
    parser.add_argument(
        "--inputs", nargs="+", required=True,
        help=(
            "Paths to Agent 3 batch JSON files to merge, in slide order "
            "(space-separated). Example: inputs/agent3_m1_fr-fr_slides_001_020.json "
            "inputs/agent3_m1_fr-fr_slides_021_040.json"
        ),
    )
    parser.add_argument(
        "--output", required=True,
        help=(
            "Destination path for the merged JSON. "
            "Example: inputs/agent3_m1_fr-fr_merged.json"
        ),
    )
    args = parser.parse_args()

    input_paths = [Path(p) for p in args.inputs]
    output_path = Path(args.output)

    merge_batches(input_paths, output_path)


if __name__ == "__main__":
    main()
