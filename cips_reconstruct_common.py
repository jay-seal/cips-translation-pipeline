"""
cips_reconstruct_common.py
==========================
CIPS Translation Pipeline — Shared reconstruction helpers.

Format-agnostic logic used by every reconstruction script
(cips_reconstruct_pptx.py, cips_reconstruct_docx.py, and future
cips_reconstruct_<format>.py scripts).

This module deliberately has NO dependency on python-pptx, python-docx, or
any format-specific library. It operates only on the translation JSON, so
the same validation behaves identically across all output formats and is
maintained in exactly one place.

Contents
--------
1. decode_newlines()   — normalise LLM double-escaped newlines to real U+000A.
2. run_preflight()     — format-agnostic validation gates over the segments,
                         folded into the match report under a "preflight" key.

Preflight gates (replacing the former Agent 4 QA agent)
-------------------------------------------------------
These checks were previously performed by an LLM QA agent whose output never
actually reached reconstruction. They are deterministic — counting lines,
measuring length, reading a status field — so they belong in a script, not a
language model. run_preflight() records every finding in the report; whether
any of them halts the run is controlled by the caller.

    LINE_COUNT_MISMATCH  The number of newline-separated lines differs between
                         source_text and translated_text. This is the one gate
                         that can corrupt output (translated lines get mapped
                         onto the wrong paragraphs), so it is the only gate
                         that can optionally block — see `strict_line_count`.

    FLAGGED_SEGMENT      A segment carries flag_status == "FLAG", meaning a
                         terminology decision was left for a human. Always a
                         warning: flagged text is still rendered, but the
                         reviewer should confirm the chosen wording.

    EXPANSION_WARNING    Translated text is materially longer than the source
                         on a length-constrained element (titles, labels,
                         bullets, text boxes). Always a warning: this is a
                         layout risk visible only after reconstruction and is
                         correctable in the output document, so it must never
                         block — consistent with the pipeline's routing rules.
"""

from __future__ import annotations

import logging
from typing import Any

log = logging.getLogger("cips_reconstruct")


# ---------------------------------------------------------------------------
# Newline decoding
# ---------------------------------------------------------------------------

def decode_newlines(text: str | None) -> str | None:
    """
    Convert the literal two-character sequence backslash+n into a real newline
    (U+000A).

    LLM agents frequently double-escape newlines when emitting JSON, storing
    "\\n" (backslash followed by n) rather than an actual line break. Left
    undecoded, every line-count comparison and paragraph mapping downstream is
    wrong. Both source_text and translated_text must be passed through this
    before any comparison or matching.

    Returns the input unchanged if it is None or contains no escaped newlines.
    """
    if not text:
        return text
    return text.replace('\\n', '\x0a')


# ---------------------------------------------------------------------------
# Preflight validation
# ---------------------------------------------------------------------------

# Element types where text sits in a size-constrained space and expansion is
# a genuine layout risk worth flagging. Unconstrained contexts (footers,
# speaker notes, flowing body text) are not checked for expansion.
_CONSTRAINED_ELEMENT_TYPES = frozenset({
    "slide_title", "slide_subtitle", "heading",
    "label", "text_box", "bullet_point",
})

# Expansion thresholds for constrained elements. A finding is recorded only
# when BOTH the absolute and percentage thresholds are exceeded, so that a
# tiny source string growing by a few characters does not raise noise.
_EXPANSION_ABS_CHARS = 10      # minimum absolute character growth
_EXPANSION_PCT = 20.0          # minimum percentage growth


def _line_count(text: str | None) -> int:
    """Number of newline-separated lines in already-decoded text."""
    if not text:
        return 0
    return text.count('\x0a') + 1


def run_preflight(segments: list[dict[str, Any]],
                  strict_line_count: bool = False) -> dict[str, Any]:
    """
    Run the deterministic preflight gates over the translation segments.

    Parameters
    ----------
    segments
        The list of segment dicts from the Agent 3 translation JSON.
    strict_line_count
        If True, a line-count mismatch is recorded as blocking (severity
        "BLOCKER") and the caller is expected to halt. If False (default),
        it is recorded as a warning and the run proceeds. All other gates
        are always non-blocking regardless of this flag.

    Returns
    -------
    A dict suitable for folding into the match report under a "preflight" key:

        {
          "blocking": bool,          # True only if a BLOCKER-severity issue exists
          "summary": {
            "line_count_mismatch": int,
            "flagged_segments":    int,
            "expansion_warnings":  int,
          },
          "issues": [ {segment_id, check, severity, detail, ...}, ... ]
        }
    """
    issues: list[dict[str, Any]] = []
    counts = {
        "line_count_mismatch": 0,
        "flagged_segments": 0,
        "expansion_warnings": 0,
    }

    for seg in segments:
        seg_id = seg.get("segment_id", "UNKNOWN")
        status = seg.get("translation_status", "")

        # Only inspect segments that actually carry a translation. Untranslated
        # (PENDING), kept, or do-not-translate segments have nothing to check.
        translated_raw = seg.get("translated_text")
        if not translated_raw:
            continue
        if status in ("DO_NOT_TRANSLATE", "KEPT", "KEPT — IMAGE TEXT"):
            continue

        source = decode_newlines(seg.get("source_text", "")) or ""
        translated = decode_newlines(translated_raw) or ""

        # --- Gate 1: line-count mismatch ---------------------------------
        src_lines = _line_count(source)
        tgt_lines = _line_count(translated)
        if src_lines != tgt_lines:
            counts["line_count_mismatch"] += 1
            severity = "BLOCKER" if strict_line_count else "WARNING"
            issues.append({
                "segment_id": seg_id,
                "check": "LINE_COUNT_MISMATCH",
                "severity": severity,
                "source_lines": src_lines,
                "translated_lines": tgt_lines,
                "detail": (
                    f"Source has {src_lines} line(s), translation has "
                    f"{tgt_lines}. Line breaks must match exactly or content "
                    f"may map onto the wrong paragraphs during reconstruction."
                ),
            })

        # --- Gate 2: flagged segment -------------------------------------
        if seg.get("flag_status") == "FLAG":
            counts["flagged_segments"] += 1
            issues.append({
                "segment_id": seg_id,
                "check": "FLAGGED_SEGMENT",
                "severity": "WARNING",
                "flag_options": seg.get("flag_options"),
                "detail": (
                    "Segment carries an unresolved terminology FLAG. The "
                    "translation was rendered using a best-guess option; a "
                    "reviewer should confirm the correct term."
                ),
            })

        # --- Gate 3: expansion on constrained elements -------------------
        element_type = seg.get("element_type", "")
        in_text_box = seg.get("is_in_text_box") is True
        if element_type in _CONSTRAINED_ELEMENT_TYPES or in_text_box:
            src_len = len(source)
            tgt_len = len(translated)
            abs_diff = tgt_len - src_len
            pct = (abs_diff / src_len * 100) if src_len else 0.0
            if abs_diff > _EXPANSION_ABS_CHARS and pct > _EXPANSION_PCT:
                counts["expansion_warnings"] += 1
                issues.append({
                    "segment_id": seg_id,
                    "check": "EXPANSION_WARNING",
                    "severity": "WARNING",
                    "source_length": src_len,
                    "translated_length": tgt_len,
                    "absolute_difference": abs_diff,
                    "expansion_percentage": round(pct, 1),
                    "element_type": element_type,
                    "detail": (
                        f"Translation is {abs_diff} chars longer "
                        f"({round(pct, 1)}%) than source on a constrained "
                        f"element. Check layout in the output document."
                    ),
                })

    blocking = any(i["severity"] == "BLOCKER" for i in issues)

    # --- Log a concise summary ------------------------------------------
    log.info("=" * 60)
    log.info("PREFLIGHT VALIDATION")
    log.info("  Line-count mismatches : %d", counts["line_count_mismatch"])
    log.info("  Flagged segments      : %d", counts["flagged_segments"])
    log.info("  Expansion warnings    : %d", counts["expansion_warnings"])
    if strict_line_count and counts["line_count_mismatch"]:
        log.info("  Strict line-count is ON — mismatches will BLOCK.")
    log.info("=" * 60)

    return {
        "blocking": blocking,
        "strict_line_count": strict_line_count,
        "summary": counts,
        "issues": issues,
    }
