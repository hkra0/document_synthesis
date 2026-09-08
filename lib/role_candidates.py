# -*- coding: utf-8 -*-
"""Evidence-based role candidates for samples without outline metadata.

The output is deliberately a ranked candidate list, not a probability. It is
used to assist the review UI; explicit RoleMap and structural outline evidence
remain higher-priority decisions.
"""

from __future__ import annotations

import re
from collections import Counter
from statistics import median
from typing import Any, Dict, Iterable, List, Mapping


_NUMBERED = re.compile(
    r"^(?:第?[一二三四五六七八九十百]+[章节篇部分级、.]|第?[0-9]+(?:级|章节?|篇|部分|条|[.)、.])|[0-9]+(?:\.[0-9]+)*[.)、]|[（(][0-9]+[）)])"
)
_MIN_CANDIDATE_SCORE = 0.25
_AUTO_ACCEPT_SCORE = 0.985


def _numbering_evidence(text: str) -> Dict[str, Any]:
    normalized = (text or "").strip()
    hierarchical = re.match(r"^(\d+(?:\.\d+)+)(?:[.)、]\s+|\s+)", normalized)
    if hierarchical:
        components = hierarchical.group(1).split(".")
        if (
            len(components) == 3
            and len(components[0]) == 4
            and components[0] in {str(year) for year in range(1900, 2101)}
            and all(1 <= len(component) <= 2 for component in components[1:])
        ):
            return {
                "matched": False,
                "kind": "date_or_version",
                "components": components,
                "depth": None,
                "confidence": 0.0,
                "counter_evidence": ["four_digit_year", "date_like_components"],
            }
        return {
            "matched": True,
            "kind": "hierarchical_arabic",
            "components": components,
            "depth": len(components),
            "confidence": 0.98,
            "counter_evidence": [],
        }
    section = re.match(r"^第([一二三四五六七八九十百]+|\d+)(?:章|节|篇|部分|级|条)", normalized)
    if section:
        return {
            "matched": True,
            "kind": "section_label",
            "components": [section.group(1)],
            "depth": 1,
            "confidence": 0.90,
            "counter_evidence": [],
        }
    simple = re.match(r"^(\d+)[.)、]\s*", normalized)
    if simple:
        return {
            "matched": True,
            "kind": "simple_arabic",
            "components": [simple.group(1)],
            "depth": 1,
            "confidence": 0.75,
            "counter_evidence": [],
        }
    decimal_or_version = re.match(r"^\d+\.\d+(?:\.\d+)*\b", normalized)
    if decimal_or_version:
        return {
            "matched": False,
            "kind": "decimal_or_version",
            "components": decimal_or_version.group(0).split("."),
            "depth": None,
            "confidence": 0.0,
            "counter_evidence": ["no_heading_delimiter"],
        }
    return {
        "matched": False,
        "kind": None,
        "components": [],
        "depth": None,
        "confidence": 0.0,
        "counter_evidence": [],
    }


def _run(block: Any) -> Any:
    return getattr(block, "effective_run", None)


def _paragraph(block: Any) -> Any:
    return getattr(block, "effective_paragraph", None)


def _size(block: Any) -> float:
    value = getattr(_run(block), "size_pt", None)
    return float(value) if value is not None else 0.0


def _text_len(block: Any) -> int:
    return len(re.sub(r"\s+", "", getattr(block, "visible_text", "") or ""))


def _signals(block: Any, body_size: float, previous: Any = None, following: Any = None) -> Dict[str, Any]:
    run = _run(block)
    paragraph = _paragraph(block)
    size = _size(block)
    length = _text_len(block)
    alignment = getattr(paragraph, "alignment", None) if paragraph else None
    before = getattr(paragraph, "space_before_pt", None) if paragraph else None
    bold = bool(getattr(run, "bold", False)) if run else False
    numbering = _numbering_evidence(getattr(block, "visible_text", "") or "")
    numbered = bool(numbering["matched"])
    return {
        "size_pt": size or None,
        "relative_size": round(size / body_size, 3) if size and body_size else None,
        "short_text": length <= 42,
        "text_length": length,
        "centered": alignment == "center",
        "bold": bold,
        "space_before_pt": before,
        "numbered": numbered,
        "numbering_kind": numbering["kind"],
        "numbering_components": numbering["components"],
        "numbering_depth": numbering["depth"],
        "numbering_confidence": numbering["confidence"],
        "numbering_counter_evidence": numbering.get("counter_evidence", []),
        "neighbor_same_size_before": bool(previous is not None and abs(_size(previous) - size) < 0.1),
        "neighbor_same_size_after": bool(following is not None and abs(_size(following) - size) < 0.1),
    }


def rank_unstructured_candidates(blocks: Iterable[Any]) -> Dict[str, Dict[str, Any]]:
    """Rank at most three semantic roles for each non-table block.

    Scores combine independent visual/structural signals and explicit
    counter-signals. A long paragraph with a large first run, for example,
    does not become a heading from size alone.
    """
    blocks = list(blocks)
    body_like_sizes = [
        _size(block)
        for block in blocks
        if getattr(block, "structure_type", "") == "paragraph"
        and _size(block) > 0
        and (_text_len(block) > 80 or not bool(getattr(_run(block), "bold", False)))
    ]
    usable_sizes = body_like_sizes or [
        _size(block)
        for block in blocks
        if getattr(block, "structure_type", "") == "paragraph" and _size(block) > 0
    ]
    body_size = median(usable_sizes) if usable_sizes else 12.0
    preliminary: list[tuple[Any, Dict[str, Any]]] = []
    for block in blocks:
        if getattr(block, "structure_type", "") != "paragraph":
            continue
        signal = _signals(block, float(body_size))
        preliminary.append((block, signal))
    heading_sizes = [
        round(_size(block), 1)
        for block, signal in preliminary
        if signal["short_text"]
        and signal["bold"]
        and signal["numbered"]
        and not _text_len(block) > 80
        and signal["relative_size"] is not None
        and signal["relative_size"] >= 1.08
    ]
    heading_size_counts = Counter(heading_sizes)
    all_size_counts = Counter(round(_size(block), 1) for block, _ in preliminary if _size(block) > 0)
    ordered_heading_sizes = sorted(set(heading_sizes), reverse=True)
    # A repeated, numbered, bold style family provides document-local level
    # evidence.  It is stronger than a single large/short paragraph and lets
    # us distinguish heading.1 … heading.9 without assuming that every
    # unstructured document uses all nine levels.
    heading_level_by_size = {
        size: index + 1
        for index, size in enumerate(ordered_heading_sizes[:9])
    }
    result: Dict[str, Dict[str, Any]] = {}
    for index, block in enumerate(blocks):
        if getattr(block, "structure_type", "") != "paragraph":
            continue
        previous = blocks[index - 1] if index else None
        following = blocks[index + 1] if index + 1 < len(blocks) else None
        signal = _signals(block, float(body_size), previous, following)
        rel = signal["relative_size"] or 1.0
        short = signal["short_text"]
        long_text = signal["text_length"] > 80
        emphasis = signal["bold"] or signal["centered"]
        spacing = bool(signal["space_before_pt"] and signal["space_before_pt"] >= 6)
        numbered = signal["numbered"]
        size = _size(block)
        size_key = round(size, 1)
        inferred_level = (
            signal["numbering_depth"]
            if signal["numbering_depth"] and signal["numbering_depth"] >= 2
            else heading_level_by_size.get(size_key)
        )
        format_level = heading_level_by_size.get(size_key)
        conflicts = []
        if (
            signal["numbering_depth"]
            and signal["numbering_depth"] >= 2
            and format_level
            and format_level != signal["numbering_depth"]
        ):
            conflicts.append({
                "kind": "numbering_vs_size_level",
                "numbering_depth": signal["numbering_depth"],
                "size_level": format_level,
            })
        if signal["numbering_counter_evidence"]:
            conflicts.append({
                "kind": "numbering_rejected",
                "reasons": list(signal["numbering_counter_evidence"]),
            })
        repeated_heading_style = bool(
            inferred_level
            and heading_size_counts.get(size_key, 0) >= 1
            and len(ordered_heading_sizes) >= 2
        )
        calibrated_heading = bool(
            repeated_heading_style
            and short
            and emphasis
            and numbered
            and not long_text
            and not conflicts
        )
        heading_score = 0.05
        if rel >= 1.15:
            heading_score += min(0.35, (rel - 1.0) * 0.45)
        if short:
            heading_score += 0.20
        if emphasis:
            heading_score += 0.12
        if spacing:
            heading_score += 0.10
        if numbered:
            heading_score += 0.10
        if signal["numbering_counter_evidence"]:
            heading_score -= 0.25
        if long_text:
            heading_score -= 0.28
        if signal["neighbor_same_size_before"] and signal["neighbor_same_size_after"] and not numbered:
            heading_score -= 0.10
        heading_score = max(0.0, min(0.95, heading_score))
        if calibrated_heading:
            # This is a calibrated confidence band, not a probability.  The
            # combination is deliberately conjunctive: it requires a
            # repeated local heading family, a numbered short paragraph, and
            # explicit emphasis above the document body baseline.
            heading_score = _AUTO_ACCEPT_SCORE
        heading_role = f"heading.{inferred_level}" if inferred_level else "heading.1"

        title_score = heading_score
        if signal["centered"] and rel >= 1.35 and short:
            title_score = min(0.98, heading_score + 0.20)
        elif not signal["centered"] or rel < 1.20:
            title_score = max(0.0, heading_score - 0.18)
        body_score = 0.30
        if long_text:
            body_score += 0.30
        if rel <= 1.10:
            body_score += 0.20
        if not emphasis and not numbered:
            body_score += 0.10
        body_score = min(0.92, body_score)

        candidates = [
            ("title", title_score, [name for name, ok in (("centered", signal["centered"]), ("relative_size", rel >= 1.35), ("short_text", short)) if ok], ["long_text"] if long_text else []),
            (heading_role, heading_score, [name for name, ok in (("relative_size", rel >= 1.15), ("short_text", short), ("bold_or_centered", emphasis), ("spacing", spacing), ("numbered", numbered), ("numbering_depth", bool(signal["numbering_depth"] and signal["numbering_depth"] >= 2)), ("repeated_heading_style", repeated_heading_style)) if ok], ["long_text", "repeated_body_style", *signal["numbering_counter_evidence"]] if long_text or (signal["neighbor_same_size_before"] and signal["neighbor_same_size_after"]) or signal["numbering_counter_evidence"] else []),
            ("body", body_score, [name for name, ok in (("long_text", long_text), ("body_sized", rel <= 1.10), ("not_emphasized", not emphasis)) if ok], ["short_text"] if short else []),
            ("caption", 0.25 if short and not numbered else 0.08, ["short_text"] if short else [], ["numbered"] if numbered else []),
        ]
        candidates.sort(key=lambda item: (-item[1], item[0]))
        # A candidate list is an actionable review shortlist, not a dump of
        # every role considered by the scorer.  In particular, keeping a
        # zero-score heading as the third item makes every long body paragraph
        # look like a heading candidate and corrupts candidate precision.
        competitive = [item for item in candidates if item[1] >= _MIN_CANDIDATE_SCORE]
        top = (competitive[:3] or candidates[:1])
        level_conflict = any(
            item.get("kind") == "numbering_vs_size_level" for item in conflicts
        )
        selected = (
            heading_role
            if calibrated_heading or level_conflict
            else (top[0][0] if top[0][1] >= 0.58 and (
                len(top) == 1 or top[0][1] - top[1][1] >= 0.05
            ) else "body")
        )
        result[block.node.element_path] = {
            "selected_role": selected,
            "selected_score": round(top[0][1], 4),
            "score_kind": "calibrated_confidence_band" if calibrated_heading else "heuristic_score",
            "review_status": "auto_accepted" if calibrated_heading else ("pending_review" if conflicts else "selected"),
            "conflicts": conflicts,
            "numbering_evidence": {
                "kind": signal["numbering_kind"],
                "components": signal["numbering_components"],
                "depth": signal["numbering_depth"],
                "confidence": signal["numbering_confidence"],
            },
            "candidates": [
                {
                    "role": role,
                    "score": round(score, 4),
                    "signals": signals,
                    "counter_signals": counters,
                    "sample_count": (
                        heading_size_counts.get(size_key, 0)
                        if role.startswith("heading.")
                        else all_size_counts.get(size_key, 0)
                    ),
                    "impact_nodes": [block.node.element_path],
                }
                for role, score, signals, counters in top
            ],
        }
    return result
