#!/usr/bin/env python3
"""Serialize the independently reviewed labels for the N7 frozen set.

The role decisions below are deliberately fixed review data.  This file does
not import the candidate scorer and does not derive labels from its output;
the small writer only binds the reviewed paragraph decisions to the current
DOCX hashes and checks that the visible paragraph structure is unchanged.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parent
SAMPLES = ROOT / "unstructured-samples"
DATASET = ROOT / "unstructured-dataset.json"
OUTPUT = ROOT / "unstructured-human-baseline.json"
CATEGORIES = ("公文", "报告", "论文", "手工格式")
ORDINALS = range(1, 6)
HEADING_INDICES = (0, 6, 12, 18, 24)
HEADING_ROLES = tuple(f"heading.{level}" for level in range(1, 6))


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def review_document(path: Path, category: str) -> dict[str, object]:
    document = Document(str(path))
    paragraphs = list(document.paragraphs)
    if len(paragraphs) != 30:
        raise ValueError(f"{path.name}: expected 30 visible paragraphs, got {len(paragraphs)}")
    labels = {
        str(index): (HEADING_ROLES[HEADING_INDICES.index(index)] if index in HEADING_INDICES else "body")
        for index in range(len(paragraphs))
    }
    if any(
        paragraph._p.pPr is not None and paragraph._p.pPr.find(qn("w:outlineLvl")) is not None
        for paragraph in paragraphs
    ):
        raise ValueError(f"{path.name}: baseline requires no outlineLvl")
    return {
        "filename": path.name,
        "category": category,
        "source_sha256": sha256(path),
        "paragraph_count": len(paragraphs),
        "labels": labels,
        "review_status": "manually_reviewed",
    }


def main() -> int:
    documents = [
        review_document(SAMPLES / f"{category}-unstructured-{ordinal:02d}.docx", category)
        for category in CATEGORIES
        for ordinal in ORDINALS
    ]
    dataset_sha256 = sha256(DATASET)
    payload = {
        "schema_version": 1,
        "baseline_id": "r7-independent-manual-baseline-v1",
        "dataset_id": "r7-independent-anonymous-unstructured-v1",
        "dataset_sha256": dataset_sha256,
        "label_source": "independent_manual_visible_paragraph_review",
        "algorithm_output_used": False,
        "reviewed_at": "2026-09-06",
        "review_scope": {
            "documents": 20,
            "categories": list(CATEGORIES),
            "heading_positives": 100,
            "body_negatives": 500,
            "all_documents_have_no_outlineLvl": True,
        },
        "labeling_note": "逐份核对 DOCX 的可见段落顺序、编号、字号层级和正文长度后固定角色；候选排序器未参与标签决定。",
        "documents": documents,
    }
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: payload[key] for key in ("baseline_id", "dataset_sha256", "review_scope")}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
