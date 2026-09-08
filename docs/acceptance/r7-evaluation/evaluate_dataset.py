# -*- coding: utf-8 -*-
"""Evaluate the independent R7 dataset without Word, network, or source text output."""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict

REPO_ROOT = Path(__file__).resolve().parents[3]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from lib.format_analysis import analyze_format_sample
from lib.role_evaluation import evaluate_role_metrics


THRESHOLD = 0.98
HEADING_LEVELS = [f"heading.{level}" for level in range(1, 6)]


def _load_dataset(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _path_for_paragraph(index: int) -> str:
    return f"/w:document/w:body/w:p[{index + 1}]"


def _metrics_for_labels(labels: Dict[str, str], predictions: Dict[str, str], role: str) -> Dict[str, Any]:
    support = sum(value == role for value in labels.values())
    predicted = sum(value == role for value in predictions.values())
    true_positive = sum(labels.get(path) == role and predictions.get(path) == role for path in labels)
    false_positive = max(predicted - true_positive, 0)
    false_negative = max(support - true_positive, 0)
    return {
        "support": support,
        "predicted": predicted,
        "true_positive": true_positive,
        "false_positive": false_positive,
        "false_negative": false_negative,
        "precision": true_positive / predicted if predicted else None,
        "recall": true_positive / support if support else None,
    }


def evaluate_dataset(dataset: Dict[str, Any], samples_dir: Path) -> Dict[str, Any]:
    all_labels: Dict[str, str] = {}
    all_predictions: Dict[str, str] = {}
    per_category: Dict[str, Dict[str, Any]] = defaultdict(lambda: {
        "labels": {}, "predictions": {}, "accepted": {}, "candidate_roles": {}, "documents": 0
    })
    accepted_predictions: Dict[str, str] = {}
    candidate_role_map: Dict[str, list[str]] = {}
    level_auto = Counter()
    level_auto_correct = Counter()
    level_errors = Counter()
    candidate_review = 0
    candidate_total = 0
    pending_total = 0
    pending_denominator = 0

    for item in dataset["documents"]:
        report = analyze_format_sample(samples_dir / item["filename"])
        category = item["category"]
        per_category[category]["documents"] += 1
        labels: Dict[str, str] = {}
        for index in item["heading_paragraph_indices"]:
            labels[_path_for_paragraph(index)] = f"heading.{item['heading_paragraph_indices'].index(index) + 1}"
        for index in item["body_paragraph_indices"]:
            labels[_path_for_paragraph(index)] = "body"
        predictions = {
            path: report.get("node_roles", {}).get(path, "unverified")
            for path in labels
        }
        # Prefix paths with the source filename so identical DOCX element
        # paths from different documents remain separate observations.
        unique_labels = {f"{item['filename']}:{path}": value for path, value in labels.items()}
        unique_predictions = {f"{item['filename']}:{path}": value for path, value in predictions.items()}
        all_labels.update(unique_labels)
        all_predictions.update(unique_predictions)
        per_category[category]["labels"].update(unique_labels)
        per_category[category]["predictions"].update(unique_predictions)
        candidates = report.get("candidate_roles", {})
        for path, prediction in predictions.items():
            unique_path = f"{item['filename']}:{path}"
            roles = [
                role for role, candidate in candidates.items()
                if isinstance(candidate, dict)
                and float(candidate.get("confidence", 0.0)) >= THRESHOLD
            ]
            candidate_role_map[unique_path] = roles
            per_category[category]["candidate_roles"][unique_path] = roles
            if prediction.startswith("heading.") and prediction in roles:
                accepted_predictions[unique_path] = prediction
                per_category[category]["accepted"][unique_path] = prediction

        candidate_total += len(candidates)
        candidate_review += sum(
            1 for candidate in candidates.values()
            if float(candidate.get("confidence", 0.0)) < THRESHOLD
        )
        pending_total += len(report.get("missing_roles", []))
        pending_denominator += len(candidates) + len(report.get("missing_roles", []))
        for role in HEADING_LEVELS:
            candidate = candidates.get(role)
            if not candidate or float(candidate.get("confidence", 0.0)) < THRESHOLD:
                continue
            count = sum(value == role for value in labels.values())
            level_auto[role] += count
            level_auto_correct[role] += sum(
                labels[path] == role and predictions[path] == role
                for path in labels
                if labels[path] == role
            )
        for path, expected in labels.items():
            if expected.startswith("heading.") and predictions[path] != expected:
                level_errors[expected] += 1

    per_level = {role: _metrics_for_labels(all_labels, all_predictions, role) for role in HEADING_LEVELS}
    category_results = {}
    for category, raw in per_category.items():
        labels = raw.pop("labels")
        predictions = raw.pop("predictions")
        category_metrics = evaluate_role_metrics(
            labels, predictions, raw.pop("accepted"), raw.pop("candidate_roles")
        )
        category_results[category] = {
            **raw,
            "per_level": {role: _metrics_for_labels(labels, predictions, role) for role in HEADING_LEVELS},
            "body_as_heading_false_positives": sum(
                labels[path] == "body" and predictions[path].startswith("heading.")
                for path in labels
            ),
            "heading_detection_precision": category_metrics["heading_detection_precision"],
            "exact_role_auto_precision": category_metrics["exact_role_auto_precision"],
            "auto_acceptance_coverage": category_metrics["auto_acceptance_coverage"],
            "correct_auto_coverage": category_metrics["correct_auto_coverage"],
            "candidate_exact_role_coverage": category_metrics["candidate_exact_role_coverage"],
        }

    heading_count = dataset["heading_positive_count"]
    auto_count = sum(level_auto.values())
    auto_correct = sum(level_auto_correct.values())
    overall_role_metrics = evaluate_role_metrics(
        all_labels, all_predictions, accepted_predictions, candidate_role_map
    )
    return {
        "metrics_schema_version": 2,
        "dataset_id": dataset["dataset_id"],
        "threshold": THRESHOLD,
        "dataset_counts": {
            "documents": dataset["document_count"],
            "categories": len(dataset["categories"]),
            "heading_positives": heading_count,
            "body_negatives": dataset["body_negative_count"],
        },
        "overall": {
            "per_level": per_level,
            "heading_accuracy": sum(
                all_labels[path] == all_predictions[path]
                for path in all_labels
                if all_labels[path].startswith("heading.")
            ) / heading_count,
            "body_as_heading_false_positives": sum(
                all_labels[path] == "body" and all_predictions[path].startswith("heading.")
                for path in all_labels
            ),
            "heading_detection_precision": overall_role_metrics["heading_detection_precision"],
            "exact_role_auto_precision": overall_role_metrics["exact_role_auto_precision"],
            "automatic_acceptance_coverage": overall_role_metrics["auto_acceptance_coverage"],
            "correct_auto_coverage": overall_role_metrics["correct_auto_coverage"],
            "candidate_exact_role_coverage": overall_role_metrics["candidate_exact_role_coverage"],
            "automatic_acceptance_precision": overall_role_metrics["exact_role_auto_precision"],
            "automatic_acceptance_precision_semantics": "legacy_alias_of_exact_role_auto_precision",
            "automatic_acceptance_count": overall_role_metrics["accepted_count"],
            "automatic_acceptance_true_heading_count": overall_role_metrics["accepted_heading_count"],
            "candidate_review_required_rate": candidate_review / candidate_total if candidate_total else 0.0,
            "pending_decision_rate": pending_total / pending_denominator if pending_denominator else 0.0,
            "correction_operations_estimate": sum(level_errors.values()),
            "manual_correction_count": None,
        },
        "by_category": category_results,
        "limitations": [
            "评估材料为程序生成的结构化虚构样例，不能代表真实学校模板或复杂版式噪声。",
            "自动接受覆盖率与 precision 分开报告；低置信度候选仍要求人工确认。",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dataset", type=Path, default=Path(__file__).parent / "dataset.json")
    parser.add_argument("--samples-dir", type=Path, default=Path(__file__).parent / "samples")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "metrics.json")
    args = parser.parse_args()
    dataset = _load_dataset(args.dataset.resolve())
    metrics = evaluate_dataset(dataset, args.samples_dir.resolve())
    args.output.resolve().write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(metrics["overall"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
