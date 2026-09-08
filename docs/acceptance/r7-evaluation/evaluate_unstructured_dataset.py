#!/usr/bin/env python3
"""Evaluate no-outline candidate coverage without claiming probabilities."""

from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from pathlib import Path
import sys

REPO = Path(__file__).resolve().parents[3]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from lib.format_analysis import analyze_format_sample
from lib.role_evaluation import evaluate_role_metrics


THRESHOLD = 0.98


def _is_heading_role(role: str) -> bool:
    return isinstance(role, str) and (role == "title" or role.startswith("heading."))


def _role_metrics(labels: dict[str, str], predictions: dict[str, str], role: str) -> dict[str, object]:
    support = sum(value == role for value in labels.values())
    predicted = sum(value == role for value in predictions.values())
    true_positive = sum(
        labels.get(path) == role and predictions.get(path) == role
        for path in labels
    )
    return {
        "support": support,
        "predicted": predicted,
        "true_positive": true_positive,
        "false_positive": max(predicted - true_positive, 0),
        "false_negative": max(support - true_positive, 0),
        "precision": true_positive / predicted if predicted else None,
        "recall": true_positive / support if support else None,
    }


def evaluate(dataset: dict, samples_dir: Path, baseline: dict | None = None) -> dict:
    per_category = defaultdict(lambda: {
        "documents": 0,
        "labels": {},
        "predictions": {},
        "heading_total": 0,
        "body_total": 0,
        "candidate_heading_hits": 0,
        "candidate_heading_false_positives": 0,
        "candidate_block_count": 0,
        "candidate_true_heading_blocks": 0,
        "accepted": {},
        "candidate_roles": {},
        "correction_operations_estimate": 0,
        "unresolved_count": 0,
    })
    per_level = defaultdict(lambda: Counter())
    all_labels: dict[str, str] = {}
    all_predictions: dict[str, str] = {}
    accepted_predictions: dict[str, str] = {}
    candidate_roles: dict[str, list[str]] = {}
    heading_candidate_hits = 0
    body_selected_as_heading = 0
    body_candidate_as_heading = 0
    candidate_block_count = 0
    candidate_true_heading_blocks = 0
    auto_accept = 0
    auto_accept_correct = 0
    auto_accept_true_heading = 0
    heading_total = 0
    body_total = 0
    correction_operations_estimate = 0
    unresolved_count = 0
    candidate_count_max = 0
    selected_scores = []
    baseline_documents = {
        item["filename"]: item for item in (baseline or {}).get("documents", [])
    }
    for item in dataset["documents"]:
        report = analyze_format_sample(samples_dir / item["filename"])
        evidence = report.get("unstructured_candidate_evidence", {})
        reviewed = baseline_documents.get(item["filename"])
        labels_by_index = (
            {int(index): role for index, role in reviewed["labels"].items()}
            if reviewed else
            {index: f"heading.{position + 1}" for position, index in enumerate(item["heading_paragraph_indices"])}
        )
        headings = {
            f"/w:document/w:body/w:p[{index + 1}]"
            for index, role in labels_by_index.items()
            if role.startswith("heading.")
        }
        bodies = {
            f"/w:document/w:body/w:p[{index + 1}]"
            for index, role in labels_by_index.items()
            if role == "body"
        }
        category = item["category"]
        labels = {
            path: labels_by_index[int(path.rsplit("[", 1)[1][:-1]) - 1]
            for path in (headings | bodies)
        }
        predictions = {
            path: evidence.get(path, {}).get("selected_role", "body")
            for path in labels
        }
        unique_labels = {f"{item['filename']}:{path}": value for path, value in labels.items()}
        unique_predictions = {f"{item['filename']}:{path}": value for path, value in predictions.items()}
        all_labels.update(unique_labels)
        all_predictions.update(unique_predictions)
        category_data = per_category[category]
        category_data["documents"] += 1
        category_data["labels"].update(unique_labels)
        category_data["predictions"].update(unique_predictions)
        for path in labels:
            unique_path = f"{item['filename']}:{path}"
            selected = predictions[path]
            evidence_item = evidence.get(path, {})
            score = float(evidence_item.get("selected_score", 0.0))
            if _is_heading_role(selected) and score >= THRESHOLD:
                accepted_predictions[unique_path] = selected
                category_data["accepted"][unique_path] = selected
            roles = [candidate.get("role") for candidate in evidence_item.get("candidates", [])]
            candidate_roles[unique_path] = [role for role in roles if isinstance(role, str)]
            category_data["candidate_roles"][unique_path] = candidate_roles[unique_path]
        for path in labels:
            is_heading = labels[path].startswith("heading.")
            candidates = evidence.get(path, {}).get("candidates", [])
            candidate_count_max = max(candidate_count_max, len(candidates))
            candidate_heading = any(_is_heading_role(candidate.get("role")) for candidate in candidates)
            selected = predictions[path]
            selected_heading = _is_heading_role(selected)
            score = float(evidence.get(path, {}).get("selected_score", 0.0))
            selected_scores.append(score)
            if candidate_heading:
                candidate_block_count += 1
                category_data["candidate_block_count"] += 1
                if is_heading:
                    candidate_true_heading_blocks += 1
                    category_data["candidate_true_heading_blocks"] += 1
            if selected != labels[path]:
                correction_operations_estimate += 1
                unresolved_count += 1
                category_data["correction_operations_estimate"] += 1
                category_data["unresolved_count"] += 1
            if is_heading:
                heading_total += 1
                category_data["heading_total"] += 1
                level = labels[path]
                level_data = per_level[level]
                level_data["support"] += 1
                if candidate_heading:
                    heading_candidate_hits += 1
                    category_data["candidate_heading_hits"] += 1
                    level_data["candidate_hits"] += 1
                if selected == level:
                    level_data["selected_exact"] += 1
                if selected_heading:
                    level_data["selected_heading"] += 1
                if selected_heading and score >= THRESHOLD:
                    auto_accept += 1
                    auto_accept_correct += 1
                    auto_accept_true_heading += 1
            else:
                body_total += 1
                category_data["body_total"] += 1
                if candidate_heading:
                    body_candidate_as_heading += 1
                    category_data["candidate_heading_false_positives"] += 1
                if selected_heading:
                    body_selected_as_heading += 1
                if selected_heading and score >= THRESHOLD:
                    auto_accept += 1
    overall_role_metrics = evaluate_role_metrics(
        all_labels, all_predictions, accepted_predictions, candidate_roles
    )
    per_level_result = {}
    for role in sorted(per_level):
        data = per_level[role]
        support = int(data["support"])
        candidate_hits = int(data["candidate_hits"])
        selected_exact = int(data["selected_exact"])
        per_level_result[role] = {
            "support": support,
            "candidate_hits": candidate_hits,
            "candidate_coverage": candidate_hits / support if support else 0.0,
            "selected_exact": selected_exact,
            "selected_exact_recall": selected_exact / support if support else 0.0,
            "selected_as_heading": int(data["selected_heading"]),
        }
    category_results = {}
    for category, raw in sorted(per_category.items()):
        labels = raw["labels"]
        predictions = raw["predictions"]
        category_metrics = evaluate_role_metrics(
            labels,
            predictions,
            raw["accepted"],
            raw["candidate_roles"],
        )
        category_results[category] = {
            "documents": raw["documents"],
            "heading_total": raw["heading_total"],
            "body_total": raw["body_total"],
            "candidate_heading_coverage": raw["candidate_heading_hits"] / raw["heading_total"] if raw["heading_total"] else 0.0,
            "candidate_block_precision": raw["candidate_true_heading_blocks"] / raw["candidate_block_count"] if raw["candidate_block_count"] else None,
            "candidate_heading_false_positives": raw["candidate_heading_false_positives"],
            "body_as_heading_false_positives": sum(
                labels[path] == "body" and _is_heading_role(predictions[path])
                for path in labels
            ),
            "correction_operations_estimate": raw["correction_operations_estimate"],
            "unresolved_count": raw["unresolved_count"],
            "heading_detection_precision": category_metrics["heading_detection_precision"],
            "exact_role_auto_precision": category_metrics["exact_role_auto_precision"],
            "auto_acceptance_coverage": category_metrics["auto_acceptance_coverage"],
            "correct_auto_coverage": category_metrics["correct_auto_coverage"],
            "candidate_exact_role_coverage": category_metrics["candidate_exact_role_coverage"],
            "per_level": {
                role: _role_metrics(labels, predictions, role)
                for role in sorted({value for value in labels.values() if value.startswith("heading.")})
            },
        }
    return {
        "metrics_schema_version": 2,
        "dataset_id": dataset["dataset_id"],
        "label_baseline": {
            "baseline_id": (baseline or {}).get("baseline_id"),
            "dataset_sha256": (baseline or {}).get("dataset_sha256"),
            "label_source": (baseline or {}).get("label_source", "dataset_fixture_manifest"),
            "algorithm_output_used": (baseline or {}).get("algorithm_output_used"),
        },
        "threshold": THRESHOLD,
        "dataset_counts": {
            "documents": dataset["document_count"],
            "categories": len(dataset["categories"]),
            "heading_positives": heading_total,
            "body_negatives": body_total,
        },
        "overall": {
            "heading_candidate_coverage": heading_candidate_hits / heading_total if heading_total else 0.0,
            "per_level": per_level_result,
            "body_as_heading_false_positives": body_selected_as_heading,
            "candidate_heading_false_positives": body_candidate_as_heading,
            "candidate_block_precision": candidate_true_heading_blocks / candidate_block_count if candidate_block_count else None,
            "candidate_count_max": candidate_count_max,
            "heading_detection_precision": overall_role_metrics["heading_detection_precision"],
            "exact_role_auto_precision": overall_role_metrics["exact_role_auto_precision"],
            "automatic_acceptance_coverage": overall_role_metrics["auto_acceptance_coverage"],
            "correct_auto_coverage": overall_role_metrics["correct_auto_coverage"],
            "candidate_exact_role_coverage": overall_role_metrics["candidate_exact_role_coverage"],
            "automatic_acceptance_precision": overall_role_metrics["exact_role_auto_precision"],
            "automatic_acceptance_precision_semantics": "legacy_alias_of_exact_role_auto_precision",
            "automatic_acceptance_count": len(accepted_predictions),
            "automatic_acceptance_true_heading_count": overall_role_metrics["accepted_heading_count"],
            "automatic_acceptance_threshold": THRESHOLD,
            "selected_score_range": {
                "min": min(selected_scores) if selected_scores else None,
                "max": max(selected_scores) if selected_scores else None,
            },
            "automatic_threshold_reachable": any(score >= THRESHOLD for score in selected_scores),
            "candidate_count_limit": 3,
            "candidate_count_within_limit": candidate_count_max <= 3,
            "correction_operations_estimate": correction_operations_estimate,
            "manual_correction_count": None,
            "unresolved_count": unresolved_count,
        },
        "by_category": category_results,
        "limitations": [
            "材料为独立设计的匿名虚构无大纲样例，不能外推为真实学校模板准确率。",
            "候选覆盖率表示前三名角色中出现标题候选，不等于最终角色正确率。",
            "correction_operations_estimate 是达到数据标签所需的理论纠正数，不是实际人工点击数。",
            "固定标签来自独立人工可见段落复核，候选排序器未参与标签决定；manual_correction_count 仍为 null，因为当前证据没有真实人工操作日志。" if baseline else
            "manual_correction_count 为 null：无真实人工校正日志时不报告人工一致性为通过。",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dataset", type=Path, default=Path(__file__).parent / "unstructured-dataset.json")
    parser.add_argument("--samples-dir", type=Path, default=Path(__file__).parent / "unstructured-samples")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "unstructured-metrics.json")
    parser.add_argument("--baseline", type=Path, default=Path(__file__).parent / "unstructured-human-baseline.json")
    args = parser.parse_args()
    data = json.loads(args.dataset.resolve().read_text(encoding="utf-8"))
    baseline = json.loads(args.baseline.resolve().read_text(encoding="utf-8")) if args.baseline.resolve().is_file() else None
    if baseline and baseline.get("dataset_id") != data.get("dataset_id"):
        raise SystemExit("baseline dataset_id does not match dataset")
    metrics = evaluate(data, args.samples_dir.resolve(), baseline=baseline)
    args.output.resolve().write_text(json.dumps(metrics, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(metrics["overall"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
