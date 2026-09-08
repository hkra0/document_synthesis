"""Pure, label-driven metrics for automatic role decisions.

This module intentionally has no dependency on the candidate scorer.  The
caller supplies the fixed labels, the selected predictions, the nodes actually
accepted automatically, and the candidate roles.  That keeps evaluation
definitions testable and prevents a scorer from manufacturing its own labels.
"""

from __future__ import annotations

from typing import Iterable, Mapping


def _is_heading_role(role: str) -> bool:
    return isinstance(role, str) and (role == "title" or role.startswith("heading."))


def _ratio(numerator: int, denominator: int) -> float | None:
    return numerator / denominator if denominator else None


def evaluate_role_metrics(
    labels: Mapping[str, str],
    predictions: Mapping[str, str],
    accepted: Mapping[str, str] | Iterable[str],
    candidate_roles: Mapping[str, Iterable[str]] | None = None,
) -> dict[str, object]:
    """Compute detection and exact-role metrics from fixed decisions.

    ``accepted`` is the actual set of nodes released without human review.  A
    mapping may be supplied to make the accepted role explicit; an iterable
    uses ``predictions`` for those node IDs.  Every accepted node must have a
    label so an unlabelled prediction cannot silently disappear from a metric
    denominator.
    """

    labels = dict(labels)
    predictions = dict(predictions)
    if isinstance(accepted, Mapping):
        accepted_roles = dict(accepted)
    else:
        accepted_roles = {node: predictions[node] for node in accepted}
    missing_labels = sorted(set(accepted_roles) - set(labels))
    if missing_labels:
        raise ValueError(f"自动接受节点缺少人工标签: {missing_labels[:3]}")

    heading_nodes = {node for node, role in labels.items() if _is_heading_role(role)}
    accepted_nodes = set(accepted_roles)
    accepted_heading_nodes = {
        node
        for node, role in accepted_roles.items()
        if _is_heading_role(role) and _is_heading_role(labels[node])
    }
    exact_nodes = {
        node for node in accepted_nodes if accepted_roles[node] == labels[node]
    }
    candidate_roles = candidate_roles or {}
    candidate_exact_nodes = {
        node
        for node in heading_nodes
        if labels[node] in set(candidate_roles.get(node, ()))
    }

    return {
        "heading_detection_precision": _ratio(
            len(accepted_heading_nodes), len(accepted_nodes)
        ),
        "exact_role_auto_precision": _ratio(len(exact_nodes), len(accepted_nodes)),
        "auto_acceptance_coverage": _ratio(
            len(accepted_heading_nodes), len(heading_nodes)
        ),
        "correct_auto_coverage": _ratio(len(exact_nodes), len(heading_nodes)),
        "candidate_exact_role_coverage": _ratio(
            len(candidate_exact_nodes), len(heading_nodes)
        ),
        "accepted_count": len(accepted_nodes),
        "accepted_heading_count": len(accepted_heading_nodes),
        "exact_role_count": len(exact_nodes),
        "heading_count": len(heading_nodes),
        "candidate_exact_role_count": len(candidate_exact_nodes),
        "body_as_heading_false_positives": sum(
            labels[node] == "body" and _is_heading_role(role)
            for node, role in accepted_roles.items()
        ),
    }
