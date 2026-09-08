# -*- coding: utf-8 -*-
"""Shared JSON contracts, semantic roles, and provenance helpers.

This module deliberately has no dependency on the configuration or rendering
layers.  It is therefore safe to use from the analyser, compiler, resolver,
and CLI without creating import cycles.
"""

from __future__ import annotations

import copy
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable, Mapping, MutableMapping, Optional, Tuple

import jsonschema


SCHEMA_DIR = Path(__file__).resolve().parent.parent / "schemas"
SHA256_RE = re.compile(r"^[0-9a-fA-F]{64}$")


SEMANTIC_ROLES = frozenset(
    {
        "title",
        "subtitle",
        "body",
        "quote",
        "caption",
        "bibliography",
        "table.body",
        "header",
        "footer",
        "toc.title",
        *(f"heading.{level}" for level in range(1, 10)),
        *(f"toc.{level}" for level in range(1, 10)),
    }
)
CORE_ROLES = (
    "title",
    "heading.1",
    "heading.2",
    "heading.3",
    "body",
    "table.body",
)
STRUCTURE_TYPES = frozenset({"paragraph", "table", "section_break", "drawing"})


class ContractError(ValueError):
    """A JSON contract error with a stable JSON Pointer location."""

    def __init__(self, message: str, pointer: str = ""):
        self.pointer = pointer or "/"
        super().__init__(f"{message} (JSON Pointer: {self.pointer})")


def json_pointer(parts: Iterable[Any]) -> str:
    encoded = []
    for part in parts:
        text = str(part).replace("~", "~0").replace("/", "~1")
        encoded.append(text)
    return "/" + "/".join(encoded) if encoded else "/"


def load_contract_schema(name: str) -> Dict[str, Any]:
    path = SCHEMA_DIR / name
    try:
        with path.open("r", encoding="utf-8") as handle:
            schema = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise ContractError(f"无法读取契约模式 {path}: {exc}") from exc
    if not isinstance(schema, dict):
        raise ContractError(f"契约模式 {path} 根节点必须是对象")
    return schema


def validate_contract(instance: Any, schema_name: str, *, context: str = "契约") -> Any:
    """Validate an instance and report the first deterministic JSON Pointer."""

    schema = load_contract_schema(schema_name)
    validator = jsonschema.Draft7Validator(schema)
    errors = sorted(
        validator.iter_errors(instance),
        key=lambda error: (tuple(str(x) for x in error.absolute_path), error.validator),
    )
    if errors:
        error = errors[0]
        raise ContractError(f"{context}无效: {error.message}", json_pointer(error.absolute_path))
    return instance


def is_sha256(value: Any) -> bool:
    return isinstance(value, str) and bool(SHA256_RE.fullmatch(value))


def require_sha256(value: Any, *, pointer: str = "/source_sha256") -> str:
    if not is_sha256(value):
        raise ContractError("source_sha256 必须是 64 位十六进制字符串", pointer)
    return value.lower()


def is_semantic_role(role: Any) -> bool:
    return isinstance(role, str) and role in SEMANTIC_ROLES


def text_hash(text: str) -> str:
    normalized = " ".join(str(text or "").split())
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def merge_dict_with_provenance(
    base: Mapping[str, Any],
    patch: Mapping[str, Any],
    base_provenance: Optional[Mapping[str, str]] = None,
    patch_source_label: str = "override",
    current_path: str = "",
) -> Tuple[Dict[str, Any], Dict[str, str]]:
    """Deep merge config data while preserving the declaring file per leaf.

    Dictionaries merge recursively. Arrays, scalars, null, false and zero are
    values and replace the previous value exactly.  Provenance is stored at
    every value path, including the container path when a whole value is
    replaced.
    """

    result = copy.deepcopy(dict(base))
    provenance = dict(base_provenance or {})

    for key, value in patch.items():
        path = f"{current_path}/{str(key).replace('~', '~0').replace('/', '~1')}"
        old_value = result.get(key)
        if isinstance(value, Mapping) and isinstance(old_value, Mapping):
            merged, nested_provenance = merge_dict_with_provenance(
                old_value,
                value,
                provenance,
                patch_source_label,
                path,
            )
            result[key] = merged
            provenance.update(nested_provenance)
        else:
            result[key] = copy.deepcopy(value)
            for old_path in list(provenance):
                if old_path == path or old_path.startswith(path + "/"):
                    del provenance[old_path]
            provenance.update(provenance_for_value(value, patch_source_label, path))
    return result, provenance


def provenance_for_value(value: Any, source_label: str, current_path: str = "") -> Dict[str, str]:
    result = {current_path or "/": source_label}
    if isinstance(value, Mapping):
        for key, child in value.items():
            path = f"{current_path}/{str(key).replace('~', '~0').replace('/', '~1')}"
            result.update(provenance_for_value(child, source_label, path))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            path = f"{current_path}/{index}"
            result.update(provenance_for_value(child, source_label, path))
    return result


def normalize_decisions(data: Mapping[str, Any]) -> Dict[str, Any]:
    """Normalize pre-R2 decision files at the boundary, then use one shape."""

    if not isinstance(data, Mapping):
        raise ContractError("决策文件根节点必须是对象", "/")
    normalized = copy.deepcopy(dict(data))
    # Existing reports were already deterministic but did not carry the
    # explicit version marker. Keep reading them while all new output is v1.
    normalized.setdefault("decisions_schema_version", 1)
    normalized.setdefault("role_styles", {})
    normalized.setdefault("style_overrides", {})
    normalized.setdefault("missing_roles", {})
    if normalized.get("on_unmapped") == "body":
        normalized["on_unmapped"] = "error"
    return normalized


def validate_analysis_data(data: Mapping[str, Any]) -> Dict[str, Any]:
    if not isinstance(data, Mapping):
        raise ContractError("分析报告根节点必须是对象", "/")
    normalized = dict(data)
    validate_contract(normalized, "analysis-v1.schema.json", context="analysis-v1 报告")
    require_sha256(normalized.get("source_sha256"))
    for role, candidate in normalized.get("candidate_roles", {}).items():
        if not is_semantic_role(role):
            raise ContractError(f"报告包含未知语义角色: {role}", f"/candidate_roles/{role}")
        if candidate.get("role") != role:
            raise ContractError(
                f"候选角色键 {role} 与候选内部 role 不一致: {candidate.get('role')}",
                f"/candidate_roles/{role}/role",
            )
        cluster_ids = {c.get("cluster_id") for c in normalized.get("clusters", [])}
        if candidate.get("cluster_id") not in cluster_ids:
            raise ContractError(
                f"候选角色 {role} 引用了不存在的候选 ID: {candidate.get('cluster_id')}",
                f"/candidate_roles/{role}/cluster_id",
            )
    for path, role in normalized.get("node_roles", {}).items():
        if not is_semantic_role(role):
            raise ContractError(f"节点 {path} 的角色无效: {role}", json_pointer(["node_roles", path]))
        if normalized.get("node_refs") is not None and path not in normalized["node_refs"]:
            raise ContractError(f"节点角色没有对应的完整 NodeRef: {path}", json_pointer(["node_roles", path]))
    for path, node_ref in normalized.get("node_refs", {}).items():
        if node_ref["source_sha256"].lower() != normalized["source_sha256"].lower():
            raise ContractError(
                "node_refs 的源哈希与报告不一致",
                json_pointer(["node_refs", path, "source_sha256"]),
            )
    return normalized


def _validate_review_audit(
    decisions: Mapping[str, Any],
    analysis_data: Optional[Mapping[str, Any]] = None,
) -> None:
    """Validate auditable review operations without treating them as proof of a human review."""
    audit = decisions.get("review_audit")
    if audit is None:
        return
    if audit.get("report_id") != decisions.get("report_id"):
        raise ContractError("review_audit 的 report_id 与决策不一致", "/review_audit/report_id")
    if str(audit.get("source_sha256", "")).lower() != str(decisions.get("source_sha256", "")).lower():
        raise ContractError("review_audit 的源哈希与决策不一致", "/review_audit/source_sha256")
    if analysis_data is not None and audit.get("report_id") != analysis_data.get("report_id"):
        raise ContractError("review_audit 的 report_id 与分析报告不一致", "/review_audit/report_id")
    if analysis_data is not None and str(audit.get("source_sha256", "")).lower() != str(analysis_data.get("source_sha256", "")).lower():
        raise ContractError("review_audit 的源哈希与分析报告不一致", "/review_audit/source_sha256")

    operations = list(audit.get("operations", ()))
    if int(audit.get("operation_count", -1)) != len(operations):
        raise ContractError(
            "review_audit.operation_count 与操作记录数量不一致",
            "/review_audit/operation_count",
        )
    expected_sequences = list(range(1, len(operations) + 1))
    actual_sequences = [item.get("sequence") for item in operations]
    if actual_sequences != expected_sequences:
        raise ContractError(
            "review_audit 操作序号必须从 1 连续递增",
            "/review_audit/operations",
        )
    # A test harness or an automated browser may produce a useful trace, but
    # it must never be serialized as a human attestation.  An imported trace
    # is provenance, not evidence that the reviewer performed the action in
    # this browser UI.
    if bool(audit.get("automation")) and bool(audit.get("reviewer_attested")):
        raise ContractError(
            "自动化 review_audit 不能声明 reviewer_attested=true",
            "/review_audit/reviewer_attested",
        )
    if audit.get("origin") == "test_fixture" and bool(audit.get("reviewer_attested")):
        raise ContractError(
            "test_fixture review_audit 不能声明 reviewer_attested=true",
            "/review_audit/reviewer_attested",
        )
    if bool(audit.get("reviewer_attested")) and audit.get("origin") != "browser_ui":
        raise ContractError(
            "只有 browser_ui review_audit 可以声明 reviewer_attested=true",
            "/review_audit/origin",
        )
    if bool(audit.get("reviewer_attested")) and not operations:
        raise ContractError(
            "存在 reviewer_attested 时至少需要一条操作记录",
            "/review_audit/operations",
        )


def validate_decisions_data(
    data: Mapping[str, Any],
    analysis_data: Optional[Mapping[str, Any]] = None,
) -> Dict[str, Any]:
    normalized = normalize_decisions(data)
    validate_contract(normalized, "decisions-v1.schema.json", context="decisions-v1 决策")
    require_sha256(normalized.get("source_sha256"))
    _validate_review_audit(normalized, analysis_data)
    if analysis_data is not None:
        if normalized.get("report_id") != analysis_data.get("report_id"):
            raise ContractError("决策文件与分析报告的 report_id 不一致", "/report_id")
        if normalized.get("source_sha256", "").lower() != str(analysis_data.get("source_sha256", "")).lower():
            raise ContractError("决策文件与分析报告的源哈希不一致", "/source_sha256")
        cluster_ids = {c.get("cluster_id") for c in analysis_data.get("clusters", [])}
        for role, cluster_id in normalized.get("role_styles", {}).items():
            if not is_semantic_role(role):
                raise ContractError(f"决策包含未知语义角色: {role}", f"/role_styles/{role}")
            if cluster_id not in cluster_ids:
                raise ContractError(
                    f"角色 {role} 引用了不存在的候选 ID: {cluster_id}",
                    f"/role_styles/{role}",
                )
        for role in normalized.get("style_overrides", {}):
            if not is_semantic_role(role):
                raise ContractError(f"样式覆写包含未知语义角色: {role}", f"/style_overrides/{role}")
        for path, role in normalized.get("node_roles", {}).items():
            if not is_semantic_role(role):
                raise ContractError(f"节点 {path} 的角色无效: {role}", f"/node_roles/{path}")

        # New v1 files must make an explicit decision for every role the
        # analyser says is missing. Legacy files receive the historical
        # inherit default only at this boundary.
        missing_roles = {item.get("role") for item in analysis_data.get("missing_roles", [])}
        missing_decisions = normalized.get("missing_roles", {})
        selected_roles = set(normalized.get("role_styles", {}))
        if data.get("decisions_schema_version") == 1:
            # A role initially reported as missing may be resolved during
            # review by assigning an observed cluster to it.  The UI removes
            # that stale missing entry; require a decision only when neither
            # an explicit style assignment nor inherit/reject is present.
            unresolved = sorted(
                role for role in missing_roles
                if role not in missing_decisions and role not in selected_roles
            )
            if unresolved:
                raise ContractError(
                    "缺失角色必须显式选择 inherit 或 reject: " + ", ".join(unresolved),
                    "/missing_roles",
                )
    return normalized


def require_final_decisions(
    decisions_data: Mapping[str, Any],
    analysis_data: Mapping[str, Any],
    *,
    base_format_ref: Optional[str] = None,
    require_role_styles: bool = True,
) -> Dict[str, Any]:
    """Reject draft decisions before a package or mapping is published.

    The HTML report deliberately emits a draft starter.  A caller may inspect
    it, but a compiler must never silently treat a pending choice as a user
    confirmation or silently fall back to every observed cluster.
    """
    normalized = validate_decisions_data(decisions_data, analysis_data)
    if normalized.get("decisions_schema_version") != 1:
        return normalized

    pending = sorted(
        role for role, value in normalized.get("missing_roles", {}).items()
        if value == "pending"
    )
    if pending:
        raise ContractError(
            "缺失角色仍处于待确认状态: " + ", ".join(pending),
            "/missing_roles",
        )

    if require_role_styles and base_format_ref is None:
        observed_roles = set(analysis_data.get("candidate_roles", {}))
        selected_roles = set(normalized.get("role_styles", {}))
        unresolved = sorted(observed_roles - selected_roles)
        if unresolved:
            raise ContractError(
                "观测到的角色必须选择候选样式或明确继承基础格式: " + ", ".join(unresolved),
                "/role_styles",
            )
    return normalized


def validate_role_map_data(data: Mapping[str, Any]) -> Dict[str, Any]:
    if not isinstance(data, Mapping):
        raise ContractError("RoleMap 根节点必须是对象", "/")
    normalized = dict(data)
    validate_contract(normalized, "role-map-v1.schema.json", context="role-map-v1 映射")
    require_sha256(normalized.get("source_sha256"))
    seen = set()
    for index, assignment in enumerate(normalized.get("assignments", [])):
        node_ref = assignment["node_ref"]
        if node_ref["source_sha256"].lower() != normalized["source_sha256"].lower():
            raise ContractError("assignment 的 NodeRef 源哈希与映射不一致", f"/assignments/{index}/node_ref/source_sha256")
        key = (node_ref["part_uri"], node_ref["element_path"])
        if key in seen:
            raise ContractError("RoleMap 不允许重复绑定同一 NodeRef", f"/assignments/{index}/node_ref")
        seen.add(key)
        if not is_semantic_role(assignment["role"]):
            raise ContractError(f"映射包含未知语义角色: {assignment['role']}", f"/assignments/{index}/role")
    return normalized
