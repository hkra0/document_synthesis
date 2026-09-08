# -*- coding: utf-8 -*-
"""
样本文档格式分析器 (lib/format_analysis.py)
对参考 DOCX 进行只读检查，提取版面几何与有效属性，执行样式聚类、角色候选推导与字段级证据链生成。
绝不泄露样本文档的敏感正文至分析结果的规范字段。
"""

import hashlib
import json
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from .docx_inspector import (
    BlockInspection,
    DocumentInspection,
    DocumentInspectionLimits,
    inspect_docx,
)
from .contracts import CORE_ROLES, SEMANTIC_ROLES
from .format_schema import (
    Diagnostic,
    FormatDiagnosticCode,
    LengthValue,
    LineSpacing,
    PageSpec,
    ParagraphStyle,
    RunStyle,
    StyleDefinition,
)
from .role_candidates import rank_unstructured_candidates

ANALYZER_VERSION = "1.0.0"

STANDARD_CORE_ROLES = list(CORE_ROLES)


@dataclass
class StyleCluster:
    """视觉样式聚类组"""
    cluster_id: str
    run_style: Dict[str, Any]
    paragraph_style: Dict[str, Any]
    structure_type: str
    outline_level: Optional[int]
    p_style_ids: List[str]
    occurrence_count: int
    sample_nodes: List[str]
    sample_texts: List[str]
    suggested_role: str
    confidence: float

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def _compute_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def _extract_cluster_key(block: BlockInspection) -> Tuple:
    """提取用于聚类的归一化特征元组"""
    eff_r = block.effective_run
    eff_p = block.effective_paragraph

    east_asia = (eff_r.east_asia or "").strip().lower() if eff_r else ""
    latin = (eff_r.latin or "").strip().lower() if eff_r else ""
    size_pt = round(eff_r.size_pt, 1) if (eff_r and eff_r.size_pt) else None
    bold = bool(eff_r.bold) if eff_r else False
    italic = bool(eff_r.italic) if eff_r else False
    color = (eff_r.color or "auto").lower() if eff_r else "auto"

    alignment = (eff_p.alignment or "left").lower() if eff_p else "left"
    first_indent_char = None
    if eff_p and eff_p.first_line_indent:
        if eff_p.first_line_indent.unit == "char":
            first_indent_char = round(eff_p.first_line_indent.value, 1)
        elif eff_p.first_line_indent.unit == "pt" and size_pt:
            first_indent_char = round(eff_p.first_line_indent.value / size_pt, 1)

    line_mode = eff_p.line_spacing.mode if (eff_p and eff_p.line_spacing) else "single"
    line_val = round(eff_p.line_spacing.value, 2) if (eff_p and eff_p.line_spacing and eff_p.line_spacing.value) else None

    outline_lvl = block.outline_level
    is_table = block.structure_type == "table" or "/w:tbl[" in block.node.element_path

    return (
        outline_lvl,
        is_table,
        alignment,
        east_asia,
        latin,
        size_pt,
        bold,
        italic,
        color,
        first_indent_char,
        line_mode,
        line_val,
    )


def _serialize_dataclass(value: Any) -> Any:
    """将有效属性转换为稳定、可写入分析契约的普通 JSON 值。"""
    if value is None:
        return None
    if hasattr(value, "__dataclass_fields__"):
        return {key: _serialize_dataclass(getattr(value, key)) for key in value.__dataclass_fields__}
    if isinstance(value, (list, tuple)):
        return [_serialize_dataclass(item) for item in value]
    if isinstance(value, dict):
        return {str(key): _serialize_dataclass(item) for key, item in value.items()}
    return value


def _field_evidence(inspection: DocumentInspection) -> List[Dict[str, Any]]:
    """为每个节点记录有效 run/paragraph 字段及保守的来源状态。

    ``effective`` 是求值后的可观察结果；``source`` 只表示证据层级，
    不把“没有属性”猜成某个 Word 默认值。后续格式编译只消费确认后的
    cluster，不会把这些样本片段写进格式包。
    """
    evidence: List[Dict[str, Any]] = []
    run_fields = ("east_asia", "latin", "complex_script", "size_pt", "bold", "italic", "color")
    para_fields = (
        "alignment", "first_line_indent", "hanging_indent", "left_indent",
        "right_indent", "space_before_pt", "space_after_pt", "line_spacing",
        "keep_with_next", "keep_lines", "widow_control", "page_break_before", "snap_to_grid",
    )
    for block in inspection.blocks:
        for run_index, run in enumerate(block.effective_runs or ((block.effective_run,) if block.effective_run else ())):
            for field_name in run_fields:
                value = getattr(run, field_name, None)
                evidence.append({
                    "node_ref": block.node.element_path,
                    "scope": "run",
                    "run_index": run_index,
                    "field": field_name,
                    "value": _serialize_dataclass(value),
                    "status": "observed" if value is not None else "unverified",
                    "source": "effective_style_cascade" if value is not None else "missing_attribute",
                })
        paragraph = block.effective_paragraph
        if paragraph is None:
            continue
        for field_name in para_fields:
            value = getattr(paragraph, field_name, None)
            evidence.append({
                "node_ref": block.node.element_path,
                "scope": "paragraph",
                "field": field_name,
                "value": _serialize_dataclass(value),
                "status": "observed" if value is not None else "unverified",
                "source": "effective_style_cascade" if value is not None else "missing_attribute",
            })
    return evidence


def _style_inventory(inspection: DocumentInspection) -> Dict[str, Any]:
    """区分 styles.xml 中定义但未使用的样式与实际使用样式。"""
    usage = Counter(block.p_style_id for block in inspection.blocks if block.p_style_id)
    used = []
    defined_only = []
    records = []
    for style_id, style in sorted(inspection.styles.items()):
        count = usage.get(style_id, 0)
        record = {
            "style_id": style_id,
            "name": style.name,
            "style_type": style.style_type,
            "based_on": style.based_on,
            "usage_count": count,
            "used": bool(count),
            "run": _serialize_dataclass(style.run_style),
            "paragraph": _serialize_dataclass(style.paragraph_style),
        }
        records.append(record)
        (used if count else defined_only).append(style_id)
    return {
        "defined_count": len(records),
        "used_count": len(used),
        "defined_only_count": len(defined_only),
        "used_style_ids": used,
        "defined_only_style_ids": defined_only,
        "styles": records,
    }


def _unsupported_features(sample_path: Path, inspection: DocumentInspection) -> List[Dict[str, Any]]:
    """记录分析器不会替用户确认的对象，而不是静默忽略它们。"""
    features: Counter[str] = Counter()
    locations: Dict[str, List[str]] = {}
    try:
        with zipfile.ZipFile(sample_path) as package:
            for part_name in package.namelist():
                if not part_name.startswith("word/") or not part_name.endswith(".xml"):
                    continue
                raw = package.read(part_name)
                for token, feature in (
                    (b"<w:altChunk", "alt_chunk"),
                    (b"<w:object", "ole_object"),
                    (b"<w:txbxContent", "text_box"),
                    (b"<w:customXml", "custom_xml"),
                ):
                    if token in raw:
                        features[feature] += raw.count(token)
                        locations.setdefault(feature, []).append(part_name)
    except (OSError, zipfile.BadZipFile):
        features["unreadable_package"] += 1
        locations["unreadable_package"] = [str(sample_path)]
    return [
        {
            "feature": feature,
            "count": count,
            "locations": sorted(set(locations.get(feature, []))),
            "decision": "review_required",
            "message": "该对象不会被自动映射为语义角色；请人工确认保留或拒绝。",
        }
        for feature, count in sorted(features.items())
    ]


def analyze_format_sample(
    sample_path: Union[str, Path],
    base_format_ref: Optional[str] = None,
    options: Optional[Dict[str, Any]] = None,
    limits: Optional[DocumentInspectionLimits] = None,
) -> Dict[str, Any]:
    """
    深度分析样本文档，聚类有效样式并推导候选语义角色与证据链。
    返回结构完备的分析数据字典。
    """
    path = Path(sample_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"样本文档不存在: {path}")

    inspection = inspect_docx(path, limits=limits)
    source_sha256 = inspection.source_sha256
    report_id = f"sample-analysis-{source_sha256[:8]}"

    # 1. 提取页面几何配置
    page_dict: Dict[str, Any] = {
        "width_mm": 210.0,
        "height_mm": 297.0,
        "orientation": "portrait",
        "margin_top_mm": 25.4,
        "margin_bottom_mm": 25.4,
        "margin_left_mm": 31.8,
        "margin_right_mm": 31.8,
        "header_distance_mm": 15.0,
        "footer_distance_mm": 15.0,
        "snap_to_grid": True,
    }
    if inspection.sections:
        first_sec = inspection.sections[0]
        ps = first_sec.page_spec
        page_dict = {
            "width_mm": ps.width_mm,
            "height_mm": ps.height_mm,
            "orientation": ps.orientation,
            "margin_top_mm": ps.margin_top_mm,
            "margin_bottom_mm": ps.margin_bottom_mm,
            "margin_left_mm": ps.margin_left_mm,
            "margin_right_mm": ps.margin_right_mm,
            "header_distance_mm": ps.header_distance_mm,
            "footer_distance_mm": ps.footer_distance_mm,
            "snap_to_grid": ps.snap_to_grid,
        }

    # 2. 块级节点聚类 (Clustering)
    groups: Dict[Tuple, List[BlockInspection]] = {}
    for block in inspection.blocks:
        key = _extract_cluster_key(block)
        groups.setdefault(key, []).append(block)

    # 排序：按出现次数降序排列
    sorted_groups = sorted(groups.items(), key=lambda item: len(item[1]), reverse=True)

    structural_style_markers = ("heading", "标题", "title", "subtitle")
    has_structural_evidence = any(
        block.outline_level is not None
        or any(marker in (block.p_style_id or "").lower() for marker in structural_style_markers)
        or any(marker in (block.p_style_name or "").lower() for marker in structural_style_markers)
        for block in inspection.blocks
    )
    unstructured_evidence = (
        {} if has_structural_evidence else rank_unstructured_candidates(inspection.blocks)
    )

    clusters: List[StyleCluster] = []
    cluster_id_counter = 0

    for key, blocks in sorted_groups:
        cluster_id_counter += 1
        cid = f"cluster-{cluster_id_counter:02d}"
        first = blocks[0]
        eff_r = first.effective_run
        eff_p = first.effective_paragraph

        r_style_dict: Dict[str, Any] = {}
        if eff_r:
            if eff_r.east_asia:
                r_style_dict["east_asia"] = eff_r.east_asia
            if eff_r.latin:
                r_style_dict["latin"] = eff_r.latin
            if eff_r.size_pt:
                r_style_dict["size_pt"] = eff_r.size_pt
            if eff_r.bold:
                r_style_dict["bold"] = True
            if eff_r.italic:
                r_style_dict["italic"] = True
            if eff_r.color and eff_r.color.lower() != "auto":
                r_style_dict["color"] = eff_r.color.upper()

        p_style_dict: Dict[str, Any] = {}
        if eff_p:
            if eff_p.alignment:
                p_style_dict["alignment"] = eff_p.alignment
            if eff_p.first_line_indent:
                p_style_dict["first_line_indent"] = {
                    "value": eff_p.first_line_indent.value,
                    "unit": eff_p.first_line_indent.unit,
                }
            if eff_p.space_before_pt:
                p_style_dict["space_before_pt"] = eff_p.space_before_pt
            if eff_p.space_after_pt:
                p_style_dict["space_after_pt"] = eff_p.space_after_pt
            if eff_p.line_spacing:
                p_style_dict["line_spacing"] = {
                    "mode": eff_p.line_spacing.mode,
                    "value": eff_p.line_spacing.value,
                }
            if eff_p.keep_with_next:
                p_style_dict["keep_with_next"] = True

        p_styles = sorted(list({b.p_style_id for b in blocks if b.p_style_id}))
        sample_nodes = [b.node.element_path for b in blocks[:5]]
        sample_texts = [b.visible_text[:60].strip() for b in blocks[:3] if b.visible_text.strip()]

        # 角色推导
        outline_lvl = first.outline_level
        is_table = first.structure_type == "table" or "/w:tbl[" in first.node.element_path
        p_style_lower = (first.p_style_id or "").lower()

        suggested_role = "body"
        confidence = 0.70

        if is_table:
            suggested_role = "table.body"
            confidence = 0.95
        elif outline_lvl is not None and 1 <= outline_lvl <= 9:
            suggested_role = f"heading.{outline_lvl}"
            # A direct w:outlineLvl is an explicit structural signal.  Keep
            # style-inherited outline levels reviewable because a sample can
            # inherit an outline level without that style being a heading.
            confidence = 0.99 if first.outline_source == "direct" else 0.95
        elif "title" in p_style_lower and "sub" not in p_style_lower:
            suggested_role = "title"
            confidence = 0.90
        elif "subtitle" in p_style_lower:
            suggested_role = "subtitle"
            confidence = 0.90
        elif any(h in p_style_lower for h in ("heading 1", "heading1", "标题 1", "标题1")):
            suggested_role = "heading.1"
            confidence = 0.90
        elif any(h in p_style_lower for h in ("heading 2", "heading2", "标题 2", "标题2")):
            suggested_role = "heading.2"
            confidence = 0.90
        elif any(h in p_style_lower for h in ("heading 3", "heading3", "标题 3", "标题3")):
            suggested_role = "heading.3"
            confidence = 0.90
        elif first.node.element_path == "/w:document/w:body/w:p[1]" and eff_p and eff_p.alignment == "center" and eff_r and (eff_r.size_pt or 0) >= 16.0:
            suggested_role = "title"
            confidence = 0.85
        elif not has_structural_evidence and blocks:
            evidence = max(
                (unstructured_evidence.get(block.node.element_path, {}) for block in blocks),
                key=lambda item: float(item.get("selected_score", 0.0)),
                default={},
            )
            suggested_role = evidence.get("selected_role", "body")
            confidence = float(evidence.get("selected_score", 0.70))
        else:
            suggested_role = "body"
            confidence = 0.85

        if not has_structural_evidence and suggested_role == "title" and confidence < 0.58:
            suggested_role = "body"

        clusters.append(StyleCluster(
            cluster_id=cid,
            run_style=r_style_dict,
            paragraph_style=p_style_dict,
            structure_type="table" if is_table else first.structure_type,
            outline_level=outline_lvl,
            p_style_ids=p_styles,
            occurrence_count=len(blocks),
            sample_nodes=sample_nodes,
            sample_texts=sample_texts,
            suggested_role=suggested_role,
            confidence=confidence,
        ))

    # 3. 角色候选绑定 (Role Candidates)
    candidate_roles: Dict[str, Dict[str, Any]] = {}
    assigned_clusters: Set[str] = set()

    # 优先分配标题与大纲级别
    for lvl in range(1, 10):
        h_role = f"heading.{lvl}"
        matching = [c for c in clusters if c.suggested_role == h_role and c.cluster_id not in assigned_clusters]
        if matching:
            best = max(matching, key=lambda c: (c.confidence, c.occurrence_count))
            candidate_roles[h_role] = {
                "cluster_id": best.cluster_id,
                "role": h_role,
                "outline_level": lvl,
                "run": best.run_style,
                "paragraph": best.paragraph_style,
                "confidence": best.confidence,
                "evidence_count": best.occurrence_count,
                "confidence_band": "review_required" if best.confidence < 0.98 else "auto_candidate",
                "selection_basis": "outline_level_and_effective_style",
                "candidate_evidence": [
                    unstructured_evidence.get(path, {}) for path in best.sample_nodes
                    if path in unstructured_evidence
                ],
            }
            assigned_clusters.add(best.cluster_id)

    # 分配标题 Title
    title_matches = [c for c in clusters if c.suggested_role == "title" and c.cluster_id not in assigned_clusters]
    if title_matches:
        best_title = max(title_matches, key=lambda c: (c.confidence, c.occurrence_count))
        candidate_roles["title"] = {
            "cluster_id": best_title.cluster_id,
            "role": "title",
            "run": best_title.run_style,
            "paragraph": best_title.paragraph_style,
            "confidence": best_title.confidence,
            "evidence_count": best_title.occurrence_count,
            "confidence_band": "review_required" if best_title.confidence < 0.98 else "auto_candidate",
            "selection_basis": "style_name_or_position_and_effective_style",
            "candidate_evidence": [
                unstructured_evidence.get(path, {}) for path in best_title.sample_nodes
                if path in unstructured_evidence
            ],
        }
        assigned_clusters.add(best_title.cluster_id)

    # 分配表格 Table Body
    tbl_matches = [c for c in clusters if c.suggested_role == "table.body" and c.cluster_id not in assigned_clusters]
    if tbl_matches:
        best_tbl = max(tbl_matches, key=lambda c: c.occurrence_count)
        candidate_roles["table.body"] = {
            "cluster_id": best_tbl.cluster_id,
            "role": "table.body",
            "run": best_tbl.run_style,
            "paragraph": best_tbl.paragraph_style,
            "confidence": best_tbl.confidence,
            "evidence_count": best_tbl.occurrence_count,
            "confidence_band": "review_required" if best_tbl.confidence < 0.98 else "auto_candidate",
            "selection_basis": "table_structure_and_effective_style",
            "candidate_evidence": [
                unstructured_evidence.get(path, {}) for path in best_tbl.sample_nodes
                if path in unstructured_evidence
            ],
        }
        assigned_clusters.add(best_tbl.cluster_id)

    # 分配正文 Body (取剩余中最常出现的)
    body_matches = [c for c in clusters if c.suggested_role == "body" and c.cluster_id not in assigned_clusters]
    if not body_matches and clusters:
        body_matches = [c for c in clusters if c.cluster_id not in assigned_clusters]

    if body_matches:
        best_body = max(body_matches, key=lambda c: c.occurrence_count)
        candidate_roles["body"] = {
            "cluster_id": best_body.cluster_id,
            "role": "body",
            "run": best_body.run_style,
            "paragraph": best_body.paragraph_style,
            "confidence": best_body.confidence,
            "evidence_count": best_body.occurrence_count,
            "confidence_band": "review_required" if best_body.confidence < 0.98 else "auto_candidate",
            "selection_basis": "remaining_occurrence_frequency",
            "candidate_evidence": [
                unstructured_evidence.get(path, {}) for path in best_body.sample_nodes
                if path in unstructured_evidence
            ],
        }
        assigned_clusters.add(best_body.cluster_id)

    # 4. 节点到角色映射 (Node -> Role)
    cluster_role_map = {c.cluster_id: c.suggested_role for c in clusters}
    for role_name, cand in candidate_roles.items():
        cluster_role_map[cand["cluster_id"]] = role_name

    node_roles: Dict[str, str] = {}
    node_refs: Dict[str, Dict[str, str]] = {}
    cluster_lookup = {}
    for c in clusters:
        for n_path in c.sample_nodes:
            cluster_lookup[n_path] = c.cluster_id

    for block in inspection.blocks:
        p_path = block.node.element_path
        node_refs[p_path] = {
            "source_sha256": block.node.source_sha256,
            "part_uri": block.node.part_uri,
            "element_path": p_path,
            "text_hash": block.node.text_hash,
        }
        matched_cid = cluster_lookup.get(p_path)
        if not matched_cid:
            key = _extract_cluster_key(block)
            for idx, (grp_key, _) in enumerate(sorted_groups):
                if key == grp_key:
                    matched_cid = f"cluster-{idx + 1:02d}"
                    break
        role = cluster_role_map.get(matched_cid, "body") if matched_cid else "body"
        node_roles[p_path] = role

    # 5. 缺失角色分析 (Missing Roles)
    missing_roles: List[Dict[str, Any]] = []
    for required_role in STANDARD_CORE_ROLES:
        if required_role not in candidate_roles:
            reason = f"样本文档中未观测到满足 {required_role} 特征的独立段落"
            suggested_base = "body"
            if required_role.startswith("heading."):
                lvl = int(required_role.split(".")[1])
                suggested_base = f"heading.{lvl - 1}" if lvl > 1 and f"heading.{lvl - 1}" in candidate_roles else "body"
            elif required_role == "title":
                suggested_base = "heading.1" if "heading.1" in candidate_roles else "body"
            elif required_role == "table.body":
                suggested_base = "body"

            missing_roles.append({
                "role": required_role,
                "suggested_fallback": f"inherit_from:{suggested_base}",
                "reason": reason,
            })

    candidate_confidences = [float(item.get("confidence", 0.0)) for item in candidate_roles.values()]
    sample_counts = {
        "blocks": len(inspection.blocks),
        "paragraphs": sum(1 for block in inspection.blocks if block.structure_type == "paragraph"),
        "tables": sum(1 for block in inspection.blocks if block.structure_type == "table"),
        "runs": sum(block.runs_count for block in inspection.blocks),
        "headings": sum(1 for block in inspection.blocks if block.outline_level is not None),
        "nodes_with_protected_objects": sum(1 for block in inspection.blocks if block.protected_objects),
    }
    return {
        "analysis_schema_version": 1,
        "report_id": report_id,
        "source_sha256": source_sha256,
        "source_file": path.name,
        "analyzer_version": ANALYZER_VERSION,
        "base_format_ref": base_format_ref,
        "page": page_dict,
        "clusters": [c.to_dict() for c in clusters],
        "candidate_roles": candidate_roles,
        "node_roles": node_roles,
        "node_refs": node_refs,
        "unstructured_candidate_evidence": unstructured_evidence,
        "missing_roles": missing_roles,
        "sample_counts": sample_counts,
        "style_inventory": _style_inventory(inspection),
        "field_evidence": _field_evidence(inspection),
        "unsupported_features": _unsupported_features(path, inspection),
        "confidence_summary": {
            "automatic_accept_threshold": 0.98,
            "candidate_count": len(candidate_confidences),
            "review_required_count": sum(1 for value in candidate_confidences if value < 0.98),
            "automatic_candidate_count": sum(1 for value in candidate_confidences if value >= 0.98),
            "minimum": min(candidate_confidences) if candidate_confidences else None,
            "maximum": max(candidate_confidences) if candidate_confidences else None,
            "note": "置信度只是候选排序证据；低于阈值必须人工确认。",
        },
        "diagnostics": [
            {
                "code": d.code.value if hasattr(d.code, "value") else str(d.code),
                "severity": d.severity,
                "location": d.location,
                "message": d.message,
            }
            for d in inspection.diagnostics
        ],
    }
