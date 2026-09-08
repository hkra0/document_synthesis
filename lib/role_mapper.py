# -*- coding: utf-8 -*-
"""
确定性语义角色映射器 (lib/role_mapper.py)
将输入 DOCX 的段落与表格等块级元素，基于唯一 NodeRef 确定性映射至语义角色。
消除传统正则匹配因重复同名标题造成的书签错绑。
"""

import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from .docx_inspector import (
    BlockInspection,
    DocumentInspectionLimits,
    DocumentInspection,
    NodeRef,
    inspect_docx,
)
from .format_schema import Diagnostic, FormatDiagnosticCode
from .contracts import SEMANTIC_ROLES, text_hash, validate_role_map_data

VALID_ROLES: Set[str] = set(SEMANTIC_ROLES)


@dataclass(frozen=True)
class RoleAssignment:
    """语义角色分配项 (与确定的 NodeRef 绑定)"""
    node_ref: NodeRef
    role: str
    level: Optional[int] = None
    bookmark_name: Optional[str] = None
    title_text: Optional[str] = None
    provenance: str = "auto"  # "explicit" | "outline_level" | "style" | "highlight" | "default"
    confirmed: bool = False
    text_hash: Optional[str] = None
    story_type: Optional[str] = None

    def __post_init__(self):
        if self.role not in VALID_ROLES:
            raise ValueError(f"无效的语义角色: {self.role}，允许的角色集合: {VALID_ROLES}")

    def to_dict(self) -> Dict[str, Any]:
        """Return the single canonical role-map assignment representation."""
        result = {
            "node_ref": {
                "source_sha256": self.node_ref.source_sha256,
                "part_uri": self.node_ref.part_uri,
                "element_path": self.node_ref.element_path,
                "text_hash": self.text_hash or self.node_ref.text_hash or text_hash(self.title_text or ""),
            },
            "role": self.role,
            "provenance": self.provenance,
            "confirmed": bool(self.confirmed),
            "level": self.level,
            "bookmark_name": self.bookmark_name,
        }
        if self.story_type:
            result["story_type"] = self.story_type
        return result


def serialize_role_map(
    assignments: List[RoleAssignment],
    *,
    source_file: str,
    analyzer_version: str = "1.0.0",
    on_unmapped: str = "error",
) -> Dict[str, Any]:
    """Serialize assignments using the canonical role-map-v1 contract."""
    if on_unmapped not in {"preserve", "error"}:
        raise ValueError("on_unmapped 只支持 preserve 或 error")
    source_hashes = {assignment.node_ref.source_sha256 for assignment in assignments}
    if len(source_hashes) > 1:
        raise ValueError("RoleMap 的 assignments 必须来自同一源文件")
    source_sha256 = next(iter(source_hashes), "")
    mapping = {
        "mapping_schema_version": 1,
        "source_sha256": source_sha256,
        "source_file": source_file,
        "analyzer_version": analyzer_version,
        "assignments": [assignment.to_dict() for assignment in assignments],
        "on_unmapped": on_unmapped,
    }
    validate_role_map_data(mapping)
    return mapping


def validate_role_map(
    assignments: List[RoleAssignment],
    inspection_result: Optional[DocumentInspection] = None,
) -> List[Diagnostic]:
    """校验角色映射列表的完备性与唯一性"""
    diags: List[Diagnostic] = []

    # 1. 检查书签唯一性
    seen_bookmarks: Set[str] = set()
    for a in assignments:
        if a.bookmark_name:
            if a.bookmark_name in seen_bookmarks:
                diags.append(Diagnostic(
                    code=FormatDiagnosticCode.ROLE_UNRESOLVED,
                    severity="error",
                    location=a.node_ref.element_path,
                    message=f"重复的书签名称: {a.bookmark_name}",
                ))
            seen_bookmarks.add(a.bookmark_name)

    # 2. 如果提供了检查结果，核对 NodeRef 是否存在
    if inspection_result is not None:
        for a in assignments:
            if a.node_ref.source_sha256 != inspection_result.source_sha256:
                diags.append(Diagnostic(
                    code=FormatDiagnosticCode.MAPPING_STALE,
                    severity="error",
                    location=a.node_ref.element_path,
                    message="NodeRef 的源文件哈希与当前检查结果不一致，映射已过期。",
                ))
        block_by_key = {
            (b.node.part_uri, b.node.element_path): b
            for b in inspection_result.blocks
        }
        for a in assignments:
            block = block_by_key.get((a.node_ref.part_uri, a.node_ref.element_path))
            if block is None:
                diags.append(Diagnostic(
                    code=FormatDiagnosticCode.ROLE_UNRESOLVED,
                    severity="error",
                    location=a.node_ref.element_path,
                    message=f"NodeRef 指向的节点在源文档中不存在: {a.node_ref.element_path}",
                ))
                continue
            expected_hash = block.node.text_hash or text_hash(block.visible_text)
            actual_hash = a.text_hash or a.node_ref.text_hash
            if actual_hash and actual_hash != expected_hash:
                diags.append(Diagnostic(
                    code=FormatDiagnosticCode.MAPPING_STALE,
                    severity="error",
                    location=a.node_ref.element_path,
                    message="NodeRef 的文本哈希与当前源节点不一致，映射已过期。",
                ))

    return diags


class RoleMapper:
    """语义角色映射引擎"""

    def __init__(
        self,
        mode: str = "auto",
        explicit_rules: Optional[Dict[str, str]] = None,
    ):
        self.mode = mode
        self.explicit_rules = explicit_rules or {}

    def serialize(
        self,
        assignments: List[RoleAssignment],
        *,
        source_file: str,
        analyzer_version: str = "1.0.0",
        on_unmapped: str = "error",
    ) -> Dict[str, Any]:
        """Serialize mapper output using the canonical role-map-v1 shape."""
        return serialize_role_map(
            assignments,
            source_file=source_file,
            analyzer_version=analyzer_version,
            on_unmapped=on_unmapped,
        )

    def map_document(
        self,
        doc_path: Union[str, Path],
        limits: Optional[DocumentInspectionLimits] = None,
    ) -> Tuple[List[RoleAssignment], DocumentInspection]:
        """对目标文档进行只读分析，并输出确定性的角色分配"""
        result = inspect_docx(doc_path, limits)
        assignments = self.map_inspection(result)
        return assignments, result

    def map_inspection(
        self,
        result: DocumentInspection,
    ) -> List[RoleAssignment]:
        """从已完成的只读检查结果生成映射，避免 plan/render 重复扫描源文档。"""
        assignments: List[RoleAssignment] = []
        heading_counter = 0

        for block in result.blocks:
            path_key = block.node.element_path

            # 1. 显式规则优先匹配
            if path_key in self.explicit_rules:
                role = self.explicit_rules[path_key]
                level = int(role.split(".")[1]) if role.startswith("heading.") else None
                bm_name = None
                if role.startswith("heading.") or role == "title":
                    heading_counter += 1
                    bm_name = f"_Toc_{heading_counter:04d}_{hashlib.md5(path_key.encode()).hexdigest()[:6]}"
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role=role,
                    level=level,
                    bookmark_name=bm_name,
                    title_text=block.visible_text.strip() if bm_name else None,
                    provenance="explicit",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))
                continue

            # 2. 表格
            if block.structure_type == "table" or "/w:tbl[" in block.node.element_path:
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role="table.body",
                    provenance="block_type",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))
                continue

            # 3. 启发式：基于大纲级别 (outline_level) 或样式
            outline_lvl = block.outline_level
            style_id = (block.p_style_id or "").lower()
            style_name = (block.p_style_name or "").lower()

            if outline_lvl is None:
                m = re.search(r"(?:heading|标题)\s*([1-9])", style_id) or re.search(r"(?:heading|标题)\s*([1-9])", style_name)
                if m:
                    outline_lvl = int(m.group(1))

            if outline_lvl is not None and 1 <= outline_lvl <= 9:
                heading_counter += 1
                bm_name = f"_Toc_{heading_counter:04d}_{hashlib.md5(path_key.encode()).hexdigest()[:6]}"
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role=f"heading.{outline_lvl}",
                    level=outline_lvl,
                    bookmark_name=bm_name,
                    title_text=block.visible_text.strip(),
                    provenance="outline_level",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))
            elif "title" in style_id and "sub" not in style_id:
                heading_counter += 1
                bm_name = f"_Toc_{heading_counter:04d}_{hashlib.md5(path_key.encode()).hexdigest()[:6]}"
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role="title",
                    bookmark_name=bm_name,
                    title_text=block.visible_text.strip(),
                    provenance="style",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))
            elif "subtitle" in style_id:
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role="subtitle",
                    provenance="style",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))
            else:
                assignments.append(RoleAssignment(
                    node_ref=block.node,
                    role="body",
                    provenance="default",
                    text_hash=block.node.text_hash or text_hash(block.visible_text),
                    story_type=block.story_type,
                ))

        return assignments


def load_role_map(
    map_path: Union[str, Path],
    inspection_result: DocumentInspection,
) -> Tuple[List[RoleAssignment], Dict[str, Any]]:
    """载入并绑定 canonical RoleMap；任何过期或未知 NodeRef 都在渲染前拒绝。"""
    path = Path(map_path).resolve()
    if not path.is_file():
        raise ValueError(f"RoleMap 文件不存在: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取 RoleMap: {path}: {exc}") from exc
    try:
        normalized = validate_role_map_data(data)
    except Exception as exc:
        raise ValueError(f"RoleMap 契约校验失败: {exc}") from exc

    if normalized["source_sha256"].lower() != inspection_result.source_sha256.lower():
        raise ValueError("RoleMap 的源文件哈希与当前源文档不一致，映射已过期。")

    blocks = {
        (block.node.part_uri, block.node.element_path): block
        for block in inspection_result.blocks
    }
    assignments: List[RoleAssignment] = []
    seen = set()
    heading_counter = 0
    for index, item in enumerate(normalized.get("assignments", [])):
        node_ref_data = item["node_ref"]
        key = (node_ref_data["part_uri"], node_ref_data["element_path"])
        block = blocks.get(key)
        if block is None:
            raise ValueError(f"RoleMap[{index}] 的 NodeRef 在当前源文档中不存在: {key[1]}")
        if key in seen:
            raise ValueError(f"RoleMap 包含重复 NodeRef: {key[1]}")
        seen.add(key)
        expected_hash = block.node.text_hash or text_hash(block.visible_text)
        if node_ref_data["text_hash"].lower() != expected_hash.lower():
            raise ValueError(f"RoleMap[{index}] 的文本哈希与源节点不一致，映射已过期: {key[1]}")
        declared_story = item.get("story_type")
        if declared_story and declared_story != block.story_type:
            raise ValueError(
                f"RoleMap[{index}] 的 story_type 与源节点不一致: {key[1]}"
            )
        if item.get("level") is not None and not item["role"].startswith("heading."):
            raise ValueError(f"RoleMap[{index}] 只有 heading 角色可以声明 level: {key[1]}")

        role = item["role"]
        bookmark_name = item.get("bookmark_name")
        if (role.startswith("heading.") or role == "title") and not bookmark_name:
            heading_counter += 1
            bookmark_name = f"_Toc_{heading_counter:04d}_{hashlib.md5(key[1].encode()).hexdigest()[:6]}"
        level = item.get("level")
        if level is None and role.startswith("heading."):
            level = int(role.split(".", 1)[1])
        assignments.append(RoleAssignment(
            node_ref=block.node,
            role=role,
            level=level,
            bookmark_name=bookmark_name,
            title_text=block.visible_text.strip() if (role.startswith("heading.") or role == "title") else None,
            provenance=item.get("provenance", "explicit"),
            confirmed=bool(item.get("confirmed", False)),
            text_hash=expected_hash,
            story_type=block.story_type,
        ))

    if normalized.get("on_unmapped") == "error":
        missing = [
            block.node.element_path
            for block in inspection_result.blocks
            if (block.node.part_uri, block.node.element_path) not in seen
        ]
        if missing:
            raise ValueError(
                "RoleMap 的 on_unmapped=error 但未覆盖源文档节点: "
                + ", ".join(missing[:8])
                + (" ..." if len(missing) > 8 else "")
            )

    diagnostics = validate_role_map(assignments, inspection_result)
    errors = [d.message for d in diagnostics if d.severity == "error"]
    if errors:
        raise ValueError("RoleMap 节点校验失败: " + "; ".join(errors))
    return assignments, normalized
