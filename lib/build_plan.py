# -*- coding: utf-8 -*-
"""一次准备、多阶段消费的只读构建计划。

``PreparedBuild`` 把配置解析、源文档检查、NodeRef、角色映射和交付声明
固定在同一个不可替换的准备结果上。其运行时字段可以保留文档配置对象，
但 ``to_public_dict`` 只返回不含源正文的可审阅计划。
"""

from copy import deepcopy
from dataclasses import asdict, dataclass, field, is_dataclass
import hashlib
import json
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Tuple

from .docx_inspector import DocumentInspection, inspect_docx
from .role_mapper import RoleAssignment, RoleMapper, load_role_map
from .scanner import flatten_tree_nodes
from .source_strategies import (
    SourceStrategyError,
    build_outline,
    resolve_docx_document_path,
)


def compute_sha256(path: Path) -> str:
    if not path.is_file():
        return ""
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(65536):
            digest.update(chunk)
    return digest.hexdigest()


def _json_safe(value: Any) -> Any:
    """把配置的有效值转换为稳定、可哈希的 JSON 结构。"""
    if is_dataclass(value):
        return _json_safe(asdict(value))
    if isinstance(value, Enum):
        return _json_safe(value.value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, Mapping):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    if isinstance(value, set):
        return sorted(_json_safe(item) for item in value)
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def config_contract_snapshot(config: Any) -> Dict[str, Any]:
    """返回构建会消费的配置快照，不依赖外层 dataclass 的浅冻结。"""
    resolved = getattr(config, "resolved_format", None)
    return _json_safe({
        "schema_version": getattr(config, "schema_version", None),
        "project_name": getattr(config, "project_name", None),
        "raw": getattr(config, "raw", {}),
        "source": getattr(config, "source", {}),
        "cover": getattr(config, "cover", {}),
        "cover_spec": getattr(config, "cover_spec", None),
        "formatting": getattr(config, "formatting", None),
        "parts": getattr(config, "parts", {}),
        "regions": getattr(config, "regions", {}),
        "page_sequences": getattr(config, "page_sequences", {}),
        "layout": getattr(config, "layout", {}),
        "documents": getattr(config, "documents", []),
        "node_overrides": getattr(config, "node_overrides", {}),
        "tree_order": getattr(config, "tree_order", []),
        "ignore_nodes": getattr(config, "ignore_nodes", []),
        "resolved_format": {
            "id": getattr(resolved, "id", None),
            "version": getattr(resolved, "version", None),
            "content_hash": getattr(resolved, "content_hash", ""),
            "source_hashes": getattr(resolved, "source_hashes", {}),
        },
    })


def compute_config_contract_hash(config: Any) -> str:
    payload = json.dumps(
        config_contract_snapshot(config),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _node_key(source_sha256: str, part_uri: str, element_path: str) -> str:
    return json.dumps(
        [source_sha256, part_uri, element_path],
        ensure_ascii=False,
        separators=(",", ":"),
    )


def _diagnostic_dict(diagnostic: Any) -> Dict[str, Any]:
    code = getattr(diagnostic, "code", "UNKNOWN")
    return {
        "code": getattr(code, "value", code),
        "severity": getattr(diagnostic, "severity", "warning"),
        "location": getattr(diagnostic, "location", ""),
        "message": getattr(diagnostic, "message", str(diagnostic)),
    }


def _part_dict(part: Any) -> Dict[str, Any]:
    if isinstance(part, Mapping):
        return deepcopy(dict(part))
    data = asdict(part)
    # Enum values are used in a few downstream JSON consumers.
    for key, value in list(data.items()):
        data[key] = getattr(value, "value", value)
    return data


def _without_source_body(value: Any) -> Any:
    """Remove source prose from the external plan while retaining structure."""
    if isinstance(value, list):
        return [_without_source_body(item) for item in value]
    if isinstance(value, tuple):
        return [_without_source_body(item) for item in value]
    if isinstance(value, dict):
        hidden = {"raw_text", "source_text", "visible_text", "text_hash", "title_text"}
        return {
            key: _without_source_body(item)
            for key, item in value.items()
            if key not in hidden
        }
    return value


def _source_files(source_path: Path, strategy: str, target_file: Optional[Path]) -> List[Path]:
    if source_path.is_file():
        return [target_file or source_path]
    if strategy == "docx_document" and target_file is not None:
        # An explicit docx_document source is a single semantic source.  Other
        # DOCX files beside it may be prior evidence or generated artifacts and
        # must not enter the content-integrity inventory for this delivery.
        return [target_file.resolve()]
    files = [
        path.resolve()
        for path in source_path.rglob("*")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if target_file and target_file.is_file() and target_file.resolve() not in files:
        files.append(target_file.resolve())
    return sorted(set(files))


@dataclass(frozen=True)
class PreparedBuild:
    """所有后续阶段共享的只读准备结果。"""

    source_path: Path
    config: Any = field(repr=False, compare=False)
    config_hash: str = ""
    config_provenance: Dict[str, str] = field(default_factory=dict)
    source_hashes: Dict[str, str] = field(default_factory=dict)
    format_hash: str = ""
    format_source_hashes: Dict[str, str] = field(default_factory=dict)
    cover_template_path: Optional[Path] = None
    cover_template_hash: str = ""
    role_map_hash: Optional[str] = None
    role_map_path: Optional[Path] = None
    source_order: Tuple[str, ...] = ()
    nodes: Tuple[Dict[str, Any], ...] = ()
    node_index: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    selection_plan: Dict[str, Any] = field(default_factory=dict)
    assignments: Tuple[RoleAssignment, ...] = ()
    inspections: Tuple[DocumentInspection, ...] = ()
    parts: Dict[str, Any] = field(default_factory=dict)
    sections: Tuple[Dict[str, Any], ...] = ()
    deliveries: Tuple[Dict[str, Any], ...] = ()
    diagnostics: Tuple[Dict[str, Any], ...] = ()
    strategy: str = "directory_tree"
    final_tree: Tuple[Dict[str, Any], ...] = ()
    source_docx_path: Optional[Path] = None
    warnings: Tuple[str, ...] = ()
    counts: Dict[str, int] = field(default_factory=dict)

    def verify_inputs_unchanged(self) -> None:
        """在创建 run_dir 前确认准备阶段的所有输入仍是同一份材料。"""
        if self.config_hash and compute_config_contract_hash(self.config) != self.config_hash:
            raise ValueError("配置自准备后已发生变更 (契约哈希不匹配)。")
        for file_name, expected in self.source_hashes.items():
            path = Path(file_name)
            if not path.is_file():
                raise ValueError(f"源文件自准备后已丢失: {path.name}")
            if compute_sha256(path) != expected:
                raise ValueError(f"源文件自准备后已发生变更 (哈希不匹配): {path.name}")
        if self.role_map_path and self.role_map_hash is not None:
            if not self.role_map_path.is_file():
                raise ValueError(f"RoleMap 自准备后已丢失: {self.role_map_path.name}")
            if compute_sha256(self.role_map_path) != self.role_map_hash:
                raise ValueError(f"RoleMap 自准备后已发生变更 (哈希不匹配): {self.role_map_path.name}")
        for file_name, expected in self.format_source_hashes.items():
            path = Path(file_name)
            if not path.is_file():
                raise ValueError(f"格式包自准备后已丢失: {path.name}")
            if compute_sha256(path) != expected:
                raise ValueError(f"格式包自准备后已发生变更 (哈希不匹配): {path.name}")
        if self.cover_template_path and self.cover_template_hash:
            if not self.cover_template_path.is_file():
                raise ValueError(f"封面模板自准备后已丢失: {self.cover_template_path.name}")
            if compute_sha256(self.cover_template_path) != self.cover_template_hash:
                raise ValueError(f"封面模板自准备后已发生变更 (哈希不匹配): {self.cover_template_path.name}")

    def to_public_dict(self) -> Dict[str, Any]:
        """兼容旧 plan() 返回值，同时剥离运行时对象和源正文。"""
        public_nodes = _without_source_body(list(self.nodes))
        public_assignments = [assignment.to_dict() for assignment in self.assignments]
        public_parts = {key: _part_dict(value) for key, value in self.parts.items()}
        template_path = str(self.cover_template_path) if self.cover_template_path else None
        data = {
            "documents": deepcopy(list(self.deliveries)),
            "manifest_path": str(getattr(self.config, "manifest_path", "")) or None,
            "override_paths": [str(path) for path in getattr(self.config, "override_paths", [])],
            "warnings": list(self.warnings),
            "parts": public_parts,
            "project": getattr(self.config, "project_name", self.source_path.stem),
            "source_dir": str(self.source_path),
            "source_docx_path": str(self.source_docx_path) if self.source_docx_path else None,
            "config_loaded": getattr(self.config, "manifest_path", None) is not None,
            "config_provenance": dict(self.config_provenance),
            "config_contract_sha256": self.config_hash,
            "template": template_path,
            "root_count": len(self.final_tree),
            "node_count": len(self.nodes),
            "nodes": public_nodes,
            "content_block_count": self.counts.get("content_block_count", 0),
            "heading_count": self.counts.get("heading_count", 0),
            "source_hashes": dict(self.source_hashes),
            "format_hash": self.format_hash,
            "format_source_hashes": dict(self.format_source_hashes),
            "cover": {
                **(getattr(getattr(self.config, "cover_spec", None), "to_dict", lambda: {})() or {}),
                "template_sha256": self.cover_template_hash or None,
            },
            "role_map_hash": self.role_map_hash,
            "role_map_path": str(self.role_map_path) if self.role_map_path else None,
            "source_order": list(self.source_order),
            "node_index": deepcopy(self.node_index),
            "selection_plan": deepcopy(self.selection_plan),
            "role_assignments": public_assignments,
            "sections": deepcopy(list(self.sections)),
            "diagnostics": deepcopy(list(self.diagnostics)),
            "strategy": self.strategy,
            "legacy_doc_count": self.counts.get("legacy_doc_count", 0),
            "pptx_count": self.counts.get("pptx_count", 0),
            "source_body_omitted": True,
        }
        return data


@dataclass(frozen=True)
class RenderedBody:
    """渲染阶段唯一输出；不把 Document 句柄放进准备计划或序列化结构。"""

    path: Path
    nodes: Tuple[Dict[str, Any], ...]


def prepare_project_build(source_path: Path, config: Any) -> PreparedBuild:
    """只读准备项目；不创建输出目录、不创建 run_dir、不修改输入。"""
    source_path = Path(source_path).resolve()
    strategy = config.source.get(
        "strategy",
        "highlighted_docx" if source_path.is_file() else "directory_tree",
    )
    configured_role_map = getattr(getattr(config, "formatting", None), "role_map", None)
    if configured_role_map and strategy != "docx_document":
        raise SourceStrategyError(
            "formatting.role_map 当前只能与 docx_document 策略配合使用；"
            "多来源/高亮来源必须先完成 R4/R5 的范围映射。"
        )
    target_file: Optional[Path] = None
    inspections: List[DocumentInspection] = []
    assignments: List[RoleAssignment] = []
    role_map_path: Optional[Path] = None
    role_map_hash: Optional[str] = None

    if strategy == "docx_document":
        if source_path.is_file() and source_path.suffix.lower() != ".docx":
            raise SourceStrategyError(f"docx_document 仅支持 DOCX 文件: {source_path.name}")
        target_file = resolve_docx_document_path(source_path, config)
        inspection = inspect_docx(target_file)
        inspections.append(inspection)
        if configured_role_map:
            role_map_path = Path(configured_role_map).resolve()
            role_map_hash = compute_sha256(role_map_path)
            if not role_map_hash:
                raise SourceStrategyError(f"RoleMap 文件不存在: {role_map_path}")
            assignments, _ = load_role_map(role_map_path, inspection)
        else:
            assignments = RoleMapper().map_inspection(inspection)
        final_tree = build_outline(
            source_path,
            config,
            assignments=assignments,
            inspection=inspection,
        )
    elif strategy == "highlighted_docx":
        if not source_path.is_file() or source_path.suffix.lower() != ".docx":
            raise SourceStrategyError(f"highlighted_docx 仅支持 DOCX 文件: {source_path}")
        from .highlighted_docx import extract_and_format_highlighted_headings

        target_file = source_path
        inspection = inspect_docx(target_file)
        inspections.append(inspection)
        headings = extract_and_format_highlighted_headings(
            __import__("docx").Document(str(source_path)),
            config.source.get("highlight", {}),
            config.fonts,
            normalize=False,
        )
        if not headings:
            raise SourceStrategyError(f"未在 {source_path.name} 发现黄色高亮标题。")
        final_tree = [{
            "title": source_path.stem,
            "level": 1,
            "type": "highlighted_docx",
            "file": source_path.name,
            "children": [],
            "content_block_count": len(headings),
            "heading_count": len(headings),
        }]
        nodes_override = [{**item, "file": source_path.name} for item in headings]
    elif strategy in {"directory_tree", "explicit_tree"}:
        if not source_path.is_dir():
            raise SourceStrategyError(f"来源策略 {strategy} 需要输入目录: {source_path}")
        final_tree = build_outline(source_path, config)
        nodes_override = None
    else:
        raise SourceStrategyError(f"不支持的来源策略: {strategy}")

    if strategy != "highlighted_docx":
        nodes_override = None
    flat_nodes = nodes_override or flatten_tree_nodes(final_tree)
    source_files = _source_files(source_path, strategy, target_file)
    source_hashes = {str(path): compute_sha256(path) for path in source_files}

    # 目录型输入也先建立 DOCX 的原始 NodeRef；跨文件身份包含各自源哈希。
    if strategy in {"directory_tree", "explicit_tree"}:
        for path in source_files:
            if path.suffix.lower() != ".docx" or any(item.file_path == str(path) for item in inspections):
                continue
            inspection = inspect_docx(path)
            inspections.append(inspection)
            assignments.extend(RoleMapper().map_inspection(inspection))

    node_index: Dict[str, Dict[str, Any]] = {}
    diagnostics: List[Dict[str, Any]] = []
    sections: List[Dict[str, Any]] = []
    for inspection in inspections:
        diagnostics.extend(_diagnostic_dict(item) for item in inspection.diagnostics)
        sections.extend(asdict(item) for item in inspection.sections)
        for block in inspection.blocks:
            ref = block.node
            node_index[_node_key(ref.source_sha256, ref.part_uri, ref.element_path)] = {
                "source_sha256": ref.source_sha256,
                "part_uri": ref.part_uri,
                "element_path": ref.element_path,
                "structure_type": block.structure_type,
                "story_type": block.story_type,
                "text_hash": ref.text_hash,
            }

    deliveries = tuple(deepcopy(getattr(config, "documents", [])))
    counts = {
        "content_block_count": (
            sum(item.get("content_block_count", 0) for item in final_tree)
            if strategy == "docx_document"
            else len(flat_nodes)
        ),
        "heading_count": (
            sum(1 for assignment in assignments if assignment.role.startswith("heading."))
            if strategy == "docx_document"
            else sum(
                item.get("heading_count", 0) for item in final_tree
            ) if strategy == "highlighted_docx" else sum(
                1 for item in flat_nodes
                if item.get("type") in ("embedded_heading", "folder")
                or (item.get("level") or 99) <= 9
            )
        ),
        "legacy_doc_count": sum(path.suffix.lower() == ".doc" for path in source_files),
        "pptx_count": sum(path.suffix.lower() == ".pptx" for path in source_files),
    }
    warnings = list(getattr(config, "warnings", []))
    if any("toc" in item.get("parts", []) for item in deliveries) and counts["heading_count"] == 0:
        warnings.append("输入材料未包含各级标题，若请求生成目录将无有效条目。")

    selection_plan = {
        "strategy": strategy,
        "source_files": [str(path) for path in source_files],
        "source_order": [str(path) for path in source_files],
        "on_unmapped": getattr(getattr(config, "formatting", None), "on_unmapped", "preserve"),
        "role_map": str(role_map_path) if role_map_path else None,
        "node_identity": "source_sha256 + part_uri + element_path + text_hash",
    }
    format_hash = getattr(getattr(config, "resolved_format", None), "content_hash", "")
    format_source_hashes = dict(
        getattr(getattr(config, "resolved_format", None), "source_hashes", {}) or {}
    )
    cover_template_path: Optional[Path] = None
    cover_template_hash = ""
    if any("cover" in item.get("parts", []) for item in deliveries):
        cover_template_path = config.get_template_path()
        if cover_template_path:
            cover_template_hash = compute_sha256(cover_template_path)
    return PreparedBuild(
        source_path=source_path,
        config=config,
        config_hash=compute_config_contract_hash(config),
        config_provenance=dict(getattr(config, "provenance", {})),
        source_hashes=source_hashes,
        format_hash=format_hash,
        format_source_hashes=format_source_hashes,
        cover_template_path=cover_template_path,
        cover_template_hash=cover_template_hash,
        role_map_hash=role_map_hash,
        role_map_path=role_map_path,
        source_order=tuple(str(path) for path in source_files),
        nodes=tuple(deepcopy(item) for item in flat_nodes),
        node_index=node_index,
        selection_plan=selection_plan,
        assignments=tuple(assignments),
        inspections=tuple(inspections),
        parts=dict(getattr(config, "parts", {})),
        sections=tuple(sections),
        deliveries=deliveries,
        diagnostics=tuple(diagnostics),
        strategy=strategy,
        final_tree=tuple(deepcopy(item) for item in final_tree),
        source_docx_path=target_file,
        warnings=tuple(warnings),
        counts=counts,
    )


def prepare_build(source: Path, config: Any) -> PreparedBuild:
    """R3 公开函数别名：使用调用方已解析的 ProjectConfig 做只读准备。"""
    return prepare_project_build(Path(source).resolve(), config)


def verify_inputs_unchanged(plan: PreparedBuild) -> None:
    """R3 公开校验函数，便于预检器在渲染前复用同一结果。"""
    plan.verify_inputs_unchanged()
