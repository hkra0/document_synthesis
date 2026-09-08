# -*- coding: utf-8 -*-
"""可配置的材料大纲来源策略。"""

import copy
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from .scanner import SUPPORTED_EXTENSIONS, scan_directory_to_tree


class SourceStrategyError(ValueError):
    """来源策略无法生成一个安全、完整的大纲时抛出。"""


def resolve_docx_document_path(source_dir: Path, config: Any) -> Path:
    """解析 docx_document 的唯一源文件，统一供 prepare、plan 与 render 使用。"""
    source_dir = Path(source_dir).resolve()
    source_cfg = getattr(config, "source", {}) or {}
    file_name = source_cfg.get("file") or source_cfg.get("path")
    if file_name:
        candidate = Path(file_name)
        if not candidate.is_absolute():
            base = source_dir if source_dir.is_dir() else source_dir.parent
            candidate = base / candidate
        file_path = candidate.resolve()
    elif source_dir.is_file():
        file_path = source_dir
    else:
        docx_files = [
            p.resolve() for p in source_dir.glob("*.docx")
            if not p.name.startswith(("~$", "."))
        ]
        if len(docx_files) != 1:
            raise SourceStrategyError(
                f"docx_document 策略需要 source.file 或目录内唯一 .docx 文件: {source_dir}"
            )
        file_path = docx_files[0]

    if not file_path.is_file() or file_path.suffix.lower() != ".docx":
        raise SourceStrategyError(f"docx_document 必须引用存在的 .docx 文件: {file_path}")
    return file_path


def _assign_bookmarks(nodes: List[Dict[str, Any]], counter: List[int] = None) -> List[Dict[str, Any]]:
    """为所有节点重建唯一且与书签名称一致的 ID。"""
    if counter is None:
        counter = [0]
    for node in nodes:
        counter[0] += 1
        if "bm_id" not in node:
            node["bm_id"] = counter[0]
        if not node.get("bookmark_name"):
            node["bookmark_name"] = f"_Toc_auto_{counter[0]:03d}"
        children = node.get("children") or []
        if children:
            _assign_bookmarks(children, counter)
    return nodes


def _validate_explicit_node(node: Dict[str, Any], source_dir: Path, expected_level: int) -> None:
    if not isinstance(node, dict):
        raise SourceStrategyError("source.tree 中的每个节点必须是对象。")
    if node.get("level") != expected_level:
        raise SourceStrategyError(
            f"节点 {node.get('title', '<未命名>')} 的 level 应为 {expected_level}。"
        )
    if not isinstance(node.get("title"), str) or not node["title"].strip():
        raise SourceStrategyError("每个显式目录节点必须包含非空 title。")

    node_type = node.get("type")
    if node_type == "folder":
        children = node.get("children")
        if not isinstance(children, list) or not children:
            raise SourceStrategyError(f"文件夹节点 {node['title']} 必须包含非空 children。")
        for child in children:
            _validate_explicit_node(child, source_dir, expected_level + 1)
        return

    if node_type == "docx_outline":
        rel_file = node.get("file")
        if not isinstance(rel_file, str) or not rel_file:
            raise SourceStrategyError(f"内嵌目录节点 {node['title']} 必须包含 file。")
        file_path = (source_dir / rel_file).resolve()
        try:
            file_path.relative_to(source_dir.resolve())
        except ValueError as exc:
            raise SourceStrategyError(f"节点 {node['title']} 的 file 不能指向输入目录外。") from exc
        if not file_path.is_file() or file_path.suffix.lower() != ".docx":
            raise SourceStrategyError(f"内嵌目录节点 {node['title']} 必须引用存在的 .docx 文件: {rel_file}")
        rules = node.get("outline_rules")
        if not isinstance(rules, list) or not rules:
            raise SourceStrategyError(f"内嵌目录节点 {node['title']} 必须包含非空 outline_rules。")
        for rule in rules:
            if not isinstance(rule, dict) or not isinstance(rule.get("pattern"), str):
                raise SourceStrategyError(f"内嵌目录节点 {node['title']} 的每条规则必须包含正则 pattern。")
            if not isinstance(rule.get("level"), int) or rule["level"] < expected_level:
                raise SourceStrategyError(f"内嵌目录节点 {node['title']} 的规则 level 不能小于 {expected_level}。")
            try:
                re.compile(rule["pattern"])
            except re.error as exc:
                raise SourceStrategyError(f"内嵌目录节点 {node['title']} 的 pattern 无效: {rule['pattern']}") from exc
            replacements = rule.get("title_replacements", {})
            if not isinstance(replacements, dict) or not all(
                isinstance(key, str) and isinstance(value, str) for key, value in replacements.items()
            ):
                raise SourceStrategyError(f"内嵌目录节点 {node['title']} 的 title_replacements 必须是字符串映射。")
        return

    if node_type not in set(SUPPORTED_EXTENSIONS.values()):
        raise SourceStrategyError(f"节点 {node['title']} 使用了不支持的 type: {node_type}")
    rel_file = node.get("file")
    if not isinstance(rel_file, str) or not rel_file:
        raise SourceStrategyError(f"文件节点 {node['title']} 必须包含 file。")
    file_path = (source_dir / rel_file).resolve()
    try:
        file_path.relative_to(source_dir.resolve())
    except ValueError as exc:
        raise SourceStrategyError(f"节点 {node['title']} 的 file 不能指向输入目录外。") from exc
    if not file_path.is_file():
        raise SourceStrategyError(f"节点 {node['title']} 引用的文件不存在: {rel_file}")


def _expand_docx_outline_nodes(nodes: List[Dict[str, Any]], source_dir: Path) -> None:
    """将 DOCX 正文标题转成虚拟目录节点；正文仍只在渲染阶段复制一次。"""
    from docx import Document

    for node in nodes:
        if node.get("type") == "folder":
            _expand_docx_outline_nodes(node.get("children", []), source_dir)
            continue
        if node.get("type") != "docx_outline":
            continue

        rules = node["outline_rules"]
        virtual_children: List[Dict[str, Any]] = []
        for paragraph in Document(str(source_dir / node["file"])).paragraphs:
            source_text = paragraph.text.strip()
            if not source_text:
                continue
            for rule in rules:
                if re.search(rule["pattern"], source_text):
                    title = rule.get("title_replacements", {}).get(source_text, source_text)
                    virtual_children.append({
                        "level": rule["level"],
                        "title": title,
                        "toc_title": title,
                        "type": "embedded_heading",
                        "source_text": source_text,
                    })
                    break
        if not virtual_children:
            raise SourceStrategyError(
                f"内嵌目录节点 {node['title']} 未在 {node['file']} 中匹配到任何标题。"
            )
        node["children"] = virtual_children


def build_outline(
    source_dir: Path,
    config: Any,
    assignments: Optional[List[Any]] = None,
    inspection: Optional[Any] = None,
) -> List[Dict[str, Any]]:
    """依据 config.source.strategy 生成最终树，并保证书签 ID 全局唯一。"""
    strategy = config.source.get("strategy", "directory_tree")
    if strategy == "directory_tree":
        tree = config.apply_overrides_to_tree(scan_directory_to_tree(source_dir))
    elif strategy == "explicit_tree":
        raw_tree = config.source.get("tree")
        if not isinstance(raw_tree, list) or not raw_tree:
            raise SourceStrategyError("explicit_tree 策略需要 source.tree 非空数组。")
        tree = copy.deepcopy(raw_tree)
        for node in tree:
            _validate_explicit_node(node, source_dir, 1)
        _expand_docx_outline_nodes(tree, source_dir)
        tree = config.apply_overrides_to_tree(tree)
    elif strategy == "docx_document":
        file_path = resolve_docx_document_path(source_dir, config)
        if assignments is None or inspection is None:
            from .role_mapper import RoleMapper
            mapper = RoleMapper()
            assignments, inspection = mapper.map_document(file_path)

        heading_children: List[Dict[str, Any]] = []
        for idx, a in enumerate(assignments):
            if a.role.startswith("heading."):
                heading_children.append({
                    "idx": idx,
                    "level": a.level or 1,
                    "title": a.title_text or "",
                    "toc_title": a.title_text or "",
                    "type": "embedded_heading",
                    "file": file_path.name,
                    "node_ref_path": a.node_ref.element_path,
                    "bookmark_name": a.bookmark_name,
                    "bm_id": len(heading_children) + 1,
                })

        tree = [{
            "title": file_path.stem,
            "level": 1,
            "type": "docx_document",
            "file": file_path.name,
            "include_in_toc": False,
            "children": heading_children,
            "content_block_count": len(inspection.blocks),
            "heading_count": len(heading_children),
        }]
        tree = config.apply_overrides_to_tree(tree)
    else:
        raise SourceStrategyError(f"不支持的材料来源策略: {strategy}")

    if not tree:
        raise SourceStrategyError("来源策略没有生成任何可合成节点。")
    return _assign_bookmarks(tree)
