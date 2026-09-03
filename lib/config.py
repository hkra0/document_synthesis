# -*- coding: utf-8 -*-
"""
配置清单与复写合并模块 (lib/config.py)
支持加载 JSON 清单文件并与自动发现的大纲树执行深度合并与参数复写。
"""

import os
import json
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Union


SCHEMA_VERSION = 1


class ConfigError(ValueError):
    """项目配置不符合契约时抛出，禁止静默退回默认行为。"""


DEFAULT_PAGE_SETUP = {
    "width_cm": 21.0,
    "height_cm": 29.7,
    "margin_top_cm": 3.7,
    "margin_bottom_cm": 3.5,
    "margin_left_cm": 2.8,
    "margin_right_cm": 2.6
}

DEFAULT_FONTS = {
    "title": "方正小标宋简体",
    "h1": "黑体",
    "h2": "楷体_GB2312",
    "h3": "仿宋_GB2312",
    "body": "仿宋_GB2312",
    "en": "Times New Roman"
}


def clean_match_key(text: str) -> str:
    """清理字符串用于稳健的键名匹配"""
    return re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', text)


def find_template_file(template_hint: Optional[Union[str, Path]] = None, project_dir: Optional[Path] = None) -> Optional[Path]:
    """
    智能动态查找封面与目录模板文件，不硬编码文件名：
    1. 优先使用显式指定的 template_hint（若文件存在）
    2. 动态扫描 templates 目录下的所有 .docx 模板文件：
       - 若目录下仅有 1 个 .docx 文件，直接作为模板使用（无论其为何名称）
       - 若有多个，按文件名相关度评分匹配（优先匹配包含 封面/目录/模版/模板/cover/template 的文档）
    3. 在项目目录或根目录下兜底查找相关的 .docx 模板
    """
    project_root = Path(__file__).resolve().parent.parent

    # 1. 显式指定的提示路径
    if template_hint:
        p = Path(template_hint)
        if p.is_file():
            return p.resolve()
        for base in [project_root, project_root / "templates", project_dir, project_dir / "templates" if project_dir else None, Path(".")]:
            if base:
                cand = (base / p.name).resolve()
                if cand.is_file():
                    return cand
                cand2 = (base / p).resolve()
                if cand2.is_file():
                    return cand2

    # 2. 动态扫描 templates 目录（优先查找项目专属 templates，次之全局 templates）。
    # 不扫描项目材料根目录，避免把任意支撑 Word 文档误当作封面模板。
    candidate_dirs = []
    if project_dir:
        candidate_dirs.append(Path(project_dir) / "templates")
    candidate_dirs.extend([project_root / "templates", Path("templates").resolve()])

    seen_dirs = set()
    for t_dir in candidate_dirs:
        if not t_dir or not t_dir.exists() or not t_dir.is_dir():
            continue
        real_dir = str(t_dir.resolve())
        if real_dir in seen_dirs:
            continue
        seen_dirs.add(real_dir)

        docx_files = [
            f for f in t_dir.iterdir()
            if f.is_file() and f.suffix.lower() == ".docx"
            and not f.name.startswith("~$") and not f.name.startswith(".")
        ]
        if not docx_files:
            continue

        # 只有一个 docx 模板，直接使用
        if len(docx_files) == 1:
            return docx_files[0].resolve()

        # 多个模板按名称意图择优
        def score_candidate(f: Path) -> int:
            name = f.stem.lower()
            score = 0
            if "封面" in name and "目录" in name:
                score += 100
            elif "封面" in name or "cover" in name:
                score += 60
            elif "模版" in name or "模板" in name or "template" in name:
                score += 50
            elif "目录" in name or "toc" in name:
                score += 30
            else:
                score += 10
            return score

        docx_files.sort(key=score_candidate, reverse=True)
        return docx_files[0].resolve()

    # 3. 兜底扫描：仅接受文件名明确表明为封面或模板的文档。
    fallback_dirs = [project_root]
    if project_dir:
        fallback_dirs.append(Path(project_dir))
    fallback_dirs.append(Path("."))
    for f_dir in fallback_dirs:
        if not f_dir or not f_dir.exists():
            continue
        candidates = [
            f for f in f_dir.glob("*.docx")
            if not f.name.startswith("~$") and not f.name.startswith(".")
            and ("封面" in f.name or "模版" in f.name or "模板" in f.name)
        ]
        if candidates:
            return candidates[0].resolve()

    return None


def _is_comment_key(key: str) -> bool:
    return key.startswith("_comment")


def _validate_filename(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"output.{field_name} 必须是非空 DOCX 文件名。")
    candidate = Path(value)
    if candidate.is_absolute() or candidate.name != value or candidate.suffix.lower() != ".docx":
        raise ConfigError(f"output.{field_name} 只能是输出目录内的 .docx 文件名，不能包含路径。")
    return candidate.name


def _validate_config(raw: Dict[str, Any]) -> None:
    if not isinstance(raw, dict):
        raise ConfigError("配置根节点必须是 JSON 对象。")

    allowed = {
        "schema_version", "project_name", "page_setup", "fonts", "cover", "output",
        "source", "node_overrides", "tree_order", "ignore_nodes",
    }
    unknown = [key for key in raw if key not in allowed and not _is_comment_key(key)]
    if unknown:
        raise ConfigError(f"配置包含不支持的字段: {', '.join(unknown)}")
    if raw.get("schema_version") != SCHEMA_VERSION:
        raise ConfigError(f"schema_version 必须为 {SCHEMA_VERSION}。")
    if not isinstance(raw.get("project_name"), str) or not raw["project_name"].strip():
        raise ConfigError("project_name 必须是非空字符串。")

    for field in ("page_setup", "fonts", "cover", "output", "source", "node_overrides"):
        if field in raw and not isinstance(raw[field], dict):
            raise ConfigError(f"{field} 必须是对象。")
    for field in ("tree_order", "ignore_nodes"):
        if field in raw and not isinstance(raw[field], list) or (
            field in raw and not all(isinstance(item, str) for item in raw[field] if not _is_comment_key(item))
        ):
            raise ConfigError(f"{field} 必须是字符串数组。")

    def validate_mapping(field: str, allowed_keys: set, value_validator) -> None:
        for key, value in raw.get(field, {}).items():
            if _is_comment_key(key):
                continue
            if key not in allowed_keys:
                raise ConfigError(f"{field} 包含不支持的字段: {key}")
            if not value_validator(value):
                raise ConfigError(f"{field}.{key} 的值类型不正确。")

    validate_mapping("page_setup", set(DEFAULT_PAGE_SETUP), lambda value: isinstance(value, (int, float)))
    validate_mapping("fonts", set(DEFAULT_FONTS), lambda value: isinstance(value, str) and bool(value.strip()))
    validate_mapping(
        "cover", {"header_title", "sub_title", "main_title", "author", "date", "template"},
        lambda value: value is None or isinstance(value, str) or value is False,
    )

    source = raw.get("source", {"strategy": "directory_tree"})
    allowed_source_keys = {"strategy", "tree", "highlight"}
    invalid_source_keys = [key for key in source if key not in allowed_source_keys and not _is_comment_key(key)]
    if invalid_source_keys:
        raise ConfigError(f"source 包含不支持的字段: {', '.join(invalid_source_keys)}")
    strategy = source.get("strategy", "directory_tree")
    if strategy not in {"directory_tree", "explicit_tree", "highlighted_docx"}:
        raise ConfigError(f"source.strategy 不受支持: {strategy}")
    if strategy == "explicit_tree" and not isinstance(source.get("tree"), list):
        raise ConfigError("explicit_tree 策略需要 source.tree 数组。")
    if strategy == "highlighted_docx":
        if "tree" in source:
            raise ConfigError("highlighted_docx 策略不接受 source.tree。")
        highlight = source.get("highlight", {})
        if not isinstance(highlight, dict):
            raise ConfigError("source.highlight 必须是对象。")
        allowed_highlight_keys = {"colors", "heading_rules", "title_replacements"}
        invalid_highlight_keys = [key for key in highlight if key not in allowed_highlight_keys and not _is_comment_key(key)]
        if invalid_highlight_keys:
            raise ConfigError(f"source.highlight 包含不支持的字段: {', '.join(invalid_highlight_keys)}")
        if "colors" in highlight and (not isinstance(highlight["colors"], list) or not all(isinstance(color, str) for color in highlight["colors"])):
            raise ConfigError("source.highlight.colors 必须是字符串数组。")
        if "heading_rules" in highlight:
            rules = highlight["heading_rules"]
            if not isinstance(rules, list) or not rules:
                raise ConfigError("source.highlight.heading_rules 必须是非空数组。")
            for rule in rules:
                if not isinstance(rule, dict) or not isinstance(rule.get("level"), int) or not isinstance(rule.get("pattern"), str):
                    raise ConfigError("source.highlight.heading_rules 的每项必须包含整数 level 和字符串 pattern。")
                try:
                    re.compile(rule["pattern"])
                except re.error as exc:
                    raise ConfigError(f"source.highlight.heading_rules 包含无效 pattern: {rule['pattern']}") from exc
        if "title_replacements" in highlight and (
            not isinstance(highlight["title_replacements"], dict)
            or not all(isinstance(k, str) and isinstance(v, str) for k, v in highlight["title_replacements"].items())
        ):
            raise ConfigError("source.highlight.title_replacements 必须是字符串映射。")
    if strategy == "directory_tree" and "tree" in source:
        raise ConfigError("directory_tree 策略不接受 source.tree。")

    for key, value in raw.get("output", {}).items():
        if _is_comment_key(key):
            continue
        if key not in {"compiled_document", "cover_toc", "toc_body"}:
            raise ConfigError(f"output 包含不支持的字段: {key}")
        _validate_filename(value, key)
    configured_outputs = [
        value for key, value in raw.get("output", {}).items()
        if not _is_comment_key(key)
    ]
    if len(configured_outputs) != len(set(configured_outputs)):
        raise ConfigError("output 中的交付文件名必须彼此不同。")

    node_overrides = raw.get("node_overrides", {})
    for key, value in node_overrides.items():
        if not _is_comment_key(key) and not isinstance(value, dict):
            raise ConfigError(f"node_overrides.{key} 必须是对象。")


class ProjectConfig:
    """项目排版合成配置实体类"""
    def __init__(self, raw_data: Optional[Dict[str, Any]] = None, project_dir: Optional[Path] = None):
        self.project_dir = Path(project_dir) if project_dir else Path(".")
        raw = raw_data if raw_data is not None else {
            "schema_version": SCHEMA_VERSION,
            "project_name": self.project_dir.name,
        }
        _validate_config(raw)
        self.project_name = raw["project_name"]
        self.manifest_path: Optional[Path] = None
        
        # 1. 页面与字体配置
        self.page_setup = {**DEFAULT_PAGE_SETUP, **raw.get("page_setup", {})}
        self.fonts = {**DEFAULT_FONTS, **raw.get("fonts", {})}
        
        # 2. 封面配置
        cover_raw = raw.get("cover", {})
        dir_name = self.project_name
        self.cover = {
            "header_title": cover_raw.get("header_title", ""),
            "sub_title": cover_raw.get("sub_title", ""),
            "main_title": cover_raw.get("main_title", dir_name),
            "author": cover_raw.get("author", ""),
            "date": cover_raw.get("date", ""),
            "template": cover_raw.get("template", None)
        }
        
        # 3. 输出配置
        output_raw = raw.get("output", {})
        default_out_doc = f"{dir_name}_合成材料.docx"
        default_out_toc = f"{dir_name}_封面+目录.docx"
        default_out_toc_body = f"{dir_name}_目录+正文.docx"
        self.output = {
            "compiled_document": _validate_filename(output_raw.get("compiled_document", default_out_doc), "compiled_document"),
            "cover_toc": _validate_filename(output_raw.get("cover_toc", default_out_toc), "cover_toc"),
            "toc_body": _validate_filename(output_raw.get("toc_body", default_out_toc_body), "toc_body"),
        }
        
        # 4. 节点复写与自定义
        self.node_overrides = raw.get("node_overrides", {})
        self.tree_order = raw.get("tree_order", [])
        self.ignore_nodes = raw.get("ignore_nodes", [])
        self.source = {"strategy": "directory_tree", **raw.get("source", {})}
        self.raw = raw

    def get_template_path(self) -> Optional[Path]:
        """动态解析项目封面模板路径，优先清单显式配置，次之自动探测 templates/ 目录下的模板文档"""
        if self.cover.get("template") is False:
            return None
        return find_template_file(self.cover.get("template"), self.project_dir)

    def _is_ignored(self, node: Dict[str, Any]) -> bool:
        rel = node.get("file") or node.get("folder", "")
        return any(clean_match_key(ignored) in clean_match_key(rel) for ignored in self.ignore_nodes)

    def apply_overrides_to_node(self, node: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """对单个大纲树节点应用复写规则"""
        rel_file = node.get("file") or node.get("folder", "")
        title = node.get("title", "")
        
        # 查找匹配的复写规则 (按相对路径或按标题匹配)
        match_cfg = None
        if rel_file in self.node_overrides:
            match_cfg = self.node_overrides[rel_file]
        else:
            for k, v in self.node_overrides.items():
                if clean_match_key(k) in clean_match_key(rel_file) or clean_match_key(k) in clean_match_key(title):
                    match_cfg = v
                    break
                    
        if match_cfg:
            for attr, val in match_cfg.items():
                node[attr] = val
                
        # 递归处理子节点
        if "children" in node and node["children"]:
            node["children"] = [
                child for child in node["children"]
                if not self._is_ignored(child)
                and self.apply_overrides_to_node(child) is not None
            ]
            if node.get("type") == "folder" and not node["children"]:
                return None
        return node

    def apply_overrides_to_tree(self, tree_nodes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """对整棵大纲树应用忽略规则、排序规则与属性复写"""
        filtered: List[Dict[str, Any]] = []
        for n in tree_nodes:
            if self._is_ignored(n):
                continue
            updated = self.apply_overrides_to_node(n)
            if updated is not None:
                filtered.append(updated)
            
        # 如果指定了自定义顶级排序
        if self.tree_order:
            def get_order_idx(node):
                rel = node.get("file") or node.get("folder", "")
                for i, ord_k in enumerate(self.tree_order):
                    if clean_match_key(ord_k) in clean_match_key(rel):
                        return i
                return 999
            filtered.sort(key=get_order_idx)
            
        return filtered


def load_project_config(
    project_dir: Path,
    manifest_path: Optional[Path] = None,
    project_name: Optional[str] = None,
) -> ProjectConfig:
    """
    智能加载项目配置：
    1. 优先读取显式指定的 manifest_path
    2. 其次查找 project_dir / "manifest.json"
    3. 再次查找 "profiles" / f"{project_name}.json"（经过验收的通用案例基线）
    4. 最后查找 "manifests" / f"{project_name}.json"（兼容旧清单）
    5. 若均不存在，返回默认配置（纯文件名驱动模式）
    """
    project_dir = Path(project_dir).resolve()
    resolved_name = project_name or project_dir.name
    project_root = Path(__file__).resolve().parent.parent
    candidate_paths: List[Path] = []
    if manifest_path:
        explicit_path = Path(manifest_path).resolve()
        if not explicit_path.is_file():
            raise ConfigError(f"指定的配置文件不存在: {explicit_path}")
        candidate_paths.append(explicit_path)
        
    candidate_paths.append(project_dir / "manifest.json")
    candidate_paths.append(project_root / "profiles" / f"{resolved_name}.json")
    candidate_paths.append(project_root / "manifests" / f"{resolved_name}.json")

    target_json = None
    for p in candidate_paths:
        if "example" not in p.name and p.is_file():
            target_json = p
            break
            
    if target_json:
        try:
            with open(target_json, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigError(f"无法解析配置文件 {target_json}: {exc}") from exc
        try:
            config = ProjectConfig(data, project_dir)
            config.manifest_path = target_json
            return config
        except ConfigError as exc:
            raise ConfigError(f"配置文件 {target_json} 无效: {exc}") from exc

    config = ProjectConfig({"schema_version": SCHEMA_VERSION, "project_name": resolved_name}, project_dir)
    config.manifest_path = None
    return config
