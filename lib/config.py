# -*- coding: utf-8 -*-
"""
配置清单与复写合并模块 (lib/config.py)
支持加载 JSON 清单文件并与自动发现的大纲树执行深度合并与参数复写。
"""

import os
import json
import re
from copy import deepcopy
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any, List, Optional, Union

from .contracts import (
    ContractError,
    merge_dict_with_provenance as _merge_dict_with_provenance,
    provenance_for_value,
    validate_contract,
)


DEFAULT_SCHEMA_VERSION = 2
SUPPORTED_SCHEMA_VERSIONS = {1, 2, 3}
SCHEMA_VERSION = DEFAULT_SCHEMA_VERSION


class ConfigError(ValueError):
    """项目配置不符合契约时抛出，禁止静默退回默认行为。"""


@dataclass(frozen=True)
class CoverSpec:
    """解析后的封面来源契约，供 plan/render/QA 共享。"""

    mode: str
    template_path: Optional[Path]
    dynamic_fields: Dict[str, str]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "mode": self.mode,
            "template_path": str(self.template_path) if self.template_path else None,
            "dynamic_fields": dict(self.dynamic_fields),
        }


def merge_dict(base: Dict[str, Any], patch: Dict[str, Any]) -> Dict[str, Any]:
    """纯函数合并字典：字典递归合并，数组与标量直接替换，保留 null/0/False。"""
    result = deepcopy(base)
    for key, value in patch.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = merge_dict(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def merge_dict_with_provenance(
    base: Dict[str, Any],
    patch: Dict[str, Any],
    base_provenance: Optional[Dict[str, str]] = None,
    patch_source_label: str = "override",
    current_path: str = "",
):
    """Config counterpart of format inheritance merge with source tracking."""
    return _merge_dict_with_provenance(
        base,
        patch,
        base_provenance,
        patch_source_label,
        current_path,
    )


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
            f for f in f_dir.iterdir()
            if f.is_file() and f.suffix.lower() == ".docx"
            and not f.name.startswith(("~$", "."))
            and ("封面" in f.name or "模版" in f.name or "模板" in f.name)
        ]
        if candidates:
            return candidates[0].resolve()

    return None


def _is_comment_key(key: str) -> bool:
    return isinstance(key, str) and key.startswith("_comment")


WINDOWS_RESERVED_NAMES = frozenset({
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
})
WINDOWS_FORBIDDEN_CHARS = frozenset('<>:"/\\|?*')


def _is_windows_reserved_name(name: str) -> bool:
    stem = name.split(".")[0].strip().upper()
    return stem in WINDOWS_RESERVED_NAMES


def _validate_filename(value: Any, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"output.{field_name} 必须是非空 DOCX 文件名。")
    candidate = Path(value)
    if (candidate.is_absolute() or candidate.name != value or candidate.suffix.lower() != ".docx"
            or any(char in WINDOWS_FORBIDDEN_CHARS for char in value)
            or any(ord(char) < 32 for char in value)
            or any(unicodedata.category(char) == "Cc" for char in value)
            or value.startswith((".", "~$"))
            or value.endswith((".", " "))
            or candidate.stem.endswith((".", " "))
            or value != value.strip()):
        raise ConfigError(f"output.{field_name} 只能是输出目录内的合法 .docx 文件名，不能包含路径、Windows 非法字符或尾部空格/点。")
    if _is_windows_reserved_name(value) or _is_windows_reserved_name(candidate.stem):
        raise ConfigError(f"output.{field_name} 不能使用 Windows 保留设备名称 ({candidate.stem})。")
    return candidate.name


def _collision_key(value: str) -> str:
    return unicodedata.normalize("NFC", value).casefold()


def _validate_project_name(value: Any) -> str:
    """The name also becomes an output directory and profile lookup component."""
    if (not isinstance(value, str) or not value or value != value.strip()
            or value.startswith((".", "~$"))
            or value.endswith((".", " "))
            or any(char in WINDOWS_FORBIDDEN_CHARS for char in value)
            or any(ord(char) < 32 for char in value)
            or any(unicodedata.category(char) == "Cc" for char in value)):
        raise ConfigError("project_name 必须是安全的单级目录名，不能包含路径、Windows 非法字符、控制字符、首尾空白/点或以 .、~$ 开头。")
    if _is_windows_reserved_name(value):
        raise ConfigError(f"project_name 不能使用 Windows 保留设备名称 ({value})。")
    return value


def _get_parts_registry(raw: Dict[str, Any]):
    from .document_parts import DocumentPart, get_default_parts
    layout = raw.get("layout", {})
    if isinstance(layout, dict) and "parts" in layout and isinstance(layout["parts"], dict):
        reg = {}
        for pid, pdata in layout["parts"].items():
            if isinstance(pdata, dict):
                reg[pid] = DocumentPart(
                    id=pid,
                    kind=pdata.get("kind", "content"),
                    source_region=pdata.get("source_region"),
                    section_ref=pdata.get("section_ref"),
                    section_type=pdata.get("section_type", "oddPage" if pdata.get("section_ref") in ("odd_page", "oddPage") else "nextPage"),
                    page_sequence=pdata.get("page_sequence"),
                    include_in_toc=pdata.get("include_in_toc", True),
                )
        if reg:
            return reg
    return get_default_parts()


def _get_source_regions(raw: Dict[str, Any]):
    from .document_parts import SourceRegion
    regs_raw = (
        raw.get("source", {}).get("regions")
        or raw.get("formatting", {}).get("regions")
        or raw.get("regions")
        or {}
    )
    res = {}
    if isinstance(regs_raw, dict):
        for rid, rdef in regs_raw.items():
            if isinstance(rdef, dict) and "start" in rdef and "end" in rdef:
                res[rid] = SourceRegion(
                    id=rid,
                    start=rdef["start"],
                    end=rdef["end"],
                    exclude=bool(rdef.get("exclude", False)),
                )
    return res


def _get_page_sequences(raw: Dict[str, Any]):
    from .pagination_types import PageSequence
    layout = raw.get("layout", {})
    if isinstance(layout, dict) and "page_sequences" in layout and isinstance(layout["page_sequences"], dict):
        res = {}
        for sid, sdef in layout["page_sequences"].items():
            if isinstance(sdef, dict):
                res[sid] = PageSequence(
                    id=sid,
                    format=sdef.get("format", "decimal"),
                    start=sdef.get("start"),
                )
        if res:
            return res
    return {}


def _normalize_documents(raw: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Validate the delivery contract, including references, before any build work."""
    output = raw.get("output", {})
    project = raw["project_name"]
    parts_registry = _get_parts_registry(raw)
    is_custom_parts = (
        "layout" in raw
        and isinstance(raw["layout"], dict)
        and "parts" in raw["layout"]
        and bool(raw["layout"]["parts"])
    )

    if raw["schema_version"] == 1:
        allowed = {"compiled_document", "cover_toc", "toc_body"}
        defaults = {
            "compiled_document": (f"{project}_合成材料.docx", ["body"]),
            "cover_toc": (f"{project}_封面+目录.docx", ["cover", "toc"]),
            "toc_body": (f"{project}_目录+正文.docx", ["toc", "body"]),
        }
        documents = []
        for doc_id, (filename, parts) in defaults.items():
            doc = {"id": doc_id, "filename": output.get(doc_id, filename), "parts": parts}
            if "toc" in parts:
                doc["toc"] = {"reference": "toc_body", "links": "none" if doc_id == "cover_toc" else "internal"}
            documents.append(doc)
    else:
        allowed = {"documents"}
        documents = deepcopy(output.get("documents", [{
            "id": "main", "filename": f"{project}_完整文档.docx", "parts": ["cover", "toc", "body"],
        }]))
    for key in output:
        if key not in allowed and not _is_comment_key(key):
            raise ConfigError(f"output 包含不支持的字段: {key}；v2 请使用 output.documents，迁移时移除旧输出字段。")
    if not isinstance(documents, list) or not documents:
        raise ConfigError("output.documents 必须是非空数组；省略该字段才使用默认交付文档。")
    ids, names = set(), set()
    normalized = []
    for index, document in enumerate(documents):
        field = f"documents[{index}]"
        if not isinstance(document, dict):
            raise ConfigError(f"output.{field} 必须是对象。")
        unknown = set(document) - {"id", "filename", "parts", "toc"}
        if any(not _is_comment_key(key) for key in unknown):
            raise ConfigError(f"output.{field} 包含不支持的字段: {', '.join(sorted(unknown))}")
        doc_id = document.get("id")
        if not isinstance(doc_id, str) or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", doc_id):
            raise ConfigError(f"output.{field}.id 必须以英文字母开头，且只含字母、数字、下划线或连字符。")
        filename = _validate_filename(document.get("filename"), f"{field}.filename")
        if _collision_key(doc_id) in ids or _collision_key(filename) in names:
            raise ConfigError("output 中的文档 ID 和交付文件名必须彼此不同（不区分大小写）。")
        ids.add(_collision_key(doc_id))
        names.add(_collision_key(filename))
        parts = document.get("parts")

        if not is_custom_parts:
            canonical = [part for part in ("cover", "toc", "body") if isinstance(parts, list) and part in parts]
            if not isinstance(parts, list) or not parts or parts != canonical:
                raise ConfigError(f"output.{field}.parts 必须是 cover、toc、body 的非空有序子集，不能重复。")
            has_content = "body" in parts
            has_toc = "toc" in parts
        else:
            if not isinstance(parts, list) or not parts:
                raise ConfigError(f"output.{field}.parts 必须是非空数组。")
            if len(parts) != len(set(parts)):
                raise ConfigError(f"output.{field}.parts 不能包含重复部件: {parts}")
            unknown_parts = [p for p in parts if p not in parts_registry]
            if unknown_parts:
                raise ConfigError(f"output.{field}.parts 包含未在 layout.parts 中定义的部件 ID: {', '.join(unknown_parts)}")
            seqs = _get_page_sequences(raw)
            for p in parts:
                part_def = parts_registry.get(p)
                if part_def and part_def.page_sequence and seqs:
                    if part_def.page_sequence not in seqs:
                        raise ConfigError(f"部件 {p} 引用的 page_sequence '{part_def.page_sequence}' 未在 layout.page_sequences 中定义。")
            has_content = any(parts_registry[p].kind == "content" for p in parts)
            has_toc = any(parts_registry[p].kind == "generated_toc" for p in parts)

        doc = {"id": doc_id, "filename": filename, "parts": list(parts)}
        if not has_toc and "toc" in document:
            raise ConfigError(f"output.{field} 不包含目录部件，不能配置 toc 选项。")
        if has_toc:
            toc = document.get("toc", {})
            if not isinstance(toc, dict):
                raise ConfigError(f"output.{field}.toc 必须是对象。")
            if any(key not in {"reference", "links"} and not _is_comment_key(key) for key in toc):
                raise ConfigError(f"output.{field}.toc 包含不支持的字段。")
            reference = toc.get("reference", doc_id if has_content else None)
            if not isinstance(reference, str) or not reference:
                raise ConfigError(f"output.{field}.toc.reference 必须指向含正文/内容的交付文档 ID。")
            links = toc.get("links", "internal" if has_content else "none")
            if links not in ("internal", "none"):
                raise ConfigError(f"output.{field}.toc.links 只支持 internal 或 none。")
            if not has_content and links != "none":
                raise ConfigError(f"output.{field} 不含正文/内容，toc.links 必须为 none。")
            if has_content and reference != doc_id:
                raise ConfigError(f"output.{field} 含正文/内容，目录必须引用本文件 ID {doc_id}。")
            doc["toc"] = {"reference": reference, "links": links}
        normalized.append(doc)
    by_id = {doc["id"]: doc for doc in normalized}
    for doc in normalized:
        if "toc" in doc:
            target = by_id.get(doc["toc"]["reference"])
            if not is_custom_parts:
                target_has_content = target is not None and "body" in target["parts"]
            else:
                target_has_content = target is not None and any(parts_registry[p].kind == "content" for p in target["parts"])
            if target is None or not target_has_content:
                raise ConfigError(f"文档 {doc['id']} 的 toc.reference 必须指向存在且含正文的交付文档。")
    return normalized


def _validate_config(raw: Dict[str, Any]) -> None:
    if not isinstance(raw, dict):
        raise ConfigError("配置根节点必须是 JSON 对象。")

    schema_version = raw.get("schema_version")
    if type(schema_version) is not int or schema_version not in SUPPORTED_SCHEMA_VERSIONS:
        raise ConfigError("schema_version 必须为 1、2 或 3；v2 为默认配置，v3 为自定义格式配置。")

    _validate_project_name(raw.get("project_name"))

    if schema_version == 3:
        if "page_setup" in raw or "fonts" in raw:
            raise ConfigError("schema_version 3 不再接受顶层 fonts 与 page_setup，请在 format 中配置。")

        allowed_v3 = {
            "schema_version", "project_name", "format", "formatting", "output",
            "source", "cover", "node_overrides", "tree_order", "ignore_nodes",
            "layout", "regions",
        }
        unknown = [key for key in raw if key not in allowed_v3 and not _is_comment_key(key)]
        if unknown:
            raise ConfigError(f"配置包含不支持的字段: {', '.join(unknown)}")

        for field in ("format", "formatting", "cover", "output", "source", "node_overrides", "layout"):
            if field in raw and not isinstance(raw[field], dict):
                raise ConfigError(f"{field} 必须是对象。")
        for field in ("tree_order", "ignore_nodes"):
            if field in raw and (not isinstance(raw[field], list) or not all(isinstance(item, str) for item in raw[field])):
                raise ConfigError(f"{field} 必须是字符串数组。")

        if "layout" in raw:
            layout_raw = raw["layout"]
            for lk in layout_raw:
                if lk not in ("parts", "page_sequences") and not _is_comment_key(lk):
                    raise ConfigError(f"layout 包含不支持的字段: {lk}")
            if "parts" in layout_raw:
                parts_map = layout_raw["parts"]
                if not isinstance(parts_map, dict) or not parts_map:
                    raise ConfigError("layout.parts 必须是非空对象。")
                for pid, pdef in parts_map.items():
                    if not isinstance(pdef, dict):
                        raise ConfigError(f"layout.parts.{pid} 必须是对象。")
                    if "kind" not in pdef:
                        raise ConfigError(f"layout.parts.{pid} 必须包含 kind 字段。")
                    if pdef["kind"] == "toc":
                        pdef["kind"] = "generated_toc"
                    if pdef["kind"] not in ("cover", "generated_toc", "content"):
                        raise ConfigError(f"layout.parts.{pid}.kind 只支持 cover, generated_toc 或 content，收到: {pdef['kind']}")
            if "page_sequences" in layout_raw:
                seqs_map = layout_raw["page_sequences"]
                if not isinstance(seqs_map, dict):
                    raise ConfigError("layout.page_sequences 必须是对象。")
                valid_formats = {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"}
                for sid, sdef in seqs_map.items():
                    if not isinstance(sdef, dict):
                        raise ConfigError(f"layout.page_sequences.{sid} 必须是对象。")
                    fmt = sdef.get("format", "decimal")
                    if fmt not in valid_formats:
                        raise ConfigError(f"layout.page_sequences.{sid}.format 无效: {fmt}，允许值: {', '.join(sorted(valid_formats))}")
                    if "start" in sdef:
                        start = sdef["start"]
                        if not isinstance(start, int) or start < 1:
                            raise ConfigError(f"layout.page_sequences.{sid}.start 必须是大于等于 1 的整数。")

        # 校验 format
        format_raw = raw.get("format")
        if not isinstance(format_raw, dict):
            raise ConfigError("schema_version 3 必须包含 format 对象。")
        ref = format_raw.get("ref")
        if not isinstance(ref, str) or not ref.strip():
            raise ConfigError("format.ref 必须是非空字符串 (如 preset:academic-basic@1.0.0 或格式文件路径)。")
        if "overrides" in format_raw and not isinstance(format_raw["overrides"], dict):
            raise ConfigError("format.overrides 必须是对象。")

        # 校验 formatting
        formatting_raw = raw.get("formatting")
        if not isinstance(formatting_raw, dict):
            raise ConfigError("schema_version 3 必须包含 formatting 对象。")
        mode = formatting_raw.get("mode")
        if mode not in ("preserve", "restyle", "mixed"):
            raise ConfigError(f"formatting.mode 只支持 preserve, restyle 或 mixed，收到: {mode}")
        if "on_unmapped" in formatting_raw and formatting_raw["on_unmapped"] not in ("preserve", "error"):
            raise ConfigError(f"formatting.on_unmapped 只支持 preserve 或 error，收到: {formatting_raw['on_unmapped']}")
        if "page_policy" in formatting_raw and formatting_raw["page_policy"] not in ("target", "source"):
            raise ConfigError(f"formatting.page_policy 只支持 target 或 source，收到: {formatting_raw['page_policy']}")
        if "inline_emphasis" in formatting_raw and formatting_raw["inline_emphasis"] not in ("preserve", "target"):
            raise ConfigError(f"formatting.inline_emphasis 只支持 preserve 或 target，收到: {formatting_raw['inline_emphasis']}")
        if "source_toc" in formatting_raw and formatting_raw["source_toc"] not in ("preserve", "error"):
            raise ConfigError(f"formatting.source_toc 只支持 preserve 或 error，收到: {formatting_raw['source_toc']}")
        if mode == "mixed" and "page_policy" not in formatting_raw:
            raise ConfigError("formatting.mode 为 mixed 时必须显式声明 page_policy。")
        if "scopes" in formatting_raw:
            scopes = formatting_raw["scopes"]
            if not isinstance(scopes, list):
                raise ConfigError("formatting.scopes 必须是数组。")
            for idx, sc in enumerate(scopes):
                if not isinstance(sc, dict) or "file" not in sc or "mode" not in sc:
                    raise ConfigError(f"formatting.scopes[{idx}] 必须包含 file 和 mode。")
                if sc["mode"] not in ("preserve", "restyle"):
                    raise ConfigError(f"formatting.scopes[{idx}].mode 只支持 preserve 或 restyle。")

        # 校验 source
        source = raw.get("source", {"strategy": "directory_tree"})
        strategy = source.get("strategy", "directory_tree")
        if strategy not in {"directory_tree", "explicit_tree", "highlighted_docx", "docx_document"}:
            raise ConfigError(f"source.strategy 不受支持: {strategy}")

        # Python validation above retains the public, domain-specific error
        # messages. The schema pass then closes nested unknown-field and type
        # gaps before any resolver, cache, or Word operation is reached.
        try:
            validate_contract(raw, "project-v3.schema.json", context="project-v3 配置")
        except ContractError as exc:
            raise ConfigError(str(exc)) from exc
    else:
        # v1 / v2
        allowed = {
            "schema_version", "project_name", "page_setup", "fonts", "cover", "output",
            "source", "node_overrides", "tree_order", "ignore_nodes",
        }
        unknown = [key for key in raw if key not in allowed and not _is_comment_key(key)]
        if unknown:
            raise ConfigError(f"配置包含不支持的字段: {', '.join(unknown)}")

        for field in ("page_setup", "fonts", "cover", "output", "source", "node_overrides"):
            if field in raw and not isinstance(raw[field], dict):
                raise ConfigError(f"{field} 必须是对象。")
        for field in ("tree_order", "ignore_nodes"):
            if field in raw and (not isinstance(raw[field], list) or not all(isinstance(item, str) for item in raw[field])):
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

        source = raw.get("source", {"strategy": "directory_tree"})
        allowed_source_keys = {"strategy", "tree", "highlight", "file", "path"}
        invalid_source_keys = [key for key in source if key not in allowed_source_keys and not _is_comment_key(key)]
        if invalid_source_keys:
            raise ConfigError(f"source 包含不支持的字段: {', '.join(invalid_source_keys)}")
        strategy = source.get("strategy", "directory_tree")
        if strategy not in {"directory_tree", "explicit_tree", "highlighted_docx", "docx_document"}:
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

    # Common for all versions
    validate_cover = raw.get("cover", {})
    if isinstance(validate_cover, dict):
        for key, value in validate_cover.items():
            if _is_comment_key(key):
                continue
            if key not in {"mode", "header_title", "sub_title", "main_title", "author", "date", "template"}:
                raise ConfigError(f"cover 包含不支持的字段: {key}")
            if not (value is None or isinstance(value, str) or value is False):
                raise ConfigError(f"cover.{key} 的值类型不正确。")
        mode = validate_cover.get("mode")
        if mode is not None and mode not in {"generated", "template", "static_template"}:
            raise ConfigError(f"cover.mode 只支持 generated、template 或 static_template，收到: {mode}")
        if raw.get("schema_version") == 3:
            template = validate_cover.get("template")
            effective_mode = mode
            if effective_mode is None:
                effective_mode = "template" if isinstance(template, str) and template else "generated"
            if effective_mode == "generated" and template not in (None, False):
                raise ConfigError("cover.mode=generated 不能同时指定 template。")
            if effective_mode in {"template", "static_template"} and not (isinstance(template, str) and template.strip()):
                raise ConfigError(f"cover.mode={effective_mode} 必须提供有效 template 路径。")
            if effective_mode == "static_template":
                dynamic = {key for key in ("header_title", "sub_title", "main_title", "author", "date") if key in validate_cover}
                if dynamic:
                    raise ConfigError(
                        "cover.mode=static_template 不能同时设置动态封面字段: " + ", ".join(sorted(dynamic))
                    )

    _normalize_documents(raw)

    node_overrides = raw.get("node_overrides", {})
    for key, value in node_overrides.items():
        if not _is_comment_key(key) and not isinstance(value, dict):
            raise ConfigError(f"node_overrides.{key} 必须是对象。")


def _declared_base_dir(
    provenance: Dict[str, str],
    pointer: str,
    fallback: Path,
) -> Path:
    """Resolve a relative reference against the file that declared it."""
    label = provenance.get(pointer)
    if not label:
        # A recursive merge may only have recorded the parent container. Walk
        # upward until a leaf source is found.
        current = pointer
        while current and current != "/":
            current = current.rsplit("/", 1)[0] or "/"
            label = provenance.get(current)
            if label:
                break
    if not label or label in {"ProjectConfig", "默认配置"}:
        return Path(fallback)
    source = str(label).split("#", 1)[0]
    candidate = Path(source)
    if candidate.suffix.lower() in {".json", ".yaml", ".yml"} or candidate.is_file():
        return candidate.parent
    return Path(fallback)


def _declared_source_label(
    provenance: Dict[str, str],
    pointer: str,
    fallback: Path,
) -> str:
    label = provenance.get(pointer)
    current = pointer
    while not label and current and current != "/":
        current = current.rsplit("/", 1)[0] or "/"
        label = provenance.get(current)
    if not label:
        return str(fallback)
    return str(label).split("#", 1)[0]


class ProjectConfig:
    """项目排版合成配置实体类"""
    def __init__(
        self,
        raw_data: Optional[Dict[str, Any]] = None,
        project_dir: Optional[Path] = None,
        manifest_path: Optional[Path] = None,
        provenance: Optional[Dict[str, str]] = None,
    ):
        self.project_dir = Path(project_dir) if project_dir else Path(".")
        self.manifest_path: Optional[Path] = Path(manifest_path).resolve() if manifest_path else None
        raw = raw_data if raw_data is not None else {
            "schema_version": DEFAULT_SCHEMA_VERSION,
            "project_name": self.project_dir.name,
        }
        _validate_config(raw)
        self.project_name = raw["project_name"]
        self.schema_version = raw["schema_version"]
        self.override_paths: List[Path] = []
        self.provenance: Dict[str, str] = dict(provenance or provenance_for_value(
            raw,
            str(self.manifest_path) if self.manifest_path else "ProjectConfig",
        ))
        self.warnings = (["schema_version 1 保留三份旧交付物；请迁移为 v2 的 output.documents，默认将交付一份完整文档。"]
                         if self.schema_version == 1 else [])

        # 延迟导入以解耦并避免循环引用
        from lib.format_schema import FormattingPolicy
        from lib.format_resolver import resolve_format_package
        from lib.legacy_format import from_project_config, LegacyPolicy

        # 1. 解析格式包与排版策略
        if self.schema_version == 3:
            manifest_dir = self.manifest_path.parent if self.manifest_path else self.project_dir
            f_raw = raw["formatting"]
            formatting_raw = deepcopy(f_raw)
            role_map_ref = formatting_raw.get("role_map")
            if isinstance(role_map_ref, str) and role_map_ref:
                role_map_base = _declared_base_dir(self.provenance, "/formatting/role_map", manifest_dir)
                if not Path(role_map_ref).is_absolute():
                    formatting_raw["role_map"] = str((role_map_base / role_map_ref).resolve())
            if isinstance(formatting_raw.get("scopes"), list):
                scopes = []
                scopes_base = _declared_base_dir(self.provenance, "/formatting/scopes", manifest_dir)
                for scope in formatting_raw["scopes"]:
                    scope_copy = deepcopy(scope)
                    if isinstance(scope_copy.get("file"), str) and not Path(scope_copy["file"]).is_absolute():
                        scope_copy["file"] = str((scopes_base / scope_copy["file"]).resolve())
                    scopes.append(scope_copy)
                formatting_raw["scopes"] = scopes
            self.formatting = FormattingPolicy(
                mode=formatting_raw.get("mode", "restyle"),
                role_map=formatting_raw.get("role_map"),
                on_unmapped=formatting_raw.get(
                    "on_unmapped",
                    "preserve" if formatting_raw.get("mode") == "mixed" else "error",
                ),
                page_policy=formatting_raw.get(
                    "page_policy",
                    "source" if formatting_raw.get("mode") == "preserve" else "target",
                ),
                inline_emphasis=formatting_raw.get("inline_emphasis", "preserve"),
                source_toc=formatting_raw.get("source_toc", "preserve"),
                scopes=deepcopy(formatting_raw.get("scopes", [])),
            )
            fmt_cfg = raw["format"]
            format_declared_in = _declared_base_dir(self.provenance, "/format/ref", manifest_dir)
            self.resolved_format = resolve_format_package(
                ref=fmt_cfg["ref"],
                overrides=fmt_cfg.get("overrides"),
                declared_in=format_declared_in,
                override_source_label=(
                    _declared_source_label(
                        self.provenance,
                        "/format/overrides",
                        self.manifest_path if self.manifest_path else self.project_dir,
                    )
                    + "#format.overrides"
                    if fmt_cfg.get("overrides")
                    else None
                ),
            )
            # 兼容旧代码直接读取 page_setup 和 fonts
            pg = self.resolved_format.page
            self.page_setup = {
                "width_cm": pg.width_mm / 10.0,
                "height_cm": pg.height_mm / 10.0,
                "margin_top_cm": pg.margin_top_mm / 10.0,
                "margin_bottom_cm": pg.margin_bottom_mm / 10.0,
                "margin_left_cm": pg.margin_left_mm / 10.0,
                "margin_right_cm": pg.margin_right_mm / 10.0,
            }
            body_style = self.resolved_format.styles.get("body")
            body_run = body_style.run if body_style else None
            h1_style = self.resolved_format.styles.get("heading.1")
            h1_run = h1_style.run if h1_style else None
            title_style = self.resolved_format.styles.get("title")
            title_run = title_style.run if title_style else None
            self.fonts = {
                "body": (body_run.east_asia if body_run and body_run.east_asia else "仿宋_GB2312"),
                "h1": (h1_run.east_asia if h1_run and h1_run.east_asia else "黑体"),
                "title": (title_run.east_asia if title_run and title_run.east_asia else "方正小标宋简体"),
                "en": (body_run.latin if body_run and body_run.latin else "Times New Roman"),
            }
            self.legacy_policy = LegacyPolicy.default_for_version(3)
        else:
            self.page_setup = {**DEFAULT_PAGE_SETUP, **raw.get("page_setup", {})}
            self.fonts = {**DEFAULT_FONTS, **raw.get("fonts", {})}
            self.legacy_policy = LegacyPolicy.default_for_version(self.schema_version)
            self.formatting = FormattingPolicy(mode="restyle" if self.schema_version == 2 else "preserve")
            self.resolved_format = from_project_config(self)

        # 2. 封面配置
        cover_raw = raw.get("cover", {})
        dir_name = self.project_name
        cover_template = cover_raw.get("template", None)
        if isinstance(cover_template, str) and cover_template:
            cover_base = _declared_base_dir(
                self.provenance,
                "/cover/template",
                self.manifest_path.parent if self.manifest_path else self.project_dir,
            )
            if not Path(cover_template).is_absolute():
                cover_template = str((cover_base / cover_template).resolve())
        if self.schema_version >= 3:
            cover_mode = cover_raw.get("mode")
            if cover_mode is None:
                cover_mode = "template" if isinstance(cover_template, str) and cover_template else "generated"
            if cover_mode == "generated":
                cover_template = None
        else:
            cover_mode = cover_raw.get("mode")
        self.cover = {
            "mode": cover_mode,
            "header_title": cover_raw.get("header_title", ""),
            "sub_title": cover_raw.get("sub_title", ""),
            "main_title": cover_raw.get("main_title", dir_name),
            "author": cover_raw.get("author", ""),
            "date": cover_raw.get("date", ""),
            "template": cover_template
        }
        dynamic_fields = {} if cover_mode == "static_template" else {
            "header_title": str(cover_raw.get("header_title") or ""),
            "sub_title": str(cover_raw.get("sub_title") or ""),
            "main_title": str(cover_raw.get("main_title") or self.project_name),
            "author": str(cover_raw.get("author") or ""),
            "date": str(cover_raw.get("date") or ""),
        }
        self.cover_spec = CoverSpec(
            mode=cover_mode or ("legacy_auto" if self.schema_version < 3 else "generated"),
            template_path=Path(cover_template).resolve()
            if isinstance(cover_template, str) and cover_template
            else None,
            dynamic_fields=dynamic_fields,
        )

        # 3. 部件与选区配置
        self.parts = _get_parts_registry(raw)
        self.regions = _get_source_regions(raw)
        self.page_sequences = _get_page_sequences(raw)
        self.layout = raw.get("layout", {})

        # 4. 输出配置
        self.documents = _normalize_documents(raw)
        self.output = ({doc["id"]: doc["filename"] for doc in self.documents}
                       if self.schema_version == 1 else {"documents": self.documents})

        # 5. 节点复写与自定义
        self.node_overrides = raw.get("node_overrides", {})
        self.tree_order = raw.get("tree_order", [])
        self.ignore_nodes = raw.get("ignore_nodes", [])
        self.source = {"strategy": "directory_tree", **deepcopy(raw.get("source", {}))}
        source_file = self.source.get("file")
        if isinstance(source_file, str) and source_file:
            source_base = _declared_base_dir(
                self.provenance,
                "/source/file",
                self.manifest_path.parent if self.manifest_path else self.project_dir,
            )
            if not Path(source_file).is_absolute():
                self.source["file"] = str((source_base / source_file).resolve())
        self.raw = raw

    def get_template_path(self) -> Optional[Path]:
        """解析封面模板；v3 只有显式 template 才允许使用模板。"""
        if self.cover.get("template") is False:
            return None
        if self.schema_version >= 3:
            mode = self.cover.get("mode", "generated")
            if mode == "generated":
                return None
            template = self.cover.get("template")
            if not isinstance(template, str) or not template:
                raise ConfigError(f"cover.mode={mode} 未提供模板路径。")
            path = Path(template).resolve()
            if not path.is_file():
                raise ConfigError(f"封面模板不存在: {path}")
            return path
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
    override_paths: Optional[List[Path]] = None,
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
    resolved_name = _validate_project_name(project_name if project_name is not None else project_dir.name)
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
        if p.is_file() and (manifest_path or "example" not in p.name):
            target_json = p
            break
            
    def read_json(path: Path) -> Dict[str, Any]:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigError(f"无法解析配置文件 {path}: {exc}") from exc
        if not isinstance(data, dict):
            raise ConfigError(f"配置文件 {path} 根节点必须是 JSON 对象。")
        return data

    data = read_json(target_json) if target_json else {"schema_version": DEFAULT_SCHEMA_VERSION, "project_name": resolved_name}
    data_provenance = provenance_for_value(
        data,
        str(target_json) if target_json else "默认配置",
    )
    resolved_overrides = []
    for path in override_paths or []:
        path = Path(path).resolve()
        data, data_provenance = merge_dict_with_provenance(
            data,
            read_json(path),
            data_provenance,
            str(path),
        )
        resolved_overrides.append(path)
    try:
        config = ProjectConfig(
            data,
            project_dir,
            manifest_path=target_json,
            provenance=data_provenance,
        )
    except ConfigError as exc:
        provenance = str(target_json) if target_json else "默认配置"
        if resolved_overrides:
            provenance += " + 覆写 " + ", ".join(str(path) for path in resolved_overrides)
        raise ConfigError(f"配置 {provenance} 无效: {exc}") from exc
    config.manifest_path = target_json
    config.override_paths = resolved_overrides
    config.provenance = data_provenance
    return config
