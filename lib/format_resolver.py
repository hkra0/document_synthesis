# -*- coding: utf-8 -*-
"""
格式包解析、继承与校验模块 (lib/format_resolver.py)
负责解析 format 引用、处理 extends 继承链、循环检测、来源追踪与能力匹配。
"""

import os
import json
import hashlib
from copy import deepcopy
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Set

from lib.config import ConfigError
from lib.contracts import ContractError, is_semantic_role, validate_contract
from lib.format_schema import (
    FormatDiagnosticCode,
    Diagnostic,
    LengthValue,
    LineSpacing,
    RunStyle,
    ParagraphStyle,
    StyleDefinition,
    PageSpec,
    RoleSpec,
    TocSpec,
    HeaderFooterSpec,
    ResolvedFormat,
)

# -----------------------------------------------------------------------------
# 引擎能力注册表 (Capability Registry)
# -----------------------------------------------------------------------------

SUPPORTED_ENGINE_CAPABILITIES: Set[str] = {
    "styles.paragraph.v1",
    "styles.character.v1",
    "layout.single_column.v1",
    "docx.single_source_preservation.v1",
    "docx.multisource_basic.v1",
    "numbering.preserve.v1",
    "layout.parts.v1",
    "pagination.roman.v1",
    "notes.merge.v1",
    "fields.managed_update.v1",
}

PLANNED_M2_CAPABILITIES: Set[str] = set()

PLANNED_M3_CAPABILITIES: Set[str] = {
    "bibliography.transform.v1",
}

KNOWN_ENGINE_CAPABILITIES: Set[str] = (
    SUPPORTED_ENGINE_CAPABILITIES | PLANNED_M2_CAPABILITIES | PLANNED_M3_CAPABILITIES
)

MAX_INHERITANCE_DEPTH = 8


# -----------------------------------------------------------------------------
# 递归合并与来源追踪 (Deep Merge with Provenance)
# -----------------------------------------------------------------------------

def deep_merge_with_provenance(
    base: Dict[str, Any],
    patch: Dict[str, Any],
    base_prov: Dict[str, str],
    patch_source_label: str,
    current_path: str = "",
) -> Tuple[Dict[str, Any], Dict[str, str]]:
    """
    纯函数深层合并：
    - 对象：递归合并，展开记录每个叶子节点的来源
    - 数组：整体替换
    - 标量：替换并记录来源（保留显式 0、False，拒绝 null）
    """
    result = deepcopy(base)
    prov = deepcopy(base_prov)

    for key, val in patch.items():
        sub_path = f"{current_path}/{key}" if current_path else f"/{key}"

        if val is None:
            raise ConfigError(f"格式配置属性 {sub_path} 不能为 null；清除属性请使用具体默认值。")

        if isinstance(val, dict):
            if not isinstance(result.get(key), dict):
                result[key] = {}
            sub_res, sub_prov = deep_merge_with_provenance(
                result[key], val, prov, patch_source_label, sub_path
            )
            result[key] = sub_res
            prov.update(sub_prov)
        else:
            result[key] = deepcopy(val)
            prov[sub_path] = patch_source_label

    return result, prov


# -----------------------------------------------------------------------------
# 格式加载与继承解析 (Format Resolver)
# -----------------------------------------------------------------------------

def _resolve_preset_path(preset_ref: str) -> Path:
    """
    解析 'preset:<id>@<version>' 为内置预设文件绝对路径。
    例如: 'preset:report-basic@1.0.0' -> formats/presets/report-basic/1.0.0.json
    """
    prefix = "preset:"
    if not preset_ref.startswith(prefix):
        raise ConfigError(f"内置预设引用必须以 'preset:' 开头: {preset_ref}")
    spec = preset_ref[len(prefix):]
    if "@" not in spec:
        raise ConfigError(f"预设引用必须明确指定版本号 (如 preset:academic-basic@1.0.0): {preset_ref}")
    preset_id, version = spec.split("@", 1)
    if not preset_id or not version:
        raise ConfigError(f"无效的预设引用格式: {preset_ref}")

    project_root = Path(__file__).resolve().parent.parent
    preset_file = project_root / "formats" / "presets" / preset_id / f"{version}.json"
    if not preset_file.is_file():
        raise ConfigError(f"未找到内置格式预设: {preset_ref} (路径: {preset_file})")
    return preset_file.resolve()


def resolve_format_reference(
    ref: str,
    declared_in: Optional[Path] = None,
) -> Path:
    """
    解析格式引用为磁盘文件路径。
    只允许本地文件或 preset: 语法，拒绝 URL 与通配符。
    """
    if ref.startswith("http://") or ref.startswith("https://"):
        raise ConfigError(f"格式引用禁止使用网络 URL: {ref}")

    if ref.startswith("preset:"):
        return _resolve_preset_path(ref)

    ref_path = Path(ref)
    if ref_path.is_absolute():
        if not ref_path.is_file():
            raise ConfigError(f"格式包文件不存在: {ref_path}")
        return ref_path.resolve()

    base_dir = declared_in if declared_in else Path.cwd()
    resolved = (base_dir / ref_path).resolve()
    if not resolved.is_file():
        raise ConfigError(f"格式包文件不存在: {ref} (解析路径: {resolved})")
    return resolved


def load_format_raw(file_path: Path) -> Dict[str, Any]:
    """读取并校验格式包 JSON 文件的结构契约。"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"无法读取格式包文件 {file_path}: {exc}") from exc

    if not isinstance(data, dict):
        raise ConfigError(f"格式包 {file_path} 根节点必须是 JSON 对象。")

    try:
        validate_contract(data, "format-v1.schema.json", context=f"格式包 {file_path}")
    except ContractError as exc:
        raise ConfigError(str(exc)) from exc

    return data


def resolve_format_package(
    ref: str,
    overrides: Optional[Dict[str, Any]] = None,
    declared_in: Optional[Path] = None,
    current_capabilities: Optional[Set[str]] = None,
    override_source_label: Optional[str] = None,
) -> ResolvedFormat:
    """
    解析并构建最终的 ResolvedFormat。
    执行 extends 递归继承、循环检测、来源追踪与 Schema 模型映射。
    """
    active_capabilities = current_capabilities or SUPPORTED_ENGINE_CAPABILITIES

    # 1. 递归构建继承链 (Child -> Parent -> Grandparent)
    chain: List[Tuple[Path, Dict[str, Any]]] = []
    seen_paths: Set[Path] = set()

    current_ref = ref
    current_declared = declared_in

    while current_ref:
        # Detect cycles by canonical paths, so aliases such as ./base.json and
        # base.json cannot bypass the inheritance guard.
        path = resolve_format_reference(current_ref, current_declared)
        if path in seen_paths:
            raise ConfigError(f"[{FormatDiagnosticCode.FORMAT_CYCLE}] 检测到格式继承循环: {current_ref}")
        if len(chain) >= MAX_INHERITANCE_DEPTH:
            raise ConfigError(f"[{FormatDiagnosticCode.FORMAT_CYCLE}] 格式继承深度超过最大限制 ({MAX_INHERITANCE_DEPTH} 层)")

        seen_paths.add(path)
        raw_pkg = load_format_raw(path)
        chain.append((path, raw_pkg))

        parent_ref = raw_pkg.get("extends")
        if parent_ref:
            current_ref = parent_ref
            current_declared = path.parent
        else:
            break

    # 2. 从最顶层父包向下依次深层合并
    chain.reverse()  # 现在最上层父包在前

    merged_data: Dict[str, Any] = {}
    provenance: Dict[str, str] = {}

    for path, pkg in chain:
        source_label = f"{pkg['id']}@{pkg['version']} ({path.name})"
        merged_data, provenance = deep_merge_with_provenance(
            merged_data, pkg, provenance, source_label
        )

    # 3. 合并来自清单的显式 overrides
    if overrides:
        if not isinstance(overrides, dict):
            raise ConfigError("format.overrides 必须是 JSON 对象。")
        override_label = override_source_label or "manifest.json#format.overrides"
        merged_data, provenance = deep_merge_with_provenance(
            merged_data, overrides, provenance, override_label
        )

    # 4. 校验必需能力支持 (required_capabilities)
    req_caps = merged_data.get("required_capabilities", [])
    if not isinstance(req_caps, list):
        raise ConfigError("required_capabilities 必须是字符串数组。")

    for cap in req_caps:
        if cap not in KNOWN_ENGINE_CAPABILITIES:
            raise ConfigError(f"[{FormatDiagnosticCode.UNSUPPORTED_CAPABILITY}] 格式包声明了未知引擎能力: {cap}")
        if cap not in active_capabilities:
            if cap in PLANNED_M3_CAPABILITIES:
                raise ConfigError(f"[{FormatDiagnosticCode.UNSUPPORTED_CAPABILITY}] 能力 {cap} 属于后续扩展规划，当前引擎尚未启用。")
            if cap in PLANNED_M2_CAPABILITIES:
                raise ConfigError(f"[{FormatDiagnosticCode.UNSUPPORTED_CAPABILITY}] 能力 {cap} 属于 M2 学位论文扩展，当前引擎尚未启用。")
            raise ConfigError(f"[{FormatDiagnosticCode.UNSUPPORTED_CAPABILITY}] 当前引擎不支持能力: {cap}")

    # 5. 模型映射与几何核验
    # Page
    raw_page = merged_data.get("page", {})
    page_spec = PageSpec(
        width_mm=float(raw_page.get("width_mm", 210.0)),
        height_mm=float(raw_page.get("height_mm", 297.0)),
        orientation=raw_page.get("orientation", "portrait"),
        margin_top_mm=float(raw_page.get("margin_top_mm", 25.4)),
        margin_bottom_mm=float(raw_page.get("margin_bottom_mm", 25.4)),
        margin_left_mm=float(raw_page.get("margin_left_mm", 31.8)),
        margin_right_mm=float(raw_page.get("margin_right_mm", 31.8)),
        header_distance_mm=float(raw_page.get("header_distance_mm", 15.0)),
        footer_distance_mm=float(raw_page.get("footer_distance_mm", 15.0)),
        snap_to_grid=bool(raw_page.get("snap_to_grid", True)),
    )
    for diag in page_spec.validate():
        if diag.severity == "error":
            raise ConfigError(f"[{diag.code}] 页面几何配置错误: {diag.message} (位置: {diag.location})")

    # Styles
    styles: Dict[str, StyleDefinition] = {}
    raw_styles = merged_data.get("styles", {})
    for s_name, s_data in raw_styles.items():
        run_spec = None
        if "run" in s_data and s_data["run"]:
            r = s_data["run"]
            run_spec = RunStyle(
                east_asia=r.get("east_asia"),
                latin=r.get("latin"),
                complex_script=r.get("complex_script"),
                size_pt=float(r["size_pt"]) if "size_pt" in r and r["size_pt"] is not None else None,
                bold=r.get("bold"),
                italic=r.get("italic"),
                color=r.get("color"),
            )
        para_spec = None
        if "paragraph" in s_data and s_data["paragraph"]:
            p = s_data["paragraph"]
            
            def parse_len(raw_val: Optional[Dict[str, Any]]) -> Optional[LengthValue]:
                if not raw_val or not isinstance(raw_val, dict):
                    return None
                return LengthValue(value=float(raw_val["value"]), unit=raw_val["unit"])

            line_sp = None
            if "line_spacing" in p and p["line_spacing"]:
                lsp = p["line_spacing"]
                line_sp = LineSpacing(
                    mode=lsp.get("mode", "single"),
                    value=float(lsp["value"]) if "value" in lsp and lsp["value"] is not None else None,
                )

            para_spec = ParagraphStyle(
                alignment=p.get("alignment"),
                first_line_indent=parse_len(p.get("first_line_indent")),
                hanging_indent=parse_len(p.get("hanging_indent")),
                left_indent=parse_len(p.get("left_indent")),
                right_indent=parse_len(p.get("right_indent")),
                space_before_pt=float(p["space_before_pt"]) if "space_before_pt" in p and p["space_before_pt"] is not None else None,
                space_after_pt=float(p["space_after_pt"]) if "space_after_pt" in p and p["space_after_pt"] is not None else None,
                line_spacing=line_sp,
                keep_with_next=p.get("keep_with_next"),
                keep_lines=p.get("keep_lines"),
                widow_control=p.get("widow_control"),
                page_break_before=p.get("page_break_before"),
                snap_to_grid=p.get("snap_to_grid"),
            )
            for diag in para_spec.validate(f"/styles/{s_name}/paragraph"):
                if diag.severity == "error":
                    raise ConfigError(f"[{diag.code}] 样式 {s_name} 段落配置错误: {diag.message}")

        styles[s_name] = StyleDefinition(
            name=s_name,
            based_on=s_data.get("based_on"),
            run=run_spec,
            paragraph=para_spec,
        )

    # Roles
    roles: Dict[str, RoleSpec] = {}
    raw_roles = merged_data.get("roles", {})
    for r_name, r_data in raw_roles.items():
        if not is_semantic_role(r_name):
            raise ConfigError(f"{FormatDiagnosticCode.ROLE_UNRESOLVED}: 格式包包含未知语义角色: {r_name}")
        st_ref = r_data.get("style")
        if not st_ref or st_ref not in styles:
            raise ConfigError(f"[{FormatDiagnosticCode.STYLE_NOT_FOUND}] 角色 {r_name} 引用的样式不存在: {st_ref}")
        roles[r_name] = RoleSpec(
            style=st_ref,
            outline_level=r_data.get("outline_level"),
            include_in_toc=bool(r_data.get("include_in_toc", False)),
        )

    # Toc
    raw_toc = merged_data.get("toc", {})
    toc_spec = TocSpec(
        title=raw_toc.get("title", "目  录"),
        title_role=raw_toc.get("title_role", "toc.title"),
        max_level=int(raw_toc.get("max_level", 3)),
        leader=raw_toc.get("leader", "dots"),
        page_number_gap_mm=float(raw_toc.get("page_number_gap_mm", 5.0)),
    )
    if not is_semantic_role(toc_spec.title_role):
        raise ConfigError(f"{FormatDiagnosticCode.ROLE_UNRESOLVED}: toc.title_role 不是已注册语义角色: {toc_spec.title_role}")

    # Header / Footer
    raw_header = merged_data.get("header", {})
    header_spec = HeaderFooterSpec(
        mode=raw_header.get("mode", "none"),
        format=raw_header.get("format", "none"),
        alignment=raw_header.get("alignment", "right"),
        style=raw_header.get("style"),
        text=raw_header.get("text"),
    )
    raw_footer = merged_data.get("footer", {})
    footer_spec = HeaderFooterSpec(
        mode=raw_footer.get("mode", "managed"),
        format=raw_footer.get("format", "number"),
        alignment=raw_footer.get("alignment", "center"),
        style=raw_footer.get("style"),
        text=raw_footer.get("text"),
    )

    # Style inheritance must resolve deterministically as well as package
    # inheritance. A changed package with the same semantic version must still
    # produce a different content hash.
    for style_name, style_def in styles.items():
        if style_def.based_on and style_def.based_on not in styles:
            raise ConfigError(
                f"[{FormatDiagnosticCode.STYLE_NOT_FOUND}] 样式 {style_name} 引用的 based_on 不存在: {style_def.based_on}"
            )
    for style_name in styles:
        style_chain: Set[str] = set()
        current_style = style_name
        while current_style:
            if current_style in style_chain:
                raise ConfigError(f"[{FormatDiagnosticCode.FORMAT_CYCLE}] 样式继承存在循环: {current_style}")
            style_chain.add(current_style)
            current_style = styles[current_style].based_on

    # Content Hash (基于完整规范合并数据生成，不排除同版本的包修改)
    canonical_dict = deepcopy(merged_data)
    content_str = json.dumps(canonical_dict, sort_keys=True, ensure_ascii=False)
    content_hash = hashlib.sha256(content_str.encode("utf-8")).hexdigest()

    return ResolvedFormat(
        id=merged_data["id"],
        version=merged_data["version"],
        page=page_spec,
        styles=styles,
        roles=roles,
        toc=toc_spec,
        header=header_spec,
        footer=footer_spec,
        required_capabilities=req_caps,
        provenance=provenance,
        content_hash=content_hash,
        source_hashes={
            str(path): hashlib.sha256(path.read_bytes()).hexdigest()
            for path, _ in chain
        },
    )
