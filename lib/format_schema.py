# -*- coding: utf-8 -*-
"""
格式包契约与数据模型 (lib/format_schema.py)
定义排版规格、单位转换、几何检查与诊断数据结构。
不依赖 engine、renderer 或 Word 自动化。
"""

import math
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple, Union

from .contracts import CORE_ROLES, SEMANTIC_ROLES, STRUCTURE_TYPES


# -----------------------------------------------------------------------------
# 诊断代码与数据结构
# -----------------------------------------------------------------------------

class FormatDiagnosticCode:
    FORMAT_CYCLE = "FORMAT_CYCLE"
    INVALID_GEOMETRY = "INVALID_GEOMETRY"
    UNSUPPORTED_CAPABILITY = "UNSUPPORTED_CAPABILITY"
    ROLE_UNRESOLVED = "ROLE_UNRESOLVED"
    MAPPING_STALE = "MAPPING_STALE"
    UNSUPPORTED_OBJECT = "UNSUPPORTED_OBJECT"
    FONT_MISSING = "FONT_MISSING"
    UNKNOWN_KEY = "UNKNOWN_KEY"
    INVALID_UNIT = "INVALID_UNIT"
    STYLE_NOT_FOUND = "STYLE_NOT_FOUND"
    INVALID_SCHEMA = "INVALID_SCHEMA"


@dataclass(frozen=True)
class Diagnostic:
    code: str
    severity: str  # "info" | "warning" | "error"
    location: str  # JSON Pointer 或 NodeRef 标识
    message: str


# -----------------------------------------------------------------------------
# 单位转换与度量
# -----------------------------------------------------------------------------

MM_TO_PT = 72.0 / 25.4
PT_TO_TWIP = 20.0
MM_TO_TWIP = MM_TO_PT * PT_TO_TWIP  # ≈ 56.6929133858


@dataclass(frozen=True)
class LengthValue:
    """带单位的长度或缩进量"""
    value: float
    unit: str  # "char" | "pt" | "mm"

    def to_twip(self, base_font_size_pt: float = 12.0) -> int:
        if self.unit == "pt":
            return round(self.value * PT_TO_TWIP)
        elif self.unit == "mm":
            return round(self.value * MM_TO_TWIP)
        elif self.unit == "char":
            # 1 char 缩进在 Word 中按所在段落基准字号的宽（通常 1 个全角字符或磅值）计算
            return round(self.value * base_font_size_pt * PT_TO_TWIP)
        raise ValueError(f"不支持的单位: {self.unit}")

    def to_pt(self, base_font_size_pt: float = 12.0) -> float:
        if self.unit == "pt":
            return self.value
        elif self.unit == "mm":
            return self.value * MM_TO_PT
        elif self.unit == "char":
            return self.value * base_font_size_pt
        raise ValueError(f"不支持的单位: {self.unit}")


@dataclass(frozen=True)
class LineSpacing:
    """段落行距规格"""
    mode: str  # "single" | "multiple" | "exact" | "at_least"
    value: Optional[float] = None  # multiple 时为倍数 (如 1.5)，exact/at_least 时为 pt

    def validate(self, location: str = "") -> List[Diagnostic]:
        diags = []
        if self.mode not in ("single", "multiple", "exact", "at_least"):
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_SCHEMA, "error", f"{location}/mode",
                f"行距模式只支持 single, multiple, exact, at_least，收到: {self.mode}"
            ))
        if self.mode == "multiple":
            if self.value is None or self.value <= 0:
                diags.append(Diagnostic(
                    FormatDiagnosticCode.INVALID_SCHEMA, "error", f"{location}/value",
                    f"multiple 模式行距必须指定大于 0 的倍数值，收到: {self.value}"
                ))
        elif self.mode in ("exact", "at_least"):
            if self.value is None or self.value <= 0:
                diags.append(Diagnostic(
                    FormatDiagnosticCode.INVALID_SCHEMA, "error", f"{location}/value",
                    f"{self.mode} 模式行距必须指定大于 0 的磅值 (pt)，收到: {self.value}"
                ))
        return diags


# -----------------------------------------------------------------------------
# 样式与排版模型
# -----------------------------------------------------------------------------

@dataclass(frozen=True)
class RunStyle:
    """字符运行级别样式"""
    east_asia: Optional[str] = None
    latin: Optional[str] = None
    complex_script: Optional[str] = None
    size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[str] = None  # "auto" 或 6位十六进制颜色


@dataclass(frozen=True)
class ParagraphStyle:
    """段落级别样式"""
    alignment: Optional[str] = None  # "left" | "center" | "right" | "justify"
    first_line_indent: Optional[LengthValue] = None
    hanging_indent: Optional[LengthValue] = None
    left_indent: Optional[LengthValue] = None
    right_indent: Optional[LengthValue] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[LineSpacing] = None
    keep_with_next: Optional[bool] = None
    keep_lines: Optional[bool] = None
    widow_control: Optional[bool] = None
    page_break_before: Optional[bool] = None
    snap_to_grid: Optional[bool] = None

    def validate(self, location: str = "") -> List[Diagnostic]:
        diags = []
        if self.first_line_indent and self.hanging_indent:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_SCHEMA, "error", f"{location}",
                "首行缩进 (first_line_indent) 与悬挂缩进 (hanging_indent) 互斥，不能同时指定"
            ))
        if self.line_spacing:
            diags.extend(self.line_spacing.validate(f"{location}/line_spacing"))
        return diags


@dataclass(frozen=True)
class StyleDefinition:
    """具名样式定义"""
    name: str
    based_on: Optional[str] = None
    run: Optional[RunStyle] = None
    paragraph: Optional[ParagraphStyle] = None


@dataclass(frozen=True)
class PageSpec:
    """页面几何与分节规格"""
    width_mm: float = 210.0
    height_mm: float = 297.0
    orientation: str = "portrait"  # "portrait" | "landscape"
    margin_top_mm: float = 25.4
    margin_bottom_mm: float = 25.4
    margin_left_mm: float = 31.8
    margin_right_mm: float = 31.8
    header_distance_mm: float = 15.0
    footer_distance_mm: float = 15.0
    snap_to_grid: bool = True

    def validate(self, location: str = "/page") -> List[Diagnostic]:
        diags = []
        if self.width_mm <= 0 or self.height_mm <= 0:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_GEOMETRY, "error", location,
                f"页面尺寸必须大于 0: width={self.width_mm}, height={self.height_mm}"
            ))
            return diags

        printable_width = self.width_mm - (self.margin_left_mm + self.margin_right_mm)
        printable_height = self.height_mm - (self.margin_top_mm + self.margin_bottom_mm)

        if printable_width <= 10.0:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_GEOMETRY, "error", location,
                f"有效版心宽度不足 ({printable_width:.1f}mm)，左右边距过大"
            ))
        if printable_height <= 10.0:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_GEOMETRY, "error", location,
                f"有效版心高度不足 ({printable_height:.1f}mm)，上下边距过大"
            ))
        if self.header_distance_mm >= self.margin_top_mm:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_GEOMETRY, "warning", f"{location}/header_distance_mm",
                f"页眉边距 ({self.header_distance_mm}mm) 大于或等于上边距 ({self.margin_top_mm}mm)，可能产生版面遮挡"
            ))
        if self.footer_distance_mm >= self.margin_bottom_mm:
            diags.append(Diagnostic(
                FormatDiagnosticCode.INVALID_GEOMETRY, "warning", f"{location}/footer_distance_mm",
                f"页脚边距 ({self.footer_distance_mm}mm) 大于或等于下边距 ({self.margin_bottom_mm}mm)，可能产生版面遮挡"
            ))
        return diags


@dataclass(frozen=True)
class RoleSpec:
    """语义角色到样式的绑定规格"""
    style: str
    outline_level: Optional[int] = None  # 1-9
    include_in_toc: bool = False


@dataclass(frozen=True)
class TocSpec:
    """目录生成规则"""
    title: str = "目  录"
    title_role: str = "toc.title"
    max_level: int = 3
    leader: str = "dots"  # "dots" | "hyphens" | "underline" | "none"
    page_number_gap_mm: float = 5.0


@dataclass(frozen=True)
class HeaderFooterSpec:
    """页眉与页脚定义"""
    mode: str = "managed"  # "managed" | "none" | "source"
    format: str = "number"  # "number" | "dash_number" | "none"
    alignment: str = "center"  # "left" | "center" | "right"
    style: Optional[str] = None
    text: Optional[str] = None


@dataclass(frozen=True)
class FormattingPolicy:
    """排版与重排策略"""
    mode: str = "restyle"  # "preserve" | "restyle" | "mixed"
    role_map: Optional[str] = None
    on_unmapped: str = "error"  # "preserve" | "error"
    page_policy: str = "target"  # "target" | "source"
    inline_emphasis: str = "preserve"  # "preserve" | "target"
    source_toc: str = "preserve"  # "preserve" | "error"
    scopes: List[Dict[str, Any]] = field(default_factory=list)


@dataclass(frozen=True)
class FormatPackage:
    """未解析继承的格式包原始数据结构"""
    format_schema_version: int
    id: str
    version: str
    extends: Optional[str] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    page: Optional[PageSpec] = None
    styles: Dict[str, StyleDefinition] = field(default_factory=dict)
    roles: Dict[str, RoleSpec] = field(default_factory=dict)
    toc: Optional[TocSpec] = None
    header: Optional[HeaderFooterSpec] = None
    footer: Optional[HeaderFooterSpec] = None
    required_capabilities: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class ResolvedFormat:
    """完全继承合并后的最终排版格式上下文"""
    id: str
    version: str
    page: PageSpec
    styles: Dict[str, StyleDefinition]
    roles: Dict[str, RoleSpec]
    toc: TocSpec
    header: HeaderFooterSpec
    footer: HeaderFooterSpec
    required_capabilities: List[str]
    provenance: Dict[str, str]  # 属性 JSON Pointer -> 来源描述
    content_hash: str  # 规范化内容摘要哈希
    source_hashes: Dict[str, str] = field(default_factory=dict)  # 继承链上的格式包文件哈希
