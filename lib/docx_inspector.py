# -*- coding: utf-8 -*-
"""
只读 DOCX 分析、NodeRef 规范身份与有效排版属性求值 (lib/docx_inspector.py)
实现基于 OOXML 规范的有向图求值：
docDefaults -> Theme -> Style basedOn 链 -> Direct Formatting
支持资源限额安全检查与复杂/受保护对象识别。
"""

import os
import posixpath
import re
import zipfile
import hashlib
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Set, Union
import lxml.etree

from docx.oxml.ns import qn
from lib.config import ConfigError
from lib.format_schema import (
    FormatDiagnosticCode,
    Diagnostic,
    LengthValue,
    LineSpacing,
    RunStyle,
    ParagraphStyle,
    PageSpec,
)
from lib.contracts import text_hash


# -----------------------------------------------------------------------------
# 节点身份与资源限额
# -----------------------------------------------------------------------------


def _identity_text_without_field_cache(element: Any) -> str:
    """Return author text while excluding Word-controlled field caches.

    Word is allowed to rewrite a field result (for example a PAGEREF value)
    during pagination.  That result is not an author-authored identity signal,
    so NodeRef hashes must use the surrounding text and field instruction but
    not the cached result.
    """
    text_parts: List[str] = []
    complex_field_depth = 0
    simple_fields = {node for node in element.iter() if node.tag == qn("w:fldSimple")}
    for node in element.iter():
        if node.tag == qn("w:fldChar"):
            field_type = (node.get(qn("w:fldCharType")) or "").lower()
            if field_type == "begin":
                complex_field_depth += 1
            elif field_type == "end":
                complex_field_depth = max(0, complex_field_depth - 1)
            continue
        if node.tag != qn("w:t"):
            continue
        if complex_field_depth:
            continue
        if any(parent in simple_fields for parent in node.iterancestors()):
            continue
        text_parts.append(node.text or "")
    return "".join(text_parts).strip()

@dataclass(frozen=True)
class NodeRef:
    """文档中任意块级元素的唯一、规范身份"""
    source_sha256: str
    part_uri: str      # 例如 "word/document.xml"
    element_path: str  # 例如 "/w:document/w:body/w:p[1]" 或 "/w:document/w:body/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]"
    text_hash: Optional[str] = None  # 节点正文归一化哈希；仅用于绑定校验，不替代路径身份


@dataclass(frozen=True)
class DocumentInspectionLimits:
    """ZIP 与 XML 安全限额配置"""
    max_uncompressed_bytes: int = 256 * 1024 * 1024  # 256 MiB
    max_single_xml_bytes: int = 32 * 1024 * 1024     # 32 MiB
    max_part_count: int = 10000
    max_compression_ratio: float = 200.0


# -----------------------------------------------------------------------------
# 检查数据结构
# -----------------------------------------------------------------------------

@dataclass(frozen=True)
class ThemeFonts:
    """从 word/theme/theme1.xml 解析的主题字体映射"""
    major_east_asia: Optional[str] = None
    major_ascii: Optional[str] = None
    minor_east_asia: Optional[str] = None
    minor_ascii: Optional[str] = None

    def resolve_theme_font(self, theme_key: Optional[str]) -> Optional[str]:
        if not theme_key:
            return None
        k = theme_key.lower()
        if "majoreastasia" in k:
            return self.major_east_asia
        elif "minoreastasia" in k:
            return self.minor_east_asia
        elif "majorascii" in k or "majorhansi" in k:
            return self.major_ascii
        elif "minorascii" in k or "minorhansi" in k:
            return self.minor_ascii
        return None


@dataclass(frozen=True)
class StyleNode:
    """样式有向图节点"""
    style_id: str
    style_type: str  # "paragraph" | "character" | "table"
    name: Optional[str] = None
    based_on: Optional[str] = None
    next_style: Optional[str] = None
    outline_level: Optional[int] = None
    run_style: Optional[RunStyle] = None
    paragraph_style: Optional[ParagraphStyle] = None


@dataclass(frozen=True)
class BlockInspection:
    """单个块级节点（段落/表格）的只读检查结果"""
    node: NodeRef
    structure_type: str      # "paragraph" | "table" | "section_break"
    story_type: str          # "body" | "header" | "footer" | "footnote" | "endnote"
    visible_text: str
    p_style_id: Optional[str]
    p_style_name: Optional[str]
    outline_level: Optional[int]
    effective_run: Optional[RunStyle]
    effective_paragraph: Optional[ParagraphStyle]
    protected_objects: List[str]  # ["field_complex", "omml_math", "footnote_ref", "hyperlink", "image"]
    runs_count: int
    effective_runs: Tuple[RunStyle, ...] = ()
    # Logical text intervals retain spaces, tabs, and breaks so inline
    # properties remain independent of Word's run splitting/merging while
    # keeping deterministic character coordinates.
    effective_run_spans: Tuple[Dict[str, Any], ...] = ()
    # Direct run-level emphasis is separate from effective style.  A heading
    # style may be bold by inheritance without that being authorial inline
    # emphasis that restyle mode must preserve.
    inline_emphasis_spans: Tuple[Dict[str, Any], ...] = ()
    outline_source: Optional[str] = None  # "direct" or "style", for confidence evidence


@dataclass(frozen=True)
class SectionInspection:
    """分节配置检查"""
    section_index: int
    page_spec: PageSpec
    section_type: str        # "continuous" | "nextPage" | "oddPage" | "evenPage"
    has_header: bool
    has_footer: bool
    page_start: Optional[int] = None
    page_format: Optional[str] = None  # "decimal" | "upperRoman" 等
    header_references: Tuple[Dict[str, Any], ...] = ()
    footer_references: Tuple[Dict[str, Any], ...] = ()
    different_first_page: bool = False


@dataclass(frozen=True)
class SectionStoryBinding:
    section_index: int
    story_type: str
    variant: str
    part_uri: str
    linked_to_previous: bool = False


@dataclass(frozen=True)
class StoryInspection:
    story_type: str
    variant: str
    part_uri: str
    blocks: Tuple[BlockInspection, ...] = ()


@dataclass(frozen=True)
class DocumentInspection:
    """完整的只读文档检查结果"""
    source_sha256: str
    file_path: str
    file_size: int
    blocks: List[BlockInspection]
    styles: Dict[str, StyleNode]
    theme_fonts: ThemeFonts
    sections: List[SectionInspection]
    diagnostics: List[Diagnostic]
    stories: List[StoryInspection] = field(default_factory=list)
    section_bindings: List[SectionStoryBinding] = field(default_factory=list)


# -----------------------------------------------------------------------------
# XML 解析辅助工具
# -----------------------------------------------------------------------------

def _get_attr(elem: Optional[lxml.etree._Element], name: str) -> Optional[str]:
    if elem is None:
        return None
    return elem.get(qn(name))


def _logical_run_text(run: Any) -> str:
    """Return deterministic run coordinates, retaining whitespace controls."""
    parts: List[str] = []
    for node in run.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag in {qn("w:br"), qn("w:cr")}:
            parts.append("\n")
    return "".join(parts)


def _parse_toggle(rPr: Optional[lxml.etree._Element], tag_name: str) -> Optional[bool]:
    """
    解析 OOXML toggle 属性 (如 w:b, w:i, w:caps, w:strike)
    <w:b/> 或 <w:b w:val="true|1|on"/> -> True
    <w:b w:val="false|0|off"/> -> False (显式取消!)
    不存在 -> None (未声明，继承父层)
    """
    if rPr is None:
        return None
    el = rPr.find(qn(tag_name))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return True
    val_lower = val.strip().lower()
    if val_lower in ("true", "1", "on"):
        return True
    elif val_lower in ("false", "0", "off"):
        return False
    return True


def _parse_half_points(elem: Optional[lxml.etree._Element], tag_name: str) -> Optional[float]:
    if elem is None:
        return None
    el = elem.find(qn(tag_name))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is not None:
        try:
            return float(val) / 2.0
        except ValueError:
            return None
    return None


def _parse_twips_to_pt(elem: Optional[lxml.etree._Element], tag_name: str, attr_name: str = "w:val") -> Optional[float]:
    if elem is None:
        return None
    el = elem.find(qn(tag_name))
    if el is None:
        return None
    val = el.get(qn(attr_name))
    if val is not None:
        try:
            return float(val) / 20.0
        except ValueError:
            return None
    return None


def _parse_r_fonts(rPr: Optional[lxml.etree._Element], theme_fonts: ThemeFonts) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """解析字符运行字体 (east_asia, latin, complex_script)"""
    if rPr is None:
        return None, None, None
    fonts_el = rPr.find(qn("w:rFonts"))
    if fonts_el is None:
        return None, None, None

    # East Asia
    ea = fonts_el.get(qn("w:eastAsia"))
    if not ea:
        ea = theme_fonts.resolve_theme_font(fonts_el.get(qn("w:eastAsiaTheme")))

    # Latin (ascii / hAnsi)
    lat = fonts_el.get(qn("w:ascii")) or fonts_el.get(qn("w:hAnsi"))
    if not lat:
        lat = theme_fonts.resolve_theme_font(fonts_el.get(qn("w:asciiTheme")) or fonts_el.get(qn("w:hAnsiTheme")))

    # Complex Script
    cs = fonts_el.get(qn("w:cs"))
    if not cs:
        cs = theme_fonts.resolve_theme_font(fonts_el.get(qn("w:cstheme")))

    return ea, lat, cs


def _parse_color(rPr: Optional[lxml.etree._Element]) -> Optional[str]:
    if rPr is None:
        return None
    col_el = rPr.find(qn("w:color"))
    if col_el is None:
        return None
    val = col_el.get(qn("w:val"))
    if val:
        return val.strip().lower()
    return None


# -----------------------------------------------------------------------------
# 样式图与求值引擎 (Effective Style Evaluator)
# -----------------------------------------------------------------------------

class EffectiveStyleEvaluator:
    """计算段落与 Run 的最终有效排版属性"""

    def __init__(
        self,
        styles: Dict[str, StyleNode],
        theme_fonts: ThemeFonts,
        doc_defaults_rPr: Optional[lxml.etree._Element] = None,
        doc_defaults_pPr: Optional[lxml.etree._Element] = None,
    ):
        self.styles = styles
        self.theme_fonts = theme_fonts
        self.doc_defaults_rPr = doc_defaults_rPr
        self.doc_defaults_pPr = doc_defaults_pPr
        self._style_run_cache: Dict[str, RunStyle] = {}
        self._style_para_cache: Dict[str, ParagraphStyle] = {}

    def get_style_chain(self, style_id: str) -> List[StyleNode]:
        """获取基于 basedOn 的继承链（从最顶层父样式到当前样式）"""
        chain = []
        visited = set()
        curr_id = style_id
        while curr_id and curr_id in self.styles:
            if curr_id in visited:
                break
            visited.add(curr_id)
            node = self.styles[curr_id]
            chain.append(node)
            curr_id = node.based_on or ""
        chain.reverse()
        return chain

    def compute_effective_paragraph(
        self,
        p_style_id: Optional[str],
        direct_pPr: Optional[lxml.etree._Element],
    ) -> ParagraphStyle:
        """层叠计算段落有效属性: docDefaults -> style chain -> direct_pPr"""
        # 1. 初始属性来自 docDefaults
        alignment = None
        space_before_pt = None
        space_after_pt = None
        line_spacing = None
        first_line_indent = None
        hanging_indent = None
        left_indent = None
        right_indent = None
        keep_with_next = None
        keep_lines = None
        page_break_before = None
        widow_control = True
        snap_to_grid = True

        def apply_ppr(ppr: Optional[lxml.etree._Element]) -> None:
            """应用一层 pPr；调用顺序决定继承优先级。"""
            nonlocal alignment, space_before_pt, space_after_pt, line_spacing
            nonlocal first_line_indent, hanging_indent, left_indent, right_indent
            nonlocal keep_with_next, keep_lines, page_break_before, widow_control, snap_to_grid
            if ppr is None:
                return
            jc_el = ppr.find(qn("w:jc"))
            if jc_el is not None:
                val = jc_el.get(qn("w:val"))
                alignment = "justify" if val == "both" else val
            sp_el = ppr.find(qn("w:spacing"))
            if sp_el is not None:
                for attr, target in (("before", "before"), ("after", "after")):
                    value = sp_el.get(qn(f"w:{attr}"))
                    if value is not None:
                        try:
                            if target == "before":
                                space_before_pt = float(value) / 20.0
                            else:
                                space_after_pt = float(value) / 20.0
                        except ValueError:
                            pass
                line_val = sp_el.get(qn("w:line"))
                if line_val is not None:
                    try:
                        line_value = float(line_val)
                        rule = sp_el.get(qn("w:lineRule"), "auto")
                        if rule == "auto":
                            line_spacing = LineSpacing(mode="multiple", value=round(line_value / 240.0, 2))
                        elif rule == "exact":
                            line_spacing = LineSpacing(mode="exact", value=round(line_value / 20.0, 1))
                        elif rule == "atLeast":
                            line_spacing = LineSpacing(mode="at_least", value=round(line_value / 20.0, 1))
                    except ValueError:
                        pass
            ind_el = ppr.find(qn("w:ind"))
            if ind_el is not None:
                fl_chars, fl_pt = ind_el.get(qn("w:firstLineChars")), ind_el.get(qn("w:firstLine"))
                hg_chars, hg_pt = ind_el.get(qn("w:hangingChars")), ind_el.get(qn("w:hanging"))
                if fl_chars is not None:
                    first_line_indent = LengthValue(value=float(fl_chars) / 100.0, unit="char")
                    hanging_indent = None
                elif fl_pt is not None:
                    first_line_indent = LengthValue(value=float(fl_pt) / 20.0, unit="pt")
                    hanging_indent = None
                elif hg_chars is not None:
                    hanging_indent = LengthValue(value=float(hg_chars) / 100.0, unit="char")
                    first_line_indent = None
                elif hg_pt is not None:
                    hanging_indent = LengthValue(value=float(hg_pt) / 20.0, unit="pt")
                    first_line_indent = None
                left = ind_el.get(qn("w:left"))
                right = ind_el.get(qn("w:right"))
                if left is not None:
                    left_indent = LengthValue(value=float(left) / 20.0, unit="pt")
                if right is not None:
                    right_indent = LengthValue(value=float(right) / 20.0, unit="pt")
            for tag, target in (
                ("keepNext", "keep_with_next"),
                ("keepLines", "keep_lines"),
                ("pageBreakBefore", "page_break_before"),
                ("widowControl", "widow_control"),
            ):
                if ppr.find(qn(f"w:{tag}")) is not None:
                    value = _parse_toggle(ppr, f"w:{tag}")
                    if target == "keep_with_next":
                        keep_with_next = value
                    elif target == "keep_lines":
                        keep_lines = value
                    elif target == "page_break_before":
                        page_break_before = value
                    else:
                        widow_control = value
            if ppr.find(qn("w:snapToGrid")) is not None:
                snap_to_grid = _parse_toggle(ppr, "w:snapToGrid")

        # 2. 文档默认值先于样式链。
        apply_ppr(self.doc_defaults_pPr)

        # 3. 叠加上级样式链
        if p_style_id:
            for s_node in self.get_style_chain(p_style_id):
                if s_node.paragraph_style:
                    ps = s_node.paragraph_style
                    if ps.alignment is not None:
                        alignment = ps.alignment
                    if ps.space_before_pt is not None:
                        space_before_pt = ps.space_before_pt
                    if ps.space_after_pt is not None:
                        space_after_pt = ps.space_after_pt
                    if ps.line_spacing is not None:
                        line_spacing = ps.line_spacing
                    if ps.first_line_indent is not None:
                        first_line_indent = ps.first_line_indent
                        hanging_indent = None
                    if ps.hanging_indent is not None:
                        hanging_indent = ps.hanging_indent
                        first_line_indent = None
                    if ps.left_indent is not None:
                        left_indent = ps.left_indent
                    if ps.right_indent is not None:
                        right_indent = ps.right_indent
                    if ps.keep_with_next is not None:
                        keep_with_next = ps.keep_with_next
                    if ps.keep_lines is not None:
                        keep_lines = ps.keep_lines
                    if ps.page_break_before is not None:
                        page_break_before = ps.page_break_before
                    if ps.widow_control is not None:
                        widow_control = ps.widow_control
                    if ps.snap_to_grid is not None:
                        snap_to_grid = ps.snap_to_grid

        # 3. 叠加直接格式 (direct_pPr)
        if direct_pPr is not None:
            # jc
            jc_el = direct_pPr.find(qn("w:jc"))
            if jc_el is not None:
                val = jc_el.get(qn("w:val"))
                if val == "both":
                    alignment = "justify"
                elif val in ("left", "center", "right", "justify"):
                    alignment = val

            # spacing
            sp_el = direct_pPr.find(qn("w:spacing"))
            if sp_el is not None:
                b_val = sp_el.get(qn("w:before"))
                if b_val is not None:
                    try:
                        space_before_pt = float(b_val) / 20.0
                    except ValueError:
                        pass
                a_val = sp_el.get(qn("w:after"))
                if a_val is not None:
                    try:
                        space_after_pt = float(a_val) / 20.0
                    except ValueError:
                        pass
                line_val = sp_el.get(qn("w:line"))
                line_rule = sp_el.get(qn("w:lineRule"), "auto")
                if line_val is not None:
                    try:
                        lv = float(line_val)
                        if line_rule == "auto":
                            line_spacing = LineSpacing(mode="multiple", value=round(lv / 240.0, 2))
                        elif line_rule == "exact":
                            line_spacing = LineSpacing(mode="exact", value=round(lv / 20.0, 1))
                        elif line_rule == "atLeast":
                            line_spacing = LineSpacing(mode="at_least", value=round(lv / 20.0, 1))
                    except ValueError:
                        pass

            # ind
            ind_el = direct_pPr.find(qn("w:ind"))
            if ind_el is not None:
                fl_chars = ind_el.get(qn("w:firstLineChars"))
                fl_pt = ind_el.get(qn("w:firstLine"))
                hg_chars = ind_el.get(qn("w:hangingChars"))
                hg_pt = ind_el.get(qn("w:hanging"))
                left_pt = ind_el.get(qn("w:left"))
                right_pt = ind_el.get(qn("w:right"))

                if fl_chars is not None:
                    first_line_indent = LengthValue(value=float(fl_chars) / 100.0, unit="char")
                    hanging_indent = None
                elif fl_pt is not None:
                    first_line_indent = LengthValue(value=float(fl_pt) / 20.0, unit="pt")
                    hanging_indent = None
                elif hg_chars is not None:
                    hanging_indent = LengthValue(value=float(hg_chars) / 100.0, unit="char")
                    first_line_indent = None
                elif hg_pt is not None:
                    hanging_indent = LengthValue(value=float(hg_pt) / 20.0, unit="pt")
                    first_line_indent = None

                if left_pt is not None:
                    left_indent = LengthValue(value=float(left_pt) / 20.0, unit="pt")
                if right_pt is not None:
                    right_indent = LengthValue(value=float(right_pt) / 20.0, unit="pt")

            # toggles
            if direct_pPr.find(qn("w:keepNext")) is not None:
                keep_with_next = _parse_toggle(direct_pPr, "w:keepNext")
            if direct_pPr.find(qn("w:keepLines")) is not None:
                keep_lines = _parse_toggle(direct_pPr, "w:keepLines")
            if direct_pPr.find(qn("w:pageBreakBefore")) is not None:
                page_break_before = _parse_toggle(direct_pPr, "w:pageBreakBefore")
            if direct_pPr.find(qn("w:widowControl")) is not None:
                widow_control = _parse_toggle(direct_pPr, "w:widowControl")
            if direct_pPr.find(qn("w:snapToGrid")) is not None:
                snap_to_grid = _parse_toggle(direct_pPr, "w:snapToGrid")

        return ParagraphStyle(
            alignment=alignment,
            first_line_indent=first_line_indent,
            hanging_indent=hanging_indent,
            left_indent=left_indent,
            right_indent=right_indent,
            space_before_pt=space_before_pt,
            space_after_pt=space_after_pt,
            line_spacing=line_spacing,
            keep_with_next=keep_with_next,
            keep_lines=keep_lines,
            page_break_before=page_break_before,
            widow_control=widow_control,
            snap_to_grid=snap_to_grid,
        )

    def compute_effective_run(
        self,
        p_style_id: Optional[str],
        r_style_id: Optional[str],
        direct_rPr: Optional[lxml.etree._Element],
    ) -> RunStyle:
        """层叠计算字符运行有效属性: docDefaults -> pStyle -> rStyle -> direct_rPr"""
        east_asia = None
        latin = None
        complex_script = None
        size_pt = None
        bold = False
        italic = False
        # OOXML 未声明颜色时的有效默认值是 auto，而不是“无法验证”。
        color = "auto"

        # 1. docDefaults
        if self.doc_defaults_rPr is not None:
            ea, lat, cs = _parse_r_fonts(self.doc_defaults_rPr, self.theme_fonts)
            if ea: east_asia = ea
            if lat: latin = lat
            if cs: complex_script = cs
            sz = _parse_half_points(self.doc_defaults_rPr, "w:sz")
            if sz is not None: size_pt = sz
            b = _parse_toggle(self.doc_defaults_rPr, "w:b")
            if b is not None: bold = b
            it = _parse_toggle(self.doc_defaults_rPr, "w:i")
            if it is not None: italic = it
            c = _parse_color(self.doc_defaults_rPr)
            if c: color = c

        # 2. 段落样式及其基于的样式链
        if p_style_id:
            for s_node in self.get_style_chain(p_style_id):
                if s_node.run_style:
                    rs = s_node.run_style
                    if rs.east_asia is not None: east_asia = rs.east_asia
                    if rs.latin is not None: latin = rs.latin
                    if rs.complex_script is not None: complex_script = rs.complex_script
                    if rs.size_pt is not None: size_pt = rs.size_pt
                    if rs.bold is not None: bold = rs.bold
                    if rs.italic is not None: italic = rs.italic
                    if rs.color is not None: color = rs.color

        # 3. 字符样式 (rStyle) 及其继承链
        if r_style_id:
            for s_node in self.get_style_chain(r_style_id):
                if s_node.run_style:
                    rs = s_node.run_style
                    if rs.east_asia is not None: east_asia = rs.east_asia
                    if rs.latin is not None: latin = rs.latin
                    if rs.complex_script is not None: complex_script = rs.complex_script
                    if rs.size_pt is not None: size_pt = rs.size_pt
                    if rs.bold is not None: bold = rs.bold
                    if rs.italic is not None: italic = rs.italic
                    if rs.color is not None: color = rs.color

        # 4. 直接格式覆盖 (direct_rPr)
        if direct_rPr is not None:
            ea, lat, cs = _parse_r_fonts(direct_rPr, self.theme_fonts)
            if ea: east_asia = ea
            if lat: latin = lat
            if cs: complex_script = cs
            sz = _parse_half_points(direct_rPr, "w:sz")
            if sz is not None: size_pt = sz
            b = _parse_toggle(direct_rPr, "w:b")
            if b is not None: bold = b  # 包含了显式 w:b w:val="0" -> False!
            it = _parse_toggle(direct_rPr, "w:i")
            if it is not None: italic = it
            c = _parse_color(direct_rPr)
            if c: color = c

        return RunStyle(
            east_asia=east_asia,
            latin=latin,
            complex_script=complex_script,
            size_pt=size_pt,
            bold=bold,
            italic=italic,
            color=color,
        )


# -----------------------------------------------------------------------------
# 主检查器 (DOCX Inspector)
# -----------------------------------------------------------------------------

def _parse_theme_fonts(theme_xml_bytes: Optional[bytes]) -> ThemeFonts:
    if not theme_xml_bytes:
        return ThemeFonts()
    try:
        root = lxml.etree.fromstring(theme_xml_bytes)
    except Exception:
        return ThemeFonts()

    # namespaces
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    font_scheme = root.find(".//a:fontScheme", namespaces=ns)
    if font_scheme is None:
        return ThemeFonts()

    major = font_scheme.find("a:majorFont", namespaces=ns)
    minor = font_scheme.find("a:minorFont", namespaces=ns)

    def get_font_names(parent):
        if parent is None:
            return None, None
        ea_el = parent.find("a:ea", namespaces=ns)
        latin_el = parent.find("a:latin", namespaces=ns)
        ea_typeface = ea_el.get("typeface") if ea_el is not None else None
        latin_typeface = latin_el.get("typeface") if latin_el is not None else None
        return ea_typeface, latin_typeface

    major_ea, major_lat = get_font_names(major)
    minor_ea, minor_lat = get_font_names(minor)

    return ThemeFonts(
        major_east_asia=major_ea,
        major_ascii=major_lat,
        minor_east_asia=minor_ea,
        minor_ascii=minor_lat,
    )


def _parse_styles_xml(
    styles_xml_bytes: bytes,
    theme_fonts: ThemeFonts,
) -> Tuple[Dict[str, StyleNode], Optional[lxml.etree._Element], Optional[lxml.etree._Element]]:
    root = lxml.etree.fromstring(styles_xml_bytes)
    styles: Dict[str, StyleNode] = {}

    doc_defaults = root.find(qn("w:docDefaults"))
    doc_def_rPr = None
    doc_def_pPr = None
    if doc_defaults is not None:
        rDef = doc_defaults.find(qn("w:rPrDefault"))
        if rDef is not None:
            doc_def_rPr = rDef.find(qn("w:rPr"))
        pDef = doc_defaults.find(qn("w:pPrDefault"))
        if pDef is not None:
            doc_def_pPr = pDef.find(qn("w:pPr"))

    for s_el in root.findall(qn("w:style")):
        sid = s_el.get(qn("w:styleId"))
        if not sid:
            continue
        stype = s_el.get(qn("w:type"), "paragraph")
        name_el = s_el.find(qn("w:name"))
        sname = name_el.get(qn("w:val")) if name_el is not None else sid
        based_el = s_el.find(qn("w:basedOn"))
        based_on = based_el.get(qn("w:val")) if based_el is not None else None
        next_el = s_el.find(qn("w:next"))
        next_style = next_el.get(qn("w:val")) if next_el is not None else None

        rPr = s_el.find(qn("w:rPr"))
        pPr = s_el.find(qn("w:pPr"))

        run_spec = None
        if rPr is not None:
            ea, lat, cs = _parse_r_fonts(rPr, theme_fonts)
            run_spec = RunStyle(
                east_asia=ea,
                latin=lat,
                complex_script=cs,
                size_pt=_parse_half_points(rPr, "w:sz"),
                bold=_parse_toggle(rPr, "w:b"),
                italic=_parse_toggle(rPr, "w:i"),
                color=_parse_color(rPr),
            )

        para_spec = None
        outline_lvl = None
        if pPr is not None:
            out_el = pPr.find(qn("w:outlineLvl"))
            if out_el is not None:
                try:
                    outline_lvl = int(out_el.get(qn("w:val"))) + 1
                except (ValueError, TypeError):
                    pass

            jc_el = pPr.find(qn("w:jc"))
            alignment = jc_el.get(qn("w:val")) if jc_el is not None else None
            if alignment == "both":
                alignment = "justify"

            sp_before = _parse_twips_to_pt(pPr, "w:spacing", "w:before")
            sp_after = _parse_twips_to_pt(pPr, "w:spacing", "w:after")

            # line spacing
            sp_el = pPr.find(qn("w:spacing"))
            line_sp = None
            if sp_el is not None:
                lv = sp_el.get(qn("w:line"))
                lr = sp_el.get(qn("w:lineRule"), "auto")
                if lv:
                    try:
                        val_fl = float(lv)
                        if lr == "auto":
                            line_sp = LineSpacing(mode="multiple", value=round(val_fl / 240.0, 2))
                        elif lr == "exact":
                            line_sp = LineSpacing(mode="exact", value=round(val_fl / 20.0, 1))
                        elif lr == "atLeast":
                            line_sp = LineSpacing(mode="at_least", value=round(val_fl / 20.0, 1))
                    except ValueError:
                        pass

            para_spec = ParagraphStyle(
                alignment=alignment,
                space_before_pt=sp_before,
                space_after_pt=sp_after,
                line_spacing=line_sp,
                keep_with_next=_parse_toggle(pPr, "w:keepNext"),
                keep_lines=_parse_toggle(pPr, "w:keepLines"),
                page_break_before=_parse_toggle(pPr, "w:pageBreakBefore"),
                widow_control=_parse_toggle(pPr, "w:widowControl"),
                snap_to_grid=_parse_toggle(pPr, "w:snapToGrid"),
            )

            ind_el = pPr.find(qn("w:ind"))
            if ind_el is not None:
                def parse_length(chars_name: str, pt_name: str):
                    chars = ind_el.get(qn(chars_name))
                    pts = ind_el.get(qn(pt_name))
                    if chars is not None:
                        try:
                            return LengthValue(float(chars) / 100.0, "char")
                        except ValueError:
                            return None
                    if pts is not None:
                        try:
                            return LengthValue(float(pts) / 20.0, "pt")
                        except ValueError:
                            return None
                    return None
                first = parse_length("w:firstLineChars", "w:firstLine")
                hanging = parse_length("w:hangingChars", "w:hanging")
                left = parse_length("w:leftChars", "w:left")
                right = parse_length("w:rightChars", "w:right")
                para_spec = ParagraphStyle(
                    **{**para_spec.__dict__,
                       "first_line_indent": first,
                       "hanging_indent": hanging,
                       "left_indent": left,
                       "right_indent": right}
                )

        styles[sid] = StyleNode(
            style_id=sid,
            style_type=stype,
            name=sname,
            based_on=based_on,
            next_style=next_style,
            outline_level=outline_lvl,
            run_style=run_spec,
            paragraph_style=para_spec,
        )

    return styles, doc_def_rPr, doc_def_pPr


def _parse_section_spec(sectPr: lxml.etree._Element, section_idx: int) -> SectionInspection:
    pgSz = sectPr.find(qn("w:pgSz"))
    w_mm = 210.0
    h_mm = 297.0
    orientation = "portrait"
    if pgSz is not None:
        try:
            w_twip = float(pgSz.get(qn("w:w"), "11906"))
            h_twip = float(pgSz.get(qn("w:h"), "16838"))
            w_mm = round(w_twip / 20.0 * 25.4 / 72.0, 1)
            h_mm = round(h_twip / 20.0 * 25.4 / 72.0, 1)
        except ValueError:
            pass
        if pgSz.get(qn("w:orient")) == "landscape":
            orientation = "landscape"

    pgMar = sectPr.find(qn("w:pgMar"))
    top_mm = 25.4
    bot_mm = 25.4
    left_mm = 31.8
    right_mm = 31.8
    header_mm = 15.0
    footer_mm = 15.0
    if pgMar is not None:
        def twip_to_mm(val_str: Optional[str], default_mm: float) -> float:
            if not val_str: return default_mm
            try: return round(float(val_str) / 20.0 * 25.4 / 72.0, 1)
            except ValueError: return default_mm

        top_mm = twip_to_mm(pgMar.get(qn("w:top")), top_mm)
        bot_mm = twip_to_mm(pgMar.get(qn("w:bottom")), bot_mm)
        left_mm = twip_to_mm(pgMar.get(qn("w:left")), left_mm)
        right_mm = twip_to_mm(pgMar.get(qn("w:right")), right_mm)
        header_mm = twip_to_mm(pgMar.get(qn("w:header")), header_mm)
        footer_mm = twip_to_mm(pgMar.get(qn("w:footer")), footer_mm)

    page_spec = PageSpec(
        width_mm=w_mm,
        height_mm=h_mm,
        orientation=orientation,
        margin_top_mm=top_mm,
        margin_bottom_mm=bot_mm,
        margin_left_mm=left_mm,
        margin_right_mm=right_mm,
        header_distance_mm=header_mm,
        footer_distance_mm=footer_mm,
    )

    type_el = sectPr.find(qn("w:type"))
    stype = type_el.get(qn("w:val"), "nextPage") if type_el is not None else "nextPage"
    header_references = tuple(
        {
            "variant": item.get(qn("w:type"), "default"),
            "relationship_id": item.get(qn("r:id")),
        }
        for item in sectPr.findall(qn("w:headerReference"))
    )
    footer_references = tuple(
        {
            "variant": item.get(qn("w:type"), "default"),
            "relationship_id": item.get(qn("r:id")),
        }
        for item in sectPr.findall(qn("w:footerReference"))
    )
    has_header = bool(header_references)
    has_footer = bool(footer_references)

    pgNum = sectPr.find(qn("w:pgNumType"))
    p_start = None
    p_fmt = None
    if pgNum is not None:
        s_val = pgNum.get(qn("w:start"))
        if s_val:
            try: p_start = int(s_val)
            except ValueError: pass
        p_fmt = pgNum.get(qn("w:fmt"))

    return SectionInspection(
        section_index=section_idx,
        page_spec=page_spec,
        section_type=stype,
        has_header=has_header,
        has_footer=has_footer,
        page_start=p_start,
        page_format=p_fmt,
        header_references=header_references,
        footer_references=footer_references,
        different_first_page=bool(
            sectPr.find(qn("w:titlePg")) is not None
        ),
    )


def inspect_docx(
    path: Union[str, Path],
    limits: Optional[DocumentInspectionLimits] = None,
) -> DocumentInspection:
    """
    只读检查 DOCX 样例文档。
    - 绝不创建磁盘缓存
    - 绝不调用 Word、模型或网络
    - 精确求值有效排版属性与受保护对象
    """
    file_path = Path(path).resolve()
    if not file_path.is_file():
        raise ConfigError(f"待分析 DOCX 文档不存在: {file_path}")

    # 拒绝 docm / doc 二进制
    suffix = file_path.suffix.lower()
    if suffix in (".docm", ".doc"):
        raise ConfigError(f"安全策略拒绝分析宏文档或旧版二进制文档 ({suffix}): {file_path.name}")

    active_limits = limits or DocumentInspectionLimits()

    # 1. 计算文件哈希与文件大小
    hasher = hashlib.sha256()
    file_size = file_path.stat().st_size
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    source_sha256 = hasher.hexdigest()

    diagnostics: List[Diagnostic] = []

    # 2. 安全检查 ZIP 容器
    try:
        zf = zipfile.ZipFile(file_path, "r")
    except (zipfile.BadZipFile, OSError) as exc:
        raise ConfigError(f"无法以 ZIP 格式读取 DOCX 包: {exc}") from exc

    with zf:
        infolist = zf.infolist()
        if len(infolist) > active_limits.max_part_count:
            raise ConfigError(
                f"[{FormatDiagnosticCode.UNSUPPORTED_OBJECT}] 文档部件数量 ({len(infolist)}) 超过安全上限 ({active_limits.max_part_count})"
            )

        total_uncompressed = 0
        for info in infolist:
            # 路径逃逸检查
            if info.filename.startswith("/") or ".." in info.filename:
                raise ConfigError(f"DOCX 包中包含非法路径逃逸部件: {info.filename}")

            total_uncompressed += info.file_size
            if info.file_size > active_limits.max_single_xml_bytes:
                raise ConfigError(
                    f"[{FormatDiagnosticCode.UNSUPPORTED_OBJECT}] 单部件解压大小 ({info.file_size / 1024 / 1024:.1f} MiB) 超过安全上限 ({active_limits.max_single_xml_bytes / 1024 / 1024:.1f} MiB): {info.filename}"
                )

            # 压缩比检查 (仅针对有实际压缩且解压后 > 1KB 的文件)
            if info.compress_size > 0 and info.file_size > 1024:
                ratio = info.file_size / float(info.compress_size)
                if ratio > active_limits.max_compression_ratio:
                    raise ConfigError(
                        f"[{FormatDiagnosticCode.UNSUPPORTED_OBJECT}] 部件压缩比异常 ({ratio:.1f})，疑似 ZIP 炸弹: {info.filename}"
                    )

        if total_uncompressed > active_limits.max_uncompressed_bytes:
            raise ConfigError(
                f"[{FormatDiagnosticCode.UNSUPPORTED_OBJECT}] 文档总解压大小 ({total_uncompressed / 1024 / 1024:.1f} MiB) 超过安全上限 ({active_limits.max_uncompressed_bytes / 1024 / 1024:.1f} MiB)"
            )

        # 3. 读取主题与样式
        theme_bytes = None
        if "word/theme/theme1.xml" in zf.namelist():
            theme_bytes = zf.read("word/theme/theme1.xml")
        theme_fonts = _parse_theme_fonts(theme_bytes)

        if "word/styles.xml" not in zf.namelist():
            raise ConfigError("DOCX 包中缺少必需的 word/styles.xml 部件")
        styles_bytes = zf.read("word/styles.xml")
        styles_map, doc_def_rPr, doc_def_pPr = _parse_styles_xml(styles_bytes, theme_fonts)

        evaluator = EffectiveStyleEvaluator(styles_map, theme_fonts, doc_def_rPr, doc_def_pPr)

        # 4. 遍历 word/document.xml 并构建节点索引与属性求值
        if "word/document.xml" not in zf.namelist():
            raise ConfigError("DOCX 包中缺少主文档部件 word/document.xml")
        doc_bytes = zf.read("word/document.xml")
        doc_root = lxml.etree.fromstring(doc_bytes)
        relationship_targets: Dict[str, str] = {}
        rels_name = "word/_rels/document.xml.rels"
        if rels_name in zf.namelist():
            rels_root = lxml.etree.fromstring(zf.read(rels_name))
            rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            for relationship in rels_root.findall(f"{{{rel_ns}}}Relationship"):
                rid = relationship.get("Id")
                target = relationship.get("Target")
                if not rid or not target or relationship.get("TargetMode") == "External":
                    continue
                relationship_targets[rid] = posixpath.normpath(
                    target.lstrip("/") if target.startswith("/")
                    else posixpath.join("word", target)
                )
        story_part_bytes = {
            name: zf.read(name)
            for name in zf.namelist()
            if (
                name.startswith("word/header") and name.endswith(".xml")
            ) or (
                name.startswith("word/footer") and name.endswith(".xml")
            )
        }

    body = doc_root.find(qn("w:body"))
    if body is None:
        raise ConfigError("word/document.xml 中未找到 w:body 根节点")

    blocks: List[BlockInspection] = []
    sections: List[SectionInspection] = []
    sec_idx = 0

    def inspect_paragraph(
        p_el: lxml.etree._Element,
        path_str: str,
        story: str = "body",
        part_uri: str = "word/document.xml",
    ) -> BlockInspection:
        pPr = p_el.find(qn("w:pPr"))
        pStyle = pPr.find(qn("w:pStyle")) if pPr is not None else None
        p_sid = pStyle.get(qn("w:val")) if pStyle is not None else None
        p_sname = styles_map[p_sid].name if p_sid and p_sid in styles_map else p_sid

        # 提取文字
        texts = []
        for t_el in p_el.findall(f".//{qn('w:t')}"):
            if t_el.text:
                texts.append(t_el.text)
        visible_text = "".join(texts).strip()

        # 识别受保护对象
        protected: List[str] = []
        if p_el.find(f".//{qn('w:fldChar')}") is not None or p_el.find(f".//{qn('w:fldSimple')}") is not None:
            protected.append("field_complex")
        # omml
        if p_el.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath") is not None:
            protected.append("omml_math")
        # footnote
        if p_el.find(f".//{qn('w:footnoteReference')}") is not None:
            protected.append("footnote_ref")
        # drawing / picture
        if p_el.find(f".//{qn('w:drawing')}") is not None or p_el.find(f".//{qn('w:pict')}") is not None:
            protected.append("image")
        # hyperlink
        if p_el.find(f".//{qn('w:hyperlink')}") is not None:
            protected.append("hyperlink")

        # outline level
        outline_lvl = None
        outline_source = None
        if pPr is not None:
            out_el = pPr.find(qn("w:outlineLvl"))
            if out_el is not None:
                try:
                    outline_lvl = int(out_el.get(qn("w:val"))) + 1
                    outline_source = "direct"
                except (ValueError, TypeError):
                    pass
        if outline_lvl is None and p_sid:
            for s_node in evaluator.get_style_chain(p_sid):
                if s_node.outline_level is not None:
                    outline_lvl = s_node.outline_level
                    outline_source = "style"

        # runs：保留每个 run 的有效值；首个 run 仅作为旧 API 的兼容代表值。
        runs = list(p_el.iter(qn("w:r")))
        effective_runs: List[RunStyle] = []
        effective_run_spans: List[Dict[str, Any]] = []
        inline_emphasis_spans: List[Dict[str, Any]] = []
        logical_offset = 0
        for r in runs:
            rPr = r.find(qn("w:rPr"))
            r_style_id = None
            if rPr is not None:
                rStyle_el = rPr.find(qn("w:rStyle"))
                if rStyle_el is not None:
                    r_style_id = rStyle_el.get(qn("w:val"))
            effective_run = evaluator.compute_effective_run(p_sid, r_style_id, rPr)
            effective_runs.append(effective_run)
            logical_text = _logical_run_text(r)
            if logical_text:
                effective_run_spans.append({
                    "start": logical_offset,
                    "end": logical_offset + len(logical_text),
                    "style": effective_run,
                })
                inline_emphasis_spans.append({
                    "start": logical_offset,
                    "end": logical_offset + len(logical_text),
                    "bold": _parse_toggle(rPr, "w:b"),
                    "italic": _parse_toggle(rPr, "w:i"),
                })
                logical_offset += len(logical_text)
        eff_para = evaluator.compute_effective_paragraph(p_sid, pPr)
        eff_run = effective_runs[0] if effective_runs else evaluator.compute_effective_run(p_sid, None, None)

        node_ref = NodeRef(
            source_sha256=source_sha256,
            part_uri=part_uri,
            element_path=path_str,
            text_hash=text_hash(_identity_text_without_field_cache(p_el)),
        )

        return BlockInspection(
            node=node_ref,
            structure_type="paragraph",
            story_type=story,
            visible_text=visible_text,
            p_style_id=p_sid,
            p_style_name=p_sname,
            outline_level=outline_lvl,
            effective_run=eff_run,
            effective_paragraph=eff_para,
            protected_objects=protected,
            runs_count=len(runs),
            effective_runs=tuple(effective_runs),
            effective_run_spans=tuple(effective_run_spans),
            inline_emphasis_spans=tuple(inline_emphasis_spans),
            outline_source=outline_source,
        )

    # 遍历 body 顶级子节点
    p_counter = 0
    tbl_counter = 0

    for child in body:
        tag = child.tag
        if tag == qn("w:p"):
            p_counter += 1
            p_path = f"/w:document/w:body/w:p[{p_counter}]"
            blk = inspect_paragraph(child, p_path, "body")
            blocks.append(blk)

            # 检查段内分节符 (w:p/w:pPr/w:sectPr)
            pPr = child.find(qn("w:pPr"))
            if pPr is not None:
                sectPr = pPr.find(qn("w:sectPr"))
                if sectPr is not None:
                    sections.append(_parse_section_spec(sectPr, sec_idx))
                    sec_idx += 1

        elif tag == qn("w:tbl"):
            tbl_counter += 1
            tbl_path = f"/w:document/w:body/w:tbl[{tbl_counter}]"

            # 遍历表格单元格中的段落
            r_counter = 0
            for row in child.findall(qn("w:tr")):
                r_counter += 1
                c_counter = 0
                for cell in row.findall(qn("w:tc")):
                    c_counter += 1
                    cell_p_counter = 0
                    for cell_p in cell.findall(qn("w:p")):
                        cell_p_counter += 1
                        cell_p_path = f"{tbl_path}/w:tr[{r_counter}]/w:tc[{c_counter}]/w:p[{cell_p_counter}]"
                        blk = inspect_paragraph(cell_p, cell_p_path, "body")
                        blocks.append(blk)

        elif tag == qn("w:sectPr"):
            # 文档末尾最终分节
            sections.append(_parse_section_spec(child, sec_idx))
            sec_idx += 1

    stories_by_key: Dict[Tuple[str, str, str], StoryInspection] = {}
    section_bindings: List[SectionStoryBinding] = []
    previous_effective: Dict[Tuple[str, str], str] = {}

    def register_story(
        story_type: str,
        variant: str,
        part_uri: str,
    ) -> None:
        key = (story_type, variant, part_uri)
        if key in stories_by_key:
            return
        story_bytes = story_part_bytes.get(part_uri)
        if story_bytes is None:
            diagnostics.append(Diagnostic(
                code=FormatDiagnosticCode.UNSUPPORTED_OBJECT,
                severity="error",
                location=part_uri,
                message="页眉/页脚关系目标部件不存在",
            ))
            return
        try:
            story_root = lxml.etree.fromstring(story_bytes)
        except lxml.etree.XMLSyntaxError as exc:
            diagnostics.append(Diagnostic(
                code=FormatDiagnosticCode.UNSUPPORTED_OBJECT,
                severity="error",
                location=part_uri,
                message=f"页眉/页脚部件 XML 无法解析: {exc}",
            ))
            return
        story_blocks: List[BlockInspection] = []
        story_p_counter = 0
        story_table_counter = 0

        def inspect_story_table(table_el: lxml.etree._Element, path_str: str) -> BlockInspection:
            """Inspect a header/footer table as one story block.

            Story tables must remain outside the body block sequence, but they
            still need stable identity and protected-object evidence so story
            contracts can compare them after package import/remapping.
            """
            paragraphs = list(table_el.iter(qn("w:p")))
            visible_text = " ".join(
                _identity_text_without_field_cache(paragraph)
                for paragraph in paragraphs
                if _identity_text_without_field_cache(paragraph)
            ).strip()
            protected: List[str] = []
            for tag_name, feature in (
                ("w:fldChar", "field_complex"),
                ("w:fldSimple", "field_complex"),
                ("m:oMath", "omml_math"),
                ("w:footnoteReference", "footnote_ref"),
                ("w:endnoteReference", "endnote_ref"),
                ("w:drawing", "image"),
                ("w:pict", "image"),
                ("w:hyperlink", "hyperlink"),
            ):
                if table_el.find(f".//{qn(tag_name)}") is not None and feature not in protected:
                    protected.append(feature)

            first = None
            if paragraphs:
                first = inspect_paragraph(
                    paragraphs[0],
                    f"{path_str}/w:tr[1]/w:tc[1]/w:p[1]",
                    story_type,
                    part_uri,
                )
            node = NodeRef(
                source_sha256=source_sha256,
                part_uri=part_uri,
                element_path=path_str,
                text_hash=text_hash(visible_text),
            )
            if first is None:
                return BlockInspection(
                    node=node,
                    structure_type="table",
                    story_type=story_type,
                    visible_text=visible_text,
                    p_style_id=None,
                    p_style_name=None,
                    outline_level=None,
                    effective_run=None,
                    effective_paragraph=None,
                    protected_objects=protected,
                    runs_count=len(list(table_el.iter(qn("w:r")))),
                )
            return replace(
                first,
                node=node,
                structure_type="table",
                visible_text=visible_text,
                protected_objects=protected,
                runs_count=len(list(table_el.iter(qn("w:r")))),
            )

        for story_child in story_root:
            if story_child.tag != qn("w:p"):
                if story_child.tag == qn("w:tbl"):
                    story_table_counter += 1
                    story_blocks.append(inspect_story_table(
                        story_child,
                        f"/{'w:hdr' if story_type == 'header' else 'w:ftr'}/w:tbl[{story_table_counter}]",
                    ))
                continue
            story_p_counter += 1
            story_blocks.append(inspect_paragraph(
                story_child,
                f"/{'w:hdr' if story_type == 'header' else 'w:ftr'}/w:p[{story_p_counter}]",
                story_type,
                part_uri,
            ))
        stories_by_key[key] = StoryInspection(
            story_type=story_type,
            variant=variant,
            part_uri=part_uri,
            blocks=tuple(story_blocks),
        )

    for section in sections:
        current_effective: Dict[Tuple[str, str], str] = {}
        for story_type, references in (
            ("header", section.header_references),
            ("footer", section.footer_references),
        ):
            explicit = {
                str(reference.get("variant") or "default"): reference
                for reference in references
            }
            variants = set(explicit) | {
                variant for (bound_story, variant) in previous_effective
                if bound_story == story_type
            }
            for variant in sorted(variants):
                reference = explicit.get(variant)
                linked_to_previous = reference is None
                if reference is not None:
                    rid = reference.get("relationship_id")
                    part_uri = relationship_targets.get(rid or "")
                    if not part_uri:
                        diagnostics.append(Diagnostic(
                            code=FormatDiagnosticCode.UNSUPPORTED_OBJECT,
                            severity="error",
                            location=f"section[{section.section_index}].{story_type}.{variant}",
                            message=f"无法解析 {story_type} 关系目标: {rid}",
                        ))
                        continue
                else:
                    part_uri = previous_effective[(story_type, variant)]
                current_effective[(story_type, variant)] = part_uri
                section_bindings.append(SectionStoryBinding(
                    section_index=section.section_index,
                    story_type=story_type,
                    variant=variant,
                    part_uri=part_uri,
                    linked_to_previous=linked_to_previous,
                ))
                register_story(story_type, variant, part_uri)
        previous_effective = current_effective

    effective_story_keys = {
        (binding.section_index, binding.story_type)
        for binding in section_bindings
    }
    sections = [
        replace(
            section,
            has_header=(section.section_index, "header") in effective_story_keys,
            has_footer=(section.section_index, "footer") in effective_story_keys,
        )
        for section in sections
    ]

    return DocumentInspection(
        source_sha256=source_sha256,
        file_path=str(file_path),
        file_size=file_size,
        blocks=blocks,
        styles=styles_map,
        theme_fonts=theme_fonts,
        sections=sections,
        diagnostics=diagnostics,
        stories=list(stories_by_key.values()),
        section_bindings=section_bindings,
    )
