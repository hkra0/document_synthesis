# -*- coding: utf-8 -*-
"""
自适应版心几何度量与渲染上下文 (lib/layout.py)
提供 ContentBox、compute_content_box、RenderContext 以及图片/表格自适应版心约束。
消除对全局可变状态与硬编码纸张尺寸的依赖。
"""

import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Tuple, Union

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm, Mm
from docx.enum.section import WD_ORIENT

from .format_schema import (
    LengthValue,
    PageSpec,
    ResolvedFormat,
    FormattingPolicy,
    MM_TO_TWIP,
    PT_TO_TWIP,
)


@dataclass(frozen=True)
class ContentBox:
    """版心几何度量数据结构 (所有内部度量均以 twip 为基准单位)"""
    page_width_twip: int
    page_height_twip: int
    margin_top_twip: int
    margin_bottom_twip: int
    margin_left_twip: int
    margin_right_twip: int

    @property
    def content_width_twip(self) -> int:
        """可用内容宽度 (twip)"""
        return max(0, self.page_width_twip - self.margin_left_twip - self.margin_right_twip)

    @property
    def content_height_twip(self) -> int:
        """可用内容高度 (twip)"""
        return max(0, self.page_height_twip - self.margin_top_twip - self.margin_bottom_twip)

    @property
    def content_width_pt(self) -> float:
        return self.content_width_twip / PT_TO_TWIP

    @property
    def content_height_pt(self) -> float:
        return self.content_height_twip / PT_TO_TWIP

    @property
    def content_width_mm(self) -> float:
        return self.content_width_twip / MM_TO_TWIP

    @property
    def content_height_mm(self) -> float:
        return self.content_height_twip / MM_TO_TWIP

    @property
    def content_width_cm(self) -> float:
        return self.content_width_mm / 10.0

    @property
    def content_height_cm(self) -> float:
        return self.content_height_mm / 10.0

    @property
    def orientation(self) -> str:
        return "landscape" if self.page_width_twip > self.page_height_twip else "portrait"


# 默认 A4 纸张与公文默认边距 (twip)
# A4: 210mm x 297mm ≈ 11906 x 16838 twip
DEFAULT_A4_WIDTH_TWIP = 11906
DEFAULT_A4_HEIGHT_TWIP = 16838
DEFAULT_MARGIN_TOP_TWIP = int(round(37.0 * MM_TO_TWIP))     # 37mm ≈ 2098 twip
DEFAULT_MARGIN_BOTTOM_TWIP = int(round(35.0 * MM_TO_TWIP))  # 35mm ≈ 1984 twip
DEFAULT_MARGIN_LEFT_TWIP = int(round(28.0 * MM_TO_TWIP))    # 28mm ≈ 1587 twip
DEFAULT_MARGIN_RIGHT_TWIP = int(round(26.0 * MM_TO_TWIP))   # 26mm ≈ 1474 twip


def compute_content_box(spec_or_section: Any = None) -> ContentBox:
    """计算当前节或规格的版心几何度量。
    
    支持:
    - docx.section.Section: 从实际 Word 节对象提取当前页宽、页高与页边距
    - PageSpec: 从 ResolvedFormat.page 提取规格
    - ResolvedFormat: 提取其 page 规格
    - None: 返回默认 A4 纸张与边距
    """
    if spec_or_section is None:
        return ContentBox(
            page_width_twip=DEFAULT_A4_WIDTH_TWIP,
            page_height_twip=DEFAULT_A4_HEIGHT_TWIP,
            margin_top_twip=DEFAULT_MARGIN_TOP_TWIP,
            margin_bottom_twip=DEFAULT_MARGIN_BOTTOM_TWIP,
            margin_left_twip=DEFAULT_MARGIN_LEFT_TWIP,
            margin_right_twip=DEFAULT_MARGIN_RIGHT_TWIP,
        )

    # 1. 传入 ResolvedFormat
    if isinstance(spec_or_section, ResolvedFormat):
        return compute_content_box(spec_or_section.page)

    # 2. 传入 PageSpec
    if isinstance(spec_or_section, PageSpec):
        w = int(round(spec_or_section.width_mm * MM_TO_TWIP))
        h = int(round(spec_or_section.height_mm * MM_TO_TWIP))
        top = int(round(spec_or_section.margin_top_mm * MM_TO_TWIP))
        bottom = int(round(spec_or_section.margin_bottom_mm * MM_TO_TWIP))
        left = int(round(spec_or_section.margin_left_mm * MM_TO_TWIP))
        right = int(round(spec_or_section.margin_right_mm * MM_TO_TWIP))
        if spec_or_section.orientation == "landscape" and w < h:
            w, h = h, w
        return ContentBox(
            page_width_twip=w,
            page_height_twip=h,
            margin_top_twip=top,
            margin_bottom_twip=bottom,
            margin_left_twip=left,
            margin_right_twip=right,
        )

    # 3. 传入 RenderContext
    if isinstance(spec_or_section, RenderContext):
        return spec_or_section.get_content_box()

    # 4. 传入 docx.section.Section
    # Section 对象具有 page_width, page_height, top_margin 等属性（Length 对象）
    if hasattr(spec_or_section, "page_width") and hasattr(spec_or_section, "top_margin"):
        sec = spec_or_section
        def _to_twip(val: Any, default: int) -> int:
            if val is None:
                return default
            if hasattr(val, "pt"):
                return int(round(val.pt * PT_TO_TWIP))
            if hasattr(val, "twips"):
                return int(round(val.twips))
            return default

        pw = _to_twip(sec.page_width, DEFAULT_A4_WIDTH_TWIP)
        ph = _to_twip(sec.page_height, DEFAULT_A4_HEIGHT_TWIP)
        top = _to_twip(sec.top_margin, DEFAULT_MARGIN_TOP_TWIP)
        bot = _to_twip(sec.bottom_margin, DEFAULT_MARGIN_BOTTOM_TWIP)
        left = _to_twip(sec.left_margin, DEFAULT_MARGIN_LEFT_TWIP)
        right = _to_twip(sec.right_margin, DEFAULT_MARGIN_RIGHT_TWIP)
        return ContentBox(
            page_width_twip=pw,
            page_height_twip=ph,
            margin_top_twip=top,
            margin_bottom_twip=bot,
            margin_left_twip=left,
            margin_right_twip=right,
        )

    # 回退为默认 A4
    return compute_content_box(None)


def constrain_image_box(
    image_width_px: int,
    image_height_px: int,
    content_box: ContentBox,
    max_height_ratio: float = 1.0
) -> Tuple[float, float]:
    """等比约束图片尺寸，保证其在 content_box 的可用宽度和最大可用高度内。
    
    返回:
        (w_cm, h_cm): 以厘米为单位的尺寸元组，保持原纵横比。
    """
    max_w_cm = content_box.content_width_cm
    max_h_cm = content_box.content_height_cm * max_height_ratio

    if image_width_px <= 0 or image_height_px <= 0:
        return (round(max_w_cm, 3), round(max_h_cm, 3))

    aspect = float(image_height_px) / float(image_width_px)
    
    # 首先尝试以最大宽度排布
    candidate_w = max_w_cm
    candidate_h = candidate_w * aspect

    # 若高度超过最大高度限制，则以最大高度约束宽度
    if candidate_h > max_h_cm:
        candidate_h = max_h_cm
        candidate_w = candidate_h / aspect

    return (round(candidate_w, 3), round(candidate_h, 3))


def adapt_table_to_content_box(table: Any, content_box: ContentBox) -> bool:
    """自适应表格宽度：当表格总宽度超出 content_box.content_width_twip 时，
    按比例缩放各列宽度和单元格宽度，防止表格冲出页边距。
    
    返回:
        bool: 是否对表格进行了缩放修改。
    """
    tbl_elem = getattr(table, "_tbl", table)
    tblGrid = tbl_elem.find(qn("w:tblGrid"))
    if tblGrid is None:
        return False

    gridCols = tblGrid.findall(qn("w:gridCol"))
    if not gridCols:
        return False

    current_col_widths = []
    for col in gridCols:
        val = col.get(qn("w:w"))
        try:
            w_int = int(val) if val else 0
        except ValueError:
            w_int = 0
        current_col_widths.append(w_int)

    total_table_width = sum(current_col_widths)
    avail_width = content_box.content_width_twip

    if total_table_width <= avail_width or total_table_width <= 0:
        return False

    # 按比例缩放列宽
    scale = avail_width / float(total_table_width)
    new_col_widths = []
    accum = 0
    for idx, orig_w in enumerate(current_col_widths):
        if idx == len(current_col_widths) - 1:
            # 最后一列吸收取整误差
            scaled_w = max(1, avail_width - accum)
        else:
            scaled_w = max(1, int(round(orig_w * scale)))
            accum += scaled_w
        new_col_widths.append(scaled_w)

    # 1. 更新 tblGrid
    for col, new_w in zip(gridCols, new_col_widths):
        col.set(qn("w:w"), str(new_w))

    # 2. 更新 tblPr/w:tblW
    tblPr = tbl_elem.find(qn("w:tblPr"))
    if tblPr is not None:
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is not None:
            tblW.set(qn("w:w"), str(avail_width))
            tblW.set(qn("w:type"), "dxa")

    # 3. 更新每一行的单元格宽度
    for tr in tbl_elem.findall(qn("w:tr")):
        tc_elements = tr.findall(qn("w:tc"))
        for col_idx, tc in enumerate(tc_elements):
            if col_idx < len(new_col_widths):
                tcPr = tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    tcW = tcPr.find(qn("w:tcW"))
                    if tcW is not None:
                        tcW.set(qn("w:w"), str(new_col_widths[col_idx]))
                        tcW.set(qn("w:type"), "dxa")

    return True


def isolate_section_boundaries(section: Any) -> None:
    """断开节的页眉页脚与上一节的关联（is_linked_to_previous = False），
    确保封面、目录、主文或保留区跨节时不发生页眉页脚泄漏。
    """
    for attr in ("header", "footer", "first_page_header", "first_page_footer", "even_page_header", "even_page_footer"):
        if hasattr(section, attr):
            try:
                hf = getattr(section, attr)
                if hf is not None:
                    hf.is_linked_to_previous = False
            except Exception:
                pass


@dataclass
class RenderContext:
    """统一渲染上下文
    
    维护排版生命周期中的状态、度量与配置，替代原有的全局变量传递。
    """
    resolved_format: Optional[ResolvedFormat] = None
    formatting_policy: Optional[FormattingPolicy] = None
    current_section: Optional[Any] = None
    current_section_index: int = 0
    last_rendered_landscape: bool = False
    exact_pages: Dict[str, int] = field(default_factory=dict)
    resolved_files: Mapping[str, Path] = field(default_factory=dict)
    source_dir: Optional[Path] = None
    fonts: Mapping[str, str] = field(default_factory=dict)
    page_spec: Optional[PageSpec] = None

    @classmethod
    def from_config(
        cls,
        config: Any,
        *,
        exact_pages: Optional[Mapping[str, int]] = None,
        resolved_files: Optional[Mapping[str, Path]] = None,
        source_dir: Optional[Path] = None,
    ) -> "RenderContext":
        """从已解析配置建立一次构建生命周期内共享的上下文。"""
        resolved = getattr(config, "resolved_format", None)
        return cls(
            resolved_format=resolved,
            formatting_policy=getattr(config, "formatting", None),
            exact_pages=dict(exact_pages or {}),
            resolved_files=dict(resolved_files or {}),
            source_dir=source_dir,
            fonts=dict(getattr(config, "fonts", {}) or {}),
            page_spec=getattr(resolved, "page", None),
        )

    def get_content_box(self, section: Optional[Any] = None) -> ContentBox:
        """获取当前节或指定节的版心几何度量"""
        target = section or self.current_section
        if target is not None:
            return compute_content_box(target)
        if self.resolved_format and self.resolved_format.page:
            return compute_content_box(self.resolved_format.page)
        return compute_content_box(None)

    def track_landscape(self, is_landscape: bool) -> None:
        """记录当前节点排版是否以横版分节结尾"""
        self.last_rendered_landscape = is_landscape

    def consume_landscape(self) -> bool:
        """消费并重置横版状态（用于决定后继节点是否需要强制新页换行）"""
        val = self.last_rendered_landscape
        self.last_rendered_landscape = False
        return val


def apply_section_spec(
    section: Any,
    spec: Optional[Union[PageSpec, ResolvedFormat]] = None,
    *,
    source_section: Optional[Any] = None,
) -> Any:
    """把页面规格真正落到 Word 节对象上。

    ``source_section`` 用于 preserve/source 路径，逐项复制源节几何；否则从
    ``PageSpec`` 或 ``ResolvedFormat.page`` 应用目标纸张、边距、页眉页脚距离
    与网格策略。该函数是所有新执行路径的唯一页面几何入口。
    """
    if source_section is not None:
        for attr in (
            "page_width", "page_height", "top_margin", "bottom_margin",
            "left_margin", "right_margin", "header_distance", "footer_distance",
        ):
            value = getattr(source_section, attr, None)
            if value is not None:
                setattr(section, attr, value)
        try:
            section.orientation = source_section.orientation
        except Exception:
            pass
        source_grid = source_section._sectPr.find(qn("w:docGrid"))
        if source_grid is not None:
            sect_grid = section._sectPr.find(qn("w:docGrid"))
            if sect_grid is not None:
                section._sectPr.remove(sect_grid)
            section._sectPr.append(copy.deepcopy(source_grid))
        return section

    page = spec.page if isinstance(spec, ResolvedFormat) else spec
    if page is None:
        return section

    width_mm, height_mm = page.width_mm, page.height_mm
    if page.orientation == "landscape":
        width_mm, height_mm = max(width_mm, height_mm), min(width_mm, height_mm)
        section.orientation = WD_ORIENT.LANDSCAPE
    else:
        width_mm, height_mm = min(width_mm, height_mm), max(width_mm, height_mm)
        section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Mm(width_mm)
    section.page_height = Mm(height_mm)
    section.top_margin = Mm(page.margin_top_mm)
    section.bottom_margin = Mm(page.margin_bottom_mm)
    section.left_margin = Mm(page.margin_left_mm)
    section.right_margin = Mm(page.margin_right_mm)
    section.header_distance = Mm(page.header_distance_mm)
    section.footer_distance = Mm(page.footer_distance_mm)

    # docGrid 的存在与否是页面规格的一部分；具体行距仍由样式托管。
    sect_pr = section._sectPr
    grid = sect_pr.find(qn("w:docGrid"))
    if page.snap_to_grid:
        if grid is None:
            grid = OxmlElement("w:docGrid")
            sect_pr.append(grid)
        grid.set(qn("w:type"), "lines")
    elif grid is not None:
        sect_pr.remove(grid)
    return section
