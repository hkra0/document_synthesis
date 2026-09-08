# -*- coding: utf-8 -*-
"""
具名 Word 样式安装、属性所有权清理与排版应用器 (lib/style_applier.py)
安装 Synth 系列原生具名样式，严格保护行内语义（上下标、局部强调、公式、字段、超链接），
杜绝删除 run 节点再写回纯文本的破坏性重排。
"""

import copy
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Tuple, Set, Union

from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

from lib.config import ConfigError
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
    FormattingPolicy,
    ResolvedFormat,
)
from lib.styles import PPR_ORDER, RPR_ORDER, normalize_openxml_element_order


# -----------------------------------------------------------------------------
# 具名样式注册映射表 (Role -> (styleId, displayName))
# -----------------------------------------------------------------------------

ROLE_STYLE_ID_MAP: Dict[str, Tuple[str, str]] = {
    "body": ("SynthBody", "Synth Body"),
    "title": ("SynthTitle", "Synth Title"),
    "subtitle": ("SynthSubtitle", "Synth Subtitle"),
    "heading.1": ("SynthHeading1", "Synth Heading 1"),
    "heading.2": ("SynthHeading2", "Synth Heading 2"),
    "heading.3": ("SynthHeading3", "Synth Heading 3"),
    "heading.4": ("SynthHeading4", "Synth Heading 4"),
    "heading.5": ("SynthHeading5", "Synth Heading 5"),
    "heading.6": ("SynthHeading6", "Synth Heading 6"),
    "heading.7": ("SynthHeading7", "Synth Heading 7"),
    "heading.8": ("SynthHeading8", "Synth Heading 8"),
    "heading.9": ("SynthHeading9", "Synth Heading 9"),
    "toc.title": ("SynthTocTitle", "Synth TOC Title"),
    "toc.1": ("SynthToc1", "Synth TOC 1"),
    "toc.2": ("SynthToc2", "Synth TOC 2"),
    "toc.3": ("SynthToc3", "Synth TOC 3"),
    "toc.4": ("SynthToc4", "Synth TOC 4"),
    "toc.5": ("SynthToc5", "Synth TOC 5"),
    "toc.6": ("SynthToc6", "Synth TOC 6"),
    "toc.7": ("SynthToc7", "Synth TOC 7"),
    "toc.8": ("SynthToc8", "Synth TOC 8"),
    "toc.9": ("SynthToc9", "Synth TOC 9"),
    "caption": ("SynthCaption", "Synth Caption"),
    "table.body": ("SynthTableBody", "Synth Table Body"),
    "header": ("SynthHeader", "Synth Header"),
    "footer": ("SynthFooter", "Synth Footer"),
    "quote": ("SynthQuote", "Synth Quote"),
    "bibliography": ("SynthBibliography", "Synth Bibliography"),
}


@dataclass(frozen=True)
class ApplyReport:
    """排版应用执行报告"""
    styled_paragraphs_count: int
    cleared_direct_formats_count: int
    preserved_inline_objects_count: int
    role_distribution: Dict[str, int] = field(default_factory=dict)
    diagnostics: List[Diagnostic] = field(default_factory=list)


# -----------------------------------------------------------------------------
# 样式生成与安装 (install_styles)
# -----------------------------------------------------------------------------

def _build_style_element(
    style_id: str,
    display_name: str,
    style_def: StyleDefinition,
    outline_level: Optional[int] = None,
    based_on_id: Optional[str] = None,
    next_style_id: Optional[str] = None,
) -> OxmlElement:
    """构造完整的 w:style XML 元素"""
    style_el = parse_xml(f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="{style_id}"/>')

    # w:name
    name_el = parse_xml(f'<w:name {nsdecls("w")} w:val="{display_name}"/>')
    style_el.append(name_el)

    # w:basedOn
    if based_on_id:
        based_el = parse_xml(f'<w:basedOn {nsdecls("w")} w:val="{based_on_id}"/>')
        style_el.append(based_el)

    # w:next
    if next_style_id:
        next_el = parse_xml(f'<w:next {nsdecls("w")} w:val="{next_style_id}"/>')
        style_el.append(next_el)

    # w:pPr
    pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
    has_pPr = False

    # outline level (0-8 in OOXML)
    if outline_level is not None and 1 <= outline_level <= 9:
        out_el = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{outline_level - 1}"/>')
        pPr.append(out_el)
        has_pPr = True

    p_style = style_def.paragraph
    if p_style:
        # alignment (w:jc)
        if p_style.alignment:
            jc_val = "both" if p_style.alignment == "justify" else p_style.alignment
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="{jc_val}"/>'))
            has_pPr = True

        # spacing (w:spacing)
        sp_attrs = []
        if p_style.space_before_pt is not None:
            sp_attrs.append(f'w:before="{int(round(p_style.space_before_pt * 20))}"')
        if p_style.space_after_pt is not None:
            sp_attrs.append(f'w:after="{int(round(p_style.space_after_pt * 20))}"')
        if p_style.line_spacing:
            lsp = p_style.line_spacing
            if lsp.mode == "multiple" and lsp.value:
                sp_attrs.append(f'w:line="{int(round(lsp.value * 240))}" w:lineRule="auto"')
            elif lsp.mode == "exact" and lsp.value:
                sp_attrs.append(f'w:line="{int(round(lsp.value * 20))}" w:lineRule="exact"')
            elif lsp.mode == "at_least" and lsp.value:
                sp_attrs.append(f'w:line="{int(round(lsp.value * 20))}" w:lineRule="atLeast"')
            elif lsp.mode == "single":
                sp_attrs.append('w:line="240" w:lineRule="auto"')
        if sp_attrs:
            pPr.append(parse_xml(f'<w:spacing {nsdecls("w")} {" ".join(sp_attrs)}/>'))
            has_pPr = True

        # indents (w:ind)
        ind_attrs = []
        if p_style.first_line_indent:
            fli = p_style.first_line_indent
            if fli.unit == "char":
                ind_attrs.append(f'w:firstLineChars="{int(round(fli.value * 100))}"')
            else:
                ind_attrs.append(f'w:firstLine="{fli.to_twip()}"')
        elif p_style.hanging_indent:
            hgi = p_style.hanging_indent
            if hgi.unit == "char":
                ind_attrs.append(f'w:hangingChars="{int(round(hgi.value * 100))}"')
            else:
                ind_attrs.append(f'w:hanging="{hgi.to_twip()}"')
        if p_style.left_indent:
            ind_attrs.append(f'w:left="{p_style.left_indent.to_twip()}"')
        if p_style.right_indent:
            ind_attrs.append(f'w:right="{p_style.right_indent.to_twip()}"')
        if ind_attrs:
            pPr.append(parse_xml(f'<w:ind {nsdecls("w")} {" ".join(ind_attrs)}/>'))
            has_pPr = True

        # toggles
        if p_style.keep_with_next is True:
            pPr.append(parse_xml(f'<w:keepNext {nsdecls("w")}/>'))
            has_pPr = True
        if p_style.page_break_before is True:
            pPr.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
            has_pPr = True
        if p_style.widow_control is not None:
            w_val = "1" if p_style.widow_control else "0"
            pPr.append(parse_xml(f'<w:widowControl {nsdecls("w")} w:val="{w_val}"/>'))
            has_pPr = True
        if p_style.snap_to_grid is not None:
            s_val = "1" if p_style.snap_to_grid else "0"
            pPr.append(parse_xml(f'<w:snapToGrid {nsdecls("w")} w:val="{s_val}"/>'))
            has_pPr = True

    if has_pPr:
        normalize_openxml_element_order(pPr, PPR_ORDER)
        style_el.append(pPr)

    # w:rPr
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    has_rPr = False

    r_style = style_def.run
    if r_style:
        # rFonts
        fonts_attrs = []
        latin = getattr(r_style, "latin", None)
        if latin:
            fonts_attrs.append(f'w:ascii="{latin}" w:hAnsi="{latin}"')
        east_asia = getattr(r_style, "east_asia", None)
        if east_asia:
            fonts_attrs.append(f'w:eastAsia="{east_asia}"')
        cs = getattr(r_style, "complex_script", None)
        if cs:
            fonts_attrs.append(f'w:cs="{cs}"')
        if fonts_attrs:
            rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} {" ".join(fonts_attrs)}/>'))
            has_rPr = True

        # sz / szCs (in half-points)
        if r_style.size_pt is not None:
            half_pts = int(round(r_style.size_pt * 2))
            rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{half_pts}"/>'))
            rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{half_pts}"/>'))
            has_rPr = True

        # bold
        if r_style.bold is True:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
            rPr.append(parse_xml(f'<w:bCs {nsdecls("w")}/>'))
            has_rPr = True
        elif r_style.bold is False:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")} w:val="0"/>'))
            has_rPr = True

        # italic
        if r_style.italic is True:
            rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
            rPr.append(parse_xml(f'<w:iCs {nsdecls("w")}/>'))
            has_rPr = True
        elif r_style.italic is False:
            rPr.append(parse_xml(f'<w:i {nsdecls("w")} w:val="0"/>'))
            has_rPr = True

        # color
        if r_style.color and r_style.color != "auto":
            rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{r_style.color}"/>'))
            has_rPr = True

    if has_rPr:
        normalize_openxml_element_order(rPr, RPR_ORDER)
        style_el.append(rPr)

    return style_el


def install_styles(document: Document, resolved_format: ResolvedFormat) -> Dict[str, str]:
    """
    在目标文档的 styles.xml 中注入全套 Synth 具名原生样式。
    保证样式具备完整的 basedOn 继承、大纲层级与可视化名称。
    返回: role_name -> installed_style_id 的映射表
    """
    styles_el = document.part.styles._element

    # 索引现有样式 ID
    existing_map = {}
    for s in styles_el.findall(qn("w:style")):
        sid = s.get(qn("w:styleId"))
        if sid:
            existing_map[sid] = s

    role_to_style_id: Dict[str, str] = {}

    # 遍历所有角色进行样式生成
    for role_name, role_spec in resolved_format.roles.items():
        style_id, display_name = ROLE_STYLE_ID_MAP.get(role_name, (f"Synth_{role_name.replace('.', '_')}", f"Synth {role_name}"))
        role_to_style_id[role_name] = style_id

        # 查找目标样式规格
        style_def = resolved_format.styles.get(role_spec.style)
        if not style_def:
            continue

        # 处理 based_on
        based_on_id = None
        if style_def.based_on and style_def.based_on in resolved_format.roles:
            parent_role = style_def.based_on
            based_on_id = ROLE_STYLE_ID_MAP.get(parent_role, (f"Synth_{parent_role}", ""))[0]
        elif style_id != "SynthBody" and "Normal" in existing_map:
            based_on_id = "Normal"

        next_id = "SynthBody" if "heading" in role_name else None

        style_el = _build_style_element(
            style_id=style_id,
            display_name=display_name,
            style_def=style_def,
            outline_level=role_spec.outline_level,
            based_on_id=based_on_id,
            next_style_id=next_id,
        )

        # 替换已有同名样式
        if style_id in existing_map:
            styles_el.remove(existing_map[style_id])
        styles_el.append(style_el)
        existing_map[style_id] = style_el

    return role_to_style_id


# -----------------------------------------------------------------------------
# 角色应用与属性所有权清理 (apply_roles)
# -----------------------------------------------------------------------------

def apply_roles(
    document: Document,
    role_assignments: Dict[str, str],
    resolved_format: ResolvedFormat,
    policy: Optional[FormattingPolicy] = None,
) -> ApplyReport:
    """
    按角色分配应用目标样式，根据属性所有权安全清理直接格式。
    严格保护行内局部语义（上下标、局部强调、数学公式、复杂字段、超链接、图片）。
    """
    active_policy = policy or FormattingPolicy()

    # preserve 是明确的保真合同：不安装 Synth 样式、不写 pStyle，也不清理
    # 源段落/字符直接格式。映射仍由准备阶段读取，但不能在这里改变正文。
    if active_policy.mode == "preserve":
        return ApplyReport(
            styled_paragraphs_count=0,
            cleared_direct_formats_count=0,
            preserved_inline_objects_count=0,
            role_distribution={},
            diagnostics=[],
        )

    role_to_style_id = install_styles(document, resolved_format)

    styled_count = 0
    cleared_count = 0
    preserved_inline_count = 0
    role_dist: Dict[str, int] = {}
    diagnostics: List[Diagnostic] = []

    paragraph_entries = []
    for idx, paragraph in enumerate(document.paragraphs):
        paragraph_entries.append((
            f"/w:document/w:body/w:p[{idx + 1}]",
            paragraph,
            idx,
        ))
    # ``Document.paragraphs`` excludes table-cell paragraphs, while the
    # inspector/RoleMapper assigns those nodes the canonical table.body role.
    # Visit them explicitly so a restyle build cannot leave cells on Normal.
    for table_index, table in enumerate(document.tables, start=1):
        for row_index, row in enumerate(table.rows, start=1):
            for cell_index, cell in enumerate(row.cells, start=1):
                for paragraph_index, paragraph in enumerate(cell.paragraphs, start=1):
                    paragraph_entries.append((
                        f"/w:document/w:body/w:tbl[{table_index}]"
                        f"/w:tr[{row_index}]/w:tc[{cell_index}]/w:p[{paragraph_index}]",
                        paragraph,
                        None,
                    ))

    for p_path, p, legacy_idx in paragraph_entries:
        
        # 查找角色映射
        target_role = None
        if p_path in role_assignments:
            target_role = role_assignments[p_path]
        elif legacy_idx is not None and str(legacy_idx) in role_assignments:
            target_role = role_assignments[str(legacy_idx)]
        elif legacy_idx is not None and f"p[{legacy_idx}]" in role_assignments:
            target_role = role_assignments[f"p[{legacy_idx}]"]

        if not target_role:
            if active_policy.on_unmapped == "error":
                raise ConfigError(
                    f"[{FormatDiagnosticCode.ROLE_UNRESOLVED}] 段落未指定角色且 on_unmapped=error: {p_path} (文本: '{p.text[:30]}')"
                )
            # on_unmapped=preserve 的语义是整段保持不动，而不是偷偷降级为正文。
            # mixed 的 NodeRef scope 依赖这个分支保护未托管区域。
            continue

        role_dist[target_role] = role_dist.get(target_role, 0) + 1
        target_style_id = role_to_style_id.get(target_role, "SynthBody")
        role_spec = resolved_format.roles.get(target_role)
        style_def = resolved_format.styles.get(role_spec.style) if role_spec else None

        p_pr = p._p.get_or_add_pPr()

        # Numbering is semantic content, even when the source paragraph gets
        # its ``numPr`` through a paragraph style such as List Bullet.  Move
        # that inherited numbering to the paragraph before replacing pStyle;
        # otherwise restyle mode silently turns lists into ordinary prose.
        if p_pr.find(qn("w:numPr")) is None:
            source_style = p_pr.find(qn("w:pStyle"))
            source_style_id = source_style.get(qn("w:val")) if source_style is not None else None
            if source_style_id:
                source_style_el = document.part.styles._element.find(
                    f".//{qn('w:style')}[@{qn('w:styleId')}='{source_style_id}']"
                )
                inherited_num_pr = (
                    source_style_el.find(f"./{qn('w:pPr')}/{qn('w:numPr')}")
                    if source_style_el is not None
                    else None
                )
                if inherited_num_pr is not None:
                    p_pr.append(copy.deepcopy(inherited_num_pr))

        # 1. 设置 pStyle
        pStyle = p_pr.find(qn("w:pStyle"))
        if pStyle is None:
            pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{target_style_id}"/>')
            p_pr.insert(0, pStyle)
        else:
            pStyle.set(qn("w:val"), target_style_id)

        # 2. 设置大纲级别
        if role_spec and role_spec.outline_level is not None:
            out_lvl = p_pr.find(qn("w:outlineLvl"))
            if out_lvl is None:
                out_lvl = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{role_spec.outline_level - 1}"/>')
                p_pr.append(out_lvl)
            else:
                out_lvl.set(qn("w:val"), str(role_spec.outline_level - 1))
        styled_count += 1

        # 3. 属性所有权清理 (段落级)
        if style_def and style_def.paragraph:
            ps = style_def.paragraph
            if ps.alignment is not None:
                jc = p_pr.find(qn("w:jc"))
                if jc is not None:
                    p_pr.remove(jc)
                    cleared_count += 1
            if ps.space_before_pt is not None or ps.space_after_pt is not None or ps.line_spacing is not None:
                sp = p_pr.find(qn("w:spacing"))
                if sp is not None:
                    p_pr.remove(sp)
                    cleared_count += 1
            if ps.first_line_indent is not None or ps.hanging_indent is not None:
                ind = p_pr.find(qn("w:ind"))
                if ind is not None:
                    p_pr.remove(ind)
                    cleared_count += 1
            if ps.keep_with_next is not None:
                kn = p_pr.find(qn("w:keepNext"))
                if kn is not None:
                    p_pr.remove(kn)
                    cleared_count += 1
            if ps.page_break_before is not None:
                pbb = p_pr.find(qn("w:pageBreakBefore"))
                if pbb is not None:
                    p_pr.remove(pbb)
                    cleared_count += 1

        # 4. 属性所有权清理与行内语义保护 (字符/Run 级)
        runs = p.runs
        # 统计本段是否全段加粗 (全段加粗通常为旧直接格式；部分加粗为局部强调)
        bold_runs_count = sum(1 for r in runs if r.bold)
        is_all_bold = (len(runs) > 1 and bold_runs_count == len(runs)) or (len(runs) == 1 and bold_runs_count == 1 and "heading" in target_role)

        for r in runs:
            r_pr = r._r.find(qn("w:rPr"))
            if r_pr is None:
                continue

            # 统计受保护对象
            if r_pr.find(qn("w:vertAlign")) is not None:
                preserved_inline_count += 1
            if r._r.find(f".//{qn('w:fldChar')}") is not None or r._r.find(f".//{qn('w:fldSimple')}") is not None:
                preserved_inline_count += 1

            # 清理受目标样式托管的字体、字号、颜色
            if style_def and style_def.run:
                rs = style_def.run
                if rs.east_asia or rs.latin:
                    rf = r_pr.find(qn("w:rFonts"))
                    if rf is not None:
                        r_pr.remove(rf)
                        cleared_count += 1
                if rs.size_pt is not None:
                    sz = r_pr.find(qn("w:sz"))
                    szCs = r_pr.find(qn("w:szCs"))
                    if sz is not None:
                        r_pr.remove(sz)
                        cleared_count += 1
                    if szCs is not None:
                        r_pr.remove(szCs)
                        cleared_count += 1
                if rs.color:
                    col = r_pr.find(qn("w:color"))
                    if col is not None:
                        r_pr.remove(col)
                        cleared_count += 1

            # 加粗处理：如果目标是标题，由样式统一负责；如果是正文全段加粗，清理；局部加粗保留
            if "heading" in target_role or is_all_bold:
                b_el = r_pr.find(qn("w:b"))
                bCs_el = r_pr.find(qn("w:bCs"))
                if b_el is not None:
                    r_pr.remove(b_el)
                    cleared_count += 1
                if bCs_el is not None:
                    r_pr.remove(bCs_el)
            else:
                # 局部加粗作为行内语义保留
                if r.bold:
                    preserved_inline_count += 1

            # 斜体处理：局部斜体作为强调保留
            if r.italic and active_policy.inline_emphasis == "preserve":
                preserved_inline_count += 1

            normalize_openxml_element_order(r_pr, RPR_ORDER)

        normalize_openxml_element_order(p_pr, PPR_ORDER)

    return ApplyReport(
        styled_paragraphs_count=styled_count,
        cleared_direct_formats_count=cleared_count,
        preserved_inline_objects_count=preserved_inline_count,
        role_distribution=role_dist,
        diagnostics=diagnostics,
    )


# -----------------------------------------------------------------------------
# 格式上下文驱动的输出适配器 (Output Adapters)
# -----------------------------------------------------------------------------

def _apply_direct_run_style(run, style_def: Optional[StyleDefinition]) -> None:
    """给新生成的 run 写入目标样式的显式属性，确保跨文档导入后仍可见。"""
    if not style_def or not style_def.run:
        return
    rs = style_def.run
    if rs.latin:
        run.font.name = rs.latin
    if rs.size_pt is not None:
        run.font.size = Pt(rs.size_pt)
    if rs.bold is not None:
        run.font.bold = rs.bold
    if rs.italic is not None:
        run.font.italic = rs.italic
    if rs.color and rs.color != "auto":
        try:
            run.font.color.rgb = RGBColor.from_string(rs.color.replace("#", ""))
        except ValueError:
            pass
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None and (rs.latin or rs.east_asia or rs.complex_script):
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    if r_fonts is not None:
        if rs.latin:
            r_fonts.set(qn("w:ascii"), rs.latin)
            r_fonts.set(qn("w:hAnsi"), rs.latin)
        if rs.east_asia:
            r_fonts.set(qn("w:eastAsia"), rs.east_asia)
        if rs.complex_script:
            r_fonts.set(qn("w:cs"), rs.complex_script)
    normalize_openxml_element_order(r_pr, RPR_ORDER)


def _apply_direct_paragraph_style(paragraph: Paragraph, style_def: Optional[StyleDefinition]) -> None:
    """将生成段落的必要属性显式落地；完整定义仍保留在 Synth 样式中。"""
    if not style_def or not style_def.paragraph:
        return
    ps = style_def.paragraph
    pf = paragraph.paragraph_format
    if ps.alignment:
        alignment = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }.get(ps.alignment)
        if alignment is not None:
            pf.alignment = alignment
    if ps.space_before_pt is not None:
        pf.space_before = Pt(ps.space_before_pt)
    if ps.space_after_pt is not None:
        pf.space_after = Pt(ps.space_after_pt)
    if ps.line_spacing:
        if ps.line_spacing.mode == "multiple" and ps.line_spacing.value:
            pf.line_spacing = ps.line_spacing.value
        elif ps.line_spacing.value is not None:
            pf.line_spacing = Pt(ps.line_spacing.value)
        else:
            pf.line_spacing = 1.0
    if ps.keep_with_next is not None:
        pf.keep_with_next = ps.keep_with_next
    if ps.keep_lines is not None:
        pf.keep_together = ps.keep_lines
    if ps.page_break_before is not None:
        pf.page_break_before = ps.page_break_before
    if ps.first_line_indent and ps.first_line_indent.unit != "char":
        pf.first_line_indent = Pt(ps.first_line_indent.to_pt())
    if ps.left_indent and ps.left_indent.unit != "char":
        pf.left_indent = Pt(ps.left_indent.to_pt())
    if ps.right_indent and ps.right_indent.unit != "char":
        pf.right_indent = Pt(ps.right_indent.to_pt())


def _apply_direct_xml_run_style(run_element, style_def: Optional[StyleDefinition]) -> None:
    """为超链接/域等非 python-docx Run 对象写入同样的字符属性。"""
    if not style_def or not style_def.run:
        return
    rs = style_def.run
    r_pr = run_element.find(qn("w:rPr"))
    if r_pr is None:
        r_pr = OxmlElement("w:rPr")
        run_element.insert(0, r_pr)
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None and (rs.latin or rs.east_asia or rs.complex_script):
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    if r_fonts is not None:
        if rs.latin:
            r_fonts.set(qn("w:ascii"), rs.latin)
            r_fonts.set(qn("w:hAnsi"), rs.latin)
        if rs.east_asia:
            r_fonts.set(qn("w:eastAsia"), rs.east_asia)
        if rs.complex_script:
            r_fonts.set(qn("w:cs"), rs.complex_script)
    if rs.size_pt is not None:
        for name in ("sz", "szCs"):
            element = r_pr.find(qn(f"w:{name}"))
            if element is None:
                element = OxmlElement(f"w:{name}")
                r_pr.append(element)
            element.set(qn("w:val"), str(int(round(rs.size_pt * 2))))
    if rs.bold is not None:
        element = r_pr.find(qn("w:b"))
        if element is None:
            element = OxmlElement("w:b")
            r_pr.append(element)
        if not rs.bold:
            element.set(qn("w:val"), "0")
    if rs.italic is not None:
        element = r_pr.find(qn("w:i"))
        if element is None:
            element = OxmlElement("w:i")
            r_pr.append(element)
        if not rs.italic:
            element.set(qn("w:val"), "0")
    if rs.color and rs.color != "auto":
        element = r_pr.find(qn("w:color"))
        if element is None:
            element = OxmlElement("w:color")
            r_pr.append(element)
        element.set(qn("w:val"), rs.color.replace("#", ""))
    normalize_openxml_element_order(r_pr, RPR_ORDER)

def add_styled_heading(
    document: Document,
    text: str,
    level: int,
    resolved_format: ResolvedFormat,
    bookmark_name: Optional[str] = None,
    bookmark_id: Optional[int] = None,
    need_page_break: bool = False,
    role_name: Optional[str] = None,
) -> Paragraph:
    """插入具名标题段落，应用 SynthHeading 样式，不再硬编码中西文字体或内联字号"""
    role_to_style = install_styles(document, resolved_format)
    role_name = role_name or f"heading.{level}"
    target_style_id = role_to_style.get(role_name, f"SynthHeading{level}")

    p = document.add_paragraph()
    p_pr = p._p.get_or_add_pPr()

    # 设置 pStyle
    pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{target_style_id}"/>')
    p_pr.append(pStyle)

    # 只有真正的 heading 角色才写大纲级别；目录标题不能污染导航结构。
    if role_name.startswith("heading."):
        out_lvl = parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level - 1}"/>')
        p_pr.append(out_lvl)
    normalize_openxml_element_order(p_pr, PPR_ORDER)
    p.paragraph_format.page_break_before = need_page_break

    # 插入书签
    if bookmark_name:
        bm_id = bookmark_id or 100
        start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bookmark_name}"/>')
        end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
        p._p.append(start)
        p.add_run(text)
        p._p.append(end)
    else:
        p.add_run(text)

    _apply_direct_paragraph_style(p, resolved_format.styles.get(resolved_format.roles.get(role_name).style) if role_name in resolved_format.roles else None)
    for run in p.runs:
        _apply_direct_run_style(run, resolved_format.styles.get(resolved_format.roles.get(role_name).style) if role_name in resolved_format.roles else None)

    return p


def add_styled_toc_entry(
    document: Document,
    text: str,
    page_num: int,
    level: int,
    resolved_format: ResolvedFormat,
    bookmark_name: Optional[str] = None,
) -> Paragraph:
    """插入目录条目段落，应用 SynthToc 样式与点号制表符，支持书签超链接"""
    role_to_style = install_styles(document, resolved_format)
    target_style_id = role_to_style.get(f"toc.{level}", f"SynthToc{level}")

    p = document.add_paragraph()
    p_pr = p._p.get_or_add_pPr()

    pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{target_style_id}"/>')
    p_pr.append(pStyle)

    style_def = None
    role_spec = resolved_format.roles.get(f"toc.{level}")
    if role_spec:
        style_def = resolved_format.styles.get(role_spec.style)

    # 计算当前页面版心右侧制表位，不能假设 A4 或固定 8900 twip。
    from .layout import compute_content_box
    content_box = compute_content_box(resolved_format)
    tab_twip = content_box.content_width_twip
    leader = {"dots": "dot", "hyphens": "hyphen", "underline": "underscore", "none": "none"}.get(
        resolved_format.toc.leader, "dot"
    )
    tabs_el = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="{leader}" w:pos="{tab_twip}"/></w:tabs>')
    p_pr.append(tabs_el)
    normalize_openxml_element_order(p_pr, PPR_ORDER)

    # 如果有书签跳转则封装超链接
    if bookmark_name:
        hlink = parse_xml(f'<w:hyperlink {nsdecls("w")} w:anchor="{bookmark_name}"/>')
        r1 = parse_xml(f'<w:r {nsdecls("w")}><w:t>{text}</w:t></w:r>')
        r2 = parse_xml(f'<w:r {nsdecls("w")}><w:tab/><w:t>{page_num}</w:t></w:r>')
        _apply_direct_xml_run_style(r1, style_def)
        _apply_direct_xml_run_style(r2, style_def)
        hlink.append(r1)
        hlink.append(r2)
        p._p.append(hlink)
    else:
        r1 = p.add_run(text)
        r2 = p.add_run(f"\t{page_num}")
        _apply_direct_run_style(r1, style_def)
        _apply_direct_run_style(r2, style_def)

    _apply_direct_paragraph_style(p, style_def)

    return p


def setup_styled_footer(
    section,
    resolved_format: ResolvedFormat,
    start_page: Optional[int] = None,
):
    """基于 resolved_format.footer 配置分节页脚"""
    footer_spec = resolved_format.footer
    if footer_spec.mode == "source":
        # 源页脚由调用方在 preserve/source 路径复制；目标适配器不能覆盖它。
        return
    if footer_spec.mode == "none":
        section.footer.is_linked_to_previous = False
        section.footer.paragraphs[0].text = ""
        return

    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.text = ""
    p_pr = p._p.get_or_add_pPr()

    for child in list(p_pr.findall(qn("w:jc"))):
        p_pr.remove(child)

    role_spec = resolved_format.roles.get("footer")
    footer_style = resolved_format.styles.get(role_spec.style) if role_spec else resolved_format.styles.get("footer")
    if footer_style:
        p_style = p_pr.find(qn("w:pStyle"))
        if p_style is None:
            p_style = OxmlElement("w:pStyle")
            p_pr.insert(0, p_style)
        p_style.set(qn("w:val"), ROLE_STYLE_ID_MAP.get("footer", ("SynthFooter", ""))[0])
        _apply_direct_paragraph_style(p, footer_style)

    # 对齐
    jc_val = footer_spec.alignment
    jc = p_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        p_pr.append(jc)
    jc.set(qn("w:val"), jc_val)

    # 页码格式
    if footer_spec.format == "dash_number":
        run = p.add_run("- ")
        _apply_direct_run_style(run, footer_style)
        fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>')
        for run_element in fld.iter(qn("w:r")):
            _apply_direct_xml_run_style(run_element, footer_style)
        p._p.append(fld)
        run = p.add_run(" -")
        _apply_direct_run_style(run, footer_style)
    elif footer_spec.format == "number":
        fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>')
        for run_element in fld.iter(qn("w:r")):
            _apply_direct_xml_run_style(run_element, footer_style)
        p._p.append(fld)

    normalize_openxml_element_order(p_pr, PPR_ORDER)

    # 连续页码或起始页码
    secPr = section._sectPr
    if start_page is not None:
        pgNumType = secPr.find(qn("w:pgNumType"))
        if pgNumType is None:
            pgNumType = parse_xml(f'<w:pgNumType {nsdecls("w")}/>')
            secPr.append(pgNumType)
        pgNumType.set(qn("w:start"), str(start_page))
    else:
        pgNumType = secPr.find(qn("w:pgNumType"))
        if pgNumType is not None and qn("w:start") in pgNumType.attrib:
            del pgNumType.attrib[qn("w:start")]
