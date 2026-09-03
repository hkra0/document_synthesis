# -*- coding: utf-8 -*-
"""
通用材料排版与两趟编译合成引擎 (lib/engine.py)
整合文件名大纲自发现、JSON 清单复写合并、多源材料渲染与 Word 原生页码反查。
彻底消除硬编码，提供端到端的一键合成能力。
"""

import os
import sys
import copy
import io
import re
import subprocess
import uuid
from pathlib import Path
from typing import List, Dict, Any, Optional, Mapping

import pymupdf
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from .config import ConfigError, ProjectConfig, load_project_config
from .scanner import flatten_tree_nodes
from .source_strategies import SourceStrategyError, build_outline
from .styles import (
    DEFAULT_FONTS, DEFAULT_PAGE_SETUP,
    apply_standard_page_setup, setup_footer, set_run_fonts,
    add_toc_entry_with_hyperlink, add_heading_paragraph,
    neutralize_document_styles
)
from .renderers import (
    render_docx_file, render_pdf_file, get_cached_pdf_page_image,
    copy_element_with_rels, append_element_to_body, LAST_RENDERED_LANDSCAPE,
    is_standalone_cover_doc, add_invisible_heading_anchor, set_pdf_render_cache_dir
)
from .qa import OfficeExportError, export_docx_to_pdf, get_exact_printed_heading_pages, run_qa_assertions


class BuildError(RuntimeError):
    """构建无法安全完成时抛出的错误。"""


def _applescript_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _run_applescript(script: str, label: str):
    if sys.platform != "darwin" or not os.path.exists("/usr/bin/osascript"):
        raise BuildError(f"{label} 需要 macOS 上的 Microsoft Office。")
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise BuildError(f"{label}失败: {detail or '未知错误'}")


def prepare_conversions(source_dir: Path, work_dir: Path) -> Dict[str, Path]:
    """把必要转换写入本次工作目录，绝不修改输入材料目录。"""
    converted_root = work_dir / "converted"
    resolved: Dict[str, Path] = {}

    for doc_path in source_dir.rglob("*.doc"):
        if doc_path.name.startswith("~$") or doc_path.with_suffix(".docx").exists():
            continue
        relative = doc_path.relative_to(source_dir)
        output_path = (converted_root / relative).with_suffix(".docx")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        print(f"  [转换] {relative} → 工作目录中的 DOCX")
        script = f'''
        with timeout of 600 seconds
            tell application "Microsoft Word"
                set display alerts to none
                open (POSIX file "{_applescript_string(str(doc_path.resolve()))}") confirm conversions false
                set sourceDoc to active document
                save as active document file name "{_applescript_string(str(output_path.resolve()))}" file format format document default
                close sourceDoc saving no
            end tell
        end timeout
        '''
        _run_applescript(script, f"转换 DOC 文件 {relative}")
        if not output_path.exists():
            raise BuildError(f"转换 DOC 文件后未找到产物: {output_path}")
        resolved[relative.as_posix()] = output_path

    for pptx_path in source_dir.rglob("*.pptx"):
        if pptx_path.name.startswith("~$"):
            continue
        relative = pptx_path.relative_to(source_dir)
        existing_pdf = pptx_path.with_suffix(".pdf")
        if existing_pdf.exists():
            resolved[relative.as_posix()] = existing_pdf
            continue
        output_path = (converted_root / relative).with_suffix(".pdf")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        print(f"  [转换] {relative} → 工作目录中的 PDF")
        script = f'''
        with timeout of 600 seconds
            tell application "Microsoft PowerPoint"
                open (POSIX file "{_applescript_string(str(pptx_path.resolve()))}")
                set sourcePresentation to active presentation
                save active presentation in (POSIX file "{_applescript_string(str(output_path.resolve()))}") as save as PDF
                close sourcePresentation saving no
            end tell
        end timeout
        '''
        _run_applescript(script, f"转换 PPTX 文件 {relative}")
        if not output_path.exists():
            raise BuildError(f"转换 PPTX 文件后未找到产物: {output_path}")
        resolved[relative.as_posix()] = output_path

    return resolved


def _resolve_node_path(node: Dict[str, Any], source_dir: Path, resolved_files: Mapping[str, Path]) -> Optional[Path]:
    rel_file = node.get("file")
    if not rel_file:
        return None
    return resolved_files.get(rel_file, source_dir / rel_file)


def render_tree_node_recursive(
    doc: Document,
    node: Dict[str, Any],
    source_dir: Path,
    fonts: Dict[str, str],
    is_first_section: bool = False,
    is_first_child: bool = False,
    exact_pages: Dict[str, int] = None,
    resolved_files: Optional[Mapping[str, Path]] = None
):
    """递归排版大纲树节点"""
    ntype = node.get("type", "folder")
    resolved_files = resolved_files or {}
    
    # 确定是否需要强制换页（首大节、首子节点、以及刚从横版恢复的分节无需换页）
    need_page_break = False
    if not is_first_section and not is_first_child and not LAST_RENDERED_LANDSCAPE[0]:
        need_page_break = True
    LAST_RENDERED_LANDSCAPE[0] = False

    if ntype == "folder":
        add_heading_paragraph(doc, node, fonts, need_page_break=need_page_break)
        children = node.get("children", [])
        for idx, child in enumerate(children):
            render_tree_node_recursive(
                doc, child, source_dir, fonts,
                is_first_child=(idx == 0),
                exact_pages=exact_pages,
                resolved_files=resolved_files
            )
            
    elif ntype == "docx_outline":
        rel_file = node.get("file")
        fpath = _resolve_node_path(node, source_dir, resolved_files)
        if node.get("add_heading_before_content", True):
            add_heading_paragraph(doc, node, fonts, need_page_break=need_page_break)
        elif need_page_break:
            # 根节点不显示标题时，仍让其正文在正确的新页开始。
            anchor = doc.add_paragraph()
            anchor.paragraph_format.page_break_before = True
        if fpath and fpath.exists():
            render_docx_file(
                doc, fpath, exact_pages=exact_pages, fonts=fonts,
                embedded_outline=node.get("children", []),
            )
        elif rel_file:
            raise BuildError(f"找不到 Word 文件: {fpath}")

    elif ntype == "docx":
        rel_file = node.get("file")
        fpath = _resolve_node_path(node, source_dir, resolved_files)
                
        has_cover = bool(fpath and fpath.exists() and is_standalone_cover_doc(fpath))
        if has_cover:
            add_invisible_heading_anchor(doc, node, need_page_break=need_page_break)
        else:
            add_heading_paragraph(doc, node, fonts, need_page_break=need_page_break)
            
        if fpath and fpath.exists():
            render_docx_file(doc, fpath, exact_pages=exact_pages, fonts=fonts)
        elif rel_file:
            raise BuildError(f"找不到 Word 文件: {fpath}")
                
    elif ntype == "pdf":
        add_heading_paragraph(doc, node, fonts, need_page_break=need_page_break)
        rel_file = node.get("file")
        if rel_file:
            fpath = _resolve_node_path(node, source_dir, resolved_files)
            if fpath and fpath.exists():
                render_pdf_file(doc, fpath, has_headings_on_page=True, dpi=node.get("pdf_dpi", 300))
            else:
                raise BuildError(f"找不到 PDF 文件: {fpath}")
                
    elif ntype == "pptx":
        rel_file = node.get("file")
        if rel_file:
            pptx_path = source_dir / rel_file
            pdf_path = resolved_files.get(rel_file, pptx_path.with_suffix(".pdf"))
            if pdf_path.exists():
                sec_pptx = doc.add_section(WD_SECTION_START.NEW_PAGE)
                sec_pptx.orientation = WD_ORIENT.LANDSCAPE
                sec_pptx.page_width = Cm(29.7)
                sec_pptx.page_height = Cm(21.0)
                sec_pptx.top_margin = Cm(2.2)
                sec_pptx.bottom_margin = Cm(2.2)
                sec_pptx.left_margin = Cm(2.8)
                sec_pptx.right_margin = Cm(2.6)
                pgSz = sec_pptx._sectPr.find(qn('w:pgSz'))
                if pgSz is not None:
                    pgSz.set(qn('w:w'), '16838')
                    pgSz.set(qn('w:h'), '11906')
                    pgSz.set(qn('w:orient'), 'landscape')
                setup_footer(sec_pptx, start_page=None)
                
                add_heading_paragraph(doc, node, fonts, need_page_break=False)
                
                pdf_doc = pymupdf.open(str(pdf_path))
                for i in range(len(pdf_doc)):
                    img_bytes = get_cached_pdf_page_image(pdf_path, i, dpi=300)
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pf = p_img.paragraph_format
                    if i > 0:
                        pf.page_break_before = True
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(0)
                    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
                    pf.line_spacing = None
                    w_cm = 23.2 if i == 0 else 24.0
                    r = p_img.add_run()
                    r.add_picture(io.BytesIO(img_bytes), width=Cm(w_cm))
                LAST_RENDERED_LANDSCAPE[0] = True
            else:
                raise BuildError(f"找不到 PPTX 对应的 PDF: {pdf_path}")
                
    elif ntype == "image":
        add_heading_paragraph(doc, node, fonts, need_page_break=need_page_break)
        rel_file = node.get("file")
        if rel_file:
            img_path = _resolve_node_path(node, source_dir, resolved_files)
            if img_path and img_path.exists():
                from PIL import Image as PILImage
                try:
                    with PILImage.open(img_path) as pimg:
                        pw, ph = pimg.size
                        aspect = pw / max(1, ph)
                except Exception:
                    aspect = 1.0
                
                # 动态计算安全可用高度（17.5cm），按比例缩放宽度，防止高图被挤到下一页产生单标题孤立
                max_avail_h_cm = 17.5
                default_w_cm = 14.8
                w_cm = min(default_w_cm, max_avail_h_cm * aspect)
                
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(4)
                p_img.paragraph_format.space_after = Pt(4)
                p_img.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p_img.paragraph_format.line_spacing = None
                r = p_img.add_run()
                r.add_picture(str(img_path), width=Cm(w_cm))
            else:
                raise BuildError(f"找不到图片文件: {img_path}")


def build_cover_and_toc_doc(
    config: ProjectConfig,
    flat_nodes: List[Dict[str, Any]],
    exact_pages: Dict[str, int],
    out_cover_path: Path
):
    """根据真实打印物理页码，结合模板生成公文红头封面与超链接目录文档"""
    fonts = config.fonts
    page_setup = config.page_setup
    
    tmpl_path = config.get_template_path()

    main_title = config.cover.get("main_title") or config.project_name
    header_title = config.cover.get("header_title")
    sub_title = config.cover.get("sub_title")
    
    if tmpl_path:
        doc = Document(str(tmpl_path))
        
        # 依据大标题字数自适应计算字号 (单位: half-points)
        n_chars = len(main_title)
        if n_chars <= 6:
            main_sz = "76"
        elif n_chars <= 8:
            main_sz = "72"
        elif n_chars <= 10:
            main_sz = "66"
        else:
            main_sz = "56"

        placeholders = {
            "{{HEADER_TITLE}}": header_title or "",
            "{{SUB_TITLE}}": sub_title or "",
            "{{MAIN_TITLE}}": main_title,
        }
        for p in doc.paragraphs:
            txt = p.text
            if "{{MAIN_TITLE}}" in txt:
                for r_el in list(p._element.findall(qn("w:r"))):
                    p._element.remove(r_el)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pf = p.paragraph_format
                pf.space_before = Pt(14)
                pf.space_after = Pt(14)
                r = p.add_run(main_title)
                rPr = r._element.get_or_add_rPr()
                
                rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="方正小标宋简体" w:eastAsia="方正小标宋简体" w:hAnsi="方正小标宋简体" w:cs="方正小标宋简体" w:hint="eastAsia"/>')
                rPr.append(rFonts)
                color = parse_xml(f'<w:color {nsdecls("w")} w:val="C00000"/>')
                rPr.append(color)
                sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="{main_sz}"/>')
                rPr.append(sz)
                szCs = parse_xml(f'<w:szCs {nsdecls("w")} w:val="{main_sz}"/>')
                rPr.append(szCs)
                b = parse_xml(f'<w:b {nsdecls("w")}/>')
                rPr.append(b)
            else:
                for marker, value in placeholders.items():
                    if marker in txt:
                        p.text = txt.replace(marker, value)
                    
        doc.add_page_break()
    else:
        doc = Document()
        sec = doc.sections[0]
        apply_standard_page_setup(sec, page_setup)
        
        # 顶部标头
        p_top = doc.add_paragraph()
        p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf_top = p_top.paragraph_format
        pf_top.space_before = Pt(36)
        pf_top.space_after = Pt(12)
        if header_title:
            r_top = p_top.add_run(header_title)
            set_run_fonts(r_top, fonts.get("h2", "楷体_GB2312"), 16, bold=True, font_en=fonts.get("en", "Times New Roman"))
        
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf_sub = p_sub.paragraph_format
        pf_sub.space_before = Pt(10)
        pf_sub.space_after = Pt(36)
        if sub_title:
            r_sub = p_sub.add_run(sub_title)
            set_run_fonts(r_sub, fonts.get("h2", "楷体_GB2312"), 18, bold=True, font_en=fonts.get("en", "Times New Roman"))
        
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf_title = p_title.paragraph_format
        pf_title.space_before = Pt(30)
        pf_title.space_after = Pt(100)
        r_title = p_title.add_run(main_title)
        set_run_fonts(r_title, fonts.get("title", "方正小标宋简体"), 26, bold=True, color_rgb=RGBColor(192, 0, 0), font_en=fonts.get("en", "Times New Roman"))
        
        sec_toc = doc.add_section(WD_SECTION_START.NEW_PAGE)
        apply_standard_page_setup(sec_toc, page_setup)
        
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf_th = p_toc_head.paragraph_format
    pf_th.space_before = Pt(18)
    pf_th.space_after = Pt(16)
    r_th = p_toc_head.add_run("目  录")
    set_run_fonts(r_th, fonts.get("title", "方正小标宋简体"), 22, bold=True, font_en=fonts.get("en", "Times New Roman"))
    
    for item in flat_nodes:
        if item.get("include_in_toc") is False:
            continue
        if item.get("level", 1) > 2 and item.get("include_h3_in_toc") is False:
            continue
            
        t_text = item.get("title") or item.get("toc_title", "")
        lvl = item.get("level", 1)
        bm_name = item.get("bookmark_name", "")
        pg = exact_pages.get(item.get("title", ""), 1)
        add_toc_entry_with_hyperlink(doc, t_text, pg, lvl, bookmark_name=bm_name, fonts=fonts)
        
    out_cover_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_cover_path))


def build_toc_and_body_doc(
    config: ProjectConfig,
    final_tree: List[Dict[str, Any]],
    flat_nodes: List[Dict[str, Any]],
    exact_pages: Dict[str, int],
    source_dir: Path,
    out_toc_body_path: Path,
    resolved_files: Optional[Mapping[str, Path]] = None
):
    """根据真实打印物理页码生成独立的《目录+正文》交付文档（正文编页码从1开始）"""
    doc_full = Document()
    neutralize_document_styles(doc_full)
    fonts = config.fonts
    page_setup = config.page_setup
    
    # 1. 目录分节（无页脚页码）
    sec_toc = doc_full.sections[0]
    apply_standard_page_setup(sec_toc, page_setup)
    sec_toc.footer.is_linked_to_previous = False
    
    p_toc_head = doc_full.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf_th = p_toc_head.paragraph_format
    pf_th.space_before = Pt(18)
    pf_th.space_after = Pt(16)
    r_th = p_toc_head.add_run("目  录")
    set_run_fonts(r_th, fonts.get("title", "方正小标宋简体"), 22, bold=True, font_en=fonts.get("en", "Times New Roman"))
    
    for item in flat_nodes:
        if item.get("include_in_toc") is False:
            continue
        if item.get("level", 1) > 2 and item.get("include_h3_in_toc") is False:
            continue
            
        t_text = item.get("title") or item.get("toc_title", "")
        lvl = item.get("level", 1)
        bm_name = item.get("bookmark_name", "")
        pg = exact_pages.get(item.get("title", ""), 1)
        add_toc_entry_with_hyperlink(doc_full, t_text, pg, lvl, bookmark_name=bm_name, fonts=fonts)
        
    # 2. 正文分节（从第 1 页开始编号）
    sec_body = doc_full.add_section(WD_SECTION_START.NEW_PAGE)
    apply_standard_page_setup(sec_body, page_setup)
    sec_body.footer.is_linked_to_previous = False
    setup_footer(sec_body, start_page=1, font_en=fonts.get("en", "Times New Roman"))
    
    LAST_RENDERED_LANDSCAPE[0] = False
    for idx, root_node in enumerate(final_tree):
        render_tree_node_recursive(
            doc_full, root_node, source_dir, fonts,
            is_first_section=(idx == 0),
            is_first_child=False,
            exact_pages=exact_pages,
            resolved_files=resolved_files
        )
        
    out_toc_body_path.parent.mkdir(parents=True, exist_ok=True)
    doc_full.save(str(out_toc_body_path))


def build_toc_and_body_from_existing_doc(
    config: ProjectConfig,
    body_doc_path: Path,
    flat_nodes: List[Dict[str, Any]],
    exact_pages: Dict[str, int],
    out_path: Path,
) -> None:
    """为保留原版式的单一 DOCX 在前端插入目录，而不重新渲染正文。"""
    body_doc = Document(str(body_doc_path))
    front_doc = Document()
    neutralize_document_styles(front_doc)
    apply_standard_page_setup(front_doc.sections[0], config.page_setup)
    title = front_doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(18)
    title.paragraph_format.space_after = Pt(16)
    set_run_fonts(title.add_run("目 录"), config.fonts.get("title", "方正小标宋简体"), 22, bold=True, font_en=config.fonts.get("en", "Times New Roman"))
    for item in flat_nodes:
        if item.get("include_in_toc") is False:
            continue
        add_toc_entry_with_hyperlink(
            front_doc, item["title"], exact_pages.get(item["title"], 1), item["level"],
            bookmark_name=item["bookmark_name"], fonts=config.fonts,
        )

    section_break = parse_xml(f'''<w:p {nsdecls("w")}><w:pPr><w:sectPr>
        <w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2100" w:right="1474" w:bottom="1985" w:left="1587" w:header="851" w:footer="992" w:gutter="0"/>
        <w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>''')
    body = body_doc.element.body
    body.insert(0, section_break)
    for element in reversed(list(front_doc.element.body)):
        if element.tag != qn("w:sectPr"):
            body.insert(0, copy_element_with_rels(element, front_doc, body_doc))

    payload = io.BytesIO()
    body_doc.save(payload)
    payload.seek(0)
    final_doc = Document(payload)
    first_section = final_doc.sections[0]
    first_section.footer.is_linked_to_previous = False
    for paragraph in first_section.footer.paragraphs:
        paragraph.text = ""
    body_section = final_doc.sections[1]
    body_section.footer.is_linked_to_previous = False
    setup_footer(body_section, start_page=1, font_en=config.fonts.get("en", "Times New Roman"))
    for section in final_doc.sections[2:]:
        section.footer.is_linked_to_previous = True
        page_number = section._sectPr.find(qn("w:pgNumType"))
        if page_number is not None and qn("w:start") in page_number.attrib:
            del page_number.attrib[qn("w:start")]
    out_path.parent.mkdir(parents=True, exist_ok=True)
    final_doc.save(str(out_path))


class UnifiedSynthesizer:
    """通用材料排版与两趟编译调度器"""

    @classmethod
    def plan(
        cls,
        source_dir: Path,
        manifest_path: Optional[Path] = None
    ) -> Dict[str, Any]:
        """只读扫描输入目录，返回构建前可供人工确认的计划。"""
        source_path = Path(source_dir).resolve()
        if not source_path.exists() or not (source_path.is_dir() or source_path.is_file()):
            raise BuildError(f"输入路径不存在: {source_path}")
        config_dir = source_path if source_path.is_dir() else source_path.parent
        try:
            config = load_project_config(config_dir, manifest_path, project_name=source_path.stem)
        except ConfigError as exc:
            raise BuildError(str(exc)) from exc

        if source_path.is_file():
            if config.source["strategy"] != "highlighted_docx":
                raise BuildError(f"文件型输入只支持 highlighted_docx 策略: {source_path.name}")
            if source_path.suffix.lower() != ".docx":
                raise BuildError(f"highlighted_docx 仅支持 DOCX 文件: {source_path.name}")
            from .highlighted_docx import extract_and_format_highlighted_headings

            headings = extract_and_format_highlighted_headings(
                Document(str(source_path)), config.source.get("highlight", {}), config.fonts, normalize=False,
            )
            if not headings:
                raise BuildError(f"未在 {source_path.name} 发现黄色高亮标题。")
            return {
                "project": config.project_name,
                "source_dir": source_path,
                "config_loaded": config.manifest_path is not None,
                "template": config.get_template_path(),
                "root_count": sum(item["level"] == 1 for item in headings),
                "node_count": len(headings),
                "nodes": [{**item, "type": "docx", "file": source_path.name} for item in headings],
                "legacy_doc_count": 0,
                "pptx_count": 0,
            }

        try:
            final_tree = build_outline(source_path, config)
        except SourceStrategyError as exc:
            raise BuildError(str(exc)) from exc
        flat_nodes = flatten_tree_nodes(final_tree)
        source_files = [p for p in source_path.rglob("*") if p.is_file() and not p.name.startswith("~$")]
        return {
            "project": config.project_name,
            "source_dir": source_path,
            "config_loaded": config.manifest_path is not None,
            "template": config.get_template_path(),
            "root_count": len(final_tree),
            "node_count": len(flat_nodes),
            "nodes": flat_nodes,
            "legacy_doc_count": sum(p.suffix.lower() == ".doc" for p in source_files),
            "pptx_count": sum(p.suffix.lower() == ".pptx" for p in source_files),
        }
    
    @classmethod
    def synthesize(
        cls,
        source_dir: Path,
        output_dir: Optional[Path] = None,
        manifest_path: Optional[Path] = None
    ) -> Dict[str, Path]:
        source_path = Path(source_dir).resolve()
        if not source_path.exists() or not (source_path.is_dir() or source_path.is_file()):
            raise BuildError(f"输入路径不存在: {source_path}")
        config_dir = source_path if source_path.is_dir() else source_path.parent
        try:
            config = load_project_config(config_dir, manifest_path, project_name=source_path.stem)
        except ConfigError as exc:
            raise BuildError(str(exc)) from exc
        proj_name = config.project_name

        if source_path.is_file():
            if config.source["strategy"] != "highlighted_docx":
                raise BuildError(f"文件型输入只支持 highlighted_docx 策略: {source_path.name}")
            return cls._synthesize_highlighted_docx(source_path, output_dir, config)

        source_dir = source_path
        
        print("\n" + "=" * 65)
        print(f"=== 开始合成项目: 《{proj_name}》 ===")
        print("=" * 65)
        
        # 1. 确定输出目录与隔离工作目录。旧交付物只会在所有检查通过后替换。
        if output_dir:
            out_base = Path(output_dir).resolve()
        else:
            out_base = Path(f"output/{proj_name}").resolve()
        out_base.mkdir(parents=True, exist_ok=True)
        run_dir = out_base / ".work" / f"run-{uuid.uuid4().hex}"
        run_dir.mkdir(parents=True, exist_ok=False)
        set_pdf_render_cache_dir(run_dir / "pdf-render-cache")
        print(f"  隔离工作目录: {run_dir}")

        cover_toc_path = out_base / config.output["cover_toc"]
        toc_body_path = out_base / config.output["toc_body"]
        compiled_docx_path = out_base / config.output["compiled_document"]
        staged_cover_toc_path = run_dir / cover_toc_path.name
        staged_toc_body_path = run_dir / toc_body_path.name
        staged_compiled_docx_path = run_dir / compiled_docx_path.name
        
        # 3. 自动发现大纲树并应用复写
        print(f"[阶段 1/4] 扫描目录构建大纲树: {source_dir}")
        try:
            final_tree = build_outline(source_dir, config)
        except SourceStrategyError as exc:
            raise BuildError(str(exc)) from exc
        flat_nodes = flatten_tree_nodes(final_tree)
        if not flat_nodes:
            raise BuildError("未发现可合成的输入材料。")
        print(f"  大纲树就绪: 顶级大纲 {len(final_tree)} 项，全量节点 {len(flat_nodes)} 项")

        # 4. 所有 Office 转换只写入工作目录，输入材料保持只读。
        print("  准备旧版 Office 文件的隔离转换产物...")
        resolved_files = prepare_conversions(source_dir, run_dir)

        # 5. 第一趟编译：构建纯正文骨架以精确计算物理页码
        print(f"\n[阶段 2/4] 第一趟编译：生成正文骨架文档...")
        doc_body = Document()
        neutralize_document_styles(doc_body)
        sec_body = doc_body.sections[0]
        apply_standard_page_setup(sec_body, config.page_setup)
        setup_footer(sec_body, start_page=1, font_en=config.fonts.get("en", "Times New Roman"))
        
        LAST_RENDERED_LANDSCAPE[0] = False
        for idx, root_node in enumerate(final_tree):
            render_tree_node_recursive(
                doc_body, root_node, source_dir, config.fonts,
                is_first_section=(idx == 0),
                is_first_child=False,
                resolved_files=resolved_files
            )
            
        doc_body.save(str(staged_compiled_docx_path))
        print(f"  正文骨架构建完成: {staged_compiled_docx_path}")
        
        # 6. 本地 Word 导出 PDF 反查实际物理页码（包含各小文档内部目录项）
        print(f"\n[阶段 3/4] 本地 Word 引擎排版与真实物理打印页码反查...")
        try:
            exact_pages = get_exact_printed_heading_pages(
                str(staged_compiled_docx_path), flat_nodes, source_dir=source_dir
            )
        except OfficeExportError as exc:
            raise BuildError(str(exc)) from exc
        
        # 7. 第二趟编译：生成《封面+目录》与《目录+正文》交付文档
        print(f"\n[阶段 4/4] 第二趟编译：生成交付文档...")
        build_cover_and_toc_doc(config, flat_nodes, exact_pages, staged_cover_toc_path)
        print(f"  1. 封面与目录构建完成: {staged_cover_toc_path}")
        
        build_toc_and_body_doc(
            config, final_tree, flat_nodes, exact_pages, source_dir, staged_toc_body_path,
            resolved_files=resolved_files
        )
        print(f"  2. 目录与正文构建完成: {staged_toc_body_path}")

        # 8. 发布前强制进行最终文档页级质检。
        print("\n[发布前检查] 使用 Microsoft Word 导出最终文档并执行页级 QA...")
        qa_pdf_path = run_dir / f"{proj_name}_qa.pdf"
        if not export_docx_to_pdf(str(staged_toc_body_path), str(qa_pdf_path)):
            raise BuildError("无法导出最终文档用于 QA；旧交付物未被替换。")
        if not run_qa_assertions(str(qa_pdf_path), ignore_front_pages=1):
            raise BuildError(f"最终文档未通过页级 QA；请查看工作目录: {run_dir}")

        # 9. 所有检查通过后再发布，避免中途失败破坏既有交付物。
        for staged_path, published_path in (
            (staged_compiled_docx_path, compiled_docx_path),
            (staged_cover_toc_path, cover_toc_path),
            (staged_toc_body_path, toc_body_path),
        ):
            os.replace(staged_path, published_path)
        
        print("\n" + "-" * 65)
        print(f"《{proj_name}》全部交付文档已输出至: {out_base}")
        print("-" * 65 + "\n")
        
        return {
            "cover_toc": cover_toc_path,
            "toc_body": toc_body_path,
            "compiled_document": compiled_docx_path
        }

    @classmethod
    def _synthesize_highlighted_docx(
        cls, source_path: Path, output_dir: Optional[Path], config: ProjectConfig
    ) -> Dict[str, Path]:
        """统一引擎的单文件高亮标题路径；不依赖任何项目专项脚本。"""
        from .highlighted_docx import (
            extract_and_format_highlighted_headings,
            find_heading_pages_in_pdf,
            normalize_document_pagination,
        )

        proj_name = config.project_name
        out_base = Path(output_dir).resolve() if output_dir else Path(f"output/{proj_name}").resolve()
        out_base.mkdir(parents=True, exist_ok=True)
        run_dir = out_base / ".work" / f"run-{uuid.uuid4().hex}"
        run_dir.mkdir(parents=True, exist_ok=False)
        compiled_path = out_base / config.output["compiled_document"]
        cover_path = out_base / config.output["cover_toc"]
        toc_body_path = out_base / config.output["toc_body"]
        staged_compiled = run_dir / compiled_path.name
        staged_cover = run_dir / cover_path.name
        staged_toc_body = run_dir / toc_body_path.name

        print(f"\n=== 开始合成项目: 《{proj_name}》（高亮标题 DOCX） ===")
        document = Document(str(source_path))
        headings = extract_and_format_highlighted_headings(
            document, config.source.get("highlight", {}), config.fonts,
        )
        if not headings:
            raise BuildError(f"未在 {source_path.name} 发现符合配置的高亮标题。")
        normalize_document_pagination(document)
        document.save(str(staged_compiled))
        page_reference_pdf = run_dir / f"{proj_name}_pageref.pdf"
        if not export_docx_to_pdf(str(staged_compiled), str(page_reference_pdf)):
            raise BuildError("无法取得 Microsoft Word 的精确物理页码；已停止构建。")
        exact_pages = find_heading_pages_in_pdf(str(page_reference_pdf), headings)
        build_cover_and_toc_doc(config, headings, exact_pages, staged_cover)
        build_toc_and_body_from_existing_doc(config, staged_compiled, headings, exact_pages, staged_toc_body)
        qa_pdf = run_dir / f"{proj_name}_qa.pdf"
        if not export_docx_to_pdf(str(staged_toc_body), str(qa_pdf)):
            raise BuildError("无法导出最终文档用于 QA；旧交付物未被替换。")
        if not run_qa_assertions(str(qa_pdf), ignore_front_pages=1):
            raise BuildError(f"最终文档未通过页级 QA；请查看工作目录: {run_dir}")
        for staged, published in ((staged_compiled, compiled_path), (staged_cover, cover_path), (staged_toc_body, toc_body_path)):
            os.replace(staged, published)
        return {"compiled_document": compiled_path, "cover_toc": cover_path, "toc_body": toc_body_path}
