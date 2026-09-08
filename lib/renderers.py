# -*- coding: utf-8 -*-
"""
多源材料渲染调度模块
包含 Word 节点克隆与样式展平、PDF 渲染排版、图片居中与大纲树递归渲染调度
"""
import io
import re
import copy
from pathlib import Path
from typing import Dict, Any, List, Optional
from PIL import Image
import pymupdf
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from .styles import (
    DEFAULT_FONTS,
    apply_standard_page_setup,
    setup_footer,
    merge_styles_into_doc,
    inline_style_properties,
    add_heading_paragraph,
    set_run_fonts
)
from .layout import RenderContext, apply_section_spec, adapt_table_to_content_box
from .sanitizers import PdfPageSanitizer
from .scanner import CN_NUMS

# 记录上一个节点是否以横版分节结尾
LAST_RENDERED_LANDSCAPE = [False]


def _context_is_preserve(context: Optional[RenderContext]) -> bool:
    policy = getattr(context, "formatting_policy", None) if context is not None else None
    return bool(policy and policy.mode == "preserve")


def _context_preserves_geometry(context: Optional[RenderContext]) -> bool:
    policy = getattr(context, "formatting_policy", None) if context is not None else None
    return bool(policy and (policy.mode == "preserve" or (policy.mode == "mixed" and policy.page_policy == "source")))


def _copy_section_stories(source_section, target_section) -> None:
    """复制简单页眉页脚内容，保持 preserve/source 下的故事流不被默认模板替换。"""
    for attr in ("header", "footer", "first_page_header", "first_page_footer", "even_page_header", "even_page_footer"):
        try:
            source_story = getattr(source_section, attr)
            target_story = getattr(target_section, attr)
            target_story.is_linked_to_previous = False
            for child in list(target_story._element):
                target_story._element.remove(child)
            for child in source_story._element:
                target_story._element.append(copy.deepcopy(child))
        except Exception:
            continue


def copy_element_with_rels(elem, src_doc: Document, dst_doc: Document):
    """深拷贝 XML 节点，并将 OLE 对象和浮动图片处理为可安全合并的静态内容。"""
    cloned = copy.deepcopy(elem)
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'o': 'urn:schemas-microsoft-com:office:office',
        'v': 'urn:schemas-microsoft-com:vml'
    }
    
    # 将包含 v:imagedata 的 w:object 转为静态 w:pict，避免保留 Excel 动态链接。
    for obj in list(cloned.findall(f'.//{{{ns["w"]}}}object')):
        shapes = obj.findall(f'.//{{{ns["v"]}}}shape')
        if shapes:
            pict = OxmlElement('w:pict')
            for sh in shapes:
                sh.attrib.pop(f'{{{ns["o"]}}}ole', None)
                sh.attrib.pop('ole', None)
                style = sh.get('style', '')
                if 'position:absolute' in style:
                    cleaned_style = ';'.join([s for s in style.split(';') if not any(s.strip().startswith(k) for k in ['position', 'margin-left', 'margin-top', 'z-index', 'left', 'top'])])
                    sh.set('style', cleaned_style)
                pict.append(sh)
            parent = obj.getparent()
            if parent is not None:
                parent.replace(obj, pict)
        else:
            parent = obj.getparent()
            if parent is not None:
                parent.remove(obj)

    # 移除残留的 OLEObject
    for ole in list(cloned.findall(f'.//{{{ns["o"]}}}OLEObject')):
        parent = ole.getparent()
        if parent is not None:
            parent.remove(ole)

    # 将 v:textbox 内的图片转为普通内联图片，避免绝对定位对象相互遮挡。
    for tb in list(cloned.findall(f'.//{{{ns["v"]}}}textbox')):
        inner_shapes = tb.findall(f'.//{{{ns["v"]}}}shape')
        tb_shape = tb.getparent()
        if tb_shape is not None and inner_shapes:
            pict_parent = tb_shape.getparent()
            if pict_parent is not None:
                run_parent = pict_parent.getparent()
                if run_parent is not None:
                    for ish in inner_shapes:
                        style = ish.get('style', '')
                        cleaned_style = ';'.join([s for s in style.split(';') if not any(s.strip().startswith(k) for k in ['position', 'margin-left', 'margin-top', 'z-index', 'left', 'top'])])
                        ish.set('style', cleaned_style)
                        new_pict = OxmlElement('w:pict')
                        new_pict.append(ish)
                        run_parent.append(new_pict)
                    run_parent.remove(pict_parent)

    # 3. 对所有残留的 v:shape 统一清除 position:absolute
    for sh in cloned.findall(f'.//{{{ns["v"]}}}shape'):
        style = sh.get('style', '')
        if 'position:absolute' in style:
            cleaned_style = ';'.join([s for s in style.split(';') if not any(s.strip().startswith(k) for k in ['position', 'margin-left', 'margin-top', 'z-index', 'left', 'top'])])
            sh.set('style', cleaned_style)
            
    # 4. 重新映射所有图片资源 ID，清理悬空无效关系
    r_embeds = cloned.findall('.//*[@r:embed]', ns)
    for el in r_embeds:
        old_rid = el.get(f'{{{ns["r"]}}}embed')
        if old_rid in src_doc.part.rels:
            rel = src_doc.part.rels[old_rid]
            if "image" in rel.reltype:
                new_rid, _ = dst_doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                el.set(f'{{{ns["r"]}}}embed', new_rid)
            else:
                el.attrib.pop(f'{{{ns["r"]}}}embed', None)
        else:
            el.attrib.pop(f'{{{ns["r"]}}}embed', None)
                
    r_links = cloned.findall('.//*[@r:id]', ns)
    for el in r_links:
        old_rid = el.get(f'{{{ns["r"]}}}id')
        if old_rid in src_doc.part.rels:
            rel = src_doc.part.rels[old_rid]
            if "image" in rel.reltype:
                new_rid, _ = dst_doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                el.set(f'{{{ns["r"]}}}id', new_rid)
            else:
                el.attrib.pop(f'{{{ns["r"]}}}id', None)
        else:
            el.attrib.pop(f'{{{ns["r"]}}}id', None)
                
    return cloned


def _prepare_source_bookmark_remap(
    doc_src: Document,
    target: Document,
    docx_path: Path,
    extra_names=(),
):
    """Prepare collision-safe bookmark names and IDs for one DOCX import.

    Directory-tree rendering imports each source document by cloning its XML
    directly.  Unlike ``PackageImporter``, that path previously left source
    bookmark names untouched, so two valid sources containing the same local
    bookmark (for example ``fig_caption``) produced an invalid merged package.
    The remap is scoped to one source import and is applied consistently to
    every cloned paragraph/table, including simple-field instructions.
    """
    existing_names = {
        marker.get(qn("w:name"))
        for marker in target.element.body.iter(qn("w:bookmarkStart"))
        if marker.get(qn("w:name"))
    }
    source_names = []
    source_ids = []
    source_name_ids = {}
    for marker in doc_src.element.body.iter(qn("w:bookmarkStart")):
        name = marker.get(qn("w:name"))
        marker_id = marker.get(qn("w:id"))
        if name and name not in source_names:
            source_names.append(name)
            if marker_id is not None:
                source_name_ids[name] = marker_id
        if marker_id is not None and marker_id not in source_ids:
            source_ids.append(marker_id)
    for name in extra_names:
        if name and name not in source_names:
            source_names.append(name)

    name_map = {}
    source_token = hashlib.sha256(str(Path(docx_path).resolve()).encode("utf-8")).hexdigest()[:10]
    for index, old_name in enumerate(source_names, 1):
        new_name = old_name
        if new_name in existing_names:
            new_name = f"_SynthBm_{source_token}_{index:02d}"
            while new_name in existing_names:
                index += 1
                new_name = f"_SynthBm_{source_token}_{index:02d}"
        existing_names.add(new_name)
        name_map[old_name] = new_name

    existing_ids = [
        int(marker.get(qn("w:id")))
        for marker in target.element.body.iter(qn("w:bookmarkStart"))
        if (marker.get(qn("w:id")) or "").lstrip("-").isdigit()
    ]
    next_id = max(existing_ids, default=-1) + 1
    id_map = {}
    for old_id in source_ids:
        id_map[old_id] = str(next_id)
        next_id += 1
    name_id_map = {}
    for name in source_names:
        old_id = source_name_ids.get(name)
        if old_id in id_map:
            name_id_map[name] = id_map[old_id]
        else:
            name_id_map[name] = str(next_id)
            next_id += 1
    return name_map, id_map, name_id_map


def _remap_source_bookmarks(element, name_map, id_map):
    """Apply one source import's bookmark and field-reference remap."""
    field_pattern = re.compile(
        r"(\b(?:REF|PAGEREF|STYLEREF)\s+)([^\s\\]+)",
        flags=re.IGNORECASE,
    )
    for item in element.iter():
        if item.tag == qn("w:bookmarkStart"):
            old_id = item.get(qn("w:id"))
            if old_id in id_map:
                item.set(qn("w:id"), id_map[old_id])
            old_name = item.get(qn("w:name"))
            if old_name in name_map:
                item.set(qn("w:name"), name_map[old_name])
        elif item.tag == qn("w:bookmarkEnd"):
            old_id = item.get(qn("w:id"))
            if old_id in id_map:
                item.set(qn("w:id"), id_map[old_id])
        elif item.tag == qn("w:hyperlink"):
            old_name = item.get(qn("w:anchor"))
            if old_name in name_map:
                item.set(qn("w:anchor"), name_map[old_name])
        elif item.tag == qn("w:instrText") and item.text:
            item.text = field_pattern.sub(
                lambda match: match.group(1) + name_map.get(match.group(2), match.group(2)),
                item.text,
            )
        elif item.tag == qn("w:fldSimple"):
            instruction = item.get(qn("w:instr"))
            if instruction:
                item.set(
                    qn("w:instr"),
                    field_pattern.sub(
                        lambda match: match.group(1) + name_map.get(match.group(2), match.group(2)),
                        instruction,
                    ),
                )


def _remap_source_note_references(element, footnote_ids, endnote_ids):
    """Rewrite note references to the IDs allocated in the target package."""
    for reference in element.iter(qn("w:footnoteReference")):
        old_id = reference.get(qn("w:id"))
        if old_id in footnote_ids:
            reference.set(qn("w:id"), footnote_ids[old_id])
    for reference in element.iter(qn("w:endnoteReference")):
        old_id = reference.get(qn("w:id"))
        if old_id in endnote_ids:
            reference.set(qn("w:id"), endnote_ids[old_id])


def append_element_to_body(doc: Document, elem):
    """将 XML 节点插入文档正文尾部，保持分节属性处于最后"""
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(elem)
    else:
        body.append(elem)


import hashlib
import os
from typing import Optional

# 由构建调度器指定本次运行的缓存位置；延迟创建，避免 import 时写入调用方目录。
CACHE_DIR: Optional[Path] = None


def set_pdf_render_cache_dir(cache_dir: Path):
    """设置 PDF 页渲染缓存目录。"""
    global CACHE_DIR
    CACHE_DIR = Path(cache_dir)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


def _get_pdf_render_cache_dir() -> Path:
    global CACHE_DIR
    if CACHE_DIR is None:
        configured = os.environ.get("DOCUMENT_SYNTHESIS_CACHE_DIR")
        set_pdf_render_cache_dir(Path(configured) if configured else Path(".cache/pdf_render"))
    return CACHE_DIR


def get_cached_pdf_page_image(
    pdf_path: Path,
    page_idx: int,
    dpi: int = 300,
    sanitize_footer: bool = True,
    crop_whitespace: bool = True,
    cache_version: str = "v1",
    format_hash: str = "",
    page_geometry: str = "",
    formatting_strategy: str = "",
) -> bytes:
    """基于文件修改状态、页面编号、DPI与清洗策略的哈希缓存，避免不同清洗策略复用错误缓存"""
    stat = pdf_path.stat()
    # 使用绝对路径与策略哈希，避免不同配置复用错误图片。
    key = (
        f"{pdf_path.resolve()}_{stat.st_mtime_ns}_{stat.st_size}_p{page_idx}"
        f"_dpi{dpi}_sf{int(sanitize_footer)}_cw{int(crop_whitespace)}"
        f"_fmt{format_hash}_page{page_geometry}_strategy{formatting_strategy}_{cache_version}"
    )
    h = hashlib.md5(key.encode('utf-8')).hexdigest()
    cache_file = _get_pdf_render_cache_dir() / f"{h}.png"
    
    if cache_file.exists():
        return cache_file.read_bytes()
        
    pdf_doc = pymupdf.open(str(pdf_path))
    zoom = dpi / 72.0
    mat = pymupdf.Matrix(zoom, zoom)
    page = pdf_doc[page_idx]
    if sanitize_footer:
        PdfPageSanitizer.erase_footer_page_number(page, page_idx)
    
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    if crop_whitespace:
        img = PdfPageSanitizer.crop_vertical_whitespace(img)
    
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=False)
    data = buf.getvalue()
    
    try:
        cache_file.write_bytes(data)
    except Exception:
        pass
        
    pdf_doc.close()
    return data


def render_pdf_file(
    doc: Document,
    pdf_path: Path,
    has_headings_on_page: bool = True,
    dpi: int = 300,
    render_context: Optional[RenderContext] = None,
):
    """PDF 页面渲染排版，结合可用高度自适应等比缩放"""
    pdf_doc = pymupdf.open(str(pdf_path))
    total_pages = len(pdf_doc)
    
    for i in range(total_pages):
        page_geometry = ""
        formatting_strategy = ""
        format_hash = ""
        if render_context is not None:
            box = render_context.get_content_box()
            page_geometry = f"{box.page_width_twip}x{box.page_height_twip}:{box.content_width_twip}x{box.content_height_twip}"
            formatting_strategy = getattr(render_context.formatting_policy, "mode", "")
            format_hash = getattr(render_context.resolved_format, "content_hash", "")
        img_bytes = get_cached_pdf_page_image(
            pdf_path,
            i,
            dpi=dpi,
            format_hash=format_hash,
            page_geometry=page_geometry,
            formatting_strategy=formatting_strategy,
        )
        img = Image.open(io.BytesIO(img_bytes))
        
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p_img.paragraph_format
        # 不插入独立的换页段落：前一张图恰好占满页面时，独立换页段落会自己
        # 溢出成空白页。将换页属性附在下一张实际图片段落上才是原子操作。
        if i > 0:
            pf.page_break_before = True
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        w_px, h_px = img.size
        aspect = h_px / w_px
        
        rect = pdf_doc[i].rect
        content_box = render_context.get_content_box() if render_context is not None else None
        default_w_cm = content_box.content_width_cm if content_box is not None else (15.6 if rect.width > rect.height else 15.0)
        
        # 首页若与标题同页，按多级标题场景预留足够高度，防止标题成为孤页。
        if i == 0 and has_headings_on_page:
            max_h_cm = min(15.5, content_box.content_height_cm) if content_box is not None else 15.5
        else:
            max_h_cm = content_box.content_height_cm if content_box is not None else 22.0
            
        w_cm = min(default_w_cm, max_h_cm / aspect)
        p_img.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(w_cm))
        
    pdf_doc.close()


def extract_numbering_rules(doc_src: Document) -> Dict[str, Any]:
    """提取源文档中的编号规则（数字列表、项目符号等）"""
    num_info = {}
    try:
        num_part = doc_src.part.numbering_part
    except Exception:
        num_part = None
        
    if num_part is not None:
        root = num_part._element
        abs_map = {}
        for abs_el in root.findall(qn('w:abstractNum')):
            abs_id = abs_el.get(qn('w:abstractNumId'))
            lvl = abs_el.find(qn('w:lvl'))
            if lvl is not None:
                numFmt = lvl.find(qn('w:numFmt')).get(qn('w:val')) if lvl.find(qn('w:numFmt')) is not None else 'decimal'
                lvlText = lvl.find(qn('w:lvlText')).get(qn('w:val')) if lvl.find(qn('w:lvlText')) is not None else '%1.'
                start = int(lvl.find(qn('w:start')).get(qn('w:val'))) if lvl.find(qn('w:start')) is not None else 1
                abs_map[abs_id] = {'fmt': numFmt, 'lvlText': lvlText, 'start': start}
        for num_el in root.findall(qn('w:num')):
            num_id = num_el.get(qn('w:numId'))
            abs_ref = num_el.find(qn('w:abstractNumId')).get(qn('w:val')) if num_el.find(qn('w:abstractNumId')) is not None else None
            if abs_ref in abs_map:
                num_info[num_id] = copy.deepcopy(abs_map[abs_ref])
    return num_info


def format_numbering_prefix(rule: Dict[str, Any], count: int) -> str:
    """根据编号规则与当前序号生成前缀文本（如 '1.  '）"""
    fmt = rule.get('fmt', 'decimal')
    lvlText = rule.get('lvlText', '%1.')
    if fmt == 'decimal':
        s = str(count)
    elif fmt == 'chineseCounting':
        idx = count - 1
        s = CN_NUMS[idx] if 0 <= idx < len(CN_NUMS) else str(count)
    elif fmt == 'upperLetter':
        s = chr(64 + count) if 1 <= count <= 26 else str(count)
    elif fmt == 'decimalEnclosedCircleChinese':
        circles = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩', '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳']
        idx = count - 1
        s = circles[idx] if 0 <= idx < len(circles) else str(count)
    else:
        s = str(count)
    res = lvlText.replace('%1', s)
    if not res.endswith(' ') and not res.endswith('\t'):
        res += '  '
    return res


def split_element_by_soft_breaks(p_elem) -> List[Any]:
    """将包含软回车（<w:br/>）的段落拆分成独立的段落，杜绝分散对齐与字符拉伸"""
    soft_brs = [b for b in p_elem.findall('.//' + qn('w:br')) if b.get(qn('w:type')) is None]
    if not soft_brs:
        return [p_elem]
        
    pPr = p_elem.find(qn('w:pPr'))
    
    def make_pPr():
        if pPr is None:
            return None
        cp = copy.deepcopy(pPr)
        ind = cp.find(qn('w:ind'))
        if ind is not None and ind.get(qn('w:hanging')) is not None:
            ind.attrib.pop(qn('w:hanging'), None)
            ind.attrib.pop(qn('w:hangingChars'), None)
        return cp

    new_paragraphs = []
    curr_p = OxmlElement('w:p')
    pPr_clone = make_pPr()
    if pPr_clone is not None:
        curr_p.append(pPr_clone)
    new_paragraphs.append(curr_p)
    
    for child in list(p_elem):
        if child.tag.endswith('pPr'):
            continue
        if child.tag.endswith('r'):
            r_children = list(child)
            rPr = child.find(qn('w:rPr'))
            has_br = any(rc.tag.endswith('br') and rc.get(qn('w:type')) is None for rc in r_children)
            if not has_br:
                curr_p.append(copy.deepcopy(child))
            else:
                curr_r = OxmlElement('w:r')
                if rPr is not None:
                    curr_r.append(copy.deepcopy(rPr))
                for rc in r_children:
                    if rc.tag.endswith('rPr'):
                        continue
                    if rc.tag.endswith('br') and rc.get(qn('w:type')) is None:
                        if len(curr_r) > (1 if rPr is not None else 0):
                            curr_p.append(curr_r)
                        curr_p = OxmlElement('w:p')
                        pPr_clone = make_pPr()
                        if pPr_clone is not None:
                            curr_p.append(pPr_clone)
                        new_paragraphs.append(curr_p)
                        curr_r = OxmlElement('w:r')
                        if rPr is not None:
                            curr_r.append(copy.deepcopy(rPr))
                    else:
                        curr_r.append(copy.deepcopy(rc))
                if len(curr_r) > (1 if rPr is not None else 0):
                    curr_p.append(curr_r)
        else:
            curr_p.append(copy.deepcopy(child))
            
    result = []
    for p in new_paragraphs:
        txt = ''.join([t.text or '' for t in p.findall('.//' + qn('w:t'))]).strip()
        if txt or p.findall('.//' + qn('w:drawing')):
            result.append(p)
    return result or [p_elem]


def inline_numbering_to_paragraph(p_elem, num_info: Dict[str, Any], num_counters: Dict[str, int]):
    """将 Word 自动编号内联为显式文本前缀，杜绝模板间 numbering.xml 映射冲突变成黑圆点"""
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is None:
        return
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return
    
    numId_el = numPr.find(qn('w:numId'))
    numId = numId_el.get(qn('w:val')) if numId_el is not None else None
    if not numId:
        return
        
    rule = num_info.get(numId, {'fmt': 'decimal', 'lvlText': '%1.', 'start': 1})
    count = num_counters.get(numId, rule.get('start', 1))
    num_counters[numId] = count + 1
    
    prefix = format_numbering_prefix(rule, count)
    pPr.remove(numPr)
    
    tabs = pPr.find(qn('w:tabs'))
    if tabs is not None:
        pPr.remove(tabs)
        
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        if ind.get(qn('w:hanging')) is not None or ind.get(qn('w:hangingChars')) is not None:
            ind.attrib.pop(qn('w:hanging'), None)
            ind.attrib.pop(qn('w:hangingChars'), None)
        ind.set(qn('w:firstLine'), '480')
        ind.set(qn('w:firstLineChars'), '200')
    else:
        new_ind = parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="480" w:firstLineChars="200"/>')
        pPr.append(new_ind)
        
    first_r = p_elem.find(qn('w:r'))
    prefix_r = OxmlElement('w:r')
    if first_r is not None:
        first_rPr = first_r.find(qn('w:rPr'))
        if first_rPr is not None:
            prefix_r.append(copy.deepcopy(first_rPr))
    prefix_t = OxmlElement('w:t')
    prefix_t.text = prefix
    prefix_t.set(qn('xml:space'), 'preserve')
    prefix_r.append(prefix_t)
    
def sanitize_ole_and_external_links(element):
    """
    清理可能导致 Word 打开时弹窗提示更新外部文件链接的 OLE 对象与外部链接
    1. 将 <w:object> 中的 <o:OLEObject> 移除，并将 <w:object> 转为静态图片容器 <w:pict>，保留其内嵌的预渲染图像快照 (<v:shape><v:imagedata.../>)
    2. 移除任何指向外部工作簿的 LINK/DDE/INCLUDETEXT 域代码，防止 Word 尝试联网/找本地不存在文件
    """
    # 使用 QName/iter 而非带 w: 前缀的 xpath：克隆后的 lxml 元素不一定携带
    # python-docx 的预注册命名空间，带前缀查询会在含 OLE 的真实材料上直接失败。
    for obj in list(element.iter(qn('w:object'))):
        for ole in [item for item in obj.iter() if item.tag.split('}')[-1] == "OLEObject"]:
            parent = ole.getparent()
            if parent is not None:
                parent.remove(ole)
        obj.tag = qn('w:pict')
        
    # 清理 VML shape 上的 o:ole 和 o:gfxdata 外部工作簿引用属性，将其完全脱敏为纯静态图像快照
    for shape in [item for item in element.iter() if item.tag.split('}')[-1] == "shape"]:
        for k in list(shape.attrib.keys()):
            if 'ole' in k.lower() or 'gfxdata' in k.lower():
                shape.attrib.pop(k)
        
    external_field_keywords = ("LINK", "INCLUDETEXT", "DATABASE", "DDE")
    for fld in list(element.iter(qn('w:fldSimple'))):
        instruction = (fld.get(qn('w:instr')) or "").upper()
        if not any(keyword in instruction for keyword in external_field_keywords):
            continue
        parent = fld.getparent()
        if parent is not None:
            idx = parent.index(fld)
            for child in list(fld):
                parent.insert(idx, child)
                idx += 1
            parent.remove(fld)


def create_static_toc_paragraph(title: str, pg_str: str, level: int = 1, fonts: Dict[str, str] = None):
    """构建高保真目录段落（还原源文档的蓝色下划线超链接样式、点号引线、精准页码与规范层级缩进）"""
    fonts = fonts or DEFAULT_FONTS
    font_en = fonts.get("en", "Times New Roman")
    
    if level == 1:
        indent_twips = '0'
        is_bold = True
        font_cn = '宋体'
    elif level == 2:
        indent_twips = '280'
        is_bold = False
        font_cn = '仿宋'
    else:
        indent_twips = '560'
        is_bold = False
        font_cn = '仿宋'

    p = OxmlElement('w:p')
    pPr_xml = f'''<w:pPr {nsdecls("w")}>
        <w:tabs>
            <w:tab w:val="right" w:leader="dot" w:pos="8326"/>
        </w:tabs>
        <w:spacing w:line="240" w:lineRule="auto" w:before="0" w:after="0"/>
        <w:ind w:left="{indent_twips}"/>
    </w:pPr>'''
    p.append(parse_xml(pPr_xml))
    
    # 标题 Run (蓝色单下划线超链接外观)
    r_t = OxmlElement('w:r')
    b_tag = '<w:b/>' if is_bold else ''
    rPr_t = f'''<w:rPr {nsdecls("w")}>
        <w:rFonts w:ascii="{font_en}" w:eastAsia="{font_cn}" w:hAnsi="{font_en}" w:cs="{font_en}"/>
        {b_tag}
        <w:color w:val="0000FF"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:u w:val="single"/>
    </w:rPr>'''
    r_t.append(parse_xml(rPr_t))
    t_el = OxmlElement('w:t')
    t_el.text = title
    r_t.append(t_el)
    p.append(r_t)
    
    # Tab 引导线 Run (黑色无下划线)
    r_tab = OxmlElement('w:r')
    rPr_tab = f'''<w:rPr {nsdecls("w")}>
        <w:rFonts w:ascii="{font_en}" w:eastAsia="宋体" w:hAnsi="{font_en}" w:cs="{font_en}"/>
        <w:color w:val="000000"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
    </w:rPr>'''
    r_tab.append(parse_xml(rPr_tab))
    r_tab.append(OxmlElement('w:tab'))
    p.append(r_tab)
    
    # 页码 Run (黑色无下划线)
    r_pg = OxmlElement('w:r')
    r_pg.append(parse_xml(rPr_tab))
    t_pg = OxmlElement('w:t')
    t_pg.text = str(pg_str)
    r_pg.append(t_pg)
    p.append(r_pg)
    
    return p


def _add_embedded_heading_bookmark(paragraph, node: Dict[str, Any]) -> None:
    """在克隆后的原始正文段落内写入目录锚点，不额外生成可见标题。"""
    p_pr = paragraph.find(qn("w:pPr"))
    insert_at = paragraph.index(p_pr) + 1 if p_pr is not None else 0
    bookmark_id = str(node["bm_id"])
    start = parse_xml(
        f'<w:bookmarkStart {nsdecls("w")} w:id="{bookmark_id}" w:name="{node["bookmark_name"]}"/>'
    )
    end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bookmark_id}"/>')
    paragraph.insert(insert_at, start)
    paragraph.append(end)


def render_docx_file(
    doc: Document,
    docx_path: Path,
    exact_pages: Dict[str, int] = None,
    fonts: Dict[str, str] = None,
    embedded_outline: List[Dict[str, Any]] = None,
    render_context: Optional[RenderContext] = None,
    suppress_first_page_break: bool = False,
):
    """Word 文件保真渲染，支持全动态横竖版分节自适应、内部目录高保真还原与格式微调"""
    doc_src = Document(str(docx_path))
    derived_outline = []
    source_assignments = []
    if embedded_outline is None:
        # A directory-tree DOCX node has no explicit outline children, but
        # the source may still contain PAGEREF/REF fields pointing to the
        # deterministic heading anchors produced by RoleMapper.  Materialize
        # those anchors while importing the source so such fields remain
        # closed and measurable in the merged package.
        try:
            from .role_mapper import RoleMapper

            source_assignments, _ = RoleMapper().map_document(docx_path)
            derived_outline = [
                {
                    "source_text": assignment.title_text or "",
                    "bookmark_name": assignment.bookmark_name,
                    "bm_id": assignment.node_ref.element_path,
                }
                for assignment in source_assignments
                if assignment.bookmark_name and assignment.title_text
            ]
        except Exception:
            derived_outline = []
        embedded_outline = derived_outline
    extra_bookmark_names = [
        item.get("bookmark_name")
        for item in (embedded_outline or [])
        if item.get("bookmark_name")
    ]
    source_bookmark_names, source_bookmark_ids, source_bookmark_name_ids = _prepare_source_bookmark_remap(
        doc_src, doc, docx_path, extra_names=extra_bookmark_names
    )
    from .notes_merger import NotesMerger

    source_footnote_ids = NotesMerger.merge_notes(
        doc_src, doc, is_endnote=False, strict_rels=False
    )
    source_endnote_ids = NotesMerger.merge_notes(
        doc_src, doc, is_endnote=True, strict_rels=False
    )
    normalized_outline = []
    for item in embedded_outline or []:
        normalized = dict(item)
        old_name = normalized.get("bookmark_name")
        if old_name:
            normalized["bookmark_name"] = source_bookmark_names.get(old_name, old_name)
            mapped_id = source_bookmark_name_ids.get(old_name)
            if mapped_id is not None:
                normalized["bm_id"] = int(mapped_id)
        normalized_outline.append(normalized)
    embedded_outline = normalized_outline
    if (
        render_context is not None
        and getattr(render_context, "resolved_format", None) is not None
        and getattr(getattr(render_context, "formatting_policy", None), "mode", "restyle") == "restyle"
        and source_assignments
    ):
        # Directory-tree imports use the low-level clone path below, so apply
        # the target role styles to the source package before cloning.  This
        # keeps a restyle build consistent with the docx_document path while
        # preserving the source's inline semantic objects.
        from .style_applier import apply_roles

        apply_roles(
            doc_src,
            {item.node_ref.element_path: item.role for item in source_assignments},
            render_context.resolved_format,
            policy=render_context.formatting_policy,
        )
    merge_styles_into_doc(doc_src, doc)
    source_sections = list(doc_src.sections)
    preserve_source = _context_is_preserve(render_context)
    preserve_geometry = _context_preserves_geometry(render_context)
    if preserve_geometry and source_sections and doc.sections:
        # 来源边界由 engine 在外层标题写入前创建；这里仅把首节的几何与故事
        # 流绑定到当前目标节，避免把外层标题留在新节之前形成孤行。
        apply_section_spec(doc.sections[-1], source_section=source_sections[0])
        _copy_section_stories(source_sections[0], doc.sections[-1])
    num_info = extract_numbering_rules(doc_src)
    num_counters = {}
    remaining_anchors: Dict[str, List[Dict[str, Any]]] = {}
    for item in embedded_outline or []:
        source_text = item.get("source_text", "").strip()
        if source_text:
            remaining_anchors.setdefault(source_text, []).append(item)

    # 1. 元素分节切片 (依据段落 sectPr 与 body sectPr 自动识别横版/竖版)
    body = doc_src._body._element
    raw_chunks = []
    current_chunk = []

    for elem in body:
        tag = elem.tag.split('}')[-1]
        if tag not in ('p', 'tbl', 'sdt'):
            continue
        current_chunk.append(elem)
        if tag == 'p':
            pPr = elem.find(qn('w:pPr'))
            if pPr is not None:
                sectPr = pPr.find(qn('w:sectPr'))
                if sectPr is not None:
                    pgSz = sectPr.find(qn('w:pgSz'))
                    orient = pgSz.get(qn('w:orient')) if pgSz is not None and pgSz.get(qn('w:orient')) else 'portrait'
                    raw_chunks.append((orient, current_chunk))
                    current_chunk = []

    final_sectPr = body.find(qn('w:sectPr'))
    if final_sectPr is not None:
        pgSz = final_sectPr.find(qn('w:pgSz'))
        orient = pgSz.get(qn('w:orient')) if pgSz is not None and pgSz.get(qn('w:orient')) else 'portrait'
    else:
        orient = 'portrait'
    if current_chunk:
        raw_chunks.append((orient, current_chunk))

    # 合并相邻相同朝向的切片
    chunks = []
    for orient, ch_elems in raw_chunks:
        if chunks and chunks[-1][0] == orient:
            chunks[-1][1].extend(ch_elems)
        else:
            chunks.append([orient, list(ch_elems)])

    # 2. 逐切片渲染
    # 旧项目的文件名特判已移除；保留统一渲染路径。
    is_wen_doc = False
    is_huaxiang_doc = False
    is_keti_doc = False
    is_xize_doc = False
    is_growth_doc = False
    in_wen_expected = False
    in_huaxiang_appendix = False
    is_pyfa_doc = False
    in_pyfa_toc = False
    first_pyfa_body = False
    first_content_element = True
    for chunk_index, (orient, chunk_elems) in enumerate(chunks):
        last_landscape = (
            render_context.last_rendered_landscape
            if render_context is not None else LAST_RENDERED_LANDSCAPE[0]
        )
        need_new_section = False
        if orient == 'landscape' and not last_landscape:
            need_new_section = True
        elif orient == 'portrait' and last_landscape:
            need_new_section = True
            
        if need_new_section:
            sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
            source_section = source_sections[min(chunk_index, len(source_sections) - 1)] if source_sections else None
            if preserve_geometry and source_section is not None:
                apply_section_spec(sec, source_section=source_section)
                _copy_section_stories(source_section, sec)
            elif orient == 'landscape':
                if render_context is not None and render_context.page_spec is not None:
                    from dataclasses import replace
                    base = render_context.page_spec
                    target = replace(base, width_mm=max(base.width_mm, base.height_mm), height_mm=min(base.width_mm, base.height_mm), orientation="landscape")
                    apply_section_spec(sec, target)
                else:
                    sec.orientation = WD_ORIENT.LANDSCAPE
                    sec.page_width = Cm(29.7)
                    sec.page_height = Cm(21.0)
                    sec.top_margin = Cm(2.2)
                    sec.bottom_margin = Cm(2.2)
                    sec.left_margin = Cm(2.8)
                    sec.right_margin = Cm(2.6)
                secPr_land = sec._sectPr
                pgSz = secPr_land.find(qn('w:pgSz'))
                if pgSz is not None and not preserve_geometry:
                    pgSz.set(qn('w:orient'), 'landscape')
                    pgSz.set(qn('w:w'), '16838')
                    pgSz.set(qn('w:h'), '11906')
                pgNumType = secPr_land.find(qn('w:pgNumType'))
                if pgNumType is not None and not preserve_geometry:
                    secPr_land.remove(pgNumType)
                if not preserve_geometry and render_context is not None and render_context.resolved_format:
                    from .style_applier import setup_styled_footer
                    setup_styled_footer(sec, render_context.resolved_format)
                elif not preserve_geometry:
                    setup_footer(sec, font_en='Times New Roman')
                if render_context is not None:
                    render_context.track_landscape(True)
                else:
                    LAST_RENDERED_LANDSCAPE[0] = True
            else:
                if preserve_geometry and source_section is not None:
                    apply_section_spec(sec, source_section=source_section)
                elif render_context is not None and render_context.page_spec is not None:
                    apply_section_spec(sec, render_context.page_spec)
                else:
                    sec.orientation = WD_ORIENT.PORTRAIT
                    sec.page_width = Cm(21.0)
                    sec.page_height = Cm(29.7)
                    apply_standard_page_setup(sec)
                secPr_port = sec._sectPr
                pgSz = secPr_port.find(qn('w:pgSz'))
                if pgSz is not None and not preserve_geometry:
                    pgSz.set(qn('w:orient'), 'portrait')
                    pgSz.set(qn('w:w'), '11906')
                    pgSz.set(qn('w:h'), '16838')
                pgNumType = secPr_port.find(qn('w:pgNumType'))
                if pgNumType is not None and not preserve_geometry:
                    secPr_port.remove(pgNumType)
                if not preserve_geometry and render_context is not None and render_context.resolved_format:
                    from .style_applier import setup_styled_footer
                    setup_styled_footer(sec, render_context.resolved_format)
                elif not preserve_geometry:
                    setup_footer(sec, font_en='Times New Roman')
                if render_context is not None:
                    render_context.track_landscape(False)
                else:
                    LAST_RENDERED_LANDSCAPE[0] = False

        consecutive_empty_p = 0
        has_standalone_cover = is_standalone_cover_doc(docx_path)
        last_elem_was_table = False
        if render_context is not None and doc.sections:
            render_context.current_section = doc.sections[-1]

        for elem in chunk_elems:
            tag = elem.tag.split('}')[-1]
            if tag == 'p':
                splits = [elem] if preserve_source else split_element_by_soft_breaks(elem)
                for sp in splits:
                    if not preserve_source:
                        inline_numbering_to_paragraph(sp, num_info, num_counters)
                    txt = ''.join(sp.itertext()).strip()
                    drawings = sp.findall('.//' + qn('w:drawing')) + sp.findall('.//' + qn('w:pict'))
                    brs = sp.findall('.//' + qn('w:br'))
                    
                    is_empty = (not txt and not drawings)
                    if is_empty and not preserve_source:
                        # 过滤掉连续超过 2 个的冗余空段落，防止原文档中大量空行撑爆版面产生空白页
                        if consecutive_empty_p >= 2:
                            continue
                        consecutive_empty_p += 1
                        continue
                    else:
                        consecutive_empty_p = 0
                        # 若是独立封面文档，移除段首多余的无意义软换行符
                        if has_standalone_cover and not preserve_source:
                            for r_node in list(sp.findall(qn('w:r'))):
                                t_node = r_node.find(qn('w:t'))
                                if t_node is not None and t_node.text and t_node.text.strip():
                                    break
                                br_nodes = r_node.findall(qn('w:br'))
                                if br_nodes and (t_node is None or not t_node.text or not t_node.text.strip()):
                                    sp.remove(r_node)
                    
                    # 提取纯净可见文本（包含制表符）
                    part_list = []
                    for n_el in sp.iter():
                        if n_el.tag == qn('w:t'):
                            part_list.append(n_el.text or '')
                        elif n_el.tag == qn('w:tab'):
                            part_list.append('\t')
                    vis_txt = ''.join(part_list).strip()
                    vis_clean = re.sub(r'\s+', '', vis_txt)

                    # 处理内部目录项
                    if is_pyfa_doc:
                        if (re.match(r'^(目录|目次)+$', vis_clean) or vis_txt in ['目  录', '目录', '目 录']) and not in_pyfa_toc:
                            in_pyfa_toc = True
                            # 目录作为独立新页起始：在“目  录”段落上直接设置 pageBreakBefore，避免额外空段落吃掉行高
                            p_h = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:pageBreakBefore/><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto" w:before="0" w:after="100"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="黑体" w:eastAsia="黑体" w:hAnsi="黑体"/><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>目  录</w:t></w:r></w:p>')
                            append_element_to_body(doc, p_h)
                            continue
                            
                        if in_pyfa_toc:
                            # 过滤域定义指令和空白段落
                            if not vis_txt or 'TOC \\o' in vis_txt or 'TOC \\O' in vis_txt or vis_txt.startswith('TOC'):
                                continue
                                
                            if '\t' in vis_txt or re.search(r'[\.\·\s]+\d+\s*$', vis_txt):
                                parts = vis_txt.split('\t')
                                title = parts[0].strip()
                                pg_str = parts[-1].strip() if len(parts) > 1 else ''
                                m_pg = re.search(r'(\d+)\s*$', pg_str or title)
                                pg_num = m_pg.group(1) if m_pg else '1'
                                title = re.sub(r'[\t\.\·\s]*\d+\s*$', '', title).strip()
                                
                                # 优先从 exact_pages 动态反查真实打印页码
                                target_pg = exact_pages.get(title) if exact_pages else None
                                if target_pg is None and exact_pages:
                                    t_norm = re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', title)
                                    for ek, ev in exact_pages.items():
                                        ek_norm = re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', ek)
                                        if t_norm and (t_norm in ek_norm or ek_norm in t_norm):
                                            target_pg = ev
                                            break
                                final_pg = str(target_pg) if target_pg is not None else pg_num
                                
                                lvl = 1
                                if title.startswith(('（', '(')):
                                    lvl = 2
                                elif re.match(r'^\d+\.', title):
                                    lvl = 3
                                    
                                p_toc = create_static_toc_paragraph(title, final_pg, level=lvl, fonts=fonts)
                                append_element_to_body(doc, p_toc)
                                continue
                            else:
                                # 遇到第一个正文段落，结束目录模式并在正文段落上开启新页
                                in_pyfa_toc = False
                                first_pyfa_body = True

                    cloned = copy_element_with_rels(sp, doc_src, doc)
                    _remap_source_bookmarks(
                        cloned, source_bookmark_names, source_bookmark_ids
                    )
                    _remap_source_note_references(
                        cloned, source_footnote_ids, source_endnote_ids
                    )
                    pPr_cl = cloned.find(qn('w:pPr'))
                    if pPr_cl is not None:
                        sPr = pPr_cl.find(qn('w:sectPr'))
                        if sPr is not None:
                            pPr_cl.remove(sPr)
                    if suppress_first_page_break and first_content_element and (vis_txt or drawings):
                        # The outer directory-tree heading already establishes
                        # the file boundary. Override a source Heading 1's
                        # inherited page break so heading and first body
                        # content do not create a title-only page.
                        first_ppr = cloned.get_or_add_pPr()
                        first_break = first_ppr.find(qn("w:pageBreakBefore"))
                        if first_break is None:
                            first_break = OxmlElement("w:pageBreakBefore")
                            first_ppr.append(first_break)
                        first_break.set(qn("w:val"), "0")
                        first_content_element = False
                    if not preserve_source:
                        inline_style_properties(cloned, doc_src)

                    # itertext() 会把部分 Word XML 文本重复计入；用前面逐节点提取的可见文本匹配。
                    matching = remaining_anchors.get(vis_txt)
                    if matching:
                        _add_embedded_heading_bookmark(cloned, matching.pop(0))

                    # 针对封面空白段落，精准控制行高确保封面元素完美单页舒展，绝不溢出到第2页
                    if is_pyfa_doc:
                        pPr_cur = cloned.get_or_add_pPr()
                        if not in_pyfa_toc and not first_pyfa_body:
                            if not txt and not drawings and not brs:
                                sp_el = pPr_cur.find(qn('w:spacing'))
                                if sp_el is None:
                                    sp_el = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                    pPr_cur.append(sp_el)
                                sp_el.set(qn('w:line'), '380')
                                sp_el.set(qn('w:lineRule'), 'exact')
                                sp_el.set(qn('w:before'), '0')
                                sp_el.set(qn('w:after'), '0')
                            elif txt.startswith(('主要编制人员', '编制人员')) or (txt.startswith(('一、专业名称', '一、培养目标')) and last_elem_was_table):
                                if pPr_cur.find(qn('w:pageBreakBefore')) is None:
                                    pPr_cur.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
                        elif first_pyfa_body:
                            # 确保目录结束后的首个正文标题开启新页
                            if pPr_cur.find(qn('w:pageBreakBefore')) is None:
                                pPr_cur.append(parse_xml(f'<w:pageBreakBefore {nsdecls("w")}/>'))
                            first_pyfa_body = False
                    
                    # 清理从源文档继承的孤立 TOC 域符号（防止 Word 报文档损坏）
                    for f_char in cloned.findall('.//' + qn('w:fldChar')):
                        f_type = f_char.get(qn('w:fldCharType')) or f_char.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType')
                        if f_type == 'end':
                            p_parent = f_char.getparent()
                            while p_parent is not None and p_parent.tag != qn('w:p'):
                                p_parent = p_parent.getparent()
                            if p_parent is not None:
                                begins = [fc for fc in p_parent.findall('.//' + qn('w:fldChar')) if (fc.get(qn('w:fldCharType')) == 'begin' or fc.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType') == 'begin')]
                                if not begins:
                                    r_parent = f_char.getparent()
                                    if r_parent is not None:
                                        r_parent.remove(f_char)
                    
                    pPr = cloned.find(qn('w:pPr'))
                    if pPr is not None:
                        if txt and (re.match(r'^[一二三四五六七八九十]+[、\.]', txt) or re.match(r'^（[一二三四五六七八九十]+）', txt)):
                            if pPr.find(qn('w:keepNext')) is None:
                                pPr.append(parse_xml(f'<w:keepNext {nsdecls("w")}/>'))

                        if is_wen_doc:
                            if '六、预期成果' in txt or '六、 预期成果' in txt:
                                in_wen_expected = True
                            if in_wen_expected:
                                sp_el = pPr.find(qn('w:spacing'))
                                if sp_el is None:
                                    sp_el = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                    pPr.append(sp_el)
                                sp_el.set(qn('w:before'), '0')
                                sp_el.set(qn('w:after'), '0')
                                sp_el.set(qn('w:line'), '260')
                                sp_el.set(qn('w:lineRule'), 'auto')
                        elif is_huaxiang_doc:
                            if '附录' in txt:
                                in_huaxiang_appendix = True
                            if in_huaxiang_appendix:
                                sp_el = pPr.find(qn('w:spacing'))
                                if sp_el is None:
                                    sp_el = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                    pPr.append(sp_el)
                                sp_el.set(qn('w:line'), '240')
                                sp_el.set(qn('w:lineRule'), 'auto')
                        elif is_keti_doc or is_xize_doc:
                            sp_el = pPr.find(qn('w:spacing'))
                            if sp_el is not None:
                                sp_el.set(qn('w:line'), '530')
                                sp_el.set(qn('w:lineRule'), 'exact')
                        elif is_growth_doc:
                            sp_el = pPr.find(qn('w:spacing'))
                            if sp_el is None:
                                sp_el = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                pPr.append(sp_el)
                            sp_el.set(qn('w:line'), '235')
                            sp_el.set(qn('w:lineRule'), 'auto')
                            
                    if not preserve_source:
                        sanitize_ole_and_external_links(cloned)
                    append_element_to_body(doc, cloned)
                    if txt:
                        last_elem_was_table = False

            elif tag in ('tbl', 'sdt'):
                last_elem_was_table = True
                cloned = copy_element_with_rels(elem, doc_src, doc)
                _remap_source_bookmarks(
                    cloned, source_bookmark_names, source_bookmark_ids
                )
                _remap_source_note_references(
                    cloned, source_footnote_ids, source_endnote_ids
                )
                if suppress_first_page_break and first_content_element:
                    first_content_element = False
                if not preserve_source:
                    inline_style_properties(cloned, doc_src)
                    if render_context is not None:
                        adapt_table_to_content_box(cloned, render_context.get_content_box())
                if tag == 'tbl' and is_pyfa_doc:
                    t_txt = ''.join(cloned.itertext())
                    if '版本号' in t_txt and '单位名称' in t_txt:
                        # 精确匹配封面表头列宽：Logo 1100 dxa (完整不裁切), 校名 5100 dxa (确保单行不折行), 间隔 300 dxa, 版本号 2300 dxa (单行垂直居中)
                        widths = ['1100', '5100', '300', '2300']
                        tblPr = cloned.find(qn('w:tblPr'))
                        if tblPr is not None:
                            tblLayout = tblPr.find(qn('w:tblLayout'))
                            if tblLayout is None:
                                tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
                            else:
                                tblLayout.set(qn('w:type'), 'fixed')
                            tblW = tblPr.find(qn('w:tblW'))
                            if tblW is not None:
                                tblW.set(qn('w:w'), '8800')
                                tblW.set(qn('w:type'), 'dxa')
                        tblGrid = cloned.find(qn('w:tblGrid'))
                        if tblGrid is not None:
                            for idx, gc in enumerate(tblGrid.findall(qn('w:gridCol'))):
                                if idx < len(widths):
                                    gc.set(qn('w:w'), widths[idx])
                        for tr in cloned.findall(qn('w:tr')):
                            for idx, tc in enumerate(tr.findall(qn('w:tc'))):
                                tcPr = tc.get_or_add_tcPr()
                                tcW = tcPr.find(qn('w:tcW'))
                                if tcW is not None and idx < len(widths):
                                    tcW.set(qn('w:w'), widths[idx])
                                if idx == 0:
                                    # 校徽单元格：必须重置为 auto 行距与 0 缩进，防止继承正文样式导致裁剪顶部与右移
                                    for p in tc.findall(qn('w:p')):
                                        pPr = p.get_or_add_pPr()
                                        sp = pPr.find(qn('w:spacing'))
                                        if sp is None:
                                            sp = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                            pPr.append(sp)
                                        sp.set(qn('w:line'), '240')
                                        sp.set(qn('w:lineRule'), 'auto')
                                        sp.set(qn('w:before'), '0')
                                        sp.set(qn('w:after'), '0')
                                        ind = pPr.find(qn('w:ind'))
                                        if ind is None:
                                            ind = parse_xml(f'<w:ind {nsdecls("w")}/>')
                                            pPr.append(ind)
                                        ind.set(qn('w:firstLine'), '0')
                                        ind.set(qn('w:firstLineChars'), '0')
                                        ind.set(qn('w:left'), '0')
                                elif idx == 3:
                                    vAlign = tcPr.find(qn('w:vAlign'))
                                    if vAlign is None:
                                        tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                                    else:
                                        vAlign.set(qn('w:val'), 'center')
                                    for p in tc.findall(qn('w:p')):
                                        pPr = p.get_or_add_pPr()
                                        sp = pPr.find(qn('w:spacing'))
                                        if sp is None:
                                            sp = parse_xml(f'<w:spacing {nsdecls("w")}/>')
                                            pPr.append(sp)
                                        sp.set(qn('w:line'), '240')
                                        sp.set(qn('w:lineRule'), 'auto')
                                        sp.set(qn('w:before'), '0')
                                        sp.set(qn('w:after'), '0')
                                        jc = pPr.find(qn('w:jc'))
                                        if jc is None:
                                            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
                                        else:
                                            jc.set(qn('w:val'), 'center')
                                        # 显式指定五号字(21半磅)，确保单行水平垂直居中容纳
                                        for r in p.findall('.//' + qn('w:r')):
                                            rPr = r.get_or_add_rPr()
                                            sz = rPr.find(qn('w:sz'))
                                            if sz is None:
                                                rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="21"/>'))
                                            else:
                                                sz.set(qn('w:val'), '21')
                        # 校名单元格设为单倍行距，避免折行
                        tr0 = cloned.findall(qn('w:tr'))[0] if cloned.findall(qn('w:tr')) else None
                        if tr0 is not None and len(tr0.findall(qn('w:tc'))) > 1:
                            for p in tr0.findall(qn('w:tc'))[1].findall(qn('w:p')):
                                pPr = p.get_or_add_pPr()
                                sp = pPr.find(qn('w:spacing'))
                                if sp is not None:
                                    sp.set(qn('w:line'), '240')
                                    sp.set(qn('w:lineRule'), 'auto')

                # 严密防御：OpenXML 规定表格所有单元格 (w:tc) 必须包含至少一个块级元素 (w:p)
                for tc_el in cloned.findall('.//' + qn('w:tc')):
                    if len(tc_el.findall(qn('w:p'))) == 0:
                        tc_el.append(parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:before="0" w:after="0"/></w:pPr></w:p>'))
                if not preserve_source:
                    sanitize_ole_and_external_links(cloned)
                append_element_to_body(doc, cloned)

    missing = [item["title"] for matches in remaining_anchors.values() for item in matches]
    if missing:
        raise RuntimeError(f"内嵌目录标题未能写入书签: {', '.join(missing)}")


def render_image_file(
    doc: Document,
    img_path: Path,
    max_w_cm: Optional[float] = None,
    max_h_cm: Optional[float] = None,
    render_context: Optional[RenderContext] = None,
):
    """图片居中插入，限制最大宽高"""
    if render_context is not None:
        box = render_context.get_content_box()
        max_w_cm = max_w_cm if max_w_cm is not None else box.content_width_cm
        max_h_cm = max_h_cm if max_h_cm is not None else box.content_height_cm
    else:
        max_w_cm = max_w_cm if max_w_cm is not None else 15.0
        max_h_cm = max_h_cm if max_h_cm is not None else 18.2
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    
    with Image.open(str(img_path)) as im:
        w_px, h_px = im.size
        ratio = w_px / h_px
        
    calc_h = max_w_cm / ratio
    if calc_h > max_h_cm:
        p.add_run().add_picture(str(img_path), height=Cm(max_h_cm))
    else:
        p.add_run().add_picture(str(img_path), width=Cm(max_w_cm))


def is_standalone_cover_doc(docx_path: Path) -> bool:
    """判断文档是否带有独立封面页（如表格表头、居中大标题等），避免在封面顶部强插冗余大纲文字"""
    try:
        d = Document(str(docx_path))
        if d.tables:
            t0_txt = ''.join(d.tables[0]._element.itertext()).strip()
            if '版本号' in t0_txt:
                return True
        empty_count = 0
        for p in d.paragraphs[:12]:
            txt = p.text.strip()
            if not txt:
                empty_count += 1
            if ('人才培养方案' in txt or '实施方案' in txt) and empty_count >= 3:
                return True
        return False
    except Exception:
        return False


def add_invisible_heading_anchor(doc: Document, node: Dict[str, Any], need_page_break: bool = False):
    """
    针对带有独立封面的子文档，插入零高度书签锚点并处理分页，
    既能使整本大书的目录精准链接与统计页码，又不会在封面顶端破坏原有版式。
    """
    bm_name = node.get("bookmark_name")
    if not bm_name:
        import hashlib
        bm_name = f"_Toc_{hashlib.md5(node.get('title', '').encode('utf-8')).hexdigest()[:8]}"
        node["bookmark_name"] = bm_name
        
    p = doc.add_paragraph()
    if need_page_break:
        p.paragraph_format.page_break_before = True
    pPr = p._element.get_or_add_pPr()
    sp = parse_xml(f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="0" w:lineRule="exact"/>')
    pPr.append(sp)
    
    r_bm = parse_xml(f'''
        <w:r {nsdecls("w")}>
            <w:bookmarkStart w:id="0" w:name="{bm_name}"/>
            <w:bookmarkEnd w:id="0"/>
        </w:r>
    ''')
    p._element.append(r_bm)


def render_tree_node(
    doc: Document,
    node: Dict[str, Any],
    source_root: Path,
    is_first_section: bool,
    fonts: Dict[str, str],
    is_first_child: bool = False,
    exact_pages: Dict[str, int] = None,
    render_context: Optional[RenderContext] = None,
):
    """
    递归渲染大纲节点
    使用段落属性 page_break_before 控制换节分页；当节点为首个子节点时，避免分页以便与父标题同页。
    """
    ntype = node.get("type")
    fpath = source_root / node.get("file", "") if node.get("file") else None
    has_cover = (ntype == "docx" and fpath and fpath.exists() and is_standalone_cover_doc(fpath))
    
    need_pb = False
    last_landscape = render_context.last_rendered_landscape if render_context is not None else LAST_RENDERED_LANDSCAPE[0]
    if not is_first_section and not is_first_child and not last_landscape:
        need_pb = True
    if render_context is not None:
        render_context.last_rendered_landscape = False
    else:
        LAST_RENDERED_LANDSCAPE[0] = False
        
    if has_cover:
        add_invisible_heading_anchor(doc, node, need_page_break=need_pb)
    else:
        if render_context is not None and render_context.resolved_format:
            from .style_applier import add_styled_heading
            add_styled_heading(doc, node.get("title", ""), int(node.get("level") or 1), render_context.resolved_format,
                               bookmark_name=node.get("bookmark_name"), bookmark_id=node.get("bm_id"), need_page_break=need_pb)
        else:
            add_heading_paragraph(doc, node, fonts, need_page_break=need_pb)
    
    if ntype == "pdf":
        fpath = source_root / node["file"]
        render_pdf_file(doc, fpath, has_headings_on_page=True, render_context=render_context)
    elif ntype == "docx":
        fpath = source_root / node["file"]
        render_docx_file(
            doc,
            fpath,
            exact_pages=exact_pages,
            fonts=fonts,
            render_context=render_context,
            suppress_first_page_break=True,
        )
    elif ntype == "image":
        fpath = source_root / node["file"]
        render_image_file(doc, fpath, render_context=render_context)
    elif ntype == "folder":
        children = node.get("children", [])
        for idx, child in enumerate(children):
            render_tree_node(doc, child, source_root, is_first_section=False, fonts=fonts, is_first_child=(idx == 0), exact_pages=exact_pages, render_context=render_context)
