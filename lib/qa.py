# -*- coding: utf-8 -*-
"""
质量检测与排版质检模块
包含调用 Word 导出 PDF、反查打印页码与版面断言
"""
import os
from pathlib import Path
import shutil
import sys
import subprocess
import tempfile
import uuid
import pypdf
import pymupdf
import re
from typing import Dict, Any, List, Tuple


from .office import get_backend, OfficeBackendError


class OfficeExportError(RuntimeError):
    """无法使用本地 Word 取得精确分页时抛出的错误。"""


def _word_access_directory() -> Path:
    """Return one stable directory for all files touched by Word automation."""
    backend = get_backend()
    if hasattr(backend, "get_word_access_directory"):
        return backend.get_word_access_directory()
    configured = os.environ.get("DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR")
    if configured:
        directory = Path(configured).expanduser()
    else:
        directory = Path(tempfile.gettempdir()) / "document-synthesis-word-access"
    directory.mkdir(parents=True, exist_ok=True)
    return directory.resolve()


def word_export_status() -> Tuple[bool, str]:
    """检查精确分页的静态前置条件，不触发系统授权弹窗。"""
    stat = get_backend().static_status()
    return stat.available, stat.reason


def _automation_failure_message(detail: str) -> str:
    """把底层错误转成可执行的排障提示。"""
    from .office.mac_applescript import _automation_failure_message as _afm
    return _afm(detail)


def word_automation_status() -> Tuple[bool, str]:
    """真实验证当前调用方能否控制 Word；不读取或修改任何文档。"""
    stat = get_backend().probe()
    return stat.available, stat.reason


def _applescript_string(value: str) -> str:
    """将 POSIX 路径安全嵌入 AppleScript 字符串。"""
    return value.replace("\\", "\\\\").replace('"', '\\"')


def export_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    """调用本次打开的 Microsoft Word 文档导出 PDF，不影响其他已打开文档。"""
    backend = get_backend()
    stat = backend.static_status()
    if not stat.available:
        print(f"[错误] {stat.reason}")
        return False

    try:
        backend.export_pdf(Path(docx_path), Path(pdf_path))
        return os.path.exists(pdf_path)
    except OfficeBackendError as exc:
        print(f"[错误] {exc}")
        return False
    except OSError as exc:
        print(f"[错误] 无法保存 Word 导出的 PDF: {exc}")
        return False


def get_exact_printed_heading_pages(docx_path: str, toc_items: List[Dict[str, Any]], source_dir: Any = None) -> Dict[str, int]:
    """通过本地 Word 导出 PDF 反查实际打印页码表（已废弃）。"""
    import warnings
    warnings.warn(
        "get_exact_printed_heading_pages 已弃用，将在后续版本中删除。请改用书签报告或 inspect_document。",
        DeprecationWarning,
        stacklevel=2,
    )
    pdf_tmp = docx_path.replace(".docx", "_tmp_pageref.pdf")
    if not export_docx_to_pdf(docx_path, pdf_tmp):
        raise OfficeExportError(
            "无法取得 Microsoft Word 的精确物理页码；为避免生成错误目录，已停止构建。"
        )
        
    reader = pypdf.PdfReader(pdf_tmp)
    total_pages = len(reader.pages)
    
    # 寻找正文起始页（跳过封面与目录）
    body_start_idx = 0
    first_title = (toc_items[0].get("title") or toc_items[0].get("text")) if toc_items else ""
    first_clean = re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', first_title[:10])
    
    for idx in range(1, total_pages):
        txt = reader.pages[idx].extract_text() or ""
        clean_txt = re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', txt)
        if first_clean in clean_txt and "目  录" not in txt and "目录" not in txt:
            body_start_idx = idx
            break
            
    def norm(t): 
        return re.sub(r'[\s\.\、\，\。\（\）\(\)\|｜·\/\-\_\:\：\—\t\n\r]', '', t)
        
    from pathlib import Path
    from docx import Document
    s_root = Path(source_dir) if source_dir else None
    
    # 第一遍：定位所有顶层/主大纲节点的起始打印物理页码
    top_pages = []
    cur_page_idx = body_start_idx
    for i_top, item in enumerate(toc_items):
        if i_top == 0:
            top_pages.append((item, body_start_idx))
            cur_page_idx = body_start_idx
            continue
            
        t = item.get("title") or item.get("text", "")
        dclean = norm(t[:12])
        
        found_page = None
        for p_idx in range(cur_page_idx, total_pages):
            page = reader.pages[p_idx]
            text = page.extract_text() or ""
            clean_t = norm(text)
            
            if dclean in clean_t:
                found_page = p_idx
                break
                
        if found_page is None:
            found_page = cur_page_idx
            
        top_pages.append((item, found_page))
        cur_page_idx = found_page
        
    exact_pages = {}
    
    # 第二遍：按各文档的作用域界限（[start_idx, next_start_idx]）检索内部子目录各项
    for i, (item, start_idx) in enumerate(top_pages):
        t = item.get("title") or item.get("text", "")
        exact_pages[t] = start_idx - body_start_idx + 1
        next_start_idx = top_pages[i + 1][1] if i + 1 < len(top_pages) else total_pages
        
        fpath = None
        if item.get("type") == "docx" and "file" in item:
            fpath = s_root / item["file"] if s_root else Path(item["file"])
        elif item.get("path"):
            fpath = Path(item["path"])
            
        if fpath and fpath.exists():
            try:
                doc_src = Document(str(fpath))
                in_sub_toc = False
                sub_toc_titles = []
                for p in doc_src.paragraphs:
                    vt = p.text.strip()
                    cns = re.sub(r'[\s\t\r\n]', '', vt)
                    if cns in ['目录', '目次'] or (cns.startswith('目') and cns.endswith('录') and len(cns) <= 4):
                        in_sub_toc = True
                        continue
                    if in_sub_toc:
                        if ('\t' in vt or bool(re.search(r'[\.\·]{3,}\s*\d+', vt)) or bool(re.search(r'\d+\s*$', vt))) and vt:
                            m = re.search(r'(\d+)\s*$', vt)
                            if m:
                                sub_title = vt[:m.start()].strip()
                                sub_title = re.sub(r'[\.\·\s\t]+$', '', sub_title).strip()
                            else:
                                sub_title = vt.strip()
                            if sub_title:
                                sub_toc_titles.append(sub_title)
                        elif vt:
                            break
                            
                if sub_toc_titles:
                    sub_cur_pno = start_idx
                    for stitle in sub_toc_titles:
                        st_clean = norm(stitle)
                        st_pure = norm(re.sub(r'^[一二三四五六七八九十\d\.\、\s\（\）\(\)]+', '', stitle))
                        
                        st_found = None
                        for sp_idx in range(sub_cur_pno, next_start_idx):
                            clean_sp = norm(reader.pages[sp_idx].extract_text() or "")
                            if (len(st_clean) >= 3 and st_clean in clean_sp) or (len(st_pure) >= 3 and st_pure in clean_sp) or (len(st_pure) >= 2 and st_pure[:4] in clean_sp):
                                st_found = sp_idx
                                sub_cur_pno = sp_idx
                                break
                        if st_found is not None:
                            exact_pages[stitle] = st_found - body_start_idx + 1
                        else:
                            exact_pages[stitle] = sub_cur_pno - body_start_idx + 1
            except Exception:
                pass
        
    if os.path.exists(pdf_tmp):
        try:
            os.remove(pdf_tmp)
        except Exception:
            pass
            
    return exact_pages


def run_qa_assertions(pdf_path: str, ignore_front_pages: int = 4, page_roles=None, allowed_blank_pages=None) -> bool:
    """逐页扫描断言，检查是否存在空白页、单标题页与孤行"""
    if not os.path.exists(pdf_path):
        return False
        
    pdoc = pymupdf.open(pdf_path)
    total_pages = len(pdoc)
    
    blank_pages = []
    heading_only_pages = []
    orphan_tails = []
    
    def has_visible_page_content(page: pymupdf.Page) -> bool:
        """识别无文本的扫描件/凭证页，避免把有效位图误判为空白。"""
        # 排除页脚区域，避免仅凭页码把真正空白页判为有内容。
        rect = page.rect
        content_rect = pymupdf.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.9)
        pix = page.get_pixmap(
            matrix=pymupdf.Matrix(0.3, 0.3),
            colorspace=pymupdf.csRGB,
            alpha=False,
            clip=content_rect,
        )
        samples = pix.samples
        dark_samples = sum(1 for value in samples[::3] if value < 245)
        return dark_samples > max(12, (pix.width * pix.height) // 10000)

    def is_header_footer_page_num(line_str: str) -> bool:
        clean = line_str.strip("- ").strip()
        if not clean:
            return False
        if clean.isdigit():
            return True
        from .pagination_types import roman_to_int
        if roman_to_int(clean) is not None:
            return True
        return False

    for i, page in enumerate(pdoc):
        txt = page.get_text().strip()
        raw_lines = [l.strip() for l in txt.split('\n') if l.strip()]
        content_lines = [
            l for l in raw_lines 
            if not is_header_footer_page_num(l)
        ]
        images = page.get_images()
        drawings = page.get_drawings()
        has_graphical_content = bool(images or drawings) or (not content_lines and has_visible_page_content(page))
        
        if len(content_lines) == 0 and not has_graphical_content:
            if i >= ignore_front_pages and (i + 1) not in (allowed_blank_pages or set()):
                blank_pages.append(i + 1)
        elif total_pages > 1 and len(content_lines) <= 2 and not has_graphical_content and (
            page_roles is None or page_roles.get(i + 1) == "body"
        ):
            # A complete one-page document can legitimately contain only a
            # short paragraph.  "Orphan tail" is a page-boundary defect, so
            # it is meaningful only when the artifact has another page to
            # establish that boundary; blank-page checks above still apply to
            # one-page artifacts.
            if i >= ignore_front_pages:
                if any("目录" in l or "目  录" in l or "....." in l or "……" in l for l in content_lines):
                    continue
                # 忽略公文正规的落款区（签名、盖章、日期、附件与表格清单）及子文件大标题
                if any(any(k in l for k in ["签名", "公章", "盖章", "落款", "附件：", "附件1", "附件2", "考核细则", "申请表", "登记表", "汇总表", "统计表", "申报表", "日常工作制度", "管理细则", "管理办法", "建设制度", "评价办法"]) for l in content_lines):
                    continue
                first_line = content_lines[0] if content_lines else ""
                if any(first_line.startswith(pfx) for pfx in ["一、", "二、", "三、", "四、", "五、", "六、", "七、", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "（1）", "（2）"]):
                    heading_only_pages.append((i + 1, content_lines))
                else:
                    orphan_tails.append((i + 1, content_lines))
                    
    pdoc.close()
    
    passed = (len(blank_pages) == 0 and len(heading_only_pages) == 0 and len(orphan_tails) == 0)
    if passed:
        print(f"  [通过] 逐页断言全部通过（共 {total_pages} 页），无空白页、单标题页或孤行。")
    else:
        if blank_pages:
            print(f"  [未通过] 发现空白页: 第 {blank_pages} 页")
        if heading_only_pages:
            print(f"  [未通过] 发现单标题页: {heading_only_pages}")
        if orphan_tails:
            print(f"  [未通过] 发现孤行: {orphan_tails}")
            
    return passed
