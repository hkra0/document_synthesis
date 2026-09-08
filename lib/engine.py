# -*- coding: utf-8 -*-
"""
统一来源调度、正文复用、声明式装配与验证后发布。
"""

import os
import sys
import io
import json
import re
import hashlib
import subprocess
import uuid
import shutil
from pathlib import Path
from typing import List, Dict, Any, Optional, Mapping

import pymupdf
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml.ns import qn

from .config import ConfigError, load_project_config
from .scanner import flatten_tree_nodes
from .source_strategies import SourceStrategyError, build_outline
from .build_plan import PreparedBuild, RenderedBody, prepare_project_build
from .layout import RenderContext, apply_section_spec
from .styles import (
    apply_standard_page_setup, setup_footer, add_heading_paragraph,
    neutralize_document_styles
)
from .renderers import (
    render_docx_file, render_pdf_file, get_cached_pdf_page_image,
    LAST_RENDERED_LANDSCAPE,
    is_standalone_cover_doc, add_invisible_heading_anchor, set_pdf_render_cache_dir,
    _copy_section_stories,
)
from .qa import OfficeExportError


class BuildError(RuntimeError):
    """构建无法安全完成时抛出的错误。"""


def _source_docx_paths(source_path: Path, config) -> List[Path]:
    """返回本次内容门禁可追溯的 DOCX 源节点。"""
    if source_path.is_file():
        return [source_path] if source_path.suffix.lower() == ".docx" else []

    source_cfg = getattr(config, "source", {}) or {}
    strategy = source_cfg.get("strategy")
    if strategy == "docx_document":
        file_name = source_cfg.get("file") or source_cfg.get("path")
        if file_name:
            candidate = (source_path / file_name).resolve()
            return [candidate] if candidate.is_file() and candidate.suffix.lower() == ".docx" else []
        candidates = sorted(
            p.resolve() for p in source_path.glob("*.docx")
            if not p.name.startswith(("~$", "."))
        )
        return candidates if len(candidates) == 1 else []

    return sorted(
        p.resolve() for p in source_path.rglob("*.docx")
        if not p.name.startswith(("~$", "."))
    )


def _write_failure_diagnostics(output_dir: Path, error: BaseException, qa_summary=None) -> None:
    """把失败原因写到独立文件，绝不触碰既有交付物或成功元数据。"""
    payload = {
        "status": "failed",
        "error": {
            "type": type(error).__name__,
            "message": str(error),
        },
        "qa": qa_summary or {},
    }
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        diagnostic_path = output_dir / "build-diagnostics.json"
        if diagnostic_path.is_symlink() or (diagnostic_path.exists() and not diagnostic_path.is_file()):
            return
        diagnostic_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
    except Exception:
        # 诊断文件属于辅助信息，不能覆盖原始构建错误。
        pass


def _apply_target_page_policy(document: Document, config) -> None:
    """在 restyle/target 模式下把格式包页面几何落到最终正文。"""
    if getattr(config, "schema_version", 1) < 3:
        return
    formatting = getattr(config, "formatting", None)
    mode = getattr(formatting, "mode", "restyle")
    page_policy = getattr(formatting, "page_policy", "target")
    if mode == "restyle" or (mode == "mixed" and page_policy == "target"):
        for section in document.sections:
            apply_section_spec(section, config.resolved_format)


def _add_heading_with_context(document, node, fonts, *, need_page_break=False, context=None):
    """R4 统一生成标题入口；旧版本调用仍保留在 legacy renderer 中。"""
    if context is not None and context.resolved_format is not None:
        from .style_applier import add_styled_heading

        return add_styled_heading(
            document,
            node.get("title", ""),
            level=int(node.get("level") or 1),
            resolved_format=context.resolved_format,
            bookmark_name=node.get("bookmark_name"),
            bookmark_id=node.get("bm_id"),
            need_page_break=need_page_break,
        )
    return add_heading_paragraph(document, node, fonts, need_page_break=need_page_break)


def _start_preserved_source_boundary(document, source_path: Optional[Path], context) -> None:
    """Start a new source section before writing the source wrapper heading."""
    if source_path is None or context is None:
        return
    policy = getattr(context, "formatting_policy", None)
    preserves_geometry = bool(
        policy is not None
        and (
            getattr(policy, "mode", "restyle") == "preserve"
            or (
                getattr(policy, "mode", "restyle") == "mixed"
                and getattr(policy, "page_policy", "target") == "source"
            )
        )
    )
    if not preserves_geometry or not getattr(document, "_synth_source_rendered", False):
        return
    source_document = Document(str(source_path))
    if not source_document.sections:
        return
    section = document.add_section(WD_SECTION_START.NEW_PAGE)
    apply_section_spec(section, source_section=source_document.sections[0])
    _copy_section_stories(source_document.sections[0], section)


def _policy_for_source_file(policy, source_file: Path):
    """把 mixed 的文件级 scope 收敛为本次源文件实际使用的模式。"""
    if policy is None or policy.mode != "mixed":
        return policy
    # mixed 的样式模式不能被压扁为单一文件模式；NodeRef scope 由下面的
    # role-map 过滤器处理。此函数只保留统一的策略对象，供页面策略判断。
    return policy


def _node_path_order(path: str):
    return tuple(int(value) for value in re.findall(r"\[(\d+)\]", path or ""))


def _node_in_scope(path: str, node_range: Optional[str]) -> bool:
    if not node_range or node_range in ("*", "all"):
        return True
    if path == node_range:
        return True
    range_parts = re.split(r"\s*\.\.\s*|\s+-\s+", node_range)
    if len(range_parts) == 2 and all("[" in part for part in range_parts):
        start_order = _node_path_order(range_parts[0])
        end_order = _node_path_order(range_parts[1])
        path_order = _node_path_order(path)
        if start_order and end_order and path_order:
            return start_order <= path_order <= end_order
    range_numbers = [int(value) for value in re.findall(r"\[(\d+)\]", node_range)]
    path_order = _node_path_order(path)
    if len(range_numbers) >= 2 and path_order:
        start, end = range_numbers[0], range_numbers[-1]
        # 兼容 p[2]..p[8]、p[2]:p[8] 等声明；端点均为包含范围。
        if len(path_order) == 1:
            return start <= path_order[0] <= end
    return False


def _scoped_role_map(assignments, policy, source_file: Path) -> Dict[str, str]:
    """按文件与 NodeRef 范围筛选 mixed 模式真正托管的角色。"""
    if policy is None or policy.mode == "preserve":
        return {}
    if policy.mode != "mixed":
        return {item.node_ref.element_path: item.role for item in assignments}
    source_file = Path(source_file).resolve()
    matching_scopes = []
    for scope in policy.scopes or []:
        declared = scope.get("file")
        if not declared:
            continue
        candidate = Path(declared)
        try:
            same_file = candidate.resolve() == source_file
        except OSError:
            same_file = candidate.name == source_file.name
        if same_file:
            matching_scopes.append(scope)
    if not matching_scopes:
        return {item.node_ref.element_path: item.role for item in assignments}
    result = {}
    for item in assignments:
        applicable = [
            scope for scope in matching_scopes
            if _node_in_scope(item.node_ref.element_path, scope.get("node_range"))
        ]
        if applicable and applicable[-1].get("mode") == "restyle":
            result[item.node_ref.element_path] = item.role
    return result


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
    resolved_files: Optional[Mapping[str, Path]] = None,
    context: Optional[RenderContext] = None,
):
    """递归排版大纲树节点"""
    ntype = node.get("type", "folder")
    resolved_files = resolved_files or {}
    if context is not None:
        context.resolved_files = resolved_files
        context.current_section = doc.sections[-1] if doc.sections else None
    
    # 确定是否需要强制换页（首大节、首子节点、以及刚从横版恢复的分节无需换页）
    need_page_break = False
    last_landscape = context.last_rendered_landscape if context is not None else LAST_RENDERED_LANDSCAPE[0]
    if not is_first_section and not is_first_child and not last_landscape:
        need_page_break = True
    if context is not None:
        context.last_rendered_landscape = False
    else:
        LAST_RENDERED_LANDSCAPE[0] = False

    if ntype == "folder":
        _add_heading_with_context(doc, node, fonts, need_page_break=need_page_break, context=context)
        children = node.get("children", [])
        for idx, child in enumerate(children):
            render_tree_node_recursive(
                doc, child, source_dir, fonts,
                is_first_child=(idx == 0),
                exact_pages=exact_pages,
                resolved_files=resolved_files,
                context=context,
            )
            
    elif ntype == "docx_outline":
        rel_file = node.get("file")
        fpath = _resolve_node_path(node, source_dir, resolved_files)
        _start_preserved_source_boundary(doc, fpath, context)
        if node.get("add_heading_before_content", True):
            _add_heading_with_context(doc, node, fonts, need_page_break=need_page_break, context=context)
        elif need_page_break:
            # 根节点不显示标题时，仍让其正文在正确的新页开始。
            anchor = doc.add_paragraph()
            anchor.paragraph_format.page_break_before = True
        if fpath and fpath.exists():
            preceding_elements = set(doc.element.body)
            render_docx_file(
                doc, fpath, exact_pages=exact_pages, fonts=fonts,
                embedded_outline=node.get("children", []),
                render_context=context,
                suppress_first_page_break=True,
            )
            doc._synth_source_rendered = True
            if not node.get("add_heading_before_content", True):
                from .composition import add_bookmark
                first_paragraph = next((p for element in doc.element.body if element not in preceding_elements
                                        for p in element.iter(qn("w:p"))
                                        if list(p.iter(qn("w:t"))) or list(p.iter(qn("w:drawing")))), None)
                if first_paragraph is None:
                    raise BuildError(f"无法定位隐藏父标题的正文锚点: {node['title']}")
                add_bookmark(first_paragraph, node["bookmark_name"], node["bm_id"])
        elif rel_file:
            raise BuildError(f"找不到 Word 文件: {fpath}")

    elif ntype == "docx":
        rel_file = node.get("file")
        fpath = _resolve_node_path(node, source_dir, resolved_files)
        _start_preserved_source_boundary(doc, fpath, context)
                
        has_cover = bool(fpath and fpath.exists() and is_standalone_cover_doc(fpath))
        if has_cover:
            add_invisible_heading_anchor(doc, node, need_page_break=need_page_break)
        else:
            _add_heading_with_context(doc, node, fonts, need_page_break=need_page_break, context=context)
            
        if fpath and fpath.exists():
            render_docx_file(
                doc,
                fpath,
                exact_pages=exact_pages,
                fonts=fonts,
                render_context=context,
                suppress_first_page_break=True,
            )
            doc._synth_source_rendered = True
        elif rel_file:
            raise BuildError(f"找不到 Word 文件: {fpath}")
                
    elif ntype == "pdf":
        _add_heading_with_context(doc, node, fonts, need_page_break=need_page_break, context=context)
        rel_file = node.get("file")
        if rel_file:
            fpath = _resolve_node_path(node, source_dir, resolved_files)
            if fpath and fpath.exists():
                render_pdf_file(doc, fpath, has_headings_on_page=True, dpi=node.get("pdf_dpi", 300), render_context=context)
            else:
                raise BuildError(f"找不到 PDF 文件: {fpath}")
                
    elif ntype == "pptx":
        rel_file = node.get("file")
        if rel_file:
            pptx_path = source_dir / rel_file
            pdf_path = resolved_files.get(rel_file, pptx_path.with_suffix(".pdf"))
            if pdf_path.exists():
                sec_pptx = doc.add_section(WD_SECTION_START.NEW_PAGE)
                from dataclasses import replace
                from .format_schema import PageSpec
                base_page = context.page_spec if context and context.page_spec else PageSpec()
                landscape_page = replace(
                    base_page,
                    width_mm=max(base_page.width_mm, base_page.height_mm),
                    height_mm=min(base_page.width_mm, base_page.height_mm),
                    orientation="landscape",
                )
                apply_section_spec(sec_pptx, landscape_page)
                pgSz = sec_pptx._sectPr.find(qn('w:pgSz'))
                if pgSz is not None:
                    pgSz.set(qn('w:w'), '16838')
                    pgSz.set(qn('w:h'), '11906')
                    pgSz.set(qn('w:orient'), 'landscape')
                if context and context.resolved_format:
                    from .style_applier import setup_styled_footer
                    setup_styled_footer(sec_pptx, context.resolved_format, start_page=None)
                else:
                    setup_footer(sec_pptx, start_page=None)
                
                _add_heading_with_context(doc, node, fonts, need_page_break=False, context=context)
                
                pdf_doc = pymupdf.open(str(pdf_path))
                for i in range(len(pdf_doc)):
                    img_bytes = get_cached_pdf_page_image(
                        pdf_path,
                        i,
                        dpi=300,
                        format_hash=getattr(context.resolved_format, "content_hash", "") if context else "",
                        page_geometry=(
                            f"{context.get_content_box().page_width_twip}x{context.get_content_box().page_height_twip}"
                            if context else ""
                        ),
                        formatting_strategy=getattr(context.formatting_policy, "mode", "") if context else "",
                    )
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
                if context is not None:
                    context.last_rendered_landscape = True
                else:
                    LAST_RENDERED_LANDSCAPE[0] = True
            else:
                raise BuildError(f"找不到 PPTX 对应的 PDF: {pdf_path}")
                
    elif ntype == "image":
        _add_heading_with_context(doc, node, fonts, need_page_break=need_page_break, context=context)
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
                content_box = context.get_content_box() if context is not None else None
                max_avail_h_cm = content_box.content_height_cm if content_box is not None else 17.5
                default_w_cm = content_box.content_width_cm if content_box is not None else 14.8
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


def compute_file_sha256(path: Path) -> str:
    """计算单个文件的 SHA-256 哈希值"""
    if not path.is_file():
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def _find_paragraph_element(document: Document, element_path: str):
    """按 inspector 的规范路径定位段落，覆盖正文段落与表格内段落。"""
    top_match = re.fullmatch(r"/w:document/w:body/w:p\[(\d+)\]", element_path)
    body = document.element.body
    if top_match:
        paragraphs = [child for child in body if child.tag == qn("w:p")]
        index = int(top_match.group(1)) - 1
        return paragraphs[index] if 0 <= index < len(paragraphs) else None

    table_match = re.fullmatch(
        r"/w:document/w:body/w:tbl\[(\d+)\]/w:tr\[(\d+)\]/w:tc\[(\d+)\]/w:p\[(\d+)\]",
        element_path,
    )
    if not table_match:
        return None
    table_index, row_index, cell_index, paragraph_index = (
        int(value) - 1 for value in table_match.groups()
    )
    tables = [child for child in body if child.tag == qn("w:tbl")]
    if not 0 <= table_index < len(tables):
        return None
    rows = tables[table_index].findall(qn("w:tr"))
    if not 0 <= row_index < len(rows):
        return None
    cells = rows[row_index].findall(qn("w:tc"))
    if not 0 <= cell_index < len(cells):
        return None
    paragraphs = cells[cell_index].findall(qn("w:p"))
    return paragraphs[paragraph_index] if 0 <= paragraph_index < len(paragraphs) else None


class UnifiedSynthesizer:
    """通用材料来源与声明式交付调度器。"""

    @classmethod
    def verify_build_plan_hashes(cls, plan: Any) -> None:
        """校验旧字典计划或 R3 PreparedBuild 中引用的输入哈希。"""
        if isinstance(plan, PreparedBuild):
            try:
                plan.verify_inputs_unchanged()
            except ValueError as exc:
                raise BuildError(str(exc)) from exc
            return
        source_hashes = plan.get("source_hashes", {})
        for file_str, expected_hash in source_hashes.items():
            fpath = Path(file_str)
            if not fpath.exists():
                raise BuildError(f"源文件自计划后已丢失: {fpath.name}")
            current_hash = compute_file_sha256(fpath)
            if current_hash != expected_hash:
                raise BuildError(f"源文件自计划后已发生变更 (哈希不匹配): {fpath.name}")

    @classmethod
    def prepare_build(
        cls,
        source_dir: Path,
        manifest_path: Optional[Path] = None,
        override_paths=None,
    ) -> PreparedBuild:
        """只读准备一次项目，供 plan、render 和交付核验共享。"""
        source_path = Path(source_dir).resolve()
        if not source_path.exists() or not (source_path.is_dir() or source_path.is_file()):
            raise BuildError(f"输入路径不存在: {source_path}")
        try:
            if manifest_path is not None and not isinstance(manifest_path, (str, Path)):
                # 供上层预检器直接传入已解析配置，避免再次解析配置和 provenance。
                config = manifest_path
            else:
                config_dir = source_path if source_path.is_dir() else source_path.parent
                config = load_project_config(config_dir, manifest_path, project_name=source_path.stem, override_paths=override_paths)
            return prepare_project_build(source_path, config)
        except (ConfigError, SourceStrategyError, ValueError) as exc:
            raise BuildError(str(exc)) from exc

    @classmethod
    def plan(
        cls,
        source_dir: Path,
        manifest_path: Optional[Path] = None,
        override_paths=None,
    ) -> Dict[str, Any]:
        """只读准备输入并返回兼容旧调用方的可审阅计划。"""
        prepared = cls.prepare_build(source_dir, manifest_path, override_paths=override_paths)
        selected = {spec["filename"] for spec in prepared.deliveries}
        existing = Path("output") / prepared.config.project_name
        unselected = sorted(p.name for p in existing.glob("*.docx") if p.name not in selected)
        if unselected:
            # 不修改 PreparedBuild；只在公共计划中附加环境提示。
            public = prepared.to_public_dict()
            public["warnings"].append("默认输出目录中有不属于本次交付的 DOCX（保留不删除）: " + ", ".join(unselected))
            return public
        return prepared.to_public_dict()

    @classmethod
    def render_body(cls, prepared: PreparedBuild, run_dir: Path) -> RenderedBody:
        """消费既有 PreparedBuild 渲染正文；不重新扫描、推导或载入 RoleMap。"""
        from .composition import add_bookmark, normalize_body_sections
        from .highlighted_docx import normalize_document_pagination

        try:
            prepared.verify_inputs_unchanged()
        except ValueError as exc:
            raise BuildError(str(exc)) from exc
        source_path = prepared.source_path
        config = prepared.config
        run_dir = Path(run_dir)
        context = RenderContext.from_config(
            config,
            exact_pages=prepared.exact_pages if hasattr(prepared, "exact_pages") else None,
            source_dir=source_path if source_path.is_dir() else source_path.parent,
        )
        internal = run_dir / "internal"
        internal.mkdir(parents=True, exist_ok=True)
        body_path = internal / "body.docx"
        print(f"[1/3] 合成可复用正文: {config.project_name}")

        strategy = prepared.strategy
        is_single_docx = source_path.is_file() or strategy == "docx_document"
        if is_single_docx:
            target_file = prepared.source_docx_path
            if target_file is None:
                raise BuildError(f"准备结果未包含 DOCX 源文件: {source_path}")
            if strategy == "docx_document":
                source_policy = getattr(config, "formatting", None)
                document = Document(str(target_file))
                from .style_applier import apply_roles

                assignments = list(prepared.assignments)
                if getattr(config, "resolved_format", None):
                    source_policy = _policy_for_source_file(config.formatting, target_file)
                    role_map = _scoped_role_map(assignments, source_policy, target_file)
                    if role_map or source_policy.mode not in ("preserve", "mixed"):
                        apply_roles(document, role_map, config.resolved_format, policy=source_policy)

                nodes = []
                for idx, assignment in enumerate(assignments):
                    if not assignment.role.startswith("heading."):
                        continue
                    paragraph = _find_paragraph_element(document, assignment.node_ref.element_path)
                    if paragraph is not None:
                        existing_bms = {
                            item.get(qn("w:name"))
                            for item in paragraph.iter(qn("w:bookmarkStart"))
                        }
                        if assignment.bookmark_name and assignment.bookmark_name not in existing_bms:
                            add_bookmark(paragraph, assignment.bookmark_name, len(nodes) + 1)
                    nodes.append({
                        "idx": idx,
                        "level": assignment.level or 1,
                        "title": assignment.title_text or "",
                        "toc_title": assignment.title_text or "",
                        "type": "embedded_heading",
                        "bookmark_name": assignment.bookmark_name,
                        "bm_id": len(nodes) + 1,
                        "file": target_file.name,
                    })
                if any("toc" in item.get("parts", []) for item in config.documents) and not nodes:
                    raise BuildError("文档未包含各级标题，无法生成目录；请在配置中移除 toc 交付部分或为文档添加标题。")
                if not source_policy or (source_policy.mode != "preserve" and not (source_policy.mode == "mixed" and source_policy.page_policy == "source")):
                    normalize_document_pagination(document)
            elif strategy == "highlighted_docx":
                from .highlighted_docx import extract_and_format_highlighted_headings

                document = Document(str(target_file))
                source_policy = _policy_for_source_file(config.formatting, target_file)
                nodes = extract_and_format_highlighted_headings(
                    document,
                    config.source.get("highlight", {}),
                    config.fonts,
                    normalize=not source_policy or source_policy.mode != "preserve",
                    render_context=context,
                )
                if not source_policy or (source_policy.mode != "preserve" and not (source_policy.mode == "mixed" and source_policy.page_policy == "source")):
                    normalize_document_pagination(document)
            else:
                raise BuildError(f"不支持的来源策略: {strategy}")
        else:
            final_tree = list(prepared.final_tree)
            nodes = list(prepared.nodes)
            resolved_files = prepare_conversions(source_path, run_dir)
            document = Document()
            neutralize_document_styles(document)
            if getattr(config, "schema_version", 1) >= 3 and config.resolved_format:
                apply_section_spec(document.sections[0], config.resolved_format)
                from .style_applier import setup_styled_footer
                setup_styled_footer(document.sections[0], config.resolved_format, start_page=1)
            else:
                apply_standard_page_setup(document.sections[0], config.page_setup)
                setup_footer(document.sections[0], start_page=1, font_en=config.fonts["en"])
            context.last_rendered_landscape = False
            for index, node in enumerate(final_tree):
                render_tree_node_recursive(
                    document,
                    node,
                    source_path,
                    context.fonts or config.fonts,
                    is_first_section=index == 0,
                    resolved_files=resolved_files,
                    context=context,
                )

        effective_mode = getattr(getattr(config, "formatting", None), "mode", "restyle")
        if effective_mode != "preserve" and not (
            effective_mode == "mixed" and getattr(config.formatting, "page_policy", "target") == "source"
        ):
            normalize_body_sections(document, config)
        _apply_target_page_policy(document, config)
        document.save(str(body_path))
        return RenderedBody(path=body_path, nodes=tuple(nodes))
    
    @classmethod
    def synthesize(
        cls, source_dir: Path, output_dir: Optional[Path] = None,
        manifest_path: Optional[Path] = None, override_paths=None, keep_work: bool = False,
    ) -> Dict[str, Path]:
        """Normalize the source once, assemble declared documents, verify, then publish."""
        prepared = cls.prepare_build(source_dir, manifest_path, override_paths=override_paths)
        cls.verify_build_plan_hashes(prepared)
        source_path = prepared.source_path
        config = prepared.config
        plan = prepared.to_public_dict()
        out_base = Path(output_dir).resolve() if output_dir else Path("output", config.project_name).resolve()
        if out_base == source_path or (source_path.is_dir() and source_path in out_base.parents):
            raise BuildError("输出目录不能位于输入材料目录内。")
        for warning in plan["warnings"]:
            print(f"[提示] {warning}")
        selected = {spec["filename"] for spec in config.documents}
        extras = sorted(p.name for p in out_base.glob("*.docx") if p.name not in selected)
        if extras:
            print("[保留] 以下文件不属于本次交付，不会删除: " + ", ".join(extras))
        out_base.mkdir(parents=True, exist_ok=True)
        run_dir = out_base / ".work" / f"run-{uuid.uuid4().hex}"
        run_dir.mkdir(parents=True, exist_ok=False)
        from . import renderers
        from .composition import normalize_body_sections
        from .delivery import (
            DeliveryGateError,
            DeliveryReport,
            CheckResult,
            assert_publishable,
            build_deliveries,
            summarize_delivery_reports,
        )
        previous_cache = renderers.CACHE_DIR
        is_v3 = getattr(config, "schema_version", 1) >= 3
        qa_summary = None
        page_evidence = {}
        try:
            set_pdf_render_cache_dir(run_dir / "pdf-render-cache")
            rendered_body = cls.render_body(prepared, run_dir)
            body_path = rendered_body.path
            nodes = list(rendered_body.nodes)
            print(f"[2/3] 装配并核验 {len(config.documents)} 份交付物")
            staged = build_deliveries(
                config,
                body_path,
                nodes,
                run_dir,
                page_maps_out=page_evidence,
                prepared=prepared,
            )

            # 2.5 内容完整性与有效格式核验及元数据留存 (针对 v3 清单工程)
            if is_v3:
                from .content_integrity import (
                    build_expected_inventory,
                    verify_content_integrity,
                    verify_delivery_format,
                    verify_generated_object_allowances,
                )

                reports = []
                parts_registry = getattr(config, "parts", {}) or {}
                from .verification_contracts import build_verification_context
                from .content_integrity import verify_generated_content
                has_content = lambda spec: any(
                    part == "body" or (parts_registry.get(part) and parts_registry[part].kind == "content")
                    for part in spec.get("parts", [])
                )
                for spec in config.documents:
                    doc_id = spec["id"]
                    staged_path = staged.get(doc_id)
                    checks = []
                    unverified_attributes = []
                    expected_inventory = None
                    verification_context = None
                    content_evidence = {}

                    if has_content(spec):
                        if not staged_path or not Path(staged_path).is_file():
                            checks.append(CheckResult(
                                "content_integrity", "failed", True,
                                ("未找到该交付物的 staging 文件",),
                            ))
                        else:
                            try:
                                expected_inventory = build_expected_inventory(prepared, spec=spec)
                                if not expected_inventory.source_order:
                                    raise ValueError("没有可用于内容完整性核验的 DOCX 源节点清单")
                                verify_content_integrity(
                                    expected_inventory,
                                    staged_path,
                                    evidence_out=content_evidence,
                                )
                                checks.append(CheckResult(
                                    "content_integrity", "passed", True,
                                    (f"清单范围: {', '.join(spec.get('parts', []))}",),
                                    evidence={
                                        "expected_inventory": expected_inventory.to_dict(),
                                        **content_evidence,
                                    },
                                ))
                            except Exception as exc:
                                checks.append(CheckResult(
                                    "content_integrity", "failed", True,
                                    (str(exc),),
                                    evidence={
                                        "expected_inventory": expected_inventory.to_dict()
                                        if expected_inventory is not None else {},
                                        **content_evidence,
                                    },
                                ))
                    else:
                        checks.append(CheckResult(
                            "content_integrity", "not_run", False,
                            ("该交付物未声明 content 部件",),
                        ))

                        # A cover/TOC-only v3 delivery does not pass through
                        # verify_content_integrity, but its template media is
                        # still a finite, part-scoped generated allowance.
                        if staged_path and Path(staged_path).is_file() and "cover" in spec.get("parts", ()):
                            try:
                                expected_inventory = build_expected_inventory(prepared, spec=spec)
                                generated_object_report = verify_generated_object_allowances(
                                    staged_path,
                                    expected_inventory,
                                )
                                if generated_object_report["checks"]:
                                    checks.append(CheckResult(
                                        "generated_objects",
                                        "passed" if generated_object_report["passed"] else "failed",
                                        True,
                                        tuple(item.get("kind", "generated_object") for item in generated_object_report["failures"]),
                                        evidence=generated_object_report,
                                    ))
                            except Exception as exc:
                                checks.append(CheckResult(
                                    "generated_objects", "failed", True,
                                    (f"生成对象许可核验器异常: {exc}",),
                                ))

                    if staged_path and Path(staged_path).is_file():
                        try:
                            verification_context = build_verification_context(
                                prepared,
                                rendered_body,
                                spec,
                                staged_path,
                            )
                            generated_report = verify_generated_content(
                                staged_path,
                                verification_context.expected_delivery,
                            )
                            if generated_report["checks"]:
                                generated_diagnostics = tuple(
                                    f"{item.get('part_id')}/{item.get('field')}: {item.get('status')}"
                                    for item in generated_report["failures"]
                                )
                                checks.append(CheckResult(
                                    "generated_content",
                                    "passed" if generated_report["passed"] else "failed",
                                    True,
                                    generated_diagnostics,
                                    evidence=generated_report,
                                ))
                        except Exception as exc:
                            checks.append(CheckResult(
                                "generated_content", "failed", True,
                                (f"生成内容核验器异常: {exc}",),
                            ))

                    if not staged_path or not Path(staged_path).is_file():
                        checks.append(CheckResult(
                            "format_verification", "failed", True,
                            ("未找到该交付物的 staging 文件",),
                        ))
                    elif not getattr(config, "resolved_format", None):
                        checks.append(CheckResult(
                            "format_verification", "unsupported", True,
                            ("v3 清单工程没有解析出的格式包",),
                        ))
                    else:
                        try:
                            if verification_context is None:
                                verification_context = build_verification_context(
                                    prepared,
                                    rendered_body,
                                    spec,
                                    staged_path,
                                )
                            fmt_rep = verify_delivery_format(
                                staged_path,
                                config.resolved_format,
                                verification_context=verification_context,
                            )
                            unverified_attributes.extend(fmt_rep.unverified_attributes)
                            checks.append(CheckResult(
                                "format_verification",
                                "passed" if fmt_rep.passed else "failed",
                                True,
                                tuple(fmt_rep.violations),
                                evidence={
                                    **fmt_rep.to_dict(),
                                    "verification_context": verification_context.to_dict(),
                                },
                            ))
                        except Exception as exc:
                            checks.append(CheckResult(
                                "format_verification", "failed", True,
                                (f"格式核验器异常: {exc}",),
                            ))

                    reports.append(DeliveryReport(
                        delivery_id=doc_id,
                        filename=spec["filename"],
                        checks=tuple(checks),
                        source_nodes=(
                            expected_inventory.source_order
                            if expected_inventory is not None
                            else tuple(str(item) for item in prepared.source_order)
                        ),
                        unverified_attributes=tuple(sorted(set(unverified_attributes))),
                    ))

                qa_summary = summarize_delivery_reports(reports)
                assert_publishable(reports)
                # The runtime config remains available to renderers for
                # compatibility, but the prepared contract must still be the
                # one that is published after all verification has completed.
                prepared.verify_inputs_unchanged()

                metadata_path = run_dir / "build-metadata.json"
                metadata_dict = generate_build_metadata(
                    config,
                    staged,
                    plan.get("source_hashes", {}),
                    qa_summary=qa_summary,
                    page_evidence=page_evidence,
                )
                with open(metadata_path, "w", encoding="utf-8") as f:
                    json.dump(metadata_dict, f, ensure_ascii=False, indent=2)
                staged["__metadata__"] = metadata_path

            print("[3/3] 所有交付物已通过检查，发布本次成果")
            return _publish_deliveries(staged, out_base, run_dir)
        except (ValueError, OSError, SourceStrategyError, OfficeExportError) as exc:
            error = BuildError(str(exc))
            if is_v3:
                _write_failure_diagnostics(out_base, error, qa_summary)
            raise error from exc
        except Exception as exc:
            from .content_integrity import ContentIntegrityError, FormatVerificationError
            if isinstance(exc, (ContentIntegrityError, FormatVerificationError, DeliveryGateError)):
                error = BuildError(str(exc))
                if is_v3:
                    _write_failure_diagnostics(out_base, error, qa_summary)
                raise error from exc
            if isinstance(exc, BuildError) and is_v3:
                _write_failure_diagnostics(out_base, exc, qa_summary)
            raise
        finally:
            renderers.CACHE_DIR = previous_cache
            if keep_work or getattr(sys.exc_info()[1], "preserve_work", False):
                print(f"[诊断] 已保留本次工作目录: {run_dir}")
            else:
                shutil.rmtree(run_dir)
                try:
                    run_dir.parent.rmdir()
                except OSError:
                    pass


def generate_build_metadata(
    config: Any,
    staged: Dict[str, Path],
    source_hashes: Dict[str, str],
    qa_summary: Optional[Dict[str, Any]] = None,
    page_evidence: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """生成符合 B7 审计规范的构建元数据 build-metadata.json 内容。"""
    from datetime import datetime, timezone
    import platform
    from .build_plan import compute_config_contract_hash
    from .content_integrity import compute_file_sha256

    is_v3 = getattr(config, "schema_version", 1) >= 3
    format_info = {}
    if getattr(config, "resolved_format", None):
        rf = config.resolved_format
        format_info = {
            "id": rf.id,
            "version": rf.version,
            "capabilities": getattr(rf, "required_capabilities", []),
            "source_sha256": getattr(rf, "content_hash", ""),
        }
    else:
        format_info = {
            "id": "legacy-builtin",
            "version": "legacy",
            "capabilities": ["styles.fonts.basic", "pagination.continuous"],
        }

    deliveries_info = {}
    page_evidence = page_evidence or {}
    for spec in getattr(config, "documents", []):
        doc_id = spec["id"]
        file_path = staged.get(doc_id)
        if file_path and file_path.is_file():
            delivery_info = {
                "filename": spec["filename"],
                "parts": spec.get("parts", []),
                "sha256": compute_file_sha256(file_path),
            }
            evidence = page_evidence.get(doc_id, {})
            pdf_path = Path(evidence.get("pdf_path", "")) if evidence.get("pdf_path") else None
            if pdf_path and pdf_path.is_file():
                delivery_info["pdf"] = {
                    "filename": pdf_path.name,
                    "sha256": compute_file_sha256(pdf_path),
                }
            if evidence.get("page_map"):
                delivery_info["page_map"] = evidence["page_map"]
            if evidence.get("page_records"):
                delivery_info["page_records"] = evidence["page_records"]
            if evidence.get("pagination_convergence") is not None:
                delivery_info["pagination_convergence"] = evidence["pagination_convergence"]
            if evidence.get("field_cache") is not None:
                delivery_info["field_cache"] = evidence["field_cache"]
            deliveries_info[doc_id] = delivery_info

    from .qa import word_automation_status
    word_available = False
    word_reason = "未执行 Word Automation 探针。"
    try:
        word_available, word_reason = word_automation_status()
    except Exception:
        pass

    manifest_path = getattr(config, "manifest_path", None)
    manifest_hash = compute_file_sha256(Path(manifest_path)) if manifest_path else ""
    override_hashes = {
        str(path): compute_file_sha256(Path(path))
        for path in getattr(config, "override_paths", [])
        if compute_file_sha256(Path(path))
    }

    return {
        "generator": "document-synthesis",
        "engine_version": "3.0.0",
        "verification_contract_version": 1,
        "schema_version": getattr(config, "schema_version", 3 if is_v3 else 2),
        "project_name": getattr(config, "project_name", "unknown"),
        "build_timestamp": datetime.now(timezone.utc).isoformat(),
        "format": format_info,
        "cover": getattr(getattr(config, "cover_spec", None), "to_dict", lambda: {})(),
        "configuration": {
            "manifest_sha256": manifest_hash,
            "override_sha256": override_hashes,
            "config_contract_sha256": compute_config_contract_hash(config),
        },
        "source_hashes": source_hashes,
        "deliveries": deliveries_info,
        "environment": {
            "python_version": sys.version.split()[0],
            "platform": platform.platform(),
            "word_available": word_available,
            "word_automation": {
                "available": word_available,
                "reason": word_reason,
            },
        },
        "qa": qa_summary or {
            "content_integrity": "passed",
            "format_verification": "passed" if getattr(config, "resolved_format", None) else "skipped_legacy",
            "unverified_attributes": [],
        },
    }


def _publish_deliveries(staged, output_dir, run_dir):
    """Rollback per-file replacement failures; no files change before all QA succeeds."""
    backups = Path(run_dir) / "publish-backups"
    backups.mkdir()
    targets, originals = {}, {}
    for key, source in staged.items():
        target = Path(output_dir) / source.name
        if target.is_symlink() or (target.exists() and not target.is_file()):
            raise BuildError(f"交付目标不是普通文件，拒绝替换: {target}")
        targets[key] = target
        if target.exists():
            backup = backups / source.name
            shutil.copy2(target, backup)
            originals[key] = backup
    replaced = []
    try:
        for key, source in staged.items():
            os.replace(source, targets[key])
            replaced.append(key)
    except BaseException as publication_error:
        recovery_errors = []
        for key in reversed(replaced):
            try:
                if key in originals:
                    os.replace(originals[key], targets[key])
                else:
                    targets[key].unlink()
            except OSError as exc:
                recovery_errors.append(str(exc))
        if recovery_errors:
            error = BuildError(f"发布失败且回滚未完成。原文件备份已保留于 {backups}: {'; '.join(recovery_errors)}")
            error.preserve_work = True
            raise error from publication_error
        raise
    return targets
