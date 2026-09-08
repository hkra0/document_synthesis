# -*- coding: utf-8 -*-
"""Self-contained, fictional HTML preview for a resolved format package.

The preview is intentionally an approximation: it demonstrates the role
hierarchy and managed attributes without rendering sample prose or claiming
Word/PDF pagination evidence.
"""

from __future__ import annotations

import html
import json
import platform
from pathlib import Path
from typing import Any, Dict, Optional, Union

from .format_resolver import resolve_format_package
from .qa import word_export_status


_FICTIONAL_TEXT = {
    "title": "通用文档格式预览",
    "subtitle": "虚构副标题，仅用于检查层级",
    "body": "这是用于预览的虚构正文。它不来自任何样本文档，也不代表实际交付内容。",
    "quote": "这是一段虚构引用，用于检查缩进和段落间距。",
    "caption": "图 1  虚构图片占位说明",
    "bibliography": "[1] 虚构参考资料，格式预览用。",
}


def _css_value(value: Any) -> str:
    if value is None:
        return ""
    return html.escape(str(value), quote=True)


def _role_css(resolved: Any, role: str) -> str:
    role_spec = resolved.roles.get(role)
    if role_spec is None:
        return ""
    style = resolved.styles.get(role_spec.style)
    if style is None:
        return ""
    rules = []
    run = style.run
    para = style.paragraph
    if run:
        if run.east_asia:
            rules.append(f"font-family:'{_css_value(run.east_asia)}', sans-serif")
        elif run.latin:
            rules.append(f"font-family:'{_css_value(run.latin)}', sans-serif")
        if run.size_pt:
            rules.append(f"font-size:{_css_value(run.size_pt)}pt")
        if run.bold is not None:
            rules.append(f"font-weight:{'700' if run.bold else '400'}")
        if run.italic is not None:
            rules.append(f"font-style:{'italic' if run.italic else 'normal'}")
        if run.color and run.color != "auto":
            rules.append(f"color:#{_css_value(run.color).lstrip('#')}")
    if para:
        if para.alignment:
            rules.append(f"text-align:{_css_value(para.alignment)}")
        if para.space_before_pt is not None:
            rules.append(f"margin-top:{_css_value(para.space_before_pt)}pt")
        if para.space_after_pt is not None:
            rules.append(f"margin-bottom:{_css_value(para.space_after_pt)}pt")
        if para.line_spacing:
            if para.line_spacing.mode == "multiple" and para.line_spacing.value:
                rules.append(f"line-height:{_css_value(para.line_spacing.value)}")
            elif para.line_spacing.value:
                rules.append(f"line-height:{_css_value(para.line_spacing.value)}pt")
        if para.first_line_indent:
            unit = "em" if para.first_line_indent.unit == "char" else para.first_line_indent.unit
            rules.append(f"text-indent:{_css_value(para.first_line_indent.value)}{unit}")
        if para.left_indent:
            rules.append(f"margin-left:{_css_value(para.left_indent.value)}{para.left_indent.unit}")
        if para.right_indent:
            rules.append(f"margin-right:{_css_value(para.right_indent.value)}{para.right_indent.unit}")
    return ";".join(rules)


def _preview_body(resolved: Any) -> str:
    def para(role: str, text: str, extra: str = "") -> str:
        css = _role_css(resolved, role)
        return f'<p data-role="{html.escape(role)}" style="{css};{extra}">{html.escape(text)}</p>'

    parts = [para("title", _FICTIONAL_TEXT["title"])]
    for level in range(1, 10):
        role = f"heading.{level}"
        text = f"{level}. 虚构第 {level} 级标题"
        parts.append(para(role, text))
    parts.append(para("subtitle", _FICTIONAL_TEXT["subtitle"]))
    parts.append(para("body", _FICTIONAL_TEXT["body"]))
    parts.append(para("body", "第二段虚构正文，检查连续正文的统一格式。"))
    parts.append(para("quote", _FICTIONAL_TEXT["quote"]))
    parts.append('<ul data-role="list"><li>虚构列表项目一</li><li>虚构列表项目二</li></ul>')
    parts.append(
        '<table data-role="table.body"><thead><tr><th>虚构表头 A</th><th>虚构表头 B</th></tr></thead>'
        '<tbody><tr><td>示例值 1</td><td>示例值 2</td></tr></tbody></table>'
    )
    parts.append('<figure data-role="image"><div class="image-placeholder">虚构图片占位</div>' + para("caption", _FICTIONAL_TEXT["caption"]) + "</figure>")
    parts.append('<section data-role="toc"><h2>目录（虚构预览）</h2><ol><li>1. 虚构第 1 级标题 …… 1</li><li>2. 虚构第 2 级标题 …… 2</li></ol></section>')
    parts.append(para("bibliography", _FICTIONAL_TEXT["bibliography"]))
    return "\n".join(parts)


def build_preview(
    format_ref: Union[str, Path],
    output_path: Union[str, Path],
    *,
    engine: str = "html-css-approximate",
    metadata_path: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """Render a fictional preview and return its evidence metadata."""
    format_path = Path(format_ref).resolve()
    resolved = resolve_format_package(str(format_path))
    output = Path(output_path).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    word_available, word_reason = word_export_status()
    metadata = {
        "preview_schema_version": 1,
        "format_id": resolved.id,
        "format_version": resolved.version,
        "engine": engine,
        "approximate": True,
        "fictional_content": True,
        "source_content_included": False,
        "word_environment": {
            "platform": platform.platform(),
            "available": word_available,
            "status": "not_run",
            "message": word_reason,
        },
        "roles_rendered": ["title", "subtitle", "body", "quote", "caption", "bibliography"]
        + [f"heading.{level}" for level in range(1, 10)]
        + ["list", "table.body", "image", "toc", "footer"],
    }
    css = """body{background:#e5e7eb;color:#111827;font-family:system-ui,sans-serif;margin:0;padding:2rem}.page{background:#fff;max-width:210mm;min-height:297mm;margin:auto;padding:25mm;box-sizing:border-box;box-shadow:0 2px 8px #0002}.notice{border:1px solid #f59e0b;background:#fffbeb;padding:.75rem;margin-bottom:1rem;font-size:.9rem}.image-placeholder{height:90px;background:#dbeafe;border:1px dashed #2563eb;display:flex;align-items:center;justify-content:center;color:#1d4ed8}table{border-collapse:collapse;width:100%;margin:1rem 0}th,td{border:1px solid #9ca3af;padding:.5rem}figure{margin:1rem 0}footer{border-top:1px solid #9ca3af;margin-top:2rem;padding-top:.5rem;text-align:center;color:#4b5563}"""
    body = _preview_body(resolved)
    html_text = """<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>格式近似预览</title><style>""" + css + """</style></head><body><main class="page"><div class="notice"><strong>近似预览</strong> · 引擎：""" + html.escape(engine) + """ · 使用虚构文本，不构成 Word/PDF 页面证据。</div>""" + body + "<footer data-role=\"footer\">— 1 —（虚构页脚预览）</footer></main></body></html>"
    output.write_text(html_text, encoding="utf-8")
    if metadata_path:
        meta_path = Path(metadata_path).resolve()
        meta_path.parent.mkdir(parents=True, exist_ok=True)
        meta_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    return metadata
