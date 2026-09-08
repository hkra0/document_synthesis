# -*- coding: utf-8 -*-
"""
格式校正报告生成与格式包编译器 (lib/format_review.py)
生成独立自包含的 HTML 交互校正报告，提供候选角色核对、参数微调与决策导出；
依据确认的决策与候选规格，严格编译生成符合 schemas/format-v1.schema.json 的独立格式包，
且坚决不泄露样本文档中的任何敏感正文字符串。
"""

import copy
import html
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

import jsonschema

from .config import ConfigError
from .contracts import (
    ContractError,
    require_final_decisions,
    validate_analysis_data,
    validate_decisions_data,
    validate_role_map_data,
)
from .format_resolver import resolve_format_package
from .docx_inspector import NodeRef
from .role_mapper import RoleAssignment, serialize_role_map

SCHEMA_PATH = Path(__file__).resolve().parent.parent / "schemas" / "format-v1.schema.json"


def load_format_schema() -> Dict[str, Any]:
    with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def build_decisions_example(analysis_data: Dict[str, Any]) -> Dict[str, Any]:
    """Create the canonical v1 decision starter for CLI and HTML consumers."""
    return {
        "decisions_schema_version": 1,
        "report_id": analysis_data.get("report_id"),
        "source_sha256": analysis_data.get("source_sha256"),
        "role_styles": {
            role: cand["cluster_id"]
            for role, cand in analysis_data.get("candidate_roles", {}).items()
        },
        "style_overrides": {},
        "missing_roles": {
            item["role"]: "pending"
            for item in analysis_data.get("missing_roles", [])
        },
    }


def generate_html_review(
    analysis_data: Dict[str, Any],
    output_path: Union[str, Path],
) -> Path:
    """
    生成单文件自包含的交互式 HTML 校正报告。
    所有内嵌数据与样本片断严格执行 HTML 实体转义。
    """
    out_file = Path(output_path).resolve()
    out_file.parent.mkdir(parents=True, exist_ok=True)

    report_id = html.escape(analysis_data.get("report_id", "report-unknown"))
    source_sha256 = html.escape(analysis_data.get("source_sha256", ""))
    source_file = html.escape(analysis_data.get("source_file", "unknown.docx"))
    analyzer_version = html.escape(analysis_data.get("analyzer_version", "1.0.0"))

    clusters = analysis_data.get("clusters", [])
    candidate_roles = analysis_data.get("candidate_roles", {})
    node_roles = analysis_data.get("node_roles", {})
    node_refs = analysis_data.get("node_refs", {})
    missing_roles = analysis_data.get("missing_roles", [])
    page = analysis_data.get("page", {})

    # 预准备默认决策 JSON
    initial_decisions = build_decisions_example(analysis_data)
    decisions_json_str = html.escape(json.dumps(initial_decisions, ensure_ascii=False, indent=2))
    initial_decisions_js_escaped = json.dumps(initial_decisions).replace('\\', '\\\\').replace("'", "\\'")
    analysis_json_str = html.escape(json.dumps(analysis_data, ensure_ascii=False))

    # 构造表格行
    cluster_rows = []
    cluster_role_options = (
        "body", "title", "subtitle",
        *(f"heading.{level}" for level in range(1, 10)),
        "table.body", "quote", "caption", "bibliography",
    )
    for c in clusters:
        cid = html.escape(c["cluster_id"])
        r = c["run_style"]
        p = c["paragraph_style"]

        font_desc = f"{html.escape(r.get('east_asia', '宋体'))} / {html.escape(r.get('latin', 'Times New Roman'))} · {r.get('size_pt', 12)}pt"
        if r.get("bold"):
            font_desc += " · <b>加粗</b>"
        if r.get("italic"):
            font_desc += " · <i>倾斜</i>"

        align_desc = html.escape(p.get("alignment", "left"))
        indent_desc = f"首行 {p.get('first_line_indent', {}).get('value', 0)}{p.get('first_line_indent', {}).get('unit', '')}" if p.get("first_line_indent") else "无缩进"
        line_desc = f"{p.get('line_spacing', {}).get('mode', 'single')} ({p.get('line_spacing', {}).get('value', 1.0)})" if p.get("line_spacing") else "单倍行距"

        sample_escaped = "<br>".join(f'<span class="sample-snippet">“{html.escape(t)}”</span>' for t in c.get("sample_texts", [])[:2])
        if not sample_escaped:
            sample_escaped = '<span class="text-muted">（无文字段落或空行）</span>'

        sug_role = c.get("suggested_role", "body")
        conf_pct = int(c.get("confidence", 0.5) * 100)
        size_value = html.escape(str(r.get("size_pt", "")))
        alignment_value = html.escape(str(p.get("alignment", "left")))

        role_options = "".join(
            f'<option value="{html.escape(option, quote=True)}" '
            f'{"selected" if option == sug_role else ""}>{html.escape(option)}</option>'
            for option in cluster_role_options
        )
        cluster_rows.append(f"""
        <tr data-cluster-row="{cid}">
            <td><code>{cid}</code></td>
            <td><strong>{font_desc}</strong><br><small>{align_desc} · {indent_desc} · {line_desc}</small></td>
            <td><span class="badge badge-count">{c.get('occurrence_count', 0)} 段</span></td>
            <td>
                <select class="role-select" data-cluster="{cid}">
                    {role_options}
                </select>
                <div class="conf-bar"><span style="width: {conf_pct}%"></span></div>
                <small class="text-muted">置信度 {conf_pct}%</small>
                <div class="attribute-editor">
                    <label>字号 <input class="size-input" data-cluster="{cid}" type="number" min="1" max="200" step="0.1" value="{size_value}" placeholder="未设置"></label>
                    <label>对齐 <select class="alignment-input" data-cluster="{cid}">
                        <option value="left" {'selected' if alignment_value == 'left' else ''}>左</option>
                        <option value="center" {'selected' if alignment_value == 'center' else ''}>中</option>
                        <option value="right" {'selected' if alignment_value == 'right' else ''}>右</option>
                        <option value="justify" {'selected' if alignment_value == 'justify' else ''}>两端</option>
                    </select></label>
                </div>
            </td>
            <td>{sample_escaped}</td>
        </tr>
        """)

    # 缺失角色行
    missing_rows = []
    for m in missing_roles:
        m_role = html.escape(m["role"])
        reason = html.escape(m.get("reason", ""))
        suggested = html.escape(m.get("suggested_fallback", "inherit"))
        missing_rows.append(f"""
        <tr>
            <td><code>{m_role}</code></td>
            <td><span class="text-muted">{reason}</span></td>
            <td>
                <select class="missing-select" data-role="{m_role}">
                    <option value="pending" selected>待人工确认（不可直接编译）</option>
                    <option value="inherit">从基础预设/父级继承 ({suggested})</option>
                    <option value="reject">明确拒绝此角色</option>
                </select>
            </td>
        </tr>
        """)
    if not missing_rows:
        missing_rows.append('<tr><td colspan="3" class="text-muted">样本文档覆盖了所有基础语义角色，无缺失项。</td></tr>')

    # Cluster choices are the batch operation.  Keep per-node choices as
    # explicit exceptions so the exported decisions file does not silently
    # turn every inferred assignment into a hand-authored override.
    node_role_options = (
        "body", "title", "subtitle", "heading.1", "heading.2", "heading.3",
        "heading.4", "heading.5", "heading.6", "heading.7", "heading.8", "heading.9",
        "table.body", "quote", "caption", "bibliography",
    )
    node_rows = []
    for element_path, role in node_roles.items():
        node_ref = node_refs.get(element_path, {})
        safe_path = html.escape(str(element_path), quote=True)
        safe_role = html.escape(str(role), quote=True)
        text_hash = html.escape(str(node_ref.get("text_hash", ""))[:12], quote=True)
        options = "".join(
            f'<option value="{html.escape(option, quote=True)}" '
            f'{"selected" if option == role else ""}>{html.escape(option)}</option>'
            for option in node_role_options
        )
        node_rows.append(f"""
        <tr>
            <td><code class="node-ref">{safe_path}</code><br><small class="text-muted">文本指纹 {text_hash}</small></td>
            <td><select class="node-role-select" data-node="{safe_path}" data-initial-role="{safe_role}">{options}</select></td>
        </tr>
        """)
    if not node_rows:
        node_rows.append('<tr><td colspan="2" class="text-muted">没有可供例外校正的节点。</td></tr>')

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>格式分析校正报告 · {source_file}</title>
    <style>
        :root {{
            --bg: #f8fafc;
            --surface: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --success: #16a34a;
            --warning: #ea580c;
            --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}
        body {{
            margin: 0;
            padding: 0;
            font-family: var(--font-family);
            background: var(--bg);
            color: var(--text-main);
            line-height: 1.5;
        }}
        .container {{
            max-width: 1200px;
            margin: 2rem auto;
            padding: 0 1.5rem;
        }}
        .card {{
            background: var(--surface);
            border-radius: 12px;
            border: 1px solid var(--border);
            padding: 1.5rem;
            margin-bottom: 1.5rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 1rem;
            margin-bottom: 1.5rem;
        }}
        h1, h2, h3 {{ margin: 0 0 0.5rem 0; font-weight: 600; }}
        h1 {{ font-size: 1.5rem; }}
        h2 {{ font-size: 1.2rem; color: #334155; }}
        .meta-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1rem;
            margin-bottom: 1rem;
        }}
        .meta-item {{ font-size: 0.875rem; }}
        .meta-item .label {{ color: var(--text-muted); display: block; margin-bottom: 0.2rem; }}
        .meta-item .val {{ font-family: monospace; font-size: 0.85rem; word-break: break-all; }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.9rem;
            margin-top: 0.5rem;
        }}
        th, td {{
            text-align: left;
            padding: 0.75rem 1rem;
            border-bottom: 1px solid var(--border);
            vertical-align: middle;
        }}
        th {{ background: #f1f5f9; color: var(--text-muted); font-weight: 600; font-size: 0.8rem; text-transform: uppercase; }}
        select {{
            padding: 0.4rem 0.6rem;
            border-radius: 6px;
            border: 1px solid var(--border);
            background: #fff;
            font-size: 0.85rem;
            color: var(--text-main);
            outline: none;
        }}
        select:focus {{ border-color: var(--primary); }}
        .badge {{
            display: inline-block;
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
        }}
        .badge-count {{ background: #e0e7ff; color: #3730a3; }}
        .conf-bar {{
            height: 4px;
            background: #e2e8f0;
            border-radius: 2px;
            overflow: hidden;
            margin-top: 4px;
        }}
        .conf-bar span {{
            display: block;
            height: 100%;
            background: var(--primary);
        }}
        .sample-snippet {{
            font-size: 0.8rem;
            color: #475569;
            background: #f8fafc;
            padding: 2px 4px;
            border-radius: 4px;
            display: inline-block;
            margin-bottom: 2px;
        }}
        .btn {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 0.6rem 1.2rem;
            border-radius: 8px;
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            border: none;
            transition: all 0.15s ease;
            background: var(--primary);
            color: #fff;
        }}
        .btn:hover {{ background: var(--primary-hover); }}
        .btn-secondary {{ background: #e2e8f0; color: #1e293b; }}
        .btn-secondary:hover {{ background: #cbd5e1; }}
        pre {{
            background: #0f172a;
            color: #f8fafc;
            padding: 1rem;
            border-radius: 8px;
            font-size: 0.85rem;
            overflow-x: auto;
            max-height: 280px;
        }}
        .preview-box {{
            background: #ffffff;
            border: 1px dashed #cbd5e1;
            padding: 1.5rem;
            border-radius: 8px;
            margin-top: 1rem;
        }}
        .text-muted {{ color: var(--text-muted); font-size: 0.8rem; }}
        .attribute-editor {{ display: flex; gap: 0.5rem; flex-wrap: wrap; margin-top: 0.45rem; font-size: 0.75rem; }}
        .attribute-editor input {{ width: 4.5rem; padding: 0.25rem; border: 1px solid var(--border); border-radius: 4px; }}
        .attribute-editor select {{ padding: 0.25rem; font-size: 0.75rem; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="header">
                <div>
                    <h1>样本文档格式分析与校正报告</h1>
                    <span class="text-muted">报告编号: {report_id} · 分析引擎 v{analyzer_version}</span>
                </div>
                <div>
                    <label class="text-muted" style="display: block; margin-bottom: 0.35rem;">
                        <input type="checkbox" id="reviewerAttested">
                        我确认已人工审阅并核对上述决定
                    </label>
                    <button class="btn" id="btnDownloadDecisions">导出 decisions.json</button>
                </div>
            </div>

            <div class="meta-grid">
                <div class="meta-item">
                    <span class="label">源文档材料</span>
                    <span class="val">{source_file}</span>
                </div>
                <div class="meta-item">
                    <span class="label">SHA-256 唯一指纹</span>
                    <span class="val">{source_sha256}</span>
                </div>
                <div class="meta-item">
                    <span class="label">版面尺寸 (宽×高)</span>
                    <span class="val">{page.get('width_mm', 210)}mm × {page.get('height_mm', 297)}mm ({page.get('orientation', 'portrait')})</span>
                </div>
                <div class="meta-item">
                    <span class="label">边距 (上下 / 左右)</span>
                    <span class="val">{page.get('margin_top_mm', 25.4)}/{page.get('margin_bottom_mm', 25.4)}mm · {page.get('margin_left_mm', 31.8)}/{page.get('margin_right_mm', 31.8)}mm</span>
                </div>
            </div>
        </div>

        <div class="card">
            <h2>1. 视觉样式聚类与语义角色候选</h2>
            <p class="text-muted">系统已对源文档所有段落完成聚类推导。您可以在下拉菜单中核对并调整各聚类对应的目标角色：</p>
            <table>
                <thead>
                    <tr>
                        <th style="width: 10%;">聚类标识</th>
                        <th style="width: 30%;">排版属性 (字号/字体/段距/行距)</th>
                        <th style="width: 10%;">出现频次</th>
                        <th style="width: 25%;">推导语义角色</th>
                        <th style="width: 25%;">样本片断证据</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(cluster_rows)}
                </tbody>
            </table>
        </div>

        <div class="card">
            <h2>2. 缺失规范角色处理</h2>
            <p class="text-muted">若样本文档未出现规范所需的角色，可指定继承行为：</p>
            <table>
                <thead>
                    <tr>
                        <th style="width: 20%;">缺失角色</th>
                        <th style="width: 50%;">分析原因</th>
                        <th style="width: 30%;">处理决策</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(missing_rows)}
                </tbody>
            </table>
        </div>

        <div class="card">
            <h2>2.5. 节点例外校正</h2>
            <p class="text-muted">聚类选择会批量作用于同类节点；只有确需偏离聚类结论时，才在此为单个 NodeRef 指定例外角色。导出文件仅记录发生变化的例外。</p>
            <table>
                <thead>
                    <tr><th style="width: 70%;">来源节点（NodeRef）</th><th style="width: 30%;">例外角色</th></tr>
                </thead>
                <tbody>{''.join(node_rows)}</tbody>
            </table>
        </div>

        <div class="card">
            <h2>3. 虚构内容排版预览 (零正文泄露保护)</h2>
            <p class="text-muted">预览使用中性测试文字展示各角色规格视觉层次，杜绝展示样本文档真实涉密文字：</p>
            <div class="preview-box" id="previewArea">
                <div style="text-align: center; font-size: 22pt; font-weight: bold; margin-bottom: 12pt;">
                    通用公文排版大标题测试样例
                </div>
                <div style="font-size: 16pt; font-weight: bold; margin-top: 14pt; margin-bottom: 6pt;">
                    一、 第一章 项目建设背景与总体原则
                </div>
                <div style="font-size: 15pt; font-weight: bold; margin-top: 10pt; margin-bottom: 4pt;">
                    （一） 建设目标与核心任务
                </div>
                <div style="font-size: 14pt; text-indent: 2em; line-height: 28pt; margin-bottom: 6pt;">
                    这是标准正文排版效果预览段落。为了保护原始样本文档的知识产权与数据隐私，此处的预览完全采用虚构中性文本展示字号、行距、首行缩进与字体的视觉搭配效果。
                </div>
                <div style="font-size: 14pt; text-indent: 2em; line-height: 28pt;">
                    第二段正文继续保持完全一致的段落格式与排版规范，确保整篇公文阅读体验流畅和谐。
                </div>
            </div>
        </div>

        <div class="card">
            <h2>4. 生成的 decisions.json 预览</h2>
            <p class="text-muted">此决策数据可保存至本地，用于后续执行 <code>synthesize.py --compile-format</code> 自动化编译输出标准格式包：</p>
            <pre id="decisionsPre">{decisions_json_str}</pre>
        </div>
    </div>

    <script>
        const initialData = JSON.parse('{initial_decisions_js_escaped}');
        let currentDecisions = Object.assign({{}}, initialData);
        let reviewStartedAt = null;
        const reviewOperations = [];

        function recordReviewOperation(operation, target, fromValue, toValue) {{
            const from = fromValue === undefined ? null : (fromValue === '' ? null : String(fromValue));
            const to = toValue === undefined ? null : (toValue === '' ? null : String(toValue));
            if (from === to) return;
            if (!reviewStartedAt) reviewStartedAt = new Date().toISOString();
            reviewOperations.push({{
                sequence: reviewOperations.length + 1,
                operation: operation,
                target: String(target),
                from: from,
                to: to
            }});
            syncReviewAudit();
        }}

        function syncReviewAudit(completed) {{
            if (!reviewOperations.length) return;
            const attested = Boolean(document.getElementById('reviewerAttested')?.checked);
            currentDecisions.review_audit = {{
                audit_schema_version: 1,
                report_id: currentDecisions.report_id,
                source_sha256: currentDecisions.source_sha256,
                origin: 'browser_ui',
                automation: false,
                reviewer_attested: attested,
                started_at: reviewStartedAt,
                operation_count: reviewOperations.length,
                operations: reviewOperations.map(item => Object.assign({{}}, item))
            }};
            if (completed) currentDecisions.review_audit.completed_at = new Date().toISOString();
        }}

        function updateDecisions() {{
            const roleStyles = {{}};
            document.querySelectorAll('.role-select').forEach(sel => {{
                const cid = sel.dataset.cluster;
                const role = sel.value;
                roleStyles[role] = cid;
            }});
            currentDecisions.role_styles = roleStyles;

            const styleOverrides = {{}};
            document.querySelectorAll('.role-select').forEach(sel => {{
                const role = sel.value;
                const row = sel.closest('tr');
                const size = row.querySelector('.size-input')?.value;
                const alignment = row.querySelector('.alignment-input')?.value;
                const run = {{}};
                const paragraph = {{}};
                if (size) run.size_pt = Number(size);
                if (alignment) paragraph.alignment = alignment;
                if (Object.keys(run).length || Object.keys(paragraph).length) {{
                    styleOverrides[role] = {{}};
                    if (Object.keys(run).length) styleOverrides[role].run = run;
                    if (Object.keys(paragraph).length) styleOverrides[role].paragraph = paragraph;
                }}
            }});
            currentDecisions.style_overrides = styleOverrides;

            const missingDecisions = {{}};
            document.querySelectorAll('.missing-select').forEach(sel => {{
                const role = sel.dataset.role;
                // Selecting a cluster explicitly resolves that role.  Do not
                // export the stale "missing" decision from the first-pass
                // analysis alongside the new role_styles assignment.
                if (!roleStyles[role]) missingDecisions[role] = sel.value;
            }});
            currentDecisions.missing_roles = missingDecisions;

            const nodeRoles = {{}};
            document.querySelectorAll('.node-role-select').forEach(sel => {{
                // Do not serialize unchanged inferred roles.  This keeps the
                // distinction between analysis output and human exceptions.
                if (sel.value !== sel.dataset.initialRole) {{
                    nodeRoles[sel.dataset.node] = sel.value;
                }}
            }});
            if (Object.keys(nodeRoles).length) currentDecisions.node_roles = nodeRoles;
            else delete currentDecisions.node_roles;

            document.getElementById('decisionsPre').textContent = JSON.stringify(currentDecisions, null, 2);
        }}

        document.querySelectorAll('.role-select').forEach(sel => {{
            sel.dataset.lastValue = sel.value;
            sel.addEventListener('change', () => {{
                recordReviewOperation('select_cluster', sel.dataset.cluster, sel.dataset.lastValue, sel.value);
                sel.dataset.lastValue = sel.value;
                updateDecisions();
            }});
        }});
        document.querySelectorAll('.missing-select').forEach(sel => {{
            sel.dataset.lastValue = sel.value;
            sel.addEventListener('change', () => {{
                recordReviewOperation('set_missing_role', sel.dataset.role, sel.dataset.lastValue, sel.value);
                sel.dataset.lastValue = sel.value;
                updateDecisions();
            }});
        }});
        document.querySelectorAll('.node-role-select').forEach(sel => {{
            sel.dataset.lastValue = sel.value;
            sel.addEventListener('change', () => {{
                recordReviewOperation('set_node_role', sel.dataset.node, sel.dataset.lastValue, sel.value);
                sel.dataset.lastValue = sel.value;
                updateDecisions();
            }});
        }});
        document.querySelectorAll('.size-input, .alignment-input').forEach(input => {{
            input.dataset.lastValue = input.value;
            input.addEventListener('input', () => {{
                recordReviewOperation('set_style_override', input.dataset.cluster, input.dataset.lastValue, input.value);
                input.dataset.lastValue = input.value;
                updateDecisions();
            }});
        }});
        document.getElementById('reviewerAttested')?.addEventListener('change', () => {{
            syncReviewAudit(false);
            updateDecisions();
        }});

        document.getElementById('btnDownloadDecisions').addEventListener('click', () => {{
            updateDecisions();
            syncReviewAudit(true);
            const blob = new Blob([JSON.stringify(currentDecisions, null, 2)], {{ type: 'application/json' }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'decisions.json';
            a.click();
            URL.revokeObjectURL(url);
        }});
    </script>
</body>
</html>
"""

    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    return out_file


def _verify_no_sample_text_leaks(target_dict: Any, forbidden_snippets: List[str]) -> None:
    """递归核对生成的格式包字典，确保绝无任何样本文本泄露"""
    if isinstance(target_dict, str):
        for snippet in forbidden_snippets:
            if len(snippet) >= 6 and snippet in target_dict:
                raise ConfigError(f"安全违规：在导出的格式包中检测到样本文本片段泄漏: '{snippet}'")
    elif isinstance(target_dict, dict):
        for v in target_dict.values():
            _verify_no_sample_text_leaks(v, forbidden_snippets)
    elif isinstance(target_dict, list):
        for item in target_dict:
            _verify_no_sample_text_leaks(item, forbidden_snippets)


def compile_format_package(
    analysis_data: Dict[str, Any],
    decisions_data: Dict[str, Any],
    base_format_ref: Optional[str] = None,
    format_id: str = "custom-format",
    format_version: str = "1.0.0",
    output_path: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """
    结合分析报告与用户决策，编译生成符合 schemas/format-v1.schema.json 的标准格式包。
    支持继承 base_format_ref，零样本正文泄露。
    """
    # 1. Validate the report and normalize the one accepted decisions shape at
    # the boundary.  Legacy files without the version marker are read-only
    # compatibility input; emitted decisions and packages are always v1.
    try:
        analysis_data = validate_analysis_data(analysis_data)
        decisions_data = require_final_decisions(
            decisions_data,
            analysis_data,
            base_format_ref=base_format_ref,
        )
    except ContractError as exc:
        raise ConfigError(str(exc)) from exc

    # 2. 严格校验报告身份与源指纹匹配
    if analysis_data.get("report_id") != decisions_data.get("report_id"):
        raise ConfigError(
            f"决策文件与分析报告不匹配: report_id {decisions_data.get('report_id')} != {analysis_data.get('report_id')}"
        )
    if analysis_data.get("source_sha256") != decisions_data.get("source_sha256"):
        raise ConfigError("决策文件引用的源文件 SHA-256 哈希与分析报告不一致，拒绝编译。")

    # 3. 索引可用聚类样式
    cluster_map = {c["cluster_id"]: c for c in analysis_data.get("clusters", [])}

    # 4. 收集角色到样式的分配
    role_styles = decisions_data.get("role_styles") or {r: cand["cluster_id"] for r, cand in analysis_data.get("candidate_roles", {}).items()}
    style_overrides = decisions_data.get("style_overrides", {})
    missing_decisions = decisions_data.get("missing_roles", {})

    compiled_styles: Dict[str, Any] = {}
    compiled_roles: Dict[str, Any] = {}

    for role_name, cid in role_styles.items():
        cluster = cluster_map.get(cid)
        if not cluster:
            raise ConfigError(f"角色 {role_name} 引用了不存在的候选样式 ID: {cid}")

        style_name = f"style_{role_name.replace('.', '_')}"
        r_style = copy.deepcopy(cluster.get("run_style", {}))
        p_style = copy.deepcopy(cluster.get("paragraph_style", {}))

        # 应用微调覆写
        if role_name in style_overrides:
            ov = style_overrides[role_name]
            r_style.update(ov.get("run", {}))
            p_style.update(ov.get("paragraph", {}))

        # 构造有效 StyleDefinition 字典
        s_def: Dict[str, Any] = {}
        if r_style:
            s_def["run"] = r_style
        if p_style:
            s_def["paragraph"] = p_style

        compiled_styles[style_name] = s_def

        # 角色规格
        outline_lvl = None
        include_in_toc = False
        if role_name.startswith("heading."):
            outline_lvl = int(role_name.split(".")[1])
            include_in_toc = outline_lvl <= 3
        elif role_name == "title":
            outline_lvl = None
            include_in_toc = False

        r_spec: Dict[str, Any] = {"style": style_name}
        if outline_lvl:
            r_spec["outline_level"] = outline_lvl
        if include_in_toc:
            r_spec["include_in_toc"] = True

        compiled_roles[role_name] = r_spec

    # 5. 如果声明了 base_format_ref，解析基础包并作为兜底继承
    base_pkg_dict: Dict[str, Any] = {}
    if base_format_ref:
        base_resolved = resolve_format_package(base_format_ref)
        # 将 base_resolved 转换为可用字典
        for r_name, r_spec in base_resolved.roles.items():
            if r_name not in compiled_roles:
                compiled_roles[r_name] = {
                    "style": r_spec.style,
                    "include_in_toc": r_spec.include_in_toc,
                }
                if r_spec.outline_level:
                    compiled_roles[r_name]["outline_level"] = r_spec.outline_level
                if r_spec.style in base_resolved.styles and r_spec.style not in compiled_styles:
                    s_orig = base_resolved.styles[r_spec.style]
                    p_dict: Dict[str, Any] = {}
                    if s_orig.paragraph and s_orig.paragraph.alignment:
                        p_dict["alignment"] = s_orig.paragraph.alignment
                    else:
                        p_dict["alignment"] = "left"
                    compiled_styles[r_spec.style] = {
                        "run": {
                            "east_asia": s_orig.run.east_asia if (s_orig.run and s_orig.run.east_asia) else "宋体",
                            "latin": s_orig.run.latin if (s_orig.run and s_orig.run.latin) else "Times New Roman",
                            "size_pt": s_orig.run.size_pt if (s_orig.run and s_orig.run.size_pt) else 12.0,
                        },
                        "paragraph": p_dict,
                    }

    # 6. 组装完整 format package 字典
    package: Dict[str, Any] = {
        "format_schema_version": 1,
        "id": format_id,
        "version": format_version,
        "metadata": {
            "description": f"Compiled from {analysis_data.get('source_file')} via format-analysis",
            "author": "Document Synthesis Format Compiler",
            "year": "2026",
        },
        "page": analysis_data.get("page", {
            "width_mm": 210.0,
            "height_mm": 297.0,
            "orientation": "portrait",
            "margin_top_mm": 25.4,
            "margin_bottom_mm": 25.4,
            "margin_left_mm": 31.8,
            "margin_right_mm": 31.8,
            "header_distance_mm": 15.0,
            "footer_distance_mm": 15.0,
            "snap_to_grid": True,
        }),
        "styles": compiled_styles,
        "roles": compiled_roles,
        "toc": {
            "title": "目  录",
            "title_role": "toc.title",
            "max_level": 3,
            "leader": "dots",
            "page_number_gap_mm": 5.0,
        },
        "header": {
            "mode": "none",
        },
        "footer": {
            "mode": "managed",
            "format": "dash_number",
            "alignment": "center",
        },
    }

    if base_format_ref:
        package["extends"] = base_format_ref

    # 7. Schema 验证
    schema = load_format_schema()
    try:
        jsonschema.validate(instance=package, schema=schema)
    except jsonschema.ValidationError as err:
        raise ConfigError(f"编译出的格式包未通过模式校验: {err.message} (路径: {err.json_path})") from err

    # 8. 零样本正文泄露严格断言
    snippets = []
    for c in analysis_data.get("clusters", []):
        snippets.extend(c.get("sample_texts", []))
    _verify_no_sample_text_leaks(package, snippets)

    # 9. 保存输出
    if output_path:
        out_p = Path(output_path).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)
        with open(out_p, "w", encoding="utf-8") as f:
            json.dump(package, f, ensure_ascii=False, indent=2)

    return package


def compile_role_mapping(
    analysis_data: Dict[str, Any],
    decisions_data: Dict[str, Any],
    output_path: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """
    将分析出的文档节点与最终确认角色编译为绑定的 role-map.json
    """
    try:
        analysis_data = validate_analysis_data(analysis_data)
        decisions_data = require_final_decisions(
            decisions_data,
            analysis_data,
            require_role_styles=False,
        )
    except ContractError as exc:
        raise ConfigError(str(exc)) from exc

    if analysis_data.get("report_id") != decisions_data.get("report_id"):
        raise ConfigError("决策文件与分析报告不匹配。")
    if analysis_data.get("source_sha256") != decisions_data.get("source_sha256"):
        raise ConfigError("决策文件与分析报告的源材料 SHA-256 哈希不一致。")

    confirmed_node_roles = dict(analysis_data.get("node_roles", {}))
    # 允许决策中手工指定节点角色覆写
    if "node_roles" in decisions_data:
        unknown_paths = sorted(set(decisions_data["node_roles"]) - set(analysis_data.get("node_refs", {})))
        if unknown_paths:
            raise ConfigError(
                "决策引用了分析报告不存在的节点: " + ", ".join(unknown_paths)
            )
        confirmed_node_roles.update(decisions_data["node_roles"])

    node_refs = analysis_data.get("node_refs", {})
    role_assignments = []
    explicit_paths = set(decisions_data.get("node_roles", {}))
    for element_path, role in confirmed_node_roles.items():
        node_ref = node_refs.get(element_path)
        if not node_ref:
            raise ConfigError(f"分析报告缺少节点的完整 NodeRef: {element_path}")
        role_assignments.append(RoleAssignment(
            node_ref=NodeRef(
                source_sha256=node_ref["source_sha256"],
                part_uri=node_ref["part_uri"],
                element_path=node_ref["element_path"],
            ),
            role=role,
            provenance="explicit" if element_path in explicit_paths else "analysis",
            confirmed=element_path in explicit_paths,
            text_hash=node_ref["text_hash"],
        ))

    on_unmapped = decisions_data.get("on_unmapped", "error")
    # Pre-R2 decisions used the fallback role name here. It is not a valid
    # policy, so normalize it to the safe, explicit gate behaviour.
    if on_unmapped == "body":
        on_unmapped = "error"
    mapper = serialize_role_map(
        role_assignments,
        source_file=analysis_data.get("source_file"),
        analyzer_version=analysis_data.get("analyzer_version", "1.0.0"),
        on_unmapped=on_unmapped,
    )
    mapping: Dict[str, Any] = mapper

    try:
        validate_role_map_data(mapping)
    except ContractError as exc:
        raise ConfigError(str(exc)) from exc

    if output_path:
        out_p = Path(output_path).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)
        with open(out_p, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)

    return mapping
