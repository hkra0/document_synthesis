#!/usr/bin/env python3
"""Write the auditable N0–N10 status sidecars from current evidence."""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]
OUT = Path(__file__).resolve().parent

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from docs.acceptance.tools.sanitize_paths import relativize, sanitize_text


def digest(path: Path) -> str:
    if not path.is_file():
        return ""
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def status(
    batch: str,
    state: str,
    closes: list[str],
    commands: list[str],
    tests: dict,
    code: list[str],
    fixtures: list[str],
    artifacts: list[str],
    remaining: list[str],
    word_cases: list[dict] | None = None,
) -> dict:
    return {
        "schema_version": 1,
        "batch": batch,
        "status": state,
        "verified_at": "2026-09-06",
        "closes": closes,
        "code_sha256": {str(path.relative_to(ROOT)): digest(path) for path in map(Path, code)},
        "fixture_sha256": {str(path.relative_to(ROOT)): digest(path) for path in map(Path, fixtures)},
        "commands": [sanitize_text(cmd, ROOT) for cmd in commands],
        "tests": tests,
        "word_cases": word_cases or [],
        "artifact_sha256": {str(path.relative_to(ROOT)): digest(path) for path in map(Path, artifacts)},
        "remaining_limits": remaining,
    }


def main() -> int:
    contracts = ROOT / "tests" / "test_post_review_contracts.py"
    word = ROOT / "tests" / "test_post_review_word.py"
    format_analysis = ROOT / "lib" / "format_analysis.py"
    candidates = ROOT / "lib" / "role_candidates.py"
    content = ROOT / "lib" / "content_integrity.py"
    contracts_module = ROOT / "lib" / "contracts.py"
    fields = ROOT / "lib" / "field_updater.py"
    delivery = ROOT / "lib" / "delivery.py"
    migration = ROOT / "lib" / "manifest_migration.py"
    config = ROOT / "lib" / "config.py"
    composition = ROOT / "lib" / "composition.py"
    build_plan = ROOT / "lib" / "build_plan.py"
    engine = ROOT / "lib" / "engine.py"
    verification_contracts = ROOT / "lib" / "verification_contracts.py"
    diagnostics = OUT / "pdf-diagnostics" / "summary.json"
    metrics = ROOT / "docs" / "acceptance" / "r7-evaluation" / "unstructured-metrics.json"
    structured_metrics = ROOT / "docs" / "acceptance" / "r7-evaluation" / "metrics.json"
    unstructured_evaluator = ROOT / "docs" / "acceptance" / "r7-evaluation" / "evaluate_unstructured_dataset.py"
    unstructured_generator = ROOT / "docs" / "acceptance" / "r7-evaluation" / "generate_unstructured_dataset.py"
    unstructured_baseline_writer = ROOT / "docs" / "acceptance" / "r7-evaluation" / "write_unstructured_baseline.py"
    unstructured_baseline = ROOT / "docs" / "acceptance" / "r7-evaluation" / "unstructured-human-baseline.json"
    format_review = ROOT / "lib" / "format_review.py"
    docx_inspector = ROOT / "lib" / "docx_inspector.py"
    smoke = ROOT / "smoke_test.py"
    manual_correction_smoke = ROOT / "docs" / "acceptance" / "r7-evaluation" / "manual-correction-smoke.json"
    offline = OUT / "offline-verification.json"
    common = [str(path) for path in (contracts,)]
    word_cases = [
        {"id": f"NW{index:02d}", "status": "passed", "evidence": "真实 Word NW01–NW12 固定 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR 矩阵；12/12 通过，无新授权弹窗。"}
        for index in range(1, 13)
    ]
    batches = {
        "N0": status("N0", "verified", ["C01", "C02", "C03", "C04", "C05", "C06"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
        ], {"run": 26, "failed": 0, "errors": 0, "skipped": 0}, [contracts], [], [contracts], [
            "C08 的最终发布/Word 证据归入 N8，当前尚未关闭。",
        ]),
        "N1": status("N1", "verified", ["C01"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> TMPDIR=<word-access-dir>/document-synthesis-nwdebug.K6jPYu DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v",
        ], {"run": 3, "failed": 0, "errors": 0, "skipped": 0}, [fields, delivery], [], [fields, delivery], [
            "固定 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR 后 NW01–NW12 矩阵 12/12 通过；随机临时目录版本的一次 NW02 超时作为历史观察保留。",
        ], [
            {"id": "NW01", "status": "passed", "evidence": "固定目录真实 Word NW01–NW12 矩阵"},
            {"id": "NW02", "status": "passed", "evidence": "固定目录真实 Word 矩阵；12/12 通过"},
            {"id": "NW10", "status": "passed", "evidence": "固定目录真实 Word NW01–NW12 矩阵"},
        ]),
        "N2": status("N2", "verified", ["C02"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
            "python3 synthesize.py --source examples/custom-format-demo/manuscript.docx --manifest examples/custom-format-demo/manifest.json --plan",
        ], {"run": 7, "failed": 0, "errors": 0, "skipped": 0}, [config, composition, build_plan, verification_contracts], [ROOT / "examples" / "custom-format-demo" / "academic-demo.format.json"], [contracts], [
            "当前证据限于 NW03 匿名夹具，不扩展为任意模板封面支持。",
        ], [{"id": "NW03", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"}]),
        "N3": status("N3", "verified", ["C03", "C04"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
        ], {"run": 6, "failed": 0, "errors": 0, "skipped": 0}, [build_plan, engine, content, verification_contracts], [], [contracts], [
            "当前证据覆盖本轮匿名夹具；不扩展为所有复杂 story 的无条件映射支持。",
        ], [{"id": "NW06", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"}]),
        "N4": status("N4", "verified", ["C03", "C04"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
        ], {"run": 4, "failed": 0, "errors": 0, "skipped": 0}, [content, engine, verification_contracts], [], [contracts], [
            "当前证据覆盖 preserve/mixed 的匿名夹具；复杂 story 仍按能力边界管理。",
        ], [
            {"id": "NW04", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"},
            {"id": "NW05", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"},
        ]),
        "N5": status("N5", "verified", ["C05"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
        ], {"run": 6, "failed": 0, "errors": 0, "skipped": 0}, [content, engine, verification_contracts, composition, ROOT / "lib" / "document_parts.py"], [], [contracts], [
            "当前证据覆盖 NW06/NW10 匿名夹具；不扩展为任意复杂 Word 对象组合支持。",
        ], [
            {"id": "NW06", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"},
            {"id": "NW10", "status": "passed", "evidence": "真实 Word NW01–NW12 矩阵"},
        ]),
        "N6": status("N6", "verified", ["C06"], [
            "python3 -m unittest discover -s tests -p 'test_post_review_contracts.py' -q",
        ], {"run": 1, "failed": 0, "errors": 0, "skipped": 0}, [migration], [], [contracts], []),
        "N7": status("N7", "verified", ["C07"], [
            "python3 docs/acceptance/r7-evaluation/generate_unstructured_dataset.py",
            "python3 docs/acceptance/r7-evaluation/write_unstructured_baseline.py",
            "python3 docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py",
            "python3 docs/acceptance/r7-evaluation/evaluate_dataset.py --output docs/acceptance/r7-evaluation/metrics.json",
            "python3 -W error::FutureWarning -m unittest tests.test_format_analysis tests.test_r7_workflows tests.test_format_review -q",
        ], {"run": 17, "failed": 0, "errors": 0, "skipped": 0}, [format_analysis, candidates, format_review, contracts_module, unstructured_evaluator, unstructured_generator, unstructured_baseline_writer, ROOT / "tests" / "test_format_analysis.py", ROOT / "tests" / "test_format_review.py", ROOT / "tests" / "test_r7_workflows.py"], [ROOT / "docs" / "acceptance" / "r7-evaluation" / "unstructured-dataset.json", unstructured_baseline], [metrics, structured_metrics, manual_correction_smoke], [
            "冻结匿名无大纲集包含 20 份文档、100 个标题正例和 500 个正文负例；标题候选覆盖率 100%、候选块 precision 1.0、自动接受覆盖率 100%、自动接受 precision 1.0，正文误报 0，已达到 98% precision/50% coverage 门槛。",
            "冻结标签由独立人工可见段落基线生成器绑定并声明 algorithm_output_used=false；手工决策→格式包→目标 RoleMap→生产 --plan 的 workflow smoke 也已保留。结论仅适用于该匿名冻结集，不外推为任意真实文档的准确率保证。",
        ]),
        "N8": status("N8", "verified", ["C08"], [
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> TMPDIR=<word-access-dir>/document-synthesis-nwdebug.K6jPYu DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v",
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> TMPDIR=<word-access-dir>/document-synthesis-nwdebug.K6jPYu DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest tests.test_post_review_word.PostReviewWordMatrixTest.test_NW02_toc_and_reference_document_use_final_labels -v",
        ], {"full_matrix": {"run": 12, "passed": 12, "failed": 0, "errors": 0, "skipped": 0, "duration_seconds": 72.668, "prior_run": {"run": 12, "passed": 11, "failed": 1, "errors": 0, "skipped": 0, "failure": "NW02 的一次 Microsoft Word→PDF 导出超时；固定 Word Automation 目录后复跑通过。"}}, "NW02_recheck": {"run": 1, "passed": 1, "failed": 0, "errors": 0, "skipped": 0, "duration_seconds": 296.557}}, [word, engine, delivery, ROOT / "lib" / "qa.py", ROOT / "lib" / "pagination.py", ROOT / "lib" / "docx_inspector.py", ROOT / "lib" / "role_mapper.py", ROOT / "lib" / "format_analysis.py", ROOT / "lib" / "composition.py", ROOT / "lib" / "scanner.py", ROOT / "lib" / "content_integrity.py", ROOT / "lib" / "verification_contracts.py", ROOT / "lib" / "renderers.py"], [ROOT / "docs" / "acceptance" / "word-followup" / "m2-fixtures" / "numbering-source.docx"], [ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "nw09-notes.docx", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "nw09-qa-main.pdf", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-source.docx", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "nw09-build-metadata.json", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-qa-main.pdf"], [
            "固定 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR 后完整矩阵 12/12 通过；随机临时目录版本的一次 NW02 导出超时仅作为历史观察保留。",
            "PDF 解析器的 Quartz wrong pointing object 警告由 N9 单独诊断；N8 不扩大为所有 PDF 结构的无条件兼容声明。",
        ], word_cases),
        "N9": status("N9", "verified", [], [
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 -c 'from lib.qa import export_docx_to_pdf; export_docx_to_pdf(<anonymous-source>, <rerun-pdf>)'",
            "python3 docs/acceptance/post-review/pdf-diagnostics/diagnose_pdfs.py",
            "qpdf --check <rerun-pdf>",
        ], {"historical_diagnostics": {"run": 1, "failed": 0, "errors": 0, "skipped": 0, "warning_classification": "quartz_orphan_zero_offset_xref"}, "same_source_word_export": {"run": 1, "passed": 1, "failed": 0, "errors": 0, "skipped": 0, "pypdf_pages": 6, "pymupdf_pages": 6, "pypdf_warning_count": 3, "warning_classification": "quartz_orphan_zero_offset_xref", "orphan_xref_objects": [6, 8, 10]}}, [ROOT / "docs" / "acceptance" / "post-review" / "pdf-diagnostics" / "diagnose_pdfs.py", ROOT / "docs" / "acceptance" / "post-review" / "pdf-diagnostics" / "README.md"], [ROOT / "docs" / "acceptance" / "r0-r11-review" / "thesis-complete.pdf", ROOT / "docs" / "acceptance" / "r0-r11-review" / "body-pageref.pdf", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-source.docx"], [diagnostics, ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-qa-main.pdf", ROOT / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-qa-main-rerun.pdf"], [
            "已知 Quartz 模式的处理是非破坏性诊断分类，不改写原始 PDF；仅当 offset=0 对象不存在、无任何引用且 pypdf/PyMuPDF/Poppler 页数与页面框一致时视为非致命。其他结构异常仍保持 blocked。",
        ]),
        "N10": status("N10", "verified", [], [
            "python3 synthesize.py --analyze-format examples/custom-format-demo/sample.docx --analysis-dir <fresh-dir>/sample --replace-output",
            "python3 synthesize.py --compile-format <fresh-dir>/sample/analysis.json --decisions examples/custom-format-demo/decisions.final.json --format-out <fresh-dir>/academic-demo.format.json --replace-output",
            "python3 synthesize.py --analyze-source examples/custom-format-demo/manuscript.docx --analysis-dir <fresh-dir>/target --replace-output",
            "python3 synthesize.py --compile-mapping <fresh-dir>/target/analysis.json --decisions examples/custom-format-demo/manuscript-decisions.final.json --mapping-out <fresh-dir>/manuscript-roles.json --replace-output",
            "python3 synthesize.py --render-preview <fresh-dir>/academic-demo.format.json --preview-out <fresh-dir>/preview.html --preview-meta <fresh-dir>/preview.json --replace-output",
            "python3 synthesize.py --migrate-manifest examples/custom-format-demo/legacy-manifest-v1.json --manifest-out <fresh-dir>/manifest.v2.json --replace-output",
            "python3 synthesize.py --source examples/custom-format-demo/manuscript.docx --manifest examples/custom-format-demo/manifest.json --plan",
            "python3 docs/acceptance/post-review/pdf-diagnostics/diagnose_pdfs.py",
            "python3 -m unittest discover -s tests -p 'test_smoke_outputs.py' -q",
            "python3 -m venv <fresh-venv>",
            "<fresh-venv>/bin/python -m pip install -r requirements.txt",
            "<fresh-venv>/bin/python -W error::FutureWarning -m unittest discover -s tests -q",
            "python3 -W error::FutureWarning -m unittest discover -s tests -q",
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> TMPDIR=<word-access-dir>/document-synthesis-nwdebug.K6jPYu DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest discover -s tests -p 'test_post_review_word.py' -v",
            "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> TMPDIR=<word-access-dir>/document-synthesis-nwdebug.K6jPYu DOCUMENT_SYNTHESIS_WORD_TEST=1 python3 -m unittest tests.test_post_review_word.PostReviewWordMatrixTest.test_NW02_toc_and_reference_document_use_final_labels -v",
            "git diff --check",
        ], {"run": 286, "failed": 0, "errors": 0, "skipped": 3, "clean_venv": {"run": 286, "failed": 0, "errors": 0, "skipped": 3}}, [ROOT / "README.md", ROOT / "DEVELOPMENT.md", ROOT / "docs" / "formatting-progress.md", format_analysis, candidates, contracts_module, build_plan, engine, fields, content, verification_contracts, docx_inspector, format_review, smoke, composition, ROOT / "lib" / "document_parts.py", word, OUT / "write_status.py", ROOT / "docs" / "acceptance" / "post-review" / "pdf-diagnostics" / "diagnose_pdfs.py"], [ROOT / "examples" / "custom-format-demo" / "academic-demo.format.json", ROOT / "examples" / "custom-format-demo" / "manuscript-roles.json"], [diagnostics, metrics, offline, manual_correction_smoke], [
            "离线 CLI 链路、真实 Word NW01–NW12 证据、固定目录授权机制与本轮文档状态已同步。",
            "N7 冻结集自动识别门槛已通过；结论限定于独立匿名冻结集，不扩大为任意文档的通用准确率声明。N9 已将可重复的 Quartz orphan xref 模式分类为非致命。",
        ], [{"id": "NW01–NW12", "status": "passed", "evidence": "acceptance/post-review/n8-status.json"}]),
    }
    OUT.mkdir(parents=True, exist_ok=True)
    for batch, payload in batches.items():
        (OUT / f"{batch.lower()}-status.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    (OUT / "status.json").write_text(
        json.dumps({
            "schema_version": 1,
            "plan": "docs/post-r0-r11-optimization-plan.md",
            "generated_at": "2026-09-06",
            "status": "verified",
            "summary": "N0–N10 的当前验收证据已齐备；N7 独立匿名冻结集达到自动接受门槛，N8 真实 Word 矩阵与 N9 PDF 警告诊断已验证。",
            "batches": {key: value["status"] for key, value in batches.items()},
            "execution_policy": {
                "automation": "real_word_used",
                "word_claims": "verified_for_N8_and_N9",
            },
            "blocking_conditions": [],
            "next_step": "当前冻结验收范围已放行；后续扩大真实文档或学校规范范围时，新增独立冻结集并重新校准 N7，继续使用固定 DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR 运行 Word 验收。",
        }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
