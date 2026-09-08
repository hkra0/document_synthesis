#!/usr/bin/env python3
"""Write current S0–S8 remediation evidence sidecars.

The sidecars intentionally distinguish verified sub-ranges from partial
batches.  They are generated from repository files and recorded commands; a
test name alone is not promoted to a verified scenario.
"""

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


def digest(path: str) -> str:
    target = ROOT / path
    if not target.is_file():
        return ""
    h = hashlib.sha256()
    with target.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def sidecar(batch: str, state: str, closes: list[str], commands: list[str], tests: dict,
            files: list[str], evidence: list[str], remaining: list[str]) -> dict:
    return {
        "schema_version": 1,
        "batch": batch,
        "plan": "docs/post-n0-n10-remediation-plan.md",
        "status": state,
        "verified_at": "2026-09-08",
        "closes": closes,
        "commands": [sanitize_text(cmd, ROOT) for cmd in commands],
        "tests": tests,
        "code_and_test_sha256": {relativize(path, ROOT).replace("<repo>/", ""): digest(path) for path in files},
        "evidence": evidence,
        "word_access_directory": relativize("/private/tmp/document-synthesis-word-access", ROOT),
        "remaining_limits": remaining,
    }


def main() -> int:
    common_offline = [
        "python3 -m unittest tests.test_role_evaluation tests.test_remediation_contracts -q",
        "python3 -m unittest tests.test_content_integrity tests.test_docx_inspector -q",
    ]
    word_prefix = (
        "DOCUMENT_SYNTHESIS_WORD_TEST=1 "
        f"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR={relativize('/private/tmp/document-synthesis-word-access', ROOT)} "
        "python3 -m unittest"
    )
    batches = {
        "S0": sidecar(
            "S0", "verified", ["D01", "D02", "D03", "D04", "D05", "D06"],
            common_offline,
            {"offline_contract_tests": {"run": 23, "passed": 23, "failed": 0, "errors": 0}},
            [
                "tests/test_remediation_contracts.py",
                "tests/test_role_evaluation.py",
                "docs/acceptance/remediation/scenario-contracts.json",
            ],
            ["D01–D06 的历史探针结果已与当前成对测试逐项对照；scenario-contracts 的 verified 状态由历史失败观测、当前测试命令和真实 Word sidecar 共同支撑，不由测试名称自动推导。"],
            [],
        ),
        "S1": sidecar(
            "S1", "verified", ["D01"],
            common_offline + [
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW01_real_word_preserves_logical_inline_spans_and_pdf_text -v",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW01_real_word_corrupted_logical_span_is_rejected_by_final_gate -v",
            ],
            {"offline": {"passed": True}, "SW01_real": {"run": 2, "passed": 2, "failed": 0}},
            ["lib/docx_inspector.py", "lib/content_integrity.py", "lib/verification_contracts.py", "tests/test_remediation_contracts.py", "tests/test_remediation_word.py"],
            ["S1-T01–T07 离线契约通过；SW01 真实 Word→PDF 正常对照通过；同一真实 CLI 成果损坏第二逻辑区间后由最终格式门禁以字号诊断拒绝。"],
            [],
        ),
        "S2": sidecar(
            "S2", "verified", ["D02"],
            [
                "python3 -m unittest tests.test_docx_inspector tests.test_remediation_contracts -q",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW02_real_word_preserves_default_first_even_story_bindings -v",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW02_real_word_two_section_shared_and_unlinked_bindings -v",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW02_real_word_wrong_story_relationship_is_rejected -v",
            ],
            {"offline": {"passed": True}, "story_table_inspection": {"run": 1, "passed": 1, "failed": 0}, "SW02_real": {"run": 3, "passed": 3, "failed": 0}},
            ["lib/docx_inspector.py", "lib/verification_contracts.py", "lib/content_integrity.py", "tests/test_docx_inspector.py", "tests/test_remediation_contracts.py", "tests/test_remediation_word.py"],
            ["story part、显式关系、linked_to_previous、different_first_page、文字/距离/有效字符属性和页眉表格边界均有断言；三项真实 SW02 夹具覆盖共享/解除链接及错误关系目标，均通过 Word→PDF/最终门禁。"],
            [],
        ),
        "S3": sidecar(
            "S3", "verified", ["D03"],
            [
                "python3 -m unittest tests.test_content_integrity tests.test_remediation_contracts -q",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW05_real_word_multi_source_notes_relationship_matrix_is_closed -v",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW06_real_word_object_instances_and_extra_image_are_checked -v",
                f"{word_prefix} tests.test_remediation_word.RemediationWordTest.test_SW06_real_word_generated_cover_image_has_finite_license -v",
            ],
            {"offline": {"passed": True}, "object_mapping_and_license": {"run": 3, "passed": 3, "failed": 0}},
            ["lib/content_integrity.py", "tests/test_content_integrity.py", "tests/test_remediation_contracts.py", "tests/test_remediation_word.py"],
            ["ObjectOccurrence 记录来源哈希、part/story、宿主与局部序号；多来源 SW05 真实交付验证 4 个 source occurrence→output anchor 映射；SW06 真实交付验证正文/封面同资源分部件映射及有限次数许可；重复、缺失、同名宿主移动和额外图片均有负向断言。"],
            [],
        ),
        "S4": sidecar(
            "S4", "verified", ["D05"],
            ["python3 -m unittest tests.test_role_evaluation -v", "python3 docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py", "python3 docs/acceptance/r7-evaluation/evaluate_dataset.py --output docs/acceptance/r7-evaluation/metrics.json"],
            {"pure_metrics": {"run": 6, "passed": 6, "failed": 0}, "metrics_schema_version": 2, "manual_correction_count": None},
            ["lib/role_evaluation.py", "tests/test_role_evaluation.py", "docs/acceptance/r7-evaluation/evaluate_unstructured_dataset.py", "docs/acceptance/r7-evaluation/evaluate_dataset.py", "docs/acceptance/r7-evaluation/metrics.json", "docs/acceptance/r7-evaluation/unstructured-metrics.json"],
            ["一对一错手算 exact precision=0.5；正文误报进入 detection 分母；无真实点击日志时 manual_correction_count=null。"],
            ["当前结论限定在匿名冻结集，未声称任意真实模板准确率。"],
        ),
        "S5": sidecar(
            "S5", "verified", ["D04"],
            [
                "python3 docs/acceptance/r7-evaluation/generate_hard_cases.py",
                "python3 synthesize.py --analyze-format docs/acceptance/r7-evaluation/unstructured-hard-samples/公文-hard-cases.docx --analysis-dir <fresh-review-dir> --replace-output",
                "node -e \"extract <fresh-review-dir>/review.html script and parse it with vm.Script\"",
                "python3 -m unittest tests.test_role_candidates tests.test_r7_workflows tests.test_remediation_contracts tests.test_format_analysis tests.test_format_review -q",
                "python3 synthesize.py --compile-format <fresh-review-dir>/analysis.json --decisions '<downloads>/decisions (2).json' --format-base preset:academic-basic@1.0.0 --format-id r7-human-reviewed-format --format-version 1.0.0 --format-out <fresh-review-dir>/attested-output/r7-human-reviewed-format.json --replace-output",
                "python3 synthesize.py --compile-mapping <fresh-review-dir>/analysis.json --decisions '<downloads>/decisions (2).json' --mapping-out <fresh-review-dir>/attested-output/r7-human-reviewed.roles.json --replace-output",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 synthesize.py --doctor --source docs/acceptance/r7-evaluation/unstructured-hard-samples/报告-hard-cases.docx --manifest <fresh-review-dir>/target-review/manifest.json",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 synthesize.py --source docs/acceptance/r7-evaluation/unstructured-hard-samples/报告-hard-cases.docx --manifest <fresh-review-dir>/target-review/manifest.json --output-dir <fresh-review-dir>/target-release",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 smoke_test.py --source docs/acceptance/r7-evaluation/unstructured-hard-samples/报告-hard-cases.docx --manifest <fresh-review-dir>/target-review/manifest.json --output-dir <fresh-review-dir>/target-release",
            ],
            {
                "offline": {"passed": True},
                "frozen_baseline": {
                    "documents": 20,
                    "categories": 4,
                    "documents_per_category": 5,
                    "heading_positives": 100,
                    "body_negatives": 500,
                },
                "hard_case_fixtures": {"documents": 4, "case_types": 6},
                "manual_reuse_workflow": {"run": 1, "passed": 1},
                "human_review": {"run": 1, "passed": 1, "reviewer_attested": True, "operation_count": 2},
                "target_real_word_delivery": {"run": 1, "passed": 1, "pages": 2, "verified_blocks": 12},
            },
            [
                "lib/role_candidates.py",
                "lib/contracts.py",
                "lib/format_review.py",
                "schemas/decisions-v1.schema.json",
                "tests/test_role_candidates.py",
                "tests/test_r7_workflows.py",
                "tests/test_format_review.py",
                "tests/test_remediation_contracts.py",
                "docs/acceptance/r7-evaluation/generate_hard_cases.py",
                "docs/acceptance/r7-evaluation/unstructured-hard-cases.json",
                "docs/acceptance/remediation/s5-human-review-evidence.json",
            ],
            ["主冻结集实际包含四类各 5 份、100 个标题正例和 500 个正文负例；额外四类匿名 DOCX 困难夹具覆盖缺失中间层级、同层异字号、日期/版本号、正文编号、短粗体正文和题注。真实浏览器审阅导出已绑定 report_id/source_sha256，2 个操作的 reviewer_attested=true 文件成功编译格式包；格式包在另一份匿名困难目标上生成 RoleMap，并通过真实 Word build/smoke。完整哈希见 s5-human-review-evidence.json。"],
            ["20 份冻结集的 aggregate manual_correction_count 仍为 null；本轮只证明一份匿名困难样例的真人审阅与跨目标交付闭环，不将单次操作外推为冻结集人工一致性。"],
        ),
        "S6": sidecar(
            "S6", "verified", ["D06"],
            [f"{word_prefix} discover -s tests -p 'test_remediation_word.py' -v", f"{word_prefix} discover -s tests -p 'test_post_review_word.py' -v"],
            {"remediation_word": {"run": 11, "passed": 11, "failed": 0}, "NW01_NW12": {"run": 12, "passed": 12, "failed": 0}},
            ["tests/test_remediation_word.py", "tests/test_post_review_word.py", "lib/qa.py", "lib/engine.py", "lib/renderers.py", "lib/content_integrity.py"],
            ["固定目录真实 Word NW01–NW12=12/12；remediation_word=11/11；SW03 完整论文部件、SW04 横向节故事、SW05 多来源脚注/尾注关系、SW06 封面图片许可和 SW07 缺层级人工确认均已通过 Word→PDF/最终断言；均无授权弹窗。"],
            [],
        ),
        "S7": sidecar(
            "S7", "verified", ["D06"],
            [
                "python3 -m unittest tests.test_delivery_pipeline.DeliveryPipelineTest.test_publish_failure_restores_all_deliveries_and_metadata -v",
                f"{word_prefix} tests.test_publication_gates.PublicationGateContractTest.test_real_cli_second_failure_keeps_previous_delivery_and_metadata -v",
                f"{word_prefix} tests.test_publication_gates.PublicationGateContractTest.test_real_word_gate_injection_after_staging_blocks_publication -v",
            ],
            {
                "rollback_all_artifacts": {"run": 1, "passed": 1},
                "real_cli_second_failure": {"run": 1, "passed": 1},
                "real_gate_injection": {"run": 1, "passed": 1},
            },
            ["lib/engine.py", "lib/delivery.py", "tests/test_delivery_pipeline.py", "tests/test_publication_gates.py", "tests/test_post_review_word.py"],
            ["真实 CLI 首次成功、第二次过期 RoleMap 失败且旧 DOCX/metadata 哈希不变；真实 Word/PDF 完成后的 staging 损坏被生产格式门禁拒绝；发布中途第二替换失败可恢复全部 DOCX 与 metadata。"],
            [],
        ),
        "S8": sidecar(
            "S8", "verified", [],
            [
                "git diff --check",
                "python3 -m unittest discover -s tests -q",
                "python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --plan",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 synthesize.py --doctor --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 synthesize.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --output-dir <word-access-dir>/document-synthesis-release-check-20260908-r2",
                "DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR=<word-access-dir> python3 smoke_test.py --source examples/minimal-demo/source --manifest examples/minimal-demo/manifest.json --output-dir <word-access-dir>/document-synthesis-release-check-20260908-r2",
                f"{word_prefix} discover -s tests -p 'test_post_review_word.py' -v",
                f"{word_prefix} discover -s tests -p 'test_remediation_word.py' -v",
                f"{word_prefix} discover -s tests -p 'test_word_formatting.py' -v",
                f"{word_prefix} discover -s tests -p 'test_m1_acceptance.py' -v",
                f"{word_prefix} discover -s tests -p 'test_thesis_acceptance.py' -v",
                f"{word_prefix} tests.test_pagination.PaginationTest.test_real_word_repeated_titles_and_section_page_restart -v",
                f"{word_prefix} tests.test_publication_gates.PublicationGateContractTest.test_real_cli_second_failure_keeps_previous_delivery_and_metadata -v",
                f"{word_prefix} tests.test_publication_gates.PublicationGateContractTest.test_real_word_gate_injection_after_staging_blocks_publication -v",
            ],
            {
                "offline": {"run": 325, "passed": 325, "failed": 0, "errors": 0, "skipped": 6},
                "real_word": {
                    "NW01_NW12": {"run": 12, "passed": 12},
                    "remediation_word": {"run": 11, "passed": 11},
                    "W04_W07": {"run": 4, "passed": 4},
                    "M1": {"run": 3, "passed": 3},
                    "thesis": {"run": 1, "passed": 1},
                    "pagination": {"run": 1, "passed": 1},
                },
                "status": "verified",
                "reason": "S0–S7 退出证据已固定；S5 真人审阅、跨目标 RoleMap 和真实 Word 交付已闭环",
            },
            ["docs/post-n0-n10-remediation-plan.md", "docs/acceptance/remediation/write_status.py", "docs/acceptance/remediation/s5-human-review-evidence.json"],
            ["离线套件 325/325 通过（6 项条件跳过）；NW01–NW12、remediation_word、W04–W07、M1、论文、分页和发布门禁索引均已固定。最小匿名示例已完成 plan→doctor→真实 Word build→smoke_test；S5 真人审阅与跨目标真实 Word 交付证据见 s5-human-review-evidence.json。"],
            ["能力结论仍限定在声明的匿名夹具、冻结集和 Word 环境；新增模板或更大冻结集需要另行复验。"],
        ),
    }
    for batch, payload in batches.items():
        (OUT / f"{batch.lower()}-status.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    summary = {
        "schema_version": 1,
        "plan": "docs/post-n0-n10-remediation-plan.md",
        "status": "verified",
        "batches": {batch: payload["status"] for batch, payload in batches.items()},
        "next_step": "S0–S8 当前退出条件已满足；后续仅在扩大模板、冻结集或 Word 环境范围时新增复验。",
    }
    (OUT / "status-summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
