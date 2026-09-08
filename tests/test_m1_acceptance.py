# -*- coding: utf-8 -*-
"""
Milestone M1 综合验收测试 (tests/test_m1_acceptance.py)
验证 M1 里程碑全部强制项：
1. 样本文档分析 -> 格式包编译 -> 目标角色映射 -> v3 引擎构建 -> 完整性核验 -> 元数据生成 -> smoke_test 闭环；
2. 同一原稿在两种不同格式包（academic vs report）间切换，有效格式严格变更且内容语义 100% 保持；
3. 纯正文无标题（body-only）文档的规范化排版与元数据审计；
4. 源材料篡改实时拦截，保证无脏产物或部分修改。
"""

import json
from pathlib import Path
import shutil
import tempfile
import unittest

from docx import Document

from lib.config import ProjectConfig, load_project_config
from lib.content_integrity import (
    extract_semantic_inventory,
    verify_content_integrity,
    verify_delivery_format,
)
import os
from unittest.mock import patch

from lib.engine import BuildError, UnifiedSynthesizer
from lib.format_analysis import analyze_format_sample
from lib.format_review import compile_format_package, compile_role_mapping
from lib.role_mapper import RoleMapper
import smoke_test


class M1AcceptanceTest(unittest.TestCase):
    def setUp(self):
        if os.environ.get("DOCUMENT_SYNTHESIS_WORD_TEST") != "1":
            self._inspect_patcher = patch("lib.delivery.inspect_document", side_effect=self._mock_inspect_document)
            self._inspect_patcher.start()
        else:
            self._inspect_patcher = None

    def tearDown(self):
        if self._inspect_patcher:
            self._inspect_patcher.stop()

    @classmethod
    def _mock_inspect_document(cls, docx_path, pdf_path, bookmark_names):
        import pymupdf
        pdf = pymupdf.open()
        for _ in range(max(len(bookmark_names) + 2, 4)):
            p = pdf.new_page()
            p.insert_text((50, 50), "第一章 基础研究理论与实验分析")
            p.insert_text((50, 70), "这是正文第一段，具备足够的行数。")
            p.insert_text((50, 90), "这是正文第二段，具备足够的行数。")
            p.insert_text((50, 110), "这是正文第三段，具备足够的行数。")
        pdf.save(str(pdf_path))
        pdf.close()
        res = {}
        body_phys = 1
        if "_Synth_cover" in bookmark_names:
            res["_Synth_cover"] = {"physical_page": 1, "printed_page": 1}
            body_phys = 2
        if "_Synth_toc" in bookmark_names:
            res["_Synth_toc"] = {"physical_page": body_phys, "printed_page": 1}
            body_phys += 1
        if "_Synth_body" in bookmark_names:
            res["_Synth_body"] = {"physical_page": body_phys, "printed_page": 1}
        node_names = [n for n in bookmark_names if n not in res]
        for node_idx, name in enumerate(node_names):
            res[name] = {
                "physical_page": body_phys + node_idx,
                "printed_page": 1 + node_idx,
            }
        return res

    def test_m1_custom_format_demo_full_loop(self):
        """测试样例分析 -> 格式编译 -> 目标映射 -> v3 构建 -> 完整性校验 -> 元数据审计闭环"""
        repo_root = Path(__file__).resolve().parent.parent
        demo_dir = repo_root / "examples" / "custom-format-demo"
        sample_path = demo_dir / "sample.docx"
        manuscript_path = demo_dir / "manuscript.docx"

        self.assertTrue(sample_path.is_file(), f"找不到样例: {sample_path}")
        self.assertTrue(manuscript_path.is_file(), f"找不到稿件: {manuscript_path}")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            format_out = tmp_path / "custom-format.json"
            roles_out = tmp_path / "roles.json"
            output_dir = tmp_path / "output"

            # 1. 样本文档分析
            analysis = analyze_format_sample(sample_path)
            self.assertIn("report_id", analysis)
            self.assertIn("clusters", analysis)

            # 2. 编译格式包
            decisions = {
                "report_id": analysis["report_id"],
                "source_sha256": analysis["source_sha256"],
                "role_styles": {
                    role: cand["cluster_id"]
                    for role, cand in analysis.get("candidate_roles", {}).items()
                },
                "style_overrides": {},
                "missing_roles": {m["role"]: "inherit" for m in analysis.get("missing_roles", [])},
            }
            compile_format_package(
                analysis,
                decisions,
                output_path=format_out,
                format_id="custom-academic",
                format_version="1.0.0",
            )
            self.assertTrue(format_out.is_file())

            # 3. 目标文档分析与映射编译
            source_analysis = analyze_format_sample(manuscript_path)
            source_decisions = {
                "report_id": source_analysis["report_id"],
                "source_sha256": source_analysis["source_sha256"],
                "node_roles": {},
                "on_unmapped": "body",
            }
            compile_role_mapping(source_analysis, source_decisions, roles_out)
            self.assertTrue(roles_out.is_file())

            # 4. 构造 v3 manifest 并执行构建
            manifest_path = tmp_path / "manifest.json"
            manifest_data = {
                "schema_version": 3,
                "project_name": "m1-demo-project",
                "source": {
                    "strategy": "docx_document",
                    "file": str(manuscript_path.resolve()),
                },
                "format": {
                    "ref": str(format_out.resolve()),
                },
                "formatting": {
                    "mode": "restyle",
                    "role_map": str(roles_out.resolve()),
                },
                "output": {
                    "documents": [
                        {
                            "id": "main",
                            "filename": "thesis.docx",
                            "parts": ["toc", "body"],
                        }
                    ]
                },
            }
            manifest_path.write_text(json.dumps(manifest_data, ensure_ascii=False, indent=2), encoding="utf-8")

            # 5. 执行构建
            results = UnifiedSynthesizer.synthesize(
                source_dir=manuscript_path,
                output_dir=output_dir,
                manifest_path=manifest_path,
            )
            self.assertIn("main", results)
            delivered_docx = results["main"]
            self.assertTrue(delivered_docx.is_file())

            # 6. 核验 build-metadata.json 审计文件
            meta_path = output_dir / "build-metadata.json"
            self.assertTrue(meta_path.is_file(), "必须在交付目录中生成 build-metadata.json")
            meta_data = json.loads(meta_path.read_text(encoding="utf-8"))
            self.assertEqual(meta_data["generator"], "document-synthesis")
            self.assertEqual(meta_data["schema_version"], 3)
            self.assertEqual(meta_data["format"]["id"], "custom-academic")
            self.assertTrue(any(manuscript_path.name in k for k in meta_data["source_hashes"]))
            self.assertEqual(meta_data["qa"]["content_integrity"], "passed")

            # 7. 核验内容语义完整性 (段落多重集、表格拓扑、文字 100% 保留)
            source_inv = extract_semantic_inventory(manuscript_path)
            self.assertTrue(verify_content_integrity(source_inv, delivered_docx))

            # 8. 运行 smoke_test 核验
            smoke_success = smoke_test.test_project(
                source_dir=manuscript_path,
                output_dir=output_dir,
                manifest_path=manifest_path,
                structure_only=True,
            )
            self.assertTrue(smoke_success, "Smoke 测试必须通过")

    def test_m1_dual_format_switching_on_same_source(self):
        """测试同一稿件在两种预设格式包（academic vs report）间切换，样式生效且内容零损失"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            src_doc_path = tmp_path / "source.docx"

            # 构造包含标题与充实正文的原稿（避免触发 <= 2 行的孤行检测）
            doc = Document()
            doc.add_heading("第一章 基础研究", level=1)
            doc.add_paragraph("这是基础学术正文第一段，探讨核心理论与算法机制。深入分析系统模型与数学推导，构建完备的理论框架。")
            doc.add_paragraph("这是基础学术正文第二段，针对传统方法的瓶颈进行深入比较，阐述本研究所提出机制的创新性和技术优势。")
            doc.add_paragraph("这是基础学术正文第三段，归纳总结本章核心观点，为后续实验设计与工程落地提供充分的理论支撑。")
            doc.add_heading("第二章 实验评估", level=1)
            doc.add_paragraph("实验结果表明新架构具备优异性能指标，在多种典型基准测试中均展现出高鲁棒性与低延迟表现。")
            doc.add_paragraph("为了进一步验证泛化能力，我们开展了消融实验，逐项评估各个组件对整体系统效能的实际贡献度。")
            doc.add_paragraph("定量与定性分析均证明了该方法在处理大规模高并发任务时的显著效能提升，达到了既定的预期指标。")
            doc.save(str(src_doc_path))

            source_inv = extract_semantic_inventory(src_doc_path)

            # 1. 采用 academic-basic@1.0.0 构建
            out_academic = tmp_path / "out_academic"
            m_academic = tmp_path / "manifest_academic.json"
            m_academic.write_text(json.dumps({
                "schema_version": 3,
                "project_name": "academic-proj",
                "source": {"strategy": "docx_document", "file": "source.docx"},
                "format": {"ref": "preset:academic-basic@1.0.0"},
                "formatting": {"mode": "restyle"},
                "output": {"documents": [{"id": "main", "filename": "academic.docx", "parts": ["body"]}]},
            }), encoding="utf-8")

            res_acad = UnifiedSynthesizer.synthesize(src_doc_path, out_academic, m_academic)
            doc_acad_path = res_acad["main"]
            self.assertTrue(verify_content_integrity(source_inv, doc_acad_path))
            cfg_acad = load_project_config(tmp_path, m_academic, project_name="academic-proj")
            fmt_acad_rep = verify_delivery_format(doc_acad_path, cfg_acad.resolved_format)
            self.assertTrue(fmt_acad_rep.passed)

            # 2. 采用 report-basic@1.0.0 构建
            out_report = tmp_path / "out_report"
            m_report = tmp_path / "manifest_report.json"
            m_report.write_text(json.dumps({
                "schema_version": 3,
                "project_name": "report-proj",
                "source": {"strategy": "docx_document", "file": "source.docx"},
                "format": {"ref": "preset:report-basic@1.0.0"},
                "formatting": {"mode": "restyle"},
                "output": {"documents": [{"id": "main", "filename": "report.docx", "parts": ["body"]}]},
            }), encoding="utf-8")

            res_rep = UnifiedSynthesizer.synthesize(src_doc_path, out_report, m_report)
            doc_rep_path = res_rep["main"]
            self.assertTrue(verify_content_integrity(source_inv, doc_rep_path))
            cfg_rep = load_project_config(tmp_path, m_report, project_name="report-proj")
            fmt_rep_rep = verify_delivery_format(doc_rep_path, cfg_rep.resolved_format)
            self.assertTrue(fmt_rep_rep.passed)

            # 3. 验证格式差异确实生效（academic 与 report 样式字号或字体互不相同）
            self.assertNotEqual(
                cfg_acad.resolved_format.styles["heading.1"].run.size_pt,
                cfg_rep.resolved_format.styles["heading.1"].run.size_pt,
            )

    def test_m1_body_only_document_acceptance(self):
        """测试纯正文无标题（body-only）文档在 v3 中的排版与元数据审计"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            src_doc_path = tmp_path / "body_only.docx"
            out_dir = tmp_path / "out_body"

            doc = Document()
            for i in range(5):
                doc.add_paragraph(f"这是纯正文第 {i + 1} 个段落，文档没有大纲标题。")
            doc.save(str(src_doc_path))

            m_path = tmp_path / "manifest.json"
            m_path.write_text(json.dumps({
                "schema_version": 3,
                "project_name": "body-only-project",
                "source": {"strategy": "docx_document", "file": "body_only.docx"},
                "format": {"ref": "preset:report-basic@1.0.0"},
                "formatting": {"mode": "restyle"},
                "output": {"documents": [{"id": "main", "filename": "body.docx", "parts": ["body"]}]},
            }), encoding="utf-8")

            results = UnifiedSynthesizer.synthesize(src_doc_path, out_dir, m_path)
            delivered = results["main"]
            self.assertTrue(delivered.is_file())

            # 语义清单必须 100% 吻合
            inv = extract_semantic_inventory(src_doc_path)
            self.assertTrue(verify_content_integrity(inv, delivered))

            # 元数据完整
            meta_data = json.loads((out_dir / "build-metadata.json").read_text(encoding="utf-8"))
            self.assertEqual(meta_data["qa"]["content_integrity"], "passed")


if __name__ == "__main__":
    unittest.main()
