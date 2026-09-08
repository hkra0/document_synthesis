# -*- coding: utf-8 -*-
"""
样本文档格式分析单元测试 (tests/test_format_analysis.py)
验证样本文档样式聚类、候选语义角色推导、字段级证据链与缺失项分析。
"""

import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor

from lib.format_analysis import analyze_format_sample, ANALYZER_VERSION


class FormatAnalysisTest(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_sample_clustering_and_role_candidates(self):
        """测试对具有鲜明层级的样本文档完成样式聚类与角色候选推导"""
        sample_path = self.test_dir / "sample_doc.docx"
        doc = Document()

        # 1. 标题 (居中, 22pt, 加粗)
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_title = p_title.add_run("机关公文排版标准规范样例")
        r_title.font.name = "方正小标宋简体"
        r_title.font.size = Pt(22)
        r_title.font.bold = True

        # 2. 一级标题 (16pt, 黑体, outlineLvl=0)
        p_h1 = doc.add_paragraph()
        p_h1._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        r_h1 = p_h1.add_run("一、 总体工作原则与指导思想")
        r_h1.font.name = "黑体"
        r_h1.font.size = Pt(16)
        r_h1.font.bold = True

        # 3. 正文 (仿宋, 14pt, 4段)
        for i in range(4):
            p_b = doc.add_paragraph()
            r_b = p_b.add_run(f"这是第 {i+1} 段测试正文内容，严格按照三号仿宋字体与固定行距排版。")
            r_b.font.name = "仿宋_GB2312"
            r_b.font.size = Pt(14)

        # 4. 表格 (1张表)
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].text = "表头A"
        tbl.rows[0].cells[1].text = "表头B"
        tbl.rows[1].cells[0].text = "数据1"
        tbl.rows[1].cells[1].text = "数据2"

        doc.save(str(sample_path))

        analysis = analyze_format_sample(sample_path)

        self.assertEqual(analysis["source_file"], "sample_doc.docx")
        self.assertEqual(analysis["analyzer_version"], ANALYZER_VERSION)
        self.assertEqual(len(analysis["source_sha256"]), 64)

        # 聚类检查
        clusters = analysis["clusters"]
        self.assertGreaterEqual(len(clusters), 3)

        # 正文聚类应出现最多（4段）
        body_cluster = max(clusters, key=lambda c: c["occurrence_count"])
        self.assertEqual(body_cluster["suggested_role"], "body")
        self.assertEqual(body_cluster["occurrence_count"], 4)

        # 候选角色检查
        candidates = analysis["candidate_roles"]
        self.assertIn("body", candidates)
        self.assertIn("heading.1", candidates)
        self.assertIn("title", candidates)
        self.assertIn("table.body", candidates)

        # 标题与大纲级别检查
        self.assertEqual(candidates["heading.1"]["outline_level"], 1)
        self.assertGreaterEqual(candidates["heading.1"]["confidence"], 0.9)

        # 节点映射检查
        node_roles = analysis["node_roles"]
        self.assertGreater(len(node_roles), 5)
        self.assertEqual(node_roles["/w:document/w:body/w:p[1]"], "title")
        self.assertEqual(node_roles["/w:document/w:body/w:p[2]"], "heading.1")

    def test_missing_roles_detection_and_fallbacks(self):
        """测试样本文档中未出现的标准角色能够被识别并提供合理继承回退建议"""
        sample_path = self.test_dir / "minimal_sample.docx"
        doc = Document()
        # 仅有普通正文
        doc.add_paragraph("只有纯正文段落，没有任何标题。")
        doc.save(str(sample_path))

        analysis = analyze_format_sample(sample_path)

        missing_roles = {m["role"]: m for m in analysis["missing_roles"]}
        self.assertIn("heading.1", missing_roles)
        self.assertIn("heading.2", missing_roles)
        self.assertIn("heading.3", missing_roles)
        self.assertIn("suggested_fallback", missing_roles["heading.1"])

    def test_unstructured_candidate_shortlist_is_bounded(self):
        """无大纲样例的每个节点只暴露可操作的前三候选角色。"""
        sample_path = self.test_dir / "unstructured_sample.docx"
        doc = Document()
        doc.add_paragraph("第一部分 工作安排")
        doc.add_paragraph(
            "这是一段足够长的普通正文，用于确认正文段落不会因为字号、邻接关系或默认样式而被扩展成标题候选。"
            "它包含连续的说明文字和多个从句，长度明显超过标题通常应有的范围。"
        )
        doc.add_paragraph("（一）具体执行事项")
        doc.add_paragraph(
            "第二段普通正文继续说明执行条件、责任边界、时间安排和验收口径，保持与正文相同的视觉属性。"
        )
        doc.save(str(sample_path))

        analysis = analyze_format_sample(sample_path)
        evidence = analysis["unstructured_candidate_evidence"]
        self.assertTrue(evidence)
        self.assertLessEqual(
            max(len(item["candidates"]) for item in evidence.values()),
            3,
        )
        self.assertTrue(
            all(
                len({candidate["role"] for candidate in item["candidates"]})
                == len(item["candidates"])
                for item in evidence.values()
            )
        )

    def test_unstructured_numbered_style_family_calibrates_levels(self):
        """重复的无大纲分级样式可自动接受，长正文仍保持 body。"""
        sample_path = self.test_dir / "unstructured_numbered_family.docx"
        doc = Document()
        for level in range(1, 4):
            heading = doc.add_paragraph(f"第{level}部分 无大纲分级标题")
            heading.runs[0].bold = True
            heading.runs[0].font.size = Pt(16 - level)
            body = doc.add_paragraph(
                "这是一段足够长的正文，用于验证局部字号族、分级编号和加粗标题的组合信号，"
                "不会把正文节点误判为标题。"
            )
            body.runs[0].font.size = Pt(12)
        doc.save(str(sample_path))

        evidence = analyze_format_sample(sample_path)["unstructured_candidate_evidence"]
        headings = [evidence[f"/w:document/w:body/w:p[{index}]"] for index in (1, 3, 5)]
        for level, item in enumerate(headings, start=1):
            self.assertEqual(item["selected_role"], f"heading.{level}")
            self.assertGreaterEqual(item["selected_score"], 0.98)
            self.assertEqual(item["score_kind"], "calibrated_confidence_band")
            self.assertTrue(item["candidates"][0]["impact_nodes"])
            self.assertGreaterEqual(item["candidates"][0]["sample_count"], 1)
        for index in (2, 4, 6):
            self.assertEqual(evidence[f"/w:document/w:body/w:p[{index}]"]["selected_role"], "body")


if __name__ == "__main__":
    unittest.main()
