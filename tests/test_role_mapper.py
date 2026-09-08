# -*- coding: utf-8 -*-
"""
确定性语义角色映射单元测试 (tests/test_role_mapper.py)
验证 NodeRef 绑定、同名标题消歧、大纲级别映射与校验诊断。
"""

import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from lib.docx_inspector import inspect_docx, NodeRef
from lib.role_mapper import RoleMapper, RoleAssignment, validate_role_map, VALID_ROLES


class RoleMapperTest(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_identical_headings_disambiguated_by_node_ref(self):
        """测试两个完全同名的标题段落通过不同的 NodeRef 映射到不同的书签"""
        doc_path = self.test_dir / "duplicate_headings.docx"
        doc = Document()
        
        # 插入两个同名标题段落，设置相同的样式或大纲级别
        p1 = doc.add_paragraph("第一章 概述")
        p1._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        
        p_body = doc.add_paragraph("正文段落")
        
        p2 = doc.add_paragraph("第一章 概述")
        p2._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        
        doc.save(str(doc_path))

        mapper = RoleMapper()
        assignments, inspection = mapper.map_document(doc_path)

        heading_assignments = [a for a in assignments if a.role == "heading.1"]
        self.assertEqual(len(heading_assignments), 2)
        
        # 标题文字相同
        self.assertEqual(heading_assignments[0].title_text, "第一章 概述")
        self.assertEqual(heading_assignments[1].title_text, "第一章 概述")

        # 但 NodeRef element_path 不同，书签也不同
        self.assertNotEqual(heading_assignments[0].node_ref.element_path, heading_assignments[1].node_ref.element_path)
        self.assertNotEqual(heading_assignments[0].bookmark_name, heading_assignments[1].bookmark_name)
        self.assertTrue(heading_assignments[0].bookmark_name.startswith("_Toc_0001_"))
        self.assertTrue(heading_assignments[1].bookmark_name.startswith("_Toc_0002_"))

        # 校验角色分配合法无冲突
        diags = validate_role_map(assignments, inspection)
        self.assertEqual(len(diags), 0)

    def test_outline_level_to_heading_mapping(self):
        """测试基于 outlineLvl 的 1-9 级标题映射与表格映射"""
        doc_path = self.test_dir / "levels.docx"
        doc = Document()
        
        # 1-3 级
        for lvl in range(1, 4):
            p = doc.add_paragraph(f"标题 {lvl}")
            p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{lvl - 1}"/>'))
        
        # 正文
        doc.add_paragraph("普通正文")

        # 表格
        table = doc.add_table(rows=2, cols=2)
        table.rows[0].cells[0].text = "表头1"
        table.rows[0].cells[1].text = "表头2"
        table.rows[1].cells[0].text = "内容1"
        table.rows[1].cells[1].text = "内容2"

        doc.save(str(doc_path))

        mapper = RoleMapper()
        assignments, inspection = mapper.map_document(doc_path)

        roles = [a.role for a in assignments]
        self.assertIn("heading.1", roles)
        self.assertIn("heading.2", roles)
        self.assertIn("heading.3", roles)
        self.assertIn("body", roles)
        self.assertIn("table.body", roles)

    def test_explicit_rules_override(self):
        """测试用户显式配置的 NodeRef 规则可以覆盖自动推导的角色"""
        doc_path = self.test_dir / "explicit.docx"
        doc = Document()
        p1 = doc.add_paragraph("本应是普通段落")
        p2 = doc.add_paragraph("第二段正文")
        doc.save(str(doc_path))

        # 显式规则指定 p[1] 为 heading.1
        explicit = {
            "/w:document/w:body/w:p[1]": "heading.1"
        }
        mapper = RoleMapper(explicit_rules=explicit)
        assignments, inspection = mapper.map_document(doc_path)

        self.assertEqual(assignments[0].role, "heading.1")
        self.assertEqual(assignments[0].provenance, "explicit")
        self.assertEqual(assignments[0].level, 1)
        self.assertIsNotNone(assignments[0].bookmark_name)

        self.assertEqual(assignments[1].role, "body")
        self.assertEqual(assignments[1].provenance, "default")

    def test_validate_role_map_catches_invalid_and_duplicate(self):
        """测试 validate_role_map 检查出重复书签和虚假 NodeRef"""
        fake_ref1 = NodeRef(source_sha256="abc", part_uri="word/document.xml", element_path="/w:document/w:body/w:p[1]")
        fake_ref2 = NodeRef(source_sha256="abc", part_uri="word/document.xml", element_path="/w:document/w:body/w:p[2]")
        nonexistent_ref = NodeRef(source_sha256="abc", part_uri="word/document.xml", element_path="/w:document/w:body/w:p[999]")

        # 构造包含重复书签的分配
        assignments = [
            RoleAssignment(node_ref=fake_ref1, role="heading.1", level=1, bookmark_name="_Toc_dup"),
            RoleAssignment(node_ref=fake_ref2, role="heading.1", level=1, bookmark_name="_Toc_dup"),
            RoleAssignment(node_ref=nonexistent_ref, role="body"),
        ]

        # 仅校验书签唯一性
        diags = validate_role_map(assignments)
        self.assertEqual(len(diags), 1)
        self.assertIn("重复的书签名称", diags[0].message)

        # 校验非法角色抛出异常
        with self.assertRaises(ValueError):
            RoleAssignment(node_ref=fake_ref1, role="invalid_custom_role")


if __name__ == "__main__":
    unittest.main()
