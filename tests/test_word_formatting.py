# -*- coding: utf-8 -*-
"""独立真实 Word/PDF 验收组。

本文件只在 ``DOCUMENT_SYNTHESIS_WORD_TEST=1`` 且 Word Automation 可用时执行。
环境不满足时显示为 BLOCKED skip，不能把离线替换 Word 的结果计入真实验收。
它与 ``test_thesis_acceptance.py`` 的装配级测试分离，专门覆盖 W04-W06。
"""

import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

import pymupdf
from docx import Document

from lib.pagination import extract_pdf_page_labels
from lib.qa import export_docx_to_pdf, word_automation_status


REPO = Path(__file__).resolve().parents[1]
WORD_FIXTURES = REPO / "docs" / "acceptance" / "word-followup"


class WordFormattingIntegrationTest(unittest.TestCase):
    """W04-W06：只接受生产 CLI 的真实 Word 交付结果。"""

    @classmethod
    def setUpClass(cls):
        if os.environ.get("DOCUMENT_SYNTHESIS_WORD_TEST") != "1":
            raise unittest.SkipTest(
                "BLOCKED: 未设置 DOCUMENT_SYNTHESIS_WORD_TEST=1，未执行真实 Word 验收。"
            )
        available, reason = word_automation_status()
        if not available:
            raise unittest.SkipTest(f"BLOCKED: Word Automation 不可用：{reason}")

    def _run_cli(self, source, manifest, output_dir):
        return subprocess.run(
            [
                sys.executable,
                "synthesize.py",
                "--source",
                str(source),
                "--manifest",
                str(manifest),
                "--output-dir",
                str(output_dir),
            ],
            cwd=str(REPO),
            capture_output=True,
            text=True,
        )

    def test_W04_real_word_publishes_configured_A4_geometry(self):
        """W04 integration with Word：A4 几何必须同时出现在 DOCX 和 PDF。"""
        manifest = WORD_FIXTURES / "manifest.json"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "output"
            result = self._run_cli(WORD_FIXTURES, manifest, output)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            delivered = output / "academic.docx"
            self.assertTrue(delivered.is_file())
            doc = Document(delivered)
            self.assertAlmostEqual(doc.sections[0].page_width.mm, 210.0, delta=0.1)
            self.assertAlmostEqual(doc.sections[0].page_height.mm, 297.0, delta=0.1)

            pdf = Path(tmp) / "academic-verification.pdf"
            self.assertTrue(export_docx_to_pdf(str(delivered), str(pdf)))
            with pymupdf.open(pdf) as rendered:
                self.assertAlmostEqual(rendered[0].rect.width * 25.4 / 72, 210.0, delta=0.1)
                self.assertAlmostEqual(rendered[0].rect.height * 25.4 / 72, 297.0, delta=0.1)

    def test_W05_real_word_verifies_sequences_and_page_level_QA(self):
        """W05 integration with Word：真实标签、分节规则和页级 QA 必须通过。"""
        fixture_dir = WORD_FIXTURES / "m2-fixtures"
        manifest = fixture_dir / "numbering-manifest.json"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "output"
            result = self._run_cli(fixture_dir / "numbering-source.docx", manifest, output)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            delivered = output / "m2.docx"
            self.assertTrue(delivered.is_file())
            pdf = Path(tmp) / "numbering-verification.pdf"
            self.assertTrue(export_docx_to_pdf(str(delivered), str(pdf)))
            labels = extract_pdf_page_labels(pdf)
            self.assertTrue(labels, "PDF 中必须能提取到真实页脚标签")
            self.assertEqual(labels[min(labels)], "i")
            self.assertIn("1", labels.values())

    def test_W06_real_word_converges_PAGEREF_and_publishes(self):
        """W06 integration with Word：PAGEREF 两轮测量必须收敛并发布。"""
        fixture_dir = WORD_FIXTURES / "m2-fixtures"
        manifest = fixture_dir / "pageref-manifest.json"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "output"
            result = self._run_cli(fixture_dir / "pageref-source.docx", manifest, output)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            delivered = output / "m2.docx"
            self.assertTrue(delivered.is_file())
            text = "\n".join(paragraph.text for paragraph in Document(delivered).paragraphs)
            self.assertNotIn("AttributeError", result.stdout + result.stderr)
            self.assertTrue(text.strip())

    def test_W07_real_word_complete_anonymous_thesis_matrix(self):
        """W07：完整匿名论文必须走生产 CLI 并完成真实 Word/PDF 门禁。"""
        fixture_script = REPO / "docs" / "acceptance" / "create_complete_thesis_fixture.py"
        with tempfile.TemporaryDirectory() as tmp, tempfile.TemporaryDirectory() as output_tmp:
            root = Path(tmp)
            fixture = subprocess.run(
                [sys.executable, str(fixture_script), str(root)],
                cwd=str(REPO),
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(fixture.returncode, 0, fixture.stdout + fixture.stderr)
            manifest = root / "manifest.json"
            output = Path(output_tmp) / "output"
            result = self._run_cli(root, manifest, output)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

            delivered = output / "thesis-complete.docx"
            self.assertTrue(delivered.is_file())
            verification_pdf = root / "thesis-verification.pdf"
            self.assertTrue(export_docx_to_pdf(str(delivered), str(verification_pdf)))
            labels = extract_pdf_page_labels(verification_pdf)
            self.assertIn("i", labels.values())
            self.assertIn("1", labels.values())

            with zipfile.ZipFile(delivered) as package:
                names = set(package.namelist())
                document_xml = package.read("word/document.xml")
                self.assertIn("word/footnotes.xml", names)
                self.assertIn("word/endnotes.xml", names)
                self.assertIn(b"PAGEREF ", document_xml)
                self.assertIn(b"SEQ Figure", document_xml)
                self.assertIn(b"SEQ Table", document_xml)
                self.assertIn(b"<w:numPr>", document_xml)
                self.assertIn(b"oMath", document_xml)
                self.assertIn(b"w:drawing", document_xml)
                self.assertIn(b"w:tbl", document_xml)

            text = "\n".join(paragraph.text for paragraph in Document(delivered).paragraphs)
            self.assertIn("第一章 绪论与研究背景", text)
            self.assertIn("第二章 算法设计与实验评估", text)
            self.assertNotIn("AttributeError", result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
