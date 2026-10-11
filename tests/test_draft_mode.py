# -*- coding: utf-8 -*-
"""P5 草稿模式：没有可精确分页的 Word 时，目录页码交给 Word 打开时更新。

覆盖：
1. 没有 --allow-draft 时构建失败，已有交付物不被替换
2. 草稿标注：*_draft.docx、显示“?”的 PAGEREF、不在打开时自动更新、终端最后一行
3. 草稿永远不覆盖精确交付物
4. 能精确分页时 --allow-draft 不起作用
5. v3 草稿元数据写入 build-metadata_draft.json，fidelity 为 draft
6. 引用另一份交付物的目录只写占位符
7. smoke --allow-draft 通过；精确 smoke 拒绝草稿元数据
8. FieldUpdater 在草稿中把所有 PAGEREF 显示为“?”（含来源文档遗留的页码）
"""

import contextlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from lib.composition import DRAFT_PAGE_PLACEHOLDER
from lib.drafting import ALLOW_DRAFT_ENV, DRAFT_METADATA_NAME, DRAFT_NOTICE, draft_filename
from lib.engine import BuildError, UnifiedSynthesizer
from lib.field_updater import FieldUpdater
from lib.office import NullBackend, set_backend

MINIMAL = ROOT / "examples" / "minimal-demo"
CUSTOM = ROOT / "examples" / "custom-format-demo"


def _document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("word/document.xml").decode("utf-8")


def _settings_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("word/settings.xml").decode("utf-8")


class DraftModeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name) / "output"
        set_backend(NullBackend())
        self.env = patch.dict(os.environ, {ALLOW_DRAFT_ENV: "1"})

    def tearDown(self) -> None:
        set_backend(None)
        self.tmp.cleanup()

    def _build(self, source, manifest, overrides=None):
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            UnifiedSynthesizer.synthesize(source, self.out, manifest, override_paths=overrides)
        return stdout.getvalue()

    def test_without_allow_draft_build_fails_and_keeps_existing_files(self) -> None:
        self.out.mkdir(parents=True)
        existing = self.out / "demo_complete.docx"
        existing.write_bytes(b"previous verified delivery")
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop(ALLOW_DRAFT_ENV, None)
            with self.assertRaises(BuildError):
                with contextlib.redirect_stdout(io.StringIO()):
                    UnifiedSynthesizer.synthesize(MINIMAL / "source", self.out, MINIMAL / "manifest.json")
        self.assertEqual(existing.read_bytes(), b"previous verified delivery")
        self.assertFalse((self.out / "demo_complete_draft.docx").exists())

    def test_draft_build_is_marked_everywhere(self) -> None:
        with self.env:
            output = self._build(MINIMAL / "source", MINIMAL / "manifest.json")
        delivered = sorted(p.name for p in self.out.iterdir() if p.is_file())
        self.assertEqual(delivered, ["demo_complete_draft.docx"])
        draft = self.out / "demo_complete_draft.docx"
        document = _document_xml(draft)
        self.assertEqual(document.count("PAGEREF _Toc_auto_"), 2)
        # 打开时自动更新会在分页前求值，得到看似合理的错误页码：两种触发方式都不能有。
        self.assertNotIn("w:dirty", document)
        self.assertNotIn("w:updateFields", _settings_xml(draft))
        self.assertNotIn("DRAFTPAGEREFTOKEN", document)
        from smoke_test import _draft_pageref_results
        with zipfile.ZipFile(draft) as package:
            results = _draft_pageref_results(package.read("word/document.xml"))
        self.assertEqual([shown for _, shown in results], ["?", "?"])
        self.assertEqual(output.strip().splitlines()[-1], DRAFT_NOTICE)

    def test_draft_never_replaces_exact_delivery(self) -> None:
        self.out.mkdir(parents=True)
        exact = self.out / "demo_complete.docx"
        exact.write_bytes(b"exact delivery")
        with self.env:
            self._build(MINIMAL / "source", MINIMAL / "manifest.json")
        self.assertEqual(exact.read_bytes(), b"exact delivery")
        self.assertTrue((self.out / "demo_complete_draft.docx").is_file())

    def test_allow_draft_has_no_effect_when_exact_backend_is_available(self) -> None:
        from test_office_backends import FakeOfficeBackend

        set_backend(FakeOfficeBackend(fidelity="exact"))
        with self.env:
            output = self._build(MINIMAL / "source", MINIMAL / "manifest.json")
        self.assertTrue((self.out / "demo_complete.docx").is_file())
        self.assertFalse((self.out / "demo_complete_draft.docx").exists())
        self.assertNotIn(DRAFT_NOTICE, output)
        self.assertNotIn("PAGEREF _Toc_auto_", _document_xml(self.out / "demo_complete.docx"))

    def test_v3_draft_metadata_is_separate_and_marked(self) -> None:
        with self.env:
            self._build(CUSTOM / "manuscript.docx", CUSTOM / "manifest.json")
        self.assertFalse((self.out / "build-metadata.json").exists())
        metadata = json.loads((self.out / DRAFT_METADATA_NAME).read_text(encoding="utf-8"))
        self.assertEqual(metadata["fidelity"], "draft")
        record = metadata["deliveries"]["main"]
        self.assertEqual(record["filename"], "academic_paper_draft.docx")
        self.assertEqual(record["fidelity"], "draft")
        self.assertNotIn("page_map", record)

    def test_toc_referencing_another_delivery_uses_placeholder(self) -> None:
        with self.env:
            output = self._build(
                MINIMAL / "source", MINIMAL / "manifest.json", [MINIMAL / "split.override.json"]
            )
        front = _document_xml(self.out / "demo_cover_toc_draft.docx")
        self.assertNotIn("PAGEREF", front)
        self.assertEqual(front.count(f"<w:t>{DRAFT_PAGE_PLACEHOLDER}</w:t>"), 2)
        self.assertIn("跨文件无法用域更新页码", output)
        self.assertEqual(_document_xml(self.out / "demo_toc_body_draft.docx").count("PAGEREF _Toc_auto_"), 2)

    def test_smoke_accepts_draft_and_exact_smoke_rejects_draft_metadata(self) -> None:
        import smoke_test

        with self.env:
            self._build(CUSTOM / "manuscript.docx", CUSTOM / "manifest.json")
        with contextlib.redirect_stdout(io.StringIO()):
            self.assertTrue(smoke_test.test_project(
                CUSTOM / "manuscript.docx", self.out, CUSTOM / "manifest.json", draft=True,
            ))

        # 把草稿冒充为精确交付物：精确 smoke 必须因保真度不符而失败。
        disguised = Path(self.tmp.name) / "disguised"
        disguised.mkdir()
        shutil.copy2(self.out / "academic_paper_draft.docx", disguised / "academic_paper.docx")
        metadata = json.loads((self.out / DRAFT_METADATA_NAME).read_text(encoding="utf-8"))
        metadata["deliveries"]["main"]["filename"] = "academic_paper.docx"
        (disguised / "build-metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            passed = smoke_test.test_project(
                CUSTOM / "manuscript.docx", disguised, CUSTOM / "manifest.json", structure_only=True,
            )
        self.assertFalse(passed)
        self.assertIn("保真度", output.getvalue())

    def test_smoke_rejects_open_time_update_and_unverified_numbers(self) -> None:
        import smoke_test
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        with self.env:
            self._build(MINIMAL / "source", MINIMAL / "manifest.json")
        draft = self.out / "demo_complete_draft.docx"
        pristine = draft.read_bytes()

        def smoke_output():
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                passed = smoke_test.test_project(MINIMAL / "source", self.out, MINIMAL / "manifest.json", draft=True)
            return passed, output.getvalue()

        self.assertTrue(smoke_output()[0])

        # 打开时自动更新域
        doc = Document(str(draft))
        flag = OxmlElement("w:updateFields")
        flag.set(qn("w:val"), "true")
        doc.settings.element.append(flag)
        doc.save(str(draft))
        passed, output = smoke_output()
        self.assertFalse(passed)
        self.assertIn("updateFields", output)

        # 页码域显示了未经核验的数字
        draft.write_bytes(pristine)
        doc = Document(str(draft))
        for text in doc.element.body.iter(qn("w:t")):
            if text.text == DRAFT_PAGE_PLACEHOLDER:
                text.text = "1"
                break
        doc.save(str(draft))
        passed, output = smoke_output()
        self.assertFalse(passed)
        self.assertIn("未经核验的页码", output)


class DraftHelpersTests(unittest.TestCase):
    def test_draft_filename(self) -> None:
        self.assertEqual(draft_filename("demo.docx"), "demo_draft.docx")
        self.assertEqual(draft_filename("材料.合成.docx"), "材料.合成_draft.docx")

    def test_field_updater_defers_pageref_but_updates_seq(self) -> None:
        doc = Document()
        paragraph = doc.add_paragraph()
        paragraph._p.append(_simple_field(" PAGEREF target \\h ", "7"))
        caption = doc.add_paragraph()
        caption._p.append(_simple_field(" SEQ Figure ", "9"))
        FieldUpdater({}, defer_pagerefs=True).update_document_fields(doc)
        texts = ["".join(t.text or "" for t in p._p.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")) for p in doc.paragraphs]
        # 来源文档遗留的页码（7）也不能保留：草稿中一律显示“?”。
        self.assertEqual(texts, ["?", "1"])


def _simple_field(instruction: str, cached: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = cached
    run.append(text)
    field.append(run)
    return field


if __name__ == "__main__":
    unittest.main()
