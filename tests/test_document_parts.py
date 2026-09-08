"""Tests for Document Parts (layout.parts) and Custom Part Lifecycle."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lib.composition import assemble_document, add_bookmark
from lib.config import ProjectConfig, ConfigError
from lib.delivery import build_deliveries, validate_delivery_structure
from lib.document_parts import DocumentPart, PartKind, SelectionError, get_part_boundary_bookmark
from lib.qa import OfficeExportError


def make_v3_manifest(**kwargs):
    base = {
        "schema_version": 3,
        "project_name": "Test Project",
        "format": {"ref": "preset:academic-basic@1.0.0"},
        "formatting": {"mode": "preserve"},
        "cover": {"template": False},
    }
    base.update(kwargs)
    return base


class DocumentPartConfigTest(unittest.TestCase):
    def test_default_parts_when_layout_omitted_v2(self):
        cfg = ProjectConfig({
            "schema_version": 2,
            "project_name": "Default Parts v2",
            "output": {"documents": [{"id": "main", "filename": "main.docx", "parts": ["cover", "toc", "body"]}]},
        })
        self.assertIn("cover", cfg.parts)
        self.assertIn("toc", cfg.parts)
        self.assertIn("body", cfg.parts)
        self.assertEqual(cfg.parts["cover"].kind, "cover")
        self.assertEqual(cfg.parts["toc"].kind, "generated_toc")
        self.assertEqual(cfg.parts["body"].kind, "content")

    def test_default_parts_when_layout_omitted_v3(self):
        cfg = ProjectConfig(make_v3_manifest(
            output={"documents": [{"id": "main", "filename": "main.docx", "parts": ["cover", "toc", "body"]}]}
        ))
        self.assertIn("cover", cfg.parts)
        self.assertIn("toc", cfg.parts)
        self.assertIn("body", cfg.parts)
        self.assertEqual(cfg.parts["cover"].kind, "cover")
        self.assertEqual(cfg.parts["toc"].kind, "generated_toc")
        self.assertEqual(cfg.parts["body"].kind, "content")

    def test_custom_parts_declaration(self):
        cfg = ProjectConfig(make_v3_manifest(
            source={
                "regions": {
                    "abstract": {
                        "start": "/w:document/w:body/w:p[1]",
                        "end": "/w:document/w:body/w:p[2]",
                    },
                    "main": {
                        "start": "/w:document/w:body/w:p[2]",
                        "end": "end_of_document",
                    },
                }
            },
            layout={
                "parts": {
                    "front_matter": {"kind": "content", "source_region": "abstract"},
                    "toc": {"kind": "generated_toc"},
                    "body": {"kind": "content", "source_region": "main"},
                }
            },
            output={
                "documents": [
                    {"id": "doc1", "filename": "doc1.docx", "parts": ["front_matter", "toc", "body"]}
                ]
            },
        ))
        self.assertIn("front_matter", cfg.parts)
        part = cfg.parts["front_matter"]
        self.assertEqual(part.id, "front_matter")
        self.assertEqual(part.kind, "content")
        self.assertEqual(part.source_region, "abstract")

    def test_invalid_part_kind(self):
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(make_v3_manifest(
                layout={
                    "parts": {
                        "p1": {"kind": "invalid_kind"}
                    }
                },
            ))
        self.assertIn("只支持 cover, generated_toc 或 content", str(ctx.exception))

    def test_layout_parts_must_be_object(self):
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(make_v3_manifest(
                layout={
                    "parts": [
                        {"id": "p1", "kind": "content"}
                    ]
                },
            ))
        self.assertIn("layout.parts 必须是非空对象", str(ctx.exception))

    def test_document_references_undeclared_part(self):
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(make_v3_manifest(
                layout={
                    "parts": {
                        "toc": {"kind": "generated_toc"}
                    }
                },
                output={
                    "documents": [
                        {"id": "main", "filename": "main.docx", "parts": ["toc", "nonexistent"]}
                    ]
                },
            ))
        self.assertIn("未在 layout.parts 中定义", str(ctx.exception))


class DocumentAssemblyCustomPartsTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.root = Path(self.temp_dir.name)

        # Create a body docx with bookmarks
        self.body_path = self.root / "body.docx"
        doc = Document()
        p1 = doc.add_paragraph("Abstract Heading")
        add_bookmark(p1._p, "_Toc_auto_001", 1)
        doc.add_paragraph("Abstract content goes here.")
        p2 = doc.add_paragraph("Chapter 1")
        add_bookmark(p2._p, "_Toc_auto_002", 2)
        doc.add_paragraph("Chapter 1 body content.")
        doc.save(self.body_path)

        self.nodes = [
            {"title": "Abstract Heading", "level": 1, "bookmark_name": "_Toc_auto_001", "bm_id": 1},
            {"title": "Chapter 1", "level": 1, "bookmark_name": "_Toc_auto_002", "bm_id": 2},
        ]

    def _make_custom_config(self):
        return ProjectConfig(make_v3_manifest(
            source={
                "regions": {
                    "abstract": {
                        "start": "/w:document/w:body/w:p[1]",
                        "end": "/w:document/w:body/w:p[3]",
                    },
                    "main": {
                        "start": "/w:document/w:body/w:p[3]",
                        "end": "end_of_document",
                    },
                }
            },
            layout={
                "parts": {
                    "preface": {"kind": "content", "source_region": "abstract"},
                    "toc": {"kind": "generated_toc"},
                    "body": {"kind": "content", "source_region": "main"},
                }
            },
            output={
                "documents": [{"id": "main", "filename": "main.docx", "parts": ["preface", "toc", "body"]}]
            },
        ))

    def test_assemble_with_custom_parts(self):
        cfg = self._make_custom_config()
        out_path = self.root / "assembled.docx"
        spec = cfg.documents[0]
        assemble_document(
            config=cfg,
            spec=spec,
            body_path=self.body_path,
            nodes=self.nodes,
            pages={},
            out_path=out_path,
        )
        self.assertTrue(out_path.exists())

        doc = Document(str(out_path))
        bm_names = set(doc.element.body.xpath(".//w:bookmarkStart/@w:name"))
        # Custom part bookmark: _Synth_part_preface_start and end
        self.assertIn("_Synth_part_preface_start", bm_names)
        self.assertIn("_Synth_part_preface_end", bm_names)
        # Legacy parts retain their standard bookmarks
        self.assertIn("_Synth_toc", bm_names)
        self.assertIn("_Synth_body", bm_names)
        # Bookmarks from each sliced region should appear exactly once
        self.assertIn("_Toc_auto_001", bm_names)
        self.assertIn("_Toc_auto_002", bm_names)

    def test_delivery_validation_with_custom_parts(self):
        cfg = self._make_custom_config()
        out_path = self.root / "assembled.docx"
        spec = cfg.documents[0]
        assemble_document(
            config=cfg,
            spec=spec,
            body_path=self.body_path,
            nodes=self.nodes,
            pages={},
            out_path=out_path,
        )

        # Should pass validation with parts_registry
        validate_delivery_structure(out_path, spec, self.nodes, parts_registry=cfg.parts)

    def test_custom_content_part_without_region_rejects_full_document_fallback(self):
        cfg = ProjectConfig(make_v3_manifest(
            source={"file": self.body_path.name},
            layout={"parts": {"preface": {"kind": "content"}}},
            output={"documents": [{"id": "main", "filename": "main.docx", "parts": ["preface"]}]},
        ))

        with self.assertRaises(SelectionError):
            assemble_document(
                config=cfg,
                spec=cfg.documents[0],
                body_path=self.body_path,
                nodes=self.nodes,
                pages={},
                out_path=self.root / "missing-region.docx",
            )
    def test_delivery_validation_fails_on_missing_custom_part_bookmark(self):
        parts_registry = {
            "preface": DocumentPart(id="preface", kind="content"),
        }
        spec = {"id": "main", "filename": "main.docx", "parts": ["preface"]}
        # Empty document without bookmarks
        empty_doc = self.root / "empty.docx"
        doc = Document()
        doc.add_paragraph("No bookmarks here")
        doc.save(empty_doc)

        with self.assertRaises(OfficeExportError) as ctx:
            validate_delivery_structure(empty_doc, spec, self.nodes, parts_registry=parts_registry)
        self.assertIn("_Synth_part_preface_start", str(ctx.exception))

    def test_build_deliveries_with_custom_parts(self):
        cfg = self._make_custom_config()
        run_dir = self.root / "run"
        run_dir.mkdir()

        import pymupdf

        def fake_inspect(path, pdf_path, names):
            with pymupdf.open() as pdf:
                for _ in range(3):
                    page = pdf.new_page()
                    page.insert_text((72, 72), "Fictional document\nSupporting content\nAdditional content\nFinal content", fontsize=12)
                pdf.save(str(pdf_path))
            return {
                "_Synth_part_preface_start": {"physical_page": 1, "printed_page": 1},
                "_Synth_toc": {"physical_page": 2, "printed_page": 1},
                "_Synth_body": {"physical_page": 3, "printed_page": 3},
                "_Toc_auto_001": {"physical_page": 1, "printed_page": 1},
                "_Toc_auto_002": {"physical_page": 3, "printed_page": 3},
            }

        with patch("lib.delivery.inspect_document", side_effect=fake_inspect):
            staged = build_deliveries(cfg, self.body_path, self.nodes, run_dir, parts_registry=cfg.parts)
            self.assertIn("main", staged)
            self.assertTrue(staged["main"].exists())


if __name__ == "__main__":
    unittest.main()
