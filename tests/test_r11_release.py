# -*- coding: utf-8 -*-
"""R11 发布收尾、示例闭环与格式继承核验回归。"""

import hashlib
import json
import re
import tempfile
import unittest
from pathlib import Path

from docx import Document

from lib.content_integrity import verify_delivery_format
from lib.format_analysis import analyze_format_sample
from lib.format_review import compile_format_package
from lib.format_resolver import resolve_format_package
from lib.layout import apply_section_spec
from lib.manifest_migration import migrate_manifest_file
from lib.preview import build_preview
from lib.style_applier import apply_roles, install_styles


ROOT = Path(__file__).resolve().parents[1]
EXAMPLE = ROOT / "examples" / "custom-format-demo"
SAMPLE = EXAMPLE / "sample.docx"
MANUSCRIPT = EXAMPLE / "manuscript.docx"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class R11ReleaseTest(unittest.TestCase):
    def test_final_decisions_are_bound_to_the_current_sources(self):
        sample_report = analyze_format_sample(SAMPLE)
        target_report = analyze_format_sample(MANUSCRIPT)
        sample_decisions = json.loads((EXAMPLE / "decisions.final.json").read_text(encoding="utf-8"))
        target_decisions = json.loads((EXAMPLE / "manuscript-decisions.final.json").read_text(encoding="utf-8"))

        self.assertEqual(sample_report["report_id"], sample_decisions["report_id"])
        self.assertEqual(sample_report["source_sha256"], sample_decisions["source_sha256"])
        self.assertEqual(target_report["report_id"], target_decisions["report_id"])
        self.assertEqual(target_report["source_sha256"], target_decisions["source_sha256"])
        self.assertEqual(sample_decisions["source_sha256"], sha256(SAMPLE))
        self.assertEqual(target_decisions["source_sha256"], sha256(MANUSCRIPT))
        self.assertNotIn("pending", json.dumps(sample_decisions, ensure_ascii=False))
        self.assertNotIn("pending", json.dumps(target_decisions, ensure_ascii=False))

    def test_compiled_package_and_preview_are_source_free(self):
        report = analyze_format_sample(SAMPLE)
        decisions = json.loads((EXAMPLE / "decisions.final.json").read_text(encoding="utf-8"))

        with tempfile.TemporaryDirectory(prefix="r11-preview-") as tmp:
            root = Path(tmp)
            package_path = root / "academic-demo.json"
            compile_format_package(
                report,
                decisions,
                base_format_ref="preset:academic-basic@1.0.0",
                format_id="academic-demo",
                format_version="1.0.0",
                output_path=package_path,
            )
            preview_path = root / "preview.html"
            metadata_path = root / "preview-metadata.json"
            metadata = build_preview(package_path, preview_path, metadata_path=metadata_path)

            package_text = package_path.read_text(encoding="utf-8")
            preview_text = preview_path.read_text(encoding="utf-8")
            for cluster in report["clusters"]:
                for sample_text in cluster.get("sample_texts", []):
                    if len(sample_text) >= 6:
                        self.assertNotIn(sample_text, package_text)
                        self.assertNotIn(sample_text, preview_text)
            self.assertEqual(metadata["format_id"], "academic-demo")
            self.assertTrue(metadata["fictional_content"])
            self.assertTrue(metadata["approximate"])
            self.assertFalse(metadata["source_content_included"])

    def test_v1_migration_is_independent_and_non_destructive(self):
        original = {
            "schema_version": 1,
            "project_name": "R11 migration",
            "source": {"strategy": "docx_document", "file": "manuscript.docx"},
            "output": {"compiled_document": "legacy.docx"},
        }
        with tempfile.TemporaryDirectory(prefix="r11-migration-") as tmp:
            root = Path(tmp)
            source = root / "legacy.json"
            target = root / "migrated" / "manifest-v2.json"
            source.write_text(json.dumps(original, ensure_ascii=False, indent=2), encoding="utf-8")
            before = sha256(source)

            migrated = migrate_manifest_file(source, target)

            self.assertEqual(migrated["schema_version"], 2)
            self.assertTrue(target.is_file())
            self.assertEqual(before, sha256(source))
            self.assertNotEqual(source.resolve(), target.resolve())
            self.assertEqual(migrated["output"]["documents"][0]["filename"], "legacy.docx")

    def test_inherited_semantic_alias_uses_the_child_declared_style(self):
        child_package = {
            "format_schema_version": 1,
            "id": "r11-child",
            "version": "1.0.0",
            "extends": "preset:academic-basic@1.0.0",
            "styles": {
                "compiled_body": {
                    "run": {"east_asia": "宋体", "latin": "Courier New", "size_pt": 10.0},
                    "paragraph": {"alignment": "left", "space_after_pt": 3.0},
                }
            },
            "roles": {"body": {"style": "compiled_body"}},
        }
        with tempfile.TemporaryDirectory(prefix="r11-inheritance-") as tmp:
            root = Path(tmp)
            package_path = root / "child.json"
            package_path.write_text(json.dumps(child_package, ensure_ascii=False), encoding="utf-8")
            resolved = resolve_format_package(str(package_path))

            doc = Document()
            apply_section_spec(doc.sections[0], resolved)
            install_styles(doc, resolved)
            doc.add_paragraph("R11 body")
            apply_roles(doc, {"/w:document/w:body/w:p[1]": "body"}, resolved)
            delivered = root / "delivered.docx"
            doc.save(delivered)

            report = verify_delivery_format(delivered, resolved)
            self.assertTrue(report.passed, report.violations)
            self.assertEqual(report.violations, [])

    def test_release_docs_use_current_cli_and_state_boundaries(self):
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        example_readme = (EXAMPLE / "README.md").read_text(encoding="utf-8")
        capabilities = (ROOT / "docs" / "formatting-capabilities.md").read_text(encoding="utf-8")
        for flag in ("--analysis-dir", "--format-out", "--mapping-out", "--preview-out", "--manifest-out"):
            self.assertIn(flag, readme)
            self.assertIn(flag, example_readme)
        for invalid in ("--output-report", "--output-html", "--output-format", "--output-mapping"):
            self.assertNotIn(invalid, readme)
            self.assertNotIn(invalid, example_readme)
        self.assertNotRegex(readme, re.compile(r"(?<![\w-])--report(?:\s|$)"))
        for boundary in ("OCR/PDF 样例学习", "双栏期刊", "参考文献著录转换", "多样例融合", "全部学校规范认证"):
            self.assertIn(boundary, readme)
            self.assertIn(boundary, capabilities)
        self.assertNotIn("R11 待实施", readme + capabilities)

    def test_r11_status_is_a_formal_release_record(self):
        status_path = ROOT / "docs" / "acceptance" / "r11-status.json"
        status = json.loads(status_path.read_text(encoding="utf-8"))
        self.assertEqual(status["batch"], "R11")
        self.assertEqual(status["status"], "verified")
        joined_commands = "\n".join(status["commands"])
        for marker in ("--analyze-format", "--compile-format", "--compile-mapping", "--render-preview", "--migrate-manifest", "--plan", "--doctor", "smoke_test.py"):
            self.assertIn(marker, joined_commands)
        self.assertEqual(status["results"]["example_workflow"]["smoke"], "passed")
        self.assertEqual(status["results"]["example_workflow"]["build"], "passed")
        self.assertFalse(status["results"]["preview"]["source_content_included"])
        self.assertEqual(status["results"]["migration"]["input_unchanged"], True)
        self.assertEqual(status["cleanup"]["work_directories"], 0)


if __name__ == "__main__":
    unittest.main()
