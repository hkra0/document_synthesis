"""封面模板解析：显式模板缺失必须报错，不能静默回退到扫描到的其他模板。"""

import json
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from docx import Document

from lib.config import ConfigError, find_template_file, load_project_config


def _save_docx(path: Path, text: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(path))
    return path


class TemplateResolutionTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name).resolve()
        self.project_dir = self.root / "project"
        self.project_dir.mkdir()
        # 项目 templates/ 中放一个“诱饵”模板：显式模板缺失时绝不能选中它。
        self.decoy = _save_docx(self.project_dir / "templates" / "封面+目录.docx", "decoy")

    def tearDown(self):
        self._tmp.cleanup()

    def _load(self, cover):
        manifest_dir = self.root / "config"
        manifest_dir.mkdir(exist_ok=True)
        manifest = manifest_dir / "manifest.json"
        manifest.write_text(json.dumps({
            "schema_version": 2,
            "project_name": "Template Resolution",
            "cover": cover,
        }), encoding="utf-8")
        return load_project_config(self.project_dir, manifest_path=manifest)

    def test_missing_explicit_template_raises_with_resolved_path(self):
        config = self._load({"template": "templates/missing.docx"})
        expected = (self.root / "config" / "templates" / "missing.docx").resolve()
        with self.assertRaises(ConfigError) as ctx:
            config.get_template_path()
        self.assertIn(str(expected), str(ctx.exception))

    def test_missing_explicit_template_does_not_fall_back_to_same_basename(self):
        _save_docx(self.project_dir / "templates" / "cover.docx", "other cover")
        config = self._load({"template": "templates/cover.docx"})
        with self.assertRaises(ConfigError):
            config.get_template_path()

    def test_explicit_template_resolves_relative_to_manifest(self):
        template = _save_docx(self.root / "config" / "templates" / "cover.docx", "cover")
        config = self._load({"template": "templates/cover.docx"})
        self.assertEqual(config.cover["template"], str(template))
        self.assertEqual(config.get_template_path(), template)

    def test_dynamic_scan_used_when_no_template_configured(self):
        config = self._load({"main_title": "标题"})
        self.assertEqual(config.get_template_path(), self.decoy.resolve())

    def test_template_false_disables_templates(self):
        config = self._load({"template": False})
        self.assertIsNone(config.get_template_path())

    def test_find_template_file_relative_hint_uses_project_dir(self):
        self.assertEqual(
            find_template_file("templates/封面+目录.docx", self.project_dir),
            self.decoy.resolve(),
        )
        with self.assertRaises(ConfigError):
            find_template_file("templates/nope.docx", self.project_dir)

    def test_chinese_demo_manifest_uses_demo_cover(self):
        demo = ROOT / "examples" / "chinese-demo"
        config = load_project_config(demo / "source", demo / "manifest.json")
        self.assertEqual(config.get_template_path(), (demo / "templates" / "cover.docx").resolve())


if __name__ == "__main__":
    unittest.main()
