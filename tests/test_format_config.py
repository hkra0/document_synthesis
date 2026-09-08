# -*- coding: utf-8 -*-
"""
v3 项目清单与格式配置测试 (tests/test_format_config.py)
"""

import unittest
import tempfile
import json
from pathlib import Path

from lib.config import (
    ProjectConfig,
    ConfigError,
    load_project_config,
    DEFAULT_SCHEMA_VERSION,
    SUPPORTED_SCHEMA_VERSIONS,
)
from lib.format_schema import FormattingPolicy, ResolvedFormat


class FormatConfigTest(unittest.TestCase):

    def test_schema_version_constants(self):
        self.assertEqual(DEFAULT_SCHEMA_VERSION, 2)
        self.assertEqual(SUPPORTED_SCHEMA_VERSIONS, {1, 2, 3})

    def test_v3_requires_format_and_formatting(self):
        # 缺少 format
        raw1 = {
            "schema_version": 3,
            "project_name": "TestV3",
            "formatting": {"mode": "restyle"},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw1)
        self.assertIn("format", str(ctx.exception))

        # 缺少 formatting
        raw2 = {
            "schema_version": 3,
            "project_name": "TestV3",
            "format": {"ref": "preset:report-basic@1.0.0"},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw2)
        self.assertIn("formatting", str(ctx.exception))

    def test_v3_rejects_top_level_fonts_and_page_setup(self):
        raw_fonts = {
            "schema_version": 3,
            "project_name": "TestV3",
            "format": {"ref": "preset:report-basic@1.0.0"},
            "formatting": {"mode": "restyle"},
            "fonts": {"body": "宋体"},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw_fonts)
        self.assertIn("不再接受顶层 fonts 与 page_setup", str(ctx.exception))

        raw_page = {
            "schema_version": 3,
            "project_name": "TestV3",
            "format": {"ref": "preset:report-basic@1.0.0"},
            "formatting": {"mode": "restyle"},
            "page_setup": {"width_cm": 21.0},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw_page)
        self.assertIn("不再接受顶层 fonts 与 page_setup", str(ctx.exception))

    def test_v3_validates_formatting_mode_and_policies(self):
        # 非法 mode
        raw = {
            "schema_version": 3,
            "project_name": "TestV3",
            "format": {"ref": "preset:report-basic@1.0.0"},
            "formatting": {"mode": "invalid_mode"},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw)
        self.assertIn("formatting.mode", str(ctx.exception))

        # mixed 模式未指定 page_policy
        raw_mixed = {
            "schema_version": 3,
            "project_name": "TestV3",
            "format": {"ref": "preset:report-basic@1.0.0"},
            "formatting": {"mode": "mixed"},
        }
        with self.assertRaises(ConfigError) as ctx:
            ProjectConfig(raw_mixed)
        self.assertIn("显式声明 page_policy", str(ctx.exception))

    def test_v3_successful_instantiation_and_compatibility_props(self):
        raw = {
            "schema_version": 3,
            "project_name": "MyAcademicPaper",
            "format": {
                "ref": "preset:academic-basic@1.0.0",
                "overrides": {
                    "page": {"margin_top_mm": 30.0}
                }
            },
            "formatting": {
                "mode": "restyle",
                "page_policy": "target",
                "on_unmapped": "error"
            },
            "source": {"strategy": "docx_document"},
            "output": {
                "documents": [{"id": "main", "filename": "thesis.docx", "parts": ["body"]}]
            }
        }
        config = ProjectConfig(raw)
        self.assertEqual(config.schema_version, 3)
        self.assertEqual(config.project_name, "MyAcademicPaper")
        self.assertIsInstance(config.formatting, FormattingPolicy)
        self.assertEqual(config.formatting.mode, "restyle")
        self.assertEqual(config.formatting.on_unmapped, "error")

        # 检查 resolved_format
        self.assertIsInstance(config.resolved_format, ResolvedFormat)
        self.assertEqual(config.resolved_format.id, "academic-basic")
        self.assertAlmostEqual(config.resolved_format.page.margin_top_mm, 30.0)

        # 检查兼容字段映射
        self.assertAlmostEqual(config.page_setup["margin_top_cm"], 3.0)
        self.assertEqual(config.fonts["body"], "宋体")
        self.assertFalse(config.legacy_policy.text_blackening)

    def test_legacy_v2_has_resolved_format_via_adapter(self):
        raw = {
            "schema_version": 2,
            "project_name": "LegacyProject",
            "page_setup": {"width_cm": 21.0, "height_cm": 29.7},
            "fonts": {"body": "仿宋_GB2312"}
        }
        config = ProjectConfig(raw)
        self.assertEqual(config.schema_version, 2)
        self.assertIsInstance(config.resolved_format, ResolvedFormat)
        self.assertEqual(config.resolved_format.id, "legacy-adapted")
        self.assertTrue(config.legacy_policy.text_blackening)
        self.assertEqual(config.formatting.mode, "restyle")


if __name__ == "__main__":
    unittest.main()
