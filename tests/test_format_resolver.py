# -*- coding: utf-8 -*-
"""
格式包解析、继承与能力验证测试 (tests/test_format_resolver.py)
"""

import unittest
import tempfile
import json
from pathlib import Path

from lib.config import ConfigError
from lib.format_schema import (
    FormatDiagnosticCode,
    LengthValue,
    LineSpacing,
    PageSpec,
    ResolvedFormat,
)
from lib.format_resolver import (
    resolve_format_package,
    resolve_format_reference,
    SUPPORTED_ENGINE_CAPABILITIES,
)


class FormatResolverTest(unittest.TestCase):

    def test_load_builtin_presets(self):
        presets = ["legacy-official@1.0.0", "report-basic@1.0.0", "academic-basic@1.0.0"]
        for p in presets:
            rf = resolve_format_package(f"preset:{p}")
            self.assertIsInstance(rf, ResolvedFormat)
            self.assertTrue(len(rf.styles) > 0)
            self.assertTrue(len(rf.roles) > 0)
            self.assertIsNotNone(rf.page)
            self.assertTrue(len(rf.content_hash) == 64)

    def test_unit_conversions(self):
        # 1 inch = 25.4 mm = 72 pt = 1440 twip
        len_pt = LengthValue(value=12.0, unit="pt")
        self.assertEqual(len_pt.to_twip(), 240)
        self.assertAlmostEqual(len_pt.to_pt(), 12.0)

        len_mm = LengthValue(value=25.4, unit="mm")
        self.assertEqual(len_mm.to_twip(), 1440)
        self.assertAlmostEqual(len_mm.to_pt(), 72.0)

        len_char = LengthValue(value=2.0, unit="char")
        self.assertEqual(len_char.to_twip(base_font_size_pt=16.0), round(2.0 * 16.0 * 20.0))

    def test_extends_inheritance_and_provenance(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            custom_pkg = tmpdir / "my_custom.json"
            custom_pkg.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "my-custom",
                "version": "1.0.0",
                "extends": "preset:report-basic@1.0.0",
                "page": {
                    "margin_top_mm": 40.0
                },
                "styles": {
                    "body": {
                        "run": {
                            "size_pt": 13.5
                        }
                    }
                }
            }, ensure_ascii=False), encoding="utf-8")

            rf = resolve_format_package(str(custom_pkg))
            self.assertIsInstance(rf, ResolvedFormat)
            # 覆写生效
            self.assertAlmostEqual(rf.page.margin_top_mm, 40.0)
            self.assertAlmostEqual(rf.styles["body"].run.size_pt, 13.5)
            # 继承自父包的属性保持
            self.assertAlmostEqual(rf.page.margin_bottom_mm, 25.4)
            self.assertEqual(rf.styles["body"].run.east_asia, "微软雅黑")

            # 检查来源追踪
            self.assertIn("/page/margin_top_mm", rf.provenance)
            self.assertIn("my-custom@1.0.0", rf.provenance["/page/margin_top_mm"])
            self.assertIn("/page/margin_bottom_mm", rf.provenance)
            self.assertIn("report-basic@1.0.0", rf.provenance["/page/margin_bottom_mm"])

    def test_inheritance_cycle_detected(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            a_file = tmpdir / "a.json"
            b_file = tmpdir / "b.json"

            a_file.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "pkg-a",
                "version": "1.0.0",
                "extends": "b.json"
            }), encoding="utf-8")

            b_file.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "pkg-b",
                "version": "1.0.0",
                "extends": "a.json"
            }), encoding="utf-8")

            with self.assertRaises(ConfigError) as ctx:
                resolve_format_package(str(a_file))
            self.assertIn(FormatDiagnosticCode.FORMAT_CYCLE, str(ctx.exception))

    def test_reject_null_in_overrides(self):
        with self.assertRaises(ConfigError) as ctx:
            resolve_format_package(
                "preset:report-basic@1.0.0",
                overrides={"styles": {"body": {"run": {"size_pt": None}}}}
            )
        self.assertIn("不能为 null", str(ctx.exception))

    def test_reject_unknown_and_unsupported_capabilities(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            bad_cap_file = tmpdir / "bad_cap.json"
            bad_cap_file.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "bad-cap",
                "version": "1.0.0",
                "extends": "preset:report-basic@1.0.0",
                "required_capabilities": ["non_existent.capability.v99"]
            }), encoding="utf-8")

            with self.assertRaises(ConfigError) as ctx:
                resolve_format_package(str(bad_cap_file))
            self.assertIn(FormatDiagnosticCode.UNSUPPORTED_CAPABILITY, str(ctx.exception))

            # 尚未实现能力（如 bibliography.transform.v1）报错拦截
            future_cap_file = tmpdir / "future_cap.json"
            future_cap_file.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "future-cap",
                "version": "1.0.0",
                "extends": "preset:report-basic@1.0.0",
                "required_capabilities": ["bibliography.transform.v1"]
            }), encoding="utf-8")

            with self.assertRaises(ConfigError) as ctx:
                resolve_format_package(str(future_cap_file))
            self.assertIn("后续扩展规划", str(ctx.exception))

            # M2 能力（pagination.roman.v1, layout.parts.v1, notes.merge.v1, fields.managed_update.v1）应当全部正常通过
            m2_all_cap_file = tmpdir / "m2_all_cap.json"
            m2_all_cap_file.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "m2-all-cap",
                "version": "1.0.0",
                "extends": "preset:report-basic@1.0.0",
                "required_capabilities": [
                    "pagination.roman.v1",
                    "layout.parts.v1",
                    "notes.merge.v1",
                    "fields.managed_update.v1",
                ]
            }), encoding="utf-8")
            resolved_m2 = resolve_format_package(str(m2_all_cap_file))
            self.assertEqual(resolved_m2.id, "m2-all-cap")

    def test_geometry_validation_errors(self):
        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            bad_geo = tmpdir / "bad_geo.json"
            bad_geo.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "bad-geo",
                "version": "1.0.0",
                "extends": "preset:report-basic@1.0.0",
                "page": {
                    "width_mm": 100.0,
                    "margin_left_mm": 60.0,
                    "margin_right_mm": 50.0  # 60 + 50 = 110 > 100
                }
            }), encoding="utf-8")

            with self.assertRaises(ConfigError) as ctx:
                resolve_format_package(str(bad_geo))
            self.assertIn(FormatDiagnosticCode.INVALID_GEOMETRY, str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
