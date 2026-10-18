# -*- coding: utf-8 -*-
"""P6 F3 回退字体：font_fallbacks、--font-fallbacks 与记录。

覆盖：
1. font_fallbacks 校验：类型、空列表、回退与原字体相同、注释键
2. 回退规则：只替换缺失的字体、按顺序取第一个已安装的、候选全缺失时失败、
   未声明的缺失字体保留、代码直接写入的字体不替换、本机字体列表不完整时失败
3. ProjectConfig：未启用时不改动且配置契约哈希不变；启用后 fonts 与格式样式一起替换；
   v1/v2 未声明时使用内置公文字体列表；v3 从格式包继承与 overrides 取得
4. 记录：构建元数据与终端摘要、--plan 中的回退行、--font-fallbacks 设置环境变量
"""

import contextlib
import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from lib import fonts  # noqa: E402
from lib.build_plan import compute_config_contract_hash  # noqa: E402
from lib.config import ConfigError, ProjectConfig  # noqa: E402
from lib.format_resolver import resolve_format_package  # noqa: E402
from lib.format_schema import RunStyle, StyleDefinition  # noqa: E402
from test_fonts import _win, build_font, make_face  # noqa: E402

# 本机“已安装”的字体：有系统自带的 仿宋/楷体，没有 仿宋_GB2312
INSTALLED = ["仿宋", "楷体", "宋体", "黑体", "楷体_GB2312", "方正小标宋简体", "Times New Roman"]
MINIMAL_DEMO = ROOT / "examples" / "minimal-demo"


def legacy_raw(**extra):
    raw = {"schema_version": 2, "project_name": "Fictional Fonts", "cover": {"template": False}}
    raw.update(extra)
    return raw


class FakeFontsCase(unittest.TestCase):
    """用合成字体代替本机字体；每个用例默认不启用回退。"""

    installed = INSTALLED

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        files = []
        for number, name in enumerate(self.installed):
            path = self.tmp / f"font{number}.ttf"
            path.write_bytes(build_font([make_face([_win(0x0804, 1, name)])]))
            files.append((path, "test"))
        self.index = fonts.FontIndex.build(files)
        self._patches = [
            patch.object(fonts, "installed_fonts", lambda: self.index),
            patch.dict(os.environ, {fonts.FONT_FALLBACKS_ENV: "0"}),
        ]
        for item in self._patches:
            item.start()

    def tearDown(self):
        for item in reversed(self._patches):
            item.stop()
        self._tmp.cleanup()

    def enable(self):
        os.environ[fonts.FONT_FALLBACKS_ENV] = "1"


class ParseFontFallbacksTest(unittest.TestCase):
    def test_valid_mapping_is_trimmed(self):
        parsed = fonts.parse_font_fallbacks({" 仿宋_GB2312 ": [" 仿宋 ", "宋体"], "_comment": "说明"})
        self.assertEqual(parsed, {"仿宋_GB2312": ("仿宋", "宋体")})
        self.assertEqual(fonts.parse_font_fallbacks(None), {})

    def test_invalid_values(self):
        cases = [
            ["仿宋"],
            {"仿宋_GB2312": "仿宋"},
            {"仿宋_GB2312": []},
            {"仿宋_GB2312": ["仿宋", " "]},
            {"": ["仿宋"]},
            {"仿宋_GB2312": ["仿宋_gb2312"]},
        ]
        for raw in cases:
            with self.subTest(raw=raw), self.assertRaises(fonts.FontFallbackError):
                fonts.parse_font_fallbacks(raw)


class ResolveFontFallbacksTest(FakeFontsCase):
    def required(self, *names):
        return [fonts.RequiredFont(name, ("样式 body（中文）",)) for name in names]

    def test_replaces_only_missing_fonts_with_first_installed_candidate(self):
        resolution = fonts.resolve_font_fallbacks(
            self.required("仿宋_GB2312", "楷体_GB2312", "P6缺失测试字体"),
            {"仿宋_GB2312": ("P6另一个缺失字体", "仿宋"), "楷体_GB2312": ("楷体",)},
            self.index,
        )
        self.assertEqual(resolution.substitutions, (("仿宋_GB2312", "仿宋"),))
        self.assertEqual(resolution.missing_unreplaced, ("P6缺失测试字体",))
        self.assertEqual(resolution.mapping, {"仿宋_gb2312": "仿宋"})
        self.assertEqual(dict(resolution.checked)["楷体_GB2312"], "installed")

    def test_all_candidates_missing_fails(self):
        with self.assertRaisesRegex(fonts.FontFallbackError, "P6另一个缺失字体"):
            fonts.resolve_font_fallbacks(self.required("仿宋_GB2312"),
                                         {"仿宋_GB2312": ("P6另一个缺失字体",)}, self.index)

    def test_renderer_fonts_are_not_replaced(self):
        empty = fonts.FontIndex.build([(self.tmp / "font0.ttf", "test")])  # 只有 仿宋
        resolution = fonts.resolve_font_fallbacks(
            [fonts.RequiredFont("宋体", ("页脚页码（styles.setup_footer）",))], {"宋体": ("仿宋",)}, empty)
        self.assertEqual(resolution.substitutions, ())
        self.assertEqual(resolution.missing_unreplaced, ("宋体",))

    def test_incomplete_index_fails(self):
        with self.assertRaisesRegex(fonts.FontFallbackError, "无法完整列出本机字体"):
            fonts.resolve_font_fallbacks(self.required("仿宋_GB2312"), {"仿宋_GB2312": ("仿宋",)},
                                         fonts.FontIndex.build([]))

    def test_substitute_format_fonts_does_not_modify_original(self):
        original = StyleDefinition("body", run=RunStyle(east_asia="仿宋_GB2312", latin="Times New Roman",
                                                        complex_script="仿宋_GB2312", size_pt=16.0))
        from dataclasses import make_dataclass

        Resolved = make_dataclass("Resolved", [("styles", dict)], frozen=True)
        resolved = Resolved({"body": original, "plain": StyleDefinition("plain")})
        replaced = fonts.substitute_format_fonts(resolved, {"仿宋_gb2312": "仿宋"})
        run = replaced.styles["body"].run
        self.assertEqual((run.east_asia, run.latin, run.complex_script, run.size_pt),
                         ("仿宋", "Times New Roman", "仿宋", 16.0))
        self.assertEqual(resolved.styles["body"].run.east_asia, "仿宋_GB2312")
        self.assertIsNone(replaced.styles["plain"].run)


class LegacyConfigTest(FakeFontsCase):
    def test_disabled_changes_nothing_and_keeps_contract_hash(self):
        config = ProjectConfig(legacy_raw())
        self.assertIsNone(config.font_resolution)
        self.assertEqual(config.fonts["body"], "仿宋_GB2312")
        with patch.dict(os.environ, {fonts.FONT_FALLBACKS_ENV: ""}):
            self.assertEqual(compute_config_contract_hash(ProjectConfig(legacy_raw())),
                             compute_config_contract_hash(config))

    def test_default_official_fallbacks(self):
        self.enable()
        config = ProjectConfig(legacy_raw())
        self.assertEqual(config.font_fallbacks, fonts.DEFAULT_LEGACY_FONT_FALLBACKS)
        self.assertEqual(config.font_resolution.substitutions, (("仿宋_GB2312", "仿宋"),))
        self.assertEqual(config.fonts["body"], "仿宋")
        self.assertEqual(config.fonts["h3"], "仿宋")
        self.assertEqual(config.fonts["h2"], "楷体_GB2312")  # 已安装，不替换
        self.assertEqual(config.resolved_format.styles["body"].run.east_asia, "仿宋")
        self.assertNotIn("仿宋_GB2312", {item.name for item in fonts.required_fonts(config)})

    def test_declared_fallbacks_replace_defaults(self):
        self.enable()
        config = ProjectConfig(legacy_raw(
            fonts={"body": "P6缺失测试字体"},
            font_fallbacks={"P6缺失测试字体": ["宋体"]},
        ))
        self.assertEqual(config.fonts["body"], "宋体")
        # 声明了 font_fallbacks 就不再使用内置列表，仿宋_GB2312 保持缺失
        self.assertEqual(config.font_resolution.missing_unreplaced, ("仿宋_GB2312",))
        self.assertEqual(config.fonts["h3"], "仿宋_GB2312")

    def test_contract_hash_records_substitution(self):
        disabled = compute_config_contract_hash(ProjectConfig(legacy_raw()))
        self.enable()
        self.assertNotEqual(compute_config_contract_hash(ProjectConfig(legacy_raw())), disabled)

    def test_failures_are_config_errors(self):
        with self.assertRaises(ConfigError):
            ProjectConfig(legacy_raw(font_fallbacks={"仿宋_GB2312": "仿宋"}))
        self.enable()
        with self.assertRaisesRegex(ConfigError, "声明的回退字体也都没有安装"):
            ProjectConfig(legacy_raw(font_fallbacks={"仿宋_GB2312": ["P6另一个缺失字体"]}))


class FormatPackageTest(FakeFontsCase):
    def write_format(self, name, data):
        path = self.tmp / name
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return path

    def base_format(self):
        return {
            "format_schema_version": 1, "id": "fallback-base", "version": "1.0.0",
            "extends": "preset:report-basic@1.0.0",
            "styles": {"body": {"run": {"east_asia": "仿宋_GB2312"}}},
            "font_fallbacks": {"仿宋_GB2312": ["仿宋"], "楷体_GB2312": ["楷体"]},
        }

    def test_inheritance_and_overrides(self):
        self.write_format("base.json", self.base_format())
        child = self.write_format("child.json", {
            "format_schema_version": 1, "id": "fallback-child", "version": "1.0.0",
            "extends": "base.json", "font_fallbacks": {"楷体_GB2312": ["宋体"]},
        })
        resolved = resolve_format_package(str(child), overrides={"font_fallbacks": {"黑体": ["宋体"]}})
        self.assertEqual(resolved.font_fallbacks,
                         {"仿宋_GB2312": ("仿宋",), "楷体_GB2312": ("宋体",), "黑体": ("宋体",)})

    def test_schema_rejects_invalid_shape(self):
        bad = self.base_format()
        bad["font_fallbacks"] = {"仿宋_GB2312": "仿宋"}
        with self.assertRaises(ConfigError):
            resolve_format_package(str(self.write_format("bad.json", bad)))
        with self.assertRaises(ConfigError):
            resolve_format_package(str(self.write_format("ok.json", self.base_format())),
                                   overrides={"font_fallbacks": {"黑体": ["黑体"]}})

    def test_v3_config_applies_format_fallbacks(self):
        self.write_format("base.json", self.base_format())
        raw = {
            "schema_version": 3, "project_name": "Fictional Fonts",
            "format": {"ref": "base.json"},
            "formatting": {"mode": "restyle"},
            "source": {"strategy": "directory_tree"},
            "output": {"documents": [{"id": "main", "filename": "main.docx", "parts": ["body"]}]},
        }
        config = ProjectConfig(raw, self.tmp)
        self.assertIsNone(config.font_resolution)
        self.enable()
        config = ProjectConfig(raw, self.tmp)
        self.assertEqual(config.resolved_format.styles["body"].run.east_asia, "仿宋")
        self.assertEqual(config.fonts["body"], "仿宋")
        self.assertIn(("仿宋_GB2312", "仿宋"), config.font_resolution.substitutions)


class RecordTest(FakeFontsCase):
    def test_record_without_fallbacks(self):
        record = fonts.font_build_record(ProjectConfig(legacy_raw()))
        self.assertFalse(record["fallbacks_enabled"])
        self.assertEqual(record["substitutions"], [])
        self.assertEqual(record["missing_unreplaced"], ["仿宋_GB2312"])
        (line,) = fonts.build_summary_lines(record)
        self.assertIn("本机缺少且未替换: 仿宋_GB2312", line)
        self.assertIn("--font-fallbacks", line)
        json.dumps(record, ensure_ascii=False)

    def test_record_with_fallbacks(self):
        self.enable()
        record = fonts.font_build_record(ProjectConfig(legacy_raw()))
        self.assertTrue(record["fallbacks_enabled"])
        self.assertEqual(record["substitutions"],
                         [{"requested": "仿宋_GB2312", "used": "仿宋", "reason": "missing"}])
        self.assertIn({"name": "仿宋_GB2312", "status": "missing"}, record["checked"])
        self.assertEqual(fonts.build_summary_lines(record), ["[字体] 回退: 仿宋_GB2312 → 仿宋（本机缺少 仿宋_GB2312）"])

    def test_build_metadata_has_fonts(self):
        from lib.engine import generate_build_metadata

        self.enable()
        metadata = generate_build_metadata(ProjectConfig(legacy_raw()), {}, {})
        self.assertEqual(metadata["fonts"]["substitutions"][0]["used"], "仿宋")

    def test_plan_shows_substitution(self):
        self.enable()
        plan_text = io.StringIO()
        from synthesize import print_plan

        with contextlib.redirect_stdout(plan_text):
            print_plan(MINIMAL_DEMO / "source", MINIMAL_DEMO / "manifest.json")
        self.assertIn("[回退] 仿宋_GB2312 → 仿宋（本机缺少 仿宋_GB2312）", plan_text.getvalue())

    def test_cli_flag_sets_environment(self):
        import synthesize

        os.environ.pop(fonts.FONT_FALLBACKS_ENV)
        argv = ["synthesize.py", "--plan", "--source", str(MINIMAL_DEMO / "source"),
                "--manifest", str(MINIMAL_DEMO / "manifest.json"), "--font-fallbacks"]
        output = io.StringIO()
        with patch.object(sys, "argv", argv), contextlib.redirect_stdout(output):
            self.assertEqual(synthesize.main(), 0)
        self.assertEqual(os.environ.get(fonts.FONT_FALLBACKS_ENV), "1")
        self.assertIn("[回退] 仿宋_GB2312 → 仿宋", output.getvalue())


if __name__ == "__main__":
    unittest.main()
