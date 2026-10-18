# -*- coding: utf-8 -*-
"""P6 F2 字体检查接入 --plan 与 --doctor。

覆盖：
1. report_lines：各状态的标签、替换提示只在缺失或无法判断时出现、索引不完整的警告
2. 计划带 fonts；缺失字体只出现在字体一节，不影响计划本身
3. --plan 终端输出字体一节
4. --doctor：无 --source 时检查默认配置；缺失字体不改变退出码；
   后端提供 Word 字体列表时用于交叉核对，读取失败时如实说明；其他后端不读取
"""

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from lib import fonts  # noqa: E402
from lib.engine import UnifiedSynthesizer  # noqa: E402
from lib.office import set_backend  # noqa: E402
from lib.office.base import BackendStatus  # noqa: E402
from lib.office.null import NullBackend  # noqa: E402
from synthesize import print_plan, run_doctor  # noqa: E402
from test_fonts import _win, build_font, make_face  # noqa: E402

MINIMAL_DEMO = ROOT / "examples" / "minimal-demo"
INSTALLED_NAMES = ["仿宋_GB2312", "楷体_GB2312", "方正小标宋简体", "黑体", "宋体", "Times New Roman"]


def report(*items, complete=True):
    return {
        "index_complete": complete,
        "index_error": None if complete else "没有找到任何字体文件",
        "word_cross_checked": False,
        "fonts": [{"name": name, "status": status, "sources": sources, "files": []}
                  for name, status, sources in items],
    }


class ReportLinesTest(unittest.TestCase):
    def test_labels_and_notice(self):
        lines = fonts.report_lines(report(
            ("仿宋_GB2312", "installed", ["样式 body（中文）"]),
            ("Aptos", "word_only", ["样式 x（西文）"]),
        ))
        self.assertEqual([label for label, _ in lines], ["已安装", "仅 Word 可用"])
        self.assertIn("云字体", lines[1][1])

        lines = fonts.report_lines(report(("P6缺失测试字体", "missing", ["样式 body（中文）"])))
        self.assertEqual(lines[-1], ("提示", fonts.FONT_ENVIRONMENT_NOTICE))

    def test_incomplete_index_is_warned(self):
        lines = fonts.report_lines(report(("仿宋", "unknown", ["fonts.body"]), complete=False))
        self.assertEqual(lines[0][0], "警告")
        self.assertIn("没有找到任何字体文件", lines[0][1])
        self.assertEqual(lines[1][0], "无法判断")
        self.assertEqual(lines[-1][0], "提示")

    def test_long_source_lists_are_shortened(self):
        sources = [f"样式 s{i}（中文）" for i in range(5)]
        self.assertEqual(fonts.describe_sources(sources), "样式 s0（中文）、样式 s1（中文）、样式 s2（中文） 等 5 处")


class FontReportingBase(unittest.TestCase):
    """用合成字体代替本机字体，结果与运行环境无关。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        tmp = Path(self._tmp.name)
        files = []
        for index, name in enumerate(INSTALLED_NAMES):
            path = tmp / f"font{index}.ttf"
            path.write_bytes(build_font([make_face([_win(0x0804, 1, name)])]))
            files.append((path, "test"))
        self.index = fonts.FontIndex.build(files)
        self._patch = patch.object(fonts, "installed_fonts", lambda: self.index)
        self._patch.start()

    def tearDown(self):
        self._patch.stop()
        set_backend(None)
        self._tmp.cleanup()

    def missing_body_override(self):
        path = Path(self._tmp.name) / "missing.override.json"
        path.write_text(json.dumps({"fonts": {"body": "P6缺失测试字体"}}, ensure_ascii=False), encoding="utf-8")
        return path

    def statuses(self, fonts_report):
        return {item["name"]: item["status"] for item in fonts_report["fonts"]}


class PlanFontsTest(FontReportingBase):
    def test_plan_reports_fonts(self):
        plan = UnifiedSynthesizer.plan(MINIMAL_DEMO / "source", MINIMAL_DEMO / "manifest.json")
        self.assertEqual(set(self.statuses(plan["fonts"]).values()), {"installed"})
        self.assertFalse(plan["fonts"]["word_cross_checked"])
        json.dumps(plan["fonts"], ensure_ascii=False)

    def test_missing_font_is_reported(self):
        plan = UnifiedSynthesizer.plan(MINIMAL_DEMO / "source", MINIMAL_DEMO / "manifest.json",
                                       override_paths=[self.missing_body_override()])
        statuses = self.statuses(plan["fonts"])
        self.assertEqual(statuses["P6缺失测试字体"], "missing")
        self.assertEqual(statuses["仿宋_GB2312"], "installed")  # h3 仍使用默认字体

    def test_print_plan_has_font_section(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            print_plan(MINIMAL_DEMO / "source", MINIMAL_DEMO / "manifest.json",
                       override_paths=[self.missing_body_override()])
        text = output.getvalue()
        self.assertIn("字体（按本机字体文件判断，不启动 Word）:", text)
        self.assertIn("  [缺失] P6缺失测试字体  ← ", text)
        self.assertIn(fonts.FONT_ENVIRONMENT_NOTICE, text)
        self.assertLess(text.index("字体（"), text.index("目录预览:"))


class WordListBackend(NullBackend):
    """能精确分页并提供 Word 字体列表的假后端。"""
    name = "fake-word"

    def __init__(self, names):
        self.names = names
        self.calls = 0

    def static_status(self):
        return BackendStatus(True, "fake", "exact")

    def probe(self):
        return BackendStatus(True, "fake probe ok", "exact")

    def font_names(self):
        self.calls += 1
        return self.names


class DoctorFontsTest(FontReportingBase):
    def doctor(self, **kwargs):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            code = run_doctor(**kwargs)
        return code, output.getvalue()

    def test_without_source_checks_default_config(self):
        set_backend(NullBackend())
        code, text = self.doctor()
        self.assertEqual(code, 2)
        self.assertIn("字体（默认配置；", text)
        self.assertIn("[已安装] 方正小标宋简体", text)

    def test_missing_font_does_not_change_exit_code(self):
        set_backend(WordListBackend(INSTALLED_NAMES))
        code, text = self.doctor(source_dir=MINIMAL_DEMO / "source", manifest_path=MINIMAL_DEMO / "manifest.json",
                                 override_paths=[self.missing_body_override()])
        self.assertEqual(code, 0)
        self.assertIn("字体（本项目配置）:", text)
        self.assertIn("[缺失] P6缺失测试字体", text)

    def test_word_font_names_cross_check(self):
        backend = WordListBackend(INSTALLED_NAMES + ["P6缺失测试字体"])
        set_backend(backend)
        code, text = self.doctor(source_dir=MINIMAL_DEMO / "source", manifest_path=MINIMAL_DEMO / "manifest.json",
                                 override_paths=[self.missing_body_override()])
        self.assertEqual(code, 0)
        self.assertEqual(backend.calls, 1)
        self.assertIn("[信息] Word 字体列表: 7 个，用于交叉核对字体", text)
        self.assertIn("[仅 Word 可用] P6缺失测试字体", text)

    def test_word_font_names_failure_is_reported(self):
        set_backend(WordListBackend(None))
        _code, text = self.doctor()
        self.assertIn("[信息] 未能读取 Word 字体列表，字体只按本机字体文件判断", text)
        self.assertNotIn("仅 Word 可用", text)

    def test_backend_without_font_names_is_not_asked(self):
        set_backend(NullBackend())
        _code, text = self.doctor()
        self.assertNotIn("Word 字体列表", text)


if __name__ == "__main__":
    unittest.main()
