# -*- coding: utf-8 -*-
"""P4 验收矩阵脚本的离线测试：跨平台页码对照与证据脱敏。"""

import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

_SPEC = importlib.util.spec_from_file_location(
    "word_matrix", ROOT / "docs" / "acceptance" / "public-release" / "word_matrix.py"
)
word_matrix = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(word_matrix)


def _case(case_id, status="verified", pages=None, count=3, labels=None):
    return {
        "id": case_id,
        "status": status,
        "page_map": pages or {"_Synth_body": {"physical_page": 2, "printed_page": 1, "verification_status": "verified"}},
        "export": {"page_count": count, "page_labels": labels or {"1": "1"}},
    }


class CrossPlatformCompareTests(unittest.TestCase):
    def test_identical_pages_ignore_platform_neutral_fields(self) -> None:
        left = {"label": "macos-word", "cases": [_case("WW01")]}
        right_case = _case("WW01")
        right_case["page_map"]["_Synth_body"]["verification_status"] = "unverified"
        right = {"label": "windows-word", "cases": [right_case]}
        report = word_matrix.compare_results(left, right)
        self.assertEqual(report["summary"], {"cases": 1, "comparable": 1, "identical": 1, "with_differences": 0})

    def test_path_hashed_bookmark_names_are_normalized(self) -> None:
        """重名书签的改名含源文件路径哈希，跨机器必然不同；只要页码一致就不算差异。"""
        def pages(first, second):
            return {
                f"_SynthBm_{first}_06": {"physical_page": 18, "printed_page": 18},
                f"_SynthBm_{second}_06": {"physical_page": 30, "printed_page": 30},
            }
        left = {"label": "macos-word", "cases": [_case("WW09", pages=pages("aaaaaaaaaa", "bbbbbbbbbb"))]}
        right = {"label": "windows-word", "cases": [_case("WW09", pages=pages("cccccccccc", "dddddddddd"))]}
        self.assertEqual(word_matrix.compare_results(left, right)["summary"]["identical"], 1)
        right["cases"][0]["page_map"]["_SynthBm_dddddddddd_06"]["physical_page"] = 31
        self.assertEqual(word_matrix.compare_results(left, right)["summary"]["with_differences"], 1)

    def test_page_and_export_differences_are_reported(self) -> None:
        left = {"label": "macos-word", "cases": [_case("WW01"), _case("WW02", count=4)]}
        right = {"label": "windows-word", "cases": [
            _case("WW01", pages={"_Synth_body": {"physical_page": 3, "printed_page": 1}}),
            _case("WW02", count=5, labels={"1": "i"}),
        ]}
        report = word_matrix.compare_results(left, right)
        self.assertEqual(report["summary"]["with_differences"], 2)
        ww01 = report["cases"][0]["differences"][0]
        self.assertEqual(ww01["bookmark"], "_Synth_body")
        self.assertEqual(ww01["macos-word"]["physical_page"], 2)
        self.assertEqual(ww01["windows-word"]["physical_page"], 3)
        kinds = {next(iter(item)) for item in report["cases"][1]["differences"]}
        self.assertEqual(kinds, {"export_page_count", "export_page_labels"})

    def test_non_verified_or_missing_cases_are_not_compared(self) -> None:
        left = {"label": "macos-word", "cases": [_case("WW01"), _case("WW02")]}
        right = {"label": "windows-word", "cases": [_case("WW01", status="failed")]}
        report = word_matrix.compare_results(left, right)
        self.assertEqual(report["summary"]["comparable"], 0)
        self.assertEqual(report["summary"]["cases"], 2)


class ScrubberTests(unittest.TestCase):
    def test_paths_and_emails_are_replaced(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            work = Path(tmp) / "ww01"
            scrub = word_matrix.Scrubber(work)
            text = scrub.text(
                f"{word_matrix.REPO}/synthesize.py {work}/out "
                r"C:\Users\someone\AppData\x /Users/someone/Library a.b@example.com"
            )
        self.assertNotIn(str(word_matrix.REPO), text)
        self.assertIn("<repo>", text)
        self.assertIn("<work>", text)
        self.assertNotRegex(text, r"[A-Za-z]:\\Users\\|/Users/")
        self.assertNotIn("example.com", text)
        self.assertEqual(scrub.data({"k": [f"{work}/a"]}), {"k": ["<work>/a"]})


if __name__ == "__main__":
    unittest.main()
