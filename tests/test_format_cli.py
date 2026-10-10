# -*- coding: utf-8 -*-
"""
CLI 格式分析与编译子命令测试 (tests/test_format_cli.py)
验证 synthesize.py 分发的 --analyze-format、--compile-format 与 --compile-mapping
的参数互斥、输出防覆盖与端到端运行。
"""

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt


class FormatCLITest(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)

        # 创建测试 DOCX
        self.sample_docx = self.test_dir / "sample.docx"
        doc = Document()
        p = doc.add_paragraph("一级标题")
        p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        doc.add_paragraph("这是正文内容段落。")
        doc.save(str(self.sample_docx))

    def tearDown(self):
        self.temp_dir.cleanup()

    def _run_cli(self, args):
        cmd = [sys.executable, "synthesize.py"] + args
        return subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(Path(__file__).resolve().parent.parent),
        )

    def test_analyze_format_cli_and_overwrite_protection(self):
        """测试 --analyze-format 正常输出及防覆盖保护"""
        analysis_dir = self.test_dir / "analysis_out"

        # 1. 正常执行分析
        res = self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(analysis_dir)])
        self.assertEqual(res.returncode, 0)
        self.assertTrue((analysis_dir / "analysis.json").is_file())
        self.assertTrue((analysis_dir / "review.html").is_file())

        # 2. 再次执行（无 --replace-output），应被拦截
        res_blocked = self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(analysis_dir)])
        self.assertNotEqual(res_blocked.returncode, 0)
        self.assertIn("已存在", res_blocked.stdout + res_blocked.stderr)

        # 3. 加上 --replace-output，执行成功
        res_replace = self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(analysis_dir), "--replace-output"])
        self.assertEqual(res_replace.returncode, 0)

    def test_compile_format_cli(self):
        """测试 --compile-format 端到端编译格式包"""
        analysis_dir = self.test_dir / "analysis_compile"
        self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(analysis_dir)])

        report_json = analysis_dir / "analysis.json"
        report_data = json.loads(report_json.read_text(encoding="utf-8"))

        decisions_json = self.test_dir / "decisions.json"
        decisions_data = {
            "report_id": report_data["report_id"],
            "source_sha256": report_data["source_sha256"],
            "role_styles": {
                "body": report_data["candidate_roles"]["body"]["cluster_id"],
                "heading.1": report_data["candidate_roles"]["heading.1"]["cluster_id"],
            },
            "style_overrides": {},
            "missing_roles": {},
        }
        decisions_json.write_text(json.dumps(decisions_data), encoding="utf-8")

        format_out = self.test_dir / "my_format.json"
        res = self._run_cli([
            "--compile-format", str(report_json),
            "--decisions", str(decisions_json),
            "--format-id", "my-test-fmt",
            "--format-version", "1.0.0",
            "--format-out", str(format_out),
        ])
        self.assertEqual(res.returncode, 0)
        self.assertTrue(format_out.is_file())

        pkg = json.loads(format_out.read_text(encoding="utf-8"))
        self.assertEqual(pkg["id"], "my-test-fmt")

    def test_mutual_exclusion_with_standard_build_flags(self):
        """测试分析/编译子命令与构建/计划参数混用时报错拦截"""
        res = self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(self.test_dir), "--plan"])
        self.assertNotEqual(res.returncode, 0)
        self.assertIn("混用", res.stderr + res.stdout)

        res2 = self._run_cli(["--analyze-format", str(self.sample_docx), "--analysis-dir", str(self.test_dir), "--all"])
        self.assertNotEqual(res2.returncode, 0)


if __name__ == "__main__":
    unittest.main()
