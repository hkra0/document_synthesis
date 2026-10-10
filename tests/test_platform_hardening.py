"""Tests for cross-platform hardening (R1, R2, R4, R5, R8)."""

import os
import re
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from lib.engine import BuildError, UnifiedSynthesizer, _replace_with_retry, _source_docx_paths
from lib.pagination import OfficeExportError, _replace_pdf_with_retry
from synthesize import run_doctor


class PlatformHardeningTests(unittest.TestCase):
    def test_file_lock_retry_succeeds_after_transient_permission_error(self):
        """测试目标文件短暂被锁时，退避重试能够自愈并成功替换。"""
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "source.docx"
            dst = Path(tmp) / "target.docx"
            src.write_text("source-content", encoding="utf-8")
            dst.write_text("old-content", encoding="utf-8")

            call_count = 0
            real_replace = os.replace

            def mock_replace(s, d):
                nonlocal call_count
                call_count += 1
                if call_count < 3:
                    raise PermissionError(13, "Permission denied (file in use)")
                return real_replace(s, d)

            sleep_calls = []

            with patch("os.replace", side_effect=mock_replace):
                _replace_with_retry(
                    src,
                    dst,
                    max_retries=5,
                    initial_delay=0.01,
                    backoff_factor=2.0,
                    sleep_func=lambda d: sleep_calls.append(d),
                )

            self.assertEqual(call_count, 3)
            self.assertEqual(len(sleep_calls), 2)
            self.assertEqual(dst.read_text(encoding="utf-8"), "source-content")

    def test_file_lock_retry_exhaustion_raises_clear_build_error(self):
        """测试重试耗尽后，抛出指明目标文件可能正在 Word 中打开的清晰错误。"""
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "source.docx"
            dst = Path(tmp) / "target.docx"
            src.write_text("source-content", encoding="utf-8")
            dst.write_text("old-content", encoding="utf-8")

            with patch("os.replace", side_effect=PermissionError(13, "File in use")):
                with self.assertRaises(BuildError) as ctx:
                    _replace_with_retry(
                        src,
                        dst,
                        max_retries=3,
                        initial_delay=0.01,
                        sleep_func=lambda _: None,
                    )
            self.assertIn("Word 中打开", str(ctx.exception))
            self.assertIn(str(dst), str(ctx.exception))

    def test_pagination_replace_pdf_retry_exhaustion_raises_office_error(self):
        """测试分页 PDF 替换重试耗尽后，抛出指明可能正在 PDF 阅读器或 Word 中打开的错误。"""
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "source.pdf"
            dst = Path(tmp) / "target.pdf"
            src.write_text("pdf-content", encoding="utf-8")
            dst.write_text("old-pdf", encoding="utf-8")

            with patch("os.replace", side_effect=PermissionError(13, "Access is denied")):
                with self.assertRaises(OfficeExportError) as ctx:
                    _replace_pdf_with_retry(
                        src,
                        dst,
                        max_retries=3,
                        initial_delay=0.01,
                        sleep_func=lambda _: None,
                    )
            self.assertIn("Word 或 PDF 阅读器", str(ctx.exception))

    def test_case_insensitive_docx_discovery_in_engine(self):
        """测试输入材料中 .DOCX、.Docx 等大写/混合大小写扩展名能被正确发现。"""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            doc_upper = root / "sample.DOCX"
            doc_upper.write_text("dummy", encoding="utf-8")

            from lib.config import ProjectConfig

            config = ProjectConfig({
                "schema_version": 2,
                "project_name": "case-test",
                "source": {"strategy": "docx_document"},
            })
            found = _source_docx_paths(root, config)
            self.assertEqual(found, [doc_upper.resolve()])

    def test_doctor_flags_overlong_path(self):
        """测试路径长度超过 240 字符时，run_doctor 产生超长预警检查失败项。"""
        long_dir_name = "a" * 250
        with tempfile.TemporaryDirectory() as tmp:
            long_dir = Path(tmp) / long_dir_name
            # Don't create actual overlong path on disk if OS restricts it; use mock Path
            mock_path = MagicMock(spec=Path)
            mock_path.resolve.return_value = Path("/" + "x" * 245)
            mock_path.exists.return_value = True

            with patch("sys.stdout"):
                exit_code = run_doctor(source_dir=mock_path)
            self.assertEqual(exit_code, 1)
