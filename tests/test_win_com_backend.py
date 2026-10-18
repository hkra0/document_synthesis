# -*- coding: utf-8 -*-
"""P3 Windows COM 后端离线测试。

不需要 Microsoft Word：worker 的 COM 调用用假对象驱动，父进程调度用假 runner
或替身子进程驱动。CI 的 Windows runner 没有 Word，也能完整运行本文件。
覆盖：
1. HRESULT / com_error 翻译为中文提示
2. worker 各操作的 COM 调用序列、数值常量与 finally 中的关闭/退出
3. 书签缺失、导出失败时仍关闭文档并退出 Word
4. PowerPoint 单实例：用户已打开时不退出应用
5. 父进程：成功、失败、异常退出、超时；失败时不替换目标文件、不留工作目录
6. 只按 pid 文件中的 PID + 映像名 + 创建时间结束进程
7. 会话：多个请求复用同一个 worker 与 Word；请求间校验失败保留实例，COM 错误、崩溃、
   超时后回收并在下一个请求时重启；超时只报告当前请求的挂起步骤；用户文档移交而不退出；
   构建元数据不再启动 Word
"""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest
from typing import Any, Dict, List
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from lib.office import (
    OfficeBackendError,
    OfficeTimeoutError,
    OfficeUnsupportedError,
    get_backend,
    set_backend,
)
from lib.office import win_com, win_com_worker
from lib.office.win_com import WinComBackend, cleanup_recorded_office, com_failure_message, run_worker
from lib.office.win_com_worker import WorkerFailure, handle, translate_exception


# ---------------------------------------------------------------------------
# COM 假对象
# ---------------------------------------------------------------------------

class FakeRange:
    def __init__(self, doc: "FakeDocument", start: int) -> None:
        self.doc = doc
        self.Start = start

    @property
    def Range(self) -> "FakeRange":  # noqa: N802 - Bookmark.Range
        return self

    def Information(self, kind: int) -> int:  # noqa: N802
        physical, printed = self.doc.pages_by_start[self.Start]
        if kind == 3:
            return physical
        if kind == 1:
            return printed
        raise AssertionError(f"unexpected Information type {kind}")


class FakeBookmarks:
    def __init__(self, doc: "FakeDocument") -> None:
        self.doc = doc
        self.ShowHidden = False

    def Exists(self, name: str) -> bool:  # noqa: N802
        return name in self.doc.bookmark_starts

    def __call__(self, name: str) -> FakeRange:
        return FakeRange(self.doc, self.doc.bookmark_starts[name])


class FakeDocument:
    def __init__(self, app: "FakeWordApp", path: str) -> None:
        self.app = app
        self.path = path
        self.bookmark_starts = dict(app.bookmark_starts)
        self.pages_by_start = dict(app.pages_by_start)
        self.FullName = path
        self.Bookmarks = FakeBookmarks(self)
        self.calls: List[Any] = []
        self.closed_with = None

    def Repaginate(self) -> None:  # noqa: N802
        self.calls.append("Repaginate")
        if self.app.on_repaginate:
            self.app.on_repaginate(self)

    def Range(self, start: int, end: int) -> FakeRange:  # noqa: N802
        assert start == end
        return FakeRange(self, start)

    def ExportAsFixedFormat(self, **kwargs: Any) -> None:  # noqa: N802
        self.calls.append(("ExportAsFixedFormat", kwargs))
        if self.app.fail_export:
            raise self.app.fail_export
        Path(kwargs["OutputFileName"]).write_bytes(b"%PDF-fake")

    def SaveAs2(self, **kwargs: Any) -> None:  # noqa: N802
        self.calls.append(("SaveAs2", kwargs))
        Path(kwargs["FileName"]).write_bytes(b"PK-fake")

    def Close(self, SaveChanges: int) -> None:  # noqa: N802,N803
        self.closed_with = SaveChanges
        if self in self.app.open_documents:
            self.app.open_documents.remove(self)


class FakeDocuments:
    def __init__(self, app: "FakeWordApp") -> None:
        self.app = app

    @property
    def Count(self) -> int:  # noqa: N802
        return len(self.app.open_documents)

    def __call__(self, index: int) -> FakeDocument:
        return self.app.open_documents[index - 1]

    def Open(self, **kwargs: Any) -> FakeDocument:  # noqa: N802
        self.app.open_kwargs = kwargs
        doc = FakeDocument(self.app, kwargs["FileName"])
        self.app.documents.append(doc)
        self.app.open_documents.append(doc)
        return doc


class FakeUserDocument:
    """用户经资源管理器打开、被系统路由进自动化实例的文档。"""

    def __init__(self, path: str) -> None:
        self.FullName = path
        self.closed_with = None

    def Close(self, SaveChanges: int) -> None:  # noqa: N802,N803
        raise AssertionError("worker 不应关闭用户的文档")


class FakeWordApp:
    Version = "16.0"
    Build = "16.0.99999"
    Path = r"C:\Program Files\Microsoft Office\root\Office16"
    FontNames = ("仿宋_GB2312", "@仿宋_GB2312", "Times New Roman")

    def __init__(self, bookmark_starts=None, pages_by_start=None, fail_export=None, on_repaginate=None) -> None:
        self.bookmark_starts = bookmark_starts or {}
        self.pages_by_start = pages_by_start or {}
        self.fail_export = fail_export
        self.on_repaginate = on_repaginate
        self.Documents = FakeDocuments(self)
        self.documents: List[FakeDocument] = []
        self.open_documents: List[Any] = []
        self.quit_with = None
        self.quit_count = 0
        self.Visible = True
        self.DisplayAlerts = -1
        self.open_kwargs: Dict[str, Any] = {}

    def Quit(self, SaveChanges: int = None) -> None:  # noqa: N802,N803
        self.quit_with = SaveChanges
        self.quit_count += 1


class FakePresentation:
    def __init__(self, app: "FakePowerPointApp") -> None:
        self.app = app
        self.closed = False

    def SaveAs(self, path: str, fmt: int) -> None:  # noqa: N802
        self.app.saved = (path, fmt)
        Path(path).write_bytes(b"%PDF-fake")

    def Close(self) -> None:  # noqa: N802
        self.closed = True


class FakePresentations:
    def __init__(self, app: "FakePowerPointApp") -> None:
        self.app = app

    def Open(self, *args: Any) -> FakePresentation:  # noqa: N802
        self.app.open_args = args
        self.app.presentation = FakePresentation(self.app)
        return self.app.presentation


class FakePowerPointApp:
    def __init__(self) -> None:
        self.Presentations = FakePresentations(self)
        self.quit_called = False
        self.saved = None

    def Quit(self) -> None:  # noqa: N802
        self.quit_called = True


class FakeComError(Exception):
    """模拟 pywin32 ``com_error``：args = (hresult, message, excepinfo, argerror)。"""


def _no_pid(_app: Any) -> int:
    return 0


# ---------------------------------------------------------------------------
# 测试
# ---------------------------------------------------------------------------

class ComErrorTranslationTests(unittest.TestCase):
    def test_known_hresults_are_translated(self) -> None:
        cases = {
            -2147221005: "未检测到已注册",        # CO_E_CLASSSTRING (signed)
            0x80040154: "未检测到已注册",
            0x80080005: "首次运行",
            -2147417846: "对象被占用 0x8001010A",  # RPC_E_SERVERCALL_RETRYLATER (signed)
            0x800706BA: "意外退出",
            0x80010108: "意外退出",
        }
        for code, expected in cases.items():
            with self.subTest(code=hex(code & 0xFFFFFFFF)):
                self.assertIn(expected, com_failure_message(code, ""))

    def test_protected_view_and_activation_by_description(self) -> None:
        self.assertIn("受保护视图", com_failure_message(None, "The file could not be opened in Protected View."))
        self.assertIn("未激活", com_failure_message(None, "This product is unlicensed."))
        self.assertIn("PowerPoint", com_failure_message(0x80080005, "", "Microsoft PowerPoint"))

    def test_translate_prefers_specific_scode(self) -> None:
        excepinfo = (0, "Microsoft Word", "The file could not be opened in Protected View.", None, 0, -2146822496)
        exc = FakeComError(-2147352567, "Exception occurred.", excepinfo, None)
        self.assertIn("受保护视图", translate_exception(exc, "Microsoft Word"))

        disconnected = FakeComError(-2147023174, "The RPC server is unavailable.", None, None)
        self.assertIn("意外退出", translate_exception(disconnected, "Microsoft Word"))

        generic = FakeComError(-2147352567, "Exception occurred.", (0, "Microsoft Word", "Bad thing", None, 0, -2146822000), None)
        self.assertIn("Bad thing", translate_exception(generic, "Microsoft Word"))


class WorkerComSequenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _dispatch(self, app: Any):
        def dispatch(progid: str) -> Any:
            self.assertEqual(progid, "Word.Application")
            return app
        return dispatch

    def test_inspect_measures_bookmarks_and_exports_pdf(self) -> None:
        app = FakeWordApp(
            bookmark_starts={"_Synth_toc": 10, "_Synth_body": 200},
            pages_by_start={10: (2, 1), 200: (4, 1)},
        )
        pdf = self.dir / "out.pdf"
        result = handle(
            {"op": "inspect", "docx": str(self.dir / "in.docx"), "pdf": str(pdf), "names": ["_Synth_toc", "_Synth_body"]},
            dispatch=self._dispatch(app),
            word_pid_of=_no_pid,
        )
        self.assertEqual(result, {"ok": True, "page_map": "_Synth_toc\t2\t1\n_Synth_body\t4\t1\n"})
        self.assertFalse(app.Visible)
        self.assertEqual(app.DisplayAlerts, 0)
        self.assertEqual(app.open_kwargs["ReadOnly"], True)
        self.assertEqual(app.open_kwargs["ConfirmConversions"], False)
        self.assertEqual(app.open_kwargs["AddToRecentFiles"], False)
        doc = app.documents[0]
        self.assertTrue(doc.Bookmarks.ShowHidden)
        self.assertEqual(doc.calls[0], "Repaginate")
        export = doc.calls[1][1]
        self.assertEqual(export["ExportFormat"], 17)
        self.assertEqual(export["CreateBookmarks"], 1)
        self.assertTrue(export["DocStructureTags"])
        self.assertTrue(export["BitmapMissingFonts"])
        self.assertTrue(pdf.is_file())
        self.assertEqual(doc.closed_with, 0)
        self.assertEqual(app.quit_with, 0)

    def test_missing_bookmark_still_closes_and_quits(self) -> None:
        app = FakeWordApp(bookmark_starts={"_Synth_body": 5}, pages_by_start={5: (1, 1)})
        with self.assertRaises(WorkerFailure) as ctx:
            handle(
                {"op": "inspect", "docx": "x", "pdf": str(self.dir / "o.pdf"), "names": ["_Synth_body", "_Synth_toc"]},
                dispatch=self._dispatch(app),
                word_pid_of=_no_pid,
            )
        self.assertIn("缺少书签: _Synth_toc", str(ctx.exception))
        self.assertEqual(app.documents[0].closed_with, 0)
        self.assertEqual(app.quit_with, 0)
        self.assertFalse((self.dir / "o.pdf").exists())

    def test_export_failure_still_closes_and_quits(self) -> None:
        app = FakeWordApp(fail_export=FakeComError(-2147023174, "The RPC server is unavailable.", None, None))
        with self.assertRaises(FakeComError):
            handle({"op": "export_pdf", "docx": "x", "pdf": str(self.dir / "o.pdf")},
                   dispatch=self._dispatch(app), word_pid_of=_no_pid)
        self.assertEqual(app.documents[0].closed_with, 0)
        self.assertEqual(app.quit_with, 0)

    def test_convert_doc_uses_document_default_format(self) -> None:
        app = FakeWordApp()
        dst = self.dir / "converted.docx"
        self.assertEqual(
            handle({"op": "convert_doc", "src": "in.doc", "dst": str(dst)}, dispatch=self._dispatch(app), word_pid_of=_no_pid),
            {"ok": True},
        )
        name, kwargs = app.documents[0].calls[0]
        self.assertEqual(name, "SaveAs2")
        self.assertEqual(kwargs["FileFormat"], 16)
        self.assertEqual(kwargs["FileName"], str(dst))

    def test_probe_reports_version_bitness_and_click_to_run(self) -> None:
        result = handle({"op": "probe"}, dispatch=self._dispatch(FakeWordApp()), word_pid_of=_no_pid)
        self.assertEqual(result["word_version"], "16.0")
        self.assertEqual(result["word_build"], "16.0.99999")
        self.assertEqual(result["word_bitness"], 64)
        self.assertTrue(result["click_to_run"])

    def test_font_names_lists_word_fonts_and_quits(self) -> None:
        app = FakeWordApp()
        result = handle({"op": "font_names"}, dispatch=self._dispatch(app), word_pid_of=_no_pid)
        self.assertEqual(result, {"ok": True, "font_names": ["仿宋_GB2312", "@仿宋_GB2312", "Times New Roman"]})
        self.assertEqual(app.quit_count, 1)

    def test_busy_office_is_retried_then_succeeds(self) -> None:
        """Word 正忙（RETRYLATER / CALL_REJECTED）时退避重试；其他错误不重试。"""
        attempts = []
        sleeps = []

        def flaky():
            attempts.append(1)
            if len(attempts) == 1:
                raise FakeComError(-2147417846, "Call was rejected by callee.", None, None)  # 0x8001010A
            if len(attempts) == 2:
                excepinfo = (0, "Microsoft Word", "busy", None, 0, -2147418111)  # scode 0x80010001
                raise FakeComError(-2147352567, "Exception occurred.", excepinfo, None)
            return "done"

        self.assertEqual(win_com_worker.call_with_busy_retry("测试步骤", flaky, sleep=sleeps.append), "done")
        self.assertEqual(len(attempts), 3)
        self.assertEqual(sleeps, [0.25, 0.5])

        other = FakeComError(-2147023174, "The RPC server is unavailable.", None, None)
        with self.assertRaises(FakeComError):
            win_com_worker.call_with_busy_retry("测试步骤", lambda: (_ for _ in ()).throw(other), sleep=self.fail)

    def test_busy_retry_gives_up_after_limit(self) -> None:
        clock = iter([0.0, 10.0, 70.0])
        busy = FakeComError(-2147417846, "Call was rejected by callee.", None, None)

        def always_busy():
            raise busy

        with self.assertRaises(FakeComError):
            win_com_worker.call_with_busy_retry(
                "测试步骤", always_busy, limit=60, sleep=lambda s: None, clock=lambda: next(clock)
            )

    def test_unknown_operation_rejected(self) -> None:
        with self.assertRaises(WorkerFailure):
            handle({"op": "format_c"})

    def test_powerpoint_owned_instance_quits(self) -> None:
        app = FakePowerPointApp()
        dst = self.dir / "slides.pdf"
        win_com_worker.convert_pptx(
            {"src": "in.pptx", "dst": str(dst)}, dispatch=lambda progid: app,
            pid_of=_no_pid, already_running=lambda image: False,
        )
        self.assertEqual(app.saved, (str(dst), 32))
        self.assertEqual(app.open_args, ("in.pptx", True, False, False))
        self.assertTrue(app.presentation.closed)
        self.assertTrue(app.quit_called)

    def test_powerpoint_user_instance_is_not_quit(self) -> None:
        app = FakePowerPointApp()
        win_com_worker.convert_pptx(
            {"src": "in.pptx", "dst": str(self.dir / "s.pdf")}, dispatch=lambda progid: app,
            pid_of=lambda a: self.fail("用户实例不应记录 PID"), already_running=lambda image: True,
        )
        self.assertTrue(app.presentation.closed)
        self.assertFalse(app.quit_called)


class FakeRunner:
    """替代 run_worker：按请求在工作目录中写出产物，或模拟失败。"""

    def __init__(self, page_map: str = "", error: Exception = None, pages: int = 3) -> None:
        self.page_map = page_map
        self.error = error
        self.pages = pages
        self.requests: List[Dict[str, Any]] = []

    def __call__(self, request: Dict[str, Any], timeout: float, pid_file: Path) -> Dict[str, Any]:
        self.requests.append(dict(request, timeout=timeout, pid_file=str(pid_file)))
        if self.error:
            raise self.error
        if request["op"] in ("inspect", "export_pdf"):
            import pymupdf
            doc = pymupdf.open()
            for _ in range(self.pages):
                doc.new_page()
            doc.save(request["pdf"])
            doc.close()
        if request["op"] in ("convert_doc", "convert_pptx"):
            Path(request["dst"]).write_bytes(b"converted")
        if request["op"] == "font_names":
            return {"ok": True, "font_names": ["仿宋_GB2312", "Aptos"]}
        return {"ok": True, "page_map": self.page_map}


class WinComBackendParentTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.work = self.dir / "word-access"
        self.patches = [
            patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR": str(self.work)}),
            patch.object(win_com, "on_windows", lambda: True),
            patch.object(win_com, "progid_registered", lambda progid: True),
            patch.object(win_com, "pywin32_available", lambda: True),
        ]
        for item in self.patches:
            item.start()
        self.docx = self.dir / "final.docx"
        self.docx.write_bytes(b"PK-docx")

    def tearDown(self) -> None:
        for item in reversed(self.patches):
            item.stop()
        set_backend(None)
        self.tmp.cleanup()

    def _leftover_jobs(self) -> List[Path]:
        return list(self.work.glob("job-*")) if self.work.exists() else []

    def test_inspect_bookmarks_success(self) -> None:
        runner = FakeRunner(page_map="_Synth_body\t2\t1\n")
        pdf = self.dir / "final.pdf"
        report = WinComBackend(runner=runner).inspect_bookmarks(self.docx, pdf, ["_Synth_body"])
        self.assertEqual(report.bookmarks, {"_Synth_body": {"physical_page": 2, "printed_page": 1}})
        self.assertEqual(report.total_pages, 3)
        self.assertEqual(report.fidelity, "exact")
        self.assertEqual(report.pdf_path, pdf.resolve())
        self.assertTrue(pdf.is_file())
        request = runner.requests[0]
        self.assertEqual(request["timeout"], 600)
        self.assertTrue(Path(request["docx"]).is_absolute())
        self.assertNotEqual(Path(request["docx"]), self.docx.resolve(), "Word 只能打开私有副本")
        self.assertEqual(self._leftover_jobs(), [])

    def test_invalid_page_map_keeps_existing_pdf(self) -> None:
        pdf = self.dir / "final.pdf"
        pdf.write_bytes(b"previous")
        runner = FakeRunner(page_map="_Synth_body\t9\t1\n", pages=3)  # 物理页超出总页数
        with self.assertRaises(OfficeBackendError):
            WinComBackend(runner=runner).inspect_bookmarks(self.docx, pdf, ["_Synth_body"])
        self.assertEqual(pdf.read_bytes(), b"previous")
        self.assertEqual(self._leftover_jobs(), [])

    def test_worker_failure_keeps_existing_pdf(self) -> None:
        pdf = self.dir / "final.pdf"
        pdf.write_bytes(b"previous")
        for error in (OfficeBackendError("Word 文档中缺少书签: x"), OfficeTimeoutError("等待超时")):
            with self.subTest(error=type(error).__name__):
                with self.assertRaises(type(error)):
                    WinComBackend(runner=FakeRunner(error=error)).inspect_bookmarks(self.docx, pdf, ["x"])
                self.assertEqual(pdf.read_bytes(), b"previous")
                self.assertEqual(self._leftover_jobs(), [])

    def test_conversions_and_export(self) -> None:
        backend = WinComBackend(runner=FakeRunner())
        doc_src = self.dir / "legacy.doc"
        doc_src.write_bytes(b"doc")
        backend.convert_doc_to_docx(doc_src, self.dir / "out" / "legacy.docx")
        self.assertEqual((self.dir / "out" / "legacy.docx").read_bytes(), b"converted")

        pptx = self.dir / "slides.pptx"
        pptx.write_bytes(b"pptx")
        backend.convert_pptx_to_pdf(pptx, self.dir / "out" / "slides.pdf")
        self.assertTrue((self.dir / "out" / "slides.pdf").is_file())

        backend.export_pdf(self.docx, self.dir / "export.pdf")
        self.assertTrue((self.dir / "export.pdf").is_file())
        self.assertEqual(self._leftover_jobs(), [])

    def test_missing_powerpoint_only_blocks_pptx(self) -> None:
        with patch.object(win_com, "progid_registered", lambda progid: progid == win_com.WORD_PROGID):
            backend = WinComBackend(runner=FakeRunner())
            self.assertTrue(backend.static_status().available)
            self.assertFalse(backend.static_status().details["powerpoint_registered"])
            pptx = self.dir / "slides.pptx"
            pptx.write_bytes(b"pptx")
            with self.assertRaises(OfficeUnsupportedError) as ctx:
                backend.convert_pptx_to_pdf(pptx, self.dir / "slides.pdf")
            self.assertIn("PowerPoint", str(ctx.exception))
            backend.export_pdf(self.docx, self.dir / "export.pdf")

    def test_static_status_distinguishes_missing_word_and_pywin32(self) -> None:
        with patch.object(win_com, "progid_registered", lambda progid: False):
            status = WinComBackend().static_status()
            self.assertFalse(status.available)
            self.assertIn("未检测到已安装的 Microsoft Word", status.reason)
        with patch.object(win_com, "pywin32_available", lambda: False):
            status = WinComBackend().static_status()
            self.assertFalse(status.available)
            self.assertIn("pywin32", status.reason)

    def test_font_names_runs_single_request_and_cleans_job(self) -> None:
        runner = FakeRunner()
        self.assertEqual(WinComBackend(runner=runner).font_names(), ["仿宋_GB2312", "Aptos"])
        self.assertEqual([request["op"] for request in runner.requests], ["font_names"])
        self.assertEqual(self._leftover_jobs(), [])

    def test_font_names_failure_returns_none(self) -> None:
        runner = FakeRunner(error=OfficeBackendError("无法启动 Microsoft Word。"))
        self.assertIsNone(WinComBackend(runner=runner).font_names())
        self.assertEqual(self._leftover_jobs(), [])
        with patch.object(win_com, "progid_registered", lambda progid: False):
            unused = FakeRunner()
            self.assertIsNone(WinComBackend(runner=unused).font_names())
            self.assertEqual(unused.requests, [])

    def test_probe_failure_is_reported_not_raised(self) -> None:
        status = WinComBackend(runner=FakeRunner(error=OfficeBackendError("无法启动 Microsoft Word。"))).probe()
        self.assertFalse(status.available)
        self.assertIn("无法启动", status.reason)

    def test_auto_backend_on_windows_is_win_com(self) -> None:
        with patch("lib.office.sys.platform", "win32"), patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_OFFICE_BACKEND": "auto"}):
            self.assertIsInstance(get_backend(), WinComBackend)
        self.assertIsInstance(get_backend("win-word"), WinComBackend)


class RecordedProcessCleanupTests(unittest.TestCase):
    def _pid_file(self, directory: str, record: Dict[str, Any]) -> Path:
        path = Path(directory) / "office.pid"
        path.write_text(json.dumps(record), encoding="utf-8")
        return path

    def test_matching_identity_is_terminated(self) -> None:
        killed = []
        with tempfile.TemporaryDirectory() as tmp:
            pid_file = self._pid_file(tmp, {"pid": 4242, "image": "winword.exe", "created": 7})
            result = cleanup_recorded_office(
                pid_file, 0, identity=lambda pid: ("winword.exe", 7), terminate=killed.append, sleep=lambda s: None
            )
        self.assertEqual(result, 4242)
        self.assertEqual(killed, [4242])

    def test_reused_pid_is_not_terminated(self) -> None:
        killed = []
        with tempfile.TemporaryDirectory() as tmp:
            pid_file = self._pid_file(tmp, {"pid": 4242, "image": "winword.exe", "created": 7})
            for identity in (("winword.exe", 8), ("notepad.exe", 7), None):
                cleanup_recorded_office(pid_file, 0, identity=lambda pid, i=identity: i, terminate=killed.append)
        self.assertEqual(killed, [])

    def test_orphan_from_stuck_launch_is_found_only_when_unambiguous(self) -> None:
        """Word 卡在启动阶段时只有启动前快照：新出现、由 COM 启动、唯一的进程才会被结束。"""
        record = {"pending": True, "image": "winword.exe", "before": [100], "started": 50}
        identities = {100: ("winword.exe", 10), 200: ("winword.exe", 60), 300: ("winword.exe", 70), 400: ("winword.exe", 40)}

        def find(pids, manual=()):
            return win_com.find_orphan_office(
                record, list_pids=lambda image: pids, identity=identities.get, automation=lambda pid: pid not in manual
            )

        self.assertEqual(find([100, 200]), 200)
        self.assertIsNone(find([100]), "启动前已存在的实例不是候选")
        self.assertIsNone(find([100, 400]), "启动前创建的进程不是候选")
        self.assertIsNone(find([100, 200], manual={200}), "用户手动打开的 Word 命令行不带 /Automation")
        self.assertIsNone(find([100, 200, 300]), "多个候选时不结束任何进程")
        self.assertEqual(find([100, 200, 300], manual={300}), 200)

    @unittest.skipUnless(sys.platform == "win32", "读取进程命令行只在 Windows 上实现")
    def test_process_command_line_reads_own_process(self) -> None:
        command_line = win_com.process_command_line(os.getpid()) or ""
        self.assertIn("python", command_line.lower())
        self.assertFalse(win_com.is_automation_instance(os.getpid()))

    def test_pending_record_cleanup_terminates_orphan(self) -> None:
        killed = []
        with tempfile.TemporaryDirectory() as tmp:
            pid_file = self._pid_file(tmp, {"pending": True, "image": "winword.exe", "before": [], "started": 1})
            with patch.object(win_com, "find_orphan_office", lambda record, identity: 777):
                self.assertEqual(cleanup_recorded_office(pid_file, 0, terminate=killed.append), 777)
        self.assertEqual(killed, [777])

    def test_late_launched_orphan_is_still_terminated(self) -> None:
        """worker 在 DispatchEx 返回前被结束，Word 稍后才由 DCOM 拉起：在等待期内找到即结束。"""
        killed = []
        polls = iter([None, None, 888])
        with tempfile.TemporaryDirectory() as tmp:
            pid_file = self._pid_file(tmp, {"pending": True, "image": "winword.exe", "before": [], "started": 1})
            result = cleanup_recorded_office(
                pid_file, 0, terminate=killed.append, sleep=lambda s: None,
                orphan_wait=30, find_orphan=lambda record: next(polls),
            )
        self.assertEqual(result, 888)
        self.assertEqual(killed, [888])

    def test_pending_record_without_wait_checks_once(self) -> None:
        calls = []
        with tempfile.TemporaryDirectory() as tmp:
            pid_file = self._pid_file(tmp, {"pending": True, "image": "winword.exe", "before": [], "started": 1})
            self.assertIsNone(cleanup_recorded_office(
                pid_file, 0, terminate=self.fail, sleep=self.fail, find_orphan=lambda record: calls.append(1),
            ))
        self.assertEqual(calls, [1])

    def test_no_pid_file_does_nothing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertIsNone(cleanup_recorded_office(Path(tmp) / "missing.pid", 0, terminate=self.fail))


class RunWorkerProcessTests(unittest.TestCase):
    """用替身子进程验证父进程的超时、异常退出与结果解析。"""

    def _run(self, code: str, timeout: float = 30, pid_file: Path = None) -> Dict[str, Any]:
        with tempfile.TemporaryDirectory() as tmp:
            return run_worker({"op": "probe"}, timeout, pid_file or Path(tmp) / "office.pid",
                              command=[sys.executable, "-c", code])

    def test_success_result_line(self) -> None:
        code = "import sys, json; sys.stdin.readline(); print('noise'); print(json.dumps({'ok': True, 'word_version': '16.0'}))"
        self.assertEqual(self._run(code)["word_version"], "16.0")

    def test_error_kinds(self) -> None:
        unsupported = "import json; print(json.dumps({'ok': False, 'kind': 'unsupported', 'error': '缺少 pywin32'}))"
        with self.assertRaises(OfficeUnsupportedError):
            self._run(unsupported)
        failed = "import json; print(json.dumps({'ok': False, 'kind': 'error', 'error': 'Word 文档中缺少书签: x'}))"
        with self.assertRaises(OfficeBackendError) as ctx:
            self._run(failed)
        self.assertIn("缺少书签", str(ctx.exception))

    def test_crash_without_result_is_readable(self) -> None:
        with self.assertRaises(OfficeBackendError) as ctx:
            self._run("import sys; sys.stderr.write('boom\\n'); sys.exit(3)")
        self.assertIn("异常退出", str(ctx.exception))
        self.assertIn("boom", str(ctx.exception))

    def test_timeout_raises_timeout_error(self) -> None:
        started = time.monotonic()
        with self.assertRaises(OfficeTimeoutError) as ctx:
            self._run("import time; time.sleep(60)", timeout=2)
        self.assertLess(time.monotonic() - started, 30)
        self.assertIn("超时", str(ctx.exception))

    def test_timeout_reports_last_step_and_saves_diagnostics(self) -> None:
        """超时报错给出最后一次 COM 调用，并把 worker 的 stderr 另存为诊断日志。"""
        code = (
            "import sys, time; sys.stdin.readline(); "
            "print('[win_com_worker] step: Documents.Open @ 10:00:00', file=sys.stderr, flush=True); "
            "print('[win_com_worker] step: Document.Repaginate @ 10:00:01', file=sys.stderr, flush=True); "
            "time.sleep(60)"
        )
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR": tmp}):
            with self.assertRaises(OfficeTimeoutError) as ctx:
                run_worker({"op": "inspect"}, 2, Path(tmp) / "office.pid", command=[sys.executable, "-c", code])
            message = str(ctx.exception)
            self.assertIn("最后一次 COM 调用: Document.Repaginate", message)
            logs = list((Path(tmp) / "diagnostics").glob("timeout-*-inspect-*.log"))
            self.assertEqual(len(logs), 1)
            # 父进程会解析长路径；CI 的临时目录可能是 8.3 短路径（RUNNER~1），只比较文件名。
            self.assertIn(logs[0].name, message)
            self.assertIn("Documents.Open", logs[0].read_text(encoding="utf-8"))

    def test_worker_receives_timeout_and_dumps_stack_before_it(self) -> None:
        """worker 收到超时秒数，并在超时前把调用栈写到 stderr。"""
        code = (
            "import faulthandler, json, sys, time; request = json.loads(sys.stdin.readline()); "
            "faulthandler.dump_traceback_later(max(request['timeout'] - 3.0, 0.5), file=sys.stderr); "
            "time.sleep(60)"
        )
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR": tmp}):
            with self.assertRaises(OfficeTimeoutError) as ctx:
                run_worker({"op": "probe"}, 4, Path(tmp) / "office.pid", command=[sys.executable, "-c", code])
            logs = list((Path(tmp) / "diagnostics").glob("timeout-*.log"))
            self.assertEqual(len(logs), 1, str(ctx.exception))
            self.assertIn("Timeout", logs[0].read_text(encoding="utf-8"))

    def test_worker_stuck_after_quit_still_returns_completed_result(self) -> None:
        """Word 关闭阶段卡住、worker 无法退出时，看门狗输出已完成的结果并结束 worker。"""
        code = (
            "import time; import lib.office.win_com_worker as w; "
            "w.QUIT_GRACE_SECONDS = 1; w._completion.watchdog_enabled = True; "
            "w._ready({'ok': True, 'page_map': 'completed-before-quit'}); "
            "w._arm_quit_watchdog(); time.sleep(60)"
        )
        started = time.monotonic()
        with tempfile.TemporaryDirectory() as tmp:
            env = {"PYTHONPATH": str(ROOT)}
            with patch.dict(os.environ, env):
                result = run_worker({"op": "inspect"}, 30, Path(tmp) / "office.pid",
                                    command=[sys.executable, "-c", code])
        self.assertEqual(result["page_map"], "completed-before-quit")
        self.assertLess(time.monotonic() - started, 20)

    def test_quit_watchdog_without_result_reports_failure(self) -> None:
        emitted, exits = [], []
        completion = win_com_worker._Completion()
        with patch.object(win_com_worker, "_completion", completion),                 patch.object(win_com_worker, "_hard_exit", exits.append),                 patch("builtins.print", lambda *a, **k: emitted.append(a[0]) if a else None):
            win_com_worker._on_quit_stalled()
        self.assertEqual(exits, [1])
        self.assertIn("退出时无响应", json.loads(emitted[-1])["error"])

    def test_quit_watchdog_is_inactive_outside_worker_main(self) -> None:
        completion = win_com_worker._Completion()
        with patch.object(win_com_worker, "_completion", completion):
            win_com_worker._arm_quit_watchdog(0.01)
            self.assertIsNone(completion.timer)

    def test_last_worker_step_parsing(self) -> None:
        self.assertIsNone(win_com.last_worker_step(""))
        stderr = "noise\n[win_com_worker] step: Documents.Open @ 10:00:00\n[win_com_worker] other\n"
        self.assertEqual(win_com.last_worker_step(stderr), "Documents.Open")

    @unittest.skipUnless(sys.platform == "win32", "按 PID 结束进程只在 Windows 上实现")
    def test_timeout_terminates_only_the_recorded_process(self) -> None:
        """替身“Office”进程被记录在 pid 文件中；另一个同名进程代表用户自己的实例。"""
        sleeper = [sys.executable, "-c", "import time; time.sleep(120)"]
        owned = subprocess.Popen(sleeper)
        user = subprocess.Popen(sleeper)
        try:
            identity = win_com.process_identity(owned.pid)
            self.assertIsNotNone(identity)
            with tempfile.TemporaryDirectory() as tmp:
                pid_file = Path(tmp) / "office.pid"
                pid_file.write_text(json.dumps({"pid": owned.pid, "image": identity[0], "created": identity[1]}), encoding="utf-8")
                with self.assertRaises(OfficeTimeoutError):
                    run_worker({"op": "probe"}, 2, pid_file, command=[sys.executable, "-c", "import time; time.sleep(60)"])
            owned.wait(timeout=15)
            self.assertIsNotNone(owned.returncode)
            self.assertIsNone(user.poll(), "未记录的同名进程不应被结束")
        finally:
            for proc in (owned, user):
                if proc.poll() is None:
                    proc.kill()
                    proc.wait()



# ---------------------------------------------------------------------------
# 会话：一次构建复用同一个 worker 与 Word 实例
# ---------------------------------------------------------------------------

def _paged_app(cls=None) -> "FakeWordApp":
    return (cls or FakeWordApp)(bookmark_starts={"cover": 1, "body": 9}, pages_by_start={1: (1, 1), 9: (3, 1)})


class SessionWorkerLoopTests(unittest.TestCase):
    """worker 会话主循环：用内存中的请求行与假 COM 对象驱动 ``serve_session``。"""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.apps: List[FakeWordApp] = []
        self.output: List[Dict[str, Any]] = []
        self.dumps: List[Any] = []
        self.consumed: List[str] = []
        self.patches = [
            patch.object(win_com_worker, "_completion", win_com_worker._Completion()),
            patch.object(win_com_worker, "_opened_paths", set()),
        ]
        for item in self.patches:
            item.start()

    def tearDown(self) -> None:
        for item in reversed(self.patches):
            item.stop()
        self.tmp.cleanup()

    def _dispatch(self, factory=None):
        def dispatch(progid: str) -> Any:
            self.assertEqual(progid, "Word.Application")
            app = (factory or _paged_app)()
            self.apps.append(app)
            return app
        return dispatch

    def _serve(self, requests, dispatch=None, pid_file=None, **kwargs) -> int:
        """``requests`` 可以是生成器：在两次请求之间修改假 Word 的状态。"""
        lines = (json.dumps(item) + "\n" for item in requests)

        def read_line() -> str:
            line = next(lines, "")
            if line:
                self.consumed.append(line)
            return line

        return win_com_worker.serve_session(
            {"op": "session", "pid_file": pid_file},
            read_line,
            lambda text: self.output.append(json.loads(text)),
            dispatch or self._dispatch(),
            _no_pid,
            arm_stack_dump=lambda seconds: self.dumps.append(("arm", seconds)),
            cancel_stack_dump=lambda: self.dumps.append("cancel"),
            **kwargs,
        )

    def _inspect(self, request_id: int, names=("cover", "body"), timeout: float = 600) -> Dict[str, Any]:
        return {"op": "inspect", "id": request_id, "timeout": timeout, "names": list(names),
                "docx": str(self.dir / f"p{request_id}.docx"), "pdf": str(self.dir / f"p{request_id}.pdf")}

    def test_requests_reuse_one_word_instance(self) -> None:
        code = self._serve([
            self._inspect(1),
            {"op": "export_pdf", "id": 2, "timeout": 600, "docx": str(self.dir / "e.docx"), "pdf": str(self.dir / "e.pdf")},
            {"op": "convert_doc", "id": 3, "timeout": 600, "src": str(self.dir / "a.doc"), "dst": str(self.dir / "a.docx")},
            self._inspect(4),
            {"op": "close"},
        ])
        self.assertEqual(code, 0)
        self.assertEqual(len(self.apps), 1, "整个会话只启动一次 Word")
        app = self.apps[0]
        self.assertEqual([item["id"] for item in self.output], [1, 2, 3, 4])
        self.assertTrue(all(item["ok"] for item in self.output))
        self.assertEqual(self.output[0]["page_map"], "cover\t1\t1\nbody\t3\t1\n")
        self.assertEqual(self.output[0]["word"]["word_version"], "16.0", "首个请求带回 Word 版本，供构建元数据复用")
        self.assertNotIn("word", self.output[1])
        self.assertEqual(len(app.documents), 4)
        self.assertTrue(all(doc.closed_with == 0 for doc in app.documents), "每个请求结束时关闭自己的文档")
        self.assertEqual(app.open_documents, [])
        self.assertEqual(app.quit_count, 1, "会话结束时退出一次 Word")
        self.assertEqual(app.quit_with, 0)
        self.assertFalse(app.Visible)
        self.assertEqual(self.dumps, [("arm", 597.0), "cancel"] * 4, "每个请求在超时前 3 秒导出调用栈，结束后取消")

    def test_end_of_input_also_quits_word(self) -> None:
        self._serve([self._inspect(1)])
        self.assertEqual(self.apps[0].quit_count, 1)

    def test_session_without_requests_never_starts_word(self) -> None:
        self.assertEqual(self._serve([{"op": "close"}]), 0)
        self.assertEqual(self.apps, [])

    def test_validation_failure_keeps_the_instance(self) -> None:
        """书签缺失等校验失败只影响当前请求；下一个请求仍在同一个 Word 中执行。"""
        self._serve([self._inspect(1, names=("cover", "absent")), self._inspect(2), {"op": "close"}])
        self.assertFalse(self.output[0]["ok"])
        self.assertIn("缺少书签: absent", self.output[0]["error"])
        self.assertNotIn("restart", self.output[0])
        self.assertFalse((self.dir / "p1.pdf").exists())
        self.assertTrue(self.output[1]["ok"])
        self.assertEqual(len(self.apps), 1)
        self.assertTrue(all(doc.closed_with == 0 for doc in self.apps[0].documents))
        self.assertEqual(self.apps[0].quit_count, 1)

    def test_unsupported_operation_keeps_the_instance(self) -> None:
        self._serve([{"op": "convert_pptx", "id": 1}, self._inspect(2), {"op": "close"}])
        self.assertIn("会话不支持", self.output[0]["error"])
        self.assertTrue(self.output[1]["ok"])

    def test_com_error_retires_the_instance_and_ends_the_session(self) -> None:
        """COM 错误之后实例状态不可信：结果带 restart，退出 Word，不再读取后续请求。"""
        error = FakeComError(-2147023174, "The RPC server is unavailable.", None, None)
        code = self._serve([
            {"op": "export_pdf", "id": 1, "timeout": 600, "docx": "x.docx", "pdf": str(self.dir / "o.pdf")},
            self._inspect(2),
        ], dispatch=self._dispatch(lambda: FakeWordApp(fail_export=error)))
        self.assertEqual(code, 1)
        self.assertEqual(len(self.output), 1)
        self.assertTrue(self.output[0]["restart"])
        self.assertIn("意外退出", self.output[0]["error"])
        self.assertEqual(len(self.consumed), 1, "失败后不再处理后续请求")
        self.assertEqual(self.apps[0].documents[0].closed_with, 0)
        self.assertEqual(self.apps[0].quit_count, 1)
        self.assertEqual(self.dumps, [("arm", 597.0), "cancel"])

    def test_launch_failure_retires_the_session(self) -> None:
        def dispatch(progid):
            raise FakeComError(-2146959355, "Server execution failed", None, None)  # CO_E_SERVER_EXEC_FAILURE

        self.assertEqual(self._serve([self._inspect(1), self._inspect(2)], dispatch=dispatch), 1)
        self.assertTrue(self.output[0]["restart"])
        self.assertIn("首次运行", self.output[0]["error"])

    def test_exited_word_between_requests_is_replaced(self) -> None:
        """两个请求之间 Word 已退出（被用户或其他程序结束）：启动新实例继续，不再调用旧实例。"""

        class ExitedApp(FakeWordApp):
            exited = False

            @property
            def Version(self):  # noqa: N802
                if self.exited:
                    raise FakeComError(-2147023174, "The RPC server is unavailable.", None, None)
                return "16.0"

        def requests():
            yield self._inspect(1)
            self.apps[0].exited = True
            yield self._inspect(2)
            yield {"op": "close"}

        self._serve(requests(), dispatch=self._dispatch(lambda: _paged_app(ExitedApp)),
                    instance_alive=lambda pid_file: False)
        self.assertTrue(all(item["ok"] for item in self.output))
        self.assertEqual(len(self.apps), 2)
        self.assertIn("word", self.output[1], "新实例的版本信息随结果返回")
        self.assertEqual(self.apps[0].quit_count, 0)
        self.assertEqual(self.apps[1].quit_count, 1)

    def test_unresponsive_word_between_requests_retires_the_session(self) -> None:
        """Word 进程仍在却无法调用：不另起实例，结束会话，由父进程按 pid 文件回收。"""

        class HungApp(FakeWordApp):
            hung = False

            @property
            def Version(self):  # noqa: N802
                if self.hung:
                    raise FakeComError(-2147023170, "The remote procedure call failed.", None, None)
                return "16.0"

        def requests():
            yield self._inspect(1)
            self.apps[0].hung = True
            yield self._inspect(2)

        code = self._serve(requests(), dispatch=self._dispatch(lambda: _paged_app(HungApp)),
                           instance_alive=lambda pid_file: True)
        self.assertEqual(code, 1)
        self.assertTrue(self.output[1]["restart"])
        self.assertEqual(len(self.apps), 1)

    def test_user_document_in_the_instance_is_handed_over_not_quit(self) -> None:
        """用户在构建期间双击打开的文档会被系统路由进隐藏实例：退出时移交而不是 Quit。"""
        pid_file = self.dir / "office.pid"
        pid_file.write_text(json.dumps({"pid": 4242, "image": "winword.exe", "created": 7}), encoding="utf-8")

        def requests():
            yield self._inspect(1)
            self.apps[0].open_documents.append(FakeUserDocument(r"C:\Users\someone\报告.docx"))
            yield self._inspect(2)
            yield {"op": "close"}

        self._serve(requests(), pid_file=str(pid_file))
        app = self.apps[0]
        self.assertTrue(all(item["ok"] for item in self.output))
        self.assertEqual(app.quit_count, 0, "有用户文档时不能退出 Word")
        self.assertTrue(app.Visible)
        self.assertEqual(app.DisplayAlerts, -1)
        self.assertEqual(len(app.open_documents), 1)
        self.assertIsNone(win_com.read_pid_record(pid_file), "pid 记录被撤销，父进程不会结束这个 Word")
        killed = []
        self.assertIsNone(cleanup_recorded_office(pid_file, 0, terminate=killed.append))
        self.assertEqual(killed, [])

    def test_leftover_own_document_is_closed_before_quit(self) -> None:
        """本次打开却未能关闭的文档不算用户文档：退出前再关闭一次，然后正常 Quit。"""
        app = FakeWordApp()
        leftover = FakeDocument(app, str(self.dir / "pagination.docx"))
        app.open_documents.append(leftover)
        win_com_worker._opened_paths.add(win_com_worker._normalized(leftover.FullName))
        win_com_worker.shutdown_word(app, None)
        self.assertEqual(leftover.closed_with, 0)
        self.assertEqual(app.quit_count, 1)

    def test_one_shot_mode_also_hands_over_user_documents(self) -> None:
        class Routed(FakeWordApp):
            def __init__(self) -> None:
                super().__init__()
                self.open_documents.append(FakeUserDocument(r"C:\Users\someone\a.docx"))

        result = handle({"op": "probe"}, dispatch=self._dispatch(Routed), word_pid_of=_no_pid)
        self.assertTrue(result["ok"])
        self.assertEqual(self.apps[0].quit_count, 0)
        self.assertTrue(self.apps[0].Visible)


FAKE_SESSION_WORKER = [sys.executable, str(ROOT / "tests" / "fake_word_session_worker.py")]


class WordSessionProcessTests(unittest.TestCase):
    """父进程 WordSession + 替身会话 worker 子进程（真实 serve_session、假 COM）。"""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.work = self.dir / "word-access"
        self.env = patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR": str(self.work)})
        self.env.start()
        self.sessions: List[win_com.WordSession] = []

    def tearDown(self) -> None:
        for session in self.sessions:
            session.close()
        self.env.stop()
        self.tmp.cleanup()

    def _session(self, **env: str) -> win_com.WordSession:
        session = win_com.WordSession(command=FAKE_SESSION_WORKER)
        self.sessions.append(session)
        self.fault_env = env
        return session

    def _inspect(self, session, index: int, names=("cover", "body"), timeout: float = 30) -> Dict[str, Any]:
        # 替身 worker 读取的环境变量在 Popen 时确定，所以每次请求都带上（只影响新启动的 worker）。
        with patch.dict(os.environ, self.fault_env):
            return session.request({"op": "inspect", "names": list(names), "docx": str(self.dir / f"p{index}.docx"),
                                    "pdf": str(self.dir / f"p{index}.pdf")}, timeout)

    def _jobs(self) -> List[Path]:
        return list(self.work.glob("job-*")) if self.work.exists() else []

    def test_requests_share_one_worker_and_close_cleans_up(self) -> None:
        session = self._session()
        self.assertEqual(self._jobs(), [], "第一个请求之前不启动 worker")
        first = self._inspect(session, 1)
        self.assertEqual(first["page_map"], "cover\t1\t1\nbody\t3\t1\n")
        self.assertEqual(first["word"]["word_version"], "16.0")
        for index in (2, 3):
            self.assertTrue(self._inspect(session, index)["ok"])
        self.assertTrue((self.dir / "p3.pdf").is_file())
        self.assertEqual(session.workers_started, 1)
        self.assertEqual(session._stderr_text().count("step: DispatchEx(Word.Application)"), 1)
        self.assertEqual([path.name.startswith("job-session-") for path in self._jobs()], [True])
        proc = session._proc
        session.close()
        self.assertEqual(proc.returncode, 0)
        self.assertEqual(self._jobs(), [], "会话结束后删除工作目录")

    def test_failure_between_requests_keeps_the_session(self) -> None:
        session = self._session()
        self.assertTrue(self._inspect(session, 1)["ok"])
        with self.assertRaises(OfficeBackendError) as ctx:
            self._inspect(session, 2, names=("cover", "absent"))
        self.assertIn("缺少书签: absent", str(ctx.exception))
        self.assertTrue(self._inspect(session, 3)["ok"])
        self.assertEqual(session.workers_started, 1)

    def test_com_error_restarts_worker_for_next_request(self) -> None:
        session = self._session(FAKE_WORD_FAULT="2:com")
        self.assertTrue(self._inspect(session, 1)["ok"])
        first_proc = session._proc
        with self.assertRaises(OfficeBackendError) as ctx:
            self._inspect(session, 2)
        self.assertNotIsInstance(ctx.exception, OfficeTimeoutError)
        self.assertIn("意外退出", str(ctx.exception))
        self.assertIsNotNone(first_proc.poll(), "失败的 worker 已退出")
        self.assertFalse((self.dir / "p2.pdf").exists())
        self.assertEqual(self._jobs(), [])
        result = self._inspect(session, 3)
        self.assertTrue(result["ok"])
        self.assertIn("word", result, "新 worker 启动了新的 Word")
        self.assertEqual(session.workers_started, 2)

    def test_worker_crash_is_readable_and_next_request_restarts(self) -> None:
        session = self._session(FAKE_WORD_FAULT="2:crash")
        self.assertTrue(self._inspect(session, 1)["ok"])
        with self.assertRaises(OfficeBackendError) as ctx:
            self._inspect(session, 2)
        self.assertIn("异常退出（退出码 3）", str(ctx.exception))
        self.assertEqual(self._jobs(), [])
        self.assertTrue(self._inspect(session, 3)["ok"])
        self.assertEqual(session.workers_started, 2)

    def test_timeout_in_second_request_reports_its_own_step(self) -> None:
        """第二个请求超时：报错只给出该请求的最后一步，保存诊断日志，下一个请求启动新 worker。"""
        session = self._session(FAKE_WORD_FAULT="2:hang")
        self.assertTrue(self._inspect(session, 1)["ok"])
        started = time.monotonic()
        with self.assertRaises(OfficeTimeoutError) as ctx:
            self._inspect(session, 2, timeout=5)
        self.assertLess(time.monotonic() - started, 20)
        message = str(ctx.exception)
        self.assertIn("超时（5 秒）", message)
        self.assertIn("最后一次 COM 调用: Document.Repaginate", message)
        logs = list((self.work / "diagnostics").glob("timeout-*-inspect-*.log"))
        self.assertEqual(len(logs), 1)
        self.assertIn(logs[0].name, message)
        log = logs[0].read_text(encoding="utf-8")
        self.assertIn("request: inspect #2", log)
        self.assertIn("Timeout", log, "超时前导出了调用栈")
        self.assertFalse((self.dir / "p2.pdf").exists())
        self.assertEqual(self._jobs(), [])
        self.assertTrue(self._inspect(session, 3)["ok"])
        self.assertEqual(session.workers_started, 2)

    def test_stalled_quit_at_close_is_bounded(self) -> None:
        """会话结束时 Word 关闭阶段卡住：worker 的 Quit 看门狗结束 worker，close 不会等满超时。"""
        session = self._session(FAKE_WORD_QUIT_HANG="1", FAKE_WORD_QUIT_GRACE="1")
        self.assertTrue(self._inspect(session, 1)["ok"])
        proc = session._proc
        started = time.monotonic()
        session.close()
        self.assertLess(time.monotonic() - started, 20)
        self.assertEqual(proc.returncode, 0)
        self.assertIn("Quit 之后", session._stderr_text())
        self.assertEqual(self._jobs(), [])

    def test_close_without_requests_is_a_no_op(self) -> None:
        session = self._session()
        session.close()
        self.assertEqual(session.workers_started, 0)

    def test_last_step_ignores_previous_request(self) -> None:
        stderr = (
            "[win_com_worker] request: inspect #1 @ 10:00:00\n"
            "[win_com_worker] step: Document.ExportAsFixedFormat @ 10:00:01\n"
            "[win_com_worker] request: inspect #2 @ 10:00:02\n"
        )
        self.assertIsNone(win_com.last_worker_step(stderr))
        self.assertEqual(win_com.last_worker_step(stderr + "[win_com_worker] step: Documents.Open @ 10:00:03\n"),
                         "Documents.Open")

    @unittest.skipUnless(sys.platform == "win32", "按 PID 结束进程只在 Windows 上实现")
    def test_session_timeout_terminates_only_the_recorded_process(self) -> None:
        """会话中超时只结束 pid 文件记录的实例；另一个同名进程代表用户自己的 Word。"""
        sleeper = [sys.executable, "-c", "import time; time.sleep(120)"]
        owned = subprocess.Popen(sleeper)
        user = subprocess.Popen(sleeper)
        try:
            session = self._session(FAKE_WORD_FAULT="2:hang", FAKE_WORD_RECORD_PID=str(owned.pid))
            self.assertTrue(self._inspect(session, 1)["ok"])
            time.sleep(1)
            self.assertIsNone(owned.poll(), "请求之间不结束会话中的实例")
            with self.assertRaises(OfficeTimeoutError):
                self._inspect(session, 2, timeout=3)
            owned.wait(timeout=15)
            self.assertIsNotNone(owned.returncode)
            self.assertIsNone(user.poll(), "未记录的同名进程不应被结束")
        finally:
            for proc in (owned, user):
                if proc.poll() is None:
                    proc.kill()
                    proc.wait()


class FakeSession:
    """替代 WordSession：记录请求，产物与 FakeRunner 相同。"""

    instances: List["FakeSession"] = []

    def __init__(self) -> None:
        self.requests: List[Dict[str, Any]] = []
        self.closed = 0
        self.runner = FakeRunner(page_map="_Synth_body\t2\t1\n")
        FakeSession.instances.append(self)

    def request(self, request: Dict[str, Any], timeout: float) -> Dict[str, Any]:
        self.requests.append(dict(request, timeout=timeout))
        result = dict(self.runner(request, timeout, Path("unused")))
        details = {"word_version": "16.0", "word_build": "16.0.20430", "word_bitness": 64}
        if request["op"] == "probe":
            result.update(details)
        if len(self.requests) == 1:
            result["word"] = details
        return result

    def close(self) -> None:
        self.closed += 1


class BackendSessionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.work = self.dir / "word-access"
        self.patches = [
            patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR": str(self.work)}),
            patch.object(win_com, "on_windows", lambda: True),
            patch.object(win_com, "progid_registered", lambda progid: True),
            patch.object(win_com, "pywin32_available", lambda: True),
        ]
        for item in self.patches:
            item.start()
        FakeSession.instances = []
        self.docx = self.dir / "final.docx"
        self.docx.write_bytes(b"PK-docx")
        self.runner = FakeRunner(page_map="_Synth_body\t2\t1\n")
        self.backend = WinComBackend(runner=self.runner, session_factory=FakeSession)

    def tearDown(self) -> None:
        for item in reversed(self.patches):
            item.stop()
        set_backend(None)
        self.tmp.cleanup()

    def test_word_operations_in_session_share_it(self) -> None:
        pptx = self.dir / "slides.pptx"
        pptx.write_bytes(b"pptx")
        with self.backend.session():
            for _ in range(3):
                report = self.backend.inspect_bookmarks(self.docx, self.dir / "final.pdf", ["_Synth_body"])
            self.assertEqual(report.bookmarks, {"_Synth_body": {"physical_page": 2, "printed_page": 1}})
            self.backend.export_pdf(self.docx, self.dir / "export.pdf")
            self.backend.convert_pptx_to_pdf(pptx, self.dir / "slides.pdf")
            self.assertEqual(self.backend.probe().details["word_version"], "16.0")
        session = FakeSession.instances[0]
        self.assertEqual([item["op"] for item in session.requests], ["inspect"] * 3 + ["export_pdf", "probe"])
        self.assertTrue(all(item["timeout"] == 600 for item in session.requests[:4]), "每个请求仍有自己的超时")
        self.assertEqual(session.requests[4]["timeout"], 120)
        self.assertNotEqual(Path(session.requests[0]["docx"]), self.docx.resolve(), "Word 只能打开私有副本")
        self.assertEqual([item["op"] for item in self.runner.requests], ["convert_pptx"], "PowerPoint 仍走单次模式")
        self.assertEqual(session.closed, 1)
        self.assertEqual(list(self.work.glob("job-*")), [])

    def test_failed_request_in_session_keeps_existing_pdf(self) -> None:
        pdf = self.dir / "final.pdf"
        pdf.write_bytes(b"previous")
        with self.backend.session():
            self.backend._session.runner.error = OfficeTimeoutError("等待超时")
            with self.assertRaises(OfficeTimeoutError):
                self.backend.inspect_bookmarks(self.docx, pdf, ["_Synth_body"])
        self.assertEqual(pdf.read_bytes(), b"previous")
        self.assertEqual(list(self.work.glob("job-*")), [])

    def test_outside_session_each_operation_is_one_shot(self) -> None:
        self.backend.inspect_bookmarks(self.docx, self.dir / "final.pdf", ["_Synth_body"])
        self.assertEqual(FakeSession.instances, [])
        self.assertEqual([item["op"] for item in self.runner.requests], ["inspect"])

    def test_session_is_closed_when_the_block_fails(self) -> None:
        with self.assertRaises(RuntimeError):
            with self.backend.session():
                self.backend.export_pdf(self.docx, self.dir / "export.pdf")
                raise RuntimeError("构建失败")
        self.assertEqual(FakeSession.instances[0].closed, 1)
        self.backend.export_pdf(self.docx, self.dir / "after.pdf")
        self.assertEqual([item["op"] for item in self.runner.requests], ["export_pdf"], "会话结束后回到单次模式")

    def test_nested_session_reuses_outer(self) -> None:
        with self.backend.session():
            with self.backend.session():
                self.backend.export_pdf(self.docx, self.dir / "a.pdf")
            self.backend.export_pdf(self.docx, self.dir / "b.pdf")
        self.assertEqual(len(FakeSession.instances), 1)
        self.assertEqual(FakeSession.instances[0].closed, 1)
        self.assertEqual(len(FakeSession.instances[0].requests), 2)

    def test_session_can_be_disabled(self) -> None:
        with patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_WORD_SESSION": "0"}), self.backend.session():
            self.backend.export_pdf(self.docx, self.dir / "a.pdf")
        self.assertEqual(FakeSession.instances, [])
        self.assertEqual([item["op"] for item in self.runner.requests], ["export_pdf"])

    def test_injected_runner_without_factory_never_starts_a_real_session(self) -> None:
        backend = WinComBackend(runner=self.runner)
        with backend.session():
            backend.export_pdf(self.docx, self.dir / "a.pdf")
        self.assertEqual([item["op"] for item in self.runner.requests], ["export_pdf"])

    def test_session_status_reuses_recorded_word_without_launching(self) -> None:
        self.assertIsNone(self.backend.session_status(), "未调用过 Word 时没有可复用的结果")
        with self.backend.session():
            self.backend.inspect_bookmarks(self.docx, self.dir / "final.pdf", ["_Synth_body"])
        status = self.backend.session_status()
        self.assertTrue(status.available)
        self.assertEqual(status.fidelity, "exact")
        self.assertIn("Microsoft Word 16.0 (build 16.0.20430), 64 位", status.reason)
        self.assertTrue(status.details["powerpoint_registered"])
        self.assertEqual(len(FakeSession.instances[0].requests), 1, "取状态不发出任何请求")

    def test_office_session_shares_backend_and_restores(self) -> None:
        from lib.office import office_session

        with patch.dict(os.environ, {"DOCUMENT_SYNTHESIS_OFFICE_BACKEND": "win-word"}), \
                patch.object(win_com, "WordSession", FakeSession):
            self.assertIsNot(get_backend(), get_backend(), "会话外每次得到新的后端实例")
            with office_session() as backend:
                self.assertIsInstance(backend, WinComBackend)
                self.assertIs(get_backend(), backend)
                get_backend().export_pdf(self.docx, self.dir / "a.pdf")
                get_backend().export_pdf(self.docx, self.dir / "b.pdf")
            self.assertIsNot(get_backend(), backend)
        self.assertEqual(len(FakeSession.instances), 1)
        self.assertEqual(len(FakeSession.instances[0].requests), 2)
        self.assertEqual(FakeSession.instances[0].closed, 1)

    def test_office_session_passes_through_backends_without_sessions(self) -> None:
        from lib.office import NullBackend, office_session
        from lib.office.mac_applescript import MacAppleScriptBackend

        self.assertFalse(hasattr(MacAppleScriptBackend(), "session"), "macOS 后端不参与会话")
        null = NullBackend()
        set_backend(null)
        with office_session() as backend:
            self.assertIs(backend, null)
        self.assertIs(get_backend(), null)


class BuildMetadataWordStatusTests(unittest.TestCase):
    """构建元数据记录 Word 可用性时不再启动 Word。"""

    class NoProbeBackend:
        name = "fake"
        fidelity = "exact"

        def __init__(self, recorded=None) -> None:
            self.recorded = recorded

        def static_status(self):
            from lib.office import BackendStatus
            return BackendStatus(True, "Microsoft Word 精确分页可用。", "exact")

        def probe(self):
            raise AssertionError("构建元数据不应再启动 Word 探针")

        def session_status(self):
            return self.recorded

    def tearDown(self) -> None:
        set_backend(None)

    def test_static_status_when_no_word_was_used(self) -> None:
        from lib.qa import word_build_status

        set_backend(self.NoProbeBackend())
        self.assertEqual(word_build_status(), (True, "Microsoft Word 精确分页可用。"))

    def test_recorded_session_status_is_reused(self) -> None:
        from lib.office import BackendStatus
        from lib.qa import word_build_status

        set_backend(self.NoProbeBackend(BackendStatus(True, "Automation 会话成功: Microsoft Word 16.0。", "exact")))
        self.assertEqual(word_build_status(), (True, "Automation 会话成功: Microsoft Word 16.0。"))

    def test_generate_build_metadata_keeps_fields_without_probe(self) -> None:
        from types import SimpleNamespace
        from lib.engine import generate_build_metadata

        set_backend(self.NoProbeBackend())
        config = SimpleNamespace(documents=[], project_name="demo", schema_version=3, override_paths=[])
        environment = generate_build_metadata(config, {}, {})["environment"]
        self.assertIs(environment["word_available"], True)
        self.assertEqual(environment["word_automation"], {"available": True, "reason": "Microsoft Word 精确分页可用。"})

if __name__ == "__main__":
    unittest.main()
