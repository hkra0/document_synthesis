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
        self.Bookmarks = FakeBookmarks(self)
        self.calls: List[Any] = []
        self.closed_with = None

    def Repaginate(self) -> None:  # noqa: N802
        self.calls.append("Repaginate")

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


class FakeDocuments:
    def __init__(self, app: "FakeWordApp") -> None:
        self.app = app

    def Open(self, **kwargs: Any) -> FakeDocument:  # noqa: N802
        self.app.open_kwargs = kwargs
        doc = FakeDocument(self.app, kwargs["FileName"])
        self.app.documents.append(doc)
        return doc


class FakeWordApp:
    Version = "16.0"
    Build = "16.0.99999"
    Path = r"C:\Program Files\Microsoft Office\root\Office16"

    def __init__(self, bookmark_starts=None, pages_by_start=None, fail_export=None) -> None:
        self.bookmark_starts = bookmark_starts or {}
        self.pages_by_start = pages_by_start or {}
        self.fail_export = fail_export
        self.Documents = FakeDocuments(self)
        self.documents: List[FakeDocument] = []
        self.quit_with = None
        self.Visible = True
        self.DisplayAlerts = -1
        self.open_kwargs: Dict[str, Any] = {}

    def Quit(self, SaveChanges: int = None) -> None:  # noqa: N802,N803
        self.quit_with = SaveChanges


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


if __name__ == "__main__":
    unittest.main()
