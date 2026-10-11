# -*- coding: utf-8 -*-
"""Windows Office COM worker 子进程入口：``python -m lib.office.win_com_worker``。

从 stdin 读取一行 JSON 请求，在 stdout 只输出一行结果 JSON，日志写到 stderr。
请求字段 ``op`` 取值：``probe``、``convert_doc``、``convert_pptx``、``export_pdf``、
``inspect``。所有路径都是父进程准备好的工作目录中的绝对路径。

worker 启动 Office 后立即把进程 PID 与创建时间写入 ``pid_file``，父进程据此在
超时或异常时只结束这一个进程。COM 对象工厂可以注入，离线测试用假对象驱动。
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import sys
import time
import uuid
from contextlib import contextmanager
from typing import Any, Callable, Dict, Iterator, List, Optional, Sequence, Tuple

from .win_com import (
    POWERPOINT_IMAGE,
    POWERPOINT_PROGID,
    WORD_IMAGE,
    WORD_PROGID,
    com_failure_message,
    current_filetime,
    find_orphan_office,
    list_process_ids,
    process_identity,
    read_pending_record,
)


# Word / PowerPoint 数值常量（不依赖 gencache）。
WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1   # 打印页码（受起始页码设置影响）
WD_ACTIVE_END_PAGE_NUMBER = 3            # 物理页（从文档开头计数）
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
PP_SAVE_AS_PDF = 32


class WorkerFailure(Exception):
    """带分类的 worker 失败：kind 为 ``error`` 或 ``unsupported``。"""

    def __init__(self, message: str, kind: str = "error") -> None:
        super().__init__(message)
        self.kind = kind


def _log(message: str) -> None:
    print(f"[win_com_worker] {message}", file=sys.stderr, flush=True)


def translate_exception(exc: BaseException, app_label: str) -> str:
    """从 pywin32 ``com_error`` 中取出 HRESULT 与描述并翻译。"""
    if isinstance(exc, WorkerFailure):
        return str(exc)
    args = getattr(exc, "args", ()) or ()
    hresult = args[0] if args and isinstance(args[0], int) else None
    detail = args[1] if len(args) > 1 and isinstance(args[1], str) else ""
    scode = None
    if len(args) > 2 and isinstance(args[2], tuple):
        excepinfo = args[2]
        # excepinfo = (wCode, source, description, helpFile, helpContext, scode)
        if len(excepinfo) > 2 and isinstance(excepinfo[2], str) and excepinfo[2].strip():
            detail = excepinfo[2]
        if len(excepinfo) > 5 and isinstance(excepinfo[5], int):
            scode = excepinfo[5]
    if hresult is None and not detail:
        detail = str(exc)
    # Office 方法内部错误的外层 HRESULT 是 DISP_E_EXCEPTION，具体原因在 scode；
    # RPC 断连等错误则只有外层 HRESULT。先用 scode，未命中已知情形再用外层。
    generic_prefix = f"{app_label} 自动化失败"
    if scode is not None:
        message = com_failure_message(scode, detail, app_label)
        if not message.startswith(generic_prefix):
            return message
    return com_failure_message(hresult, detail, app_label)


# ---------------------------------------------------------------------------
# pid 记录
# ---------------------------------------------------------------------------

def _write_pending_record(pid_file: Optional[str], image: str) -> None:
    """启动 Office 之前记录同名进程快照；启动卡住时父进程据此找回孤儿进程。"""
    if not pid_file:
        return
    try:
        record = {"pending": True, "image": image, "before": list_process_ids(image), "started": current_filetime()}
        Path(pid_file).write_text(json.dumps(record), encoding="utf-8")
    except Exception as exc:
        _log(f"无法记录启动前进程快照: {exc!r}")


def _record_instance(pid_file: Optional[str], pid: int, image: str) -> None:
    """记录本次启动的 Office 实例；按窗口取不到 PID 时，用启动前快照的差集兜底。"""
    if not pid and pid_file:
        pending = read_pending_record(Path(pid_file))
        if pending is not None:
            pid = find_orphan_office(pending) or 0
    _write_pid_record(pid_file, pid, image)


def _write_pid_record(pid_file: Optional[str], pid: int, image: str) -> None:
    if not pid_file or not pid:
        return
    identity = process_identity(pid)
    if identity is None or identity[0] != image:
        _log(f"PID {pid} 不是预期的 {image}，不记录。")
        return
    record = {"pid": pid, "image": identity[0], "created": identity[1]}
    Path(pid_file).write_text(json.dumps(record), encoding="utf-8")


def _word_pid(app: Any) -> int:
    """通过唯一窗口标题找到本实例的主窗口，再取其进程 ID。"""
    import win32gui
    import win32process

    token = "document-synthesis-" + uuid.uuid4().hex
    app.Caption = token
    hwnd = 0
    for _ in range(20):
        hwnd = win32gui.FindWindow("OpusApp", token)
        if hwnd:
            break
        time.sleep(0.1)
    if not hwnd:
        return 0
    return int(win32process.GetWindowThreadProcessId(hwnd)[1])


def _powerpoint_pid(app: Any) -> int:
    import win32process

    try:
        hwnd = int(app.HWND)
    except Exception:
        return 0
    return int(win32process.GetWindowThreadProcessId(hwnd)[1]) if hwnd else 0


def _running_images(image: str) -> bool:
    """当前会话中是否已有同名进程在运行（用于判断 PowerPoint 是否为用户实例）。"""
    return bool(list_process_ids(image))


# ---------------------------------------------------------------------------
# COM 会话
# ---------------------------------------------------------------------------

def _default_dispatch(progid: str) -> Any:
    import win32com.client

    return win32com.client.DispatchEx(progid)


@contextmanager
def word_application(
    pid_file: Optional[str],
    dispatch: Callable[[str], Any] = _default_dispatch,
    pid_of: Callable[[Any], int] = _word_pid,
) -> Iterator[Any]:
    """启动独立、隐藏、禁用宏与提示的 Word 实例，退出时一定 ``Quit``。"""
    _write_pending_record(pid_file, WORD_IMAGE)
    app = dispatch(WORD_PROGID)
    try:
        try:
            _record_instance(pid_file, pid_of(app), WORD_IMAGE)
        except Exception as exc:  # 记录失败不影响操作，父进程仍会结束 worker
            _log(f"无法记录 Word 进程: {exc!r}")
        app.Visible = False
        app.DisplayAlerts = WD_ALERTS_NONE
        try:
            app.ScreenUpdating = False
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception as exc:
            _log(f"无法设置 Word 选项: {exc!r}")
        yield app
    finally:
        try:
            app.Quit(WD_DO_NOT_SAVE_CHANGES)
        except Exception as exc:
            _log(f"Word Quit 失败: {exc!r}")
        del app


@contextmanager
def word_document(app: Any, path: str) -> Iterator[Any]:
    doc = app.Documents.Open(
        FileName=path,
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        Revert=False,
        Visible=False,
        NoEncodingDialog=True,
    )
    try:
        yield doc
    finally:
        try:
            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception as exc:
            _log(f"关闭文档失败: {exc!r}")
        del doc


def export_document_pdf(doc: Any, pdf: str) -> None:
    """与 macOS 导出对齐：全文、按标题生成 PDF 书签、保留结构标签、缺字形转位图。"""
    doc.ExportAsFixedFormat(
        OutputFileName=pdf,
        ExportFormat=WD_EXPORT_FORMAT_PDF,
        OpenAfterExport=False,
        OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
        Range=WD_EXPORT_ALL_DOCUMENT,
        Item=WD_EXPORT_DOCUMENT_CONTENT,
        IncludeDocProps=True,
        KeepIRM=True,
        CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
        DocStructureTags=True,
        BitmapMissingFonts=True,
        UseISO19005_1=False,
    )


def measure_bookmarks(doc: Any, names: Sequence[str]) -> str:
    """返回 ``名\\t物理页\\t打印页\\n`` 格式的页码报告，供 ``_parse_page_map`` 严格解析。"""
    bookmarks = doc.Bookmarks
    try:
        bookmarks.ShowHidden = True
    except Exception:
        pass
    missing = [name for name in names if not bookmarks.Exists(name)]
    if missing:
        raise WorkerFailure("Word 文档中缺少书签: " + ", ".join(missing))
    lines: List[str] = []
    for name in names:
        start = bookmarks(name).Range.Start
        point = doc.Range(start, start)
        physical = point.Information(WD_ACTIVE_END_PAGE_NUMBER)
        printed = point.Information(WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER)
        lines.append(f"{name}\t{physical}\t{printed}\n")
    return "".join(lines)


def _word_details(app: Any) -> Dict[str, Any]:
    details: Dict[str, Any] = {"word_version": str(app.Version)}
    try:
        details["word_build"] = str(app.Build)
    except Exception:
        pass
    try:
        path = str(app.Path)
        lowered = path.lower()
        details["word_bitness"] = 32 if "program files (x86)" in lowered else 64
        details["click_to_run"] = "\\root\\office" in lowered
    except Exception:
        pass
    return details


# ---------------------------------------------------------------------------
# 操作
# ---------------------------------------------------------------------------

def handle(request: Dict[str, Any], dispatch: Callable[[str], Any] = _default_dispatch,
           word_pid_of: Callable[[Any], int] = _word_pid) -> Dict[str, Any]:
    op = request.get("op")
    pid_file = request.get("pid_file")

    if op == "probe":
        with word_application(pid_file, dispatch, word_pid_of) as app:
            return dict(ok=True, **_word_details(app))

    if op == "convert_doc":
        with word_application(pid_file, dispatch, word_pid_of) as app:
            with word_document(app, request["src"]) as doc:
                doc.SaveAs2(FileName=request["dst"], FileFormat=WD_FORMAT_DOCUMENT_DEFAULT, AddToRecentFiles=False)
        return {"ok": True}

    if op == "export_pdf":
        with word_application(pid_file, dispatch, word_pid_of) as app:
            with word_document(app, request["docx"]) as doc:
                export_document_pdf(doc, request["pdf"])
        return {"ok": True}

    if op == "inspect":
        names = [str(name) for name in request.get("names", [])]
        with word_application(pid_file, dispatch, word_pid_of) as app:
            with word_document(app, request["docx"]) as doc:
                doc.Repaginate()
                page_map = measure_bookmarks(doc, names)
                export_document_pdf(doc, request["pdf"])
        return {"ok": True, "page_map": page_map}

    if op == "convert_pptx":
        return convert_pptx(request, dispatch)

    raise WorkerFailure(f"未知的 worker 操作: {op!r}")


def convert_pptx(request: Dict[str, Any], dispatch: Callable[[str], Any] = _default_dispatch,
                 pid_of: Callable[[Any], int] = _powerpoint_pid,
                 already_running: Optional[Callable[[str], bool]] = None) -> Dict[str, Any]:
    """PowerPoint 是单实例应用：若用户已打开 PowerPoint，只关闭本次演示文稿，不退出应用。"""
    running = already_running or _running_images
    user_instance = running(POWERPOINT_IMAGE)
    if not user_instance:
        _write_pending_record(request.get("pid_file"), POWERPOINT_IMAGE)
    app = dispatch(POWERPOINT_PROGID)
    try:
        if not user_instance:
            try:
                _record_instance(request.get("pid_file"), pid_of(app), POWERPOINT_IMAGE)
            except Exception as exc:
                _log(f"无法记录 PowerPoint 进程: {exc!r}")
        presentation = app.Presentations.Open(request["src"], True, False, False)  # ReadOnly, Untitled, WithWindow
        try:
            presentation.SaveAs(request["dst"], PP_SAVE_AS_PDF)
        finally:
            try:
                presentation.Close()
            except Exception as exc:
                _log(f"关闭演示文稿失败: {exc!r}")
            del presentation
    finally:
        if not user_instance:
            try:
                app.Quit()
            except Exception as exc:
                _log(f"PowerPoint Quit 失败: {exc!r}")
        del app
    return {"ok": True}


# ---------------------------------------------------------------------------
# COM 忙碌重试（IMessageFilter）
# ---------------------------------------------------------------------------

class _BusyRetryFilter:
    """Office 忙碌（RPC_E_CALL_REJECTED / RETRYLATER）时自动重试，最长约 60 秒。"""

    _public_methods_ = ["HandleInComingCall", "RetryRejectedCall", "MessagePending"]

    def HandleInComingCall(self, dwCallType, htaskCaller, dwTickCount, lpInterfaceInfo):  # noqa: N802
        return 0  # SERVERCALL_ISHANDLED

    def RetryRejectedCall(self, htaskCallee, dwTickCount, dwRejectType):  # noqa: N802
        if dwRejectType == 2 and dwTickCount < 60000:  # SERVERCALL_RETRYLATER
            return 250
        return -1

    def MessagePending(self, htaskCallee, dwTickCount, dwPendingType):  # noqa: N802
        return 2  # PENDINGMSG_WAITDEFPROCESS


def _register_message_filter() -> None:
    try:
        import pythoncom
        from win32com.server.util import wrap

        filter_object = _BusyRetryFilter()
        filter_object._com_interfaces_ = [pythoncom.IID_IMessageFilter]
        pythoncom.CoRegisterMessageFilter(wrap(filter_object, pythoncom.IID_IMessageFilter))
    except Exception as exc:
        _log(f"无法注册 COM 消息过滤器: {exc!r}")


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    try:
        request = json.loads(sys.stdin.readline() or "{}")
    except ValueError as exc:
        print(json.dumps({"ok": False, "kind": "error", "error": f"无效的 worker 请求: {exc}"}, ensure_ascii=False))
        return 1

    app_label = request.get("app_label", "Microsoft Word")
    try:
        import pythoncom
    except ImportError:
        print(json.dumps({
            "ok": False,
            "kind": "unsupported",
            "error": "缺少 Python 依赖 pywin32；请运行 pip install -r requirements.txt。",
        }, ensure_ascii=False))
        return 1

    pythoncom.CoInitialize()
    try:
        _register_message_filter()
        try:
            result = handle(request)
        except WorkerFailure as exc:
            result = {"ok": False, "kind": exc.kind, "error": str(exc)}
        except Exception as exc:
            _log(f"{type(exc).__name__}: {exc!r}")
            result = {"ok": False, "kind": "error", "error": translate_exception(exc, app_label)}
    finally:
        try:
            pythoncom.CoRegisterMessageFilter(None)
        except Exception:
            pass
        pythoncom.CoUninitialize()

    print(json.dumps(result, ensure_ascii=False), flush=True)
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
