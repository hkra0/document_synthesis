# -*- coding: utf-8 -*-
"""Windows Office COM worker 子进程入口：``python -m lib.office.win_com_worker``。

两种模式，日志都写到 stderr：

- 单次模式：从 stdin 读取一行 JSON 请求，启动 Office、执行、退出 Office，在 stdout
  只输出一行结果 JSON。请求字段 ``op`` 取值：``probe``、``convert_doc``、
  ``convert_pptx``、``export_pdf``、``inspect``。
- 会话模式：第一行为 ``{"op": "session", "pid_file": ...}``，之后每行一个带 ``id`` 的
  Word 请求，每个请求输出一行带同一 ``id`` 的结果；Word 在第一个请求时启动，
  在 ``{"op": "close"}`` 或 stdin 结束时退出。一次构建内的全部 Word 操作共用这个实例。

所有路径都是父进程准备好的工作目录中的绝对路径。worker 启动 Office 后立即把进程
PID 与创建时间写入 ``pid_file``，父进程据此在超时或异常时只结束这一个进程。
COM 对象工厂可以注入，离线测试用假对象驱动。
"""

from __future__ import annotations

import faulthandler
import json
import os
from pathlib import Path
import sys
import threading
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
WD_ALERTS_ALL = -1
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_DOCUMENT_DEFAULT = 16
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
WD_ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1   # 打印页码（受起始页码设置影响）
WD_ACTIVE_END_PAGE_NUMBER = 3            # 物理页（从文档开头计数）
MSO_AUTOMATION_SECURITY_BY_UI = 2
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
PP_SAVE_AS_PDF = 32

#: 使用 Word 的操作；``convert_pptx`` 启动的是 PowerPoint，始终走单次模式。
WORD_OPS = ("probe", "convert_doc", "export_pdf", "inspect")


#: 调用 Quit 之后 worker 仍未退出的最长等待（秒）。第三方 COM 加载项可能让 Word 在关闭阶段卡住，
#: 此时最后一次释放 COM 引用的跨进程调用会一直阻塞；结果此前已经完成，不应再等满操作超时。
QUIT_GRACE_SECONDS = 30.0


class _Completion:
    """记录已完成的操作结果，并保证结果行只输出一次。"""

    def __init__(self) -> None:
        self.result: Optional[Dict[str, Any]] = None
        self.emitted = False
        self.lock = threading.Lock()
        self.timer: Optional[threading.Timer] = None
        # 只有作为 worker 子进程运行（main）时才启用看门狗；单元测试直接调用 handle 时不启用。
        self.watchdog_enabled = False


_completion = _Completion()


def _ready(result: Dict[str, Any]) -> Dict[str, Any]:
    """操作的全部产物已经生成：记下结果，供关闭阶段卡住时直接输出。"""
    _completion.result = result
    return result


def _emit(result: Dict[str, Any]) -> bool:
    with _completion.lock:
        if _completion.emitted:
            return False
        _completion.emitted = True
    print(json.dumps(result, ensure_ascii=False), flush=True)
    return True


def _hard_exit(code: int) -> None:
    os._exit(code)


def _on_quit_stalled() -> None:
    result = _completion.result or {
        "ok": False,
        "kind": "error",
        "error": "Microsoft Word 在退出时无响应，本次操作未完成；可能是第三方 COM 加载项阻塞了关闭。",
    }
    _log(f"Quit 之后 {QUIT_GRACE_SECONDS:.0f} 秒仍未完成释放，直接输出结果并退出 worker；"
         "父进程会结束本次启动的 Office 进程。")
    _emit(result)
    _hard_exit(0 if result.get("ok") else 1)


def _arm_quit_watchdog(seconds: Optional[float] = None) -> None:
    if not _completion.watchdog_enabled or _completion.timer is not None:
        return
    timer = threading.Timer(QUIT_GRACE_SECONDS if seconds is None else seconds, _on_quit_stalled)
    timer.daemon = True
    timer.start()
    _completion.timer = timer


class WorkerFailure(Exception):
    """带分类的 worker 失败：kind 为 ``error`` 或 ``unsupported``。"""

    def __init__(self, message: str, kind: str = "error") -> None:
        super().__init__(message)
        self.kind = kind


def _log(message: str) -> None:
    print(f"[win_com_worker] {message}", file=sys.stderr, flush=True)


#: stderr 中步骤行的前缀；父进程据此在超时报错里给出最后一次 COM 调用。
STEP_PREFIX = "[win_com_worker] step: "
#: 会话模式中每个请求开始时的分隔行；父进程只在最后一个请求的步骤中查找挂起点。
REQUEST_PREFIX = "[win_com_worker] request: "


def _step(name: str) -> None:
    """在每次可能长时间阻塞的 COM 调用之前记录步骤，用于定位挂起点。"""
    print(f"{STEP_PREFIX}{name} @ {time.strftime('%H:%M:%S')}", file=sys.stderr, flush=True)


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


def _release_pid_record(pid_file: Optional[str]) -> None:
    """Word 已移交给用户（其中有用户的文档）：撤销 pid 记录，父进程不再结束它。"""
    if not pid_file:
        return
    try:
        record = json.loads(Path(pid_file).read_text(encoding="utf-8"))
        released = {"released_pid": record.get("pid") if isinstance(record, dict) else None}
        Path(pid_file).write_text(json.dumps(released), encoding="utf-8")
    except (OSError, ValueError) as exc:
        _log(f"无法撤销 pid 记录: {exc!r}")


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


#: 本 worker 交给 Word 打开过的文件（规范化路径）。Word 退出前据此区分本次的文档与
#: 用户经资源管理器打开、被系统路由进这个隐藏实例的文档。
_opened_paths: set = set()


def _normalized(path: str) -> str:
    return os.path.normcase(os.path.abspath(str(path)))


def launch_word(
    pid_file: Optional[str],
    dispatch: Callable[[str], Any] = _default_dispatch,
    pid_of: Callable[[Any], int] = _word_pid,
) -> Any:
    """启动独立、隐藏、禁用宏与提示的 Word 实例并记录其 PID；配置失败时退出该实例。"""
    _write_pending_record(pid_file, WORD_IMAGE)
    _step("DispatchEx(Word.Application)")
    app = dispatch(WORD_PROGID)
    try:
        try:
            _record_instance(pid_file, pid_of(app), WORD_IMAGE)
        except Exception as exc:  # 记录失败不影响操作，父进程仍会结束 worker
            _log(f"无法记录 Word 进程: {exc!r}")
        def configure() -> None:
            app.Visible = False
            app.DisplayAlerts = WD_ALERTS_NONE

        call_with_busy_retry("Application 选项（Visible、DisplayAlerts）", configure)
        try:
            app.ScreenUpdating = False
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception as exc:
            _log(f"无法设置 Word 选项: {exc!r}")
    except BaseException:
        shutdown_word(app, pid_file)
        raise
    return app


def _foreign_document_count(app: Any) -> int:
    """关闭本 worker 遗留的文档，返回仍打开的其他（用户的）文档数。"""
    documents = app.Documents
    foreign = 0
    for index in range(int(documents.Count), 0, -1):
        doc = documents(index)
        if _normalized(doc.FullName) in _opened_paths:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:
                _log(f"关闭遗留文档失败: {exc!r}")
                foreign += 1
        else:
            foreign += 1
        del doc
    return foreign


def _hand_over_to_user(app: Any, pid_file: Optional[str]) -> None:
    """把装有用户文档的实例交还给用户：恢复可见与提示，撤销 pid 记录，不退出。"""
    _release_pid_record(pid_file)
    for name, value in (
        ("ScreenUpdating", True),
        ("DisplayAlerts", WD_ALERTS_ALL),
        ("AutomationSecurity", MSO_AUTOMATION_SECURITY_BY_UI),
        ("Visible", True),
    ):
        try:
            setattr(app, name, value)
        except Exception as exc:
            _log(f"无法恢复 Word 选项 {name}: {exc!r}")


def shutdown_word(app: Any, pid_file: Optional[str]) -> None:
    """退出本次启动的 Word；若其中有用户的文档，则移交给用户而不退出。"""
    _step("Application.Quit")
    _arm_quit_watchdog()
    try:
        foreign = _foreign_document_count(app)
    except Exception as exc:
        _log(f"无法检查 Word 中打开的文档: {exc!r}")
        foreign = 0
    if foreign:
        _log(f"Word 中有 {foreign} 份不是本次打开的文档（用户在构建期间打开），保留并显示 Word，不退出。")
        _hand_over_to_user(app, pid_file)
        return
    try:
        app.Quit(WD_DO_NOT_SAVE_CHANGES)
    except Exception as exc:
        _log(f"Word Quit 失败: {exc!r}")


@contextmanager
def word_application(
    pid_file: Optional[str],
    dispatch: Callable[[str], Any] = _default_dispatch,
    pid_of: Callable[[Any], int] = _word_pid,
) -> Iterator[Any]:
    """单次操作用的 Word 实例，退出时一定 ``Quit``（或移交给用户）。"""
    app = launch_word(pid_file, dispatch, pid_of)
    try:
        yield app
    finally:
        shutdown_word(app, pid_file)
        del app


@contextmanager
def word_document(app: Any, path: str) -> Iterator[Any]:
    _opened_paths.add(_normalized(path))
    doc = call_with_busy_retry("Documents.Open", lambda: app.Documents.Open(
        FileName=path,
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        Revert=False,
        Visible=False,
        NoEncodingDialog=True,
    ))
    try:
        yield doc
    finally:
        _step("Document.Close")
        try:
            doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception as exc:
            _log(f"关闭文档失败: {exc!r}")
        del doc


def export_document_pdf(doc: Any, pdf: str) -> None:
    """与 macOS 导出对齐：全文、按标题生成 PDF 书签、保留结构标签、缺字形转位图。"""
    call_with_busy_retry("Document.ExportAsFixedFormat", lambda: doc.ExportAsFixedFormat(
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
    ))


def measure_bookmarks(doc: Any, names: Sequence[str]) -> str:
    """读取页码是只读操作，Office 正忙时整体重试。"""
    return call_with_busy_retry(
        f"读取 {len(names)} 个书签的页码（Range.Information）", lambda: _measure_bookmarks(doc, names)
    )


def _measure_bookmarks(doc: Any, names: Sequence[str]) -> str:
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
    """单次模式：为一个请求启动 Office，执行后退出。"""
    op = request.get("op")
    if op == "convert_pptx":
        return convert_pptx(request, dispatch)
    if op not in WORD_OPS:
        raise WorkerFailure(f"未知的 worker 操作: {op!r}")
    with word_application(request.get("pid_file"), dispatch, word_pid_of) as app:
        return run_word_op(app, request)


def run_word_op(app: Any, request: Dict[str, Any]) -> Dict[str, Any]:
    """在已启动的 Word 实例上执行一个操作；文档在返回前一定关闭。"""
    op = request.get("op")

    if op == "probe":
        return _ready(dict(ok=True, **_word_details(app)))

    if op == "convert_doc":
        _opened_paths.add(_normalized(request["dst"]))  # SaveAs2 之后文档以新路径打开
        with word_document(app, request["src"]) as doc:
            call_with_busy_retry("Document.SaveAs2", lambda: doc.SaveAs2(
                FileName=request["dst"], FileFormat=WD_FORMAT_DOCUMENT_DEFAULT, AddToRecentFiles=False
            ))
            _ready({"ok": True})
        return {"ok": True}

    if op == "export_pdf":
        with word_document(app, request["docx"]) as doc:
            export_document_pdf(doc, request["pdf"])
            _ready({"ok": True})
        return {"ok": True}

    if op == "inspect":
        names = [str(name) for name in request.get("names", [])]
        with word_document(app, request["docx"]) as doc:
            call_with_busy_retry("Document.Repaginate", doc.Repaginate)
            page_map = measure_bookmarks(doc, names)
            export_document_pdf(doc, request["pdf"])
            _ready({"ok": True, "page_map": page_map})
        return {"ok": True, "page_map": page_map}

    raise WorkerFailure(f"未知的 worker 操作: {op!r}")


# ---------------------------------------------------------------------------
# 会话模式
# ---------------------------------------------------------------------------

def _recorded_instance_alive(pid_file: Optional[str]) -> bool:
    """pid 文件记录的 Word 是否仍在运行（PID、映像名、创建时间都一致）。"""
    if not pid_file:
        return False
    try:
        record = json.loads(Path(pid_file).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False
    if not isinstance(record, dict) or not isinstance(record.get("pid"), int):
        return False
    return process_identity(record["pid"]) == (record.get("image"), record.get("created"))


def _session_app_usable(app: Any, pid_file: Optional[str],
                        alive: Callable[[Optional[str]], bool]) -> bool:
    """复用前确认上一个请求留下的 Word 仍可调用。

    返回 False 表示该实例已经退出（例如被用户或其他程序结束），可以直接启动新实例；
    实例仍在运行却无法调用时抛出异常，由调用方结束会话，父进程按 pid 文件回收它。
    """
    try:
        call_with_busy_retry("Application.Version（会话复用前检查）", lambda: app.Version)
        return True
    except Exception:
        if alive(pid_file):
            raise
        _log("会话中的 Word 已经退出，启动新的实例。")
        return False


def serve_session(
    start: Dict[str, Any],
    read_line: Callable[[], str],
    write_line: Callable[[str], None],
    dispatch: Callable[[str], Any] = _default_dispatch,
    word_pid_of: Callable[[Any], int] = _word_pid,
    *,
    arm_stack_dump: Callable[[float], None] = lambda seconds: None,
    cancel_stack_dump: Callable[[], None] = lambda: None,
    instance_alive: Callable[[Optional[str]], bool] = _recorded_instance_alive,
) -> int:
    """会话模式主循环：同一个 Word 实例依次处理多个请求。

    - 每个请求输出一行带同一 ``id`` 的结果；``timeout`` 秒前导出调用栈，超时由父进程处理。
    - 请求内的校验失败（:class:`WorkerFailure`，如书签缺失）不影响实例，会话继续。
    - 其他异常（COM 错误、Word 崩溃、对话框阻塞等）之后实例状态不可信：结果中带
      ``"restart": true``，随后退出 Word 并结束 worker；父进程在下一个请求时启动新 worker。
    """
    pid_file = start.get("pid_file")
    app: Any = None
    code = 0
    try:
        while True:
            line = read_line()
            if not line:
                break
            try:
                request = json.loads(line)
            except ValueError as exc:
                write_line(json.dumps({"ok": False, "kind": "error", "error": f"无效的 worker 请求: {exc}"},
                                      ensure_ascii=False))
                continue
            op = request.get("op")
            if op == "close":
                break
            request_id = request.get("id")
            app_label = request.get("app_label", "Microsoft Word")
            print(f"{REQUEST_PREFIX}{op} #{request_id} @ {time.strftime('%H:%M:%S')}", file=sys.stderr, flush=True)
            timeout = request.get("timeout")
            if isinstance(timeout, (int, float)) and timeout > 0:
                arm_stack_dump(max(float(timeout) - 3.0, 0.5))
            restart = False
            try:
                if op not in WORD_OPS:
                    raise WorkerFailure(f"会话不支持的 worker 操作: {op!r}")
                details: Dict[str, Any] = {}
                if app is not None and not _session_app_usable(app, pid_file, instance_alive):
                    app = None
                if app is None:
                    app = launch_word(pid_file, dispatch, word_pid_of)
                    details = _word_details(app)
                result = dict(run_word_op(app, request))
                if details:
                    result["word"] = details
            except WorkerFailure as exc:
                result = {"ok": False, "kind": exc.kind, "error": str(exc)}
            except Exception as exc:
                _log(f"{type(exc).__name__}: {exc!r}")
                result = {"ok": False, "kind": "error", "error": translate_exception(exc, app_label), "restart": True}
                restart = True
            finally:
                cancel_stack_dump()
            result["id"] = request_id
            write_line(json.dumps(result, ensure_ascii=False))
            if restart:
                code = 1
                break
    finally:
        # 所有结果都已输出：Quit 卡住时看门狗只需结束 worker。
        _completion.result = {"ok": code == 0}
        _completion.emitted = True
        if app is not None:
            shutdown_word(app, pid_file)
            del app
    return code


def convert_pptx(request: Dict[str, Any], dispatch: Callable[[str], Any] = _default_dispatch,
                 pid_of: Callable[[Any], int] = _powerpoint_pid,
                 already_running: Optional[Callable[[str], bool]] = None) -> Dict[str, Any]:
    """PowerPoint 是单实例应用：若用户已打开 PowerPoint，只关闭本次演示文稿，不退出应用。"""
    running = already_running or _running_images
    user_instance = running(POWERPOINT_IMAGE)
    if not user_instance:
        _write_pending_record(request.get("pid_file"), POWERPOINT_IMAGE)
    _step("DispatchEx(PowerPoint.Application)")
    app = dispatch(POWERPOINT_PROGID)
    try:
        if not user_instance:
            try:
                _record_instance(request.get("pid_file"), pid_of(app), POWERPOINT_IMAGE)
            except Exception as exc:
                _log(f"无法记录 PowerPoint 进程: {exc!r}")
        # 参数依次为 ReadOnly、Untitled、WithWindow
        presentation = call_with_busy_retry(
            "Presentations.Open", lambda: app.Presentations.Open(request["src"], True, False, False)
        )
        try:
            call_with_busy_retry("Presentation.SaveAs(PDF)", lambda: presentation.SaveAs(request["dst"], PP_SAVE_AS_PDF))
            _ready({"ok": True})
        finally:
            try:
                presentation.Close()
            except Exception as exc:
                _log(f"关闭演示文稿失败: {exc!r}")
            del presentation
    finally:
        _arm_quit_watchdog()
        if not user_instance:
            try:
                app.Quit()
            except Exception as exc:
                _log(f"PowerPoint Quit 失败: {exc!r}")
        del app
    return {"ok": True}


# ---------------------------------------------------------------------------
# COM 忙碌重试
# ---------------------------------------------------------------------------
# pywin32 没有提供 CoRegisterMessageFilter，无法注册 IMessageFilter；
# 改为在调用层识别“Office 正忙”的 HRESULT 并退避重试。

_BUSY_HRESULTS = {0x80010001, 0x8001010A}  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER
BUSY_RETRY_SECONDS = 60.0


def _hresults_of(exc: BaseException) -> List[int]:
    args = getattr(exc, "args", ()) or ()
    codes = []
    if args and isinstance(args[0], int):
        codes.append(args[0] & 0xFFFFFFFF)
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 5 and isinstance(args[2][5], int):
        codes.append(args[2][5] & 0xFFFFFFFF)
    return codes


def call_with_busy_retry(
    step: str,
    action: Callable[[], Any],
    *,
    limit: float = BUSY_RETRY_SECONDS,
    sleep: Callable[[float], None] = time.sleep,
    clock: Callable[[], float] = time.monotonic,
) -> Any:
    """记录步骤并执行一次 COM 调用；Office 正忙时按退避重试，最长 ``limit`` 秒。"""
    _step(step)
    deadline = clock() + limit
    delay = 0.25
    while True:
        try:
            return action()
        except Exception as exc:
            busy = [code for code in _hresults_of(exc) if code in _BUSY_HRESULTS]
            if not busy or clock() >= deadline:
                raise
            _log(f"{step}: Office 正忙（0x{busy[0]:08X}），{delay:.2f} 秒后重试")
            sleep(delay)
            delay = min(delay * 2, 2.0)


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
    session = request.get("op") == "session"
    try:
        import pythoncom
    except ImportError:
        print(json.dumps({
            "ok": False,
            "kind": "unsupported",
            "error": "缺少 Python 依赖 pywin32；请运行 pip install -r requirements.txt。",
            **({"restart": True} if session else {}),
        }, ensure_ascii=False))
        return 1

    def arm_stack_dump(seconds: float) -> None:
        # 超时前导出全部线程的 Python 调用栈；父进程超时后会保存 stderr 作为诊断日志。
        faulthandler.dump_traceback_later(seconds, exit=False, file=sys.stderr)

    _completion.watchdog_enabled = True
    if session:
        def write_line(text: str) -> None:
            print(text, flush=True)

        pythoncom.CoInitialize()
        try:
            return serve_session(
                request, sys.stdin.readline, write_line,
                arm_stack_dump=arm_stack_dump, cancel_stack_dump=faulthandler.cancel_dump_traceback_later,
            )
        finally:
            pythoncom.CoUninitialize()
            if _completion.timer is not None:
                _completion.timer.cancel()

    timeout = request.get("timeout")
    if isinstance(timeout, (int, float)) and timeout > 0:
        arm_stack_dump(max(float(timeout) - 3.0, 0.5))

    pythoncom.CoInitialize()
    try:
        try:
            result = handle(request)
        except WorkerFailure as exc:
            result = {"ok": False, "kind": exc.kind, "error": str(exc)}
        except Exception as exc:
            _log(f"{type(exc).__name__}: {exc!r}")
            result = {"ok": False, "kind": "error", "error": translate_exception(exc, app_label)}
        # 先输出结果再反初始化 COM：CoUninitialize 也可能因卡住的 Office 而阻塞，
        # 此时 Quit 看门狗会直接结束 worker，父进程已能读到结果。
        _emit(result)
    finally:
        pythoncom.CoUninitialize()
        if _completion.timer is not None:
            _completion.timer.cancel()
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
