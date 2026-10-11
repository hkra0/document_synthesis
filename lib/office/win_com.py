# -*- coding: utf-8 -*-
"""Windows Microsoft Word / PowerPoint COM 自动化后端。

COM 调用是同步的，进程内无法中断，所以所有 Office 操作都放在独立 worker 子进程
（``python -m lib.office.win_com_worker``）中执行：父进程通过 stdin 发送一行 JSON
请求，worker 在 stdout 只返回一行结果 JSON，日志写到 stderr。

可靠性约定：

- worker 用 ``DispatchEx`` 启动独立的 Word 实例，并把该实例的 PID 与创建时间写入
  本次作业的 pid 文件。父进程超时或 worker 异常退出时，只按这个 PID 结束 Word，
  并核对映像名与创建时间，绝不按进程名结束用户自己打开的 Word。
- Word 只读打开工作目录中的私有副本；目标文件只在全部检查通过后才被替换，
  失败时不会留下半成品。
- 只使用数值常量，不依赖 ``gencache`` 生成的 ``gen_py`` 缓存。

本模块在任何平台都可以导入；``winreg``、``ctypes.windll`` 与 pywin32 只在
Windows 上调用时才加载。
"""

from __future__ import annotations

import ctypes
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from .base import (
    BackendStatus,
    BookmarkReport,
    OfficeBackendError,
    OfficeTimeoutError,
    OfficeUnsupportedError,
)


WORD_PROGID = "Word.Application"
POWERPOINT_PROGID = "PowerPoint.Application"
WORD_IMAGE = "winword.exe"
POWERPOINT_IMAGE = "powerpnt.exe"

#: 单次 Office 操作的默认超时（秒），与 macOS AppleScript 的 ``with timeout of 600 seconds`` 一致。
DEFAULT_TIMEOUT = 600
#: ``--doctor`` 探针只启动 Word 读取版本，给更短的超时。
PROBE_TIMEOUT = 120
#: worker 退出后等待其启动的 Office 进程自行退出的宽限时间（秒）。
EXIT_GRACE_SECONDS = 15.0

_PROJECT_ROOT = Path(__file__).resolve().parents[2]

# 常见 HRESULT（有符号 32 位整数在 pywin32 中以负数出现，这里统一按无符号比较）。
_HR_CLASS_STRING = 0x800401F3        # CO_E_CLASSSTRING：ProgID 未注册
_HR_CLASS_NOT_REG = 0x80040154       # REGDB_E_CLASSNOTREG
_HR_SERVER_EXEC = 0x80080005         # CO_E_SERVER_EXEC_FAILURE：服务器启动失败
_HR_CALL_REJECTED = 0x80010001       # RPC_E_CALL_REJECTED
_HR_RETRY_LATER = 0x8001010A         # RPC_E_SERVERCALL_RETRYLATER：对象被占用
_HR_DISCONNECTED = 0x80010108        # RPC_E_DISCONNECTED
_HR_RPC_UNAVAILABLE = 0x800706BA     # RPC_S_SERVER_UNAVAILABLE
_HR_RPC_FAILED = 0x800706BE          # RPC_S_CALL_FAILED
_HR_ACCESS_DENIED = 0x80070005       # E_ACCESSDENIED


def _unsigned(code: Optional[int]) -> Optional[int]:
    if code is None:
        return None
    return int(code) & 0xFFFFFFFF


def com_failure_message(
    hresult: Optional[int],
    detail: str = "",
    app_label: str = "Microsoft Word",
) -> str:
    """把 COM HRESULT 与错误描述翻译成可执行的中文排障提示。

    ``hresult`` 可以是外层 HRESULT，也可以是 ``excepinfo`` 中的 scode；
    两者都会被调用方分别尝试。``detail`` 是 Office 返回的原始描述，原样附在末尾。
    """
    code = _unsigned(hresult)
    normalized = (detail or "").lower()
    suffix = f"（{detail.strip()}）" if detail and detail.strip() else ""

    if code in (_HR_CLASS_STRING, _HR_CLASS_NOT_REG):
        return f"未检测到已注册的 {app_label}，请确认已安装桌面版 Microsoft Office。{suffix}"
    if code == _HR_SERVER_EXEC:
        return (
            f"无法启动 {app_label}。请先手动打开一次 {app_label}，完成首次运行向导、"
            "登录与激活；若当前终端以管理员身份运行而 Office 不是，也会出现此错误。" + suffix
        )
    if code in (_HR_RETRY_LATER, _HR_CALL_REJECTED):
        return (
            f"{app_label} 正忙或被对话框阻塞（对象被占用 0x{code:08X}）。"
            f"请关闭 {app_label} 中打开的对话框后重试。" + suffix
        )
    if code in (_HR_DISCONNECTED, _HR_RPC_UNAVAILABLE, _HR_RPC_FAILED):
        return (
            f"{app_label} 进程在自动化过程中意外退出（可能被强制结束或崩溃），"
            "本次操作未完成，已停止。" + suffix
        )
    if code == _HR_ACCESS_DENIED:
        return f"{app_label} 拒绝访问文件或自动化接口，请检查文件权限与杀毒软件拦截。{suffix}"
    if "protected view" in normalized or "受保护的视图" in detail or "受保护视图" in detail:
        return (
            f"{app_label} 以受保护视图打开了文件，自动化无法继续。请在文件属性中“解除锁定”，"
            "或在信任中心调整受保护视图设置后重试。" + suffix
        )
    if (
        "not activated" in normalized
        or "unlicensed" in normalized
        or "reduced functionality" in normalized
        or "未激活" in detail
        or "功能受限" in detail
    ):
        return f"{app_label} 未激活或处于功能受限模式，请先完成 Office 激活。{suffix}"
    return f"{app_label} 自动化失败: {detail.strip() or (f'HRESULT 0x{code:08X}' if code is not None else '未知错误')}"


# ---------------------------------------------------------------------------
# Windows 进程与注册表工具（仅在 Windows 上调用）
# ---------------------------------------------------------------------------

_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_SYNCHRONIZE = 0x00100000
_STILL_ACTIVE = 259
_WAIT_OBJECT_0 = 0


def on_windows() -> bool:
    return sys.platform == "win32"


def progid_registered(progid: str) -> bool:
    """只查询 ``HKCR\\<ProgID>\\CLSID``，不启动应用。"""
    if sys.platform != "win32":
        return False
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid + r"\CLSID") as key:
            value, _ = winreg.QueryValueEx(key, "")
            return bool(value)
    except OSError:
        return False


def pywin32_available() -> bool:
    return importlib.util.find_spec("win32com") is not None and importlib.util.find_spec("pythoncom") is not None


def _kernel32():
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    k32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)
    ]
    k32.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    k32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    k32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    k32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    k32.WaitForSingleObject.restype = wintypes.DWORD
    return k32


def process_identity(pid: int) -> Optional[Tuple[str, int]]:
    """返回仍在运行的进程的 (小写映像名, 创建时间 FILETIME)；进程不存在时返回 None。"""
    if sys.platform != "win32" or not pid:
        return None
    from ctypes import wintypes

    k32 = _kernel32()
    handle = k32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return None
    try:
        exit_code = wintypes.DWORD()
        if not k32.GetExitCodeProcess(handle, ctypes.byref(exit_code)) or exit_code.value != _STILL_ACTIVE:
            return None
        size = wintypes.DWORD(1024)
        buffer = ctypes.create_unicode_buffer(size.value)
        if not k32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return None
        created, exited, kernel, user = (wintypes.FILETIME() for _ in range(4))
        if not k32.GetProcessTimes(
            handle, ctypes.byref(created), ctypes.byref(exited), ctypes.byref(kernel), ctypes.byref(user)
        ):
            return None
        creation = (created.dwHighDateTime << 32) | created.dwLowDateTime
        return os.path.basename(buffer.value).lower(), creation
    finally:
        k32.CloseHandle(handle)


def terminate_process(pid: int, wait_seconds: float = 10.0) -> bool:
    """结束指定 PID 并等待其退出；成功或进程已不存在时返回 True。"""
    if sys.platform != "win32":
        return False
    k32 = _kernel32()
    handle = k32.OpenProcess(_PROCESS_TERMINATE | _SYNCHRONIZE | _PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return process_identity(pid) is None
    try:
        k32.TerminateProcess(handle, 1)
        return k32.WaitForSingleObject(handle, int(wait_seconds * 1000)) == _WAIT_OBJECT_0
    finally:
        k32.CloseHandle(handle)


def read_pid_record(pid_file: Path) -> Optional[Dict[str, Any]]:
    try:
        record = json.loads(Path(pid_file).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(record, dict) or not isinstance(record.get("pid"), int):
        return None
    return record


def current_filetime() -> int:
    """当前 UTC 时间的 FILETIME，与进程创建时间可直接比较。"""
    from ctypes import wintypes

    value = wintypes.FILETIME()
    ctypes.WinDLL("kernel32").GetSystemTimeAsFileTime(ctypes.byref(value))
    return (value.dwHighDateTime << 32) | value.dwLowDateTime


def list_process_ids(image: str) -> List[int]:
    """当前能查询到的、映像名为 ``image`` 的进程 PID。"""
    import win32process

    return [pid for pid in win32process.EnumProcesses() if (process_identity(pid) or ("",))[0] == image]


def process_command_line(pid: int) -> Optional[str]:
    """读取进程命令行（``NtQueryInformationProcess`` 的 ProcessCommandLineInformation）。"""
    if sys.platform != "win32" or not pid:
        return None

    class _UnicodeString(ctypes.Structure):
        _fields_ = [("Length", ctypes.c_ushort), ("MaximumLength", ctypes.c_ushort), ("Buffer", ctypes.c_void_p)]

    k32 = _kernel32()
    ntdll = ctypes.WinDLL("ntdll")
    ntdll.NtQueryInformationProcess.restype = ctypes.c_long
    ntdll.NtQueryInformationProcess.argtypes = [
        ctypes.c_void_p, ctypes.c_ulong, ctypes.c_void_p, ctypes.c_ulong, ctypes.POINTER(ctypes.c_ulong)
    ]
    handle = k32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return None
    try:
        size = ctypes.c_ulong(0)
        ntdll.NtQueryInformationProcess(handle, 60, None, 0, ctypes.byref(size))
        if not size.value:
            return None
        buffer = ctypes.create_string_buffer(size.value)
        if ntdll.NtQueryInformationProcess(handle, 60, buffer, size.value, ctypes.byref(size)) != 0:
            return None
        text = _UnicodeString.from_buffer(buffer)
        return ctypes.wstring_at(text.Buffer, text.Length // 2) if text.Buffer else ""
    finally:
        k32.CloseHandle(handle)


def is_automation_instance(pid: int) -> bool:
    """COM 启动的 Office 命令行带 ``/Automation -Embedding``；用户手动打开的不会带。"""
    command_line = (process_command_line(pid) or "").lower()
    return "-embedding" in command_line or "/automation" in command_line


def read_pending_record(pid_file: Path) -> Optional[Dict[str, Any]]:
    try:
        record = json.loads(Path(pid_file).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(record, dict) or not record.get("pending"):
        return None
    return record


def find_orphan_office(
    record: Dict[str, Any],
    *,
    list_pids: Optional[Callable[[str], List[int]]] = None,
    identity: Callable[[int], Optional[Tuple[str, int]]] = process_identity,
    automation: Optional[Callable[[int], bool]] = None,
) -> Optional[int]:
    """启动前快照之后新出现、由 COM 启动、且唯一的 Office 进程才视为本次启动的实例。

    不以“是否有可见窗口”判断：插件弹窗会让自动化实例也出现可见窗口。
    多个候选（例如另一个构建同时在启动 Word）时不结束任何进程，避免误杀。
    """
    list_pids = list_pids or list_process_ids
    automation = automation or is_automation_instance
    image = str(record.get("image", "")).lower()
    before = set(record.get("before") or [])
    started = record.get("started")
    if not image or not isinstance(started, int):
        return None
    candidates = []
    for pid in list_pids(image):
        if pid in before:
            continue
        current = identity(pid)
        if current is None or current[0] != image or current[1] < started:
            continue
        if not automation(pid):
            continue
        candidates.append(pid)
    return candidates[0] if len(candidates) == 1 else None


def cleanup_recorded_office(
    pid_file: Path,
    grace_seconds: float = EXIT_GRACE_SECONDS,
    *,
    identity: Callable[[int], Optional[Tuple[str, int]]] = process_identity,
    terminate: Callable[[int], bool] = terminate_process,
    sleep: Callable[[float], None] = time.sleep,
) -> Optional[int]:
    """结束 worker 自己启动且仍在运行的 Office 进程，返回被结束的 PID。

    只有映像名与创建时间都和 pid 文件一致时才结束，避免 PID 复用后误杀用户进程。
    ``grace_seconds`` 内进程自行退出（正常 ``Quit`` 之后）则不做任何事。

    如果 Office 卡在启动阶段（例如首次运行或许可证对话框），worker 来不及记录 PID，
    pid 文件里只有启动前的快照，此时改用 :func:`find_orphan_office` 查找。
    """
    record = read_pending_record(pid_file)
    if record is not None:
        orphan = find_orphan_office(record, identity=identity)
        if orphan is not None:
            terminate(orphan)
        return orphan
    record = read_pid_record(pid_file)
    if record is None:
        return None
    pid = record["pid"]
    expected = (str(record.get("image", "")).lower(), record.get("created"))

    deadline = time.monotonic() + max(grace_seconds, 0.0)
    while True:
        current = identity(pid)
        if current is None or current != expected:
            return None
        if time.monotonic() >= deadline:
            break
        sleep(0.25)
    terminate(pid)
    return pid


# ---------------------------------------------------------------------------
# 父进程：调度 worker
# ---------------------------------------------------------------------------

def _timeout_seconds(default: int) -> float:
    configured = os.environ.get("DOCUMENT_SYNTHESIS_OFFICE_TIMEOUT")
    if configured:
        try:
            value = float(configured)
            if value > 0:
                return value
        except ValueError:
            pass
    return float(default)


def _worker_env() -> Dict[str, str]:
    env = dict(os.environ)
    existing = env.get("PYTHONPATH")
    env["PYTHONPATH"] = str(_PROJECT_ROOT) + (os.pathsep + existing if existing else "")
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    return env


WORKER_COMMAND = (sys.executable, "-m", "lib.office.win_com_worker")


def run_worker(
    request: Dict[str, Any],
    timeout: float,
    pid_file: Path,
    command: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    """在独立子进程中执行一次 Office 操作，并保证其启动的 Office 进程被回收。"""
    payload = dict(request, pid_file=str(pid_file))
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        proc = subprocess.Popen(
            list(command or WORKER_COMMAND),
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=_worker_env(),
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=creationflags,
        )
    except OSError as exc:
        raise OfficeBackendError(f"无法启动 Office 自动化子进程: {exc}") from exc

    timed_out = False
    try:
        try:
            stdout, stderr = proc.communicate(json.dumps(payload, ensure_ascii=False), timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            proc.kill()
            stdout, stderr = proc.communicate()
    finally:
        # 超时时 worker 已被结束，Word 不会再收到 Quit，直接清理；
        # 正常退出时给 Word 一段宽限时间自行退出。
        cleanup_recorded_office(pid_file, 0.0 if timed_out else EXIT_GRACE_SECONDS)

    label = request.get("app_label", "Microsoft Word")
    if timed_out:
        raise OfficeTimeoutError(
            f"等待 {label} 响应超时（{int(timeout)} 秒），已结束本次启动的 Office 进程。"
            f"请检查文档是否会触发 {label} 的对话框（密码、修复、宏或许可证提示）。"
        )

    result = None
    for line in reversed((stdout or "").splitlines()):
        if line.strip():
            try:
                result = json.loads(line)
            except ValueError:
                result = None
            break
    if not isinstance(result, dict):
        tail = (stderr or "").strip().splitlines()[-5:]
        raise OfficeBackendError(
            f"{label} 自动化子进程异常退出（退出码 {proc.returncode}）"
            + (": " + " | ".join(tail) if tail else "。")
        )
    if not result.get("ok"):
        message = str(result.get("error") or f"{label} 自动化失败。")
        if result.get("kind") == "unsupported":
            raise OfficeUnsupportedError(message)
        raise OfficeBackendError(message)
    return result


def _work_root() -> Path:
    configured = os.environ.get("DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR")
    if configured:
        directory = Path(configured).expanduser()
    else:
        directory = Path(tempfile.gettempdir()) / "document-synthesis-word-access"
    directory.mkdir(parents=True, exist_ok=True)
    return directory.resolve()


def _strip_zone_identifier(path: Path) -> None:
    """移除“来自网络”标记，避免副本以受保护视图打开。"""
    try:
        os.remove(str(path) + ":Zone.Identifier")
    except OSError:
        pass


def _replace_with_retry(source: Path, target: Path) -> None:
    """把工作目录中的产物原子地替换到目标路径。

    工作目录（默认在系统临时目录）可能与目标不在同一个盘符，``os.replace`` 不能跨盘，
    所以先复制到目标目录中的临时文件，再在同一目录内替换；目标被占用时按退避重试。
    """
    from lib.pagination import _replace_pdf_with_retry

    staged = target.with_name(f".{target.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        shutil.copyfile(source, staged)
        _replace_pdf_with_retry(staged, target)
    except OSError as exc:
        raise OfficeBackendError(f"无法写入目标文件 {target}: {exc}") from exc
    except Exception as exc:
        raise OfficeBackendError(str(exc)) from exc
    finally:
        try:
            staged.unlink()
        except OSError:
            pass


class WinComBackend:
    """Windows 平台 Microsoft Word / PowerPoint COM 自动化后端。"""
    name: str = "win-word"
    fidelity: str = "exact"

    def __init__(
        self,
        runner: Optional[Callable[[Dict[str, Any], float, Path], Dict[str, Any]]] = None,
    ) -> None:
        self._runner = runner or run_worker

    # -- 工作目录 ---------------------------------------------------------

    def get_word_access_directory(self) -> Path:
        return _work_root()

    def _new_job_dir(self) -> Path:
        job = _work_root() / f"job-{uuid.uuid4().hex[:16]}"
        job.mkdir(parents=True)
        return job

    def _run(self, request: Dict[str, Any], job: Path, timeout: float) -> Dict[str, Any]:
        return self._runner(request, timeout, job / "office.pid")

    # -- 状态 -------------------------------------------------------------

    def static_status(self) -> BackendStatus:
        """只查注册表与依赖，不启动 Word。"""
        if not on_windows():
            return BackendStatus(False, "Windows COM 后端只能在 Windows 上使用。", "unsupported")
        if not progid_registered(WORD_PROGID):
            return BackendStatus(
                False,
                "未检测到已安装的 Microsoft Word（注册表中没有 Word.Application），无法取得精确页码；"
                "当前可使用 --plan、格式分析类命令和 smoke_test.py --structure-only。",
                "unsupported",
            )
        if not pywin32_available():
            return BackendStatus(
                False,
                "已检测到 Microsoft Word，但缺少 Python 依赖 pywin32；请运行 pip install -r requirements.txt。",
                "unsupported",
            )
        return BackendStatus(
            True,
            "Microsoft Word 精确分页可用。",
            "exact",
            {"powerpoint_registered": progid_registered(POWERPOINT_PROGID)},
        )

    def probe(self) -> BackendStatus:
        """启动独立 Word 实例读取版本后退出；不打开任何文档。"""
        stat = self.static_status()
        if not stat.available:
            return stat
        job = self._new_job_dir()
        try:
            result = self._run({"op": "probe"}, job, _timeout_seconds(PROBE_TIMEOUT))
        except OfficeBackendError as exc:
            return BackendStatus(False, str(exc), "unsupported", dict(stat.details))
        finally:
            shutil.rmtree(job, ignore_errors=True)
        details = dict(stat.details)
        details.update({key: result[key] for key in ("word_version", "word_build", "word_bitness", "click_to_run") if key in result})
        version = result.get("word_version") or "未知版本"
        build = result.get("word_build")
        bitness = result.get("word_bitness")
        summary = f"Microsoft Word {version}" + (f" (build {build})" if build else "") + (f", {bitness} 位" if bitness else "")
        return BackendStatus(True, f"Automation 探针成功: {summary}。", "exact", details)

    # -- 转换与导出 ------------------------------------------------------

    def _require_word(self) -> None:
        stat = self.static_status()
        if not stat.available:
            raise OfficeUnsupportedError(stat.reason)

    def convert_doc_to_docx(self, src: Path, dst: Path) -> None:
        self._require_word()
        src, dst = Path(src).resolve(), Path(dst).resolve()
        dst.parent.mkdir(parents=True, exist_ok=True)
        job = self._new_job_dir()
        try:
            working_src = job / ("source" + src.suffix.lower())
            working_dst = job / "converted.docx"
            shutil.copy2(src, working_src)
            _strip_zone_identifier(working_src)
            self._run(
                {"op": "convert_doc", "src": str(working_src), "dst": str(working_dst)},
                job,
                _timeout_seconds(DEFAULT_TIMEOUT),
            )
            if not working_dst.is_file():
                raise OfficeBackendError(f"转换 DOC 文件后未找到产物: {src.name}")
            _replace_with_retry(working_dst, dst)
        except OSError as exc:
            raise OfficeBackendError(f"转换 DOC 文件 {src.name} 失败: {exc}") from exc
        finally:
            shutil.rmtree(job, ignore_errors=True)

    def convert_pptx_to_pdf(self, src: Path, dst: Path) -> None:
        if on_windows() and not progid_registered(POWERPOINT_PROGID):
            raise OfficeUnsupportedError(
                f"未检测到已安装的 Microsoft PowerPoint，无法把 {Path(src).name} 转为 PDF；"
                "请手动导出同名 PDF 放在旁边，构建会直接使用它。"
            )
        self._require_word()
        src, dst = Path(src).resolve(), Path(dst).resolve()
        dst.parent.mkdir(parents=True, exist_ok=True)
        job = self._new_job_dir()
        try:
            working_src = job / ("source" + src.suffix.lower())
            working_dst = job / "converted.pdf"
            shutil.copy2(src, working_src)
            _strip_zone_identifier(working_src)
            self._run(
                {
                    "op": "convert_pptx",
                    "src": str(working_src),
                    "dst": str(working_dst),
                    "app_label": "Microsoft PowerPoint",
                },
                job,
                _timeout_seconds(DEFAULT_TIMEOUT),
            )
            if not working_dst.is_file():
                raise OfficeBackendError(f"转换 PPTX 文件后未找到产物: {src.name}")
            _replace_with_retry(working_dst, dst)
        except OSError as exc:
            raise OfficeBackendError(f"转换 PPTX 文件 {src.name} 失败: {exc}") from exc
        finally:
            shutil.rmtree(job, ignore_errors=True)

    def export_pdf(self, docx: Path, pdf: Path) -> None:
        self._require_word()
        docx, pdf = Path(docx).resolve(), Path(pdf).resolve()
        job = self._new_job_dir()
        try:
            working_docx = job / "export.docx"
            working_pdf = job / "export.pdf"
            shutil.copy2(docx, working_docx)
            _strip_zone_identifier(working_docx)
            self._run(
                {"op": "export_pdf", "docx": str(working_docx), "pdf": str(working_pdf)},
                job,
                _timeout_seconds(DEFAULT_TIMEOUT),
            )
            if not working_pdf.is_file():
                raise OfficeBackendError("Microsoft Word 未生成 PDF。")
            _replace_with_retry(working_pdf, pdf)
        except OSError as exc:
            raise OfficeBackendError(f"无法准备或保存 Word 导出的 PDF: {exc}") from exc
        finally:
            shutil.rmtree(job, ignore_errors=True)

    def inspect_bookmarks(
        self, docx: Path, pdf: Path, names: Sequence[str]
    ) -> BookmarkReport:
        """重新分页，读取各书签物理页与打印页码，并导出最终核验 PDF。"""
        from pypdf import PdfReader
        from lib.pagination import _parse_page_map

        self._require_word()
        names = list(names)
        docx, pdf = Path(docx).resolve(), Path(pdf).resolve()
        job = self._new_job_dir()
        try:
            working_docx = job / "pagination.docx"
            working_pdf = job / "pagination.pdf"
            shutil.copy2(docx, working_docx)
            _strip_zone_identifier(working_docx)
            result = self._run(
                {"op": "inspect", "docx": str(working_docx), "pdf": str(working_pdf), "names": names},
                job,
                _timeout_seconds(DEFAULT_TIMEOUT),
            )
            if not working_pdf.is_file():
                raise OfficeBackendError("Word 未生成本次分页核验的 PDF，已停止构建。")
            try:
                total_pages = len(PdfReader(str(working_pdf)).pages)
            except Exception as exc:
                raise OfficeBackendError(f"Word 导出 PDF 无法读取: {exc}") from exc
            if total_pages < 1:
                raise OfficeBackendError("Word 导出的 PDF 没有页面。")
            try:
                page_map = _parse_page_map(str(result.get("page_map", "")), names, total_pages)
            except Exception as exc:
                raise OfficeBackendError(str(exc)) from exc
            _replace_with_retry(working_pdf, pdf)
            return BookmarkReport(bookmarks=page_map, total_pages=total_pages, pdf_path=pdf, fidelity="exact")
        except OSError as exc:
            raise OfficeBackendError(f"无法准备 Word 分页核验的工作副本: {exc}") from exc
        finally:
            shutil.rmtree(job, ignore_errors=True)


__all__: List[str] = [
    "WinComBackend",
    "com_failure_message",
    "cleanup_recorded_office",
    "process_identity",
    "progid_registered",
    "run_worker",
    "terminate_process",
]
