# -*- coding: utf-8 -*-
"""macOS AppleScript Word 与 PowerPoint 自动化后端。

原样封装现有的 AppleScript 脚本与 macOS 沙盒临时目录，行为与既有实现保持严格一致。
"""

import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import uuid
from typing import Any, Dict, List, Optional, Sequence

from pypdf import PdfReader

from .base import (
    BackendStatus,
    BookmarkReport,
    OfficeBackendError,
    OfficeTimeoutError,
    OfficeUnsupportedError,
)


def _applescript_string(value: str) -> str:
    """将 POSIX 路径安全嵌入 AppleScript 字符串。"""
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _automation_failure_message(detail: str) -> str:
    """把 macOS/AppleScript 的底层错误转成可执行的排障提示。"""
    normalized = detail.lower()
    if "-1743" in detail or "not authorized" in normalized or "not permitted" in normalized:
        return (
            "当前运行构建的应用尚未获准控制 Microsoft Word。"
            "请在“系统设置 → 隐私与安全性 → 自动化”中允许该应用控制 Word；"
            "Full Disk Access 不能替代此权限。"
        )
    if (
        "connection invalid" in normalized
        or "hiservices-xpcservice" in normalized
        or "can't get application id" in normalized
        or "can’t get application id" in normalized
    ):
        return (
            "当前命令运行在无法连接 macOS 图形自动化服务的受限上下文中。"
            "请从已获 Automation 权限的本机终端或 Codex 桌面应用运行构建。"
        )
    return f"AppleScript 调用 Microsoft Word 失败: {detail or '未知错误'}"


def _word_access_directory() -> Path:
    """Return one stable directory for all files touched by Word automation.

    macOS Word asks for folder access on the directory used by ``save as``.
    Per-call random temporary directories therefore cause an authorization
    dialog for every pagination pass.  A dedicated stable directory keeps the
    grant target constant; callers still use unique files and clean them up.
    ``DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR`` is intended for CI or a user-facing
    directory that has already been granted to Word.
    """
    configured = os.environ.get("DOCUMENT_SYNTHESIS_WORD_ACCESS_DIR")
    if configured:
        directory = Path(configured).expanduser()
    elif sys.platform == "darwin":
        directory = Path("/private/tmp/document-synthesis-word-access")
    else:
        directory = Path(tempfile.gettempdir()) / "document-synthesis-word-access"
    directory.mkdir(parents=True, exist_ok=True)
    return directory.resolve()


def _inspection_script(docx_path: Path, pdf_path: Path, bookmark_names: Sequence[str]) -> str:
    names = ", ".join('"' + _applescript_string(name) + '"' for name in bookmark_names)
    return f'''
with timeout of 600 seconds
    tell application "Microsoft Word"
        set previousAlerts to display alerts
        set inspectionDoc to missing value
        try
            set display alerts to none
            open (POSIX file "{_applescript_string(str(docx_path))}")
            set inspectionDoc to document "{_applescript_string(Path(docx_path).name)}"
            repaginate inspectionDoc
            set pageReport to ""
            repeat with bookmarkName in {{{names}}}
                set bookmarkName to bookmarkName as text
                set bookmarkStart to start of bookmark of bookmark bookmarkName of inspectionDoc
                set bookmarkRange to create range inspectionDoc start bookmarkStart end bookmarkStart
                set physicalPage to get range information bookmarkRange information type active end page number
                set printedPage to get range information bookmarkRange information type active end adjusted page number
                set pageReport to pageReport & bookmarkName & tab & (physicalPage as text) & tab & (printedPage as text) & linefeed
            end repeat
            -- Keep the document handle explicit.  Word's AppleScript
            -- dictionary accepts `save as` on a document reference, while
            -- the `active document` property expression can reject the same
            -- command with error -1708 during a real automation run.
            save as inspectionDoc file name "{_applescript_string(str(pdf_path))}" file format format PDF
            close inspectionDoc saving no
            set inspectionDoc to missing value
            set display alerts to previousAlerts
            return pageReport
        on error errorMessage number errorNumber
            if inspectionDoc is not missing value then
                try
                    close inspectionDoc saving no
                end try
            end if
            set display alerts to previousAlerts
            error errorMessage number errorNumber
        end try
    end tell
end timeout
'''


def _run_applescript(script: str, label: str) -> None:
    if sys.platform != "darwin" or not os.path.exists("/usr/bin/osascript"):
        raise OfficeUnsupportedError(f"{label} 需要 macOS 上的 Microsoft Office。")
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise OfficeBackendError(f"{label}失败: {detail or '未知错误'}")


class MacAppleScriptBackend:
    """macOS 平台原生 AppleScript Office 自动化后端。"""
    name: str = "mac-word"
    fidelity: str = "exact"

    def __init__(self) -> None:
        pass

    def get_word_access_directory(self) -> Path:
        return _word_access_directory()

    def _execute_subprocess(self, cmd: List[str], timeout: int) -> subprocess.CompletedProcess:
        # 兼容历史上直接 patch lib.pagination 或 lib.qa 中的 subprocess.run 的单元测试
        for module_name in ("lib.pagination", "lib.qa"):
            mod = sys.modules.get(module_name)
            if mod and hasattr(mod, "subprocess"):
                sub = getattr(mod, "subprocess")
                run_func = getattr(sub, "run", None)
                if (
                    run_func is not None
                    and (
                        getattr(run_func, "_mock_return_value", None) is not None
                        or getattr(run_func, "side_effect", None) is not None
                        or hasattr(run_func, "assert_called")
                    )
                ):
                    return run_func(
                        cmd,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        timeout=timeout,
                    )
        return subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )

    def static_status(self) -> BackendStatus:
        """检查精确分页的静态前置条件，不触发系统授权弹窗。"""
        if sys.platform != "darwin":
            return BackendStatus(False, "精确分页目前需要 macOS 上的 Microsoft Word。", "unsupported")
        if not os.path.exists("/usr/bin/osascript"):
            return BackendStatus(False, "未找到 osascript，无法调用 Microsoft Word。", "unsupported")
        if not os.path.exists("/Applications/Microsoft Word.app"):
            return BackendStatus(False, "未在 /Applications 找到 Microsoft Word。", "unsupported")
        return BackendStatus(True, "Microsoft Word 精确分页可用。", "exact")

    def probe(self) -> BackendStatus:
        """真实验证当前调用方能否控制 Word；不读取或修改任何文档。"""
        stat = self.static_status()
        if not stat.available:
            return stat
        try:
            result = self._execute_subprocess(
                ["osascript", "-e", 'tell application "Microsoft Word" to get name'],
                timeout=15,
            )
        except subprocess.TimeoutExpired:
            return BackendStatus(False, "等待 Microsoft Word Automation 响应超时。请确认 Word 没有被模态对话框阻塞。", "unsupported")
        if result.returncode == 0:
            return BackendStatus(True, f"Automation 探针成功: {(result.stdout or 'Microsoft Word').strip()}。", "exact")
        detail = (result.stderr or result.stdout).strip()
        return BackendStatus(False, _automation_failure_message(detail), "unsupported")

    def convert_doc_to_docx(self, src: Path, dst: Path) -> None:
        """调用 Word 将 DOC 文件另存为 DOCX。"""
        dst.parent.mkdir(parents=True, exist_ok=True)
        script = f'''
        with timeout of 600 seconds
            tell application "Microsoft Word"
                set display alerts to none
                open (POSIX file "{_applescript_string(str(src.resolve()))}") confirm conversions false
                set sourceDoc to active document
                save as active document file name "{_applescript_string(str(dst.resolve()))}" file format format document default
                close sourceDoc saving no
            end tell
        end timeout
        '''
        _run_applescript(script, f"转换 DOC 文件 {src.name}")
        if not dst.exists():
            raise OfficeBackendError(f"转换 DOC 文件后未找到产物: {dst}")

    def convert_pptx_to_pdf(self, src: Path, dst: Path) -> None:
        """调用 PowerPoint 将 PPTX 文件导出为 PDF。"""
        dst.parent.mkdir(parents=True, exist_ok=True)
        script = f'''
        with timeout of 600 seconds
            tell application "Microsoft PowerPoint"
                open (POSIX file "{_applescript_string(str(src.resolve()))}")
                set sourcePresentation to active presentation
                save active presentation in (POSIX file "{_applescript_string(str(dst.resolve()))}") as save as PDF
                close sourcePresentation saving no
            end tell
        end timeout
        '''
        _run_applescript(script, f"转换 PPTX 文件 {src.name}")
        if not dst.exists():
            raise OfficeBackendError(f"转换 PPTX 文件后未找到产物: {dst}")

    def export_pdf(self, docx: Path, pdf: Path) -> None:
        """调用 Word 将 DOCX 导出为 PDF。"""
        stat = self.static_status()
        if not stat.available:
            raise OfficeBackendError(stat.reason)

        docx_abs = docx.resolve()
        pdf_abs = pdf.resolve()

        if pdf_abs.exists():
            try:
                pdf_abs.unlink()
            except OSError:
                pass

        access_dir = _word_access_directory()
        working_docx = access_dir / f"export-{uuid.uuid4().hex}.docx"
        working_pdf = access_dir / f"export-{uuid.uuid4().hex}.pdf"
        try:
            shutil.copy2(docx_abs, working_docx)
        except OSError as exc:
            raise OfficeBackendError(f"无法准备 Word Automation 输入副本: {exc}") from exc

        script = f'''
        with timeout of 600 seconds
            tell application "Microsoft Word"
                set previousAlerts to display alerts
                set myDoc to missing value
                try
                    set display alerts to none
                    open (POSIX file "{_applescript_string(str(working_docx))}")
                    set myDoc to active document
                    -- Word's AppleScript dictionary exposes `save as` on a document
                    -- reference, not on the `active document` property expression.
                    -- Keeping the explicit handle is important when the test host has
                    -- more than one generated document open.
                    save as myDoc file name "{_applescript_string(str(working_pdf))}" file format format PDF
                    close myDoc saving no
                    set myDoc to missing value
                    set display alerts to previousAlerts
                on error errorMessage number errorNumber
                    -- Always close the generated document, including when Word
                    -- fails before the normal close.  Otherwise stale documents
                    -- remain open and can block the next Automation call.
                    if myDoc is not missing value then
                        try
                            close myDoc saving no
                        end try
                    end if
                    set display alerts to previousAlerts
                    error errorMessage number errorNumber
                end try
            end tell
        end timeout
        '''
        try:
            try:
                res = self._execute_subprocess(["osascript", "-e", script], timeout=620)
            except subprocess.TimeoutExpired as exc:
                raise OfficeTimeoutError("Microsoft Word 导出 PDF 超时；请确认 Word 没有被模态对话框阻塞。") from exc
            if res.returncode != 0:
                detail = (res.stderr or res.stdout).strip()
                raise OfficeBackendError(_automation_failure_message(detail))
            if not working_pdf.is_file():
                raise OfficeBackendError("Microsoft Word 未生成 PDF。")
            try:
                # 导入 pagination 的重试逻辑或安全复制
                from lib.pagination import _replace_pdf_with_retry
                _replace_pdf_with_retry(working_pdf, pdf_abs)
            except Exception:
                shutil.copy2(working_pdf, pdf_abs)
            if not pdf_abs.exists():
                raise OfficeBackendError(f"无法保存 Word 导出的 PDF 到目标路径: {pdf_abs}")
        finally:
            for temporary in (working_docx, working_pdf):
                try:
                    temporary.unlink()
                except FileNotFoundError:
                    pass
                except OSError:
                    pass

    def inspect_bookmarks(
        self, docx: Path, pdf: Path, names: Sequence[str]
    ) -> BookmarkReport:
        """排版重新分页并获取各书签物理页与打印页码，并导出最终核验 PDF。"""
        from lib.pagination import _parse_page_map, _replace_pdf_with_retry

        stat = self.static_status()
        if not stat.available:
            raise OfficeBackendError(stat.reason)

        access_dir = _word_access_directory()
        access_token = uuid.uuid4().hex
        working_docx = access_dir / f"pagination-{access_token}.docx"
        working_pdf = access_dir / f"pagination-{access_token}.pdf"

        try:
            shutil.copy2(docx.resolve(), working_docx)
            script = _inspection_script(working_docx, working_pdf, names)
            result = self._execute_subprocess(["osascript", "-e", script], timeout=620)
            if result.returncode:
                detail = (result.stderr or result.stdout).strip()
                message = _automation_failure_message(detail)
                if "-1708" in detail:
                    message += " 请检查 Word 是否出现 Grant File Access（文件夹访问授权）提示；该权限与 Automation 权限不同。"
                raise OfficeBackendError(message)

            if not working_pdf.is_file():
                raise OfficeBackendError("Word 未生成本次分页核验的 PDF，已停止构建。")

            try:
                total_pages = len(PdfReader(str(working_pdf)).pages)
            except Exception as exc:
                raise OfficeBackendError(f"Word 导出 PDF 无法读取: {exc}") from exc

            if total_pages < 1:
                raise OfficeBackendError("Word 导出的 PDF 没有页面。")

            page_map = _parse_page_map(result.stdout, list(names), total_pages)
            _replace_pdf_with_retry(working_pdf, pdf.resolve())
            return BookmarkReport(
                bookmarks=page_map,
                total_pages=total_pages,
                pdf_path=pdf.resolve(),
                fidelity="exact",
            )
        except subprocess.TimeoutExpired as exc:
            raise OfficeTimeoutError("等待 Word 最终分页核验超时；请检查 Word 是否有模态对话框。") from exc
        finally:
            for temporary in (working_docx, working_pdf):
                try:
                    temporary.unlink()
                except FileNotFoundError:
                    pass
                except OSError:
                    pass
