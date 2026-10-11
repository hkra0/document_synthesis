# -*- coding: utf-8 -*-
"""无 Office 或不支持环境的占位后端。"""

import sys
from pathlib import Path
from typing import Sequence
from .base import BackendStatus, BookmarkReport, OfficeUnsupportedError


class NullBackend:
    """提供明确的平台诊断信息的 Null 后端。"""
    name: str = "none"
    fidelity: str = "unsupported"

    def static_status(self) -> BackendStatus:
        available_now = "当前可使用 --plan、格式分析类命令和 smoke_test.py --structure-only。"
        if sys.platform == "win32":
            # Windows + Word 由 WinComBackend 负责；走到这里说明 Office 后端被显式设为 none，
            # 或 Windows COM 后端模块无法导入。未安装 Word 的提示由 WinComBackend 给出。
            reason = "未启用 Microsoft Word 自动化（Office 后端为 none），无法取得精确页码；" + available_now
        elif sys.platform.startswith("linux"):
            reason = "Linux 上没有可用的 Microsoft Word 自动化，无法取得精确页码；" + available_now
        else:
            reason = "未启用 Microsoft Word 自动化，无法取得精确页码；" + available_now
        return BackendStatus(available=False, reason=reason, fidelity="unsupported")

    def probe(self) -> BackendStatus:
        return self.static_status()

    def convert_doc_to_docx(self, src: Path, dst: Path) -> None:
        raise OfficeUnsupportedError(self.static_status().reason)

    def convert_pptx_to_pdf(self, src: Path, dst: Path) -> None:
        raise OfficeUnsupportedError(self.static_status().reason)

    def export_pdf(self, docx: Path, pdf: Path) -> None:
        raise OfficeUnsupportedError(self.static_status().reason)

    def inspect_bookmarks(
        self, docx: Path, pdf: Path, names: Sequence[str]
    ) -> BookmarkReport:
        raise OfficeUnsupportedError(self.static_status().reason)
