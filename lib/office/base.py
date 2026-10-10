# -*- coding: utf-8 -*-
"""Office 后端协议、状态模型与通用异常。"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Protocol, Sequence, runtime_checkable


class OfficeBackendError(RuntimeError):
    """Office 后端调用失败基础异常。"""
    pass


class OfficeUnsupportedError(OfficeBackendError):
    """当前环境或平台不支持指定的 Office 操作。"""
    pass


class OfficeTimeoutError(OfficeBackendError):
    """Office 后端调用超时。"""
    pass


@dataclass
class BackendStatus:
    """Office 后端的可用性与保真度状态。"""
    available: bool
    reason: str
    fidelity: str = "exact"  # "exact" | "draft" | "unsupported"
    details: Dict[str, Any] = field(default_factory=dict)


@dataclass
class BookmarkReport:
    """文档书签页码测量与 PDF 导出结果。"""
    bookmarks: Dict[str, Dict[str, Any]]
    total_pages: int
    pdf_path: Path
    fidelity: str = "exact"  # "exact" | "draft"


@runtime_checkable
class OfficeBackend(Protocol):
    """Office 自动化与排版后端接口协议。"""
    name: str        # "mac-word" | "win-word" | "libreoffice" | "none"
    fidelity: str    # "exact" | "draft" | "unsupported"

    def static_status(self) -> BackendStatus:
        """检查静态前置条件，不弹授权、不启动应用。"""
        ...

    def probe(self) -> BackendStatus:
        """--doctor 专用的只读探针，验证真实自动化控制能力。"""
        ...

    def convert_doc_to_docx(self, src: Path, dst: Path) -> None:
        """将 DOC 文件转换为 DOCX。"""
        ...

    def convert_pptx_to_pdf(self, src: Path, dst: Path) -> None:
        """将 PPTX 文件转换为 PDF。"""
        ...

    def export_pdf(self, docx: Path, pdf: Path) -> None:
        """调用 Word 导出 DOCX 为 PDF。"""
        ...

    def inspect_bookmarks(
        self, docx: Path, pdf: Path, names: Sequence[str]
    ) -> BookmarkReport:
        """排版重新分页并获取各书签物理页与打印页码，并导出最终核验 PDF。"""
        ...
