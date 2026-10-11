# -*- coding: utf-8 -*-
"""Office 自动化后端抽象包。

通过统一接口协议隔离具体平台 Office 驱动实现。
"""

import os
import sys
from contextlib import contextmanager
from typing import Iterator, Optional

from .base import (
    BackendStatus,
    BookmarkReport,
    OfficeBackend,
    OfficeBackendError,
    OfficeTimeoutError,
    OfficeUnsupportedError,
)
from .null import NullBackend

__all__ = [
    "BackendStatus",
    "BookmarkReport",
    "OfficeBackend",
    "OfficeBackendError",
    "OfficeTimeoutError",
    "OfficeUnsupportedError",
    "NullBackend",
    "get_backend",
    "office_session",
    "set_backend",
]

_ACTIVE_BACKEND: Optional[OfficeBackend] = None


def set_backend(backend: Optional[OfficeBackend]) -> None:
    """显式设置全局活动后端（用于单元测试与依赖注入）。"""
    global _ACTIVE_BACKEND
    _ACTIVE_BACKEND = backend


@contextmanager
def office_session() -> Iterator[OfficeBackend]:
    """一次构建（或一次 smoke 检查）内复用同一个 Office 实例。

    后端提供 ``session()`` 时（Windows COM），块内 :func:`get_backend` 都返回同一个后端
    实例，Word 操作共用一个 worker 与一个 Word；其他后端（macOS、空后端）原样执行，
    行为不变。
    """
    global _ACTIVE_BACKEND
    backend = get_backend()
    start = getattr(backend, "session", None)
    if start is None:
        yield backend
        return
    previous = _ACTIVE_BACKEND
    _ACTIVE_BACKEND = backend
    try:
        with start():
            yield backend
    finally:
        _ACTIVE_BACKEND = previous


def get_backend(name: Optional[str] = None) -> OfficeBackend:
    """根据名称、环境变量或当前平台选择 Office 自动化后端。

    选择优先级：
    1. 显式 set_backend() 设置的后端（仅在未指定 name 时生效）
    2. 显式参数 name
    3. 环境变量 DOCUMENT_SYNTHESIS_OFFICE_BACKEND
    4. 默认 'auto' 策略：
       - macOS (darwin) -> MacAppleScriptBackend
       - Windows (win32) -> WinComBackend (P3) / 降级为 NullBackend
       - Linux/其他平台 -> NullBackend

    LibreOffice 草稿后端（P5）尚未实现，显式选择时抛出 OfficeUnsupportedError，
    不静默退回其他后端。
    """
    global _ACTIVE_BACKEND
    if name is None and _ACTIVE_BACKEND is not None:
        return _ACTIVE_BACKEND

    target = name or os.environ.get("DOCUMENT_SYNTHESIS_OFFICE_BACKEND", "auto")
    target = target.strip().lower()

    if target in ("auto", "word"):
        if sys.platform == "darwin":
            from .mac_applescript import MacAppleScriptBackend
            return MacAppleScriptBackend()
        elif sys.platform == "win32":
            try:
                from .win_com import WinComBackend
                return WinComBackend()
            except ImportError:
                return NullBackend()
        else:
            return NullBackend()

    if target in ("mac-word", "mac", "applescript"):
        from .mac_applescript import MacAppleScriptBackend
        return MacAppleScriptBackend()

    if target in ("win-word", "win", "com"):
        try:
            from .win_com import WinComBackend
            return WinComBackend()
        except ImportError:
            return NullBackend()

    if target in ("libreoffice", "soffice"):
        raise OfficeUnsupportedError("LibreOffice 草稿后端尚未支持。")

    if target in ("none", "null"):
        return NullBackend()

    raise ValueError(f"未知的 Office 后端名称: {target}")
