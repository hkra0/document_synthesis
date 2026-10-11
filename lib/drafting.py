# -*- coding: utf-8 -*-
"""草稿构建：没有可精确分页的 Office 时，生成页码交给 Word 更新的交付物。

草稿只在显式 ``--allow-draft``（环境变量 ``DOCUMENT_SYNTHESIS_ALLOW_DRAFT=1``）
且当前 Office 后端无法 ``exact`` 构建时启用。草稿交付物的文件名带 ``_draft``
后缀，元数据写入独立文件，因此永远不会覆盖已有的精确交付物。
"""

import os
from pathlib import Path

ALLOW_DRAFT_ENV = "DOCUMENT_SYNTHESIS_ALLOW_DRAFT"
DRAFT_SUFFIX = "_draft"
DRAFT_METADATA_NAME = "build-metadata_draft.json"
DRAFT_NOTICE = (
    "草稿：目录页码与版式未经 Microsoft Word 核验，目前显示为“?”。"
    "在 Word 中打开后，等排版完成（状态栏页数不再变化），全选并按 F9 更新域，页码才会填入"
    "（macOS 为 Command+A 后按 F9，部分键盘需同时按 fn）。"
)


def draft_requested() -> bool:
    return os.environ.get(ALLOW_DRAFT_ENV) == "1"


def build_fidelity() -> str:
    """本次构建的保真度：能精确分页时总是 ``exact``；否则仅在显式允许时为 ``draft``。

    不能精确分页又没有允许草稿时也返回 ``exact``，由分页核验给出原有的失败原因。
    """
    from .office import get_backend

    status = get_backend().static_status()
    if status.available and status.fidelity == "exact":
        return "exact"
    return "draft" if draft_requested() else "exact"


def draft_filename(filename: str) -> str:
    """``demo.docx`` → ``demo_draft.docx``。"""
    path = Path(filename)
    return f"{path.stem}{DRAFT_SUFFIX}{path.suffix}"
