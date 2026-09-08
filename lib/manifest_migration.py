# -*- coding: utf-8 -*-
"""Explicit, non-destructive migration of legacy project manifests."""

from __future__ import annotations

import copy
import json
import os
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional, Union

from .config import ConfigError, ProjectConfig


def _rebase_reference(value: Any, source_dir: Path, target_dir: Path) -> Any:
    if not isinstance(value, str) or not value or value.startswith("preset:") or value.startswith("/"):
        return value
    candidate = (source_dir / value).resolve()
    if candidate.exists():
        return os.path.relpath(candidate, target_dir)
    return value


def _rebase_paths(data: Dict[str, Any], source_dir: Path, target_dir: Path) -> Dict[str, Any]:
    result = copy.deepcopy(data)
    source = result.get("source")
    if isinstance(source, dict) and "file" in source:
        source["file"] = _rebase_reference(source["file"], source_dir, target_dir)
    fmt = result.get("format")
    if isinstance(fmt, dict) and "ref" in fmt:
        fmt["ref"] = _rebase_reference(fmt["ref"], source_dir, target_dir)
    formatting = result.get("formatting")
    if isinstance(formatting, dict):
        for key in ("role_map", "scopes"):
            if key == "scopes" and isinstance(formatting.get(key), list):
                for scope in formatting[key]:
                    if isinstance(scope, dict) and "file" in scope:
                        scope["file"] = _rebase_reference(scope["file"], source_dir, target_dir)
            elif key in formatting:
                formatting[key] = _rebase_reference(formatting[key], source_dir, target_dir)
    cover = result.get("cover")
    if isinstance(cover, dict) and "template" in cover:
        cover["template"] = _rebase_reference(cover["template"], source_dir, target_dir)
    return result


def migrate_manifest_data(
    raw: Dict[str, Any],
    *,
    source_dir: Optional[Union[str, Path]] = None,
    target_dir: Optional[Union[str, Path]] = None,
) -> Dict[str, Any]:
    """Convert v1/v2 to an explicit v2/v3-compatible manifest.

    v1 keeps its three historical deliveries, while v2 gains an explicit
    ``output.documents`` list when it relied on defaults. v3 is validated and
    copied without an invented semantic change.
    """
    if not isinstance(raw, dict):
        raise ConfigError("manifest 根节点必须是 JSON 对象。")
    version = raw.get("schema_version")
    if type(version) is not int or version not in (1, 2, 3):
        raise ConfigError("只支持迁移 schema_version 1、2 或 3。")
    result = copy.deepcopy(raw)
    project = result.get("project_name")
    if not isinstance(project, str) or not project:
        raise ConfigError("manifest 缺少有效 project_name。")

    if version == 1:
        old_output = result.get("output") if isinstance(result.get("output"), dict) else {}
        result["schema_version"] = 2
        result["output"] = {
            "documents": [
                {
                    "id": "compiled_document",
                    "filename": old_output.get("compiled_document", f"{project}_合成材料.docx"),
                    "parts": ["body"],
                },
                {
                    "id": "cover_toc",
                    "filename": old_output.get("cover_toc", f"{project}_封面+目录.docx"),
                    "parts": ["cover", "toc"],
                    "toc": {"reference": "toc_body", "links": "none"},
                },
                {
                    "id": "toc_body",
                    "filename": old_output.get("toc_body", f"{project}_目录+正文.docx"),
                    "parts": ["toc", "body"],
                    "toc": {"reference": "toc_body", "links": "internal"},
                },
            ]
        }
    elif version == 2:
        output = result.get("output")
        if not isinstance(output, dict) or "documents" not in output:
            result["output"] = {
                "documents": [{
                    "id": "main",
                    "filename": f"{project}_完整文档.docx",
                    "parts": ["cover", "toc", "body"],
                }]
            }

    source_root = Path(source_dir).resolve() if source_dir else None
    target_root = Path(target_dir).resolve() if target_dir else source_root
    if source_root and target_root:
        result = _rebase_paths(result, source_root, target_root)
    ProjectConfig(result, project_dir=target_root or Path.cwd())
    return result


def migrate_manifest_file(
    source_path: Union[str, Path],
    output_path: Union[str, Path],
    *,
    replace: bool = False,
) -> Dict[str, Any]:
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if not source.is_file():
        raise ConfigError(f"manifest 文件不存在: {source}")
    # Migration is explicitly non-destructive.  ``resolve()`` catches normal
    # relative/symlink aliases; ``samefile()`` additionally catches hard links
    # when the destination already exists.  ``replace`` may replace an
    # independent destination, never the source inode itself.
    same_file = source == output
    if not same_file and output.exists():
        try:
            same_file = os.path.samefile(source, output)
        except OSError:
            same_file = False
    if same_file:
        raise ConfigError(
            "MIGRATION_SOURCE_EQUALS_OUTPUT: 迁移输出必须是独立文件，不能覆盖源 manifest。"
        )
    if output.exists() and not replace:
        raise ConfigError(f"迁移输出已存在，若要覆盖请显式指定 replace: {output}")
    try:
        raw = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigError(f"无法读取 manifest: {exc}") from exc
    result = migrate_manifest_data(raw, source_dir=source.parent, target_dir=output.parent)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    temporary = None
    try:
        fd, temporary = tempfile.mkstemp(prefix=f".{output.name}.", suffix=".tmp", dir=str(output.parent))
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        if output.exists() and not replace:
            raise ConfigError(f"迁移输出已存在，若要覆盖请显式指定 replace: {output}")
        os.replace(temporary, output)
        temporary = None
    except ConfigError:
        raise
    except (OSError, ValueError) as exc:
        raise ConfigError(f"无法原子写入迁移输出: {exc}") from exc
    finally:
        if temporary:
            try:
                os.unlink(temporary)
            except OSError:
                pass
    return result
