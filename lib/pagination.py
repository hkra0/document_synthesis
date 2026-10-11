"""Strict, bookmark-based pagination of the assembled document using local Word.

Unlike title-text searches, bookmark positions remain unambiguous when headings
repeat or also appear in the table of contents. No estimated pages are returned.
"""

from collections import Counter
from pathlib import Path
import os
import re
import shutil
import subprocess
import time
import uuid
import zipfile
from typing import Any, Dict

from lxml import etree
from pypdf import PdfReader

from .office import get_backend, OfficeBackendError
from .qa import (
    OfficeExportError,
    _applescript_string,
    _automation_failure_message,
    _word_access_directory,
    word_export_status,
)
from .pagination_types import PageRecord, format_page_number, roman_to_int


_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
_W = "{" + _NS["w"] + "}"


def validate_document_structure(docx_path, required_bookmarks=()):
    """Reject absent/duplicate bookmark targets and dangling internal hyperlinks.

    Checks all Word XML parts, including headers and footers. This is a package
    integrity check, not a replacement for final Word pagination or visual QA.
    """
    names = []
    links = []
    try:
        with zipfile.ZipFile(docx_path) as package:
            for part in package.namelist():
                if not part.startswith("word/") or not part.endswith(".xml"):
                    continue
                root = etree.fromstring(package.read(part), etree.XMLParser(resolve_entities=False))
                names.extend(root.xpath("//w:bookmarkStart/@w:name", namespaces=_NS))
                for link in root.xpath("//w:hyperlink[@w:anchor]", namespaces=_NS):
                    # With a relationship id the anchor addresses another file.
                    if not link.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"):
                        links.append(link.get(_W + "anchor"))
    except (OSError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        raise OfficeExportError(f"无法检查 DOCX 结构: {exc}") from exc
    counts = Counter(names)
    duplicates = sorted(name for name, count in counts.items() if count > 1)
    missing = sorted(set(required_bookmarks).difference(names))
    dangling = sorted(set(links).difference(names))
    problems = []
    if duplicates:
        problems.append("重复书签: " + ", ".join(duplicates))
    if missing:
        problems.append("缺少书签: " + ", ".join(missing))
    if dangling:
        problems.append("无效的本文件目录/超链接: " + ", ".join(dangling))
    if problems:
        raise OfficeExportError("；".join(problems))


def _parse_page_map(output, bookmark_names, total_pages):
    """Parse the deliberately narrow, tab-separated Word response strictly."""
    expected = set(bookmark_names)
    pages = {}
    for line in output.splitlines():
        if not line.strip():
            continue
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] not in expected or fields[0] in pages:
            raise OfficeExportError(f"Word 返回了无效的书签页码记录: {line!r}")
        name, physical, printed = fields
        if not re.fullmatch(r"[0-9]+", physical) or not re.fullmatch(r"[0-9]+", printed):
            raise OfficeExportError(f"Word 未返回整数页码: {name}")
        physical, printed = int(physical), int(printed)
        if not 1 <= physical <= total_pages or printed < 1:
            raise OfficeExportError(f"Word 返回的书签页码超出有效范围: {name}")
        pages[name] = {"physical_page": physical, "printed_page": printed}
    missing = expected.difference(pages)
    if missing:
        raise OfficeExportError("Word 未返回书签页码: " + ", ".join(sorted(missing)))
    return pages


def _inspection_script(docx_path, pdf_path, bookmark_names):
    from .office.mac_applescript import _inspection_script as _is
    return _is(docx_path, pdf_path, bookmark_names)


def _normalise_pdf_label(line: str) -> str:
    """Strip common footer decoration while preserving Roman/letter case."""
    value = line.strip()
    value = re.sub(r"^[\s\-–—·•|｜:：()\[\]{}]+", "", value)
    value = re.sub(r"[\s\-–—·•|｜:：()\[\]{}]+$", "", value)
    return value.strip()


def extract_pdf_page_labels(pdf_path: Path) -> Dict[int, str]:
    """
    通过 PyMuPDF 探查 PDF 页脚 (底部 15%) 与页眉 (顶部 15%) 区域，提取各物理页实际渲染的页码标签。
    返回 {physical_page (1-based): observed_label}。
    """
    labels = {}
    try:
        import pymupdf
        with pymupdf.open(str(pdf_path)) as doc:
            for idx, page in enumerate(doc, 1):
                rect = page.rect
                footer_rect = pymupdf.Rect(rect.x0, rect.y0 + rect.height * 0.85, rect.x1, rect.y1)
                footer_text = page.get_text("text", clip=footer_rect).strip()
                header_rect = pymupdf.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.15)
                header_text = page.get_text("text", clip=header_rect).strip()
                
                found_label = None
                for text_zone in (footer_text, header_text):
                    if not text_zone:
                        continue
                    for line in text_zone.splitlines():
                        candidate = _normalise_pdf_label(line)
                        if not candidate:
                            continue
                        if candidate.isdigit():
                            found_label = candidate
                            break
                        if roman_to_int(candidate) is not None:
                            found_label = candidate
                            break
                        if re.fullmatch(r"[A-Za-z]+", candidate) and len(candidate) <= 3:
                            found_label = candidate
                            break
                    if found_label:
                        break
                if found_label:
                    labels[idx] = found_label
    except Exception:
        pass
    return labels


def page_records_from_map(
    page_map: Dict[str, Dict[str, Any]],
    *,
    part_boundaries: Dict[str, str],
    part_specs: Dict[str, Any],
    sequences: Dict[str, Any],
) -> Dict[str, PageRecord]:
    """Convert compatibility page-map dictionaries to the R8 PageRecord type.

    Boundary records are assigned to the containing part by physical position;
    unknown sequence/section data remains explicit instead of being guessed.
    """
    ordered_parts = []
    for part_id, bookmark in part_boundaries.items():
        record = page_map.get(bookmark)
        if isinstance(record, dict) and isinstance(record.get("physical_page"), int):
            ordered_parts.append((record["physical_page"], part_id, bookmark))
    ordered_parts.sort()
    result: Dict[str, PageRecord] = {}
    for name, raw in page_map.items():
        if not isinstance(raw, dict):
            continue
        owner = None
        for index, (start, part_id, _) in enumerate(ordered_parts):
            end = ordered_parts[index + 1][0] if index + 1 < len(ordered_parts) else None
            if raw.get("physical_page", 0) >= start and (end is None or raw.get("physical_page", 0) < end):
                owner = part_id
                break
        part = part_specs.get(owner) if owner else None
        sequence_id = getattr(part, "page_sequence", None) if part else None
        sequence = sequences.get(sequence_id) if sequence_id else None
        fmt = getattr(sequence, "format", "decimal") if sequence else "decimal"
        expected = (
            raw["expected_label"]
            if "expected_label" in raw
            else format_page_number(int(raw.get("printed_page", 0)), fmt)
        )
        result[name] = PageRecord.from_dict(
            raw,
            section_id=(ordered_parts.index(next(item for item in ordered_parts if item[1] == owner)) + 1) if owner else 1,
            sequence_id=sequence_id,
            expected_label=expected,
        )
    return result


def _replace_pdf_with_retry(
    source: Path,
    target: Path,
    max_retries: int = 5,
    initial_delay: float = 0.1,
    backoff_factor: float = 2.0,
    sleep_func=time.sleep,
) -> None:
    delay = initial_delay
    last_exc = None
    for attempt in range(max_retries):
        try:
            os.replace(source, target)
            return
        except PermissionError as exc:
            last_exc = exc
            if attempt < max_retries - 1:
                sleep_func(delay)
                delay *= backoff_factor
            else:
                break
    raise OfficeExportError(
        f"无法更新分页诊断文件（目标文件可能正在 Word 或 PDF 阅读器中打开）: {target}"
    ) from last_exc


def inspect_document(docx_path, pdf_path, bookmark_names):
    """Export final DOCX and return 1-based physical/printed pages by bookmark.

    ``bookmark_names`` may include a body-start or TOC-start boundary bookmark.
    Word works on a unique private copy and never saves changes into the input.
    The destination PDF is replaced only after both export and page-map checks
    succeed; a stale export can therefore never satisfy verification.
    """
    source = Path(docx_path).resolve()
    destination = Path(pdf_path).resolve()
    names = list(bookmark_names)
    if source == destination or destination.suffix.lower() != ".pdf":
        raise OfficeExportError("分页导出目标必须是独立的 PDF 文件。")
    if len(set(names)) != len(names) or any(
        not isinstance(name, str) or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,39}", name)
        for name in names
    ):
        raise OfficeExportError("分页书签名称必须唯一，且符合 Word 书签命名规则。")
    validate_document_structure(source, names)
    available, reason = word_export_status()
    if not available:
        raise OfficeExportError(reason)

    backend = get_backend()
    try:
        report = backend.inspect_bookmarks(source, destination, names)
    except (OfficeBackendError, OSError) as exc:
        raise OfficeExportError(str(exc)) from exc

    # 草稿交付需要在文件名、诊断文件和终端输出中统一标注（计划 P5）。
    # 在该标注落地前，任何非 exact 结果都不能进入目录回填与发布。
    if report.fidelity != "exact":
        raise OfficeExportError(
            f"当前 Office 后端只能提供 {report.fidelity} 保真度的页码，"
            "草稿交付尚未支持；为避免发布未经 Word 核验的目录页码，已停止构建。"
        )

    page_map = report.bookmarks
    pdf_labels = extract_pdf_page_labels(destination)
    for name, record in page_map.items():
        phys = record.get("physical_page")
        if phys:
            observed = pdf_labels.get(phys)
            if observed:
                record["observed_label"] = observed
    return page_map
