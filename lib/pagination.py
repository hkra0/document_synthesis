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
import uuid
import zipfile

from lxml import etree
from pypdf import PdfReader

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
    access_dir = _word_access_directory()
    access_token = uuid.uuid4().hex
    working_docx = access_dir / f"pagination-{access_token}.docx"
    working_pdf = access_dir / f"pagination-{access_token}.pdf"
    try:
        shutil.copy2(source, working_docx)
        result = subprocess.run(
            ["osascript", "-e", _inspection_script(working_docx, working_pdf, names)],
            capture_output=True, text=True, timeout=620,
        )
        if result.returncode:
            detail = (result.stderr or result.stdout).strip()
            message = _automation_failure_message(detail)
            if "-1708" in detail:
                message += " 请检查 Word 是否出现 Grant File Access（文件夹访问授权）提示；该权限与 Automation 权限不同。"
            raise OfficeExportError(message)
        if not working_pdf.is_file():
            raise OfficeExportError("Word 未生成本次分页核验的 PDF，已停止构建。")
        try:
            total_pages = len(PdfReader(str(working_pdf)).pages)
        except Exception as exc:
            raise OfficeExportError(f"Word 导出 PDF 无法读取: {exc}") from exc
        if total_pages < 1:
            raise OfficeExportError("Word 导出的 PDF 没有页面。")
        page_map = _parse_page_map(result.stdout, names, total_pages)
        os.replace(working_pdf, destination)
        pdf_labels = extract_pdf_page_labels(destination)
        for name, record in page_map.items():
            phys = record["physical_page"]
            observed = pdf_labels.get(phys)
            if observed:
                record["observed_label"] = observed
        return page_map
    except subprocess.TimeoutExpired as exc:
        raise OfficeExportError("等待 Word 最终分页核验超时；请检查 Word 是否有模态对话框。") from exc
    except OSError as exc:
        raise OfficeExportError(f"Word 最终分页核验失败: {exc}") from exc
    finally:
        for temporary in (working_docx, working_pdf):
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass
            except OSError:
                pass
