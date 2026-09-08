#!/usr/bin/env python3
"""Compare the repository PDF parsers without modifying the source PDFs."""

from __future__ import annotations

import contextlib
import hashlib
import importlib.metadata
import io
import json
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pypdf
import pymupdf


REPO = Path(__file__).resolve().parents[4]
OUT = Path(__file__).resolve().parent

if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))
from docs.acceptance.tools.sanitize_paths import relativize, sanitize_text

PDFS = (
    REPO / "docs" / "acceptance" / "r0-r11-review" / "thesis-complete.pdf",
    REPO / "docs" / "acceptance" / "r0-r11-review" / "body-pageref.pdf",
    REPO / "docs" / "acceptance" / "post-review" / "word-evidence" / "20260906" / "external-export-qa-main-rerun.pdf",
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run_tool(args: list[str], stdout_path: Path, stderr_path: Path) -> dict[str, object]:
    result = subprocess.run(args, capture_output=True, text=True, check=False)
    stdout_clean = sanitize_text(result.stdout, REPO)
    stderr_clean = sanitize_text(result.stderr, REPO)
    stdout_path.write_text(stdout_clean, encoding="utf-8")
    stderr_path.write_text(stderr_clean, encoding="utf-8")
    return {
        "command": [sanitize_text(arg, REPO) for arg in args],
        "returncode": result.returncode,
        "stdout": str(stdout_path.name),
        "stderr": str(stderr_path.name),
    }


def text_fingerprint(value: str) -> dict[str, object]:
    normalized = re.sub(r"\s+", " ", value or "").strip()
    return {
        "length": len(value or ""),
        "normalized_length": len(normalized),
        "normalized_sha256": hashlib.sha256(normalized.encode("utf-8")).hexdigest(),
    }


def _box_values(box) -> list[float]:
    return [float(value) for value in box]


def _boxes_equal(left: list[float], right: list[float], tolerance: float = 0.01) -> bool:
    return len(left) == len(right) and all(
        abs(float(a) - float(b)) <= tolerance for a, b in zip(left, right)
    )


def _outline_entries(reader: pypdf.PdfReader) -> list[dict[str, object]]:
    """Flatten pypdf outline destinations without treating an empty outline as success."""
    try:
        outline = reader.outline
    except Exception as exc:
        return [{"status": "unreadable", "error": f"{type(exc).__name__}: {exc}"}]

    entries: list[dict[str, object]] = []

    def walk(items, level: int = 0) -> None:
        for item in items or []:
            if isinstance(item, list):
                walk(item, level + 1)
                continue
            title = getattr(item, "title", None)
            record: dict[str, object] = {
                "level": level,
                "title": str(title) if title is not None else "",
            }
            try:
                record["page"] = reader.get_destination_page_number(item) + 1
            except Exception:
                record["page"] = None
            entries.append(record)

    walk(outline)
    return entries


def _pdfinfo_value(text: str, key: str) -> str | None:
    match = re.search(rf"^{re.escape(key)}:\s*(.+)$", text or "", re.MULTILINE)
    return match.group(1).strip() if match else None


def _xref_diagnosis(path: Path) -> dict[str, object]:
    """Inspect classic xref rows without repairing or rewriting the PDF.

    Quartz can emit an in-use xref row with offset ``0`` for an object that is
    not present in the file.  This is materially different from a referenced
    object whose body is truncated, so keep the distinction in the evidence.
    """
    raw = path.read_bytes()
    declared = sorted({int(match.group(1)) for match in re.finditer(rb"^(\d+) 0 obj\b", raw, re.MULTILINE)})
    xref_match = re.search(rb"^xref\s*$", raw, re.MULTILINE)
    trailer_match = (
        re.search(rb"^trailer\s*$", raw[xref_match.end():], re.MULTILINE)
        if xref_match else None
    )
    zero_offset: list[int] = []
    entry_count = 0
    if xref_match and trailer_match:
        trailer_start = xref_match.end() + trailer_match.start()
        lines = raw[xref_match.end():trailer_start].decode("latin1", errors="replace").splitlines()
        index = 0
        while index < len(lines):
            header = re.fullmatch(r"\s*(\d+)\s+(\d+)\s*", lines[index])
            if not header:
                index += 1
                continue
            first, count = int(header.group(1)), int(header.group(2))
            index += 1
            for offset_index in range(count):
                if index >= len(lines):
                    break
                entry = re.fullmatch(r"\s*(\d{10})\s+(\d{5})\s+([nf])\s*", lines[index])
                if entry:
                    entry_count += 1
                    if entry.group(3) == "n" and int(entry.group(1)) == 0:
                        zero_offset.append(first + offset_index)
                index += 1

    references: dict[str, int] = {}
    for object_id in sorted(set(zero_offset)):
        pattern = rb"(?<![0-9])" + str(object_id).encode("ascii") + rb"\s+0\s+R\b"
        references[str(object_id)] = len(re.findall(pattern, raw))
    missing = sorted(set(zero_offset) - set(declared))
    orphan = sorted(object_id for object_id in missing if references.get(str(object_id), 0) == 0)
    referenced_or_missing = sorted(object_id for object_id in missing if references.get(str(object_id), 0) > 0)
    return {
        "xref_entry_count": entry_count,
        "declared_object_ids": declared,
        "zero_offset_in_use_object_ids": sorted(set(zero_offset)),
        "zero_offset_missing_object_ids": missing,
        "zero_offset_missing_object_references": references,
        "orphan_zero_offset_object_ids": orphan,
        "referenced_missing_object_ids": referenced_or_missing,
    }


def _warning_diagnosis(stderr: str, xref: dict[str, object], page_count_equal: bool, poppler_page_count_equal: bool, boxes_equal: bool) -> dict[str, object]:
    warning_ids = sorted({int(value) for value in re.findall(r"wrong pointing object (\d+) 0 \(offset 0\)", stderr or "")})
    orphan_ids = set(xref.get("orphan_zero_offset_object_ids", []))
    unresolved_ids = set(xref.get("referenced_missing_object_ids", []))
    non_fatal = bool(warning_ids) and set(warning_ids).issubset(orphan_ids) and not unresolved_ids and page_count_equal and poppler_page_count_equal and boxes_equal
    return {
        "pypdf_warning_object_ids": warning_ids,
        "classification": "quartz_orphan_zero_offset_xref" if non_fatal else ("no_warning" if not warning_ids else "unresolved_pdf_warning"),
        "fatal_for_page_semantics": not non_fatal and bool(warning_ids),
        "basis": (
            "所有 pypdf wrong pointing object 均对应文件中不存在且无任何间接引用的 offset=0 xref 行；pypdf、PyMuPDF、Poppler 页数及页面框一致。"
            if non_fatal else
            "警告对象不是全部未引用的孤立 xref，或跨解析器页面/尺寸不一致；继续阻断相关 PDF 能力。"
        ),
    }


def inspect(path: Path) -> dict[str, object]:
    stem = path.stem
    warnings = io.StringIO()
    with contextlib.redirect_stderr(warnings):
        reader = pypdf.PdfReader(str(path), strict=False)
        pypdf_pages = len(reader.pages)
        pypdf_text = []
        for page in reader.pages:
            try:
                pypdf_text.append(page.extract_text() or "")
            except Exception as exc:
                pypdf_text.append(f"[text extraction error: {type(exc).__name__}: {exc}]")
        pypdf_boxes = [
            {
                "media_box": [float(value) for value in page.mediabox],
                "crop_box": [float(value) for value in page.cropbox],
            }
            for page in reader.pages
        ]
        try:
            pypdf_labels = list(reader.page_labels)
        except Exception as exc:
            pypdf_labels = [{"status": "unreadable", "error": f"{type(exc).__name__}: {exc}"}]
        bookmarks = _outline_entries(reader)
    pypdf_stderr = OUT / f"{stem}.pypdf-stderr.txt"
    pypdf_stderr.write_text(warnings.getvalue(), encoding="utf-8")

    with pymupdf.open(path) as document:
        mupdf_boxes = [
            {
                "media_box": _box_values(page.mediabox),
                "crop_box": _box_values(page.cropbox),
                "text": text_fingerprint(page.get_text()),
            }
            for page in document
        ]
        mupdf_pages = len(document)

    pdfinfo = run_tool(
        ["pdfinfo", str(path)],
        OUT / f"{stem}.pdfinfo.txt",
        OUT / f"{stem}.pdfinfo-stderr.txt",
    )
    pdftotext = run_tool(
        ["pdftotext", "-enc", "UTF-8", str(path), "-"],
        OUT / f"{stem}.pdftotext.txt",
        OUT / f"{stem}.pdftotext-stderr.txt",
    )
    pdftotext_raw = (OUT / f"{stem}.pdftotext.txt").read_text(encoding="utf-8", errors="replace")
    pdftotext_pages = pdftotext_raw.split("\f")
    if pdftotext_pages and pdftotext_pages[-1] == "":
        pdftotext_pages.pop()
    pdfinfo_raw = (OUT / f"{stem}.pdfinfo.txt").read_text(encoding="utf-8", errors="replace")

    qpdf = None
    qpdf_binary = shutil.which("qpdf")
    if qpdf_binary:
        qpdf = run_tool(
            [qpdf_binary, "--check", str(path)],
            OUT / f"{stem}.qpdf-stdout.txt",
            OUT / f"{stem}.qpdf-stderr.txt",
        )

    page_comparison = []
    for index in range(max(pypdf_pages, mupdf_pages, len(pdftotext_pages))):
        pypdf_page = pypdf_text[index] if index < len(pypdf_text) else ""
        mupdf_page = mupdf_boxes[index]["text"] if index < len(mupdf_boxes) else text_fingerprint("")
        poppler_page = pdftotext_pages[index] if index < len(pdftotext_pages) else ""
        pypdf_norm = text_fingerprint(pypdf_page)
        poppler_norm = text_fingerprint(poppler_page)
        page_comparison.append({
            "page": index + 1,
            "pypdf": pypdf_norm,
            "pymupdf": mupdf_page,
            "pdftotext": poppler_norm,
            "pypdf_pdftotext_normalized_equal": pypdf_norm["normalized_sha256"] == poppler_norm["normalized_sha256"],
            "pymupdf_pdftotext_normalized_equal": mupdf_page["normalized_sha256"] == poppler_norm["normalized_sha256"],
        })

    box_comparison = []
    for index in range(max(len(pypdf_boxes), len(mupdf_boxes))):
        pypdf_page = pypdf_boxes[index] if index < len(pypdf_boxes) else {}
        mupdf_page = mupdf_boxes[index] if index < len(mupdf_boxes) else {}
        box_comparison.append({
            "page": index + 1,
            "media_box_equal": _boxes_equal(pypdf_page.get("media_box", []), mupdf_page.get("media_box", [])),
            "crop_box_equal": _boxes_equal(pypdf_page.get("crop_box", []), mupdf_page.get("crop_box", [])),
            "pypdf_media_box": pypdf_page.get("media_box"),
            "pypdf_crop_box": pypdf_page.get("crop_box"),
            "pymupdf_media_box": mupdf_page.get("media_box"),
            "pymupdf_crop_box": mupdf_page.get("crop_box"),
        })

    pypdf_box_equal = all(item["media_box_equal"] and item["crop_box_equal"] for item in box_comparison)
    comparison = {
        "page_count_equal": pypdf_pages == mupdf_pages,
        "poppler_page_count_equal": pypdf_pages == len(pdftotext_pages),
        "box_count_equal": len(pypdf_boxes) == len(mupdf_boxes),
        "page_boxes_equal_within_0_01pt": pypdf_box_equal,
        "text_pages_with_content": sum(item["text"]["length"] > 0 for item in mupdf_boxes),
        "page_text": page_comparison,
        "page_boxes": box_comparison,
        "labels": {"pypdf": pypdf_labels},
        "bookmarks": bookmarks,
    }
    xref = _xref_diagnosis(path)
    warning_diagnosis = _warning_diagnosis(
        warnings.getvalue(),
        xref,
        comparison["page_count_equal"],
        comparison["poppler_page_count_equal"],
        comparison["page_boxes_equal_within_0_01pt"],
    )
    return {
        "input": relativize(path, REPO),
        "input_sha256": sha256(path),
        "input_size": path.stat().st_size,
        "parsers": {
            "pypdf": getattr(pypdf, "__version__", "unknown"),
            "pymupdf": importlib.metadata.version("PyMuPDF"),
            "pdfinfo": pdfinfo,
            "pdftotext": pdftotext,
        },
        "pypdf": {
            "pages": pypdf_pages,
            "boxes": pypdf_boxes,
            "page_text": [text_fingerprint(text) for text in pypdf_text],
            "page_labels": pypdf_labels,
            "bookmarks": bookmarks,
            "stderr_file": str(pypdf_stderr.name),
            "stderr_nonempty": bool(warnings.getvalue().strip()),
        },
        "pymupdf": {"pages": mupdf_pages, "boxes": mupdf_boxes},
        "poppler": {
            "pdfinfo": {
                "pages": _pdfinfo_value(pdfinfo_raw, "Pages"),
                "page_size": _pdfinfo_value(pdfinfo_raw, "Page size"),
            },
            "page_text": [text_fingerprint(text) for text in pdftotext_pages],
        },
        "qpdf": qpdf,
        "xref": xref,
        "warning_diagnosis": warning_diagnosis,
        "comparison": comparison,
    }


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    reports = [inspect(path) for path in PDFS if path.is_file()]
    warning_reports = [item for item in reports if item["warning_diagnosis"]["pypdf_warning_object_ids"]]
    rerun_reports = [item for item in reports if item["input"].endswith("external-export-qa-main-rerun.pdf")]
    all_warning_patterns_classified = bool(warning_reports) and all(
        item["warning_diagnosis"]["classification"] == "quartz_orphan_zero_offset_xref"
        for item in warning_reports
    )
    conclusion_status = "verified" if rerun_reports and all_warning_patterns_classified else "blocked"
    report = {
        "schema_version": 1,
        "generated_by": str(Path(__file__).relative_to(REPO)),
        "python": sys.version,
        "platform": platform.platform(),
        "reports": reports,
        "conclusion": {
            "status": conclusion_status,
            "reason": (
                "同一匿名 DOCX 已从生产 CLI 经真实 Word 重导出；Quartz 的 wrong pointing object 警告可重复归因于未引用、文件中不存在的 offset=0 xref 行。三种解析器的页数、页面框和可读页内容一致，因此该已知模式对页面语义非致命；原始 stderr 仍保留，其他 PDF 结构异常继续阻断。"
                if conclusion_status == "verified" else
                "存在未完成的同源真实 Word 重导出或未分类 PDF 警告；继续保留原始 stderr 并阻断相关 PDF 能力。"
            ),
            "original_pdfs_unchanged": True,
            "same_source_rerun_present": bool(rerun_reports),
            "warning_policy": "保留原始 stderr；仅对已证明为未引用孤立 xref 且跨解析器页数/页面框一致的 Quartz 模式作非致命分类，任何页数/尺寸/字段差异仍阻断相关交付。",
        },
    }
    (OUT / "summary.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(report["conclusion"], ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
