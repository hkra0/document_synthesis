# -*- coding: utf-8 -*-
"""
内容语义完整性与最终有效格式验证器 (lib/content_integrity.py)

1. 语义完整性 (Semantic Inventory & Content Preservation):
   提取作者可见文本序列、段落多重集（Multiset）、表格拓扑、嵌入媒体哈希、OMML 数学公式、脚注与超链接。
   校验交付物是否完整保留源文档全部语义内容，支持显式声明的标题替换映射，严格防止同名段落重复丢失或正文截断。

2. 最终有效格式验证 (Effective Format Verification):
   基于 docx_inspector 对已装配的交付文档执行只读有效样式求值，
   逐项验证托管角色（body, heading.1..9, title, toc 等）的最终有效字体、字号、间距与大纲级别，
   确保装配、分节与页眉页脚处理后格式包契约依然 100% 成立。
"""

import collections
import copy
import hashlib
import re
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Counter, Dict, Iterable, List, Mapping, Optional, Set, Tuple, Union
import zipfile

from docx import Document
from docx.oxml.ns import qn
import lxml.etree

from .docx_inspector import (
    BlockInspection,
    DocumentInspection,
    DocumentInspectionLimits,
    inspect_docx,
)
from .document_parts import get_part_boundary_bookmark
from .format_schema import (
    Diagnostic,
    FormatDiagnosticCode,
    LengthValue,
    LineSpacing,
    ParagraphStyle,
    ResolvedFormat,
    RunStyle,
    StyleDefinition,
)

XML_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}
MEDIA_BLIP_TAG = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"


class ContentIntegrityError(Exception):
    """内容语义完整性校验失败"""
    pass


class FormatVerificationError(Exception):
    """交付物最终格式核验失败"""
    pass


@dataclass
class TableTopology:
    """表格拓扑与单元格数据"""
    row_count: int
    col_count: int
    cell_texts: List[str]
    grid_widths: List[int]

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class ObjectOccurrence:
    """One embedded object occurrence with host and relationship evidence."""
    kind: str
    fingerprint: str
    part_uri: str
    host_element: str
    local_index: int
    host_text: str = ""
    host_ordinal: Optional[int] = None
    relationship_id: Optional[str] = None
    target_part: Optional[str] = None
    story_type: str = "body"
    source_sha256: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class SemanticInventory:
    """文档语义清单快照"""
    source_sha256: str
    paragraphs: List[str]
    paragraph_counts: Dict[str, int]
    tables: List[TableTopology]
    media_hashes: Set[str]
    math_formulas: List[str]
    footnotes: List[str]
    fields: List[str]
    bookmarks: Set[str]
    hyperlinks: List[str]
    # 下面字段保留旧字段的兼容性，同时记录实例级证据，避免 set 比较吞掉重复对象。
    media_instances: List[str] = field(default_factory=list)
    media_occurrences: List[ObjectOccurrence] = field(default_factory=list)
    footnote_definitions: Dict[str, str] = field(default_factory=dict)
    endnotes: List[str] = field(default_factory=list)
    endnote_definitions: Dict[str, str] = field(default_factory=dict)
    field_dependencies: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source_sha256": self.source_sha256,
            "paragraph_total": len(self.paragraphs),
            "unique_paragraphs": len(self.paragraph_counts),
            "table_total": len(self.tables),
            "media_total": len(self.media_instances or self.media_hashes),
            "media_hashes": sorted(list(self.media_hashes)),
            "media_instance_total": len(self.media_instances),
            "media_occurrence_total": len(self.media_occurrences),
            "media_occurrences": [item.to_dict() for item in self.media_occurrences],
            "math_formula_total": len(self.math_formulas),
            "footnote_total": len(self.footnotes),
            "footnote_definition_total": len(self.footnote_definitions),
            "endnote_total": len(self.endnotes),
            "field_total": len(self.fields),
            "field_dependency_total": len(self.field_dependencies),
            "bookmark_total": len(self.bookmarks),
            "hyperlink_total": len(self.hyperlinks),
        }


@dataclass(frozen=True)
class ExpectedInventory:
    """由 PreparedBuild/实际交付部件推导出的期望清单。

    ``semantic`` 是真正参与门禁的清单；其余字段是可审阅的来源、生成内容
    和合法 ID 重映射证据。这样内容门禁不会从最终整篇 body 的字面文本猜测
    期望范围，也不会把目录或封面误当作源正文。
    """

    semantic: SemanticInventory
    source_order: Tuple[str, ...] = ()
    source_regions: Dict[str, str] = field(default_factory=dict)
    allowed_generated_text: Dict[str, int] = field(default_factory=dict)
    allowed_generated_media: Dict[str, int] = field(default_factory=dict)
    # Finite, part-scoped allowances for objects created by the assembler
    # (for example a template logo or an explicit image node).  A hash-only
    # allow-list would let an extra occurrence in another part pass.
    allowed_generated_objects: Tuple[Dict[str, Any], ...] = ()
    replacements: Dict[str, str] = field(default_factory=dict)
    allowed_id_remaps: Dict[str, str] = field(default_factory=dict)
    provenance: Dict[str, Any] = field(default_factory=dict)
    coverage: Dict[str, Any] = field(default_factory=dict)
    content_parts: Tuple[str, ...] = ()

    def to_dict(self) -> Dict[str, Any]:
        data = self.semantic.to_dict()
        data.update({
            "source_order": list(self.source_order),
            "source_regions": dict(self.source_regions),
            "allowed_generated_text": dict(self.allowed_generated_text),
            "allowed_generated_media": dict(self.allowed_generated_media),
            "allowed_generated_objects": [dict(item) for item in self.allowed_generated_objects],
            "replacements": dict(self.replacements),
            "allowed_id_remaps": dict(self.allowed_id_remaps),
            "provenance": dict(self.provenance),
            "coverage": dict(self.coverage),
            "content_parts": list(self.content_parts),
        })
        return data


@dataclass
class FormatVerificationReport:
    """交付物有效格式验证报告"""
    passed: bool
    verified_count: int
    violation_count: int
    verified_roles: List[str]
    violations: List[str]
    unverified_attributes: List[str]
    details: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def _span_style_value(span: Dict[str, Any], attribute: str) -> Any:
    style = span.get("style") if isinstance(span, dict) else None
    if isinstance(style, dict):
        return style.get(attribute)
    return getattr(style, attribute, None) if style is not None else None


def _inline_value_equal(expected: Any, actual: Any, attribute: str) -> bool:
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        return abs(float(expected) - float(actual)) <= (0.5 if attribute == "size_pt" else 0.01)
    if attribute in {"east_asia", "latin"} and isinstance(expected, str) and isinstance(actual, str):
        if expected.lower() == actual.lower():
            return True
        return {expected, actual} in [
            {"宋体", "SimSun"}, {"黑体", "SimHei"},
            {"楷体", "KaiTi"}, {"仿宋", "FangSong"},
        ]
    if attribute == "color" and isinstance(expected, str) and isinstance(actual, str):
        if {expected.lower(), actual.lower()} <= {"auto", "000000", "00000000"}:
            return True
    return expected == actual


def _logical_inline_segments(
    block: BlockInspection,
    expected_node: Any,
    attribute: str,
    target_value: Any = None,
) -> Optional[List[Dict[str, Any]]]:
    """Compare preserved inline properties over logical text intervals."""
    expected_spans = list(
        expected_node.semantic_summary.get("source_effective_run_spans", ()) or ()
    )
    actual_spans = list(getattr(block, "effective_run_spans", ()) or ())
    if not expected_spans or not actual_spans:
        return None
    if not any(_span_style_value(span, attribute) is not None for span in expected_spans):
        return None

    direct_spans = list(
        expected_node.semantic_summary.get("source_inline_emphasis_spans", ()) or ()
    )
    preserve_source = not bool(expected_node.managed_properties)
    has_direct_evidence = any(
        span.get(attribute) is not None for span in direct_spans
    )
    # In restyle/mixed, inherited heading/body style is target-owned.  Only a
    # direct run-level emphasis marker makes the source value a protected
    # inline contract.  Legacy contexts without this evidence retain their
    # previous run-level behaviour.
    if not preserve_source and not has_direct_evidence:
        return None

    expected_end = max(int(span.get("end", 0)) for span in expected_spans)
    actual_end = max(int(span.get("end", 0)) for span in actual_spans)
    if expected_end != actual_end:
        return [{
            "start": 0,
            "end": max(expected_end, actual_end),
            "expected": {"logical_text_length": expected_end},
            "actual": {"logical_text_length": actual_end},
            "ok": False,
        }]

    boundaries = {0, expected_end}
    for span in expected_spans + actual_spans:
        boundaries.add(int(span.get("start", 0)))
        boundaries.add(int(span.get("end", 0)))
    ordered = sorted(value for value in boundaries if 0 <= value <= expected_end)
    comparisons = []
    for start, end in zip(ordered, ordered[1:]):
        if start == end:
            continue
        expected_span = next(
            (span for span in expected_spans
             if int(span.get("start", 0)) <= start and end <= int(span.get("end", 0))),
            None,
        )
        actual_span = next(
            (span for span in actual_spans
             if int(span.get("start", 0)) <= start and end <= int(span.get("end", 0))),
            None,
        )
        expected_value = _span_style_value(expected_span or {}, attribute)
        actual_value = _span_style_value(actual_span or {}, attribute)
        direct_span = next(
            (span for span in direct_spans
             if int(span.get("start", 0)) <= start
             and end <= int(span.get("end", 0))
             and span.get(attribute) is not None),
            None,
        )
        enforce_source = preserve_source or direct_span is not None
        if not enforce_source and target_value is not None:
            expected_value = target_value
        comparisons.append({
            "start": start,
            "end": end,
            "expected": expected_value,
            "actual": actual_value,
            "source_protected": enforce_source,
            "ok": (
                expected_span is not None
                and actual_span is not None
                and _inline_value_equal(expected_value, actual_value, attribute)
            ),
        })
    return comparisons


def normalize_text(text: Optional[str]) -> str:
    """规范化文本：去除所有空白字符"""
    if not text:
        return ""
    return re.sub(r"\s+", "", text)


def _visible_text_without_field_cache(element: Any) -> str:
    """提取作者正文，排除 fldSimple/复杂字段的受控缓存值。"""
    text_parts: List[str] = []
    complex_field_depth = 0
    simple_fields = {node for node in element.iter(qn("w:fldSimple"))}
    for node in element.iter():
        if node.tag == qn("w:fldChar"):
            field_type = (node.get(qn("w:fldCharType")) or "").lower()
            if field_type == "begin":
                complex_field_depth += 1
            elif field_type == "end":
                complex_field_depth = max(0, complex_field_depth - 1)
            continue
        if node.tag != qn("w:t"):
            continue
        if complex_field_depth:
            continue
        if any(parent in simple_fields for parent in node.iterancestors()):
            continue
        text_parts.append(node.text or "")
    return normalize_text("".join(text_parts))


def compute_file_sha256(path: Union[str, Path]) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def extract_media_hashes_from_docx(doc_path: Path) -> Set[str]:
    """从 DOCX 包内提取所有嵌入图片部件的 SHA-256 哈希。

    NotesMerger 等关系克隆器可能把脚注/尾注中的图片落在
    ``word/synthNote*.png``，而不是传统的 ``word/media/`` 目录；完整性
    清单必须按内容类型/扩展名覆盖这些合法的图片部件。
    """
    hashes = set()
    if not doc_path.is_file():
        return hashes
    image_suffixes = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf"}
    try:
        with zipfile.ZipFile(doc_path, "r") as z:
            for info in z.infolist():
                if (
                    info.filename.startswith("word/")
                    and Path(info.filename).suffix.lower() in image_suffixes
                    and not info.is_dir()
                ):
                    data = z.read(info.filename)
                    if data:
                        hashes.add(hashlib.sha256(data).hexdigest())
    except Exception:
        pass
    return hashes


def _related_part(doc: Document, relationship_type: str):
    """返回主文档关联的部件；没有关联时返回 None。"""
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        return doc.part.part_related_by(getattr(RT, relationship_type))
    except (KeyError, AttributeError):
        return None


def _extract_note_definitions(doc: Document, relationship_type: str, note_tag: str) -> Dict[str, str]:
    part = _related_part(doc, relationship_type)
    if part is None:
        return {}
    try:
        root = lxml.etree.fromstring(part.blob)
    except (AttributeError, lxml.etree.XMLSyntaxError):
        return {}
    definitions = {}
    for note in root.findall(qn(note_tag)):
        note_id = note.get(qn("w:id"))
        if note_id in (None, "-1", "0"):
            continue
        text = normalize_text("".join(item.text or "" for item in note.iter(qn("w:t"))))
        definitions[note_id] = text
    return definitions


def _extract_media_instances(doc: Document, body_elements: Iterable[Any], package_media: Set[str]) -> List[str]:
    """解析每个 drawing 的真实媒体目标，而不是仅列出 ZIP 中的资源集合。"""
    instances: List[str] = []
    for element in body_elements:
        blips = list(element.iter(MEDIA_BLIP_TAG)) if hasattr(element, "iter") else []
        for blip in blips:
            rid = blip.get(qn("r:embed"))
            rel = getattr(doc.part, "rels", {}).get(rid) if rid else None
            target = getattr(rel, "target_part", None) if rel is not None else None
            blob = getattr(target, "blob", None) if target is not None else None
            if blob:
                instances.append(hashlib.sha256(blob).hexdigest())
    # 未被 body drawing 直接引用的媒体仍应有可比证据，但不制造重复实例。
    if not instances:
        instances.extend(sorted(package_media))
    return instances


def _extract_media_occurrences(
    doc: Document,
    body_elements: Iterable[Any],
    source_sha256: str = "",
) -> List[ObjectOccurrence]:
    occurrences: List[ObjectOccurrence] = []
    host_ordinals: Dict[Tuple[str, str], int] = {}
    for element_index, element in enumerate(body_elements, start=1):
        local_name = str(element.tag).rsplit("}", 1)[-1]
        host_element = f"/w:document/w:body/w:{local_name}[{element_index}]"
        host_text = normalize_text("".join(node.text or "" for node in element.iter(qn("w:t"))))
        host_key = (local_name, host_text)
        host_ordinals[host_key] = host_ordinals.get(host_key, 0) + 1
        for local_index, blip in enumerate(element.iter(MEDIA_BLIP_TAG), start=1):
            rid = blip.get(qn("r:embed"))
            rel = getattr(doc.part, "rels", {}).get(rid) if rid else None
            target = getattr(rel, "target_part", None) if rel is not None else None
            blob = getattr(target, "blob", None) if target is not None else None
            if not blob:
                continue
            occurrences.append(ObjectOccurrence(
                kind="image",
                fingerprint=hashlib.sha256(blob).hexdigest(),
                part_uri="word/document.xml",
                host_element=host_element,
                local_index=local_index,
                host_text=host_text,
                host_ordinal=host_ordinals[host_key],
                relationship_id=rid,
                target_part=str(getattr(target, "partname", "")) or None,
                story_type="body",
                source_sha256=source_sha256,
            ))
    return occurrences


def _extract_story_media_occurrences(doc: Document, source_sha256: str = "") -> List[ObjectOccurrence]:
    """Extract image occurrences from header/footer/notes story parts."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    rel_story_types = {
        RT.HEADER: "header",
        RT.FOOTER: "footer",
        RT.FOOTNOTES: "footnote",
        RT.ENDNOTES: "endnote",
    }
    occurrences: List[ObjectOccurrence] = []
    for story_rel in doc.part.rels.values():
        story_type = rel_story_types.get(story_rel.reltype)
        if not story_type or story_rel.is_external:
            continue
        story_part = getattr(story_rel, "target_part", None)
        if story_part is None:
            continue
        try:
            root = lxml.etree.fromstring(story_part.blob)
        except (AttributeError, lxml.etree.XMLSyntaxError):
            continue
        part_uri = str(getattr(story_part, "partname", "")).lstrip("/")
        host_elements: List[Tuple[Any, str, Optional[int]]] = []
        if story_type in {"footnote", "endnote"}:
            root_tag = "footnote" if story_type == "footnote" else "endnote"
            for note in root.findall(qn(f"w:{root_tag}")):
                note_id = note.get(qn("w:id"), "")
                for index, child in enumerate(note, start=1):
                    if child.tag in {qn("w:p"), qn("w:tbl")}:
                        host_elements.append((child, f"/w:{root_tag}[@w:id='{note_id}']/{child.tag.rsplit('}', 1)[-1]}[{index}]", None))
        else:
            root_name = "hdr" if story_type == "header" else "ftr"
            counters: Dict[str, int] = {}
            host_ordinals: Dict[Tuple[str, str], int] = {}
            for child in root:
                local_name = child.tag.rsplit("}", 1)[-1]
                if child.tag not in {qn("w:p"), qn("w:tbl")}:
                    continue
                counters[local_name] = counters.get(local_name, 0) + 1
                host_text = normalize_text("".join(node.text or "" for node in child.iter(qn("w:t"))))
                host_key = (local_name, host_text)
                host_ordinals[host_key] = host_ordinals.get(host_key, 0) + 1
                host_elements.append((child, f"/w:{root_name}/w:{local_name}[{counters[local_name]}]", host_ordinals[host_key]))
        for element, host_element, host_ordinal in host_elements:
            for local_index, blip in enumerate(element.iter(MEDIA_BLIP_TAG), start=1):
                rid = blip.get(qn("r:embed"))
                rel = getattr(story_part, "rels", {}).get(rid) if rid else None
                target = getattr(rel, "target_part", None) if rel is not None else None
                blob = getattr(target, "blob", None) if target is not None else None
                if not blob:
                    continue
                host_text = normalize_text("".join(node.text or "" for node in element.iter(qn("w:t"))))
                occurrences.append(ObjectOccurrence(
                    kind="image",
                    fingerprint=hashlib.sha256(blob).hexdigest(),
                    part_uri=part_uri,
                    host_element=host_element,
                    local_index=local_index,
                    host_text=host_text,
                    host_ordinal=host_ordinal,
                    relationship_id=rid,
                    target_part=str(getattr(target, "partname", "")) or None,
                    story_type=story_type,
                    source_sha256=source_sha256,
                ))
    return occurrences


def _resolve_hyperlink_target(doc: Document, hyperlink: Any) -> str:
    anchor = hyperlink.get(qn("w:anchor"))
    if anchor:
        return f"anchor:{anchor}"
    rid = hyperlink.get(qn("r:id"))
    rel = getattr(doc.part, "rels", {}).get(rid) if rid else None
    target = getattr(rel, "target_ref", None) if rel is not None else None
    if target:
        return f"external:{target}"
    return f"relationship:{rid or ''}"


def _math_formula_fingerprint(element: Any) -> str:
    """Hash OMML content without transient namespace declarations.

    Word commonly rewrites the package-level namespace map when it saves a
    document.  The formula tree itself is unchanged, so unused in-scope
    declarations must not make an otherwise preserved formula look different.
    """
    normalized = copy.deepcopy(element)
    lxml.etree.cleanup_namespaces(normalized)
    return hashlib.sha256(lxml.etree.tostring(normalized, method="c14n")).hexdigest()


def _normalize_field_instruction(value: Optional[str]) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())


def _field_identity_key(instruction: str) -> str:
    """Return a comparison key that tolerates collision-safe bookmark remaps.

    REF/PAGEREF/STYLEREF targets are local names inside one DOCX package.  A
    directory-tree import may have to rename a source bookmark when another
    source uses the same name; the command and its switches still carry the
    same semantics.  Other fields remain exact instruction matches.
    """
    normalized = _normalize_field_instruction(instruction)
    match = re.match(
        r"(?i)^(REF|PAGEREF|STYLEREF)\s+([^\\\s]+)(.*)$",
        normalized,
    )
    if not match:
        return normalized
    command, _target, suffix = match.groups()
    return f"{command.upper()} <bookmark>{suffix}"


def _field_dependencies(fields: Iterable[str]) -> List[str]:
    dependencies = []
    for instruction in fields:
        match = re.match(r"(?i)\s*(REF|PAGEREF|STYLEREF|SEQ|TOC|HYPERLINK)\s+([^\\\s]+)", instruction)
        if match:
            dependencies.append(f"{match.group(1).upper()}:{match.group(2)}")
    return dependencies


def _content_scope_elements(doc: Document, content_parts: Iterable[str]) -> List[Any]:
    """按交付部件边界书签返回 body 顶层元素，排除封面/目录等生成内容。"""
    from .document_parts import get_part_boundary_bookmark

    body = doc.element.body
    children = list(body)
    intervals = []
    for part_id in content_parts:
        start_name = get_part_boundary_bookmark(part_id, is_start=True)
        end_name = get_part_boundary_bookmark(part_id, is_start=False)
        start_index = next(
            (index for index, child in enumerate(children)
             if any(item.get(qn("w:name")) == start_name for item in child.iter(qn("w:bookmarkStart")))),
            None,
        )
        if start_index is None:
            raise ContentIntegrityError(f"交付物缺失内容部件边界书签: {part_id}/{start_name}")
        end_index = next(
            (index for index, child in enumerate(children)
             if any(item.get(qn("w:name")) == end_name for item in child.iter(qn("w:bookmarkStart")))),
            None,
        )
        if end_index is None:
            raise ContentIntegrityError(f"交付物缺失内容部件结束边界书签: {part_id}/{end_name}")
        if end_index < start_index:
            raise ContentIntegrityError(f"交付物内容部件边界顺序非法: {part_id}")
        intervals.append((start_index, end_index))
    selected = []
    for start, end in sorted(intervals):
        selected.extend(children[start:end + 1])
    return selected


def _declared_part_ranges(
    doc: Document,
    part_ids: Iterable[str],
) -> Dict[str, Tuple[int, int]]:
    """Return body-child intervals for the declared output parts."""
    children = list(doc.element.body)
    ranges: Dict[str, Tuple[int, int]] = {}
    for part_id in part_ids:
        start_name = get_part_boundary_bookmark(part_id, is_start=True)
        end_name = get_part_boundary_bookmark(part_id, is_start=False)
        start = next(
            (
                index for index, child in enumerate(children)
                if any(marker.get(qn("w:name")) == start_name for marker in child.iter(qn("w:bookmarkStart")))
            ),
            None,
        )
        end = next(
            (
                index for index, child in enumerate(children)
                if any(marker.get(qn("w:name")) == end_name for marker in child.iter(qn("w:bookmarkStart")))
            ),
            None,
        )
        if start is not None and end is not None and start <= end:
            ranges[str(part_id)] = (start, end)
    return ranges


def _story_part_ids_by_output_part(
    doc: Document,
    ranges: Dict[str, Tuple[int, int]],
) -> Dict[str, Set[str]]:
    """Associate effective header/footer story parts with output parts."""
    children = list(doc.element.body)
    section_positions: Dict[int, int] = {}
    for section_index, section in enumerate(getattr(doc, "sections", ())):
        for child_index, child in enumerate(children):
            if any(item is section._sectPr for item in child.iter(qn("w:sectPr"))):
                section_positions[section_index] = child_index
                break

    section_parts: Dict[int, str] = {}
    for section_index in range(len(getattr(doc, "sections", ()))):
        marker_index = section_positions.get(section_index)
        if marker_index is None:
            continue
        containing = next(
            (part_id for part_id, (start, end) in ranges.items() if start <= marker_index <= end),
            None,
        )
        if containing is None:
            preceding = [
                (end, part_id) for part_id, (start, end) in ranges.items()
                if end < marker_index
            ]
            if preceding:
                containing = max(preceding)[1]
        if containing is not None:
            section_parts[section_index] = containing

    story_to_parts: Dict[str, Set[str]] = collections.defaultdict(set)
    previous: Dict[Tuple[str, str], str] = {}
    rel_story_types = {
        "header": qn("w:headerReference"),
        "footer": qn("w:footerReference"),
    }
    for section_index, section in enumerate(getattr(doc, "sections", ())):
        current: Dict[Tuple[str, str], str] = {}
        for story_type, tag in rel_story_types.items():
            explicit: Dict[str, str] = {}
            for reference in section._sectPr.findall(tag):
                variant = reference.get(qn("w:type")) or "default"
                rid = reference.get(qn("r:id"))
                rel = getattr(doc.part, "rels", {}).get(rid) if rid else None
                target = getattr(rel, "target_part", None) if rel is not None else None
                part_uri = str(getattr(target, "partname", "")).lstrip("/") if target is not None else ""
                if part_uri:
                    explicit[variant] = part_uri
            variants = set(explicit) | {
                variant for (bound_story, variant) in previous
                if bound_story == story_type
            }
            for variant in variants:
                part_uri = explicit.get(variant) or previous.get((story_type, variant))
                if not part_uri:
                    continue
                current[(story_type, variant)] = part_uri
                output_part = section_parts.get(section_index)
                if output_part:
                    story_to_parts[part_uri].add(output_part)
        previous = current
    return story_to_parts


def _occurrence_output_parts(
    doc: Document,
    occurrences: Iterable[ObjectOccurrence],
    part_ids: Iterable[str],
) -> Dict[int, Set[str]]:
    """Map each full-package occurrence identity to declared output parts."""
    part_ids = tuple(str(item) for item in part_ids)
    ranges = _declared_part_ranges(doc, part_ids)
    story_to_parts = _story_part_ids_by_output_part(doc, ranges)
    occurrence_parts: Dict[int, Set[str]] = {}
    for occurrence in occurrences:
        parts: Set[str] = set()
        if occurrence.story_type == "body":
            match = re.search(r"\[(\d+)\]$", occurrence.host_element)
            if match:
                child_index = int(match.group(1)) - 1
                parts.update(
                    part_id for part_id, (start, end) in ranges.items()
                    if start <= child_index <= end
                )
        else:
            parts.update(story_to_parts.get(occurrence.part_uri, set()))
            if occurrence.story_type in {"footnote", "endnote"} and not parts:
                parts.update(
                    part_id for part_id in part_ids
                    if part_id == "body" or part_id not in {"cover", "toc"}
                )
        occurrence_parts[id(occurrence)] = parts
    return occurrence_parts


def merge_semantic_inventories(inventories: Iterable[SemanticInventory]) -> SemanticInventory:
    """按来源顺序合并清单，所有可重复对象均以多重集语义保留。"""
    inventories = list(inventories)
    paragraphs = [item for inv in inventories for item in inv.paragraphs]
    tables = [item for inv in inventories for item in inv.tables]
    media_instances = [item for inv in inventories for item in inv.media_instances]
    media_occurrences = [item for inv in inventories for item in inv.media_occurrences]
    if not media_instances:
        media_instances = [item for inv in inventories for item in sorted(inv.media_hashes)]
    fields = [item for inv in inventories for item in inv.fields]
    footnotes = [
        f"{index}:{item}"
        for index, inv in enumerate(inventories)
        for item in inv.footnotes
    ]
    endnotes = [
        f"{index}:{item}"
        for index, inv in enumerate(inventories)
        for item in inv.endnotes
    ]
    formulas = [item for inv in inventories for item in inv.math_formulas]
    hyperlinks = [item for inv in inventories for item in inv.hyperlinks]
    bookmarks = set().union(*(inv.bookmarks for inv in inventories)) if inventories else set()
    footnote_definitions = {
        f"{index}:{note_id}": text
        for index, inv in enumerate(inventories)
        for note_id, text in inv.footnote_definitions.items()
    }
    endnote_definitions = {
        f"{index}:{note_id}": text
        for index, inv in enumerate(inventories)
        for note_id, text in inv.endnote_definitions.items()
    }
    source_sha = hashlib.sha256(
        "|".join(inv.source_sha256 for inv in inventories).encode("utf-8")
    ).hexdigest() if inventories else ""
    return SemanticInventory(
        source_sha256=source_sha,
        paragraphs=paragraphs,
        paragraph_counts=dict(collections.Counter(paragraphs)),
        tables=tables,
        media_hashes=set(item for inv in inventories for item in inv.media_hashes),
        math_formulas=formulas,
        footnotes=footnotes,
        fields=fields,
        bookmarks=bookmarks,
        hyperlinks=hyperlinks,
        media_instances=media_instances,
        media_occurrences=media_occurrences,
        footnote_definitions=footnote_definitions,
        endnotes=endnotes,
        endnote_definitions=endnote_definitions,
        field_dependencies=[item for inv in inventories for item in inv.field_dependencies],
    )


def extract_semantic_inventory(
    doc_or_path: Union[str, Path, Document],
    content_parts: Optional[Iterable[str]] = None,
) -> SemanticInventory:
    """
    深层提取文档的语义资产清单（可见文字多重集、表格拓扑、媒体、数学公式、脚注、字段等）
    """
    if isinstance(doc_or_path, (str, Path)):
        file_path = Path(doc_or_path).resolve()
        sha256 = compute_file_sha256(file_path)
        doc = Document(str(file_path))
        media_hashes = extract_media_hashes_from_docx(file_path)
    else:
        doc = doc_or_path
        sha256 = ""
        media_hashes = set()
        try:
            for part in doc.part.package.parts:
                if "media" in str(part.partname):
                    blob = getattr(part, "blob", None)
                    if blob:
                        media_hashes.add(hashlib.sha256(blob).hexdigest())
        except Exception:
            pass

    body_element = doc.element.body
    # 普通清单覆盖整个 body；ExpectedInventory 可将实际检查范围限制在内容部件边界。
    scope_elements = (
        _content_scope_elements(doc, content_parts)
        if content_parts else list(body_element)
    )

    # 1. 段落提取（排除生成的动态目录条目）
    paragraphs = []
    for p in scope_elements:
        if p.tag != qn("w:p"):
            continue
        bms = [b.get(qn("w:name"), "") for b in p.xpath(".//w:bookmarkStart")]
        if any(b.startswith("_Synth_entry_") for b in bms):
            continue

        norm = _visible_text_without_field_cache(p)
        if norm:
            paragraphs.append(norm)

    paragraph_counts = dict(collections.Counter(paragraphs))

    # 2. 表格拓扑提取
    tables = []
    for tbl in scope_elements:
        if tbl.tag != qn("w:tbl"):
            continue
        rows = tbl.xpath("./w:tr")
        row_count = len(rows)
        col_count = 0
        grid_cols = tbl.xpath("./w:tblGrid/w:gridCol")
        if grid_cols:
            col_count = len(grid_cols)
        elif row_count > 0:
            col_count = len(rows[0].xpath("./w:tc"))

        cell_texts = []
        for r in rows:
            for tc in r.xpath("./w:tc"):
                cell_texts.append(_visible_text_without_field_cache(tc))

        grid_widths = [int(gc.get(qn("w:w"), 0)) for gc in grid_cols]
        tables.append(TableTopology(
            row_count=row_count,
            col_count=col_count,
            cell_texts=cell_texts,
            grid_widths=grid_widths,
        ))

    # 3. 数学公式提取 (OMML)。每个 oMath 是一个实例。
    math_formulas = []
    for element in scope_elements:
        for m_el in element.iter(qn("m:oMath")):
            math_formulas.append(_math_formula_fingerprint(m_el))

    # 4. 脚注/尾注引用与定义
    footnotes = []
    for element in scope_elements:
        footnotes.extend(
            fn.get(qn("w:id"), "") for fn in element.iter(qn("w:footnoteReference"))
        )
    endnotes = []
    for element in scope_elements:
        endnotes.extend(
            en.get(qn("w:id"), "") for en in element.iter(qn("w:endnoteReference"))
        )
    footnote_definitions = _extract_note_definitions(doc, "FOOTNOTES", "w:footnote")
    endnote_definitions = _extract_note_definitions(doc, "ENDNOTES", "w:endnote")

    # 5. 字段指令及依赖。缓存文本不属于字段指令，合法 SEQ 值更新不会误报。
    fields = []
    for element in scope_elements:
        for simple in element.iter(qn("w:fldSimple")):
            instruction = _normalize_field_instruction(simple.get(qn("w:instr")))
            if instruction:
                fields.append(instruction)
        for instr in element.iter(qn("w:instrText")):
            if instr.text and instr.text.strip():
                instruction = _normalize_field_instruction(instr.text)
                if instruction:
                    fields.append(instruction)

    bookmarks = {
        marker.get(qn("w:name"))
        for element in scope_elements for marker in element.iter(qn("w:bookmarkStart"))
    }
    hyperlinks = [
        _resolve_hyperlink_target(doc, link)
        for element in scope_elements for link in element.iter(qn("w:hyperlink"))
    ]
    media_occurrences = _extract_media_occurrences(doc, scope_elements, sha256)
    media_occurrences.extend(_extract_story_media_occurrences(doc, sha256))
    body_occurrences = [item for item in media_occurrences if item.story_type == "body"]
    story_occurrences = [item for item in media_occurrences if item.story_type != "body"]
    if body_occurrences:
        media_instances = [item.fingerprint for item in body_occurrences]
        media_instances.extend(item.fingerprint for item in story_occurrences)
    else:
        media_instances = sorted(media_hashes)

    return SemanticInventory(
        source_sha256=sha256,
        paragraphs=paragraphs,
        paragraph_counts=paragraph_counts,
        tables=tables,
        media_hashes=media_hashes,
        math_formulas=math_formulas,
        footnotes=footnotes,
        fields=fields,
        bookmarks=bookmarks,
        hyperlinks=hyperlinks,
        media_instances=media_instances,
        media_occurrences=media_occurrences,
        footnote_definitions=footnote_definitions,
        endnotes=endnotes,
        endnote_definitions=endnote_definitions,
        field_dependencies=_field_dependencies(fields),
    )


def verify_content_integrity(
    source_inventory: Union[SemanticInventory, ExpectedInventory],
    delivery_doc_or_path: Union[str, Path, Document],
    replacements: Optional[Dict[str, str]] = None,
    is_body_only: bool = False,
    evidence_out: Optional[Dict[str, Any]] = None,
) -> bool:
    """
    严格校验交付物的内容语义完整性：
    1. 段落多重集：确保源文档中的每一个段落文本出现频次在交付物中均被满足（>= 源频次）；
       对于声明的标题替换映射（如 raw_text -> title），自动折算对应频次；
    2. 表格完整性：源表格拓扑与单元格文字在交付物中无损保留；
    3. 媒体哈希：源文档中所有嵌入图片哈希在交付物中完整存在；
    4. 数学公式：源文档中的 OMML 数学公式无损保留。
    """
    expected = source_inventory.semantic if isinstance(source_inventory, ExpectedInventory) else source_inventory
    if isinstance(source_inventory, ExpectedInventory):
        replacements = {**source_inventory.replacements, **(replacements or {})}
    delivery_inv = extract_semantic_inventory(
        delivery_doc_or_path,
        content_parts=source_inventory.content_parts if isinstance(source_inventory, ExpectedInventory) else None,
    )
    delivery_document = (
        Document(str(delivery_doc_or_path))
        if isinstance(delivery_doc_or_path, (str, Path))
        else delivery_doc_or_path
    )
    # Text and tables must stay within the declared content boundary, while
    # object licenses can target a generated cover outside that boundary.
    # Keep a full-package inventory for the instance/anchor gate only.
    object_delivery_inv = (
        extract_semantic_inventory(delivery_doc_or_path)
        if isinstance(source_inventory, ExpectedInventory)
        else delivery_inv
    )
    rep_map = {}
    if replacements:
        for k, v in replacements.items():
            norm_k = normalize_text(k)
            norm_v = normalize_text(v)
            if norm_k and norm_v:
                rep_map[norm_k] = norm_v

    # 1. 校验段落文字序列/多重集。ExpectedInventory 的分母来自声明的
    # content 部件，因此允许按部件维护严格有序序列；旧 SemanticInventory
    # 入口保留下限兼容性，避免把历史的 cover/toc+body 调用误判为正文错误。
    delivery_counts = dict(delivery_inv.paragraph_counts)

    missing_paragraphs = []
    for text, required_count in expected.paragraph_counts.items():
        if len(text) <= 1:
            continue

        target_text = rep_map.get(text, text)
        actual_count = delivery_counts.get(target_text, 0)

        if actual_count < required_count:
            missing_paragraphs.append({
                "text": text[:60] + ("..." if len(text) > 60 else ""),
                "required_count": required_count,
                "actual_count": actual_count,
                "target_text": target_text[:60] + ("..." if len(target_text) > 60 else ""),
            })

    if missing_paragraphs:
        err_details = "\n".join(
            f"  - [{p['required_count']} != {p['actual_count']}] '{p['text']}' -> '{p['target_text']}'"
            for p in missing_paragraphs[:5]
        )
        if len(missing_paragraphs) > 5:
            err_details += f"\n  ... 另有 {len(missing_paragraphs) - 5} 项缺失"
        raise ContentIntegrityError(f"正文段落文字完整性核验未通过，发现丢失或频次不足段落:\n{err_details}")

    if isinstance(source_inventory, ExpectedInventory):
        expected_sequence = [
            rep_map.get(normalize_text(text), normalize_text(text))
            for text in expected.paragraphs
            if len(normalize_text(text)) > 1
        ]
        actual_sequence = [normalize_text(text) for text in delivery_inv.paragraphs if len(normalize_text(text)) > 1]
        generated = collections.Counter(
            normalize_text(text) for text, count in source_inventory.allowed_generated_text.items()
            for _ in range(max(0, int(count)))
        )
        cursor = 0
        unexpected = []
        for actual in actual_sequence:
            if cursor < len(expected_sequence) and actual == expected_sequence[cursor]:
                cursor += 1
            elif generated.get(actual, 0):
                generated[actual] -= 1
            else:
                unexpected.append(actual)
        if cursor != len(expected_sequence):
            missing = expected_sequence[cursor:cursor + 3]
            raise ContentIntegrityError(
                "正文来源实例顺序或次数不一致: 未按声明顺序找到预期节点 "
                + repr(missing)
            )
        if unexpected:
            raise ContentIntegrityError(
                "正文包含未声明的额外来源实例或生成内容: "
                + repr(unexpected[:3])
            )

    # 2. 校验表格完整性
    if len(delivery_inv.tables) < len(expected.tables):
        raise ContentIntegrityError(
            f"表格数量不一致: 源文档包含 {len(expected.tables)} 个表格，交付物仅含 {len(delivery_inv.tables)} 个"
        )

    def table_key(table: TableTopology) -> Tuple[Any, ...]:
        return (
            table.row_count,
            table.col_count,
            tuple(table.grid_widths),
            tuple(table.cell_texts),
        )

    required_tables = collections.Counter(table_key(table) for table in expected.tables)
    actual_tables = collections.Counter(table_key(table) for table in delivery_inv.tables)
    missing_tables = required_tables - actual_tables
    if missing_tables:
        missing_key, missing_count = next(iter(missing_tables.items()))
        raise ContentIntegrityError(
            f"表格拓扑/单元格位置或重复数量不足: 缺少 {missing_count} 个结构 "
            f"{missing_key[:2]}，单元格内容 {missing_key[3][:3]}"
        )

    # 3. 校验媒体资源的真实实例和引用目标
    required_media = collections.Counter(expected.media_instances or sorted(expected.media_hashes))
    actual_media = collections.Counter(object_delivery_inv.media_instances or sorted(object_delivery_inv.media_hashes))
    missing_media = required_media - actual_media
    missing_hashes = expected.media_hashes - object_delivery_inv.media_hashes
    allowed_media = collections.Counter()
    if isinstance(source_inventory, ExpectedInventory):
        allowed_media.update(
            item
            for item, count in source_inventory.allowed_generated_media.items()
            for _ in range(max(0, int(count)))
        )
    extra_media = actual_media - required_media - allowed_media
    # A fingerprint allowance is only a preliminary resource-level permit.
    # Defer its over-count to the occurrence mapper so the final diagnostic
    # names the missing/extra generated instance rather than stopping at a
    # hash-only count gate.
    if isinstance(source_inventory, ExpectedInventory) and extra_media:
        deferred = collections.Counter({
            fingerprint: count
            for fingerprint, count in extra_media.items()
            if fingerprint in source_inventory.allowed_generated_media
        })
        extra_media -= deferred
    if missing_media or missing_hashes or extra_media:
        missing_total = sum(missing_media.values()) + len(missing_hashes)
        extra_total = sum(extra_media.values())
        raise ContentIntegrityError(
            f"交付物嵌入媒体实例不匹配: missing={missing_total}, extra={extra_total}; "
            f"missing={list((missing_media or collections.Counter()).elements())[:3] or list(missing_hashes)[:3]}, "
            f"extra={list(extra_media.elements())[:3]}"
        )

    generated_allowances = (
        source_inventory.allowed_generated_objects
        if isinstance(source_inventory, ExpectedInventory)
        else ()
    )

    # A direct source/delivery comparison can additionally enforce the host
    # occurrence.  Prepared multi-source deliveries may legitimately remap
    # body paths during assembly; those use the instance-count gate above
    # until an explicit source-part/output mapping is available.
    if not isinstance(source_inventory, ExpectedInventory):
        def occurrence_key(item: ObjectOccurrence) -> Tuple[Any, ...]:
            return (
                item.kind,
                item.fingerprint,
                item.part_uri,
                item.host_element,
                item.local_index,
            )

        expected_occurrences = collections.Counter(
            occurrence_key(item) for item in expected.media_occurrences
        )
        actual_occurrences = collections.Counter(
            occurrence_key(item) for item in delivery_inv.media_occurrences
        )
        missing_occurrences = expected_occurrences - actual_occurrences
        extra_occurrences = actual_occurrences - expected_occurrences
        if missing_occurrences or extra_occurrences:
            raise ContentIntegrityError(
                "交付物非文字对象出现位置不匹配: "
                f"missing={list(missing_occurrences.elements())[:3]}, "
                f"extra={list(extra_occurrences.elements())[:3]}"
            )
    elif expected.media_occurrences or generated_allowances:
        # Consume actual instances one by one.  This is intentionally a
        # mapping, not a Counter comparison: every source occurrence receives
        # one output anchor, and only the explicitly declared generated
        # allowance may consume a remaining instance.
        full_occurrences = list(object_delivery_inv.media_occurrences)
        source_candidates = (
            list(delivery_inv.media_occurrences)
            if isinstance(source_inventory, ExpectedInventory) and source_inventory.content_parts
            else full_occurrences
        )
        remaining = list(full_occurrences)
        declared_part_ids = tuple(
            str(item) for item in (source_inventory.provenance.get("parts", ()) or ())
        )
        occurrence_parts = _occurrence_output_parts(
            delivery_document,
            full_occurrences,
            declared_part_ids,
        )
        object_mapping = []
        missing_objects = []
        source_used = set()
        source_cursor: Dict[str, int] = {}
        strict_host_ordinal = (
            not isinstance(source_inventory, ExpectedInventory)
            or len(source_inventory.source_order) <= 1
        )

        def matches(expected_item: ObjectOccurrence, actual_item: ObjectOccurrence) -> bool:
            if expected_item.kind != actual_item.kind:
                return False
            if expected_item.fingerprint != actual_item.fingerprint:
                return False
            if expected_item.story_type != actual_item.story_type:
                return False
            if expected_item.local_index != actual_item.local_index:
                return False
            if strict_host_ordinal and (
                expected_item.host_ordinal is not None
                and actual_item.host_ordinal is not None
                and expected_item.host_ordinal != actual_item.host_ordinal
            ):
                return False
            return not expected_item.host_text or expected_item.host_text == actual_item.host_text

        def remove_from_full(
            actual_item: ObjectOccurrence,
            preferred_parts: Optional[Set[str]] = None,
        ) -> Optional[ObjectOccurrence]:
            exact_candidates = [
                index for index, item in enumerate(remaining) if item == actual_item
            ]
            if preferred_parts:
                preferred_exact = [
                    index for index in exact_candidates
                    if occurrence_parts.get(id(remaining[index]), set()) & preferred_parts
                ]
                exact_candidates = preferred_exact
            exact_index = exact_candidates[0] if exact_candidates else None
            if exact_index is not None:
                return remaining.pop(exact_index)
            candidates = [
                index for index, item in enumerate(remaining)
                if item.kind == actual_item.kind
                and item.fingerprint == actual_item.fingerprint
                and item.story_type == actual_item.story_type
                and item.local_index == actual_item.local_index
                and item.host_text == actual_item.host_text
            ]
            if not candidates:
                return None
            if preferred_parts:
                part_candidates = [
                    index for index in candidates
                    if occurrence_parts.get(id(remaining[index]), set()) & preferred_parts
                ]
                if part_candidates:
                    candidates = part_candidates
            same_ordinal = next(
                (index for index in candidates
                 if remaining[index].host_ordinal == actual_item.host_ordinal),
                candidates[0],
            )
            return remaining.pop(same_ordinal)

        for expected_item in expected.media_occurrences:
            story_type = expected_item.story_type
            lower_bound = source_cursor.get(story_type, -1)
            match_index = next(
                (
                    index for index, actual_item in enumerate(source_candidates)
                    if index > lower_bound
                    and index not in source_used
                    and matches(expected_item, actual_item)
                ),
                None,
            )
            if match_index is None:
                missing_objects.append(expected_item.to_dict())
                continue
            source_used.add(match_index)
            source_cursor[story_type] = match_index
            actual_item = source_candidates[match_index]
            full_item = remove_from_full(
                actual_item,
                preferred_parts=set(source_inventory.content_parts),
            )
            object_mapping.append({
                "kind": "source_object",
                "source_sha256": expected_item.source_sha256,
                "source_part_uri": expected_item.part_uri,
                "source_host_element": expected_item.host_element,
                "output_part_uri": actual_item.part_uri,
                "output_host_element": actual_item.host_element,
                "story_type": actual_item.story_type,
                "fingerprint": actual_item.fingerprint,
                "local_index": actual_item.local_index,
                "source_host_ordinal": expected_item.host_ordinal,
                "output_host_ordinal": actual_item.host_ordinal,
                "output_part_ids": sorted(occurrence_parts.get(id(full_item), set())),
            })

        missing_generated = []
        for allowance in generated_allowances:
            expected_count = max(1, int(allowance.get("count", 1)))
            for _ in range(expected_count):
                match_index = next(
                    (
                        index for index, actual_item in enumerate(remaining)
                        if actual_item.kind == allowance.get("kind", "image")
                        and actual_item.fingerprint == allowance.get("fingerprint")
                        and actual_item.story_type == allowance.get("story_type", "body")
                        and (
                            allowance.get("host_text") in (None, "")
                            or actual_item.host_text == allowance.get("host_text")
                        )
                        and (
                            allowance.get("local_index") in (None, "")
                            or actual_item.local_index == int(allowance["local_index"])
                        )
                        and (
                            not allowance.get("part_id")
                            or allowance.get("part_id") in occurrence_parts.get(id(actual_item), set())
                        )
                    ),
                    None,
                )
                if match_index is None:
                    missing_generated.append(dict(allowance))
                    continue
                actual_item = remaining.pop(match_index)
                object_mapping.append({
                    "kind": "generated_object",
                    "part_id": allowance.get("part_id", "body"),
                    "rule": allowance.get("rule", "explicit"),
                    "output_part_uri": actual_item.part_uri,
                    "output_host_element": actual_item.host_element,
                    "story_type": actual_item.story_type,
                    "fingerprint": actual_item.fingerprint,
                    "local_index": actual_item.local_index,
                    "output_part_ids": sorted(occurrence_parts.get(id(actual_item), set())),
                })

        if evidence_out is not None:
            evidence_out["object_mapping"] = object_mapping
            evidence_out["missing_source_objects"] = missing_objects
            evidence_out["missing_generated_objects"] = missing_generated
            evidence_out["extra_objects"] = [item.to_dict() for item in remaining]
        if missing_objects or missing_generated or remaining:
            raise ContentIntegrityError(
                "交付物非文字对象来源锚点/生成许可不匹配（含宿主锚点）: "
                f"missing_source={len(missing_objects)}, "
                f"missing_generated={len(missing_generated)}, "
                f"extra={len(remaining)}"
            )

    # 4. 校验数学公式实例多重集
    missing_math = collections.Counter(expected.math_formulas) - collections.Counter(delivery_inv.math_formulas)
    if missing_math:
        raise ContentIntegrityError(f"交付物丢失 {sum(missing_math.values())} 个数学公式实例 (OMML)")

    # 5. 脚注/尾注按引用数量与定义内容核验；ID 可以由导入器合法重映射。
    def check_notes(kind: str, refs: List[str], definitions: Dict[str, str], actual_refs: List[str], actual_definitions: Dict[str, str]):
        if len(actual_refs) < len(refs):
            raise ContentIntegrityError(
                f"交付物丢失 {len(refs) - len(actual_refs)} 个{kind}引用实例"
            )
        required_texts = collections.Counter(
            definitions[note_id] for note_id in refs if note_id in definitions
        )
        actual_texts = collections.Counter(actual_definitions.values())
        missing_texts = required_texts - actual_texts
        if missing_texts:
            raise ContentIntegrityError(
                f"交付物丢失 {sum(missing_texts.values())} 个{kind}定义内容"
            )

    check_notes("脚注", expected.footnotes, expected.footnote_definitions,
                delivery_inv.footnotes, delivery_inv.footnote_definitions)
    check_notes("尾注", expected.endnotes, expected.endnote_definitions,
                delivery_inv.endnotes, delivery_inv.endnote_definitions)

    # 6. 字段指令、依赖和超链接目标按多重集核验；缓存值及 rId 可变化。
    missing_fields = collections.Counter(expected.fields) - collections.Counter(delivery_inv.fields)
    if missing_fields:
        remapped_missing = (
            collections.Counter(_field_identity_key(item) for item in expected.fields)
            - collections.Counter(_field_identity_key(item) for item in delivery_inv.fields)
        )
        if remapped_missing:
            raise ContentIntegrityError(
                f"交付物丢失 {sum(remapped_missing.values())} 个字段指令"
            )
    missing_dependencies = collections.Counter(expected.field_dependencies) - collections.Counter(delivery_inv.field_dependencies)
    if missing_dependencies:
        remapped_missing = (
            collections.Counter(_field_identity_key(item.replace(":", " ", 1)) for item in expected.field_dependencies)
            - collections.Counter(_field_identity_key(item.replace(":", " ", 1)) for item in delivery_inv.field_dependencies)
        )
        if remapped_missing:
            raise ContentIntegrityError(
                f"交付物丢失 {sum(remapped_missing.values())} 个字段依赖"
            )
    missing_links = collections.Counter(expected.hyperlinks) - collections.Counter(delivery_inv.hyperlinks)
    if missing_links:
        raise ContentIntegrityError(
            f"交付物超链接目标发生变化或丢失: {list(missing_links.elements())[:3]}"
        )

    missing_bookmarks = expected.bookmarks - delivery_inv.bookmarks
    if missing_bookmarks:
        raise ContentIntegrityError(f"交付物丢失 {len(missing_bookmarks)} 个书签名称: {list(missing_bookmarks)[:3]}")

    return True


def verify_generated_object_allowances(
    delivery_doc_or_path: Union[str, Path, Document],
    expected_inventory: ExpectedInventory,
) -> Dict[str, Any]:
    """Verify finite generated-object allowances for deliveries without content.

    Content deliveries run the same one-to-one consumption from
    :func:`verify_content_integrity`.  Cover-only or TOC-only deliveries do not
    enter that path, so they need an explicit gate of their own.  Source
    occurrences are consumed first, then every declared generated occurrence
    is consumed exactly once; anything left is an undeclared object.
    """
    allowances = list(expected_inventory.allowed_generated_objects)
    if not allowances:
        return {
            "checks": [],
            "failures": [],
            "passed": True,
            "mappings": [],
        }

    delivery_inventory = extract_semantic_inventory(delivery_doc_or_path)
    delivery_document = (
        Document(str(delivery_doc_or_path))
        if isinstance(delivery_doc_or_path, (str, Path))
        else delivery_doc_or_path
    )
    remaining = list(delivery_inventory.media_occurrences)
    declared_part_ids = tuple(
        str(item) for item in (expected_inventory.provenance.get("parts", ()) or ())
    )
    occurrence_parts = _occurrence_output_parts(
        delivery_document,
        remaining,
        declared_part_ids,
    )
    mappings: List[Dict[str, Any]] = []
    failures: List[Dict[str, Any]] = []

    def consume_source(expected_item: ObjectOccurrence) -> Optional[ObjectOccurrence]:
        for index, actual_item in enumerate(remaining):
            if (
                expected_item.kind == actual_item.kind
                and expected_item.fingerprint == actual_item.fingerprint
                and expected_item.story_type == actual_item.story_type
                and expected_item.local_index == actual_item.local_index
                and (not expected_item.host_text or expected_item.host_text == actual_item.host_text)
            ):
                return remaining.pop(index)
        return None

    source_occurrences = (
        expected_inventory.semantic.media_occurrences
        if expected_inventory.content_parts
        else ()
    )
    for source_item in source_occurrences:
        actual_item = consume_source(source_item)
        if actual_item is None:
            failures.append({"kind": "missing_source_object", **source_item.to_dict()})
        else:
                mappings.append({
                    "kind": "source_object",
                "source_sha256": source_item.source_sha256,
                "source_part_uri": source_item.part_uri,
                "source_host_element": source_item.host_element,
                "output_part_uri": actual_item.part_uri,
                    "output_host_element": actual_item.host_element,
                    "story_type": actual_item.story_type,
                    "fingerprint": actual_item.fingerprint,
                    "local_index": actual_item.local_index,
                    "output_part_ids": sorted(occurrence_parts.get(id(actual_item), set())),
                })

    for allowance in allowances:
        expected_count = max(1, int(allowance.get("count", 1)))
        for occurrence in range(expected_count):
            match_index = next(
                (
                    index for index, actual_item in enumerate(remaining)
                    if (
                        actual_item.kind == allowance.get("kind", "image")
                        and actual_item.fingerprint == allowance.get("fingerprint")
                        and actual_item.story_type == allowance.get("story_type", "body")
                        and (
                            allowance.get("host_text") in (None, "")
                            or actual_item.host_text == allowance.get("host_text")
                        )
                        and (
                            allowance.get("local_index") in (None, "")
                            or actual_item.local_index == int(allowance["local_index"])
                        )
                        and (
                            not allowance.get("part_id")
                            or allowance.get("part_id") in occurrence_parts.get(id(actual_item), set())
                        )
                    )
                ),
                None,
            )
            if match_index is None:
                failures.append({
                    "kind": "missing_generated_object",
                    "occurrence": occurrence,
                    **dict(allowance),
                })
                continue
            actual_item = remaining.pop(match_index)
            mappings.append({
                "kind": "generated_object",
                "part_id": allowance.get("part_id", "body"),
                "rule": allowance.get("rule", "explicit"),
                "output_part_uri": actual_item.part_uri,
                "output_host_element": actual_item.host_element,
                "story_type": actual_item.story_type,
                "fingerprint": actual_item.fingerprint,
                "local_index": actual_item.local_index,
                "output_part_ids": sorted(occurrence_parts.get(id(actual_item), set())),
            })

    for item in remaining:
        failures.append({"kind": "extra_object", **item.to_dict()})

    return {
        "checks": [{
            "name": "generated_object_allowances",
            "allowance_count": sum(max(1, int(item.get("count", 1))) for item in allowances),
            "source_occurrence_count": len(source_occurrences),
            "mapping_count": len(mappings),
        }],
        "failures": failures,
        "passed": not failures,
        "mappings": mappings,
    }


def build_expected_inventory(prepared: Any, spec: Optional[Dict[str, Any]] = None) -> ExpectedInventory:
    """从 PreparedBuild 的源顺序、部件与选区生成交付物期望清单。"""
    source_paths = [Path(item) for item in getattr(prepared, "source_order", ())]
    source_paths = [path for path in source_paths if path.suffix.lower() == ".docx" and path.is_file()]
    source_docx = getattr(prepared, "source_docx_path", None)
    if not source_paths and source_docx and Path(source_docx).is_file():
        source_paths = [Path(source_docx)]

    source_regions: Dict[str, str] = {}
    part_registry = getattr(prepared, "parts", {}) or {}
    selected_parts = (spec or {}).get("parts", [])
    for part_id in selected_parts:
        part = part_registry.get(part_id)
        region = getattr(part, "source_region", None) if part is not None else None
        if region:
            source_regions[part_id] = region

    inventories: List[SemanticInventory] = []
    selected_region_ids = set(source_regions.values())
    regions_registry = getattr(getattr(prepared, "config", None), "regions", {}) or {}
    # 自定义选区的唯一来源是已准备好的 SourceRegion；切片后再取清单，避免把全文带入期望范围。
    if selected_region_ids and len(source_paths) == 1 and selected_region_ids.issubset(regions_registry):
        from .document_parts import SelectionValidator

        source_doc = Document(str(source_paths[0]))
        spans = SelectionValidator.validate_regions(source_doc, regions_registry)
        # Do not iterate over the set above: two declared content parts may
        # intentionally reuse one region, and the expected inventory must
        # retain that instance multiplicity in delivery order.
        for part_id in selected_parts:
            region_id = source_regions.get(part_id)
            if not region_id:
                continue
            sliced, _ = SelectionValidator.slice_document_by_region_with_result(
                source_doc, spans[region_id]
            )
            inventories.append(extract_semantic_inventory(sliced))
    else:
        inventories = [extract_semantic_inventory(path) for path in source_paths]

    semantic = merge_semantic_inventories(inventories)
    nodes = list(getattr(prepared, "nodes", ()) or ())
    replacements = {}
    generated = collections.Counter()
    generated_media = collections.Counter()
    generated_objects: List[Dict[str, Any]] = []
    for node in nodes:
        source_text = node.get("raw_text") or node.get("source_text")
        title = node.get("title")
        if source_text and title:
            source_norm = normalize_text(source_text)
            title_norm = normalize_text(title)
            if source_norm and title_norm and source_norm != title_norm:
                replacements[source_norm] = title_norm
            if title_norm and title_norm not in semantic.paragraph_counts:
                generated[title_norm] += 1
        elif title and node.get("type") in {
            "folder", "docx", "docx_outline", "pdf", "pptx", "image"
        }:
            # Directory-tree rendering emits a visible outline heading for
            # every scanned folder/file node.  These headings are generated
            # from the frozen build plan and must be allowed explicitly,
            # without weakening the source paragraph denominator.
            title_norm = normalize_text(title)
            if title_norm and title_norm not in semantic.paragraph_counts:
                generated[title_norm] += 1

        if node.get("type") == "image" and node.get("file"):
            base = Path(getattr(prepared, "source_path", "."))
            if base.is_file():
                base = base.parent
            image_path = (base / str(node["file"])).resolve()
            if image_path.is_file():
                fingerprint = compute_file_sha256(image_path)
                generated_media[fingerprint] += 1
                generated_objects.append({
                    "part_id": "body",
                    "kind": "image",
                    "story_type": "body",
                    "fingerprint": fingerprint,
                    "count": 1,
                    "rule": "explicit_image_node",
                    "source_file": str(image_path),
                    "source_node_title": title or "",
                })

    # Template media are generated in the cover part and are not source
    # content.  Their exact fingerprint and occurrence count form a finite
    # allowance; they are never allowed globally across the whole document.
    cover_template = getattr(prepared, "cover_template_path", None)
    if cover_template and Path(cover_template).is_file():
        template_inventory = extract_semantic_inventory(Path(cover_template))
        for occurrence in template_inventory.media_occurrences:
            generated_media[occurrence.fingerprint] += 1
            generated_objects.append({
                "part_id": "cover",
                "kind": occurrence.kind,
                "story_type": occurrence.story_type,
                "fingerprint": occurrence.fingerprint,
                "host_text": occurrence.host_text,
                "local_index": occurrence.local_index,
                "count": 1,
                "rule": "cover_template_occurrence",
                "source_part_uri": occurrence.part_uri,
                "source_host_element": occurrence.host_element,
            })

    coverage = {
        "source_files": len(source_paths),
        "paragraph_instances": len(semantic.paragraphs),
        "table_instances": len(semantic.tables),
        "media_instances": len(semantic.media_instances),
        "formula_instances": len(semantic.math_formulas),
        "footnote_references": len(semantic.footnotes),
        "endnote_references": len(semantic.endnotes),
        "field_instructions": len(semantic.fields),
        "hyperlink_targets": len(semantic.hyperlinks),
    }
    return ExpectedInventory(
        semantic=semantic,
        source_order=tuple(str(path) for path in source_paths),
        source_regions=source_regions,
        allowed_generated_text=dict(generated),
        allowed_generated_media=dict(generated_media),
        allowed_generated_objects=tuple(generated_objects),
        replacements=replacements,
        allowed_id_remaps={
            "bookmark_id": "allowed_if_name_and_range_semantics_preserved",
            "relationship_id": "allowed_if_target_and_type_preserved",
            "style_id": "allowed_if_role_and_effective_style_preserved",
            "numbering_id": "allowed_if_numbering_semantics_preserved",
            "note_id": "allowed_if_definition_content_and_reference_count_preserved",
        },
        provenance={
            "kind": "PreparedBuild",
            "delivery_id": (spec or {}).get("id"),
            "parts": list(selected_parts),
        },
        coverage=coverage,
        # A real PreparedBuild always carries the resolved parts registry. A
        # legacy smoke compatibility view may not; keep that old whole-body
        # entry point from demanding synthetic boundary bookmarks.
        content_parts=tuple(
            part_id
            for part_id in selected_parts
            if part_registry and (
                part_id == "body"
                or (part_registry.get(part_id) is not None and getattr(part_registry[part_id], "kind", None) == "content")
            )
        ),
    )


def _part_visible_text(doc: Document, part_id: str) -> str:
    """Read one assembled part by its non-visible boundary bookmarks."""
    children = list(doc.element.body)
    start_name = get_part_boundary_bookmark(part_id, True)
    end_name = get_part_boundary_bookmark(part_id, False)
    start = next((i for i, child in enumerate(children) if any(
        marker.get(qn("w:name")) == start_name for marker in child.iter(qn("w:bookmarkStart"))
    )), None)
    if start is None:
        return ""
    end = next((i for i in range(start, len(children)) if any(
        marker.get(qn("w:name")) == end_name for marker in children[i].iter(qn("w:bookmarkStart"))
    )), None)
    if end is None:
        following = []
        for i in range(start + 1, len(children)):
            if any(
                marker.get(qn("w:name")) in {"_Synth_cover", "_Synth_toc", "_Synth_body"}
                for marker in children[i].iter(qn("w:bookmarkStart"))
            ):
                following.append(i)
        end = min(following) if following else len(children) - 1
    return normalize_text("".join(
        text.text or ""
        for child in children[start:end + 1]
        for text in child.iter(qn("w:t"))
    ))


def _bookmark_texts(doc: Document, name: str) -> List[str]:
    """Return each exact text range carrying a generated-content marker."""
    elements = list(doc.element.body.iter())
    starts = [
        (index, node)
        for index, node in enumerate(elements)
        if node.tag == qn("w:bookmarkStart") and node.get(qn("w:name")) == name
    ]
    values = []
    for start_index, start in starts:
        bookmark_id = start.get(qn("w:id"))
        end_index = next(
            (
                index for index in range(start_index + 1, len(elements))
                if elements[index].tag == qn("w:bookmarkEnd")
                and elements[index].get(qn("w:id")) == bookmark_id
            ),
            None,
        )
        if end_index is None:
            values.append(None)
            continue
        values.append("".join(
            node.text or ""
            for node in elements[start_index + 1:end_index]
            if node.tag == qn("w:t")
        ))
    return values


def verify_generated_content(
    delivery_doc_or_path: Union[str, Path, Document],
    expected_delivery: Any,
) -> Dict[str, Any]:
    """Verify generated fields by declared part and exact expected count."""
    # ``docx.Document`` is a factory function rather than the concrete
    # document class, so it cannot be used as the second argument to
    # ``isinstance``.  Distinguish path inputs explicitly and keep an already
    # opened document object untouched.
    doc = (
        Document(str(delivery_doc_or_path))
        if isinstance(delivery_doc_or_path, (str, Path))
        else delivery_doc_or_path
    )
    checks = []
    failures = []
    expected_items = list(getattr(expected_delivery, "generated_content", ()) or ())
    part_texts = {
        part_id: _part_visible_text(doc, part_id)
        for part_id in {item.get("part_id") for item in expected_items}
        if part_id
    }
    expected_counts: Dict[Tuple[str, str], int] = collections.Counter()
    for item in expected_items:
        part_id = item.get("part_id")
        expected = normalize_text(item.get("expected_text"))
        if part_id and expected:
            expected_counts[(part_id, expected)] += max(1, int(item.get("expected_count", 1)))

    for item in expected_items:
        part_id = item.get("part_id")
        expected = normalize_text(item.get("expected_text"))
        actual = part_texts.get(part_id, "")
        marker_names = list(item.get("marker_names", ()) or ())
        marker_values = []
        marker_failures = []
        if marker_names:
            for marker_name in marker_names:
                values = _bookmark_texts(doc, marker_name)
                if len(values) != 1:
                    marker_failures.append({"marker": marker_name, "values": values})
                else:
                    marker_values.append(normalize_text(values[0]))
            expected_count = max(1, int(item.get("expected_count", len(marker_names))))
            actual_count = len(marker_values)
            ok = (
                bool(expected)
                and actual_count == expected_count
                and not marker_failures
                and all(value == expected for value in marker_values)
            )
        else:
            expected_count = expected_counts.get((part_id, expected), 0)
            actual_count = actual.count(expected) if expected else 0
            ok = bool(expected) and actual_count == expected_count
        result = {
            "part_id": part_id,
            "kind": item.get("kind"),
            "field": item.get("field"),
            "expected_text_hash": item.get("expected_text_hash"),
            "actual_part_text_hash": hashlib.sha256(actual.encode("utf-8")).hexdigest(),
            "expected_count": expected_count,
            "actual_count": actual_count,
            "position": item.get("position", {"part_id": part_id}),
            "marker_names": marker_names,
            "marker_values": marker_values,
            "marker_failures": marker_failures,
            "status": "passed" if ok else "failed",
        }
        checks.append(result)
        if not ok:
            failures.append(result)
    return {"passed": not failures, "checks": checks, "failures": failures}


def verify_delivery_format(
    delivery_doc_or_path: Union[str, Path],
    resolved_format: ResolvedFormat,
    role_map: Optional[Dict[str, str]] = None,
    verification_context: Optional[Any] = None,
) -> FormatVerificationReport:
    """
    对已装配的交付文档执行只读有效样式求值，
    逐项验证托管角色（body, heading.1..9, title 等）的最终有效字体、字号、间距与大纲级别。
    """
    doc_path = Path(delivery_doc_or_path).resolve()
    inspection = inspect_docx(doc_path)

    verified_roles = set()
    violations = []
    unverified_attributes = []
    items: List[Dict[str, Any]] = []
    story_checks: List[Dict[str, Any]] = []
    logical_interval_checks: List[Dict[str, Any]] = []

    def serial(value: Any) -> Any:
        if hasattr(value, "__dataclass_fields__"):
            return asdict(value)
        return value

    def record(role: str, block: BlockInspection, attribute: str, expected: Any, actual: Any, ok: bool):
        labels = {
            "size_pt": "字号",
            "east_asia": "中文字体",
            "latin": "西文字体",
            "bold": "加粗状态",
            "italic": "斜体状态",
            "color": "文字颜色",
            "alignment": "对齐方式",
            "outline_level": "大纲级别",
        }
        label = labels.get(attribute.rsplit(".", 1)[-1], attribute)
        coverage = "verified" if actual is not None else "unverified"
        item = {
            "role": role,
            "attribute": attribute,
            "expected": serial(expected),
            "actual": serial(actual),
            "provenance": {
                "source": "ResolvedFormat",
                "node": asdict(block.node),
                "part_uri": block.node.part_uri,
            },
            "coverage": coverage,
        }
        items.append(item)
        if actual is None:
            unverified_attributes.append(f"{role}.{attribute}")
            violations.append(f"[{role} @ {block.node.element_path}] {label} 无法求得有效值")
        elif not ok:
            violations.append(
                f"[{role} @ {block.node.element_path}] {label} 不匹配: "
                f"期望 {expected}, 实际 {actual}"
            )

    def equal_value(expected: Any, actual: Any, tolerance: float = 0.01) -> bool:
        if isinstance(expected, dict) and hasattr(actual, "__dataclass_fields__"):
            return equal_value(expected, asdict(actual), tolerance)
        if hasattr(expected, "__dataclass_fields__") and isinstance(actual, dict):
            return equal_value(asdict(expected), actual, tolerance)
        if isinstance(expected, dict) and isinstance(actual, dict):
            if set(expected) != set(actual):
                return False
            return all(equal_value(expected[key], actual[key], tolerance) for key in expected)
        if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
            return abs(float(expected) - float(actual)) <= tolerance
        if isinstance(expected, LengthValue) and isinstance(actual, LengthValue):
            return expected.unit == actual.unit and abs(expected.value - actual.value) <= tolerance
        if isinstance(expected, LineSpacing) and isinstance(actual, LineSpacing):
            return expected.mode == actual.mode and (
                expected.value is None or actual.value is not None and abs(expected.value - actual.value) <= tolerance
            )
        # Word may canonicalize the theme/default black color `auto` to an
        # explicit RGB black when it saves an imported document.  They render
        # identically and represent the same default color semantics.
        if isinstance(expected, str) and isinstance(actual, str):
            if {expected.lower(), actual.lower()} <= {"auto", "000000", "00000000"}:
                return True
        return expected == actual

    def font_equal(expected: Optional[str], actual: Optional[str]) -> bool:
        if expected is None:
            return True
        if actual is None:
            return False
        if expected.lower() == actual.lower():
            return True
        return {expected, actual} in [
            {"宋体", "SimSun"}, {"黑体", "SimHei"}, {"楷体", "KaiTi"}, {"仿宋", "FangSong"}
        ]

    def compare_serialized_spans(
        expected_spans: Iterable[Mapping[str, Any]],
        actual_spans: Iterable[Mapping[str, Any]],
        attribute: str,
    ) -> Optional[List[Dict[str, Any]]]:
        """Compare one story's protected property over merged logical intervals."""
        expected_list = list(expected_spans or ())
        actual_list = list(actual_spans or ())
        if not expected_list or not actual_list:
            return None
        if not any(_span_style_value(span, attribute) is not None for span in expected_list):
            return None
        expected_end = max(int(span.get("end", 0)) for span in expected_list)
        actual_end = max(int(span.get("end", 0)) for span in actual_list)
        if expected_end != actual_end:
            return [{
                "start": 0,
                "end": max(expected_end, actual_end),
                "expected": {"logical_text_length": expected_end},
                "actual": {"logical_text_length": actual_end},
                "ok": False,
            }]
        boundaries = {0, expected_end}
        for span in expected_list + actual_list:
            boundaries.add(int(span.get("start", 0)))
            boundaries.add(int(span.get("end", 0)))
        ordered = sorted(value for value in boundaries if 0 <= value <= expected_end)
        comparisons = []
        for start, end in zip(ordered, ordered[1:]):
            if start == end:
                continue
            expected_span = next(
                (span for span in expected_list
                 if int(span.get("start", 0)) <= start and end <= int(span.get("end", 0))),
                None,
            )
            actual_span = next(
                (span for span in actual_list
                 if int(span.get("start", 0)) <= start and end <= int(span.get("end", 0))),
                None,
            )
            expected_value = _span_style_value(expected_span or {}, attribute)
            actual_value = _span_style_value(actual_span or {}, attribute)
            comparisons.append({
                "start": start,
                "end": end,
                "expected": expected_value,
                "actual": actual_value,
                "ok": (
                    expected_span is not None
                    and actual_span is not None
                    and _inline_value_equal(expected_value, actual_value, attribute)
                ),
            })
        return comparisons

    # roles.style 是实际样式 ID，允许它与语义角色名不同（例如 style_body）。
    # 核验时建立 role -> StyleDefinition 视图，避免把已声明的别名误判为未核验。
    role_styles = dict(resolved_format.styles)
    for role_name, role_spec in resolved_format.roles.items():
        style_def = role_styles.get(role_spec.style)
        if style_def is not None:
            # A child package may inherit a style with the same semantic name
            # while mapping the role to a distinct compiled style id (for
            # example ``heading.1`` -> ``style_heading_1``).  The role alias
            # must point at the declared role style, not remain shadowed by
            # the inherited style entry.
            role_styles[role_name] = style_def

    synth_role_map = {
        "synthbody": "body",
        "synthtablebody": "table.body",
        "synthtitle": "title",
        "synthsubtitle": "subtitle",
        "synthquote": "quote",
        "synthcaption": "caption",
        "synthheading1": "heading.1",
        "synthheading2": "heading.2",
        "synthheading3": "heading.3",
        "synthheading4": "heading.4",
        "synthheading5": "heading.5",
        "synthheading6": "heading.6",
        "synthheading7": "heading.7",
        "synthheading8": "heading.8",
        "synthheading9": "heading.9",
    }
    style_to_role = {
        role_spec.style.lower(): role_name
        for role_name, role_spec in resolved_format.roles.items()
        if role_spec.style
    }

    verified_count = 0

    # 页面几何默认属于目标格式；完整 VerificationContext 会明确指出
    # preserve/source 应比较的源节规格，不能无条件套用 A4。
    page = getattr(resolved_format, "page", None)
    page_geometry = []
    if page is not None:
        target_geometry = {
            "width_mm": float(page.width_mm),
            "height_mm": float(page.height_mm),
            "margin_top_mm": float(page.margin_top_mm),
            "margin_bottom_mm": float(page.margin_bottom_mm),
            "margin_left_mm": float(page.margin_left_mm),
            "margin_right_mm": float(page.margin_right_mm),
        }
        source_geometries = []
        if verification_context is not None:
            source_geometries = list(
                getattr(verification_context.expected_delivery, "page_expectations", ()) or ()
            )
            actual_section_count = len(Document(str(doc_path)).sections)
            if source_geometries and len(source_geometries) != actual_section_count:
                violations.append(
                    "输出节数量与 preserve/source 的来源节数量不一致: "
                    f"期望 {len(source_geometries)}，实际 {actual_section_count}"
                )
        for index, section in enumerate(Document(str(doc_path)).sections):
            expected_raw = source_geometries[index] if index < len(source_geometries) else target_geometry
            expected_geometry = {
                key: float(expected_raw[key])
                for key in target_geometry
                if key in expected_raw
            }
            actual_geometry = {
                "width_mm": float(section.page_width.mm),
                "height_mm": float(section.page_height.mm),
                "margin_top_mm": float(section.top_margin.mm),
                "margin_bottom_mm": float(section.bottom_margin.mm),
                "margin_left_mm": float(section.left_margin.mm),
                "margin_right_mm": float(section.right_margin.mm),
            }
            page_geometry.append({"section": index, "expected": expected_geometry, "actual": actual_geometry})
            for attribute, expected in expected_geometry.items():
                actual = actual_geometry[attribute]
                items.append({
                    "role": "page",
                    "attribute": attribute,
                    "expected": expected,
                    "actual": actual,
                    "provenance": {"source": "ResolvedFormat.page", "section": index},
                    "coverage": "verified",
                })
                if abs(expected - actual) > 0.1:
                    label = "页面尺寸" if attribute in {"width_mm", "height_mm"} else "页面边距"
                    violations.append(
                        f"[section {index}] {label} {attribute} 不匹配: "
                        f"期望 {expected:.1f}mm, 实际 {actual:.1f}mm"
                    )

    if verification_context is not None:
        expected_delivery = verification_context.expected_delivery
        if expected_delivery.policy_mode in {"preserve", "mixed"}:
            actual_stories = {
                (story.story_type, story.variant, story.part_uri): story
                for story in inspection.stories
            }
            actual_bindings = {
                (binding.section_index, binding.story_type, binding.variant): binding
                for binding in inspection.section_bindings
            }
            for expected_story in expected_delivery.story_expectations:
                key = (
                    expected_story["section_index"],
                    expected_story["story_type"],
                    expected_story["variant"],
                )
                binding = actual_bindings.get(key)
                if binding is None:
                    violations.append(
                        f"[section {key[0]}] 缺少 {key[1]} story 绑定: {key[2]}"
                    )
                    story_checks.append({"key": key, "status": "missing_binding"})
                    continue
                actual_story = actual_stories.get(
                    (binding.story_type, binding.variant, binding.part_uri)
                )
                if actual_story is None:
                    violations.append(
                        f"[section {key[0]}] 缺少 {key[1]} story 部件: {binding.part_uri}"
                    )
                    story_checks.append({"key": key, "status": "missing_part"})
                    continue
                expected_texts = list(expected_story.get("texts", ()))
                actual_texts = [block.visible_text for block in actual_story.blocks]
                text_ok = expected_texts == actual_texts
                if not text_ok:
                    violations.append(
                        f"[section {key[0]}] {key[1]} story 文字不匹配: "
                        f"期望 {expected_texts[:3]!r}, 实际 {actual_texts[:3]!r}"
                    )
                actual_section = (
                    inspection.sections[key[0]]
                    if key[0] < len(inspection.sections) else None
                )
                distance_attribute = (
                    "header_distance_mm"
                    if key[1] == "header" else "footer_distance_mm"
                )
                expected_distance = expected_story.get(distance_attribute)
                actual_distance = (
                    getattr(actual_section.page_spec, distance_attribute, None)
                    if actual_section else None
                )
                distance_ok = (
                    expected_distance is None or actual_distance is not None
                    and abs(float(expected_distance) - float(actual_distance)) <= 0.1
                )
                expected_linked = bool(expected_story.get("linked_to_previous", False))
                linked_ok = binding.linked_to_previous == expected_linked
                if not linked_ok:
                    violations.append(
                        f"[section {key[0]}] {key[1]} {key[2]} 链接状态不匹配: "
                        f"期望 linked_to_previous={expected_linked}, 实际 {binding.linked_to_previous}"
                    )
                first_page_ok = True
                expected_first_page = bool(expected_story.get("different_first_page", False))
                if actual_section is not None:
                    first_page_ok = actual_section.different_first_page == expected_first_page
                    if not first_page_ok:
                        violations.append(
                            f"[section {key[0]}] different_first_page 不匹配: "
                            f"期望 {expected_first_page}, 实际 {actual_section.different_first_page}"
                        )

                span_ok = True
                expected_blocks = list(expected_story.get("blocks", ()))
                if len(expected_blocks) != len(actual_story.blocks):
                    span_ok = False
                    violations.append(
                        f"[section {key[0]}] {key[1]} story 段落数量不匹配: "
                        f"期望 {len(expected_blocks)}, 实际 {len(actual_story.blocks)}"
                    )
                for block_index, expected_block in enumerate(expected_blocks):
                    if block_index >= len(actual_story.blocks):
                        continue
                    actual_block = actual_story.blocks[block_index]
                    for attribute in ("east_asia", "latin", "size_pt", "bold", "italic", "color"):
                        comparisons = compare_serialized_spans(
                            expected_block.get("spans", ()),
                            actual_block.effective_run_spans,
                            attribute,
                        )
                        if comparisons is None:
                            continue
                        for comparison in comparisons:
                            item = {
                                "role": f"{key[1]}.{key[2]}",
                                "attribute": f"block[{block_index}].run[{comparison['start']}:{comparison['end']}].{attribute}",
                                "expected": serial(comparison["expected"]),
                                "actual": serial(comparison["actual"]),
                                "provenance": {
                                    "source": "VerificationContext.story_expectations",
                                    "section": key[0],
                                    "part_uri": binding.part_uri,
                                },
                                "coverage": "verified" if comparison["actual"] is not None else "unverified",
                            }
                            items.append(item)
                            if comparison["actual"] is None:
                                span_ok = False
                                unverified_attributes.append(item["attribute"])
                            elif not comparison["ok"]:
                                span_ok = False
                                violations.append(
                                    f"[section {key[0]}] {key[1]} story {attribute} 区间不匹配: "
                                    f"{comparison['start']}:{comparison['end']}"
                                )
                items.append({
                    "role": f"{key[1]}.{key[2]}",
                    "attribute": "text",
                    "expected": expected_texts,
                    "actual": actual_texts,
                    "provenance": {"source": "VerificationContext.story_expectations", "section": key[0]},
                    "coverage": "verified",
                })
                items.append({
                    "role": f"{key[1]}.{key[2]}",
                    "attribute": distance_attribute,
                    "expected": expected_distance,
                    "actual": actual_distance,
                    "provenance": {"source": "SectionInspection.page_spec", "section": key[0]},
                    "coverage": "verified" if actual_distance is not None else "unverified",
                })
                if not distance_ok:
                    violations.append(
                        f"[section {key[0]}] {key[1]} 距离不匹配: "
                        f"期望 {expected_distance}, 实际 {actual_distance}"
                    )
                story_checks.append({
                    "key": key,
                    "status": "verified" if text_ok and distance_ok and linked_ok and first_page_ok and span_ok else "failed",
                    "text_ok": text_ok,
                    "distance_ok": distance_ok,
                    "linked_to_previous_ok": linked_ok,
                    "different_first_page_ok": first_page_ok,
                    "span_ok": span_ok,
                    "part_uri": binding.part_uri,
                })

    context_entries = None
    if verification_context is not None:
        expected_delivery = verification_context.expected_delivery
        if verification_context.output_node_map.duplicate_occurrences:
            violations.append(
                "输出内容部件包含未声明的重复来源实例: "
                + ", ".join(verification_context.output_node_map.duplicate_occurrences[:3])
            )
        block_by_path = {block.node.element_path: block for block in inspection.blocks}
        context_entries = []
        for expected_node in expected_delivery.ordered_nodes:
            occurrence_id = expected_node.occurrence.occurrence_id
            locations = verification_context.output_node_map.locations.get(occurrence_id, ())
            if len(locations) != 1:
                reason = "缺失" if not locations else "重复"
                unverified_attributes.append(f"{expected_node.role}@{occurrence_id}")
                violations.append(
                    f"[{expected_node.role} @ {expected_node.occurrence.element_path}] "
                    f"输出节点映射{reason}，不能缩小格式核验覆盖分母"
                )
                continue
            actual = block_by_path.get(locations[0].get("element_path"))
            if actual is None:
                unverified_attributes.append(f"{expected_node.role}@{occurrence_id}")
                violations.append(
                    f"[{expected_node.role} @ {expected_node.occurrence.element_path}] "
                    "输出节点无法重新定位"
                )
                continue
            if not locations[0].get("text_hash_matches", True):
                unverified_attributes.append(f"{expected_node.role}@{occurrence_id}")
                violations.append(
                    f"[{expected_node.role} @ {expected_node.occurrence.element_path}] "
                    "输出 provenance 节点的语义摘要与来源不一致"
                )
                continue
            context_entries.append((actual, expected_node))
    else:
        context_entries = [(block, None) for block in inspection.blocks]

    def value_of(value: Any, attribute: str) -> Any:
        if isinstance(value, dict):
            return value.get(attribute)
        return getattr(value, attribute, None) if value is not None else None

    for block, expected_node in context_entries:
        p_style = (block.p_style_id or block.p_style_name or "").lower()
        role = expected_node.role if expected_node is not None else synth_role_map.get(p_style)
        if not role:
            if role_map and block.node.element_path in role_map:
                role = role_map[block.node.element_path]
            elif p_style in synth_role_map.values():
                role = p_style
            else:
                role = style_to_role.get(p_style)
        if not role and "/w:tbl[" in block.node.element_path:
            # PackageImporter may prefix a source style ID during assembly;
            # the structural NodeRef still unambiguously identifies a cell.
            role = "table.body"
        if not role and p_style.startswith("synthpart_"):
            # PackageImporter prefixes source style IDs during part assembly.
            # The effective outline level still identifies headings, while
            # non-heading imported paragraphs are the managed body role.
            if block.outline_level and 1 <= int(block.outline_level) <= 9:
                role = f"heading.{int(block.outline_level)}"
            else:
                role = "body"

        if not role or role not in role_styles:
            if expected_node is not None:
                unverified_attributes.append(f"{role or 'unknown'}@{expected_node.occurrence.occurrence_id}")
                violations.append(
                    f"[{role or 'unknown'} @ {expected_node.occurrence.element_path}] "
                    "预期角色没有可用的格式样式"
                )
            continue

        policy_mode = (
            getattr(verification_context.expected_delivery, "policy_mode", "restyle")
            if verification_context is not None else "restyle"
        )
        inline_emphasis = (
            getattr(verification_context.expected_delivery, "inline_emphasis", "preserve")
            if verification_context is not None else "target"
        )
        style_def = role_styles[role]
        verified_roles.add(role)
        verified_count += 1

        run_eff = block.effective_run
        para_eff = block.effective_paragraph
        if expected_node is not None and not expected_node.managed_properties:
            expected_run = expected_node.semantic_summary.get("source_effective_run")
            expected_runs = expected_node.semantic_summary.get("source_effective_runs") or []
            expected_paragraph = expected_node.semantic_summary.get("source_effective_paragraph")
        else:
            expected_run = style_def.run
            expected_runs = []
            expected_paragraph = style_def.paragraph

        logical_inline_attributes = set()
        if expected_node is not None and inline_emphasis == "preserve":
            for attribute in ("east_asia", "latin", "size_pt", "bold", "italic", "color"):
                target_value = value_of(expected_run, attribute)
                comparisons = _logical_inline_segments(
                    block,
                    expected_node,
                    attribute,
                    target_value=target_value,
                )
                if comparisons is None:
                    continue
                if any(item.get("source_protected") for item in comparisons):
                    logical_inline_attributes.add(attribute)
                expected_intervals = [
                    [int(span.get("start", 0)), int(span.get("end", 0))]
                    for span in expected_node.semantic_summary.get("source_effective_run_spans", ()) or ()
                    if _span_style_value(span, attribute) is not None
                ]
                verified_intervals = [
                    [int(item["start"]), int(item["end"])]
                    for item in comparisons
                    if item.get("actual") is not None and item.get("ok")
                ]
                missing_intervals = [
                    [int(item["start"]), int(item["end"])]
                    for item in comparisons
                    if item.get("actual") is None or not item.get("ok")
                ]
                logical_interval_checks.append({
                    "role": role,
                    "node": asdict(block.node),
                    "attribute": attribute,
                    "expected_intervals": expected_intervals,
                    "verified_intervals": verified_intervals,
                    "missing_intervals": missing_intervals,
                })
                for comparison in comparisons:
                    record(
                        role,
                        block,
                        f"run[{comparison['start']}:{comparison['end']}].{attribute}",
                        comparison["expected"],
                        comparison["actual"],
                        comparison["ok"],
                    )

        # 1. 每个 run 都参与求值，不能用首个 run 代表混合段落。
        if expected_run:
            run_values = block.effective_runs or ((run_eff,) if run_eff else ())
            if not run_values:
                for attribute in ("east_asia", "latin", "size_pt", "bold", "italic", "color"):
                    if value_of(expected_run, attribute) is not None:
                        record(role, block, f"run.{attribute}", value_of(expected_run, attribute), None, False)
            for run_index, actual_run in enumerate(run_values):
                prefix = f"run[{run_index}]"
                for attribute in ("east_asia", "latin"):
                    if attribute in logical_inline_attributes:
                        continue
                    expected_value = value_of(expected_run, attribute)
                    if expected_value is not None:
                        actual_value = getattr(actual_run, attribute)
                        record(role, block, f"{prefix}.{attribute}", expected_value, actual_value, font_equal(expected_value, actual_value))
                for attribute, tolerance in (("size_pt", 0.5),):
                    if attribute in logical_inline_attributes:
                        continue
                    expected_value = value_of(expected_run, attribute)
                    if expected_value is not None:
                        actual_value = getattr(actual_run, attribute)
                        record(role, block, f"{prefix}.{attribute}", expected_value, actual_value, equal_value(expected_value, actual_value, tolerance))
                for attribute in ("bold", "italic", "color"):
                    if attribute in logical_inline_attributes:
                        continue
                    expected_source_run = expected_runs[run_index] if run_index < len(expected_runs) else expected_run
                    expected_value = value_of(expected_source_run, attribute) if inline_emphasis == "preserve" and attribute in {"bold", "italic"} else value_of(expected_run, attribute)
                    if expected_value is not None:
                        actual_value = getattr(actual_run, attribute)
                        record(role, block, f"{prefix}.{attribute}", expected_value, actual_value, equal_value(expected_value, actual_value))

        # 4. 验证大纲级别 (针对标题角色)
        if role.startswith("heading."):
            expected_level = int(role.split(".")[1])
            record(role, block, "outline_level", expected_level, block.outline_level, block.outline_level == expected_level)

        # 5. 段落最终有效属性：段前后、行距、首行/悬挂、keep/widow 等均报告证据。
        if expected_paragraph:
            for attribute in (
                "alignment", "first_line_indent", "hanging_indent", "left_indent", "right_indent",
                "space_before_pt", "space_after_pt", "line_spacing", "keep_with_next", "keep_lines",
                "widow_control", "page_break_before", "snap_to_grid",
            ):
                expected_value = value_of(expected_paragraph, attribute)
                if expected_node is not None:
                    expected_value = expected_node.semantic_summary.get(
                        "format_overrides", {}
                    ).get(attribute, expected_value)
                if expected_value is None:
                    continue
                actual_value = getattr(para_eff, attribute, None) if para_eff else None
                record(role, block, f"paragraph.{attribute}", expected_value, actual_value, equal_value(expected_value, actual_value))

    has_context = verification_context is not None
    # Cover-only/TOC-only deliveries legitimately have no source nodes. Their
    # generated-content and page checks are separate required gates; the
    # format gate must not fabricate a denominator of one source paragraph.
    passed = (len(violations) == 0) and (len(unverified_attributes) == 0) and (
        verified_count > 0 or has_context
    )
    return FormatVerificationReport(
        passed=passed,
        verified_count=verified_count,
        violation_count=len(violations),
        verified_roles=sorted(list(verified_roles)),
        violations=violations[:20],
        unverified_attributes=sorted(set(unverified_attributes)),
        details={
            "total_blocks": len(inspection.blocks),
            "verified_blocks": verified_count,
            "page_geometry": page_geometry,
            "items": items,
            "story_checks": story_checks,
            "logical_interval_checks": logical_interval_checks,
            "coverage": {
                "roles": sorted(verified_roles),
                "properties": "effective_style_per_run_and_paragraph",
                "unverified_attributes": sorted(set(unverified_attributes)),
            },
        }
    )
