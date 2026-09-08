# -*- coding: utf-8 -*-
"""
文档抽象部件与源选区规范模型与验证器 (lib/document_parts.py)

定义 DocumentPart、PartKind、SourceRegion，并提供基于稳定 NodeRef 的文档选区切片与边界校验。
"""

from dataclasses import dataclass
from enum import Enum
import hashlib
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from docx import Document
from docx.oxml.ns import qn
import lxml.etree


class PartKind(str, Enum):
    COVER = "cover"
    GENERATED_TOC = "generated_toc"
    CONTENT = "content"


@dataclass(frozen=True)
class DocumentPart:
    """具名文档部件"""
    id: str
    kind: str  # "cover" | "generated_toc" | "content"
    source_region: Optional[str] = None
    section_ref: Optional[str] = None
    section_type: str = "nextPage"  # "nextPage" | "oddPage" | "evenPage" | "continuous"
    page_sequence: Optional[str] = None
    include_in_toc: bool = True

    def __post_init__(self):
        if self.kind == "toc":
            object.__setattr__(self, "kind", PartKind.GENERATED_TOC.value)
        valid_kinds = {k.value for k in PartKind}
        if self.kind not in valid_kinds:
            raise ValueError(f"无效的部件类型 kind: {self.kind}，允许值: {valid_kinds}")
        if self.kind == PartKind.CONTENT.value and not self.source_region and self.id != "body":
            # 兼容默认 body 部件可省略 source_region
            pass


@dataclass(frozen=True)
class SourceRegion:
    """源选区定义 (基于顶层块 NodeRef 路径)"""
    id: str
    start: str  # 包含 (inclusive)，例如 "/w:document/w:body/w:p[1]"
    end: str    # 不包含 (exclusive)，例如 "/w:document/w:body/w:p[10]" 或 "end_of_document"
    exclude: bool = False  # 若为 True，该选区内容将被合法忽略不输出


class SelectionError(ValueError):
    """选区处理异常基类"""

    code = "SELECTION_ERROR"


class SelectionOverlapError(SelectionError):
    """选区相互重叠异常"""

    code = "SELECTION_OVERLAP"


class UnassignedContentError(SelectionError):
    """存在未分配实质内容异常"""

    code = "UNASSIGNED_CONTENT"


class InvalidBoundaryError(SelectionError):
    """非法选区边界异常 (例如边界切断表格内部)"""

    code = "INVALID_BOUNDARY"


class ProtectedBoundaryError(InvalidBoundaryError):
    """选区边界切断字段或书签等不可拆分对象"""

    code = "PROTECTED_BOUNDARY"


class UnsupportedObjectError(SelectionError):
    """源文档包含当前导入器无法安全处理的顶层对象"""

    code = "UNSUPPORTED_OBJECT"


def get_default_parts() -> Dict[str, DocumentPart]:
    """为未声明 layout.parts 的 v1/v2 或 v3 项目提供标准向下兼容部件注册表"""
    return {
        "cover": DocumentPart(id="cover", kind=PartKind.COVER.value, include_in_toc=False),
        "toc": DocumentPart(id="toc", kind=PartKind.GENERATED_TOC.value, include_in_toc=False),
        "body": DocumentPart(id="body", kind=PartKind.CONTENT.value, source_region="entire_document", include_in_toc=True),
    }


def get_part_boundary_bookmark(part_id: str, is_start: bool = True) -> str:
    """
    生成部件边界书签名称。
    对历史默认部件保持兼容名称：_Synth_cover, _Synth_toc, _Synth_body (起点)。
    对自定义部件及结束边界生成规范名称：_Synth_part_<id>_start / _Synth_part_<id>_end。
    """
    if is_start:
        if part_id == "cover":
            return "_Synth_cover"
        if part_id == "toc":
            return "_Synth_toc"
        if part_id == "body":
            return "_Synth_body"
    elif part_id in ("cover", "toc", "body"):
        return f"_Synth_part_{part_id}_end"
    suffix = "start" if is_start else "end"
    return f"_Synth_part_{part_id}_{suffix}"


def generated_content_bookmark_name(part_id: str, field: str, index: int) -> str:
    """Return a stable non-visible marker for one generated-content instance."""
    identity = f"{part_id}|{field}|{int(index)}"
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:24]
    return f"_SynthGen_{digest}"


class SelectionValidator:
    """文档顶层块扫描与源选区验证器"""

    @staticmethod
    def _is_substantive_element(element: Any, visible_text: str = "") -> bool:
        """判断块是否含有不可静默丢弃的内容。

        图片、公式、字段和脚注引用即使没有 ``w:t`` 也属于实质内容；
        空表格和内容控件也必须被选区覆盖，因为它们携带结构语义。
        """
        if visible_text.strip():
            return True
        if element.tag in (qn("w:tbl"), qn("w:sdt")):
            return True
        protected_tags = {
            qn("w:drawing"),
            qn("w:pict"),
            qn("w:object"),
            qn("w:fldChar"),
            qn("w:fldSimple"),
            qn("w:footnoteReference"),
            qn("w:endnoteReference"),
            qn("w:hyperlink"),
            qn("m:oMath"),
            qn("m:oMathPara"),
        }
        return any(child.tag in protected_tags for child in element.iter())

    @staticmethod
    def get_top_level_blocks(doc: Document) -> List[Tuple[str, str, Any, str]]:
        """
        遍历文档 w:body 下的所有顶层块 (段落与表格)，返回:
        [(element_path, block_type, oxml_element, visible_text), ...]
        """
        blocks: List[Tuple[str, str, Any, str]] = []
        body_el = doc.element.body
        p_count = 0
        tbl_count = 0
        sdt_count = 0

        for child in body_el:
            if child.tag == qn("w:p"):
                p_count += 1
                path_str = f"/w:document/w:body/w:p[{p_count}]"
                texts = [t.text for t in child.findall(f".//{qn('w:t')}") if t.text]
                visible_text = "".join(texts).strip()
                blocks.append((path_str, "paragraph", child, visible_text))
            elif child.tag == qn("w:tbl"):
                tbl_count += 1
                path_str = f"/w:document/w:body/w:tbl[{tbl_count}]"
                texts = [t.text for t in child.findall(f".//{qn('w:t')}") if t.text]
                visible_text = "".join(texts).strip()
                blocks.append((path_str, "table", child, visible_text))
            elif child.tag == qn("w:sdt"):
                sdt_count += 1
                path_str = f"/w:document/w:body/w:sdt[{sdt_count}]"
                texts = [t.text for t in child.findall(f".//{qn('w:t')}") if t.text]
                visible_text = "".join(texts).strip()
                blocks.append((path_str, "structured_document_tag", child, visible_text))

        return blocks

    @staticmethod
    def validate_supported_objects(doc: Document) -> None:
        """在不要求覆盖范围的情况下拒绝未知顶层对象。"""
        supported_top_level = {
            qn("w:p"),
            qn("w:tbl"),
            qn("w:sdt"),
            qn("w:sectPr"),
        }
        for child in doc.element.body:
            if child.tag not in supported_top_level:
                raise UnsupportedObjectError(
                    f"源文档包含不支持的顶层对象: {child.tag}；"
                    "无法确定其选区关系闭包"
                )

    @classmethod
    def _protected_spans(
        cls,
        blocks: List[Tuple[str, str, Any, str]],
    ) -> List[Tuple[str, str, int, int]]:
        """返回跨顶层块的书签/复杂字段范围。范围为闭区间。"""
        bookmark_positions: Dict[str, List[int]] = {}
        bookmark_starts: Dict[str, int] = {}
        bookmark_ends: Dict[str, int] = {}
        field_stack: List[int] = []
        protected: List[Tuple[str, str, int, int]] = []

        for index, (_, _, element, _) in enumerate(blocks):
            for marker in element.iter(qn("w:bookmarkStart")):
                key = marker.get(qn("w:id")) or marker.get(qn("w:name")) or f"start@{index}"
                bookmark_positions.setdefault(key, []).append(index)
                bookmark_starts[key] = index
            for marker in element.iter(qn("w:bookmarkEnd")):
                key = marker.get(qn("w:id")) or f"end@{index}"
                bookmark_positions.setdefault(key, []).append(index)
                bookmark_ends[key] = index

            for marker in element.iter(qn("w:fldChar")):
                marker_type = marker.get(qn("w:fldCharType"))
                if marker_type == "begin":
                    field_stack.append(index)
                elif marker_type == "end":
                    if not field_stack:
                        protected.append(("field", "unmatched_end", index, index))
                    else:
                        start = field_stack.pop()
                        protected.append(("field", "complex_field", start, index))

        for index in field_stack:
            protected.append(("field", "unmatched_begin", index, index))

        for key, positions in bookmark_positions.items():
            if key not in bookmark_starts or key not in bookmark_ends:
                protected.append(("bookmark", key, min(positions), max(positions)))
                continue
            protected.append(("bookmark", key, bookmark_starts[key], bookmark_ends[key]))

        return [
            (kind, identifier, min(start, end), max(start, end))
            for kind, identifier, start, end in protected
            if start != end
        ]

    @classmethod
    def _validate_protected_boundaries(
        cls,
        blocks: List[Tuple[str, str, Any, str]],
        resolved_spans: Dict[str, Tuple[int, int]],
    ) -> None:
        """禁止任何被选区部分覆盖的跨块字段或书签进入导入。"""
        for kind, identifier, start, end in cls._protected_spans(blocks):
            protected_start, protected_end = start, end + 1
            for region_id, (region_start, region_end) in resolved_spans.items():
                intersects = region_start < protected_end and region_end > protected_start
                covers_all = region_start <= protected_start and region_end >= protected_end
                if intersects and not covers_all:
                    raise ProtectedBoundaryError(
                        f"选区 '{region_id}' 切断 {kind} '{identifier}'，"
                        f"跨越顶层块 [{start}, {end}]；必须整段导入或显式排除"
                    )

    @classmethod
    def validate_regions(
        cls,
        doc: Document,
        regions: Dict[str, SourceRegion],
    ) -> Dict[str, Tuple[int, int]]:
        """
        核验所有源选区合法性，返回每个选区的 [start_idx, end_idx) 索引切片。
        检测项：
        1. 边界必须为合法顶层块（禁止切碎表格单元格内部）；
        2. 起点必须在终点之前；
        3. 选区之间不得重叠；
        4. 全文档非空文字块必须被完全覆盖，若有未分配实质内容且未标记 exclude 则报错。
        """
        cls.validate_supported_objects(doc)

        blocks = cls.get_top_level_blocks(doc)
        path_to_idx = {b[0]: idx for idx, b in enumerate(blocks)}
        total_blocks = len(blocks)

        resolved_spans: Dict[str, Tuple[int, int]] = {}
        assigned_indices: Set[int] = set()

        for rid, reg in regions.items():
            # 1. 检查起点
            if reg.start in path_to_idx:
                start_idx = path_to_idx[reg.start]
            else:
                if "/w:tbl[" in reg.start and ("/w:tc[" in reg.start or "/w:tr[" in reg.start):
                    raise InvalidBoundaryError(f"选区 '{rid}' 起点位于表格内部，非法截断表格: {reg.start}")
                raise InvalidBoundaryError(f"选区 '{rid}' 起点路径在源文档顶层块中不存在: {reg.start}")

            # 2. 检查终点
            if reg.end == "end_of_document":
                end_idx = total_blocks
            elif reg.end in path_to_idx:
                end_idx = path_to_idx[reg.end]
            else:
                if "/w:tbl[" in reg.end and ("/w:tc[" in reg.end or "/w:tr[" in reg.end):
                    raise InvalidBoundaryError(f"选区 '{rid}' 终点位于表格内部，非法截断表格: {reg.end}")
                raise InvalidBoundaryError(f"选区 '{rid}' 终点路径在源文档顶层块中不存在: {reg.end}")

            if start_idx >= end_idx:
                raise InvalidBoundaryError(
                    f"选区 '{rid}' 起点索引 ({start_idx}) 必须严格小于终点索引 ({end_idx}): "
                    f"[{reg.start} -> {reg.end}]"
                )

            current_indices = set(range(start_idx, end_idx))

            # 3. 检查重叠
            overlap = assigned_indices.intersection(current_indices)
            if overlap:
                overlapping_paths = [blocks[i][0] for i in sorted(overlap)]
                raise SelectionOverlapError(
                    f"选区 '{rid}' 与其他选区存在重叠块: {', '.join(overlapping_paths[:3])}"
                )

            assigned_indices.update(current_indices)
            resolved_spans[rid] = (start_idx, end_idx)

        cls._validate_protected_boundaries(blocks, resolved_spans)

        # 4. 检查未分配的实质内容 (排除已被 exclude 的选区)
        unassigned_indices = sorted(set(range(total_blocks)) - assigned_indices)
        unassigned_substantive = [
            (idx, blocks[idx][0], blocks[idx][3])
            for idx in unassigned_indices
            if cls._is_substantive_element(blocks[idx][2], blocks[idx][3])
        ]

        if unassigned_substantive:
            first_idx, first_path, first_text = unassigned_substantive[0]
            summary = first_text[:40] + ("..." if len(first_text) > 40 else "")
            raise UnassignedContentError(
                f"源文档中存在未分配的实质内容（共 {len(unassigned_substantive)} 个非空段落/表格未覆盖）。"
                f"首个未分配块位于 {first_path}: '{summary}'。"
                f"为防静默丢弃正文，请在 source.regions 中覆盖该区域，或显式声明 exclude=true 选区将其排除。"
            )

        return resolved_spans

    @classmethod
    def slice_document_by_region(
        cls,
        source_doc: Document,
        span: Tuple[int, int],
    ) -> Document:
        """根据 ``[start_idx, end_idx)`` 导入完整关系闭包的文档副本。"""
        new_doc, _ = cls.slice_document_by_region_with_result(source_doc, span)
        return new_doc

    @classmethod
    def slice_document_by_region_with_result(
        cls,
        source_doc: Document,
        span: Tuple[int, int],
    ) -> Tuple[Document, Any]:
        """切片并返回 ``(Document, ImportResult)``。"""
        from .layout import apply_section_spec
        from .package_importer import (
            PackageImporter,
            validate_relationship_closure,
        )

        blocks = cls.get_top_level_blocks(source_doc)
        start_idx, end_idx = span
        if start_idx < 0 or end_idx > len(blocks) or start_idx >= end_idx:
            raise InvalidBoundaryError(
                f"非法文档切片范围 [{start_idx}, {end_idx})，"
                f"有效块数量为 {len(blocks)}"
            )

        new_doc = Document()
        body = new_doc.element.body
        for child in list(body):
            if child.tag != qn("w:sectPr"):
                body.remove(child)

        if source_doc.sections and new_doc.sections:
            apply_section_spec(
                new_doc.sections[0],
                source_section=source_doc.sections[0],
            )

        selected = [blocks[idx][2] for idx in range(start_idx, end_idx)]
        importer = PackageImporter(
            source=source_doc,
            target=new_doc,
            strict_rels=True,
            sanitize_objects=False,
            style_prefix="SynthImport_",
            part_prefix="synthRegion",
            footnote_ids={
                marker.get(qn("w:id"))
                for element in selected
                for marker in element.iter(qn("w:footnoteReference"))
                if marker.get(qn("w:id")) is not None
            },
            endnote_ids={
                marker.get(qn("w:id"))
                for element in selected
                for marker in element.iter(qn("w:endnoteReference"))
                if marker.get(qn("w:id")) is not None
            },
        )
        result = importer.import_elements_with_result(selected)
        sect_pr = body.find(qn("w:sectPr"))
        for source_block, imported in zip(
            blocks[start_idx:end_idx], result.elements
        ):
            if sect_pr is None:
                body.append(imported)
            else:
                body.insert(body.index(sect_pr), imported)
            result.node_map[source_block[0]] = cls._top_level_path(body, imported)

        validate_relationship_closure(new_doc)
        return new_doc, result

    @staticmethod
    def _top_level_path(body: Any, target: Any) -> str:
        """为已插入的顶层元素生成与 NodeRef 兼容的路径。"""
        counters = {qn("w:p"): 0, qn("w:tbl"): 0, qn("w:sdt"): 0}
        for child in body:
            if child.tag not in counters:
                continue
            counters[child.tag] += 1
            if child is target or child == target:
                prefix = {
                    qn("w:p"): "p",
                    qn("w:tbl"): "tbl",
                    qn("w:sdt"): "sdt",
                }[child.tag]
                return f"/w:document/w:body/w:{prefix}[{counters[child.tag]}]"
        raise InvalidBoundaryError("导入元素未能定位到目标文档顶层")
