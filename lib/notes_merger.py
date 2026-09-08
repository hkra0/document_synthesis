# -*- coding: utf-8 -*-
"""
跨文档脚注与尾注安全合并器 (lib/notes_merger.py)
实现 'notes.merge.v1' 核心能力：
1. 解决多源文档/多部件合并时 word/footnotes.xml 与 word/endnotes.xml 的冲突；
2. 严格保留 OOXML 规范的系统分隔符 (w:id="-1", w:type="separator") 与延续分隔符 (w:id="0", w:type="continuationSeparator")；
3. 为常规脚注/尾注重新分配连续且唯一的新整数 ID；
4. 克隆并挂载脚注/尾注内部的关系（如外部超链接与媒体）；
5. 提供 ID 映射表，支持正文中 w:footnoteReference 与 w:endnoteReference 的无缝重写。
"""

import copy
from pathlib import Path
from typing import Any, Dict, Optional, Set, Tuple

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from lxml import etree


FOOTNOTES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
ENDNOTES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
)

DEFAULT_FOOTNOTES_XML = f"""<w:footnotes {nsdecls('w')}>
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
</w:footnotes>"""

DEFAULT_ENDNOTES_XML = f"""<w:endnotes {nsdecls('w')}>
  <w:endnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:endnote>
</w:endnotes>"""


class NotesMerger:
    """负责跨文档合并 footnotes 与 endnotes 部件并生成 ID 重映射字典"""

    @classmethod
    def _clone_related_part(
        cls,
        source_part: Any,
        target_doc: Any,
        cache: Dict[Any, Any],
    ) -> Any:
        """将脚注内部关系的目标部件递归克隆进目标 OPC 包。"""
        if source_part in cache:
            return cache[source_part]

        used_names = {str(part.partname) for part in target_doc.part.package.iter_parts()}
        suffix = Path(str(source_part.partname)).suffix
        index = len(used_names)
        part_name = f"/word/synthNote{index}{suffix}"
        while part_name in used_names:
            index += 1
            part_name = f"/word/synthNote{index}{suffix}"

        cloned = type(source_part).load(
            PackURI(part_name),
            source_part.content_type,
            source_part.blob,
            target_doc.part.package,
        )
        cache[source_part] = cloned

        for rel in source_part.rels.values():
            if rel.is_external:
                target = rel.target_ref
            else:
                target = cls._clone_related_part(rel.target_part, target_doc, cache)
            cloned.rels.add_relationship(rel.reltype, target, rel.rId, rel.is_external)
        return cloned

    @staticmethod
    def get_notes_part(doc: Any, rel_type: str) -> Optional[Any]:
        """安全检索与 document.xml 关联的指定笔记部件"""
        try:
            return doc.part.part_related_by(rel_type)
        except KeyError:
            return None

    @classmethod
    def get_or_create_notes_part(
        cls,
        doc: Any,
        is_endnote: bool = False,
    ) -> Tuple[Any, etree._Element]:
        """获取或创建目标文档的笔记部件及根 XML 元素"""
        rel_type = RT.ENDNOTES if is_endnote else RT.FOOTNOTES
        content_type = ENDNOTES_CONTENT_TYPE if is_endnote else FOOTNOTES_CONTENT_TYPE
        default_xml = DEFAULT_ENDNOTES_XML if is_endnote else DEFAULT_FOOTNOTES_XML
        part_name = "/word/endnotes.xml" if is_endnote else "/word/footnotes.xml"

        part = cls.get_notes_part(doc, rel_type)
        if part is None:
            part = Part(
                PackURI(part_name),
                content_type,
                default_xml.encode("utf-8"),
                doc.part.package,
            )
            doc.part.relate_to(part, rel_type)
            root = parse_xml(part.blob)
        else:
            root = parse_xml(part.blob)
        return part, root

    @classmethod
    def save_notes_part(cls, part: Any, root: etree._Element) -> None:
        """将修改后的 XML 树序列化并保存回部件 blob"""
        part._blob = etree.tostring(
            root, encoding="utf-8", xml_declaration=True, standalone="yes"
        )

    @classmethod
    def merge_notes(
        cls,
        source_doc: Any,
        target_doc: Any,
        is_endnote: bool = False,
        strict_rels: bool = False,
        include_ids: Optional[Set[str]] = None,
    ) -> Dict[str, str]:
        """
        合并 source_doc 的笔记到 target_doc 中。
        返回映射字典: {old_id_in_source: new_id_in_target}
        """
        rel_type = RT.ENDNOTES if is_endnote else RT.FOOTNOTES
        tag_name = "endnote" if is_endnote else "footnote"

        source_part = cls.get_notes_part(source_doc, rel_type)
        if source_part is None:
            return {}
        if include_ids is not None:
            include_ids = {str(value) for value in include_ids}
            if not include_ids:
                return {}

        source_root = parse_xml(source_part.blob)
        target_part, target_root = cls.get_or_create_notes_part(target_doc, is_endnote)
        related_parts: Dict[Any, Any] = {}

        # 查找目标文档中已有的常规正整数 ID
        existing_ids = []
        for note in target_root.findall(qn(f"w:{tag_name}")):
            nid_str = note.get(qn("w:id"), "")
            if nid_str.isdigit():
                existing_ids.append(int(nid_str))
        next_id = max(existing_ids, default=0) + 1

        mapping: Dict[str, str] = {}

        for note in source_root.findall(qn(f"w:{tag_name}")):
            nid_str = note.get(qn("w:id"), "")
            ntype = note.get(qn("w:type"), "")

            # 系统保留条目（分隔符与延续分隔符）跳过重新编号
            if ntype in ("separator", "continuationSeparator") or (
                nid_str.startswith("-") or nid_str == "0"
            ):
                continue
            if include_ids is not None and nid_str not in include_ids:
                continue

            # 为常规条目分配新 ID
            new_id = next_id
            next_id += 1
            mapping[nid_str] = str(new_id)

            cloned = copy.deepcopy(note)
            cloned.set(qn("w:id"), str(new_id))

            # 克隆内部可能存在的关系（如超链接与嵌入图片）
            for item in cloned.iter():
                for attr in (qn("r:id"), qn("r:embed"), qn("r:link")):
                    old_rid = item.get(attr)
                    if not old_rid:
                        continue
                    rel = source_part.rels.get(old_rid)
                    if rel is None:
                        if strict_rels:
                            raise ValueError(
                                f"{tag_name} 部件引用了不存在的关系 {old_rid}"
                            )
                        item.attrib.pop(attr, None)
                        continue
                    if rel.is_external:
                        new_rid = target_part.relate_to(
                            rel.target_ref, rel.reltype, is_external=True
                        )
                        item.set(attr, new_rid)
                    else:
                        # 内部媒体或 XML 部件必须进入目标 OPC 包，不能直接
                        # 把源包中的 part 挂到目标关系上。
                        cloned_part = cls._clone_related_part(
                            rel.target_part,
                            target_doc,
                            related_parts,
                        )
                        new_rid = target_part.relate_to(
                            cloned_part,
                            rel.reltype,
                            is_external=False,
                        )
                        item.set(attr, new_rid)

            target_root.append(cloned)

        cls.save_notes_part(target_part, target_root)
        return mapping
