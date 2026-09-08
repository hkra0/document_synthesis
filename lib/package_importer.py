# -*- coding: utf-8 -*-
"""
受控文档部件导入器 (lib/package_importer.py)
提供通用的 PackageImporter 类，负责跨 DOCX 文档/部件安全导入段落、表格与分节。

核心能力：
1. 样式隔离与重命名，避免污染目标文档 styles.xml 或覆盖 Synth* 样式；
2. 关系与媒体资源（图片、超链接）深拷贝与 rId 重映射；
3. 书签 ID (w:bookmarkStart/End) 与绘图 ID (wp:docPr) 重新编号防冲突；
4. 原生列表编号 (w:numPr / abstractNum / num) 映射与保留；
5. OLE 对象与绝对定位安全转换。
"""

import copy
import io
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union

from docx import Document
from docx.parts.numbering import NumberingPart
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from .layout import RenderContext, isolate_section_boundaries
from .notes_merger import NotesMerger


@dataclass
class ImportResult:
    """导入操作的可追溯结果与所有 ID 映射。

    ``elements`` 保留本次批量导入的 XML 元素；其余映射使用源包中的
    ID/NodeRef 作为键、目标包中的值作为值。导入器不会把运行时
    ``Document`` 句柄放进此结构，因此它可以安全地挂到 BuildPlan 的
    诊断或后续完整性检查上。
    """

    elements: List[Any] = field(default_factory=list)
    node_map: Dict[str, str] = field(default_factory=dict)
    bookmark_map: Dict[str, str] = field(default_factory=dict)
    bookmark_name_map: Dict[str, str] = field(default_factory=dict)
    relationship_map: Dict[str, str] = field(default_factory=dict)
    style_map: Dict[str, str] = field(default_factory=dict)
    numbering_map: Dict[str, str] = field(default_factory=dict)
    abstract_numbering_map: Dict[str, str] = field(default_factory=dict)
    footnote_map: Dict[str, str] = field(default_factory=dict)
    endnote_map: Dict[str, str] = field(default_factory=dict)
    diagnostics: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """返回不含 XML 元素的可序列化摘要。"""
        data = asdict(self)
        data.pop("elements", None)
        return data


class RelationshipClosureError(ValueError):
    """导入后发现缺失或类型不匹配的 OOXML 关系。"""

    code = "RELATIONSHIP_CLOSURE"


class PackageImporter:
    """通用跨文档部件导入器"""

    def __init__(
        self,
        source: Any,
        target: Any,
        render_context: Optional[RenderContext] = None,
        style_prefix: str = "SynthImport_",
        part_prefix: str = "synthPart",
        strict_rels: bool = False,
        preserve_native_numbering: bool = True,
        sanitize_objects: bool = True,
        footnote_ids: Optional[Set[str]] = None,
        endnote_ids: Optional[Set[str]] = None,
    ):
        self.source = Document(str(source)) if isinstance(source, (str, Path)) else source
        self.target = Document(str(target)) if isinstance(target, (str, Path)) else target
        self.render_context = render_context
        self.style_prefix = style_prefix
        self.part_prefix = part_prefix
        self.strict_rels = strict_rels
        self.preserve_native_numbering = preserve_native_numbering
        self.sanitize_objects = sanitize_objects

        self.import_result = ImportResult()
        self._relationship_cache: Dict[str, str] = {}
        self.bookmark_name_map: Dict[str, str] = {}

        self.parts: Dict[Any, Any] = {}
        self.used_partnames: Set[str] = {
            str(p.partname) for p in self.target.part.package.iter_parts()
        }

        self.style_ids: Dict[str, str] = {}
        self.default_styles: Dict[str, str] = {}
        self._collect_style_ids()

        self.num_ids: Dict[str, str] = {}
        self.abstract_map: Dict[str, str] = {}
        if self.preserve_native_numbering:
            self._init_numbering()

        self._clone_styles()
        self.footnote_ids: Dict[str, str] = NotesMerger.merge_notes(
            self.source,
            self.target,
            is_endnote=False,
            strict_rels=self.strict_rels,
            include_ids=footnote_ids,
        )
        self.endnote_ids: Dict[str, str] = NotesMerger.merge_notes(
            self.source,
            self.target,
            is_endnote=True,
            strict_rels=self.strict_rels,
            include_ids=endnote_ids,
        )
        self.import_result.style_map.update(self.style_ids)
        self.import_result.numbering_map.update(self.num_ids)
        self.import_result.abstract_numbering_map.update(self.abstract_map)
        self.import_result.footnote_map.update(self.footnote_ids)
        self.import_result.endnote_map.update(self.endnote_ids)

        # 书签 ID 重编号状态
        self.bookmark_ids: Dict[str, str] = {}
        self.next_bookmark_id = max(
            (
                int(e.get(qn("w:id")))
                for e in self.target.element.body.iter(qn("w:bookmarkStart"))
                if (e.get(qn("w:id")) or "").isdigit()
            ),
            default=0,
        ) + 1
        self._prepare_bookmark_name_map()

        # 绘图 docPr ID 重编号状态
        self.next_doc_pr_id = max(
            (
                int(e.get("id"))
                for e in self.target.element.body.iter(qn("wp:docPr"))
                if (e.get("id") or "").isdigit()
            ),
            default=0,
        ) + 1

    def _collect_style_ids(self) -> None:
        """预先收集所有源样式 ID 映射与默认样式"""
        used_styles = {
            s.get(qn("w:styleId")) for s in self.target.part.styles.element
        }

        for style in self.source.part.styles.element.findall(qn("w:style")):
            sid = style.get(qn("w:styleId"))
            if not sid:
                continue
            if sid not in used_styles:
                name = sid
            else:
                index = len(self.style_ids)
                name = f"{self.style_prefix}{index}"
                while name in used_styles:
                    index += 1
                    name = f"{self.style_prefix}{index}"
            used_styles.add(name)
            self.style_ids[sid] = name

            if style.get(qn("w:default")) in {"1", "true"}:
                stype = style.get(qn("w:type"))
                if stype:
                    self.default_styles[stype] = sid

    def _prepare_bookmark_name_map(self) -> None:
        """为跨来源重复书签名分配确定性的目标名称。"""
        used_names = {
            start.get(qn("w:name"))
            for start in self.target.element.body.iter(qn("w:bookmarkStart"))
            if start.get(qn("w:name"))
        }
        source_names = [
            start.get(qn("w:name"))
            for start in self.source.element.body.iter(qn("w:bookmarkStart"))
            if start.get(qn("w:name"))
        ]
        if len(source_names) != len(set(source_names)):
            raise RelationshipClosureError("源文档包含重复书签名称，无法安全导入")

        for old_name in source_names:
            new_name = old_name
            if new_name in used_names:
                base = re.sub(r"[^A-Za-z0-9_]+", "_", f"{self.style_prefix}{old_name}")
                new_name = base or "SynthBookmark"
                index = 2
                while new_name in used_names:
                    new_name = f"{base}_{index}"
                    index += 1
            used_names.add(new_name)
            self.bookmark_name_map[old_name] = new_name
        self.import_result.bookmark_name_map.update(self.bookmark_name_map)

    def _clone_styles(self) -> None:
        """克隆源样式定义至目标文档，完成 styleId 和 numId 的重写"""
        for style in self.source.part.styles.element:
            if style.tag != qn("w:style"):
                continue
            orig_id = style.get(qn("w:styleId"))
            if not orig_id or orig_id not in self.style_ids:
                continue

            cloned = copy.deepcopy(style)
            cloned.set(qn("w:styleId"), self.style_ids[orig_id])
            cloned.attrib.pop(qn("w:default"), None)

            name = cloned.find(qn("w:name"))
            if name is not None:
                name.set(qn("w:val"), self.style_ids[orig_id])

            aliases = cloned.find(qn("w:aliases"))
            if aliases is not None:
                cloned.remove(aliases)

            self._rewrite_element_styles(cloned)
            self.target.part.styles.element.append(cloned)

    def _init_numbering(self) -> None:
        """将源文档的 numbering.xml 定义（abstractNum 与 num）安全合并到目标文档"""
        try:
            source_num_part = self.source.part.numbering_part
        except (KeyError, AttributeError, NotImplementedError):
            return

        try:
            target_num_part = self.target.part.numbering_part
        except (KeyError, AttributeError, NotImplementedError):
            # python-docx 1.1.x 无法通过 NumberingPart.new() 创建空部件；
            # 这里显式挂载一个空 numbering.xml，再导入源定义，避免丢失
            # 原生编号或退回复制全文。
            empty_root = etree.fromstring(
                b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            )
            target_num_part = NumberingPart(
                PackURI("/word/numbering.xml"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                empty_root,
                self.target.part.package,
            )
            self.target.part.relate_to(target_num_part, RT.NUMBERING)

        target_numbering = target_num_part.element
        source_numbering = source_num_part.element

        for tag, attr, mapping in (
            ("abstractNum", "abstractNumId", self.abstract_map),
            ("num", "numId", self.num_ids),
        ):
            existing_ids = [
                int(e.get(qn(f"w:{attr}")))
                for e in target_numbering.findall(qn(f"w:{tag}"))
                if (e.get(qn(f"w:{attr}")) or "").isdigit()
            ]
            next_id = max(existing_ids, default=0) + 1

            for element in source_numbering.findall(qn(f"w:{tag}")):
                orig_val = element.get(qn(f"w:{attr}"))
                if not orig_val:
                    continue
                cloned = copy.deepcopy(element)
                mapping[orig_val] = str(next_id)
                cloned.set(qn(f"w:{attr}"), str(next_id))
                next_id += 1

                for ref in cloned.iter(qn("w:abstractNumId")):
                    val = ref.get(qn("w:val"))
                    if val in self.abstract_map:
                        ref.set(qn("w:val"), self.abstract_map[val])

                self._rewrite_element_styles(cloned)
                target_numbering.append(cloned)

    def _rewrite_element_styles(self, element: Any) -> None:
        """重写元素中引用的 styleId 和 numId"""
        for p in element.iter(qn("w:p")):
            if p.find("./" + qn("w:pPr") + "/" + qn("w:pStyle")) is None:
                default = self.default_styles.get("paragraph")
                if default and hasattr(p, "get_or_add_pPr"):
                    p.get_or_add_pPr().get_or_add_pStyle().val = default

        for tbl in element.iter(qn("w:tbl")):
            if tbl.find("./" + qn("w:tblPr") + "/" + qn("w:tblStyle")) is None:
                default = self.default_styles.get("table")
                if default and hasattr(tbl, "tblPr"):
                    tbl.tblPr.get_or_add_tblStyle().val = default

        for item in element.iter():
            if item.tag in {
                qn(f"w:{tag}")
                for tag in (
                    "pStyle",
                    "rStyle",
                    "tblStyle",
                    "basedOn",
                    "next",
                    "link",
                    "styleLink",
                    "numStyleLink",
                )
            }:
                value = item.get(qn("w:val"))
                if value in self.style_ids:
                    item.set(qn("w:val"), self.style_ids[value])
            if item.tag == qn("w:numId"):
                value = item.get(qn("w:val"))
                if value in self.num_ids:
                    item.set(qn("w:val"), self.num_ids[value])

    def _part(self, source_part: Any) -> Any:
        """递归克隆并登记外部部件，分配唯一的部件包路径"""
        if source_part in self.parts:
            return self.parts[source_part]

        index = len(self.used_partnames)
        suffix = Path(str(source_part.partname)).suffix
        name = f"/word/{self.part_prefix}{index}{suffix}"
        while name in self.used_partnames:
            index += 1
            name = f"/word/{self.part_prefix}{index}{suffix}"
        self.used_partnames.add(name)

        cloned = type(source_part).load(
            PackURI(name),
            source_part.content_type,
            source_part.blob,
            self.target.part.package,
        )
        self.parts[source_part] = cloned

        if hasattr(cloned, "element"):
            self._rewrite_element_styles(cloned.element)

        for rel in source_part.rels.values():
            cloned.rels.add_relationship(
                rel.reltype,
                rel.target_ref if rel.is_external else self._part(rel.target_part),
                rel.rId,
                rel.is_external,
            )
        return cloned

    def _sanitize_ole_and_shapes(self, element: Any) -> None:
        """清理 OLE 对象，转换静态图片，并移除绝对定位"""
        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns_o = "urn:schemas-microsoft-com:office:office"
        ns_v = "urn:schemas-microsoft-com:vml"

        # 1. 将包含 v:imagedata 的 w:object 转为静态 w:pict
        for obj in list(element.findall(f".//{{{ns_w}}}object")):
            shapes = obj.findall(f".//{{{ns_v}}}shape")
            if shapes:
                pict = OxmlElement("w:pict")
                for sh in shapes:
                    sh.attrib.pop(f"{{{ns_o}}}ole", None)
                    sh.attrib.pop("ole", None)
                    style = sh.get("style", "")
                    if "position:absolute" in style:
                        cleaned = ";".join(
                            [
                                s
                                for s in style.split(";")
                                if not any(
                                    s.strip().startswith(k)
                                    for k in [
                                        "position",
                                        "margin-left",
                                        "margin-top",
                                        "z-index",
                                        "left",
                                        "top",
                                    ]
                                )
                            ]
                        )
                        sh.set("style", cleaned)
                    pict.append(sh)
                parent = obj.getparent()
                if parent is not None:
                    parent.replace(obj, pict)
            else:
                parent = obj.getparent()
                if parent is not None:
                    parent.remove(obj)

        # 2. 移除残留的 OLEObject
        for ole in list(element.findall(f".//{{{ns_o}}}OLEObject")):
            parent = ole.getparent()
            if parent is not None:
                parent.remove(ole)

        # 3. 将 v:textbox 内的图片转为普通内联图片
        for tb in list(element.findall(f".//{{{ns_v}}}textbox")):
            inner_shapes = tb.findall(f".//{{{ns_v}}}shape")
            tb_shape = tb.getparent()
            if tb_shape is not None and inner_shapes:
                pict_parent = tb_shape.getparent()
                if pict_parent is not None:
                    run_parent = pict_parent.getparent()
                    if run_parent is not None:
                        for ish in inner_shapes:
                            style = ish.get("style", "")
                            cleaned = ";".join(
                                [
                                    s
                                    for s in style.split(";")
                                    if not any(
                                        s.strip().startswith(k)
                                        for k in [
                                            "position",
                                            "margin-left",
                                            "margin-top",
                                            "z-index",
                                            "left",
                                            "top",
                                        ]
                                    )
                                ]
                            )
                            ish.set("style", cleaned)
                            new_pict = OxmlElement("w:pict")
                            new_pict.append(ish)
                            run_parent.append(new_pict)
                        run_parent.remove(pict_parent)

        # 4. 清除所有残留 v:shape 的 position:absolute
        for sh in element.findall(f".//{{{ns_v}}}shape"):
            style = sh.get("style", "")
            if "position:absolute" in style:
                cleaned = ";".join(
                    [
                        s
                        for s in style.split(";")
                        if not any(
                            s.strip().startswith(k)
                            for k in [
                                "position",
                                "margin-left",
                                "margin-top",
                                "z-index",
                                "left",
                                "top",
                            ]
                        )
                    ]
                )
                sh.set("style", cleaned)

    def import_element(self, original: Any) -> Any:
        """深拷贝并导入单个 XML 元素（如段落 w:p、表格 w:tbl），重映射全部关系与 ID"""
        cloned = copy.deepcopy(original)
        self._rewrite_element_styles(cloned)
        if self.sanitize_objects:
            self._sanitize_ole_and_shapes(cloned)

        # 1. 书签 ID 重映射 (保证连续唯一，防止目标文档书签冲突)
        for b_start in cloned.iter(qn("w:bookmarkStart")):
            old_id = b_start.get(qn("w:id"))
            if old_id is not None:
                if old_id not in self.bookmark_ids:
                    self.bookmark_ids[old_id] = str(self.next_bookmark_id)
                    self.next_bookmark_id += 1
                b_start.set(qn("w:id"), self.bookmark_ids[old_id])

        for b_end in cloned.iter(qn("w:bookmarkEnd")):
            old_id = b_end.get(qn("w:id"))
            if old_id is not None and old_id in self.bookmark_ids:
                b_end.set(qn("w:id"), self.bookmark_ids[old_id])

        for item in cloned.iter():
            if item.tag == qn("w:bookmarkStart"):
                old_name = item.get(qn("w:name"))
                if old_name in self.bookmark_name_map:
                    item.set(qn("w:name"), self.bookmark_name_map[old_name])
            elif item.tag == qn("w:hyperlink"):
                old_name = item.get(qn("w:anchor"))
                if old_name in self.bookmark_name_map:
                    item.set(qn("w:anchor"), self.bookmark_name_map[old_name])
            elif item.tag == qn("w:instrText") and item.text:
                for old_name, new_name in self.bookmark_name_map.items():
                    item.text = re.sub(
                        rf"(\b(?:REF|PAGEREF|STYLEREF)\s+){re.escape(old_name)}\b",
                        rf"\g<1>{new_name}",
                        item.text,
                        flags=re.IGNORECASE,
                    )

        # 2. 绘图 docPr ID 重映射
        for doc_pr in cloned.iter(qn("wp:docPr")):
            doc_pr.set("id", str(self.next_doc_pr_id))
            self.next_doc_pr_id += 1

        # 3. 关系与媒体资源重映射
        for item in cloned.iter():
            for attr in (qn("r:id"), qn("r:embed"), qn("r:link")):
                old_rid = item.get(attr)
                if not old_rid:
                    continue
                rel = self.source.part.rels.get(old_rid)
                if rel is None:
                    if self.strict_rels:
                        raise ValueError(f"封面模板包含悬空关系: {old_rid}")
                    # 悬空无效关系，移除属性避免 Word 报错
                    item.attrib.pop(attr, None)
                    continue

                if old_rid in self._relationship_cache:
                    new_rid = self._relationship_cache[old_rid]
                elif rel.is_external:
                    new_rid = self.target.part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True
                    )
                elif "image" in rel.reltype:
                    new_rid, _ = self.target.part.get_or_add_image(
                        io.BytesIO(rel.target_part.blob)
                    )
                else:
                    cloned_part = self._part(rel.target_part)
                    new_rid = self.target.part.relate_to(
                        cloned_part, rel.reltype, is_external=False
                    )
                self._relationship_cache[old_rid] = new_rid
                self.import_result.relationship_map[old_rid] = new_rid
                item.set(attr, new_rid)

        # 4. 脚注与尾注引用 ID 重映射
        for fn_ref in cloned.iter(qn("w:footnoteReference")):
            old_fid = fn_ref.get(qn("w:id"))
            if old_fid is not None and old_fid in self.footnote_ids:
                fn_ref.set(qn("w:id"), self.footnote_ids[old_fid])
            elif self.strict_rels:
                raise RelationshipClosureError(
                    f"正文元素引用了未导入的脚注 {old_fid}"
                )

        for en_ref in cloned.iter(qn("w:endnoteReference")):
            old_eid = en_ref.get(qn("w:id"))
            if old_eid is not None and old_eid in self.endnote_ids:
                en_ref.set(qn("w:id"), self.endnote_ids[old_eid])
            elif self.strict_rels:
                raise RelationshipClosureError(
                    f"正文元素引用了未导入的尾注 {old_eid}"
                )

        self.import_result.bookmark_map.update(self.bookmark_ids)
        self.import_result.footnote_map.update(self.footnote_ids)
        self.import_result.endnote_map.update(self.endnote_ids)

        return cloned

    def element(self, original: Any) -> Any:
        """保持与旧 FrontImporter.element 完全一致的接口别名"""
        return self.import_element(original)

    def import_elements(self, elements: List[Any]) -> List[Any]:
        """批量导入元素"""
        return self.import_elements_with_result(elements).elements

    def import_elements_with_result(self, elements: List[Any]) -> ImportResult:
        """批量导入元素并返回可追溯的 ImportResult。"""
        self.import_result.elements.extend(self.import_element(el) for el in elements)
        return self.import_result

    def import_body_contents(self) -> List[Any]:
        """导入源文档所有正文段落与表格（保持先后顺序）"""
        return self.import_body_contents_with_result().elements

    def import_body_contents_with_result(self) -> ImportResult:
        """导入正文并返回包含元素及映射的 ImportResult。"""
        elements = [
            child
            for child in self.source.element.body
            if child.tag in (qn("w:p"), qn("w:tbl"), qn("w:sdt"))
        ]
        return self.import_elements_with_result(elements)


def validate_relationship_closure(document: Document) -> None:
    """验证 DOCX 包内关系的存在性、目标部件和关键关系类型。

    仅检查 XML 部件中的实际引用；外部超链接没有 ``target_part``，但仍
    必须有关系记录。对图片和超链接额外检查关系类型，避免出现 rId
    存在但指向页眉/脚注等错误部件的假闭包。
    """
    package = document.part.package
    parts = list(package.iter_parts())
    part_set = set(parts)
    attr_expectations = {
        (qn("a:blip"), qn("r:embed")): "image",
        (qn("a:blip"), qn("r:link")): "image",
        (qn("w:hyperlink"), qn("r:id")): "hyperlink",
    }

    for part in parts:
        content_type = str(getattr(part, "content_type", ""))
        if not (content_type.endswith("+xml") or content_type.endswith("/xml")):
            continue
        try:
            root = etree.fromstring(part.blob)
        except (etree.XMLSyntaxError, TypeError, ValueError):
            continue

        for element in root.iter():
            for attr in (qn("r:id"), qn("r:embed"), qn("r:link")):
                rid = element.get(attr)
                if not rid:
                    continue
                rel = part.rels.get(rid)
                if rel is None:
                    raise RelationshipClosureError(
                        f"部件 {part.partname} 的 {element.tag} 引用了不存在的关系 {rid}"
                    )
                if not rel.is_external:
                    try:
                        target_part = rel.target_part
                    except (KeyError, AttributeError) as exc:
                        raise RelationshipClosureError(
                            f"关系 {part.partname}:{rid} 无法解析目标部件"
                        ) from exc
                    if target_part not in part_set:
                        raise RelationshipClosureError(
                            f"关系 {part.partname}:{rid} 指向不属于当前包的部件"
                        )
                    expected_content_type = {
                        RT.IMAGE: lambda value: value.startswith("image/"),
                        RT.FOOTNOTES: lambda value: value
                        == "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                        RT.ENDNOTES: lambda value: value
                        == "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
                        RT.NUMBERING: lambda value: value
                        == "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                    }.get(rel.reltype)
                    if expected_content_type and not expected_content_type(
                        str(target_part.content_type)
                    ):
                        raise RelationshipClosureError(
                            f"关系 {part.partname}:{rid} 的目标内容类型 "
                            f"{target_part.content_type!r} 与 {rel.reltype!r} 不匹配"
                        )

                expected = attr_expectations.get((element.tag, attr))
                if expected and expected not in rel.reltype.lower():
                    raise RelationshipClosureError(
                        f"关系 {part.partname}:{rid} 类型 {rel.reltype!r} 与"
                        f" {element.tag} 的 {attr} 不匹配"
                    )
