# -*- coding: utf-8 -*-
"""
受控文档部件导入器单元测试 (tests/test_package_importer.py)
"""

import io
import unittest
from pathlib import Path
from PIL import Image

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Inches, Pt

from lib.package_importer import (
    PackageImporter,
    RelationshipClosureError,
    validate_relationship_closure,
)


def create_test_image_bytes(color="red", size=(30, 30)) -> bytes:
    stream = io.BytesIO()
    Image.new("RGB", size, color).save(stream, format="PNG")
    return stream.getvalue()


class PackageImporterTest(unittest.TestCase):

    def test_image_and_hyperlink_remapping(self):
        src_doc = Document()
        img_bytes = create_test_image_bytes("blue")
        src_doc.add_picture(io.BytesIO(img_bytes), width=Inches(1.0))
        
        # 添加带超链接的段落
        p = src_doc.add_paragraph("参考链接: ")
        link = parse_xml(f'<w:hyperlink {nsdecls("w", "r")}/>')
        rid = src_doc.part.relate_to("https://example.com/source", RT.HYPERLINK, is_external=True)
        link.set(qn("r:id"), rid)
        link.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>点击这里</w:t></w:r>'))
        p._p.append(link)

        dst_doc = Document()
        importer = PackageImporter(source=src_doc, target=dst_doc)

        # 导入正文
        imported_elements = importer.import_body_contents()
        self.assertEqual(len(imported_elements), 2)

        for el in imported_elements:
            dst_doc.element.body.append(el)

        # 验证目标文档中图片关系已建立且字节一致
        blips = list(dst_doc.element.body.iter(qn("a:blip")))
        self.assertTrue(len(blips) >= 1)
        new_embed_id = blips[0].get(qn("r:embed"))
        self.assertIn(new_embed_id, dst_doc.part.rels)
        self.assertEqual(dst_doc.part.rels[new_embed_id].target_part.blob, img_bytes)

        # 验证超链接关系重映射
        hyperlinks = list(dst_doc.element.body.iter(qn("w:hyperlink")))
        self.assertTrue(len(hyperlinks) >= 1)
        new_link_id = hyperlinks[0].get(qn("r:id"))
        self.assertIn(new_link_id, dst_doc.part.rels)
        self.assertEqual(dst_doc.part.rels[new_link_id].target_ref, "https://example.com/source")

    def test_import_result_records_mappings(self):
        src_doc = Document()
        src_doc.add_paragraph("可追溯段落")
        dst_doc = Document()

        importer = PackageImporter(source=src_doc, target=dst_doc)
        result = importer.import_body_contents_with_result()

        self.assertEqual(len(result.elements), 1)
        self.assertIn("Normal", result.style_map)
        self.assertIsInstance(result.to_dict(), dict)
        self.assertNotIn("elements", result.to_dict())

    def test_bookmark_name_collision_is_remapped_with_local_references(self):
        target = Document()
        target_p = target.add_paragraph("目标")
        target_p._p.append(parse_xml(
            f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="same"/>'
        ))
        target_p._p.append(parse_xml(
            f'<w:bookmarkEnd {nsdecls("w")} w:id="1"/>'
        ))

        source = Document()
        source_p = source.add_paragraph("源")
        source_p._p.append(parse_xml(
            f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="same"/>'
        ))
        source_p._p.append(parse_xml(
            f'<w:bookmarkEnd {nsdecls("w")} w:id="1"/>'
        ))
        link_p = source.add_paragraph()
        link = parse_xml(f'<w:hyperlink {nsdecls("w")}/>' )
        link.set(qn("w:anchor"), "same")
        link.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>跳转</w:t></w:r>'))
        link_p._p.append(link)

        importer = PackageImporter(source=source, target=target, strict_rels=True)
        result = importer.import_body_contents_with_result()
        for element in result.elements:
            target.element.body.insert(target.element.body.index(target.element.body[-1]), element)

        remapped = result.bookmark_name_map["same"]
        names = [b.get(qn("w:name")) for b in target.element.body.iter(qn("w:bookmarkStart"))]
        self.assertEqual(names.count("same"), 1)
        self.assertIn(remapped, names)
        anchors = [h.get(qn("w:anchor")) for h in target.element.body.iter(qn("w:hyperlink"))]
        self.assertEqual(anchors, [remapped])

    def test_relationship_type_mismatch_is_rejected(self):
        doc = Document()
        p = doc.add_paragraph()
        link = parse_xml(f'<w:hyperlink {nsdecls("w", "r")}/>' )
        rid = doc.part.get_or_add_image(io.BytesIO(create_test_image_bytes("black")))[0]
        link.set(qn("r:id"), rid)
        link.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>错误关系</w:t></w:r>'))
        p._p.append(link)

        with self.assertRaises(RelationshipClosureError):
            validate_relationship_closure(doc)

    def test_relationship_target_content_type_is_rejected(self):
        doc = Document()
        bad_part = Part(
            PackURI("/word/not-an-image.xml"),
            "text/xml",
            b"<root/>",
            doc.part.package,
        )
        rid = doc.part.relate_to(
            bad_part,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        )
        p = doc.add_paragraph()
        p._p.append(parse_xml(
            f'<w:r {nsdecls("w", "a", "r")}><a:blip r:embed="{rid}"/></w:r>'
        ))

        with self.assertRaises(RelationshipClosureError):
            validate_relationship_closure(doc)

    def test_bookmark_id_and_docpr_id_renumbering(self):
        dst_doc = Document()
        # 目标文档已存在书签 ID 1, 2 与 docPr id 1
        p_dst = dst_doc.add_paragraph("现有段落")
        bm1 = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="Bm1"/>')
        bm1_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="1"/>')
        bm2 = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="2" w:name="Bm2"/>')
        bm2_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="2"/>')
        p_dst._p.append(bm1)
        p_dst._p.append(bm1_end)
        p_dst._p.append(bm2)
        p_dst._p.append(bm2_end)

        src_doc = Document()
        p_src = src_doc.add_paragraph("源段落")
        src_bm1 = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="SourceBm"/>')
        src_bm1_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="1"/>')
        p_src._p.append(src_bm1)
        p_src._p.append(src_bm1_end)

        # 添加含 wp:docPr 的图片
        img_bytes = create_test_image_bytes("green")
        src_doc.add_picture(io.BytesIO(img_bytes), width=Inches(0.5))

        importer = PackageImporter(source=src_doc, target=dst_doc)
        imported = importer.import_body_contents()
        for el in imported:
            dst_doc.element.body.append(el)

        # 检查书签 ID 重新编号
        bm_ids = [b.get(qn("w:id")) for b in dst_doc.element.body.iter(qn("w:bookmarkStart"))]
        self.assertEqual(len(bm_ids), len(set(bm_ids)), "目标文档中所有书签 ID 必须唯一不冲突")
        self.assertIn("3", bm_ids, "导入的书签 ID 应顺延递增为 3")

    def test_native_numbering_preservation(self):
        src_doc = Document()
        # 添加带编号的段落
        p1 = src_doc.add_paragraph("列表项 1", style="List Number")
        p2 = src_doc.add_paragraph("列表项 2", style="List Number")

        dst_doc = Document()
        importer = PackageImporter(source=src_doc, target=dst_doc, preserve_native_numbering=True)
        imported = importer.import_body_contents()
        for el in imported:
            dst_doc.element.body.append(el)

        # 验证目标文档 numbering_part 包含克隆的 numbering 定义
        target_num_part = dst_doc.part.numbering_part
        nums = target_num_part.element.findall(qn("w:num"))
        self.assertTrue(len(nums) >= 1)

        # 验证导入段落的 numId 存在于目标 numbering_part
        target_num_ids = {n.get(qn("w:numId")) for n in nums}
        for el in imported:
            numPr = el.find(f".//{qn('w:numPr')}")
            if numPr is not None:
                numId_val = numPr.find(qn("w:numId")).get(qn("w:val"))
                self.assertIn(numId_val, target_num_ids)

    def test_style_isolation_no_collisions(self):
        dst_doc = Document()
        custom_dst = dst_doc.styles.add_style("SpecialStyle", WD_STYLE_TYPE.PARAGRAPH)
        custom_dst.font.name = "Arial"
        dst_p = dst_doc.add_paragraph("目标段落", style=custom_dst)

        src_doc = Document()
        custom_src = src_doc.styles.add_style("SpecialStyle", WD_STYLE_TYPE.PARAGRAPH)
        custom_src.font.name = "Courier New"
        src_p = src_doc.add_paragraph("源段落", style=custom_src)

        importer = PackageImporter(source=src_doc, target=dst_doc, style_prefix="SynthImport_")
        imported = importer.import_body_contents()
        dst_doc.element.body.append(imported[0])

        # 验证目标原有样式未被源样式冲掉
        self.assertEqual(dst_doc.styles["SpecialStyle"].font.name, "Arial")

        # 验证源样式被重命名为 SynthImport_0
        imported_pPr = imported[0].find(qn("w:pPr"))
        pStyle = imported_pPr.find(qn("w:pStyle")).get(qn("w:val"))
        self.assertNotEqual(pStyle, "SpecialStyle")
        self.assertTrue(pStyle.startswith("SynthImport_"))

    def test_ole_and_absolute_position_cleanup(self):
        src_doc = Document()
        p = src_doc.add_paragraph()
        # 构造带绝对定位与 OLE 的结构
        obj_xml = (
            '<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '  <v:shape style="position:absolute;margin-left:100pt;margin-top:50pt;width:200pt;height:100pt" o:ole="true">'
            '    <v:imagedata r:id="rIdDummy"/>'
            '  </v:shape>'
            '  <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12"/>'
            '</w:object>'
        )
        p._p.append(parse_xml(obj_xml))

        dst_doc = Document()
        importer = PackageImporter(source=src_doc, target=dst_doc)
        imported = importer.import_element(p._p)

        # 验证 OLEObject 已被剔除
        ns_o = "urn:schemas-microsoft-com:office:office"
        ns_v = "urn:schemas-microsoft-com:vml"
        self.assertEqual(list(imported.iter(f"{{{ns_o}}}OLEObject")), [])
        # 验证 w:object 已转换为静态 w:pict
        self.assertEqual(list(imported.iter(qn("w:object"))), [])
        picts = list(imported.iter(qn("w:pict")))
        self.assertEqual(len(picts), 1)
        # 验证 position:absolute 已清除
        for sh in picts[0].iter(f"{{{ns_v}}}shape"):
            style = sh.get("style", "")
            self.assertNotIn("position:absolute", style)
            self.assertNotIn("margin-left", style)


if __name__ == "__main__":
    unittest.main()
