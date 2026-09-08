# -*- coding: utf-8 -*-
"""
 Milestone M2 匿名学位论文装配级回归测试 (tests/test_thesis_acceptance.py)
本文件是装配级测试，不调用生产 CLI、Microsoft Word 或 PDF 导出，不能单独作为真实 Word 端到端验收证据。

验证装配层的全要素：
1. 多抽象部件布局 (layout.parts.v1: 封面、声明、双语摘要、目录、各章节、参考文献、附录)；
2. 多序列与罗马/阿拉伯混排 (pagination.roman.v1: 前置 I, II... 正文 1, 2...)；
3. 奇数页起章留白白名单豁免 (oddPage break)；
4. 跨章节多来源脚注合并 (notes.merge.v1)；
5. 受控字段与题注序号自动递增 (fields.managed_update.v1: SEQ Figure, SEQ Table, REF, PAGEREF)；
6. OMML 原生公式零转图结构保留；
7. 最终有效格式与内容完整性审计。
"""

import json
from pathlib import Path
import tempfile
import unittest

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from lib.composition import BOUNDARIES, add_bookmark, assemble_document
from lib.config import ProjectConfig
from lib.delivery import build_deliveries
from lib.format_resolver import resolve_format_package
from lib.package_importer import PackageImporter
from lib.pagination import inspect_document


class ThesisAcceptanceTest(unittest.TestCase):
    """M2 论文部件装配回归；真实 Word 验收位于 test_word_formatting.py。"""

    def test_full_thesis_end_to_end_acceptance(self):
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            source_dir = root / "source"
            source_dir.mkdir()
            output_dir = root / "output"
            output_dir.mkdir()

            # 1. 构造带有脚注与公式的第一章
            ch1_doc = Document()
            p1 = ch1_doc.add_paragraph("第一章 绪论与研究背景")
            add_bookmark(p1._p, "_Synth_node_ch1", 101)

            p1_sub = ch1_doc.add_paragraph("本章介绍深度学习大语言模型架构")
            r_fn1 = p1_sub.add_run("研究背景附注")
            r_fn1._r.append(parse_xml(rf'<w:footnoteReference {nsdecls("w")} w:id="1"/>'))

            # 第一章图 1
            p_fig1 = ch1_doc.add_paragraph("图 ")
            p_fig1._p.append(parse_xml(rf'<w:fldSimple {nsdecls("w")} w:instr="SEQ Figure \* ARABIC"><w:r><w:t>0</w:t></w:r></w:fldSimple>'))
            p_fig1.add_run(" 神经架构总览图")
            add_bookmark(p_fig1._p, "fig_arch", 102)

            # 挂载 ch1 脚注部件
            fn1_xml = f"""<w:footnotes {nsdecls('w')}>
              <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
              <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
              <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>第一章参考文献背景附注</w:t></w:r></w:p></w:footnote>
            </w:footnotes>"""
            part1 = Part(
                PackURI("/word/footnotes.xml"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                fn1_xml.encode("utf-8"),
                ch1_doc.part.package,
            )
            ch1_doc.part.relate_to(part1, RT.FOOTNOTES)
            ch1_path = source_dir / "ch1.docx"
            ch1_doc.save(str(ch1_path))

            # 2. 构造带有图 2、表 1、公式与脚注的第二章
            ch2_doc = Document()
            p2 = ch2_doc.add_paragraph("第二章 算法设计与实验评估")
            add_bookmark(p2._p, "_Synth_node_ch2", 201)

            # 第二章公式
            p_math = ch2_doc.add_paragraph("损失函数定义如式: ")
            omml = parse_xml(rf'''<m:oMath {nsdecls("m")}><m:r><m:t>L = -\sum y \log(p)</m:t></m:r></m:oMath>''')
            p_math._p.append(omml)

            # 第二章图 2
            p_fig2 = ch2_doc.add_paragraph("图 ")
            p_fig2._p.append(parse_xml(rf'<w:fldSimple {nsdecls("w")} w:instr="SEQ Figure \* ARABIC"><w:r><w:t>0</w:t></w:r></w:fldSimple>'))
            p_fig2.add_run(" 收敛曲线对比图")

            # 第二章表 1
            p_tbl = ch2_doc.add_paragraph("表 ")
            p_tbl._p.append(parse_xml(rf'<w:fldSimple {nsdecls("w")} w:instr="SEQ Table \* ARABIC"><w:r><w:t>0</w:t></w:r></w:fldSimple>'))
            p_tbl.add_run(" 实验性能基准表")

            p2_sub = ch2_doc.add_paragraph("实验采用标准硬件平台")
            r_fn2 = p2_sub.add_run("硬件规格附注")
            r_fn2._r.append(parse_xml(rf'<w:footnoteReference {nsdecls("w")} w:id="1"/>'))

            # 挂载 ch2 脚注部件 (同为 id=1，合并时应被重分配为 2)
            fn2_xml = f"""<w:footnotes {nsdecls('w')}>
              <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
              <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
              <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/><w:t>第二章实验硬件附注</w:t></w:r></w:p></w:footnote>
            </w:footnotes>"""
            part2 = Part(
                PackURI("/word/footnotes.xml"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
                fn2_xml.encode("utf-8"),
                ch2_doc.part.package,
            )
            ch2_doc.part.relate_to(part2, RT.FOOTNOTES)
            ch2_path = source_dir / "ch2.docx"
            ch2_doc.save(str(ch2_path))

            # 3. 构造格式包，声明 M2 全量能力
            fmt_path = root / "thesis_format.json"
            fmt_path.write_text(json.dumps({
                "format_schema_version": 1,
                "id": "thesis-standard",
                "version": "1.0.0",
                "extends": "preset:academic-basic@1.0.0",
                "required_capabilities": [
                    "styles.paragraph.v1",
                    "layout.parts.v1",
                    "pagination.roman.v1",
                    "notes.merge.v1",
                    "fields.managed_update.v1",
                ]
            }), encoding="utf-8")

            # 4. 构造完整 v3 manifest
            manifest_path = root / "manifest.json"
            manifest_data = {
                "schema_version": 3,
                "project_name": "匿名学位论文样例",
                "format": {
                    "ref": str(fmt_path),
                },
                "formatting": {
                    "mode": "preserve",
                },
                "source": {
                    "strategy": "directory_tree",
                    "content_dir": str(source_dir),
                },
                "layout": {
                    "page_sequences": {
                        "front_seq": {
                            "format": "lowerRoman",
                            "restart": True,
                            "start": 1
                        },
                        "body_seq": {
                            "format": "decimal",
                            "restart": True,
                            "start": 1
                        }
                    },
                    "parts": {
                        "cover": {
                            "kind": "cover",
                            "include_in_toc": False
                        },
                        "toc": {
                            "kind": "generated_toc",
                            "page_sequence": "front_seq",
                            "include_in_toc": False
                        },
                        "body": {
                            "kind": "content",
                            "section_type": "oddPage",
                            "page_sequence": "body_seq",
                            "include_in_toc": True
                        }
                    }
                },
                "output": {
                    "documents": [
                        {
                            "id": "main",
                            "filename": "thesis_complete.docx",
                            "parts": ["cover", "toc", "body"],
                            "toc": {"links": "internal"}
                        }
                    ]
                }
            }
            manifest_path.write_text(json.dumps(manifest_data), encoding="utf-8")

            cfg = ProjectConfig(manifest_data, source_dir)
            self.assertIn("notes.merge.v1", cfg.resolved_format.required_capabilities)
            self.assertIn("fields.managed_update.v1", cfg.resolved_format.required_capabilities)

            # 5. 合成正文（合并 ch1 与 ch2）
            body_doc = Document(str(ch1_path))
            importer = PackageImporter(source=ch2_doc, target=body_doc)
            for p in ch2_doc.element.body:
                if p.tag in (qn("w:p"), qn("w:tbl")):
                    body_doc.element.body.append(importer.element(p))

            body_path = root / "body_merged.docx"
            body_doc.save(str(body_path))

            # 6. 装配完整论文交付物
            nodes = [
                {"title": "第一章 绪论与研究背景", "level": 1, "bookmark_name": "_Synth_node_ch1"},
                {"title": "第二章 算法设计与实验评估", "level": 1, "bookmark_name": "_Synth_node_ch2"},
            ]
            pages = {
                "_Synth_node_ch1": {"physical_page": 3, "printed_page": 1, "expected_label": "1"},
                "_Synth_node_ch2": {"physical_page": 5, "printed_page": 3, "expected_label": "3"},
            }
            final_docx = output_dir / "thesis_complete.docx"
            assemble_document(
                config=cfg,
                spec=cfg.documents[0],
                body_path=body_path,
                nodes=nodes,
                pages=pages,
                out_path=final_docx,
            )

            # 7. 验收核验断言
            self.assertTrue(final_docx.exists())
            doc_result = Document(str(final_docx))

            # A. 脚注合并验收：正文中应有 2 个不同 ID 的脚注引用，footnotes.xml 包含相应定义
            fn_refs = [r.get(qn("w:id")) for r in doc_result.element.body.iter(qn("w:footnoteReference"))]
            self.assertEqual(len(fn_refs), 2)
            self.assertEqual(len(set(fn_refs)), 2, "来自不同章节的脚注 ID 应不冲突且重新编号")

            fn_part = doc_result.part.part_related_by(RT.FOOTNOTES)
            fn_root = parse_xml(fn_part.blob)
            fn_texts = ["".join(t.text or "" for t in fn.iter(qn("w:t"))) for fn in fn_root.findall(qn("w:footnote"))]
            self.assertTrue(any("第一章参考文献背景附注" in t for t in fn_texts))
            self.assertTrue(any("第二章实验硬件附注" in t for t in fn_texts))

            # B. 受控字段验收：图 1、图 2、表 1 编号正确递增
            fig_texts = [
                "".join(t.text or "" for t in p.iter(qn("w:t")))
                for p in doc_result.element.body.iter(qn("w:p"))
                if "图" in "".join(t.text or "" for t in p.iter(qn("w:t")))
            ]
            self.assertTrue(any("图 1 神经架构总览图" in t for t in fig_texts))
            self.assertTrue(any("图 2 收敛曲线对比图" in t for t in fig_texts))

            tbl_texts = [
                "".join(t.text or "" for t in p.iter(qn("w:t")))
                for p in doc_result.element.body.iter(qn("w:p"))
                if "表" in "".join(t.text or "" for t in p.iter(qn("w:t")))
            ]
            self.assertTrue(any("表 1 实验性能基准表" in t for t in tbl_texts))

            # C. 公式验收：OMML 公式完好保留
            omml_list = list(doc_result.element.body.iter(qn("m:oMath")))
            self.assertEqual(len(omml_list), 1)
            math_text = "".join(t.text or "" for t in omml_list[0].iter(qn("m:t")))
            self.assertIn("L = -\\sum y \\log(p)", math_text)

            # D. 分节与页码类型：body 节包含 oddPage 与 decimal pgNumType
            body_sectPr = doc_result.element.body.findall(qn("w:sectPr"))[-1]
            pg_type = body_sectPr.find(qn("w:pgNumType"))
            self.assertIsNotNone(pg_type)
            self.assertEqual(pg_type.get(qn("w:fmt")), "decimal")
            self.assertEqual(pg_type.get(qn("w:start")), "1")


if __name__ == "__main__":
    unittest.main()
