# -*- coding: utf-8 -*-
"""
测试样本生成器 (tests/fixtures/formatting/generate_samples.py)
用于构造高风险 OOXML 语法特征（混合字体、跨段复杂字段、脚注、多级编号、分节）的匿名验证样本。
仅在测试或探针中使用，不写入业务目录。
"""

import os
from pathlib import Path
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls


def create_mixed_fonts_document(output_path: Path) -> Path:
    """生成中西文混合字体、基于样式的继承与直接格式覆盖样本"""
    doc = Document()
    
    # 1. 默认正文段落
    p1 = doc.add_paragraph()
    r1 = p1.add_run("这是中文正文，This is English text 12345.")
    r1.font.name = "Times New Roman"
    r1._r.get_or_add_rPr().set(qn("w:rFonts"), "")
    r1_fonts = r1._r.get_or_add_rPr().find(qn("w:rFonts"))
    r1_fonts.set(qn("w:ascii"), "Times New Roman")
    r1_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r1_fonts.set(qn("w:eastAsia"), "宋体")
    r1_fonts.set(qn("w:cs"), "Times New Roman")
    
    # 2. 直接格式覆盖粗斜体（包含显式 toggle false）
    p2 = doc.add_paragraph()
    r2_normal = p2.add_run("正常文本；")
    r2_bold = p2.add_run("加粗文本；")
    r2_bold.bold = True
    r2_italic = p2.add_run("斜体文本；")
    r2_italic.italic = True
    
    # 3. 显式 w:b w:val="0" (取消粗体测试)
    p3 = doc.add_paragraph()
    r3 = p3.add_run("显式取消加粗文本")
    r3_pr = r3._r.get_or_add_rPr()
    b_elem = parse_xml(r'<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="0"/>')
    r3_pr.append(b_elem)
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def create_complex_fields_document(output_path: Path) -> Path:
    """生成包含复杂跨段/跨 run 字段 (w:fldChar) 与简单字段 (w:fldSimple) 的样本"""
    doc = Document()
    
    # 1. 简单字段 (w:fldSimple) 如 AUTHOR
    p1 = doc.add_paragraph("文档作者简单域: ")
    fld_simple = parse_xml(r'<w:fldSimple xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:instr="AUTHOR"><w:r><w:t>Anonymous</w:t></w:r></w:fldSimple>')
    p1._p.append(fld_simple)
    
    # 2. 复杂字段 (w:fldChar: begin -> instrText -> separate -> result -> end)
    p2 = doc.add_paragraph("当前页码复杂域: ")
    # begin
    r_begin = p2.add_run()
    r_begin._r.append(parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="begin"/>'))
    # instrText
    r_instr = p2.add_run()
    instr_elem = parse_xml(r'<w:instrText xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve"> PAGE </w:instrText>')
    r_instr._r.append(instr_elem)
    # separate
    r_sep = p2.add_run()
    r_sep._r.append(parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="separate"/>'))
    # result run
    r_res = p2.add_run("1")
    # end
    r_end = p2.add_run()
    r_end._r.append(parse_xml(r'<w:fldChar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fldCharType="end"/>'))
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def create_sections_and_pagination_document(output_path: Path) -> Path:
    """生成分节、罗马页码前置部分与奇数页起章 (odd-page break) 样本"""
    doc = Document()
    
    # 第一节：前置部分，罗马数字编号 I, II
    sec1 = doc.sections[0]
    sec1.different_first_page_header_footer = False
    p1 = doc.add_paragraph("第一节：前置摘要 (预期页码 I)")
    
    # 配置 sec1 页码为 upperRoman，从 1 开始
    sectPr1 = sec1._sectPr
    pgNumType1 = parse_xml(r'<w:pgNumType xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fmt="upperRoman" w:start="1"/>')
    sectPr1.append(pgNumType1)
    
    # 添加第二节：正文部分，阿拉伯数字从 1 开始，奇数页起章
    sec2 = doc.add_section()
    sectPr2 = sec2._sectPr
    # 设置 type 为 oddPage
    type_elem = parse_xml(r'<w:type xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="oddPage"/>')
    sectPr2.append(type_elem)
    # 设置 pgNumType 为 decimal, start 1
    pgNumType2 = parse_xml(r'<w:pgNumType xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fmt="decimal" w:start="1"/>')
    sectPr2.append(pgNumType2)
    
    p2 = doc.add_paragraph("第二节：正文章节 (奇数页起章，预期页码 1)")
    
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
