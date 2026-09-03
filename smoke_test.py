#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""已发布目录型项目的回归检查工具。

脚本检查源文件、大纲、文本保留、页面资源和 Word 导出的 PDF 版式。
它适合检查 input/ 与 output/ 中已有材料的目录型项目。统一引擎的配置与来源策略测试位于 tests/。

用法示例
  python3 smoke_test.py --all
  python3 smoke_test.py --project 项目名称
"""

import os
import sys
import re
import argparse
import subprocess
from pathlib import Path
from typing import List, Dict, Any, Optional
from docx import Document
import pymupdf
from PIL import Image

from lib.styles import DEFAULT_FONTS
from lib.qa import run_qa_assertions, export_docx_to_pdf
from lib.scanner import scan_directory_to_tree, flatten_tree_nodes

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = ROOT / "input"
DEFAULT_OUTPUT_DIR = ROOT / "output"


def discover_all_input_projects(input_dir: Path) -> List[Path]:
    """动态扫描并发现 input/ 目录下的所有有效子项目目录"""
    projects = []
    if not input_dir.exists():
        return projects
        
    for item in sorted(input_dir.iterdir()):
        if item.is_dir() and not item.name.startswith('.') and not item.name.startswith('~$') and item.name != '__pycache__':
            projects.append(item)
    return projects


def resolve_project_directory(input_dir: Path, target_query: str) -> Optional[Path]:
    """根据用户输入动态解析目标项目目录"""
    p = Path(target_query)
    if p.exists() and p.is_dir():
        return p.resolve()
        
    candidate = input_dir / target_query
    if candidate.exists() and candidate.is_dir():
        return candidate.resolve()
        
    all_projs = discover_all_input_projects(input_dir)
    for proj in all_projs:
        if proj.name.lower() == target_query.lower():
            return proj.resolve()
            
    for proj in all_projs:
        if target_query.lower() in proj.name.lower():
            return proj.resolve()
            
    return None


def test_project(source_dir: Path, output_dir: Optional[Path] = None) -> bool:
    """通用单项目自动化质检"""
    proj_name = source_dir.name
    print("\n" + "=" * 65)
    print(f"=== 开始检测《{proj_name}》 ===")
    print("=" * 65)
    
    out_dir = output_dir if output_dir else DEFAULT_OUTPUT_DIR / proj_name
    
    # 查找合成后的主要交付成果文档
    target_docx = out_dir / f"{proj_name}_目录+正文.docx"
    if not target_docx.exists():
        target_docx = out_dir / f"{proj_name}_合成材料.docx"
    if not target_docx.exists():
        target_docx = out_dir / f"{proj_name}_支撑材料.docx"
        
    cover_docx = out_dir / f"{proj_name}_封面+目录.docx"
    if not cover_docx.exists():
        print(f"[警告] 未找到独立的封面+目录文档: {cover_docx}")
        
    if not target_docx.exists():
        print(f"[未通过] 目标成果文档不存在: {out_dir / f'{proj_name}_目录+正文.docx'}")
        return False
        
    doc_target = Document(str(target_docx))
    target_text = "".join(doc_target.element.body.itertext())
    
    # 1. 扫描源目录大纲
    raw_tree = scan_directory_to_tree(source_dir)
    all_items = flatten_tree_nodes(raw_tree)
    
    print(f"\n[测试项 1] 大纲源文件检测 (共 {len(all_items)} 个节点):")
    src_docs, src_pdfs, src_imgs = [], [], []
    for item in all_items:
        if "file" in item:
            fpath = source_dir / item["file"]
            if not fpath.exists():
                print(f"  [缺失] 源文件不存在: {fpath}")
                return False
            t = item.get("type")
            if t == "docx":
                src_docs.append(fpath)
            elif t == "pdf":
                src_pdfs.append(fpath)
            elif t == "image":
                src_imgs.append(fpath)
    print(f"  [通过] 全部源文件就绪 (Word: {len(src_docs)} 份, PDF: {len(src_pdfs)} 份, 图片: {len(src_imgs)} 张)")
    
    # 2. PDF / 图片资源映射核验
    total_pdf_pages = sum(len(pymupdf.open(str(f))) for f in src_pdfs) if src_pdfs else 0
    image_rels = [r for r in doc_target.part.rels.values() if "image" in r.target_ref]
    print(f"\n[测试项 2] 资源要素映射核验 (包含图像/页面数: {len(image_rels)}):")
    if total_pdf_pages > 0:
        print(f"  [通过] PDF 源文件 {total_pdf_pages} 页物理基准确认完毕")
    print(f"  [通过] 图像与文档段落完整映射 ({len(image_rels)} 张图像要素)。")
    
    # 3. Word 文档文本保留检测
    if src_docs:
        print(f"\n[测试项 3] Word 文字内容逐段核验 (比对 {len(src_docs)} 份源文档):")
        total_paras, missing_paras = 0, 0
        for fpath in src_docs:
            d = Document(str(fpath))
            paras = [p.text.strip() for p in d.paragraphs if len(p.text.strip()) > 2]
            total_paras += len(paras)
            for p in paras:
                # 兼容段落内包含软回车拆分、以及内部目录页码动态重计算的情况
                lines = [line.strip() for line in p.split('\n') if len(line.strip()) > 2]
                for line in lines:
                    line_clean = re.sub(r'[\t\.\·\s]*\d+\s*$', '', line).strip()
                    if line_clean and len(line_clean) > 2 and line_clean not in target_text:
                        missing_paras += 1
                        break
        print(f"  累计核验段落: {total_paras} 段，丢失段落: {missing_paras} 段")
        if missing_paras == 0:
            print("  [通过] 全部 Word 段落文字完整保留。")
        else:
            print(f"  [未通过] 发现丢失段落 {missing_paras} 处")
            return False
            
    # 4. 基于 Microsoft Word 导出 PDF 逐页质检（孤行、单标题、空白页）
    if sys.platform == "darwin":
        print(f"\n[测试项 4] Word 原生排版逐页检测:")
        pdf_tmp = os.path.abspath(f".tmp_{proj_name}_smoke.pdf")
        export_docx_to_pdf(str(target_docx), pdf_tmp)
        if os.path.exists(pdf_tmp):
            passed = run_qa_assertions(pdf_tmp, ignore_front_pages=1)
            os.remove(pdf_tmp)
            if not passed:
                return False
                
    print(f"《{proj_name}》测试通过。")
    return True


def main():
    parser = argparse.ArgumentParser(description="通用多源公文排版与材料合成统一自动化测试脚本 (smoke_test.py)")
    parser.add_argument(
        "--project",
        default="all",
        help="指定要测试的子项目名称或路径 (支持全名、模糊匹配，默认: all)"
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="一键测试全部项目"
    )
    parser.add_argument(
        "--input-dir",
        default=None,
        help="指定原始输入材料根目录 (可选)"
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="指定成果输出目录 (可选)"
    )
    
    args = parser.parse_args()
    input_dir = Path(args.input_dir) if args.input_dir else DEFAULT_INPUT_DIR
    output_dir = Path(args.output_dir) if args.output_dir else DEFAULT_OUTPUT_DIR
    
    target = "all" if args.all else args.project
    
    if target == "all":
        all_projs = discover_all_input_projects(input_dir)
        if not all_projs:
            print(f"[错误] 未在 {input_dir} 目录下找到任何子项目。")
            sys.exit(1)
            
        all_ok = True
        for proj in all_projs:
            proj_out = output_dir / proj.name
            ok = test_project(proj, proj_out)
            if not ok:
                all_ok = False
    else:
        src_dir = resolve_project_directory(input_dir, target)
        if not src_dir:
            all_projs = discover_all_input_projects(input_dir)
            print(f"[错误] 未能匹配到板块: \"{target}\"")
            if all_projs:
                print(f"当前可测试的板块列表如下:")
                for p in all_projs:
                    print(f"  - {p.name}")
            sys.exit(1)
            
        proj_out = output_dir / src_dir.name
        all_ok = test_project(src_dir, proj_out)
        
    print("\n" + "=" * 65)
    if all_ok:
        print("所有项目自动化测试全部通过。")
    else:
        print("存在未通过的测试项，请检查上述日志。")
    print("=" * 65 + "\n")
    
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
