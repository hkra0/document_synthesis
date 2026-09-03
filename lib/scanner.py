# -*- coding: utf-8 -*-
"""
文件系统大纲自发现扫描模块 (lib/scanner.py)
用于自动遍历 input 目录，按文件名数字序号与子文件夹层级构建标准大纲树。
支持零配置发现与文件名驱动。
"""

import os
import re
from pathlib import Path
from typing import List, Dict, Any, Optional

CN_NUMS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
           '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十']

SUPPORTED_EXTENSIONS = {
    '.docx': 'docx',
    '.doc': 'docx',
    '.pdf': 'pdf',
    '.pptx': 'pptx',
    '.ppt': 'pptx',
    '.jpg': 'image',
    '.jpeg': 'image',
    '.png': 'image'
}


def parse_sort_key_and_title(name: str):
    """
    解析文件名或目录名的序号与纯标题内容，并返回数字分段列表
    """
    stem = Path(name).stem if '.' in name else name
    stem = stem.strip()
    
    # 匹配中文数字一、二、三...
    m_cn = re.match(r'^([一二三四五六七八九十]+)[\s\、\.\：\:]\s*(.*)$', stem)
    if m_cn:
        cn_str = m_cn.group(1)
        raw_title = m_cn.group(2).strip()
        idx = CN_NUMS.index(cn_str) + 1 if cn_str in CN_NUMS else 99
        return (float(idx), raw_title, cn_str, [idx])
        
    # 匹配带括号序号 (1) 或 （1）
    m_bracket = re.match(r'^[\(（](\d+)[\)）]\s*(.*)$', stem)
    if m_bracket:
        num = int(m_bracket.group(1))
        raw_title = m_bracket.group(2).strip()
        return (float(num), raw_title, None, [num])

    # 匹配点分多级序号 如 2.1, 4.2, 5.1, 6.1
    m_dot = re.match(r'^(\d+(?:\.\d+)*)[\s\.\、\_\-]*\s*(.*)$', stem)
    if m_dot:
        num_str = m_dot.group(1)
        raw_title = m_dot.group(2).strip()
        parts = [int(p) for p in num_str.split('.')]
        if len(parts) == 1:
            sort_val = float(parts[0])
        elif len(parts) == 2:
            sort_val = float(f"{parts[0]}.{parts[1]:02d}")
        else:
            sort_val = float(f"{parts[0]}.{parts[1]:02d}{parts[2]:02d}")
        return (sort_val, raw_title or stem, None, parts)
        
    return (999.0, stem, None, [999])


def infer_group_parent_title(entries: List[Any]) -> str:
    """根据多级子文件名智能推断一级大纲小标题"""
    titles = [e[2] for e in entries]
    joined = ' '.join(titles)
    if '财务培训' in joined and '珠算' in joined:
        return '社会服务与技能培训'
    elif '培训' in joined or '服务' in joined:
        return '社会服务与技能培训'
    elif '调研' in joined or '报告' in joined:
        return '专业调研与分析报告'
    elif '培养方案' in joined or '人才培养' in joined:
        return '人才培养方案与实施'
    elif '制度' in joined or '办法' in joined:
        return '管理制度与工作办法'
    return '综合支撑材料'


def format_node_title(raw_title: str, level: int, seq_num: int, orig_name: str) -> str:
    """按层级标准格式化公文大纲标题"""
    if level == 1 and orig_name and re.match(r'^[一二三四五六七八九十]+、', orig_name):
        return Path(orig_name).stem
    if level == 2 and orig_name and re.match(r'^\d+\.\s*', orig_name):
        return Path(orig_name).stem
    if level == 3 and orig_name and re.match(r'^（\d+）', orig_name):
        return Path(orig_name).stem

    if level == 1:
        cn_idx = seq_num - 1
        cn_str = CN_NUMS[cn_idx] if 0 <= cn_idx < len(CN_NUMS) else str(seq_num)
        return f"{cn_str}、{raw_title}"
    elif level == 2:
        return f"{seq_num}. {raw_title}"
    elif level == 3:
        return f"（{seq_num}）{raw_title}"
    else:
        return raw_title


def scan_directory_to_tree(
    source_dir: Path,
    base_dir: Optional[Path] = None,
    level: int = 1,
    bm_counter: Optional[List[int]] = None
) -> List[Dict[str, Any]]:
    """
    递归扫描源目录生成多级大纲树节点，支持智能识别文件名点分多级标题（如 6.1, 6.2...）并自动归类
    """
    if base_dir is None:
        base_dir = source_dir
    if bm_counter is None:
        bm_counter = [0]
        
    if not source_dir.exists():
        return []

    raw_entries = []
    for item in source_dir.iterdir():
        if item.name.startswith('.') or item.name.startswith('~$') or item.name == '__pycache__':
            continue
        if item.name.endswith('.bak') or item.name.endswith('.tmp'):
            continue
        
        # 如果存在对应的 .docx 文件，则跳过旧的 .doc 格式
        if item.is_file() and item.suffix.lower() == '.doc':
            paired_docx = item.with_suffix('.docx')
            if paired_docx.exists():
                continue
        
        # 针对已转换的 PPTX 对应的同名 PDF 进行去重（优先使用 PPTX）
        if item.is_file() and item.suffix.lower() == '.pdf':
            paired_pptx = item.with_suffix('.pptx')
            if paired_pptx.exists():
                continue
                
        sort_key, raw_title, orig_cn, num_parts = parse_sort_key_and_title(item.name)
        raw_entries.append((sort_key, item, raw_title, orig_cn, num_parts))

    # 按自然数字序号排序
    raw_entries.sort(key=lambda x: (x[0], x[1].name))

    # 按主序号 (num_parts[0]) 聚合分组，自动识别小标题层级
    from collections import OrderedDict
    groups = OrderedDict()
    for e in raw_entries:
        major = e[4][0]
        if major not in groups:
            groups[major] = []
        groups[major].append(e)

    nodes = []
    group_idx = 1
    for major_key, g_entries in groups.items():
        # 判断当前组是单文件还是包含点分小标题的复合组
        has_sub_items = any(len(e[4]) > 1 for e in g_entries)
        
        if not has_sub_items and len(g_entries) == 1:
            # 标准单一节点
            sort_val, path_obj, raw_title, orig_cn, num_parts = g_entries[0]
            rel_path = path_obj.relative_to(base_dir).as_posix()
            bm_counter[0] += 1
            bm_name = f"_Toc_auto_{bm_counter[0]:03d}"
            
            if path_obj.is_dir():
                child_nodes = scan_directory_to_tree(
                    path_obj,
                    base_dir=base_dir,
                    level=level + 1,
                    bm_counter=bm_counter
                )
                title = format_node_title(raw_title, level, group_idx, path_obj.name)
                node = {
                    "level": level,
                    "title": title,
                    "toc_title": raw_title,
                    "type": "folder",
                    "folder": rel_path,
                    "bookmark_name": bm_name,
                    "bm_id": bm_counter[0],
                    "children": child_nodes
                }
                nodes.append(node)
                group_idx += 1
            else:
                ext = path_obj.suffix.lower()
                if ext not in SUPPORTED_EXTENSIONS:
                    continue
                node_type = SUPPORTED_EXTENSIONS[ext]
                title = format_node_title(raw_title, level, group_idx, path_obj.name)
                node = {
                    "level": level,
                    "title": title,
                    "toc_title": raw_title,
                    "type": node_type,
                    "file": rel_path,
                    "bookmark_name": bm_name,
                    "bm_id": bm_counter[0]
                }
                nodes.append(node)
                group_idx += 1
        else:
            # 复合小标题组（如 6.1, 6.2, 6.3, 6.4...）
            # 1. 检查是否有顶层 1 级主文件或主文件夹
            parent_entry = next((e for e in g_entries if len(e[4]) == 1), None)
            sub_entries = [e for e in g_entries if e != parent_entry]
            
            bm_counter[0] += 1
            parent_bm_name = f"_Toc_auto_{bm_counter[0]:03d}"
            
            if parent_entry is not None:
                p_sort_val, p_path_obj, p_raw_title, p_orig_cn, _ = parent_entry
                parent_raw_title = p_raw_title
                parent_rel_path = p_path_obj.relative_to(base_dir).as_posix()
            else:
                parent_raw_title = infer_group_parent_title(g_entries)
                parent_rel_path = f"group_{group_idx}"
                
            parent_title = format_node_title(parent_raw_title, level, group_idx, "")
            
            # 2. 构建下属二级子节点
            child_nodes = []
            for child_idx, (c_sort_val, c_path, c_raw_title, c_orig_cn, c_parts) in enumerate(sub_entries, 1):
                c_rel_path = c_path.relative_to(base_dir).as_posix()
                bm_counter[0] += 1
                c_bm_name = f"_Toc_auto_{bm_counter[0]:03d}"
                c_ext = c_path.suffix.lower()
                if c_ext not in SUPPORTED_EXTENSIONS:
                    continue
                c_type = SUPPORTED_EXTENSIONS[c_ext]
                c_title = format_node_title(c_raw_title, level + 1, child_idx, "")
                c_node = {
                    "level": level + 1,
                    "title": c_title,
                    "toc_title": c_raw_title,
                    "type": c_type,
                    "file": c_rel_path,
                    "bookmark_name": c_bm_name,
                    "bm_id": bm_counter[0]
                }
                child_nodes.append(c_node)
                
            parent_node = {
                "level": level,
                "title": parent_title,
                "toc_title": parent_raw_title,
                "type": "folder",
                "folder": parent_rel_path,
                "bookmark_name": parent_bm_name,
                "bm_id": bm_counter[0],
                "children": child_nodes
            }
            nodes.append(parent_node)
            group_idx += 1

    return nodes


def flatten_tree_nodes(nodes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """展开树状大纲节点为扁平列表，方便遍历与目录构建"""
    flat = []
    for n in nodes:
        flat.append(n)
        if "children" in n and n["children"]:
            flat.extend(flatten_tree_nodes(n["children"]))
    return flat
