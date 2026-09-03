# -*- coding: utf-8 -*-
"""
页面数据清洗与图像预处理模块
包含 PDF 底部页码覆盖和图像垂直白边裁切
"""
import io
import numpy as np
from PIL import Image
import pymupdf


class PdfPageSanitizer:
    """PDF 页面数据清洗与图像预处理工具"""
    
    @staticmethod
    def erase_footer_page_number(page: pymupdf.Page, page_index: int):
        """覆盖 PDF 页面底部的旧页脚与页码"""
        rect = page.rect
        footer_height = 45
        footer_rect = pymupdf.Rect(0, rect.height - footer_height, rect.width, rect.height)
        page.draw_rect(footer_rect, color=None, fill=(1, 1, 1), overlay=True)

    @staticmethod
    def crop_vertical_whitespace(img: Image.Image, threshold: int = 250) -> Image.Image:
        """
        保持图片水平宽度不变，仅扫描并裁切上下两侧无内容的纯白空白区，
        避免图片因过高被推挤到下一页。
        """
        gray = img.convert("L")
        arr = np.array(gray)
        h, w = arr.shape
        
        row_min = arr.min(axis=1)
        non_white_rows = np.where(row_min < threshold)[0]
        
        if len(non_white_rows) == 0:
            return img
            
        top = non_white_rows[0]
        bottom = non_white_rows[-1] + 1
        
        # 保留少量边缘间距（约 8 像素）
        pad = 8
        top = max(0, top - pad)
        bottom = min(h, bottom + pad)
        
        return img.crop((0, top, w, bottom))
