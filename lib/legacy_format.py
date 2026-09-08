# -*- coding: utf-8 -*-
"""
历史版本格式适配与清洗策略隔离 (lib/legacy_format.py)
负责将 v1/v2 的 ProjectConfig 适配为等价的 ResolvedFormat，并将历史专项清洗封装为 LegacyPolicy。
保证无配置或 v1/v2 项目行为零改变，而 v3 默认不启用旧清洗。
"""

from dataclasses import dataclass
from typing import Dict, Any, Optional

from lib.format_schema import (
    LengthValue,
    LineSpacing,
    RunStyle,
    ParagraphStyle,
    StyleDefinition,
    PageSpec,
    RoleSpec,
    TocSpec,
    HeaderFooterSpec,
    ResolvedFormat,
)


@dataclass(frozen=True)
class LegacyPolicy:
    """
    旧公文业务专项清洗与容错策略。
    v1/v2 默认启用以保证兼容，v3 项目默认关闭。
    """
    text_blackening: bool = True
    long_paragraph_line_spacing_fix: bool = True
    leading_spaces_to_indent: bool = True
    pdf_strip_page_numbers: bool = True
    blank_cropping: bool = True

    @classmethod
    def default_for_version(cls, schema_version: int) -> "LegacyPolicy":
        if schema_version >= 3:
            return cls(
                text_blackening=False,
                long_paragraph_line_spacing_fix=False,
                leading_spaces_to_indent=False,
                pdf_strip_page_numbers=False,
                blank_cropping=False,
            )
        return cls()


def from_project_config(config_obj: Any) -> ResolvedFormat:
    """
    从 v1/v2 ProjectConfig 实例生成内部一致的 ResolvedFormat。
    保持原有默认尺寸 (cm -> mm) 与字体映射。
    """
    page_cfg = getattr(config_obj, "page_setup", {})
    fonts_cfg = getattr(config_obj, "fonts", {})

    # 1. 页面设置 (cm -> mm)
    page_spec = PageSpec(
        width_mm=float(page_cfg.get("width_cm", 21.0)) * 10.0,
        height_mm=float(page_cfg.get("height_cm", 29.7)) * 10.0,
        orientation="portrait",
        margin_top_mm=float(page_cfg.get("margin_top_cm", 3.7)) * 10.0,
        margin_bottom_mm=float(page_cfg.get("margin_bottom_cm", 3.5)) * 10.0,
        margin_left_mm=float(page_cfg.get("margin_left_cm", 2.8)) * 10.0,
        margin_right_mm=float(page_cfg.get("margin_right_cm", 2.6)) * 10.0,
        header_distance_mm=15.0,
        footer_distance_mm=15.0,
        snap_to_grid=True,
    )

    # 2. 字体与样式提取
    title_font = fonts_cfg.get("title", "方正小标宋简体")
    h1_font = fonts_cfg.get("h1", "黑体")
    h2_font = fonts_cfg.get("h2", "楷体_GB2312")
    h3_font = fonts_cfg.get("h3", "仿宋_GB2312")
    body_font = fonts_cfg.get("body", "仿宋_GB2312")
    en_font = fonts_cfg.get("en", "Times New Roman")

    styles: Dict[str, StyleDefinition] = {
        "body": StyleDefinition(
            name="body",
            run=RunStyle(east_asia=body_font, latin=en_font, size_pt=16.0),
            paragraph=ParagraphStyle(
                alignment="justify",
                first_line_indent=LengthValue(value=2.0, unit="char"),
                line_spacing=LineSpacing(mode="multiple", value=1.5),
                widow_control=True,
            ),
        ),
        "title": StyleDefinition(
            name="title",
            run=RunStyle(east_asia=title_font, latin=en_font, size_pt=22.0),
            paragraph=ParagraphStyle(alignment="center", space_before_pt=12.0, space_after_pt=12.0, keep_with_next=True),
        ),
        "subtitle": StyleDefinition(
            name="subtitle",
            run=RunStyle(east_asia=h2_font, latin=en_font, size_pt=16.0),
            paragraph=ParagraphStyle(alignment="center", space_before_pt=6.0, space_after_pt=6.0, keep_with_next=True),
        ),
        "heading.1": StyleDefinition(
            name="heading.1",
            based_on="body",
            run=RunStyle(east_asia=h1_font, latin=en_font, size_pt=16.0),
            paragraph=ParagraphStyle(keep_with_next=True, first_line_indent=LengthValue(value=2.0, unit="char")),
        ),
        "heading.2": StyleDefinition(
            name="heading.2",
            based_on="body",
            run=RunStyle(east_asia=h2_font, latin=en_font, size_pt=16.0),
            paragraph=ParagraphStyle(keep_with_next=True, first_line_indent=LengthValue(value=2.0, unit="char")),
        ),
        "heading.3": StyleDefinition(
            name="heading.3",
            based_on="body",
            run=RunStyle(east_asia=h3_font, latin=en_font, size_pt=16.0, bold=True),
            paragraph=ParagraphStyle(keep_with_next=True, first_line_indent=LengthValue(value=2.0, unit="char")),
        ),
        "heading.4": StyleDefinition(name="heading.4", based_on="body"),
        "heading.5": StyleDefinition(name="heading.5", based_on="body"),
        "heading.6": StyleDefinition(name="heading.6", based_on="body"),
        "heading.7": StyleDefinition(name="heading.7", based_on="body"),
        "heading.8": StyleDefinition(name="heading.8", based_on="body"),
        "heading.9": StyleDefinition(name="heading.9", based_on="body"),
        "toc.title": StyleDefinition(
            name="toc.title",
            run=RunStyle(east_asia=title_font, latin=en_font, size_pt=22.0),
            paragraph=ParagraphStyle(alignment="center", space_before_pt=12.0, space_after_pt=12.0),
        ),
        "header": StyleDefinition(name="header", run=RunStyle(east_asia=body_font, size_pt=10.5)),
        "footer": StyleDefinition(name="footer", run=RunStyle(east_asia=body_font, size_pt=10.5)),
    }

    roles: Dict[str, RoleSpec] = {
        "body": RoleSpec(style="body", include_in_toc=False),
        "title": RoleSpec(style="title", include_in_toc=False),
        "subtitle": RoleSpec(style="subtitle", include_in_toc=False),
        "heading.1": RoleSpec(style="heading.1", outline_level=1, include_in_toc=True),
        "heading.2": RoleSpec(style="heading.2", outline_level=2, include_in_toc=True),
        "heading.3": RoleSpec(style="heading.3", outline_level=3, include_in_toc=True),
        "heading.4": RoleSpec(style="heading.4", outline_level=4, include_in_toc=True),
        "heading.5": RoleSpec(style="heading.5", outline_level=5, include_in_toc=True),
        "heading.6": RoleSpec(style="heading.6", outline_level=6, include_in_toc=True),
        "heading.7": RoleSpec(style="heading.7", outline_level=7, include_in_toc=True),
        "heading.8": RoleSpec(style="heading.8", outline_level=8, include_in_toc=True),
        "heading.9": RoleSpec(style="heading.9", outline_level=9, include_in_toc=True),
        "toc.title": RoleSpec(style="toc.title", include_in_toc=False),
    }

    toc = TocSpec(title="目  录", title_role="toc.title", max_level=3, leader="dots")
    header = HeaderFooterSpec(mode="none")
    footer = HeaderFooterSpec(mode="managed", format="dash_number", alignment="center", style="footer")

    return ResolvedFormat(
        id="legacy-adapted",
        version="1.0.0",
        page=page_spec,
        styles=styles,
        roles=roles,
        toc=toc,
        header=header,
        footer=footer,
        required_capabilities=[
            "styles.paragraph.v1",
            "styles.character.v1",
            "layout.single_column.v1",
            "docx.single_source_preservation.v1",
        ],
        provenance={"/": "ProjectConfig (v1/v2 adapted)"},
        content_hash="legacy-adapted-v1",
    )
