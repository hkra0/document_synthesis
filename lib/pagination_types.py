# -*- coding: utf-8 -*-
"""
页码序列与分节测量模型 (lib/pagination_types.py)

定义多节页码格式化（罗马数字、字母、阿拉伯数字）、PageSequence 规格以及 PageRecord 测量数据结构。
"""

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple


ROMAN_VAL_MAP = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
    (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
    (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def int_to_roman(value: int) -> str:
    """将正整数转为大写罗马数字。"""
    if value <= 0:
        return str(value)
    result = []
    num = value
    for val, roman in ROMAN_VAL_MAP:
        while num >= val:
            result.append(roman)
            num -= val
    return "".join(result)


def roman_to_int(roman_str: str) -> Optional[int]:
    """尝试将罗马数字字符串解析为正整数，解析失败返回 None。"""
    if not roman_str:
        return None
    s = roman_str.upper().strip()
    roman_map = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
    total = 0
    prev_val = 0
    for char in reversed(s):
        val = roman_map.get(char)
        if val is None:
            return None
        if val < prev_val:
            total -= val
        else:
            total += val
            prev_val = val
    return total if total > 0 else None


def int_to_letters(value: int, lowercase: bool = True) -> str:
    """将 1-based 正整数转为字母编号 (1->a, 26->z, 27->aa 等)。"""
    if value <= 0:
        return str(value)
    result = []
    num = value
    while num > 0:
        num -= 1
        result.append(chr(ord('a' if lowercase else 'A') + (num % 26)))
        num //= 26
    return "".join(reversed(result))


def format_page_number(value: int, fmt: str = "decimal") -> str:
    """
    根据 Word/OOXML pgNumType 的 fmt 格式化页码整数。
    支持 decimal, upperRoman, lowerRoman, upperLetter, lowerLetter。
    """
    if fmt == "lowerRoman":
        return int_to_roman(value).lower()
    elif fmt == "upperRoman":
        return int_to_roman(value).upper()
    elif fmt == "lowerLetter":
        return int_to_letters(value, lowercase=True)
    elif fmt == "upperLetter":
        return int_to_letters(value, lowercase=False)
    else:  # decimal or fallback
        return str(value)


@dataclass(frozen=True)
class PageSequence:
    """页码序列规格定义"""
    id: str
    format: str = "decimal"  # decimal | upperRoman | lowerRoman | upperLetter | lowerLetter
    start: Optional[int] = None  # None 表示继承前节继续编号 (continue)

    def __post_init__(self):
        valid_formats = {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"}
        if self.format not in valid_formats:
            raise ValueError(f"无效的页码格式: {self.format}，允许值: {valid_formats}")
        if self.start is not None and self.start < 1:
            raise ValueError(f"页码起始值必须大于等于 1，收到: {self.start}")


@dataclass(frozen=True)
class PageRecord:
    """单个物理页的精确排版与测量记录"""
    physical_page: int          # 物理页号 (1-based PDF 页码)
    section_id: int             # 节序号 (1-based)
    sequence_id: Optional[str]  # 序列 ID (如 "front", "main", 或 None)
    number_value: int           # 打印数值 (Word adjusted page number)
    expected_label: str         # 预期标签 (如 "i", "ii", "1", "2", 或 "")
    observed_label: Optional[str] = None  # PDF 实测提取的标签
    label_verified: bool = False # 观测标签与预期是否一致
    verification_status: str = "unverified"  # verified | mismatch | unverified

    def __post_init__(self):
        if self.physical_page < 1 or self.section_id < 1 or self.number_value < 1:
            raise ValueError("PageRecord 的物理页、节序号和编号值必须从 1 开始。")
        if self.verification_status not in {"verified", "mismatch", "unverified"}:
            raise ValueError(f"未知的页码验证状态: {self.verification_status}")
        if self.label_verified and self.verification_status != "verified":
            object.__setattr__(self, "verification_status", "verified")

    @classmethod
    def from_dict(
        cls,
        record: Dict[str, Any],
        *,
        section_id: int = 1,
        sequence_id: Optional[str] = None,
        expected_label: Optional[str] = None,
    ) -> "PageRecord":
        if not isinstance(record, dict):
            raise ValueError("页码测量记录必须是对象。")
        expected = expected_label if expected_label is not None else record.get("expected_label")
        if not isinstance(expected, str):
            raise ValueError("页码测量记录缺少 expected_label。")
        observed = record.get("observed_label")
        status = record.get("verification_status")
        if status is None:
            status = "verified" if observed is not None and observed == expected else (
                "mismatch" if observed is not None else "unverified"
            )
        return cls(
            physical_page=int(record["physical_page"]),
            section_id=section_id,
            sequence_id=sequence_id,
            number_value=int(record.get("number_value", record.get("printed_page", 0))),
            expected_label=expected,
            observed_label=observed,
            label_verified=status == "verified",
            verification_status=status,
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "physical_page": self.physical_page,
            "section_id": self.section_id,
            "sequence_id": self.sequence_id,
            "number_value": self.number_value,
            "expected_label": self.expected_label,
            "observed_label": self.observed_label,
            "label_verified": self.label_verified,
            "verification_status": self.verification_status,
        }
