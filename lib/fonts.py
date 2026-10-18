# -*- coding: utf-8 -*-
"""字体可移植性：判断一次构建要写入文档的字体在本机是否已安装。

格式中写的是字体的本地化家族名（如 ``仿宋_GB2312``）。Windows 注册表值名、
磁盘文件名都不可靠（P6 F0：公文字体的注册表值名全是英文，文件名可能是乱码），
因此各平台只负责列出字体文件，统一读取字体文件 ``name`` 表中所有语言的家族名来匹配。
解析器只用标准库，支持 TTF、OTF、TTC/OTC。
"""

from dataclasses import dataclass, field
from functools import lru_cache
import os
from pathlib import Path
import shutil
import struct
import subprocess
import sys
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Tuple

FONT_SUFFIXES = {".ttf", ".otf", ".ttc", ".otc"}
NAME_IDS = {1: "family", 4: "full", 6: "postscript", 16: "typographic_family"}
# 参与匹配的名字；PostScript 名不是用户可见的字体名，不参与
MATCH_FIELDS = ("family", "typographic_family", "full")
WINDOWS_LANGS = {
    0x0409: "en-US", 0x0804: "zh-CN", 0x0404: "zh-TW", 0x0C04: "zh-HK",
    0x1004: "zh-SG", 0x1404: "zh-MO", 0x0411: "ja-JP", 0x0412: "ko-KR",
}
MAC_ENCODINGS = {0: "mac_roman", 1: "shift_jis", 2: "big5", 3: "euc_kr", 25: "gb2312"}
MAC_LANGS = {0: "en", 11: "ja", 19: "zh-Hant", 23: "ko", 33: "zh-Hans"}
WINDOWS_FONTS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"

# 不经格式包、由代码直接写入文档的字体
RENDERER_FONTS: Tuple[Tuple[str, str], ...] = (
    ("宋体", "页脚页码（styles.setup_footer）"),
)

INSTALLED = "installed"
MISSING = "missing"
UNKNOWN = "unknown"
WORD_ONLY = "word_only"


class FontParseError(Exception):
    pass


# -----------------------------------------------------------------------------
# name 表解析
# -----------------------------------------------------------------------------

@dataclass(frozen=True)
class FontFace:
    """字体文件中的一个子字体；``names`` 为 字段 → 语言 → 名字。"""
    path: Path
    index: int
    names: Dict[str, Dict[str, str]]

    def match_names(self) -> List[str]:
        result = []
        for field_name in MATCH_FIELDS:
            for text in self.names.get(field_name, {}).values():
                if text not in result:
                    result.append(text)
        return result


@dataclass(frozen=True)
class FontFile:
    path: Path
    source: str
    faces: Tuple[FontFace, ...] = ()
    status: str = "ok"  # "ok" | "unsupported" | "error"
    detail: Optional[str] = None


def _read(handle, offset: int, size: int) -> bytes:
    handle.seek(offset)
    data = handle.read(size)
    if len(data) != size:
        raise FontParseError(f"文件在偏移 {offset} 处提前结束")
    return data


def _decode_name(platform_id: int, encoding_id: int, raw: bytes) -> Optional[str]:
    if platform_id in (0, 3):
        return raw.decode("utf-16-be", errors="replace")
    if platform_id == 1:
        codec = MAC_ENCODINGS.get(encoding_id)
        if codec:
            return raw.decode(codec, errors="replace")
    return None


def _lang_tag(platform_id: int, language_id: int) -> str:
    if platform_id == 3:
        return WINDOWS_LANGS.get(language_id, f"win-0x{language_id:04x}")
    if platform_id == 1:
        return "mac-" + MAC_LANGS.get(language_id, str(language_id))
    return "unicode"


def _parse_face(handle, offset: int) -> Dict[str, Dict[str, str]]:
    sfnt_version, num_tables = struct.unpack(">4sH", _read(handle, offset, 6))
    if sfnt_version not in (b"\x00\x01\x00\x00", b"OTTO", b"true"):
        raise FontParseError(f"不支持的 sfnt 版本 {sfnt_version!r}")
    records = _read(handle, offset + 12, 16 * num_tables)
    name_offset = None
    for index in range(num_tables):
        tag, _checksum, table_offset, _length = struct.unpack(">4sIII", records[16 * index:16 * index + 16])
        if tag == b"name":
            name_offset = table_offset
            break
    if name_offset is None:
        raise FontParseError("没有 name 表")
    _format, count, string_offset = struct.unpack(">HHH", _read(handle, name_offset, 6))
    entries = _read(handle, name_offset + 6, 12 * count)
    names: Dict[str, Dict[str, str]] = {}
    for index in range(count):
        platform_id, encoding_id, language_id, name_id, length, str_offset = struct.unpack(
            ">HHHHHH", entries[12 * index:12 * index + 12]
        )
        field_name = NAME_IDS.get(name_id)
        if field_name is None or language_id >= 0x8000:
            continue
        raw = _read(handle, name_offset + string_offset + str_offset, length)
        text = _decode_name(platform_id, encoding_id, raw)
        if text is None:
            continue
        text = text.strip("\x00").strip()
        if text:
            names.setdefault(field_name, {}).setdefault(_lang_tag(platform_id, language_id), text)
    return names


def read_font_file(path: Path, source: str = "") -> FontFile:
    """解析一个字体文件；损坏或不支持的文件返回带 ``status`` 的结果，不抛出。"""
    try:
        with open(path, "rb") as handle:
            tag = _read(handle, 0, 4)
            if tag[:2] == b"MZ":
                return FontFile(path, source, status="unsupported", detail="Windows 位图字体（.fon）")
            if tag == b"ttcf":
                _version, num_fonts = struct.unpack(">II", _read(handle, 4, 8))
                offsets = struct.unpack(f">{num_fonts}I", _read(handle, 12, 4 * num_fonts))
            else:
                offsets = (0,)
            faces = []
            errors = []
            for index, offset in enumerate(offsets):
                try:
                    faces.append(FontFace(path, index, _parse_face(handle, offset)))
                except (FontParseError, struct.error) as exc:
                    errors.append(f"子字体 {index}: {exc}")
    except (OSError, FontParseError, struct.error) as exc:
        return FontFile(path, source, status="error", detail=f"{type(exc).__name__}: {exc}")
    if not faces:
        return FontFile(path, source, status="error", detail="; ".join(errors))
    return FontFile(path, source, tuple(faces), detail="; ".join(errors) or None)


# -----------------------------------------------------------------------------
# 各平台的字体文件来源
# -----------------------------------------------------------------------------

def _iter_font_files(directory: Path) -> Iterable[Path]:
    if not directory.is_dir():
        return
    for path in sorted(directory.rglob("*")):
        if path.suffix.lower() in FONT_SUFFIXES and path.is_file():
            yield path


def _read_windows_registry() -> List[Tuple[str, str, str]]:
    """返回 ``(hive, 值名, 数据)``；读取 HKLM 与 HKCU（按用户安装的字体只在 HKCU）。"""
    import winreg

    entries = []
    for hive_name, hive in (("HKLM", winreg.HKEY_LOCAL_MACHINE), ("HKCU", winreg.HKEY_CURRENT_USER)):
        try:
            with winreg.OpenKey(hive, WINDOWS_FONTS_KEY) as key:
                index = 0
                while True:
                    try:
                        value_name, data, _kind = winreg.EnumValue(key, index)
                    except OSError:
                        break
                    index += 1
                    entries.append((hive_name, value_name, str(data)))
        except OSError:
            continue
    return entries


def windows_font_files(
    registry: Optional[Sequence[Tuple[str, str, str]]] = None,
    windir: Optional[Path] = None,
    local_appdata: Optional[Path] = None,
) -> List[Tuple[Path, str]]:
    """注册表登记的字体文件，加上两个字体目录中未登记的文件。"""
    if windir is None:
        windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    if local_appdata is None and os.environ.get("LOCALAPPDATA"):
        local_appdata = Path(os.environ["LOCALAPPDATA"])
    if registry is None:
        registry = _read_windows_registry()
    system_dir = windir / "Fonts"
    user_dir = local_appdata / "Microsoft" / "Windows" / "Fonts" if local_appdata else None

    result: List[Tuple[Path, str]] = []
    seen = set()

    def add(path: Path, source: str) -> None:
        key = os.path.normcase(str(path))
        if key not in seen:
            seen.add(key)
            result.append((path, source))

    for hive, _value_name, data in registry:
        path = Path(data)
        if not path.is_absolute():
            path = system_dir / data
        if path.is_file():
            add(path, f"registry:{hive}")
    for label, directory in (("system_fonts_dir", system_dir), ("user_fonts_dir", user_dir)):
        if directory is not None:
            for path in _iter_font_files(directory):
                add(path, f"unregistered:{label}")
    return result


def macos_font_directories(home: Optional[Path] = None) -> List[Path]:
    home = home or Path.home()
    return [
        Path("/System/Library/Fonts"),
        Path("/Library/Fonts"),
        home / "Library" / "Fonts",
        # Word 自带的字体，系统字体工具看不到，但 Word 排版时可用
        Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts"),
    ]


def linux_font_directories(home: Optional[Path] = None) -> List[Path]:
    home = home or Path.home()
    return [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
            home / ".local" / "share" / "fonts", home / ".fonts"]


def _fc_list_files() -> Optional[List[Path]]:
    fc_list = shutil.which("fc-list")
    if not fc_list:
        return None
    try:
        output = subprocess.run(
            [fc_list, "--format", "%{file}\n"], capture_output=True,
            encoding="utf-8", errors="replace", timeout=60, check=True,
        ).stdout
    except (OSError, subprocess.SubprocessError):
        return None
    return [Path(line) for line in output.splitlines() if Path(line).suffix.lower() in FONT_SUFFIXES]


def directory_font_files(directories: Sequence[Path]) -> List[Tuple[Path, str]]:
    result = []
    seen = set()
    for directory in directories:
        for path in _iter_font_files(directory):
            key = str(path.resolve())
            if key not in seen:
                seen.add(key)
                result.append((path, f"dir:{directory}"))
    return result


def platform_font_files(platform: Optional[str] = None) -> List[Tuple[Path, str]]:
    platform = platform or sys.platform
    if platform == "win32":
        return windows_font_files()
    if platform == "darwin":
        return directory_font_files(macos_font_directories())
    files = _fc_list_files()
    if files is not None:
        return [(path, "fc-list") for path in files]
    return directory_font_files(linux_font_directories())


# -----------------------------------------------------------------------------
# 索引与检查
# -----------------------------------------------------------------------------

def normalize_font_name(name: str) -> str:
    """只去掉首尾空白并忽略大小写；``仿宋`` 与 ``仿宋_GB2312`` 是不同的字体。"""
    return name.strip().casefold()


@dataclass
class FontIndex:
    files: List[FontFile] = field(default_factory=list)
    error: Optional[str] = None  # 无法列出字体文件时的原因
    _by_name: Dict[str, List[FontFace]] = field(default_factory=dict, repr=False)

    @classmethod
    def build(cls, files: Iterable[Tuple[Path, str]],
              reader: Callable[[Path, str], FontFile] = read_font_file) -> "FontIndex":
        index = cls()
        for path, source in files:
            font_file = reader(path, source)
            index.files.append(font_file)
            for face in font_file.faces:
                for name in face.match_names():
                    index._by_name.setdefault(normalize_font_name(name), []).append(face)
        if not index.files:
            index.error = "没有找到任何字体文件"
        return index

    @property
    def complete(self) -> bool:
        return self.error is None

    def lookup(self, name: str) -> List[FontFace]:
        return list(self._by_name.get(normalize_font_name(name), []))


@lru_cache(maxsize=1)
def installed_fonts() -> FontIndex:
    """本机字体索引（进程内缓存）。列出字体文件失败时返回 ``complete`` 为假的索引。"""
    try:
        files = platform_font_files()
    except Exception as exc:  # 注册表、目录权限等，原因各异
        return FontIndex(error=f"无法列出本机字体: {type(exc).__name__}: {exc}")
    return FontIndex.build(files)


@dataclass(frozen=True)
class RequiredFont:
    name: str
    sources: Tuple[str, ...]


def required_fonts(config) -> List[RequiredFont]:
    """一次构建会写入文档的字体：格式各样式、``fonts`` 配置与代码直接写入的字体。

    来源文档自身带来的字体不在其中。
    """
    collected: Dict[str, Tuple[str, List[str]]] = {}

    def add(name: Optional[str], source: str) -> None:
        if not name or not name.strip():
            return
        key = normalize_font_name(name)
        entry = collected.setdefault(key, (name.strip(), []))
        if source not in entry[1]:
            entry[1].append(source)

    resolved = getattr(config, "resolved_format", None)
    if resolved is not None:
        for style_name in sorted(resolved.styles):
            run = resolved.styles[style_name].run
            if run is None:
                continue
            add(run.east_asia, f"样式 {style_name}（中文）")
            add(run.latin, f"样式 {style_name}（西文）")
            add(run.complex_script, f"样式 {style_name}（复杂文种）")
    for key, name in (getattr(config, "fonts", None) or {}).items():
        add(name, f"fonts.{key}")
    for name, source in RENDERER_FONTS:
        add(name, source)
    return [RequiredFont(name, tuple(sources)) for name, sources in collected.values()]


@dataclass(frozen=True)
class FontCheck:
    name: str
    status: str  # installed | missing | unknown | word_only
    sources: Tuple[str, ...] = ()
    files: Tuple[Path, ...] = ()


def check_fonts(required: Sequence[RequiredFont], index: FontIndex,
                word_font_names: Optional[Iterable[str]] = None) -> List[FontCheck]:
    """逐个判断状态。索引不完整时，查不到的字体记为 ``unknown``，不能写成缺失或已安装。

    ``word_font_names`` 为 Word ``FontNames``（只在 ``--doctor`` 中提供）：本机没有文件、
    但 Word 列表中有的字体记为 ``word_only``（可能是 Microsoft 365 云字体）。
    """
    word = {normalize_font_name(name) for name in word_font_names} if word_font_names is not None else None
    results = []
    for item in required:
        faces = index.lookup(item.name)
        if faces:
            files = tuple(dict.fromkeys(face.path for face in faces))
            results.append(FontCheck(item.name, INSTALLED, item.sources, files))
        elif word is not None and normalize_font_name(item.name) in word:
            results.append(FontCheck(item.name, WORD_ONLY, item.sources))
        elif not index.complete:
            results.append(FontCheck(item.name, UNKNOWN, item.sources))
        else:
            results.append(FontCheck(item.name, MISSING, item.sources))
    return results
