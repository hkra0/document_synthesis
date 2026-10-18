#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""P6 F0 字体探针：记录本机字体的注册方式、文件 name 表中的名字和 Word 字体列表。

只读，不修改注册表与字体文件；只用标准库，``--word`` 时另需 pywin32。

- Windows：读取 ``HKLM``/``HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts``
  与 ``FontSubstitutes``，并补扫 ``%WINDIR%\\Fonts``、``%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts``
  中未登记的文件。
- macOS / Linux：扫描常见字体目录（含 Word 自带的 ``DFonts``），便于对照。

每个字体文件都解析 ``name`` 表（TTF、OTF、TTC），记录各语言的家族名（nameID 1、16）、
全名（4）和 PostScript 名（6），再按名字查找目标字体。

    python docs/acceptance/public-release/font_probe.py --output font-probe.json
    python docs/acceptance/public-release/font_probe.py --word --output font-probe.json
    python docs/acceptance/public-release/font_probe.py --font 方正仿宋_GBK --output font-probe.json

``--word`` 启动一个隔离的 Word 实例读取 ``Application.FontNames`` 后退出（仅 Windows）。
输出中的用户目录与 Windows 目录替换为 ``<user-home>``、``<windir>``，但仍含本机字体清单，
请先放在仓库之外。
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import platform
import struct
import sys
from typing import Any, Dict, Iterable, List, Optional, Tuple

PROBE_VERSION = 1

# 预设与默认配置使用的字体、它们在新版 Windows 上的常见替代，以及一个保证不存在的哨兵名。
DEFAULT_TARGETS = [
    "仿宋_GB2312", "楷体_GB2312", "方正小标宋简体",
    "仿宋", "楷体", "黑体", "宋体", "微软雅黑",
    "Times New Roman", "Segoe UI", "Cambria",
    "P6缺失测试字体",
]
FONT_SUFFIXES = {".ttf", ".otf", ".ttc", ".otc"}
NAME_IDS = {1: "family", 4: "full", 6: "postscript", 16: "typographic_family"}
WINDOWS_LANGS = {
    0x0409: "en-US", 0x0804: "zh-CN", 0x0404: "zh-TW", 0x0C04: "zh-HK",
    0x1004: "zh-SG", 0x1404: "zh-MO", 0x0411: "ja-JP", 0x0412: "ko-KR",
}
MAC_ENCODINGS = {0: "mac_roman", 1: "shift_jis", 2: "big5", 3: "euc_kr", 25: "gb2312"}
MAC_LANGS = {0: "en", 11: "ja", 19: "zh-Hant", 23: "ko", 33: "zh-Hans"}
FONTS_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
SUBSTITUTES_KEY = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes"


# ---------------------------------------------------------------- 路径脱敏

def _redaction_prefixes() -> List[Tuple[str, str]]:
    prefixes = []
    for env, label in (("WINDIR", "<windir>"), ("USERPROFILE", "<user-home>"), ("HOME", "<user-home>")):
        value = os.environ.get(env)
        if value and len(value) > 3:
            prefixes.append((value.rstrip("\\/"), label))
    prefixes.sort(key=lambda item: len(item[0]), reverse=True)
    return prefixes


_PREFIXES = _redaction_prefixes()


def redact(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    lowered = text.lower()
    for prefix, label in _PREFIXES:
        if lowered.startswith(prefix.lower()):
            return label + text[len(prefix):]
    return text


# ---------------------------------------------------------------- name 表

class FontParseError(Exception):
    pass


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


def _parse_face(handle, offset: int) -> Dict[str, Any]:
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
        field = NAME_IDS.get(name_id)
        if field is None or language_id >= 0x8000:
            continue
        raw = _read(handle, name_offset + string_offset + str_offset, length)
        text = _decode_name(platform_id, encoding_id, raw)
        if text is None:
            continue
        text = text.strip("\x00").strip()
        if text:
            names.setdefault(field, {}).setdefault(_lang_tag(platform_id, language_id), text)
    return names


def read_font_names(path: Path) -> List[Dict[str, Any]]:
    """返回字体文件中每个子字体的名字；TTC 有多个子字体。"""
    with open(path, "rb") as handle:
        tag = _read(handle, 0, 4)
        if tag == b"ttcf":
            _version, num_fonts = struct.unpack(">II", _read(handle, 4, 8))
            offsets = struct.unpack(f">{num_fonts}I", _read(handle, 12, 4 * num_fonts))
        else:
            offsets = (0,)
        faces = []
        for index, offset in enumerate(offsets):
            try:
                faces.append({"index": index, "names": _parse_face(handle, offset)})
            except (FontParseError, struct.error) as exc:
                faces.append({"index": index, "error": str(exc)})
        return faces


def _all_names(face: Dict[str, Any]) -> Iterable[Tuple[str, str, str]]:
    for field, by_lang in face.get("names", {}).items():
        for lang, text in by_lang.items():
            yield field, lang, text


def _norm(name: str) -> str:
    return name.strip().casefold()


# ---------------------------------------------------------------- 字体来源

def _windows_sources() -> Tuple[List[Dict[str, Any]], Dict[str, List[Dict[str, Any]]], Dict[str, Any]]:
    import winreg

    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    system_fonts = windir / "Fonts"
    local = os.environ.get("LOCALAPPDATA")
    user_fonts = Path(local) / "Microsoft" / "Windows" / "Fonts" if local else None

    registry: List[Dict[str, Any]] = []
    files: Dict[str, List[Dict[str, Any]]] = {}
    for hive_name, hive in (("HKLM", winreg.HKEY_LOCAL_MACHINE), ("HKCU", winreg.HKEY_CURRENT_USER)):
        entry: Dict[str, Any] = {"hive": hive_name, "key": FONTS_KEY, "values": []}
        try:
            with winreg.OpenKey(hive, FONTS_KEY) as key:
                index = 0
                while True:
                    try:
                        value_name, data, _kind = winreg.EnumValue(key, index)
                    except OSError:
                        break
                    index += 1
                    data_str = str(data)
                    path = Path(data_str)
                    if not path.is_absolute():
                        path = system_fonts / data_str
                    exists = path.is_file()
                    entry["values"].append({
                        "value_name": value_name, "data": redact(data_str),
                        "path": redact(str(path)), "exists": exists,
                    })
                    if exists:
                        files.setdefault(str(path.resolve()).lower(), []).append(
                            {"source": f"registry:{hive_name}", "value_name": value_name, "path": path}
                        )
        except OSError as exc:
            entry["error"] = str(exc)
        registry.append(entry)

    unregistered: Dict[str, List[str]] = {}
    for label, directory in (("system_fonts_dir", system_fonts), ("user_fonts_dir", user_fonts)):
        if directory is None or not directory.is_dir():
            continue
        for path in sorted(directory.iterdir()):
            if path.suffix.lower() not in FONT_SUFFIXES or not path.is_file():
                continue
            key_path = str(path.resolve()).lower()
            if key_path not in files:
                files[key_path] = [{"source": f"unregistered:{label}", "value_name": None, "path": path}]
                unregistered.setdefault(label, []).append(redact(str(path)))

    substitutes: Dict[str, Any] = {}
    for hive_name, hive in (("HKLM", winreg.HKEY_LOCAL_MACHINE), ("HKCU", winreg.HKEY_CURRENT_USER)):
        try:
            with winreg.OpenKey(hive, SUBSTITUTES_KEY) as key:
                values = {}
                index = 0
                while True:
                    try:
                        value_name, data, _kind = winreg.EnumValue(key, index)
                    except OSError:
                        break
                    index += 1
                    values[value_name] = str(data)
                substitutes[hive_name] = values
        except OSError as exc:
            substitutes[hive_name] = {"error": str(exc)}

    extra = {"registry": registry, "unregistered_files": unregistered, "font_substitutes": substitutes}
    return registry, files, extra


def _directory_sources() -> Tuple[Dict[str, List[Dict[str, Any]]], Dict[str, Any]]:
    home = Path.home()
    if sys.platform == "darwin":
        directories = [
            Path("/System/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
            Path("/Library/Fonts"), home / "Library" / "Fonts",
            Path("/Applications/Microsoft Word.app/Contents/Resources/DFonts"),
        ]
    else:
        directories = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                       home / ".local" / "share" / "fonts", home / ".fonts"]
    files: Dict[str, List[Dict[str, Any]]] = {}
    scanned = []
    for directory in directories:
        if not directory.is_dir():
            continue
        scanned.append(redact(str(directory)))
        walker = directory.rglob("*") if sys.platform != "darwin" else directory.iterdir()
        for path in walker:
            if path.suffix.lower() in FONT_SUFFIXES and path.is_file():
                files.setdefault(str(path.resolve()), []).append(
                    {"source": f"dir:{redact(str(directory))}", "value_name": None, "path": path}
                )
    return files, {"scanned_directories": scanned}


# ---------------------------------------------------------------- Word

def _word_font_names() -> Dict[str, Any]:
    if sys.platform != "win32":
        return {"available": False, "reason": "仅在 Windows 上读取 Word FontNames"}
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        return {"available": False, "reason": f"缺少 pywin32: {exc}"}
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        names = [str(name) for name in app.FontNames]
        return {"available": True, "version": str(app.Version), "build": str(app.Build),
                "count": len(names), "font_names": sorted(set(names))}
    except Exception as exc:  # COM 错误种类多，原样记录
        return {"available": False, "reason": f"{type(exc).__name__}: {exc}"}
    finally:
        if app is not None:
            try:
                app.Quit(0)
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------- 汇总

def probe(targets: List[str], include_word: bool) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "probe_version": PROBE_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "platform": {
            "system": platform.system(), "release": platform.release(),
            "version": platform.version(), "machine": platform.machine(),
            "python": sys.version.split()[0],
        },
        "targets": targets,
    }
    if sys.platform == "win32":
        _registry, files, extra = _windows_sources()
    else:
        files, extra = _directory_sources()
    result["sources"] = extra

    fonts = []
    for _key, occurrences in sorted(files.items()):
        path = occurrences[0]["path"]
        record: Dict[str, Any] = {
            "path": redact(str(path)),
            "sources": [{"source": item["source"], "value_name": item["value_name"]} for item in occurrences],
        }
        try:
            record["faces"] = read_font_names(path)
        except (OSError, FontParseError, struct.error) as exc:
            record["error"] = f"{type(exc).__name__}: {exc}"
        fonts.append(record)
    result["fonts"] = fonts
    result["parse_errors"] = sum(
        1 for font in fonts
        if "error" in font or any("error" in face for face in font.get("faces", []))
    )

    word = _word_font_names() if include_word else {"available": False, "reason": "未指定 --word"}
    result["word"] = word
    word_names = {_norm(name): name for name in word.get("font_names", [])}

    matches: Dict[str, Any] = {}
    for target in targets:
        wanted = _norm(target)
        name_table = []
        value_name_hits = []
        for font in fonts:
            for face in font.get("faces", []):
                for field, lang, text in _all_names(face):
                    if field != "postscript" and _norm(text) == wanted:
                        name_table.append({"path": font["path"], "face": face["index"], "field": field, "lang": lang})
            for source in font["sources"]:
                value_name = source.get("value_name")
                if value_name and wanted in _norm(value_name):
                    value_name_hits.append({"path": font["path"], "source": source["source"], "value_name": value_name})
        matches[target] = {
            "in_name_table": bool(name_table),
            "name_table_matches": name_table,
            "registry_value_name_matches": value_name_hits,
            "in_word_font_names": (wanted in word_names) if word.get("available") else None,
        }
    result["matches"] = matches
    return result


def _print_summary(result: Dict[str, Any]) -> None:
    print(f"平台: {result['platform']['system']} {result['platform']['release']} "
          f"({result['platform']['version']})，Python {result['platform']['python']}")
    print(f"字体文件: {len(result['fonts'])} 个，解析失败 {result['parse_errors']} 个")
    unregistered = result["sources"].get("unregistered_files")
    if unregistered is not None:
        counts = ", ".join(f"{label} {len(items)}" for label, items in unregistered.items()) or "无"
        print(f"字体目录中未在注册表登记的文件: {counts}")
    word = result["word"]
    if word.get("available"):
        print(f"Word {word['version']} (build {word['build']}) FontNames: {word['count']} 个")
    else:
        print(f"Word FontNames: 未读取（{word.get('reason')}）")
    print()
    print(f"{'目标字体':<18} {'name表':<7} {'注册表值名':<9} {'Word':<6} 位置")
    for target, info in result["matches"].items():
        in_word = info["in_word_font_names"]
        word_text = "-" if in_word is None else ("有" if in_word else "无")
        hives = sorted({hit["source"] for hit in info["registry_value_name_matches"]})
        places = sorted({f"{Path(m['path']).name}[{m['face']}] {m['field']}/{m['lang']}"
                         for m in info["name_table_matches"]})
        print(f"{target:<18} {'有' if info['in_name_table'] else '无':<7} "
              f"{(','.join(hives) or '无'):<9} {word_text:<6} {'; '.join(places[:3])}"
              + (" ..." if len(places) > 3 else ""))


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="P6 字体探针（只读）")
    parser.add_argument("--output", type=Path, required=True, help="完整结果 JSON 的写入位置（建议放在仓库之外）")
    parser.add_argument("--font", action="append", default=[], help="额外检查的字体名，可重复")
    parser.add_argument("--word", action="store_true", help="启动隔离 Word 实例读取 FontNames（仅 Windows）")
    args = parser.parse_args(argv)

    targets = DEFAULT_TARGETS + [name for name in args.font if name not in DEFAULT_TARGETS]
    result = probe(targets, args.word)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    _print_summary(result)
    print(f"\n完整结果: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
