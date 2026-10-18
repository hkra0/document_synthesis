# -*- coding: utf-8 -*-
"""P6 F1 字体索引：按字体文件 name 表判断字体是否已安装。

覆盖：
1. name 表解析：Windows/Mac/Unicode 平台、语言标记、TTC 多个子字体、PostScript 名不参与匹配
2. 不支持与损坏的文件：.fon 记为不支持，截断、无 name 表、非 sfnt 记为错误，不抛出
3. Windows 来源：HKLM 相对文件名、HKCU 绝对路径、未登记文件补扫、不重复、文件名不影响识别
4. 索引与检查：大小写与首尾空白、不做模糊匹配、installed/missing/unknown/word_only
5. required_fonts：格式样式、fonts 配置、代码直接写入的字体，去重并保留来源
"""

import struct
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from lib import fonts  # noqa: E402
from lib.format_schema import RunStyle, StyleDefinition  # noqa: E402


def _win(lang, name_id, text):
    return (3, 1, lang, name_id, text.encode("utf-16-be"))


def _mac(encoding, lang, name_id, text, codec):
    return (1, encoding, lang, name_id, text.encode(codec))


def make_face(records, include_name=True, sfnt=b"\x00\x01\x00\x00"):
    """返回一个只含 name 表的最小 sfnt 子字体（表偏移相对子字体起点，由调用方修正）。"""
    strings = b""
    entries = b""
    for platform_id, encoding_id, language_id, name_id, raw in records:
        entries += struct.pack(">HHHHHH", platform_id, encoding_id, language_id, name_id, len(raw), len(strings))
        strings += raw
    name_table = struct.pack(">HHH", 0, len(records), 6 + len(entries)) + entries + strings
    tag = b"name" if include_name else b"head"
    return sfnt, tag, name_table


def build_font(faces):
    """把 make_face 的结果组装成 TTF（一个子字体）或 TTC（多个）。"""
    def face_bytes(face, base):
        sfnt, tag, table = face
        header = sfnt + struct.pack(">HHHH", 1, 16, 0, 0)
        table_offset = base + len(header) + 16
        return header + struct.pack(">4sIII", tag, 0, table_offset, len(table)) + table

    if len(faces) == 1:
        return face_bytes(faces[0], 0)
    header_size = 12 + 4 * len(faces)
    blobs = []
    offsets = []
    position = header_size
    for face in faces:
        offsets.append(position)
        blob = face_bytes(face, position)
        blobs.append(blob)
        position += len(blob)
    header = b"ttcf" + struct.pack(">II", 0x00010000, len(faces)) + struct.pack(f">{len(faces)}I", *offsets)
    return header + b"".join(blobs)


FANGSONG_GB2312 = make_face([
    _win(0x0409, 1, "FangSong_GB2312"),
    _win(0x0804, 1, "仿宋_GB2312"),
    _win(0x0804, 4, "仿宋_GB2312"),
    _win(0x0804, 6, "FangSong_GB2312-PS"),
])


class TmpDirTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def write(self, relative, data):
        path = self.tmp / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path


class NameTableTest(TmpDirTestCase):
    def test_windows_names_by_language(self):
        path = self.write("fs.ttf", build_font([FANGSONG_GB2312]))
        result = fonts.read_font_file(path, "test")
        self.assertEqual(result.status, "ok")
        (face,) = result.faces
        self.assertEqual(face.names["family"], {"en-US": "FangSong_GB2312", "zh-CN": "仿宋_GB2312"})
        self.assertEqual(face.names["postscript"], {"zh-CN": "FangSong_GB2312-PS"})
        self.assertNotIn("FangSong_GB2312-PS", face.match_names())
        self.assertIn("仿宋_GB2312", face.match_names())

    def test_mac_gb2312_and_unicode_platforms(self):
        face = make_face([
            _mac(25, 33, 1, "方正小标宋简体", "gb2312"),
            _mac(0, 0, 1, "FZXiaoBiaoSong-B05S", "mac_roman"),
            (0, 3, 0, 16, "小标宋".encode("utf-16-be")),
        ])
        result = fonts.read_font_file(self.write("fz.ttf", build_font([face])))
        names = result.faces[0].names
        self.assertEqual(names["family"], {"mac-zh-Hans": "方正小标宋简体", "mac-en": "FZXiaoBiaoSong-B05S"})
        self.assertEqual(names["typographic_family"], {"unicode": "小标宋"})

    def test_skips_language_tag_records_and_unknown_encodings(self):
        face = make_face([
            _win(0x8001, 1, "LangTagged"),
            _mac(7, 0, 1, "Unknown", "mac_roman"),
            _win(0x0409, 1, "Kept"),
        ])
        result = fonts.read_font_file(self.write("x.ttf", build_font([face])))
        self.assertEqual(result.faces[0].names["family"], {"en-US": "Kept"})

    def test_ttc_returns_each_face(self):
        simsun = make_face([_win(0x0804, 1, "宋体"), _win(0x0409, 1, "SimSun")])
        nsimsun = make_face([_win(0x0804, 1, "新宋体"), _win(0x0409, 1, "NSimSun")])
        result = fonts.read_font_file(self.write("simsun.ttc", build_font([simsun, nsimsun])))
        self.assertEqual([face.index for face in result.faces], [0, 1])
        self.assertEqual(result.faces[1].names["family"]["zh-CN"], "新宋体")

    def test_opentype_cff_signature(self):
        face = make_face([_win(0x0409, 1, "CffFont")], sfnt=b"OTTO")
        result = fonts.read_font_file(self.write("cff.otf", build_font([face])))
        self.assertEqual(result.faces[0].names["family"], {"en-US": "CffFont"})


class UnreadableFileTest(TmpDirTestCase):
    def test_fon_bitmap_font_is_unsupported(self):
        result = fonts.read_font_file(self.write("SMALLE.FON", b"MZ\xf8\x00" + b"\x00" * 60))
        self.assertEqual(result.status, "unsupported")
        self.assertEqual(result.faces, ())

    def test_broken_files_are_errors(self):
        data = build_font([FANGSONG_GB2312])
        cases = {
            "truncated.ttf": data[:30],
            "no-name.ttf": build_font([make_face([_win(0x0409, 1, "X")], include_name=False)]),
            "not-sfnt.ttf": b"?i\\E" + b"\x00" * 40,
            "empty.ttf": b"",
        }
        for filename, content in cases.items():
            with self.subTest(filename):
                result = fonts.read_font_file(self.write(filename, content))
                self.assertEqual(result.status, "error")
                self.assertTrue(result.detail)

    def test_missing_file_is_error(self):
        result = fonts.read_font_file(self.tmp / "absent.ttf")
        self.assertEqual(result.status, "error")

    def test_ttc_with_one_bad_face_keeps_the_good_one(self):
        good = make_face([_win(0x0409, 1, "Good")])
        bad = make_face([_win(0x0409, 1, "Bad")], include_name=False)
        result = fonts.read_font_file(self.write("mixed.ttc", build_font([good, bad])))
        self.assertEqual(result.status, "ok")
        self.assertEqual(len(result.faces), 1)
        self.assertIn("子字体 1", result.detail)


class WindowsSourcesTest(TmpDirTestCase):
    def test_registry_and_unregistered_files(self):
        windir = self.tmp / "Windows"
        appdata = self.tmp / "AppData" / "Local"
        font = build_font([FANGSONG_GB2312])
        simsun = self.write("Windows/Fonts/simsun.ttc", font)
        self.write("Windows/Fonts/marlett.ttf", font)
        # 解压时编码错误产生的乱码文件名，与 F0 中的楷体文件相同
        user_font = self.write("AppData/Local/Microsoft/Windows/Fonts/┐¼╠σ_GB2312.ttf", font)
        registry = [
            ("HKLM", "SimSun & NSimSun (TrueType)", "simsun.ttc"),
            ("HKLM", "Gone (TrueType)", "gone.ttf"),
            ("HKCU", "KaiTi_GB2312 (TrueType)", str(user_font)),
        ]
        files = fonts.windows_font_files(registry, windir=windir, local_appdata=appdata)
        self.assertEqual(
            [(path.name, source) for path, source in files],
            [
                ("simsun.ttc", "registry:HKLM"),
                ("┐¼╠σ_GB2312.ttf", "registry:HKCU"),
                ("marlett.ttf", "unregistered:system_fonts_dir"),
            ],
        )
        self.assertEqual(files[0][0], simsun)
        index = fonts.FontIndex.build(files)
        self.assertEqual({face.path.name for face in index.lookup("仿宋_GB2312")},
                         {"simsun.ttc", "┐¼╠σ_GB2312.ttf", "marlett.ttf"})

    def test_without_user_fonts_dir(self):
        self.write("Windows/Fonts/a.ttf", build_font([FANGSONG_GB2312]))
        files = fonts.windows_font_files([], windir=self.tmp / "Windows", local_appdata=None)
        self.assertEqual([source for _path, source in files], ["unregistered:system_fonts_dir"])


class DirectorySourcesTest(TmpDirTestCase):
    def test_recurses_filters_suffix_and_deduplicates(self):
        data = build_font([FANGSONG_GB2312])
        self.write("a/sub/One.TTF", data)
        self.write("a/readme.txt", b"x")
        self.write("b/Two.otf", data)
        files = fonts.directory_font_files([self.tmp / "a", self.tmp / "b", self.tmp / "a", self.tmp / "missing"])
        self.assertEqual(sorted(path.name for path, _ in files), ["One.TTF", "Two.otf"])

    def test_macos_directories_include_word_fonts(self):
        directories = fonts.macos_font_directories(home=Path("/home-x"))
        self.assertIn(Path("/home-x/Library/Fonts"), directories)
        self.assertTrue(any(path.name == "DFonts" for path in directories))


class IndexAndCheckTest(TmpDirTestCase):
    def setUp(self):
        super().setUp()
        fangsong = make_face([_win(0x0804, 1, "仿宋"), _win(0x0409, 1, "FangSong")])
        times = make_face([_win(0x0409, 1, "Times New Roman"), _win(0x0409, 4, "Times New Roman Bold")])
        self.files = [
            (self.write("simfang.ttf", build_font([fangsong])), "t"),
            (self.write("times.ttf", build_font([times])), "t"),
            (self.write("bad.ttf", b"garbage!"), "t"),
        ]
        self.index = fonts.FontIndex.build(self.files)

    def required(self, *names):
        return [fonts.RequiredFont(name, ("样式 body（中文）",)) for name in names]

    def test_lookup_ignores_case_and_outer_whitespace_only(self):
        self.assertTrue(self.index.lookup("  times new roman "))
        self.assertTrue(self.index.lookup("FANGSONG"))
        self.assertTrue(self.index.lookup("Times New Roman Bold"))
        self.assertEqual(self.index.lookup("仿宋_GB2312"), [])
        self.assertEqual(self.index.lookup("Times New  Roman"), [])

    def test_statuses(self):
        results = fonts.check_fonts(self.required("仿宋", "仿宋_GB2312", "Aptos"), self.index,
                                    word_font_names=["Aptos", "仿宋"])
        self.assertEqual([(r.name, r.status) for r in results],
                         [("仿宋", "installed"), ("仿宋_GB2312", "missing"), ("Aptos", "word_only")])
        self.assertEqual([path.name for path in results[0].files], ["simfang.ttf"])
        self.assertEqual(results[1].sources, ("样式 body（中文）",))

    def test_without_word_names_missing_stays_missing(self):
        (result,) = fonts.check_fonts(self.required("Aptos"), self.index)
        self.assertEqual(result.status, "missing")

    def test_incomplete_index_reports_unknown_not_missing(self):
        empty = fonts.FontIndex.build([])
        self.assertFalse(empty.complete)
        (result,) = fonts.check_fonts(self.required("仿宋"), empty)
        self.assertEqual(result.status, "unknown")

    def test_parse_errors_do_not_make_index_incomplete(self):
        self.assertTrue(self.index.complete)
        self.assertEqual([f.status for f in self.index.files], ["ok", "ok", "error"])

    def test_installed_fonts_reports_enumeration_failure(self):
        original = fonts.platform_font_files
        fonts.installed_fonts.cache_clear()
        try:
            fonts.platform_font_files = lambda platform=None: (_ for _ in ()).throw(PermissionError("denied"))
            index = fonts.installed_fonts()
            self.assertFalse(index.complete)
            self.assertIn("denied", index.error)
        finally:
            fonts.platform_font_files = original
            fonts.installed_fonts.cache_clear()


class RequiredFontsTest(unittest.TestCase):
    def test_collects_styles_config_fonts_and_renderer_fonts(self):
        styles = {
            "heading.1": StyleDefinition("heading.1", run=RunStyle(east_asia="黑体", latin="Times New Roman")),
            "body": StyleDefinition("body", run=RunStyle(east_asia="仿宋_GB2312", latin="Times New Roman",
                                                         complex_script=" ")),
            "plain": StyleDefinition("plain"),
        }
        config = SimpleNamespace(
            resolved_format=SimpleNamespace(styles=styles),
            fonts={"body": "仿宋_GB2312", "h1": "黑体", "en": "times new roman"},
        )
        required = {item.name: item.sources for item in fonts.required_fonts(config)}
        self.assertEqual(list(required), ["仿宋_GB2312", "Times New Roman", "黑体", "宋体"])
        self.assertEqual(required["仿宋_GB2312"], ("样式 body（中文）", "fonts.body"))
        self.assertEqual(required["Times New Roman"],
                         ("样式 body（西文）", "样式 heading.1（西文）", "fonts.en"))
        self.assertEqual(required["宋体"], ("页脚页码（styles.setup_footer）",))

    def test_legacy_project_config(self):
        from lib.config import DEFAULT_FONTS, ProjectConfig

        config = ProjectConfig({"schema_version": 2, "project_name": "Fictional Fonts", "cover": {"template": False}})
        names = {item.name for item in fonts.required_fonts(config)}
        self.assertTrue(set(DEFAULT_FONTS.values()) <= names, names)


class InstalledFontsSmokeTest(unittest.TestCase):
    def test_runs_on_this_machine(self):
        fonts.installed_fonts.cache_clear()
        try:
            index = fonts.installed_fonts()
        finally:
            fonts.installed_fonts.cache_clear()
        if sys.platform in ("darwin", "win32"):
            self.assertTrue(index.complete, index.error)
            self.assertTrue(any(f.status == "ok" for f in index.files))


if __name__ == "__main__":
    unittest.main()
