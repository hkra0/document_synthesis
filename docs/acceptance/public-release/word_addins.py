#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""P4 插件对照：临时禁用 Word 第三方 COM 加载项，并按备份恢复（仅 Windows）。

第三方加载项可能同时注册在三个位置；为所有用户安装的加载项以 HKLM 为准，
只改 HKCU 无法阻止加载（P4 实测 iSlide Tools 与 Acrobat PDFMaker 即如此）：

- ``HKCU\\Software\\Microsoft\\Office\\Word\\Addins\\<ProgID>``
- ``HKLM\\Software\\Microsoft\\Office\\Word\\Addins\\<ProgID>``
- ``HKLM\\Software\\WOW6432Node\\Microsoft\\Office\\Word\\Addins\\<ProgID>``

``disable`` 先把上述位置中每个已注册 ProgID 的 ``LoadBehavior`` 原值写入备份文件，
再改为 2（不加载）；``restore`` 按备份逐项写回，原本没有该值的位置不新建。
修改 HKLM 需要管理员权限，请在“以管理员身份运行”的终端中执行 disable 与 restore。
Click-to-Run 虚拟注册表中的 Microsoft 自带加载项不受影响。

    python docs/acceptance/public-release/word_addins.py status
    python docs/acceptance/public-release/word_addins.py disable --backup <文件>
    python docs/acceptance/public-release/word_addins.py restore --backup <文件>
    python docs/acceptance/public-release/word_addins.py verify

``verify`` 启动一个隔离的自动化 Word 实例，读取各 COM 加载项是否已连接后退出。
备份文件含本机加载项名称，请放在仓库之外，不入库。
"""

from __future__ import annotations

import argparse
import ctypes
import json
from pathlib import Path
import sys
from typing import Any, Dict, List, Optional, Tuple

DO_NOT_LOAD = 2
LOCATIONS: List[Tuple[str, str]] = [
    ("HKCU", r"Software\Microsoft\Office\Word\Addins"),
    ("HKLM", r"Software\Microsoft\Office\Word\Addins"),
    ("HKLM", r"Software\WOW6432Node\Microsoft\Office\Word\Addins"),
]


def _root(name: str):
    import winreg

    return {"HKCU": winreg.HKEY_CURRENT_USER, "HKLM": winreg.HKEY_LOCAL_MACHINE}[name]


def _location_key(hive: str, path: str) -> str:
    return f"{hive}\\{path}"


def _read(hive: str, path: str) -> Dict[str, Optional[int]]:
    """返回该位置下每个 ProgID 的 LoadBehavior；没有该值时为 None。"""
    import winreg

    result: Dict[str, Optional[int]] = {}
    try:
        with winreg.OpenKey(_root(hive), path) as key:
            index = 0
            while True:
                try:
                    progid = winreg.EnumKey(key, index)
                except OSError:
                    break
                index += 1
                try:
                    with winreg.OpenKey(_root(hive), path + "\\" + progid) as sub:
                        result[progid] = int(winreg.QueryValueEx(sub, "LoadBehavior")[0])
                except OSError:
                    result[progid] = None
    except OSError:
        pass
    return result


def status() -> Dict[str, Dict[str, Optional[int]]]:
    return {_location_key(hive, path): _read(hive, path) for hive, path in LOCATIONS}


def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _write(hive: str, path: str, progid: str, value: int) -> None:
    import winreg

    with winreg.OpenKey(_root(hive), path + "\\" + progid, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, int(value))


def disable(backup: Path) -> Dict[str, Any]:
    if not _is_admin():
        raise SystemExit("修改 HKLM 需要管理员权限：请在“以管理员身份运行”的终端中执行。未做任何修改。")
    if backup.exists():
        raise SystemExit(f"备份文件已存在，为避免覆盖原始值已停止: {backup}")
    current = status()
    backup.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    for hive, path in LOCATIONS:
        for progid, value in current[_location_key(hive, path)].items():
            if value is not None:
                _write(hive, path, progid, DO_NOT_LOAD)
    return status()


def restore(backup: Path) -> Dict[str, Any]:
    if not _is_admin():
        raise SystemExit("恢复 HKLM 需要管理员权限：请在“以管理员身份运行”的终端中执行。未做任何修改。")
    saved = json.loads(backup.read_text(encoding="utf-8"))
    for hive, path in LOCATIONS:
        for progid, value in saved.get(_location_key(hive, path), {}).items():
            if value is not None:
                _write(hive, path, progid, value)
    return status()


def verify() -> Dict[str, bool]:
    """在隔离的自动化实例中读取加载项连接状态；不修改任何设置。"""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    try:
        app.Visible = False
        return {app.COMAddIns.Item(i).ProgId: bool(app.COMAddIns.Item(i).Connect)
                for i in range(1, app.COMAddIns.Count + 1)}
    finally:
        app.Quit(0)
        del app
        pythoncom.CoUninitialize()


def main() -> int:
    if sys.platform != "win32":
        print("仅支持 Windows。")
        return 2
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("command", choices=("status", "disable", "restore", "verify"))
    parser.add_argument("--backup", type=Path)
    args = parser.parse_args()
    if args.command in ("disable", "restore") and args.backup is None:
        parser.error("disable / restore 需要 --backup")
    if args.command == "status":
        result: Any = status()
    elif args.command == "disable":
        result = disable(args.backup)
    elif args.command == "restore":
        result = restore(args.backup)
    else:
        result = verify()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
