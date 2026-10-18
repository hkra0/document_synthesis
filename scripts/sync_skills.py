#!/usr/bin/env python3
"""把唯一的 Skill 源文件复制到各 agent 的发现目录。

源文件：skills/<name>/SKILL.md。副本逐字节相同；用复制而不是符号链接，
因为 Windows 签出时符号链接不可靠。

    python scripts/sync_skills.py          # 写入副本
    python scripts/sync_skills.py --check  # 只检查，不一致时退出码 1
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Tuple

REPO_ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = Path("skills")

# 发现目录（2026-10-18 按各工具官方文档核对）：
# - .claude/skills：Claude Code；GitHub Copilot、Cursor 也读取
# - .agents/skills：Codex；GitHub Copilot、Cursor 也读取
TARGET_DIRS = (Path(".claude") / "skills", Path(".agents") / "skills")


def skill_pairs(root: Path = REPO_ROOT) -> List[Tuple[Path, Path]]:
    """返回 (源文件, 副本) 列表，按技能名排序。"""
    pairs = []
    for source in sorted((root / SOURCE_DIR).glob("*/SKILL.md")):
        name = source.parent.name
        for target_dir in TARGET_DIRS:
            pairs.append((source, root / target_dir / name / "SKILL.md"))
    return pairs


def out_of_sync(root: Path = REPO_ROOT) -> List[Path]:
    """返回缺失或与源文件不一致的副本。"""
    stale = []
    for source, target in skill_pairs(root):
        if not target.is_file() or target.read_bytes() != source.read_bytes():
            stale.append(target)
    return stale


def sync(root: Path = REPO_ROOT) -> List[Path]:
    """写入缺失或不一致的副本，返回写入的路径。"""
    written = []
    for source, target in skill_pairs(root):
        data = source.read_bytes()
        if target.is_file() and target.read_bytes() == data:
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)
        written.append(target)
    return written


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="同步 skills/ 下的 SKILL.md 到各 agent 目录")
    parser.add_argument("--check", action="store_true", help="只检查副本是否一致，不写入")
    args = parser.parse_args(argv)

    if not skill_pairs():
        print(f"未找到 {SOURCE_DIR}/*/SKILL.md", file=sys.stderr)
        return 1
    if args.check:
        stale = out_of_sync()
        for path in stale:
            print(f"[不一致] {path.relative_to(REPO_ROOT).as_posix()}")
        if stale:
            print("请运行 python scripts/sync_skills.py", file=sys.stderr)
            return 1
        print("所有 Skill 副本与源文件一致")
        return 0
    for path in sync():
        print(f"[已写入] {path.relative_to(REPO_ROOT).as_posix()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
