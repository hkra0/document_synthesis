#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多源公文排版与材料合成统一入口。

构建会先在 output/<项目>/.work/ 中完成转换、精确分页和 QA；只有全部通过后
才替换正式交付物。使用 --plan 可在不写入文件、不调用 Office 的前提下确认目录。
"""

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from lib.config import ConfigError, load_project_config
from lib.engine import BuildError, UnifiedSynthesizer
from lib.qa import OfficeExportError, word_automation_status

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = ROOT / "input"


def discover_all_input_projects(input_dir: Path) -> List[Path]:
    """动态发现 input/ 下的项目目录。"""
    if not input_dir.exists():
        return []
    return [
        item for item in sorted(input_dir.iterdir())
        if item.is_dir() and not item.name.startswith((".", "~$")) and item.name != "__pycache__"
    ]


def discover_standalone_input_files(input_dir: Path) -> List[Path]:
    """发现由 profile 声明处理策略的单文件项目，避免把普通附件误当成项目。"""
    if not input_dir.exists():
        return []
    projects = []
    for item in sorted(input_dir.glob("*.docx")):
        if item.name.startswith((".", "~$")):
            continue
        try:
            config = load_project_config(input_dir, project_name=item.stem)
        except ConfigError as exc:
            raise BuildError(str(exc)) from exc
        if config.source["strategy"] == "highlighted_docx":
            projects.append(item)
    return projects


def find_project_matches(input_dir: Path, target_query: str) -> List[Path]:
    """返回项目名称的精确或模糊匹配结果，不替用户猜测有歧义的结果。"""
    direct = Path(target_query)
    if direct.is_dir() or direct.is_file():
        return [direct.resolve()]
    candidate = input_dir / target_query
    if candidate.is_dir() or candidate.is_file():
        return [candidate.resolve()]

    projects = discover_all_input_projects(input_dir) + discover_standalone_input_files(input_dir)
    exact = [project for project in projects if project.stem.lower() == target_query.lower()]
    if exact:
        return exact
    return [project for project in projects if target_query.lower() in project.stem.lower()]


def resolve_project_directory(input_dir: Path, target_query: str) -> Optional[Path]:
    matches = find_project_matches(input_dir, target_query)
    return matches[0] if len(matches) == 1 else None


def run_synthesis(source_dir: Path, output_dir: Optional[Path] = None, manifest_path: Optional[Path] = None):
    return UnifiedSynthesizer.synthesize(source_dir, output_dir, manifest_path)


def print_plan(source_dir: Path, manifest_path: Optional[Path]) -> None:
    plan = UnifiedSynthesizer.plan(source_dir, manifest_path)
    print("\n" + "=" * 65)
    print(f"《{plan['project']}》构建计划（只读）")
    print("=" * 65)
    print(f"输入目录: {plan['source_dir']}")
    print(f"配置文件: {'已加载' if plan['config_loaded'] else '未加载，使用默认扫描规则'}")
    print(f"封面模板: {plan['template'] or '未找到，将使用内置封面'}")
    print(f"顶级目录: {plan['root_count']} 项；全部节点: {plan['node_count']} 项")
    print(f"需隔离转换: DOC {plan['legacy_doc_count']} 个；PPTX {plan['pptx_count']} 个")
    print("\n目录预览:")
    for node in plan["nodes"]:
        indent = "  " * (node.get("level", 1) - 1)
        path = node.get("file") or node.get("folder", "")
        print(f"{indent}- [{node.get('type')}] {node.get('title')}  ({path})")


def run_doctor(source_dir: Optional[Path] = None, manifest_path: Optional[Path] = None) -> bool:
    """检查精确构建所需的运行环境，不产生或修改项目文件。"""
    print("\n" + "=" * 65)
    print("document_synthesis 环境检查")
    print("=" * 65)

    checks = []
    checks.append((sys.version_info >= (3, 8), f"Python {sys.version.split()[0]}"))
    for module_name in ("docx", "pymupdf", "PIL", "pypdf", "lxml"):
        try:
            __import__(module_name)
            checks.append((True, f"Python 依赖: {module_name}"))
        except ImportError:
            checks.append((False, f"缺少 Python 依赖: {module_name}"))

    word_ready, word_message = word_automation_status()
    checks.append((word_ready, f"Word Automation 与精确分页: {word_message}"))

    if source_dir:
        try:
            plan = UnifiedSynthesizer.plan(source_dir, manifest_path)
            checks.append((plan["node_count"] > 0, f"输入材料: {plan['node_count']} 个大纲节点"))
        except BuildError as exc:
            checks.append((False, str(exc)))

    for passed, message in checks:
        print(f"[{'通过' if passed else '未通过'}] {message}")
    return all(passed for passed, _ in checks)


def _print_project_resolution_error(target_name: str) -> None:
    matches = find_project_matches(DEFAULT_INPUT_DIR, target_name)
    print(f"[错误] 未能唯一匹配到板块: {target_name}")
    if len(matches) > 1:
        print("匹配到多个项目，请输入完整名称:")
        for match in matches:
            print(f"  - {match.stem}")
        return
    projects = discover_all_input_projects(DEFAULT_INPUT_DIR) + discover_standalone_input_files(DEFAULT_INPUT_DIR)
    if projects:
        print("当前可用板块:")
        for project in projects:
            print(f"  - {project.stem}")


def main() -> int:
    parser = argparse.ArgumentParser(description="多源公文排版与材料合成统一主程序")
    parser.add_argument("--project", help="指定项目名称或目录路径；模糊匹配必须唯一")
    parser.add_argument("--all", action="store_true", help="构建 input/ 下全部项目")
    parser.add_argument("--source", help="直接指定原始材料目录，或由 profile 声明的单个 DOCX")
    parser.add_argument("--manifest", help="指定 JSON 配置文件")
    parser.add_argument("--output-dir", help="指定成果输出根目录")
    parser.add_argument("--plan", action="store_true", help="只读扫描并打印构建计划，不调用 Office")
    parser.add_argument("--doctor", action="store_true", help="检查 Python、Office 和可选输入目录")
    args = parser.parse_args()

    if args.all and args.source:
        parser.error("--all 不能与 --source 同时使用")
    if args.plan and args.all:
        parser.error("--plan 一次只能检查一个项目或 --source")
    if args.doctor and args.all:
        parser.error("--doctor 一次只能检查一个项目或 --source")

    manifest_path = Path(args.manifest).resolve() if args.manifest else None
    output_dir = Path(args.output_dir).resolve() if args.output_dir else None

    source_dir: Optional[Path] = None
    if args.source:
        source_dir = Path(args.source).resolve()
        if not source_dir.exists() or not (source_dir.is_dir() or source_dir.is_file()):
            print(f"[错误] --source 必须是存在的目录或文件: {source_dir}")
            return 2
    elif args.project and args.project.lower() != "all":
        source_dir = resolve_project_directory(DEFAULT_INPUT_DIR, args.project)
        if not source_dir:
            _print_project_resolution_error(args.project)
            return 2

    if args.doctor:
        return 0 if run_doctor(source_dir, manifest_path) else 1
    if args.plan:
        if not source_dir:
            print("[错误] --plan 需要 --source 或 --project。")
            return 2
        print_plan(source_dir, manifest_path)
        return 0

    try:
        if args.all or (args.project and args.project.lower() == "all"):
            projects = discover_all_input_projects(DEFAULT_INPUT_DIR) + discover_standalone_input_files(DEFAULT_INPUT_DIR)
            if not projects:
                raise BuildError(f"未在 {DEFAULT_INPUT_DIR} 找到有效输入材料。")
            failures = []
            for project in projects:
                try:
                    project_output = (output_dir / project.stem) if output_dir else None
                    run_synthesis(project, project_output, manifest_path)
                except (BuildError, OfficeExportError, RuntimeError) as exc:
                    failures.append((project.stem, str(exc)))
                    print(f"[未发布] 《{project.stem}》失败: {exc}")
            if failures:
                print("\n以下项目未发布，既有交付物保持不变:")
                for name, reason in failures:
                    print(f"  - {name}: {reason}")
                return 1
            return 0

        if not args.project and not args.source:
            parser.print_help()
            return 2

        if not source_dir:
            parser.error("请提供 --project、--source 或 --all")
        run_synthesis(source_dir, output_dir, manifest_path)
        return 0
    except (BuildError, OfficeExportError, RuntimeError) as exc:
        print(f"[未发布] 构建失败: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
