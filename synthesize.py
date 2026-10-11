#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多源公文排版与材料合成统一入口。

构建会先在 output/<项目>/.work/ 中完成转换、精确分页和 QA；只有全部通过后
才替换正式交付物。使用 --plan 可在不写入文件、不调用 Office 的前提下确认目录。
"""

import argparse
import json
import os
import sys
from pathlib import Path
from typing import List, Optional

from lib.config import ConfigError, load_project_config
from lib.engine import BuildError, UnifiedSynthesizer
from lib.qa import OfficeExportError, word_automation_status

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = ROOT / "input"


def configure_console_encoding() -> None:
    """确保标准输入输出在各平台终端（如 Windows CP936）下使用 UTF-8 编码。"""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream and hasattr(stream, "reconfigure"):
            try:
                if getattr(stream, "encoding", "").lower() != "utf-8":
                    stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


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
    for item in sorted(input_dir.iterdir()):
        if not item.is_file() or item.suffix.lower() != ".docx":
            continue
        if item.name.startswith((".", "~$")):
            continue
        try:
            config = load_project_config(input_dir, project_name=item.stem)
        except ConfigError as exc:
            raise BuildError(str(exc)) from exc
        if config.source.get("strategy") in ("highlighted_docx", "docx_document"):
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


def run_synthesis(source_dir: Path, output_dir: Optional[Path] = None, manifest_path: Optional[Path] = None,
                  override_paths: Optional[List[Path]] = None, keep_work: bool = False):
    return UnifiedSynthesizer.synthesize(source_dir, output_dir, manifest_path,
                                        override_paths=override_paths, keep_work=keep_work)


def print_plan(source_dir: Path, manifest_path: Optional[Path], override_paths: Optional[List[Path]] = None) -> None:
    plan = UnifiedSynthesizer.plan(source_dir, manifest_path, override_paths=override_paths)
    print("\n" + "=" * 65)
    print(f"《{plan['project']}》构建计划（只读）")
    print("=" * 65)
    print(f"输入目录: {plan['source_dir']}")
    print(f"配置文件: {plan.get('manifest_path') or '未加载，使用默认配置'}")
    for path in plan.get("override_paths", []):
        print(f"显式覆写: {path}")
    for warning in plan.get("warnings", []):
        print(f"[提示] {warning}")
    has_cover = any("cover" in document["parts"] for document in plan["documents"])
    print(f"封面模板: {(plan['template'] or '使用内置封面') if has_cover else '本次不生成封面'}")
    if "content_block_count" in plan:
        print(f"顶级目录: {plan['root_count']} 项；内容块: {plan['content_block_count']} 项；标题: {plan['heading_count']} 项")
    else:
        print(f"顶级目录: {plan['root_count']} 项；全部节点: {plan['node_count']} 项")
    print(f"需隔离转换: DOC {plan['legacy_doc_count']} 个；PPTX {plan['pptx_count']} 个")
    print("\n交付文档:")
    for document in plan.get("documents", []):
        print(f"  - {document['id']}: {document['filename']} ({' + '.join(document['parts'])})")
        if "toc" in document:
            print(f"    目录页码引用: {document['toc']['reference']}；链接: {document['toc']['links']}")
        if "body" in document["parts"]:
            print("    正文页码从 1 开始；封面与目录不显示页码")
    print("\n目录预览:")
    for node in plan["nodes"]:
        indent = "  " * (node.get("level", 1) - 1)
        path = node.get("file") or node.get("folder", "")
        print(f"{indent}- [{node.get('type')}] {node.get('title')}  ({path})")


def run_doctor(source_dir: Optional[Path] = None, manifest_path: Optional[Path] = None,
               override_paths: Optional[List[Path]] = None) -> int:
    """检查精确构建所需的运行环境，返回退出码:
    0: 环境完全就绪，可进行 exact 精确构建
    2: 仅支持 draft 或分析类功能（Python 依赖就绪，但 Office exact 不可用）
    1: 缺少 Python 依赖或输入配置错误
    """
    print("\n" + "=" * 65)
    print("document_synthesis 环境检查")
    print("=" * 65)

    checks = []
    dep_ok = True
    input_ok = True

    py_ok = sys.version_info >= (3, 10)
    checks.append((py_ok, f"Python {sys.version.split()[0]} (要求 >= 3.10)"))
    if not py_ok:
        dep_ok = False

    for module_name in ("docx", "pymupdf", "PIL", "pypdf", "lxml"):
        try:
            __import__(module_name)
            checks.append((True, f"Python 依赖: {module_name}"))
        except ImportError:
            checks.append((False, f"缺少 Python 依赖: {module_name}"))
            dep_ok = False

    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem") as key:
                val, _ = winreg.QueryValueEx(key, "LongPathsEnabled")
                long_paths = bool(val)
        except Exception:
            long_paths = False
        checks.append((long_paths, "Windows 长路径支持 (LongPathsEnabled 已启用)" if long_paths else "Windows 长路径支持 (未启用 LongPathsEnabled；如遇路径超限可在注册表启用)"))

    from lib.office import get_backend
    backend = get_backend()
    probe_stat = backend.probe()
    office_exact = probe_stat.available and probe_stat.fidelity == "exact"
    checks.append((office_exact, f"Office Automation ({backend.name}): {probe_stat.reason}"))

    if source_dir:
        resolved_src = str(source_dir.resolve())
        is_path_safe = len(resolved_src) <= 240
        checks.append((is_path_safe, f"输入路径长度安全 ({len(resolved_src)} 字符 <= 240)" if is_path_safe else f"输入路径过长 ({len(resolved_src)} 字符 > 240)，Windows 下容易超出 MAX_PATH 限制: {resolved_src[:60]}..."))
        if not is_path_safe:
            input_ok = False
        try:
            plan = UnifiedSynthesizer.plan(source_dir, manifest_path, override_paths=override_paths)
            has_content = plan.get("content_block_count", 0) > 0 or plan.get("node_count", 0) > 0
            detail = f"{plan.get('content_block_count', 0)} 个内容块, {plan.get('heading_count', plan.get('node_count', 0))} 个标题" if "content_block_count" in plan else f"{plan['node_count']} 个大纲节点"
            checks.append((has_content, f"输入材料: {detail}"))
            if not has_content:
                input_ok = False
        except BuildError as exc:
            checks.append((False, f"输入材料: {exc}"))
            input_ok = False

    for passed, message in checks:
        print(f"[{'通过' if passed else '未通过'}] {message}")

    if not dep_ok or not input_ok:
        return 1
    if not office_exact:
        return 2
    return 0


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
    configure_console_encoding()
    parser = argparse.ArgumentParser(description="多源公文排版与材料合成统一主程序")
    parser.add_argument("--project", help="指定项目名称或目录路径；模糊匹配必须唯一")
    parser.add_argument("--all", action="store_true", help="构建 input/ 下全部项目")
    parser.add_argument("--source", help="直接指定原始材料目录，或由 profile 声明的单个 DOCX")
    parser.add_argument("--manifest", help="指定 JSON 配置文件")
    parser.add_argument("--override", action="append", default=[], metavar="JSON", help="显式覆写基础配置，可重复；对象合并，数组整体替换")
    parser.add_argument("--keep-work", action="store_true", help="保留内部构建与诊断产物，不将其作为正式交付物")
    parser.add_argument("--output-dir", help="指定成果输出根目录")
    parser.add_argument("--plan", action="store_true", help="只读扫描并打印构建计划，不调用 Office")
    parser.add_argument("--doctor", action="store_true", help="检查 Python、Office 和可选输入目录")
    parser.add_argument(
        "--office-backend", choices=["auto", "word", "none"], default=None,
        help="指定 Office 自动化后端；未指定时读取 DOCUMENT_SYNTHESIS_OFFICE_BACKEND，默认 auto",
    )

    # B6 格式分析与编译参数
    parser.add_argument("--analyze-format", metavar="DOCX", help="分析参考样本文档，生成样式聚类与 HTML 校正报告")
    parser.add_argument("--analysis-dir", metavar="DIR", help="分析报告与校正 HTML 输出目录")
    parser.add_argument("--compile-format", metavar="REPORT", help="根据分析报告和确认决策编译标准格式包")
    parser.add_argument("--decisions", metavar="JSON", help="用户校正后的 decisions.json 文件路径")
    parser.add_argument("--format-base", metavar="REF", help="格式包继承基底，如 preset:academic-basic@1.0.0")
    parser.add_argument("--format-id", metavar="ID", default="custom-format", help="生成的格式包标识 ID")
    parser.add_argument("--format-version", metavar="VER", default="1.0.0", help="生成的格式包版本号")
    parser.add_argument("--format-out", metavar="JSON", help="编译生成的 format.json 输出路径")
    parser.add_argument("--analyze-source", metavar="DOCX", help="分析目标正文文档并提取待映射节点")
    parser.add_argument("--compile-mapping", metavar="REPORT", help="根据分析报告和确认决策编译目标文档角色映射")
    parser.add_argument("--mapping-out", metavar="JSON", help="编译生成的 roles.json 映射输出路径")
    parser.add_argument("--replace-output", action="store_true", help="允许覆盖已存在的分析报告、格式包或映射文件")
    parser.add_argument("--preview-format", metavar="JSON", help="为格式包生成虚构文本的 HTML 近似预览")
    parser.add_argument("--render-preview", metavar="JSON", help="生成格式包预览并输出可审计的预览元数据")
    parser.add_argument("--preview-out", metavar="HTML", help="预览 HTML 输出路径")
    parser.add_argument("--preview-meta", metavar="JSON", help="预览元数据输出路径")
    parser.add_argument("--preview-engine", default="html-css-approximate", help="预览引擎标识")
    parser.add_argument("--migrate-manifest", metavar="JSON", help="将旧 manifest 迁移为显式 output.documents 清单")
    parser.add_argument("--manifest-out", metavar="JSON", help="迁移后的 manifest 输出路径")

    args = parser.parse_args()

    if args.office_backend is not None:
        os.environ["DOCUMENT_SYNTHESIS_OFFICE_BACKEND"] = args.office_backend
    from lib.office import OfficeUnsupportedError, get_backend
    try:
        get_backend()
    except (OfficeUnsupportedError, ValueError) as exc:
        print(f"[错误] Office 后端配置无效: {exc}")
        return 1

    # B6 互斥检查
    b6_actions = [
        bool(args.analyze_format),
        bool(args.compile_format),
        bool(args.analyze_source),
        bool(args.compile_mapping),
        bool(args.preview_format),
        bool(args.render_preview),
        bool(args.migrate_manifest),
    ]
    if sum(b6_actions) > 1:
        parser.error("分析与编译操作互斥，一次只能执行一项。")
    if any(b6_actions):
        if args.all or args.plan or args.doctor or args.project or args.source or args.manifest:
            parser.error("分析与编译操作不能与常规构建参数 (--project, --source, --all, --plan, --doctor, --manifest) 混用。")

    if args.analyze_format:
        if not args.analysis_dir:
            parser.error("--analyze-format 需要指定 --analysis-dir")
        sample_path = Path(args.analyze_format).resolve()
        if not sample_path.is_file():
            print(f"[错误] 样本文档不存在: {sample_path}")
            return 2
        analysis_dir = Path(args.analysis_dir).resolve()
        analysis_json_path = analysis_dir / "analysis.json"
        review_html_path = analysis_dir / "review.html"
        decisions_example_path = analysis_dir / "decisions.example.json"
        if not args.replace_output and (analysis_json_path.exists() or review_html_path.exists() or decisions_example_path.exists()):
            print(f"[错误] 分析目标输出已存在，若要覆盖请使用 --replace-output: {analysis_dir}")
            return 1
        from lib.format_analysis import analyze_format_sample
        from lib.format_review import build_decisions_example, generate_html_review
        try:
            analysis_data = analyze_format_sample(sample_path)
            analysis_dir.mkdir(parents=True, exist_ok=True)
            with open(analysis_json_path, "w", encoding="utf-8") as f:
                json.dump(analysis_data, f, ensure_ascii=False, indent=2)
            generate_html_review(analysis_data, review_html_path)
            with open(decisions_example_path, "w", encoding="utf-8") as f:
                json.dump(build_decisions_example(analysis_data), f, ensure_ascii=False, indent=2)
            print("分析完成！")
            print(f"  - 分析报告: {analysis_json_path}")
            print(f"  - 校正页面: {review_html_path}")
            print(f"  - 决策示例: {decisions_example_path}")
            return 0
        except Exception as exc:
            print(f"[错误] 样例分析失败: {exc}")
            return 1

    if args.compile_format:
        if not args.decisions or not args.format_out:
            parser.error("--compile-format 需要指定 --decisions 和 --format-out")
        report_path = Path(args.compile_format).resolve()
        decisions_path = Path(args.decisions).resolve()
        format_out_path = Path(args.format_out).resolve()
        if not report_path.is_file():
            print(f"[错误] 分析报告文件不存在: {report_path}")
            return 2
        if not decisions_path.is_file():
            print(f"[错误] 决策文件不存在: {decisions_path}")
            return 2
        if not args.replace_output and format_out_path.exists():
            print(f"[错误] 格式包输出已存在，若要覆盖请使用 --replace-output: {format_out_path}")
            return 1
        from lib.format_review import compile_format_package
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                report_data = json.load(f)
            with open(decisions_path, "r", encoding="utf-8") as f:
                decisions_data = json.load(f)
            pkg = compile_format_package(
                analysis_data=report_data,
                decisions_data=decisions_data,
                base_format_ref=args.format_base,
                format_id=args.format_id,
                format_version=args.format_version,
                output_path=format_out_path,
            )
            print(f"格式包编译成功: {format_out_path}")
            return 0
        except Exception as exc:
            print(f"[错误] 格式包编译失败: {exc}")
            return 1

    if args.analyze_source:
        if not args.analysis_dir:
            parser.error("--analyze-source 需要指定 --analysis-dir")
        source_doc_path = Path(args.analyze_source).resolve()
        if not source_doc_path.is_file():
            print(f"[错误] 目标文档不存在: {source_doc_path}")
            return 2
        analysis_dir = Path(args.analysis_dir).resolve()
        analysis_json_path = analysis_dir / "analysis.json"
        decisions_example_path = analysis_dir / "decisions.example.json"
        if not args.replace_output and (analysis_json_path.exists() or decisions_example_path.exists()):
            print(f"[错误] 分析目标输出已存在，若要覆盖请使用 --replace-output: {analysis_json_path}")
            return 1
        from lib.format_analysis import analyze_format_sample
        try:
            analysis_data = analyze_format_sample(source_doc_path)
            analysis_dir.mkdir(parents=True, exist_ok=True)
            with open(analysis_json_path, "w", encoding="utf-8") as f:
                json.dump(analysis_data, f, ensure_ascii=False, indent=2)
            from lib.format_review import build_decisions_example
            with open(decisions_example_path, "w", encoding="utf-8") as f:
                json.dump(build_decisions_example(analysis_data), f, ensure_ascii=False, indent=2)
            print(f"目标分析完成: {analysis_json_path}")
            print(f"决策示例: {decisions_example_path}")
            return 0
        except Exception as exc:
            print(f"[错误] 目标分析失败: {exc}")
            return 1

    if args.compile_mapping:
        if not args.decisions or not args.mapping_out:
            parser.error("--compile-mapping 需要指定 --decisions 和 --mapping-out")
        report_path = Path(args.compile_mapping).resolve()
        decisions_path = Path(args.decisions).resolve()
        mapping_out_path = Path(args.mapping_out).resolve()
        if not report_path.is_file():
            print(f"[错误] 分析报告文件不存在: {report_path}")
            return 2
        if not decisions_path.is_file():
            print(f"[错误] 决策文件不存在: {decisions_path}")
            return 2
        if not args.replace_output and mapping_out_path.exists():
            print(f"[错误] 映射输出文件已存在，若要覆盖请使用 --replace-output: {mapping_out_path}")
            return 1
        from lib.format_review import compile_role_mapping
        try:
            with open(report_path, "r", encoding="utf-8") as f:
                report_data = json.load(f)
            with open(decisions_path, "r", encoding="utf-8") as f:
                decisions_data = json.load(f)
            mapping = compile_role_mapping(report_data, decisions_data, output_path=mapping_out_path)
            print(f"角色映射编译成功: {mapping_out_path}")
            return 0
        except Exception as exc:
            print(f"[错误] 角色映射编译失败: {exc}")
            return 1

    if args.preview_format or args.render_preview:
        if not args.preview_out:
            parser.error("--preview-format/--render-preview 需要指定 --preview-out")
        format_path = Path(args.preview_format or args.render_preview).resolve()
        preview_path = Path(args.preview_out).resolve()
        if not format_path.is_file():
            print(f"[错误] 格式包文件不存在: {format_path}")
            return 2
        if preview_path.exists() and not args.replace_output:
            print(f"[错误] 预览输出已存在，若要覆盖请使用 --replace-output: {preview_path}")
            return 1
        from lib.preview import build_preview
        try:
            metadata_path = Path(args.preview_meta).resolve() if args.preview_meta else None
            metadata = build_preview(
                format_path,
                preview_path,
                engine=args.preview_engine,
                metadata_path=metadata_path,
            )
            print(f"预览生成成功: {preview_path}")
            print(f"  - 引擎: {metadata['engine']}（近似预览）")
            if metadata_path:
                print(f"  - 预览证据: {metadata_path}")
            return 0
        except Exception as exc:
            print(f"[错误] 格式预览失败: {exc}")
            return 1

    if args.migrate_manifest:
        if not args.manifest_out:
            parser.error("--migrate-manifest 需要指定 --manifest-out")
        from lib.manifest_migration import migrate_manifest_file
        try:
            migrated = migrate_manifest_file(
                args.migrate_manifest,
                args.manifest_out,
                replace=args.replace_output,
            )
            print(f"manifest 迁移成功: {Path(args.manifest_out).resolve()}")
            print(f"  - schema_version: {migrated.get('schema_version')}")
            return 0
        except Exception as exc:
            print(f"[错误] manifest 迁移失败: {exc}")
            return 1

    if args.all and args.source:
        parser.error("--all 不能与 --source 同时使用")
    if args.plan and args.all:
        parser.error("--plan 一次只能检查一个项目或 --source")
    if args.doctor and args.all:
        parser.error("--doctor 一次只能检查一个项目或 --source")

    manifest_path = Path(args.manifest).resolve() if args.manifest else None
    output_dir = Path(args.output_dir).resolve() if args.output_dir else None
    override_paths = [Path(path).resolve() for path in args.override]

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
        if (manifest_path or override_paths) and not source_dir:
            parser.error("带配置的 --doctor 需要 --source 或 --project")
        return run_doctor(source_dir, manifest_path, override_paths)
    if args.plan:
        if not source_dir:
            print("[错误] --plan 需要 --source 或 --project。")
            return 2
        try:
            print_plan(source_dir, manifest_path, override_paths)
            return 0
        except (ConfigError, BuildError) as exc:
            print(f"[错误] 构建计划无效: {exc}")
            return 1

    try:
        if args.all or (args.project and args.project.lower() == "all"):
            projects = discover_all_input_projects(DEFAULT_INPUT_DIR) + discover_standalone_input_files(DEFAULT_INPUT_DIR)
            if not projects:
                raise BuildError(f"未在 {DEFAULT_INPUT_DIR} 找到有效输入材料。")
            failures = []
            for project in projects:
                try:
                    project_output = (output_dir / project.stem) if output_dir else None
                    run_synthesis(project, project_output, manifest_path, override_paths, args.keep_work)
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
        run_synthesis(source_dir, output_dir, manifest_path, override_paths, args.keep_work)
        return 0
    except (BuildError, OfficeExportError, RuntimeError) as exc:
        print(f"[未发布] 构建失败: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
