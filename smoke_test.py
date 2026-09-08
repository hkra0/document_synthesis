#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按构建计划核验已发布交付物；默认包含 Word 原生分页 QA。

--structure-only 可在无 Office 环境检查配置、来源、交付清单、书签和文字，
但不证明页码或版式正确。脚本不修改源文件或已发布的 DOCX。
"""

import argparse
import re
import tempfile
from types import SimpleNamespace
from pathlib import Path
from typing import Optional

from docx import Document

from lib.engine import UnifiedSynthesizer
from synthesize import (
    DEFAULT_INPUT_DIR,
    discover_all_input_projects,
    discover_standalone_input_files,
    find_project_matches,
)

ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = ROOT / "output"


def declared_delivery_paths(plan, output_dir):
    """只接受计划中的文件，不用旧文件名猜测或替代缺失交付物。"""
    paths = [(spec, Path(output_dir) / spec["filename"]) for spec in plan["documents"]]
    missing = [str(path) for _, path in paths if not path.is_file()]
    if missing:
        raise ValueError("缺少声明的交付物: " + ", ".join(missing))
    return paths


def _source_docx_paths(source, nodes):
    if source.is_file():
        return [source] if source.suffix.lower() == ".docx" else []
    paths = set()
    for node in nodes:
        if not node.get("file"):
            continue
        path = (source / node["file"]).resolve()
        if not path.is_file():
            raise ValueError(f"源文件不存在: {path}")
        if path.suffix.lower() == ".docx":
            paths.add(path)
    return sorted(paths)


def _normalized_text(text):
    return re.sub(r"\s+", "", text)


def _check_body_text(path, source_docx, nodes, expected_inventory=None):
    # smoke 与 engine 共用 PreparedBuild 期望清单/严格验证器；无构建计划时
    # 用传入的源文件构造等价的只读准备视图，兼容历史单元测试调用。
    from lib.content_integrity import build_expected_inventory, verify_content_integrity, ContentIntegrityError
    if expected_inventory is None and not source_docx:
        print("  正文文字核验: 0 段")
        return
    if expected_inventory is None:
        from types import SimpleNamespace
        expected_inventory = build_expected_inventory(
            SimpleNamespace(
                source_order=tuple(str(source) for source in source_docx),
                source_docx_path=source_docx[0] if len(source_docx) == 1 else None,
                nodes=tuple(nodes),
                parts={},
                config=SimpleNamespace(regions={}),
            ),
            spec={"id": "smoke", "parts": ["body"]},
        )
    if not expected_inventory.source_order:
        raise ValueError("没有可用于内容完整性核验的 DOCX 源节点清单")
    checked = len(expected_inventory.semantic.paragraphs)
    try:
        verify_content_integrity(expected_inventory, path)
    except ContentIntegrityError as cie:
        names = ", ".join(Path(item).name for item in expected_inventory.source_order)
        raise ValueError(f"正文文字完整性核验未通过: {names}: {cie}") from cie

    print(f"  正文文字核验: {checked} 段")


def _metadata_contract_violations(metadata, plan, prepared, paths):
    """Compare every recorded build hash with the current plan and files."""
    from lib.content_integrity import compute_file_sha256

    violations = []
    if not isinstance(metadata, dict):
        return ["build-metadata.json 根节点不是对象"]

    stored_sources = metadata.get("source_hashes")
    current_sources = plan.get("source_hashes", {})
    if stored_sources != current_sources:
        violations.append("源文件 SHA-256 与当前 plan 不一致")
    if prepared is not None and getattr(prepared, "source_hashes", {}) != current_sources:
        violations.append("PreparedBuild 与当前 plan 的源文件 SHA-256 不一致")

    if prepared is not None:
        configuration = metadata.get("configuration")
        if not isinstance(configuration, dict):
            violations.append("build-metadata.json 缺少 configuration 契约")
        else:
            stored_config = configuration.get("config_contract_sha256")
            if stored_config != prepared.config_hash:
                violations.append("配置契约 SHA-256 与当前 PreparedBuild 不一致")
        format_info = metadata.get("format", {})
        stored_format = format_info.get("source_sha256") if isinstance(format_info, dict) else None
        if prepared.format_hash and stored_format != prepared.format_hash:
            violations.append("格式包 SHA-256 与当前 PreparedBuild 不一致")

    delivery_records = metadata.get("deliveries", {})
    if not isinstance(delivery_records, dict):
        violations.append("build-metadata.json 缺少 deliveries 契约")
        delivery_records = {}
    expected_delivery_ids = {spec["id"] for spec, _ in paths}
    if set(delivery_records) != expected_delivery_ids:
        violations.append("deliveries 契约与当前声明交付物集合不一致")
    for spec, path in paths:
        record = delivery_records.get(spec["id"])
        if not isinstance(record, dict):
            violations.append(f"交付物 {spec['id']} 缺少元数据记录")
            continue
        actual_hash = compute_file_sha256(path)
        if record.get("sha256") != actual_hash:
            violations.append(f"交付物 {spec['id']} SHA-256 与元数据不一致")
        if record.get("filename") != spec.get("filename"):
            violations.append(f"交付物 {spec['id']} 文件名与当前 plan 不一致")
        if record.get("parts") != spec.get("parts"):
            violations.append(f"交付物 {spec['id']} 部件清单与当前 plan 不一致")
    return violations


def test_project(
    source_dir: Path,
    output_dir: Optional[Path] = None,
    manifest_path: Optional[Path] = None,
    override_paths=None,
    structure_only=False,
) -> bool:
    """所有来源策略共用计划和交付验证；任一声明交付物缺失即失败。"""
    try:
        import json
        from lib.delivery import validate_delivery, validate_delivery_structure

        source = Path(source_dir).resolve()
        plan = UnifiedSynthesizer.plan(source, manifest_path, override_paths=override_paths)
        out_dir = Path(output_dir) if output_dir else DEFAULT_OUTPUT_DIR / plan["project"]
        print(f"\n《{plan['project']}》交付检查: {out_dir}")
        for warning in plan.get("warnings", []):
            print(f"  [提示] {warning}")

        # 0. 读取交付元数据；具体哈希核对要等 PreparedBuild 和声明文件
        # 路径就绪后进行，不能只依赖元数据中的历史 passed 字段。
        meta_path = out_dir / "build-metadata.json"
        metadata = None
        if meta_path.is_file():
            try:
                metadata = json.loads(meta_path.read_text(encoding="utf-8"))
            except Exception as meta_exc:
                raise ValueError(f"读取交付元数据失败: {meta_exc}") from meta_exc

        paths = declared_delivery_paths(plan, out_dir)
        nodes = plan["nodes"]
        source_docx = _source_docx_paths(source, nodes)
        maps = {}
        cfg = None
        prepared_for_verification = None
        try:
            from lib.config import load_project_config
            config_dir = source if source.is_dir() else source.parent
            cfg = load_project_config(
                config_dir,
                manifest_path,
                project_name=plan["project"],
                override_paths=override_paths,
            )
            if getattr(cfg, "schema_version", 0) >= 3:
                prepared_for_verification = UnifiedSynthesizer.prepare_build(
                    source, manifest_path, override_paths=override_paths
                )
        except Exception:
            if cfg is not None and getattr(cfg, "schema_version", 0) >= 3:
                raise

        if metadata is not None:
            violations = _metadata_contract_violations(
                metadata, plan, prepared_for_verification, paths
            )
            if violations:
                raise ValueError("交付元数据契约核验未通过: " + "; ".join(violations))
        elif cfg is not None and getattr(cfg, "schema_version", 0) >= 3:
            raise ValueError("v3 交付缺少必需的 build-metadata.json")
        # plan() exposes a JSON-serializable parts registry; delivery validators
        # historically consume attribute-style DocumentPart objects.
        parts_registry = {
            key: SimpleNamespace(**value) if isinstance(value, dict) else value
            for key, value in plan.get("parts", {}).items()
        }
        if cfg is not None and getattr(cfg, "schema_version", 0) >= 3:
            parts_registry = dict(getattr(cfg, "parts", {}) or {})
        from lib.content_integrity import build_expected_inventory

        def has_content_part(parts):
            return "body" in parts or any(
                parts_registry.get(p) and parts_registry[p].kind == "content" for p in parts
            )

        prepared_view = prepared_for_verification or SimpleNamespace(
            source_order=tuple(str(source) for source in source_docx),
            source_docx_path=source_docx[0] if len(source_docx) == 1 else None,
            nodes=tuple(nodes),
            parts=parts_registry,
            config=SimpleNamespace(regions={}),
        )
        expected_by_delivery = {
            spec["id"]: build_expected_inventory(prepared_view, spec=spec)
            for spec, _ in paths
            if has_content_part(spec["parts"]) and source_docx
        }

        ordered = sorted(paths, key=lambda pair: not has_content_part(pair[0]["parts"]))
        with tempfile.TemporaryDirectory(prefix="document-synthesis-smoke-") as temporary:
            for index, (spec, path) in enumerate(ordered):
                reference = spec.get("toc", {}).get("reference")
                reference_map = maps.get(reference) if reference != spec["id"] else None
                print(f"  检查 {spec['id']}: {path.name} ({' + '.join(spec['parts'])})")
                if has_content_part(spec["parts"]):
                    _check_body_text(
                        path, source_docx, nodes,
                        expected_inventory=expected_by_delivery.get(spec["id"]),
                    )
                if structure_only:
                    validate_delivery_structure(
                        path, spec, nodes, reference_map=reference_map, parts_registry=parts_registry
                    )
                else:
                    maps[spec["id"]] = validate_delivery(
                        path, spec, nodes, Path(temporary) / f"delivery-{index}.pdf",
                        reference_map=reference_map,
                        parts_registry=parts_registry,
                    )

        # 格式有效性核验 (针对 v3 或配置了 resolved_format 的项目)。v3
        # 必须复用准备阶段的上下文；任何必需核验异常都应使 smoke 失败。
        if cfg is not None and getattr(cfg, "resolved_format", None) and getattr(cfg, "schema_version", 0) >= 3:
            from lib.content_integrity import verify_delivery_format, verify_generated_content
            from lib.verification_contracts import build_verification_context
            for spec, path in paths:
                if prepared_for_verification is None:
                    raise ValueError("v3 smoke 缺少冻结的 PreparedBuild 验证上下文")
                context = build_verification_context(
                    prepared_for_verification, None, spec, path
                )
                if metadata is not None:
                    recorded_hash = metadata["deliveries"][spec["id"]]["sha256"]
                    if context.delivery_sha256 != recorded_hash:
                        raise ValueError(f"{spec['id']} VerificationContext 与当前交付物哈希不一致")
                generated = verify_generated_content(path, context.expected_delivery)
                if not generated["passed"]:
                    raise ValueError(f"生成内容核验未通过: {generated['failures']}")
                fmt_rep = verify_delivery_format(
                    path, cfg.resolved_format, verification_context=context
                )
                if not fmt_rep.passed:
                    raise ValueError("交付格式核验未通过: " + "; ".join(fmt_rep.violations[:3]))
                print(f"  交付格式核验: {fmt_rep.verified_count} 个块符合格式包规格 ({', '.join(fmt_rep.verified_roles)})")
        elif cfg is not None and getattr(cfg, "resolved_format", None):
            from lib.content_integrity import verify_delivery_format
            for spec, path in paths:
                fmt_rep = verify_delivery_format(path, cfg.resolved_format)
                if fmt_rep.passed:
                    print(f"  交付格式核验: {fmt_rep.verified_count} 个块符合格式包规格 ({', '.join(fmt_rep.verified_roles)})")
                else:
                    print(f"  [提示] 交付格式核验提示: {len(fmt_rep.violations)} 项属性偏差")

        if structure_only:
            print("[通过] 结构检查完成；未执行 Word 导出、目录页码和页级版式核验。")
        else:
            print("[通过] 所有声明交付物的结构、Word 页码与版式检查完成。")
        return True
    except Exception as exc:
        print(f"[未通过] {exc}")
        return False


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", default="all", help="项目名称或路径（默认 all）")
    parser.add_argument("--source", help="直接指定目录或单一 DOCX 来源")
    parser.add_argument("--all", action="store_true", help="检查全部目录与配置化单文件项目")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT_DIR)
    parser.add_argument("--output-dir", type=Path, help="单项目为实际交付目录；--all 时为输出根目录")
    parser.add_argument("--manifest", type=Path, help="与构建时相同的基础配置")
    parser.add_argument("--override", type=Path, action="append", default=[], help="按顺序应用构建时的覆写文件，可重复")
    parser.add_argument("--structure-only", action="store_true", help="仅检查结构与文字，不验证 Word 页码或版式")
    args = parser.parse_args()
    if args.all and args.source:
        parser.error("--all 不能与 --source 同用")
    all_projects = args.all or (not args.source and args.project == "all")
    if all_projects and (args.manifest or args.override):
        parser.error("--manifest/--override 仅用于单项目，避免对全部项目误用同一配置")
    try:
        if all_projects:
            projects = discover_all_input_projects(args.input_dir) + discover_standalone_input_files(args.input_dir)
        else:
            projects = find_project_matches(args.input_dir, args.source or args.project)
        if not projects:
            parser.error("没有找到可检查的项目")
        if not all_projects and len(projects) != 1:
            parser.error("项目名称有歧义，请使用完整路径: " + ", ".join(str(p) for p in projects))
        results = []
        for project in projects:
            output = (args.output_dir / project.stem) if all_projects and args.output_dir else args.output_dir
            results.append(test_project(project, output, args.manifest, args.override, args.structure_only))
        return 0 if all(results) else 1
    except Exception as exc:
        print(f"[未通过] {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
