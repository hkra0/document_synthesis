"""Shared final-artifact validation and bounded pagination convergence."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn

from .composition import BOUNDARIES, assemble_document, toc_nodes
from .pagination import inspect_document, page_records_from_map, validate_document_structure
from .pagination_types import format_page_number
from .qa import OfficeExportError, run_qa_assertions
from .field_updater import required_pageref_targets, verify_final_field_caches


CHECK_STATUSES = frozenset({"not_run", "passed", "failed", "unsupported"})


@dataclass(frozen=True)
class CheckResult:
    """结果明确的单项门禁，不允许把未执行状态当成通过。"""

    name: str
    status: str
    required: bool = True
    diagnostics: Tuple[str, ...] = ()
    evidence: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        if self.status not in CHECK_STATUSES:
            raise ValueError(f"未知门禁状态: {self.status}")
        object.__setattr__(self, "diagnostics", tuple(str(item) for item in self.diagnostics))
        object.__setattr__(self, "evidence", dict(self.evidence))

    @property
    def publishable(self) -> bool:
        return not self.required or self.status == "passed"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "status": self.status,
            "required": self.required,
            "diagnostics": list(self.diagnostics),
            "evidence": self.evidence,
        }


@dataclass(frozen=True)
class DeliveryReport:
    """单个交付物的所有门禁结果及其未验证属性。"""

    delivery_id: str
    filename: str
    checks: Tuple[CheckResult, ...]
    source_nodes: Tuple[str, ...] = ()
    unverified_attributes: Tuple[str, ...] = ()

    @property
    def publishable(self) -> bool:
        return all(check.publishable for check in self.checks)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "delivery_id": self.delivery_id,
            "filename": self.filename,
            "publishable": self.publishable,
            "checks": [check.to_dict() for check in self.checks],
            "source_nodes": list(self.source_nodes),
            "unverified_attributes": list(self.unverified_attributes),
        }


class DeliveryGateError(RuntimeError):
    """交付物存在未通过的必需门禁。"""

    def __init__(self, message: str, reports: Iterable[DeliveryReport]):
        super().__init__(message)
        self.reports = tuple(reports)


def assert_publishable(reports: Iterable[DeliveryReport]) -> None:
    """聚合所有交付物的门禁结果后再决定是否允许发布。"""
    reports = tuple(reports)
    failures: List[str] = []
    for report in reports:
        for check in report.checks:
            if check.required and check.status != "passed":
                detail = "; ".join(check.diagnostics) or "无附加诊断"
                failures.append(
                    f"{report.delivery_id} ({report.filename}) / {check.name} "
                    f"= {check.status}: {detail}"
                )
    if failures:
        raise DeliveryGateError("交付门禁未通过:\n" + "\n".join(f"  - {item}" for item in failures), reports)


def summarize_delivery_reports(reports: Iterable[DeliveryReport]) -> Dict[str, Any]:
    """把逐交付物结果转换成 build-metadata.json 中稳定的 QA 摘要。"""
    reports = tuple(reports)

    def overall(check_name: str) -> str:
        checks = [
            check for report in reports for check in report.checks
            if check.name == check_name and check.required
        ]
        if not checks:
            return "not_run"
        statuses = {check.status for check in checks}
        if "failed" in statuses:
            return "failed"
        if "unsupported" in statuses:
            return "unsupported"
        if "not_run" in statuses:
            return "not_run"
        return "passed"

    violations = [
        diagnostic
        for report in reports
        for check in report.checks
        if check.status == "failed"
        for diagnostic in check.diagnostics
    ]
    unverified = sorted({
        attribute
        for report in reports
        for attribute in report.unverified_attributes
    })
    return {
        "content_integrity": overall("content_integrity"),
        "format_verification": overall("format_verification"),
        "unverified_attributes": unverified,
        "violations": violations,
        "delivery_reports": {
            report.delivery_id: report.to_dict()
            for report in reports
        },
    }


def required_bookmarks(spec, nodes, parts_registry=None):
    names = [BOUNDARIES[part] for part in spec["parts"]]
    has_content = "body" in spec["parts"] or (parts_registry and any(parts_registry.get(p) and parts_registry[p].kind == "content" for p in spec["parts"]))
    if has_content:
        names.extend(node["bookmark_name"] for node in toc_nodes(nodes))
    return names


def _check_page_map(page_map, names, *, require_observed_labels=False):
    for name in names:
        record = page_map.get(name)
        if not isinstance(record, dict) or any(
            type(record.get(key)) is not int or record[key] < 1
            for key in ("physical_page", "printed_page")
        ):
            raise OfficeExportError(f"缺失或无效的最终书签页码: {name}")
        if record.get("label_required") is False:
            continue
        observed = record.get("observed_label")
        if require_observed_labels and not isinstance(observed, str):
            record["label_verified"] = False
            record["verification_status"] = "unverified"
            raise OfficeExportError(
                f"同次导出 PDF 未能提取书签所在页的页码标签: {name} / 预期 {record.get('expected_label')!r}"
            )
        if observed is not None:
            expected = record.get("expected_label")
            if not isinstance(expected, str) or observed != expected:
                raise OfficeExportError(
                    f"实测页码标签与预期不一致: {name} / 预期 {expected!r} / 实测 {observed!r}"
                )
            record["label_verified"] = True
            record["verification_status"] = "verified"


def validate_delivery_structure(path, spec, nodes, reference_map=None, parts_registry=None):
    validate_document_structure(path, required_bookmarks(spec, nodes, parts_registry))
    doc = Document(str(path))
    actual_parts = {b.get(qn("w:name")) for b in doc.element.body.iter(qn("w:bookmarkStart"))}
    for part in spec["parts"]:
        name = BOUNDARIES[part]
        if name not in actual_parts:
            raise OfficeExportError(f"交付物组成与配置不一致: {spec['id']} / {part}")
    for legacy_part in ("cover", "toc", "body"):
        if legacy_part not in spec["parts"]:
            legacy_bm = BOUNDARIES[legacy_part]
            if legacy_bm in actual_parts:
                raise OfficeExportError(f"交付物组成与配置不一致: {spec['id']} / {legacy_part}")
    entries = {}
    for b in doc.element.body.iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name"), "")
        if name.startswith("_Synth_entry_"):
            entries[name] = b.getparent()
    has_toc = "toc" in spec["parts"] or (parts_registry and any(parts_registry.get(p) and parts_registry[p].kind == "generated_toc" for p in spec["parts"]))
    expected = toc_nodes(nodes) if has_toc else []
    if len(entries) != len(expected):
        raise OfficeExportError(f"目录条目数与大纲不一致: {spec['id']}")
    for index, node in enumerate(expected):
        p = entries.get(f"_Synth_entry_{index:05d}")
        if p is None:
            raise OfficeExportError("目录条目标识缺失。")
        texts = [t.text or "" for t in p.iter(qn("w:t"))]
        if len(texts) < 2 or "".join(texts[:-1]) != (node.get("toc_title") or node["title"]):
            raise OfficeExportError(f"目录标题不一致: {node['title']}")
        links = list(p.iter(qn("w:hyperlink")))
        name = node["bookmark_name"]
        toc_cfg = spec.get("toc", {})
        links_mode = toc_cfg.get("links", "internal")
        if links_mode == "internal":
            if len(links) != 1 or links[0].get(qn("w:anchor")) != name:
                raise OfficeExportError(f"目录跳转目标不一致: {name}")
        elif links:
            raise OfficeExportError("打印目录不应包含本文件跳转。")
        if reference_map is not None:
            ref_rec = reference_map.get(name, {})
            expected_page_str = ref_rec.get("expected_label") or str(ref_rec.get("printed_page", ""))
            if name not in reference_map or texts[-1] != expected_page_str:
                raise OfficeExportError(f"目录页码与最终 Word 书签不一致: {node['title']}")

    # 验证脚注引用完整性（禁止悬空引用）
    fn_refs = [el.get(qn("w:id")) for el in doc.element.body.iter(qn("w:footnoteReference")) if el.get(qn("w:id"))]
    if fn_refs:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml import parse_xml
        try:
            fn_part = doc.part.part_related_by(RT.FOOTNOTES)
            fn_root = parse_xml(fn_part.blob)
            available_ids = {fn.get(qn("w:id")) for fn in fn_root.findall(qn("w:footnote"))}
            for fid in fn_refs:
                if fid not in available_ids:
                    raise OfficeExportError(f"文档包含未定义的脚注引用: {fid}")
        except KeyError:
            raise OfficeExportError(f"文档包含脚注引用，但缺失 footnotes.xml 部件: {spec['id']}")


def _enrich_page_map_labels(page_map, spec, nodes, config_or_parts=None, page_sequences=None, pdf_path=None):
    from .pagination_types import format_page_number
    parts_reg = getattr(config_or_parts, "parts", config_or_parts) if config_or_parts else {}
    seqs = getattr(config_or_parts, "page_sequences", page_sequences) if config_or_parts else (page_sequences or {})

    part_fmt = {}
    for pid in spec.get("parts", []):
        pdef = parts_reg.get(pid) if parts_reg else None
        seq_id = getattr(pdef, "page_sequence", None)
        sdef = seqs.get(seq_id) if seq_id else None
        part_fmt[pid] = sdef.format if sdef else "decimal"

    observed_labels = {}
    if pdf_path is not None:
        from .pagination import extract_pdf_page_labels
        observed_labels = extract_pdf_page_labels(pdf_path)

    for name, record in page_map.items():
        label_required = True
        for pid in spec.get("parts", []):
            if BOUNDARIES.get(pid) != name:
                continue
            pdef = parts_reg.get(pid) if parts_reg else None
            label_required = bool(getattr(pdef, "page_sequence", None))
            break
        record["label_required"] = label_required
        if not label_required:
            record["expected_label"] = ""
            record["label_verified"] = False
            record["verification_status"] = "unverified"
            continue
        if "expected_label" not in record:
            target_fmt = "decimal"
            for pid, fmt in part_fmt.items():
                if BOUNDARIES.get(pid) == name:
                    target_fmt = fmt
                    break
            else:
                for pid in spec.get("parts", []):
                    pdef = parts_reg.get(pid) if parts_reg else None
                    if pid == "body" or (pdef and pdef.kind == "content"):
                        target_fmt = part_fmt.get(pid, "decimal")
                        break
            record["expected_label"] = format_page_number(record["printed_page"], target_fmt)
        if record.get("observed_label") is None:
            record["observed_label"] = observed_labels.get(record["physical_page"])
        if record.get("observed_label") is not None:
            record["label_verified"] = (record["observed_label"] == record["expected_label"])
            record["verification_status"] = "verified" if record["label_verified"] else "mismatch"
        else:
            record["label_verified"] = False
            record["verification_status"] = "unverified"


def validate_measured_delivery(path, spec, nodes, pdf_path, page_map, reference_map=None, parts_registry=None, page_sequences=None):
    _enrich_page_map_labels(page_map, spec, nodes, parts_registry, page_sequences, pdf_path=pdf_path)
    require_observed_labels = bool(page_sequences) or any(
        getattr(parts_registry.get(part), "page_sequence", None)
        for part in spec.get("parts", [])
        if parts_registry
    )
    _check_page_map(
        page_map,
        required_bookmarks(spec, nodes, parts_registry),
        require_observed_labels=require_observed_labels,
    )
    page_records = page_records_from_map(
        page_map,
        part_boundaries={part: BOUNDARIES[part] for part in spec["parts"]},
        part_specs=parts_registry or {},
        sequences=page_sequences or {},
    )
    if any(record.verification_status == "mismatch" for record in page_records.values()):
        raise OfficeExportError(f"交付物存在未通过的页码标签记录: {spec['id']}")
    reference = reference_map if reference_map is not None else page_map
    has_toc = "toc" in spec["parts"] or (parts_registry and any(parts_registry.get(p) and parts_registry[p].kind == "generated_toc" for p in spec["parts"]))
    validate_delivery_structure(path, spec, nodes, reference if has_toc else None, parts_registry=parts_registry)
    starts = [(page_map[BOUNDARIES[part]]["physical_page"], part) for part in spec["parts"]]
    if starts[0][0] != 1:
        raise OfficeExportError(f"文档各部分没有按配置独立起页: {spec['id']}")
    for (start, _), (next_start, next_part) in zip(starts, starts[1:]):
        part_def = parts_registry.get(next_part) if parts_registry else None
        section_type = getattr(part_def, "section_type", "nextPage") if part_def else "nextPage"
        if section_type == "continuous":
            invalid = next_start < start
        else:
            invalid = next_start <= start
        if invalid:
            raise OfficeExportError(f"文档各部分没有按配置顺序起始: {spec['id']} / {next_part}")
        if section_type == "oddPage" and next_start % 2 != 1:
            raise OfficeExportError(f"oddPage 部件没有落在物理奇数页: {spec['id']} / {next_part}")
        if section_type == "evenPage" and next_start % 2 != 0:
            raise OfficeExportError(f"evenPage 部件没有落在物理偶数页: {spec['id']} / {next_part}")

    # 识别由 oddPage 引起的有意空白页白名单
    allowed_blank_pages = set()
    for index, (start, part) in enumerate(starts[:-1]):
        next_start, next_part = starts[index + 1]
        part_def = parts_registry.get(next_part) if parts_registry else None
        if part_def and getattr(part_def, "section_type", None) == "oddPage":
            # A single physical blank page is the only implicit oddPage gap
            # accepted. Larger gaps need an explicit part/diagnostic.
            if next_start == start + 2 and (start + 1) % 2 == 0:
                allowed_blank_pages.add(start + 1)

    content_parts = [
        p for p in spec["parts"]
        if p == "body" or (parts_registry and parts_registry.get(p) and parts_registry[p].kind == "content")
    ]
    if content_parts:
        first_content = content_parts[0]
        first = page_map[BOUNDARIES[first_content]]
        has_custom_seq = any(
            parts_registry.get(p) and parts_registry[p].page_sequence
            for p in content_parts
        ) if parts_registry else False
        previous_by_sequence = {}
        for part_name in content_parts:
            record = page_map[BOUNDARIES[part_name]]
            part_def = parts_registry.get(part_name) if parts_registry else None
            sequence_id = getattr(part_def, "page_sequence", None) if part_def else None
            sequence = page_sequences.get(sequence_id) if sequence_id else None
            if sequence:
                if sequence.start is not None and record["printed_page"] != sequence.start:
                    raise OfficeExportError(
                        f"页码序列起始值不符: {part_name} / 预期 {sequence.start} / 实测 {record['printed_page']}"
                    )
                expected_label = format_page_number(record["printed_page"], sequence.format)
                if record.get("expected_label") != expected_label:
                    raise OfficeExportError(
                        f"页码标签格式不符: {part_name} / 预期 {expected_label} / 当前 {record.get('expected_label')}"
                    )
                previous = previous_by_sequence.get(sequence_id)
                if previous and sequence.start is None:
                    previous_page, previous_value = previous
                    expected_value = previous_value + max(1, record["physical_page"] - previous_page)
                    if record["printed_page"] != expected_value:
                        raise OfficeExportError(
                            f"页码序列未连续: {part_name} / 预期 {expected_value} / 实测 {record['printed_page']}"
                        )
                previous_by_sequence[sequence_id] = (record["physical_page"], record["printed_page"])
        if not has_custom_seq:
            if first["printed_page"] != 1:
                raise OfficeExportError("正文必须从第 1 页开始编号。")
            for node in toc_nodes(nodes):
                if node["bookmark_name"] in page_map:
                    page = page_map[node["bookmark_name"]]
                    if page["physical_page"] < first["physical_page"] or page["printed_page"] != page["physical_page"] - first["physical_page"] + 1:
                        raise OfficeExportError(f"正文编号不连续: {node['title']}")
    import pymupdf
    with pymupdf.open(str(pdf_path)) as pdf:
        page_roles = {}
        for index, (start, part) in enumerate(starts):
            end = starts[index + 1][0] if index + 1 < len(starts) else len(pdf) + 1
            page_roles.update({page: part for page in range(start, end)})
    if not run_qa_assertions(str(pdf_path), ignore_front_pages=0, page_roles=page_roles, allowed_blank_pages=allowed_blank_pages):
        raise OfficeExportError(f"交付物未通过页级 QA: {spec['filename']}")
    return page_map


def validate_delivery(path, spec, nodes, pdf_path, reference_map=None, parts_registry=None, page_sequences=None):
    """Read-only validation of a published document; used by smoke tests too."""
    validate_delivery_structure(path, spec, nodes, parts_registry=parts_registry)
    measured_names = list(dict.fromkeys(
        required_bookmarks(spec, nodes, parts_registry) + required_pageref_targets(path)
    ))
    pages = inspect_document(path, pdf_path, measured_names)
    has_toc = "toc" in spec["parts"] or (parts_registry and any(parts_registry.get(p) and parts_registry[p].kind == "generated_toc" for p in spec["parts"]))
    has_content = "body" in spec["parts"] or (parts_registry and any(parts_registry.get(p) and parts_registry[p].kind == "content" for p in spec["parts"]))
    if has_toc and not has_content and reference_map is None:
        raise OfficeExportError("独立目录核验需要其引用文档的最终页码。")
    result = validate_measured_delivery(
        path, spec, nodes, pdf_path, pages, reference_map,
        parts_registry=parts_registry, page_sequences=page_sequences,
    )
    field_targets = required_pageref_targets(path)
    if field_targets:
        field_map = reference_map if reference_map is not None else pages
        field_cache = verify_final_field_caches(path, field_map or {})
        if not field_cache["passed"]:
            detail = "; ".join(
                f"{item['instruction']}: 期望 {item['expected']!r}, 实际 {item['actual']!r}"
                for item in field_cache["failures"]
            )
            raise OfficeExportError(f"最终 DOCX 的 PAGEREF 缓存与同次测量不一致: {detail}")
    return result


def build_deliveries(
    config,
    body_path,
    nodes,
    run_dir,
    max_passes=3,
    parts_registry=None,
    page_maps_out=None,
    prepared=None,
):
    """Reassemble cheap prefixes, never re-render sources to make output variants."""
    published_maps, staged = {}, {}
    parts_registry = parts_registry or getattr(config, "parts", {})
    page_sequences = getattr(config, "page_sequences", {})
    if prepared is not None:
        # Recheck the frozen source/format/RoleMap/template inputs immediately
        # before the first assembly.  The plan owns the paths and hashes; the
        # delivery loop must not rediscover a different cover template.
        prepared.verify_inputs_unchanged()
    def has_content_check(spec):
        return "body" in spec["parts"] or any(parts_registry.get(p) and parts_registry[p].kind == "content" for p in spec["parts"])
    specs = sorted(config.documents, key=lambda spec: not has_content_check(spec))
    for spec in specs:
        path = Path(run_dir) / spec["filename"]
        pdf_path = Path(run_dir) / f"qa-{spec['id']}.pdf"
        reference = spec.get("toc", {}).get("reference")
        external = reference is not None and reference != spec["id"]
        reference_map = published_maps.get(reference) if external else None
        if external and reference_map is None:
            raise OfficeExportError(f"目录引用没有就绪的正文交付物: {reference}")
        # Keep the full measured record while using its label as the TOC
        # display value.  FieldUpdater consumes this same map for PAGEREF and
        # must never receive a bare label string.
        guess = {
            key: dict(value)
            for key, value in (reference_map or {}).items()
        }
        has_toc = "toc" in spec["parts"] or any(parts_registry.get(p) and parts_registry[p].kind == "generated_toc" for p in spec["parts"])
        # PAGEREF dependencies are independent of whether this delivery has a
        # source content part.  A cover/TOC-only artifact can still contain a
        # supported field and must not skip its measurement or cache gate.
        field_target_names = required_pageref_targets(body_path) if Path(body_path).is_file() else []
        measured_names = list(dict.fromkeys(
            required_bookmarks(spec, nodes, parts_registry) + field_target_names
        ))
        odd_page_padding = set()
        convergence_rounds = []
        for attempt in range(max_passes):
            assemble_document(
                config,
                spec,
                body_path,
                nodes,
                guess,
                path,
                odd_page_padding=odd_page_padding,
            )
            if prepared is not None and getattr(config, "schema_version", 1) >= 3:
                from .verification_contracts import build_expected_delivery, stamp_output_occurrences

                expected_delivery = build_expected_delivery(prepared, spec)
                if expected_delivery.ordered_nodes:
                    stamp_output_occurrences(path, expected_delivery)
            validate_delivery_structure(path, spec, nodes, parts_registry=parts_registry)
            page_map = inspect_document(path, pdf_path, measured_names)
            # Templates or imported front matter may introduce a PAGEREF that
            # was not present in the reusable body. Discover those targets
            # from the assembled file and measure them before accepting this
            # round.
            assembled_targets = required_pageref_targets(path)
            new_targets = [item for item in assembled_targets if item not in field_target_names]
            if new_targets:
                field_target_names = list(dict.fromkeys(field_target_names + new_targets))
                measured_names = list(dict.fromkeys(measured_names + new_targets))
                page_map = inspect_document(path, pdf_path, measured_names)
            _enrich_page_map_labels(page_map, spec, nodes, config)
            _check_page_map(page_map, required_bookmarks(spec, nodes, parts_registry))
            # Keep the correction local to parts whose measured boundary is
            # actually even.  This preserves the declared section type while
            # making imported-package assembly deterministic on Word versions
            # that suppress oddPage at this boundary.
            for _, next_part in zip(spec["parts"], spec["parts"][1:]):
                part_def = parts_registry.get(next_part) if parts_registry else None
                if getattr(part_def, "section_type", "nextPage") != "oddPage":
                    continue
                boundary = page_map.get(BOUNDARIES[next_part], {})
                if boundary.get("physical_page", 0) % 2 == 0:
                    if next_part not in odd_page_padding:
                        odd_page_padding.add(next_part)
                        continue
                    raise OfficeExportError(f"oddPage 部件没有落在物理奇数页: {spec['id']} / {next_part}")
            measured = reference_map if external else page_map
            if field_target_names:
                _check_page_map(measured, field_target_names)
            if has_toc:
                _check_page_map(measured, [node["bookmark_name"] for node in toc_nodes(nodes)])
            exact_names = list(dict.fromkeys(
                ([node["bookmark_name"] for node in toc_nodes(nodes)] if has_toc else [])
                + field_target_names
            ))
            exact = {
                name: dict(measured[name])
                for name in exact_names
                if name in measured
            }
            exact_diff = [
                {
                    "name": name,
                    "previous": (guess.get(name) or {}).get("expected_label"),
                    "current": record.get("expected_label", record.get("printed_page")),
                }
                for name, record in exact.items()
                if str((guess.get(name) or {}).get("expected_label", guess.get(name, 1)))
                != str(record.get("expected_label", record.get("printed_page", 1)))
            ]
            convergence_rounds.append({
                "attempt": attempt + 1,
                "measured_names": list(measured_names),
                "exact_diff": exact_diff,
                "odd_page_padding": sorted(odd_page_padding),
                "field_target_names": list(field_target_names),
            })
            if not exact_diff:
                break
            guess = exact
        else:
            raise OfficeExportError(f"目录页码经过 {max_passes} 趟排版仍未收敛: {spec['filename']}")
        field_cache_map = reference_map if external else page_map
        field_cache_report = verify_final_field_caches(path, field_cache_map or {}) if field_target_names else {
            "fields": [], "failures": [], "passed": True, "checked_count": 0,
        }
        if not field_cache_report["passed"]:
            detail = "; ".join(
                f"{item['instruction']}: 期望 {item['expected']!r}, 实际 {item['actual']!r}"
                for item in field_cache_report["failures"]
            )
            raise OfficeExportError(f"最终 DOCX 的 PAGEREF 缓存与同次测量不一致: {detail}")
        validate_measured_delivery(path, spec, nodes, pdf_path, page_map, reference_map, parts_registry=parts_registry, page_sequences=page_sequences)
        published_maps[spec["id"]] = page_map
        if page_maps_out is not None:
            page_maps_out[spec["id"]] = {
                "page_map": {
                    name: dict(record) for name, record in page_map.items()
                },
                "page_records": {
                    name: record.to_dict()
                    for name, record in page_records_from_map(
                        page_map,
                        part_boundaries={part: BOUNDARIES[part] for part in spec["parts"]},
                        part_specs=parts_registry,
                        sequences=page_sequences,
                    ).items()
                },
                "pdf_path": str(pdf_path),
                "pagination_convergence": convergence_rounds,
                "field_cache": field_cache_report,
            }
        staged[spec["id"]] = path
    return staged
