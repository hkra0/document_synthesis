"""Document parts and relationship-aware assembly, independent of source strategy."""

import copy
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor

from .styles import apply_standard_page_setup, neutralize_document_styles, set_run_fonts, setup_footer
from .layout import apply_section_spec
from .document_parts import generated_content_bookmark_name


def _preserves_source_layout(config) -> bool:
    policy = getattr(config, "formatting", None)
    return bool(policy and (policy.mode == "preserve" or (policy.mode == "mixed" and policy.page_policy == "source")))


def _apply_config_page_setup(section, config) -> None:
    if getattr(config, "schema_version", 1) >= 3 and getattr(config, "resolved_format", None):
        apply_section_spec(section, config.resolved_format)
    else:
        apply_standard_page_setup(section, config.page_setup)


class _BoundariesDict(dict):
    def __init__(self):
        super().__init__({
            "cover": "_Synth_cover",
            "toc": "_Synth_toc",
            "body": "_Synth_body",
        })

    def __missing__(self, key):
        from .document_parts import get_part_boundary_bookmark
        return get_part_boundary_bookmark(key, is_start=True)

    def __getitem__(self, key):
        if key in self:
            return super().__getitem__(key)
        from .document_parts import get_part_boundary_bookmark
        return get_part_boundary_bookmark(key, is_start=True)


BOUNDARIES = _BoundariesDict()


def toc_nodes(nodes):
    return [node for node in nodes if node.get("include_in_toc") is not False
            and not (node.get("level", 1) > 2 and node.get("include_h3_in_toc") is False)]


def add_bookmark(paragraph, name, bookmark_id):
    start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end.set(qn("w:id"), str(bookmark_id))
    ppr = paragraph.find(qn("w:pPr"))
    paragraph.insert(1 if ppr is not None else 0, start)
    paragraph.append(end)


def mark_start(document, part):
    paragraph = next(document.element.body.iter(qn("w:p")), None)
    if paragraph is None:
        paragraph = document.add_paragraph()._p
    ids = [int(b.get(qn("w:id"))) for b in document.element.body.iter(qn("w:bookmarkStart"))]
    add_bookmark(paragraph, BOUNDARIES[part], max(ids, default=0) + 1)


def mark_end(document, part):
    from .document_parts import get_part_boundary_bookmark
    paragraph = None
    for p in document.element.body.iter(qn("w:p")):
        paragraph = p
    if paragraph is None:
        paragraph = document.add_paragraph()._p
    ids = [int(b.get(qn("w:id"))) for b in document.element.body.iter(qn("w:bookmarkStart"))]
    bm_name = get_part_boundary_bookmark(part, is_start=False)
    add_bookmark(paragraph, bm_name, max(ids, default=0) + 1)


def renumber_bookmarks(document):
    """IDs are package-local; preserve names used by TOC links, reject ambiguous names."""
    names, active = set(), {}
    next_id = 0
    for element in document.element.body.iter():
        if element.tag == qn("w:bookmarkStart"):
            name, old = element.get(qn("w:name")), element.get(qn("w:id"))
            if name in names or old in active:
                raise ValueError(f"重复或重叠的书签标识: {name}")
            names.add(name)
            active[old] = str(next_id)
            element.set(qn("w:id"), str(next_id))
            next_id += 1
        elif element.tag == qn("w:bookmarkEnd"):
            old = element.get(qn("w:id"))
            if old not in active:
                raise ValueError(f"书签结束标记没有对应起点: {old}")
            element.set(qn("w:id"), active.pop(old))
    if active:
        raise ValueError("文档包含未闭合书签。")


def _next_bookmark_id(document):
    ids = [
        int(marker.get(qn("w:id")))
        for marker in document.element.body.iter(qn("w:bookmarkStart"))
        if (marker.get(qn("w:id")) or "").lstrip("-").isdigit()
    ]
    return max(ids, default=-1) + 1


def _bookmark_run(run, name, bookmark_id):
    """Wrap one run with an invisible bookmark without changing its text."""
    parent = run._r.getparent()
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    index = parent.index(run._r)
    parent.insert(index, start)
    parent.insert(index + 2, end)


def _clone_text_run(run_element, text):
    """Clone a plain text run while retaining its run properties."""
    clone = copy.deepcopy(run_element)
    for child in list(clone):
        if child.tag == qn("w:t"):
            clone.remove(child)
    if text:
        text_node = OxmlElement("w:t")
        text_node.text = text
        if text[:1].isspace() or text[-1:].isspace():
            text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        clone.append(text_node)
    return clone


def _bookmark_text_ranges(paragraph, ranges, next_id):
    """Bookmark replacement ranges in a paragraph's normalized text run.

    ``build_cover`` collapses split-run placeholder text into the first text
    node before calling this helper, so each generated value has an exact
    character range while the original run properties remain intact.
    """
    for start_offset, end_offset, name in sorted(ranges, key=lambda item: item[0], reverse=True):
        if end_offset <= start_offset:
            continue
        text_nodes = [node for node in paragraph.iter(qn("w:t")) if node.text]
        cursor = 0
        selected = None
        for text_node in text_nodes:
            end_cursor = cursor + len(text_node.text or "")
            if start_offset >= cursor and end_offset <= end_cursor:
                selected = (text_node, cursor)
                break
            cursor = end_cursor
        if selected is None:
            raise ValueError(f"生成封面字段范围无法定位: {name}")
        text_node, cursor = selected
        run = text_node.getparent()
        parent = run.getparent()
        local_start = start_offset - cursor
        local_end = end_offset - cursor
        original = text_node.text or ""
        before, match, after = (
            original[:local_start],
            original[local_start:local_end],
            original[local_end:],
        )
        replacement = []
        if before:
            replacement.append(_clone_text_run(run, before))
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(next_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(next_id))
        replacement.extend((start, _clone_text_run(run, match), end))
        if after:
            replacement.append(_clone_text_run(run, after))
        index = parent.index(run)
        parent.remove(run)
        for offset, element in enumerate(replacement):
            parent.insert(index + offset, element)
        next_id += 1
    return next_id


def build_cover(config):
    template = config.get_template_path()
    cover_mode = getattr(config, "cover", {}).get("mode", "template" if template else "generated")
    doc = Document(str(template)) if template else Document()
    if getattr(config, "schema_version", 1) >= 3 and getattr(config, "resolved_format", None):
        for section in doc.sections:
            _apply_config_page_setup(section, config)
    title = config.cover.get("main_title") or config.project_name
    next_generated_index = {}
    if template:
        replacements = {} if cover_mode == "static_template" else {
            "{{HEADER_TITLE}}": config.cover.get("header_title") or "",
            "{{SUB_TITLE}}": config.cover.get("sub_title") or "",
            "{{MAIN_TITLE}}": title,
            "{{AUTHOR}}": config.cover.get("author") or "",
            "{{DATE}}": config.cover.get("date") or "",
        }
        if getattr(config, "schema_version", 1) >= 3 and cover_mode == "template":
            template_text = "".join(node.text or "" for node in doc.element.body.iter(qn("w:t")))
            required = {
                token: value for token, value in replacements.items()
                if value
            }
            missing = sorted(token for token, value in required.items() if token not in template_text)
            if missing:
                raise ValueError(
                    "COVER_TEMPLATE_MISSING_FIELDS: 显式动态封面模板缺少占位符: " + ", ".join(missing)
                )
        field_tokens = {
            "{{HEADER_TITLE}}": "header_title",
            "{{SUB_TITLE}}": "sub_title",
            "{{MAIN_TITLE}}": "main_title",
            "{{AUTHOR}}": "author",
            "{{DATE}}": "date",
        }
        # Work on text nodes so table and split-run placeholders are also supported.
        for paragraph in doc.element.body.iter(qn("w:p")):
            texts = list(paragraph.iter(qn("w:t")))
            old = "".join(t.text or "" for t in texts)
            new_parts = []
            replacement_ranges = []
            cursor = 0
            while cursor < len(old):
                matches = [
                    (old.find(token, cursor), token)
                    for token in replacements
                    if old.find(token, cursor) >= 0
                ]
                if not matches:
                    new_parts.append(old[cursor:])
                    break
                match_start, token = min(matches, key=lambda item: item[0])
                new_parts.append(old[cursor:match_start])
                value = str(replacements[token])
                output_start = sum(len(item) for item in new_parts)
                new_parts.append(value)
                output_end = output_start + len(value)
                if value:
                    field = field_tokens[token]
                    occurrence = next_generated_index.get(field, 0)
                    next_generated_index[field] = occurrence + 1
                    replacement_ranges.append((
                        output_start,
                        output_end,
                        generated_content_bookmark_name("cover", field, occurrence),
                    ))
                cursor = match_start + len(token)
            new = "".join(new_parts)
            if new != old and texts:
                texts[0].text = new
                for text in texts[1:]:
                    text.text = ""
                _bookmark_text_ranges(paragraph, replacement_ranges, _next_bookmark_id(doc))
    else:
        if getattr(config, "schema_version", 1) < 3:
            apply_standard_page_setup(doc.sections[0], config.page_setup)
        for field, text, size, before, after, color in (
            ("header_title", config.cover.get("header_title") or "", 16, 36, 12, None),
            ("sub_title", config.cover.get("sub_title") or "", 18, 10, 36, None),
            ("main_title", title, 26, 30, 70, RGBColor(192, 0, 0)),
            ("author", config.cover.get("author") or "", 16, 0, 12, None),
            ("date", config.cover.get("date") or "", 16, 0, 0, None),
        ):
            if not text:
                continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(before)
            p.paragraph_format.space_after = Pt(after)
            run = p.add_run(text)
            set_run_fonts(
                run,
                config.fonts["title"] if color else config.fonts.get("h2", config.fonts.get("h1", config.fonts["body"])),
                          size, bold=bool(color), color_rgb=color, font_en=config.fonts["en"])
            occurrence = next_generated_index.get(field, 0)
            next_generated_index[field] = occurrence + 1
            _bookmark_run(
                run,
                generated_content_bookmark_name("cover", field, occurrence),
                _next_bookmark_id(doc),
            )
    mark_start(doc, "cover")
    return doc


def append_toc(doc, config, spec, nodes, pages):
    if getattr(config, "schema_version", 1) >= 3 and getattr(config, "resolved_format", None):
        from .style_applier import add_styled_heading, add_styled_toc_entry
        p = add_styled_heading(
            doc,
            config.resolved_format.toc.title,
            level=1,
            resolved_format=config.resolved_format,
            bookmark_name=BOUNDARIES["toc"],
            bookmark_id=900000,
            role_name=config.resolved_format.toc.title_role,
        )
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(16)
        set_run_fonts(p.add_run("目  录"), config.fonts["title"], 22, bold=True, font_en=config.fonts["en"])
        add_bookmark(p._p, BOUNDARIES["toc"], 900000)
    from .styles import add_toc_entry_with_hyperlink
    for index, node in enumerate(toc_nodes(nodes)):
        name = node["bookmark_name"]
        page_val = pages.get(name, 1)  # Draft only; final publication requires exact comparison.
        if isinstance(page_val, dict):
            page_label = page_val.get("expected_label") or str(page_val.get("printed_page", 1))
        elif isinstance(page_val, str):
            page_label = page_val
        else:
            page_label = str(page_val)
        level = node.get("level", 1)
        if getattr(config, "schema_version", 1) >= 3 and getattr(config, "resolved_format", None):
            toc_paragraph = add_styled_toc_entry(
                doc,
                node.get("toc_title") or node["title"],
                page_label,
                level,
                config.resolved_format,
                bookmark_name=name if spec.get("toc", {}).get("links") == "internal" else None,
            )
        else:
            add_toc_entry_with_hyperlink(
                doc, node.get("toc_title") or node["title"], page_label, level,
                bookmark_name=name if spec["toc"]["links"] == "internal" else None,
                fonts=config.fonts,
            )
            toc_paragraph = doc.paragraphs[-1]
        add_bookmark(toc_paragraph._p, f"_Synth_entry_{index:05d}", 910000 + index)


def _clear_front_footers(doc):
    for section in doc.sections:
        for footer in (section.footer, section.first_page_footer, section.even_page_footer):
            footer.is_linked_to_previous = False
            for child in list(footer._element):
                footer._element.remove(child)
            footer.add_paragraph()


def normalize_body_sections(doc, config, page_sequence=None, is_seq_start=False):
    preserve_layout = _preserves_source_layout(config)
    # preserve/source 仍允许显式页码序列落到节属性；没有序列时完全不碰
    # 源页脚、页面几何或节周围段落。
    if preserve_layout and page_sequence is None:
        return
    fmt = page_sequence.format if page_sequence else "decimal"
    start_val = page_sequence.start if (page_sequence and is_seq_start) else (1 if not page_sequence else None)
    for index, section in enumerate(doc.sections):
        if not preserve_layout and getattr(config, "schema_version", 1) >= 3 and getattr(config, "resolved_format", None):
            _apply_config_page_setup(section, config)
            from .style_applier import setup_styled_footer
            setup_styled_footer(section, config.resolved_format, start_page=start_val if index == 0 else None)
        elif not preserve_layout:
            setup_footer(section, start_page=start_val if index == 0 else None, font_en=config.fonts["en"])
        # Alternate footer variants must not leak old page numbering into a
        # managed section. preserve/source 不接管故事流。
        if not preserve_layout:
            for footer in (section.first_page_footer, section.even_page_footer):
                footer.is_linked_to_previous = False
                for child in list(footer._element):
                    footer._element.remove(child)
                for child in section.footer._element:
                    footer._element.append(copy.deepcopy(child))
        secPr = section._sectPr
        pgnum = secPr.find(qn("w:pgNumType"))
        if pgnum is None:
            pgnum = parse_xml(f'<w:pgNumType {nsdecls("w")}/>')
            secPr.append(pgnum)
        pgnum.set(qn("w:fmt"), fmt)
        if start_val is not None and index == 0:
            pgnum.set(qn("w:start"), str(start_val))
        elif qn("w:start") in pgnum.attrib and (index > 0 or not is_seq_start):
            del pgnum.attrib[qn("w:start")]
    # A prepended cover must never become the body's inherited header.
    if not preserve_layout:
        for header in (doc.sections[0].header, doc.sections[0].first_page_header, doc.sections[0].even_page_header):
            header.is_linked_to_previous = False


from .package_importer import PackageImporter, validate_relationship_closure


class FrontImporter(PackageImporter):
    """Import prefix parts without rewriting any body relationships/styles/numbering."""

    def __init__(self, source, target):
        super().__init__(
            source=source,
            target=target,
            style_prefix="SynthFront_",
            part_prefix="synthFront",
            strict_rels=True,
            preserve_native_numbering=True,
        )


def assemble_document(config, spec, body_path, nodes, pages, out_path, odd_page_padding=None):
    parts = spec["parts"]
    is_legacy = set(parts).issubset({"cover", "toc", "body"}) and not getattr(config, "regions", None) and not getattr(config, "page_sequences", None)

    if is_legacy:
        front = build_cover(config) if "cover" in parts else Document()
        if "cover" in parts:
            mark_end(front, "cover")
        if "cover" not in parts:
            neutralize_document_styles(front)
            _apply_config_page_setup(front.sections[0], config)
        if "toc" in parts:
            if "cover" in parts:
                section = front.add_section(WD_SECTION_START.NEW_PAGE)
                _apply_config_page_setup(section, config)
            append_toc(front, config, spec, nodes, pages)
            mark_end(front, "toc")
        _clear_front_footers(front)
        if "body" not in parts:
            result = front
        else:
            result = Document(str(body_path))
            normalize_body_sections(result, config)
            mark_start(result, "body")
            if len(parts) > 1:
                importer = FrontImporter(front, result)
                prefix = [importer.element(e) for e in front.element.body if e.tag != qn("w:sectPr")]
                # The prefix's last sectPr terminates the prefix, not the original body.
                boundary = OxmlElement("w:p")
                ppr = OxmlElement("w:pPr")
                spacing = OxmlElement("w:spacing")
                for key, value in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
                    spacing.set(qn(f"w:{key}"), value)
                ppr.append(spacing)
                ppr.append(importer.element(front.sections[-1]._sectPr))
                boundary.append(ppr)
                prefix.append(boundary)
                result.sections[0].start_type = WD_SECTION_START.NEW_PAGE
                for element in reversed(prefix):
                    result.element.body.insert(0, element)
            # R6 内容门禁需要可定位的实际正文范围；历史 legacy 三件套也补上
            # 结束边界，避免把封面/目录的生成文本带入正文清单。
            mark_end(result, "body")
    else:
        from .document_parts import PartKind, SelectionError, SelectionValidator
        parts_registry = getattr(config, "parts", {})
        regions_registry = getattr(config, "regions", {})
        page_sequences = getattr(config, "page_sequences", {})

        source_body_doc = Document(str(body_path)) if body_path and Path(body_path).is_file() else None
        spans = {}
        if source_body_doc:
            SelectionValidator.validate_supported_objects(source_body_doc)
        if regions_registry:
            if source_body_doc is None:
                raise SelectionError("声明了 source.regions，但正文源文件不可用")
            spans = SelectionValidator.validate_regions(source_body_doc, regions_registry)

        seen_sequences = set()
        built_parts = []
        for pid in parts:
            part_def = parts_registry.get(pid)
            pkind = part_def.kind if part_def else ("cover" if pid == "cover" else ("generated_toc" if pid == "toc" else "content"))
            seq_id = part_def.page_sequence if part_def else None
            seq_def = page_sequences.get(seq_id) if seq_id else None
            is_seq_start = False
            if seq_id and seq_id not in seen_sequences:
                is_seq_start = True
                seen_sequences.add(seq_id)

            if pkind == "cover":
                pdoc = build_cover(config)
                if pid != "cover":
                    mark_start(pdoc, pid)
                mark_end(pdoc, pid)
                _clear_front_footers(pdoc)
                built_parts.append((pid, pkind, pdoc, part_def))
            elif pkind == "generated_toc":
                pdoc = Document()
                neutralize_document_styles(pdoc)
                _apply_config_page_setup(pdoc.sections[0], config)
                append_toc(pdoc, config, spec, nodes, pages)
                if pid != "toc":
                    p_first = next(pdoc.element.body.iter(qn("w:p")), None)
                    if p_first is not None:
                        add_bookmark(p_first, BOUNDARIES[pid], 900000)
                if seq_def:
                    normalize_body_sections(pdoc, config, page_sequence=seq_def, is_seq_start=is_seq_start)
                else:
                    _clear_front_footers(pdoc)
                mark_end(pdoc, pid)
                built_parts.append((pid, pkind, pdoc, part_def))
            else:  # content
                source_region = part_def.source_region if part_def else None
                if source_region:
                    if source_body_doc is None:
                        raise SelectionError(
                            f"内容部件 '{pid}' 声明了选区 '{source_region}'，但正文源文件不可用"
                        )
                    if source_region in regions_registry and getattr(
                        regions_registry[source_region], "exclude", False
                    ):
                        raise SelectionError(
                            f"内容部件 '{pid}' 引用了明确排除的选区 '{source_region}'"
                        )
                    if source_region in spans:
                        pdoc, _import_result = SelectionValidator.slice_document_by_region_with_result(
                            source_body_doc,
                            spans[source_region],
                        )
                    elif pid == "body" and source_region == "entire_document" and not regions_registry:
                        # 默认 body 的兼容语义是整篇源文档；自定义部件不得借此回退全文。
                        pdoc = Document(str(body_path))
                    else:
                        raise SelectionError(
                            f"内容部件 '{pid}' 的选区 '{source_region}' 未在 source.regions 中定义"
                        )
                elif pid == "body" and source_body_doc:
                    pdoc = Document(str(body_path))
                elif source_body_doc:
                    raise SelectionError(
                        f"内容部件 '{pid}' 未声明 source_region，拒绝隐式复制全文"
                    )
                else:
                    pdoc = Document()
                normalize_body_sections(pdoc, config, page_sequence=seq_def, is_seq_start=is_seq_start)
                mark_start(pdoc, pid)
                mark_end(pdoc, pid)
                built_parts.append((pid, pkind, pdoc, part_def))

        if not built_parts:
            result = Document()
        elif len(built_parts) == 1:
            result = built_parts[0][2]
        else:
            result = built_parts[0][2]
            for pid, pkind, next_doc, part_def in built_parts[1:]:
                importer = PackageImporter(
                    source=next_doc,
                    target=result,
                    style_prefix=f"SynthPart_{pid}_",
                    part_prefix=f"synthPart_{pid}",
                    strict_rels=True,
                    preserve_native_numbering=True,
                    sanitize_objects=not (getattr(config, "schema_version", 1) >= 3),
                )
                curr_sectPr = result.element.body.find(qn("w:sectPr"))
                boundary = OxmlElement("w:p")
                ppr = OxmlElement("w:pPr")
                spacing = OxmlElement("w:spacing")
                for key, value in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
                    spacing.set(qn(f"w:{key}"), value)
                ppr.append(spacing)
                
                sec_type = getattr(part_def, "section_type", "nextPage") if part_def else "nextPage"
                copied_sectPr = copy.deepcopy(curr_sectPr)
                type_el = copied_sectPr.find(qn("w:type"))
                if type_el is None:
                    type_el = OxmlElement("w:type")
                    copied_sectPr.append(type_el)
                # The paragraph-level sectPr closes the preceding section;
                # its type must not turn the first section itself into an
                # odd/even-page section.  The imported final sectPr below
                # carries the start mode for the new part.
                copied_sectPr.remove(type_el)
                ppr.append(copied_sectPr)
                boundary.append(ppr)

                insert_idx = result.element.body.index(curr_sectPr)
                if odd_page_padding and pid in odd_page_padding:
                    # Some Word builds do not honor oddPage on a section
                    # assembled from imported package parts.  Once a measured
                    # pass proves that the next part landed on an even page,
                    # add one explicit page break before the section so the
                    # next pass has a deterministic blank-page slot.
                    for _ in range(2):
                        padding = OxmlElement("w:p")
                        padding_run = OxmlElement("w:r")
                        padding_break = OxmlElement("w:br")
                        padding_break.set(qn("w:type"), "page")
                        padding_run.append(padding_break)
                        padding.append(padding_run)
                        result.element.body.insert(insert_idx, padding)
                        insert_idx += 1
                result.element.body.insert(insert_idx, boundary)

                next_sectPr = next_doc.element.body.find(qn("w:sectPr"))
                for el in next_doc.element.body:
                    if el.tag != qn("w:sectPr"):
                        insert_idx = result.element.body.index(curr_sectPr)
                        result.element.body.insert(insert_idx, importer.element(el))
                
                if next_sectPr is not None:
                    imported_next_sectPr = importer.element(copy.deepcopy(next_sectPr))
                    # Word resolves the start mode of the imported section
                    # from the section properties that terminate the assembled
                    # body.  Keep the boundary marker for compatibility, but
                    # also carry the declared mode onto the imported section
                    # so oddPage/evenPage survives package assembly.
                    next_type = imported_next_sectPr.find(qn("w:type"))
                    if next_type is None:
                        next_type = OxmlElement("w:type")
                        imported_next_sectPr.append(next_type)
                    next_type.set(qn("w:val"), sec_type)
                    result.element.body.replace(curr_sectPr, imported_next_sectPr)
    renumber_bookmarks(result)
    # Drawing IDs also share a document namespace across imported parts.
    for index, element in enumerate(result.element.body.iter(qn("wp:docPr")), 1):
        element.set("id", str(index))

    # rId 存在本身不足以证明关系闭包有效；在写出文档前检查所有 XML
    # 部件的引用、目标部件和图片/超链接关系类型。
    validate_relationship_closure(result)

    # 更新文档内的受控字段（题注序列 SEQ、书签引用 REF、页码引用 PAGEREF）
    from .field_updater import FieldUpdater
    FieldUpdater(pages if isinstance(pages, dict) else {}).update_document_fields(result)

    result.save(str(out_path))
