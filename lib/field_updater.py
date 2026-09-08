# -*- coding: utf-8 -*-
"""受控字段状态机与更新器。

只更新项目明确管理的 ``SEQ``、``REF`` 和 ``PAGEREF`` 字段。未知字段的
缓存值保持不变；解析错误通过结构化异常暴露，避免把 Word 的字段问题伪装
成普通文本。
"""

import re
from dataclasses import dataclass, asdict, replace
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple, Union

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .pagination_types import PageRecord, format_page_number, int_to_letters, int_to_roman


class FieldUpdateError(ValueError):
    """字段更新失败，``code`` 可供交付 QA 定位到具体字段。"""

    def __init__(self, code: str, message: str, instruction: str = ""):
        self.code = code
        self.instruction = instruction
        super().__init__(f"{code}: {message}")


@dataclass(frozen=True)
class FieldObservation:
    """一个字段实例的稳定顺序身份与最终缓存摘要。"""

    field_id: str
    story: str
    instruction: str
    command: str
    target: Optional[str]
    cached_value: str
    measurement: Optional[Dict[str, Any]] = None
    status: str = "unverified"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class FieldIndex:
    """正文、表格、页眉页脚字段的统一索引。"""

    observations: Tuple[FieldObservation, ...]

    @property
    def pageref_targets(self) -> tuple[str, ...]:
        return tuple(
            item.target for item in self.observations
            if item.command == "PAGEREF" and item.target
        )

    def with_measurements(self, page_map: Dict[str, Any]) -> "FieldIndex":
        """Attach the same typed page measurement to every PAGEREF instance."""
        enriched = []
        for item in self.observations:
            if item.command != "PAGEREF":
                enriched.append(item)
                continue
            record = page_map.get(item.target) if item.target else None
            if isinstance(record, PageRecord):
                measurement = record.to_dict()
                expected = record.expected_label
            elif isinstance(record, dict):
                measurement = dict(record)
                expected = record.get("expected_label")
                if expected is None:
                    expected = record.get("observed_label")
                if expected is None and isinstance(record.get("printed_page"), int):
                    expected = str(record["printed_page"])
            else:
                measurement = None
                expected = None
            status = "verified" if isinstance(expected, str) and item.cached_value.strip() == expected else "failed"
            enriched.append(replace(
                item,
                measurement=measurement,
                status=status,
            ))
        return FieldIndex(tuple(enriched))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": 1,
            "fields": [item.to_dict() for item in self.observations],
            "pageref_targets": list(self.pageref_targets),
        }


def _field_text(nodes: List[Any]) -> str:
    return "".join(node.text or "" for node in nodes if node.tag == qn("w:t"))


def _iter_field_paragraphs(doc: Any):
    """Yield each body/header/footer paragraph exactly once with story name."""
    yielded: Set[int] = set()
    for paragraph in doc.element.body.iter(qn("w:p")):
        yielded.add(id(paragraph))
        yield "body", paragraph
    for index, section in enumerate(getattr(doc, "sections", [])):
        containers = (
            ("header", section.header),
            ("first_header", section.first_page_header),
            ("even_header", section.even_page_header),
            ("footer", section.footer),
            ("first_footer", section.first_page_footer),
            ("even_footer", section.even_page_footer),
        )
        for story, container in containers:
            story_name = f"{story}:{index}"
            for paragraph in container._element.iter(qn("w:p")):
                if id(paragraph) not in yielded:
                    yielded.add(id(paragraph))
                    yield story_name, paragraph


def _simple_field_observation(field: Any, field_id: str, story: str) -> FieldObservation:
    instruction = FieldUpdater._normalise_instruction(field.get(qn("w:instr"), ""))
    tokens = instruction.split()
    command = tokens[0].upper() if tokens else ""
    target = tokens[1] if len(tokens) > 1 and command in {"REF", "PAGEREF"} else None
    return FieldObservation(
        field_id=field_id,
        story=story,
        instruction=instruction,
        command=command,
        target=target,
        cached_value=_field_text(list(field.iter(qn("w:t")))),
    )


def build_field_index(doc_or_path: Union[str, Any]) -> FieldIndex:
    """统一索引简单/复杂字段，复杂字段按 begin/separate/end 状态解析。"""
    doc = Document(str(doc_or_path)) if isinstance(doc_or_path, (str, Path)) else doc_or_path
    observations: List[FieldObservation] = []
    counter = 0
    for story, paragraph in _iter_field_paragraphs(doc):
        for field in paragraph.findall(qn("w:fldSimple")):
            counter += 1
            observations.append(_simple_field_observation(field, f"field-{counter:06d}", story))

        stack: List[Dict[str, Any]] = []
        for child in paragraph:
            if child.tag != qn("w:r"):
                continue
            fld_chars = child.findall(qn("w:fldChar"))
            instr_nodes = child.findall(qn("w:instrText"))
            field_types = [item.get(qn("w:fldCharType"), "") for item in fld_chars]
            for field_type in field_types:
                if field_type == "begin":
                    stack.append({"instruction": [], "result": [], "separate": False})
                elif field_type == "separate" and stack:
                    stack[-1]["separate"] = True
                elif field_type == "end" and stack:
                    frame = stack.pop()
                    instruction = FieldUpdater._normalise_instruction("".join(frame["instruction"]))
                    tokens = instruction.split()
                    command = tokens[0].upper() if tokens else ""
                    target = tokens[1] if len(tokens) > 1 and command in {"REF", "PAGEREF"} else None
                    counter += 1
                    observations.append(FieldObservation(
                        field_id=f"field-{counter:06d}",
                        story=story,
                        instruction=instruction,
                        command=command,
                        target=target,
                        cached_value=_field_text(frame["result"]),
                    ))
            if stack and instr_nodes and not stack[-1]["separate"]:
                stack[-1]["instruction"].extend(node.text or "" for node in instr_nodes)
            if stack and not instr_nodes and not any(item in {"begin", "separate"} for item in field_types):
                for frame in stack:
                    if frame["separate"]:
                        frame["result"].extend(child.iter(qn("w:t")))
    return FieldIndex(tuple(observations))


def required_pageref_targets(doc_or_path: Union[str, Any]) -> List[str]:
    """返回字段分页依赖的有序并集，不因目录是否存在而为空。"""
    return list(dict.fromkeys(build_field_index(doc_or_path).pageref_targets))


def verify_final_field_caches(
    doc_or_path: Union[str, Any],
    page_map: Dict[str, Any],
) -> Dict[str, Any]:
    """核对最终磁盘 DOCX 中每个受控 PAGEREF 的缓存与同次测量记录。"""
    observations = build_field_index(doc_or_path).with_measurements(page_map).observations
    results = []
    failures = []
    for item in observations:
        if item.command != "PAGEREF":
            continue
        record = page_map.get(item.target) if item.target else None
        expected = None
        if isinstance(record, dict):
            if "expected_label" in record:
                expected = record.get("expected_label")
            else:
                expected = record.get("observed_label")
            if expected is None and isinstance(record.get("printed_page"), int):
                expected = str(record["printed_page"])
        actual = item.cached_value.strip()
        ok = isinstance(expected, str) and actual == expected
        result = {
            "field_id": item.field_id,
            "story": item.story,
            "instruction": item.instruction,
            "target": item.target,
            "expected": expected,
            "actual": actual,
            "measurement": item.measurement,
            "status": "verified" if ok else "failed",
        }
        results.append(result)
        if not ok:
            failures.append(result)
    return {
        "fields": results,
        "failures": failures,
        "passed": not failures,
        "checked_count": len(results),
    }


class _ComplexFieldState:
    """复杂字段栈状态帧。"""

    def __init__(self, begin_r: Any):
        self.begin_r = begin_r
        self.instr_parts: List[str] = []
        self.has_separate = False
        self.result_runs: List[Any] = []

    @property
    def instruction(self) -> str:
        return "".join(self.instr_parts).strip()


class FieldUpdater:
    """在不破坏 OMML 和未知字段的前提下更新受控字段。"""

    def __init__(self, page_map: Optional[Dict[str, Any]] = None):
        self.page_map = page_map or {}
        self.counters: Dict[str, int] = {}
        self.bookmark_text_cache: Dict[str, str] = {}

    @staticmethod
    def _normalise_instruction(instr: str) -> str:
        if not isinstance(instr, str):
            raise FieldUpdateError("INVALID_INSTRUCTION", "字段指令必须是字符串。")
        # XML/fixture 生成器有时会把 Word 的单反斜线写成两个；两者应有相同语义。
        return re.sub(r"\\{2,}", r"\\", instr).strip()

    @classmethod
    def _command(cls, instr: str) -> str:
        normalised = cls._normalise_instruction(instr)
        return normalised.split(maxsplit=1)[0].upper() if normalised else ""

    def _evaluate_instruction(self, instr: str) -> Optional[str]:
        """评估一个字段；未知字段返回 ``None`` 以保留原缓存。"""
        normalised = self._normalise_instruction(instr)
        tokens = normalised.split()
        if not tokens:
            return None

        cmd = tokens[0].upper()
        if cmd == "SEQ":
            return self._evaluate_seq(tokens, normalised)
        if cmd == "PAGEREF":
            return self._evaluate_pageref(tokens, normalised)
        if cmd == "REF":
            if len(tokens) < 2:
                return None
            return self.bookmark_text_cache.get(tokens[1])
        return None

    def _evaluate_seq(self, tokens: List[str], instruction: str) -> str:
        if len(tokens) < 2:
            return ""
        category = tokens[1].lower()
        current = self.counters.get(category, 0)

        reset = re.search(r"(?:^|\s)\\r\s+(\d+)(?:\s|$)", instruction, re.IGNORECASE)
        continue_current = bool(re.search(r"(?:^|\s)\\c(?:\s|$)", instruction, re.IGNORECASE))
        if reset:
            value = int(reset.group(1))
            if value < 1:
                raise FieldUpdateError("INVALID_SEQ_RESET", "SEQ \\r 的编号必须大于等于 1。", instruction)
            self.counters[category] = value
        elif continue_current:
            value = current if current > 0 else 1
            if current == 0:
                self.counters[category] = value
        else:
            value = current + 1
            self.counters[category] = value

        switch = re.search(r"(?:^|\s)\\\*\s+([A-Za-z]+)(?:\s|$)", instruction)
        if not switch:
            return str(value)
        picture = switch.group(1)
        upper_picture = picture.upper()
        if upper_picture == "ROMAN":
            return int_to_roman(value).lower() if picture.islower() else int_to_roman(value)
        if upper_picture == "ALPHABETIC":
            return int_to_letters(value, lowercase=picture.islower())
        if upper_picture == "ARABIC":
            return str(value)
        return str(value)

    def _evaluate_pageref(self, tokens: List[str], instruction: str) -> Optional[str]:
        if len(tokens) < 2:
            return None
        target = tokens[1]
        if target not in self.page_map:
            return None
        record = self.page_map[target]
        if isinstance(record, PageRecord):
            return record.expected_label
        if not isinstance(record, dict):
            raise FieldUpdateError(
                "INVALID_PAGE_INDEX",
                f"PAGEREF {target} 的页码索引必须是 PageRecord 或对象。",
                instruction,
            )
        expected = record.get("expected_label")
        if expected is not None:
            if not isinstance(expected, str):
                raise FieldUpdateError("INVALID_PAGE_INDEX", "expected_label 必须是字符串。", instruction)
            return expected
        value = record.get("number_value", record.get("printed_page"))
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise FieldUpdateError(
                "INVALID_PAGE_INDEX",
                f"PAGEREF {target} 缺少有效的 number_value/printed_page。",
                instruction,
            )
        fmt = record.get("format", "decimal")
        if not isinstance(fmt, str):
            raise FieldUpdateError("INVALID_PAGE_INDEX", "页码格式必须是字符串。", instruction)
        return format_page_number(value, fmt)

    @staticmethod
    def _update_runs_with_text(runs: List[Any], new_text: str, parent_p: Any = None) -> None:
        """把字段结果写入第一个文字 run，并清空其余缓存文字。"""
        text_nodes = []
        for run in runs:
            text_nodes.extend(run.iter(qn("w:t")))
        if text_nodes:
            text_nodes[0].text = new_text
            for node in text_nodes[1:]:
                node.text = ""
            return
        if runs:
            text = OxmlElement("w:t")
            text.text = new_text
            runs[0].append(text)
            return
        if parent_p is not None:
            run = OxmlElement("w:r")
            text = OxmlElement("w:t")
            text.text = new_text
            run.append(text)
            parent_p.append(run)

    def update_document_fields(self, doc: Any) -> None:
        """按依赖顺序更新正文、表格和页眉页脚中的受控字段。

        SEQ 先更新并建立书签文本，再更新 REF/PAGEREF。这样 REF 可以引用
        同一文档中后面出现的题注，也不会被段落外文本污染。
        """
        self.counters.clear()
        self.bookmark_text_cache.clear()
        self._validate_field_dependencies(doc)
        self._process_fields(doc, {"SEQ"})
        self._collect_bookmark_texts(doc)
        # A REF can point to a later bookmark, and that bookmark can itself
        # contain a REF.  Re-evaluate until the observable cache is stable;
        # cycles have already been rejected by the dependency check above.
        ref_count = sum(
            1
            for paragraph in self._iter_paragraphs(doc)
            for field in paragraph.findall(qn("w:fldSimple"))
            if self._command(field.get(qn("w:instr"), "")) in {"REF", "PAGEREF"}
        )
        # Complex fields keep their instruction in w:instrText rather than a
        # single attribute.  Count those too so a forward chain of complex
        # REF fields gets enough bounded convergence passes.
        ref_count += sum(
            1
            for paragraph in self._iter_paragraphs(doc)
            for instr in paragraph.iter(qn("w:instrText"))
            if self._command(instr.text or "") in {"REF", "PAGEREF"}
        )
        for _ in range(max(1, ref_count + 1)):
            before = {
                key: "".join(node.text or "" for node in self._bookmark_nodes(doc, key))
                for key in self.bookmark_text_cache
            }
            self._process_fields(doc, {"REF", "PAGEREF"})
            self._collect_bookmark_texts(doc)
            after = {
                key: "".join(node.text or "" for node in self._bookmark_nodes(doc, key))
                for key in self.bookmark_text_cache
            }
            if before == after:
                break

    def _iter_paragraphs(self, doc: Any):
        # body.iter 也涵盖表格中的段落；额外扫描 headers/footers，因其同样
        # 可能包含 PAGEREF，但不依赖 python-docx 的代理对象构造。
        yielded: Set[int] = set()
        for root in [doc.element.body]:
            for paragraph in root.iter(qn("w:p")):
                yielded.add(id(paragraph))
                yield paragraph
        for section in getattr(doc, "sections", []):
            for container in (
                section.header,
                section.first_page_header,
                section.even_page_header,
                section.footer,
                section.first_page_footer,
                section.even_page_footer,
            ):
                for paragraph in container._element.iter(qn("w:p")):
                    if id(paragraph) not in yielded:
                        yielded.add(id(paragraph))
                        yield paragraph

    def _process_fields(self, doc: Any, commands: Set[str]) -> None:
        for paragraph in self._iter_paragraphs(doc):
            for field in paragraph.findall(qn("w:fldSimple")):
                instruction = field.get(qn("w:instr"), "")
                if self._command(instruction) not in commands:
                    continue
                value = self._evaluate_instruction(instruction)
                if value is not None:
                    self._update_runs_with_text([field], value, paragraph)

            stack: List[_ComplexFieldState] = []
            for child in paragraph:
                if child.tag != qn("w:r"):
                    continue
                fld_chars = child.findall(qn("w:fldChar"))
                instr_nodes = child.findall(qn("w:instrText"))

                for fld_char in fld_chars:
                    field_type = fld_char.get(qn("w:fldCharType"), "")
                    if field_type == "begin":
                        stack.append(_ComplexFieldState(child))
                    elif field_type == "separate":
                        if not stack:
                            raise FieldUpdateError("UNBALANCED_FIELD", "字段出现 separate 但没有 begin。")
                        stack[-1].has_separate = True
                    elif field_type == "end":
                        if not stack:
                            raise FieldUpdateError("UNBALANCED_FIELD", "字段出现 end 但没有 begin。")
                        field = stack.pop()
                        if self._command(field.instruction) in commands:
                            value = self._evaluate_instruction(field.instruction)
                            if value is not None:
                                self._update_runs_with_text(field.result_runs, value, paragraph)

                if instr_nodes and stack and not stack[-1].has_separate:
                    stack[-1].instr_parts.extend(node.text or "" for node in instr_nodes)

                # 一个 run 可能同时承载字段控制字符和文字。只把分隔符之后
                # 的结果 run 记录给所有活跃字段，支持跨 run/跨段的常见形态。
                if stack and any(field.has_separate for field in stack):
                    if not instr_nodes and not any(
                        fc.get(qn("w:fldCharType")) in {"begin", "separate"} for fc in fld_chars
                    ):
                        for field in stack:
                            if field.has_separate and child not in field.result_runs:
                                field.result_runs.append(child)

            if stack:
                raise FieldUpdateError("UNCLOSED_FIELD", "段落结束时仍有未闭合的复杂字段。")

    def _bookmark_nodes(self, doc: Any, name: str) -> List[Any]:
        """Return the text nodes in one exact bookmark range for convergence checks."""
        nodes: List[Any] = []
        active: Dict[str, bool] = {}
        for paragraph in self._iter_paragraphs(doc):
            for element in paragraph.iter():
                if element.tag == qn("w:bookmarkStart"):
                    if element.get(qn("w:name")) == name:
                        active[str(element.get(qn("w:id")))] = True
                elif element.tag == qn("w:t") and active:
                    nodes.append(element)
                elif element.tag == qn("w:bookmarkEnd"):
                    active.pop(str(element.get(qn("w:id"))), None)
        return nodes

    def _validate_field_dependencies(self, doc: Any) -> None:
        """Build the bookmark-level REF dependency graph and reject cycles."""
        active: Dict[str, Dict[str, Any]] = {}
        stack: List[List[str]] = []
        graph: Dict[str, Set[str]] = {}
        for paragraph in self._iter_paragraphs(doc):
            for element in paragraph.iter():
                if element.tag == qn("w:bookmarkStart"):
                    name = element.get(qn("w:name"))
                    bookmark_id = element.get(qn("w:id"))
                    if name and bookmark_id is not None:
                        active[str(bookmark_id)] = {"name": name}
                        graph.setdefault(name, set())
                elif element.tag == qn("w:fldSimple"):
                    instr = element.get(qn("w:instr"), "")
                    self._record_dependency(active, graph, instr)
                elif element.tag == qn("w:fldChar"):
                    field_type = element.get(qn("w:fldCharType"), "")
                    if field_type == "begin":
                        stack.append([])
                    elif field_type == "separate":
                        if stack:
                            stack[-1].append("")
                    elif field_type == "end":
                        if stack:
                            instr = "".join(stack.pop())
                            self._record_dependency(active, graph, instr)
                elif element.tag == qn("w:instrText") and stack:
                    stack[-1].append(element.text or "")
                elif element.tag == qn("w:bookmarkEnd"):
                    active.pop(str(element.get(qn("w:id"))), None)

        visiting: Set[str] = set()
        visited: Set[str] = set()

        def visit(node: str, trail: List[str]) -> None:
            if node in visiting:
                cycle = " -> ".join(trail + [node])
                raise FieldUpdateError("FIELD_DEPENDENCY_CYCLE", f"检测到 REF 字段循环依赖: {cycle}")
            if node in visited:
                return
            visiting.add(node)
            for dependency in sorted(graph.get(node, ())):
                visit(dependency, trail + [node])
            visiting.remove(node)
            visited.add(node)

        for node in sorted(graph):
            visit(node, [])

    def _record_dependency(self, active: Dict[str, Dict[str, Any]], graph: Dict[str, Set[str]], instr: str) -> None:
        normalised = self._normalise_instruction(instr)
        tokens = normalised.split()
        if len(tokens) < 2 or tokens[0].upper() != "REF":
            return
        target = tokens[1]
        for source in active.values():
            graph.setdefault(source["name"], set()).add(target)

    def _collect_bookmark_texts(self, doc: Any) -> None:
        """只收集 bookmarkStart 到对应 bookmarkEnd 之间的文字。

        状态跨段落保留，因此跨段书签不会把段落前后无关正文一并吞入 REF。
        """
        self.bookmark_text_cache.clear()
        active: Dict[str, Dict[str, Any]] = {}
        for paragraph in self._iter_paragraphs(doc):
            for element in paragraph.iter():
                if element.tag == qn("w:bookmarkStart"):
                    name = element.get(qn("w:name"))
                    bookmark_id = element.get(qn("w:id"))
                    if name and bookmark_id is not None:
                        active[str(bookmark_id)] = {"name": name, "text": []}
                if element.tag == qn("w:t"):
                    value = element.text or ""
                    for entry in active.values():
                        entry["text"].append(value)
                if element.tag == qn("w:bookmarkEnd"):
                    bookmark_id = element.get(qn("w:id"))
                    entry = active.pop(str(bookmark_id), None)
                    if entry is not None:
                        self.bookmark_text_cache[entry["name"]] = "".join(entry["text"]).strip()
        # Malformed but recoverable input: retain an explicit partial range rather
        # than silently falling back to the whole paragraph.
        for entry in active.values():
            self.bookmark_text_cache[entry["name"]] = "".join(entry["text"]).strip()
