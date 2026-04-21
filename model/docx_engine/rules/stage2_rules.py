import re

from docx.enum.text import WD_ALIGN_PARAGRAPH

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source
from model.format_checker import FormatChecker


def _cm_value(length_obj) -> float | None:
    if length_obj is None:
        return None
    try:
        return float(length_obj.cm)
    except Exception:  # noqa: BLE001
        return None


def _approx_equal(value: float | None, target: float, tolerance: float = 0.15) -> bool:
    if value is None:
        return False
    return abs(value - target) <= tolerance


def _paragraph_text(paragraph) -> str:
    return str(getattr(paragraph, "text", "") or "").strip()


def _normalize_space(text: str) -> str:
    return re.sub(r"[\s\u3000]+", " ", text or "").strip()


def _normalize_heading_compare_text(text: str) -> str:
    normalized = _normalize_space(text)
    normalized = normalized.replace("．", ".")
    normalized = re.sub(r"[·•●▪■◆◇○◎]", "", normalized)
    return normalized.strip()


def _strip_catalogue_trailing_page(text: str) -> str:
    normalized = _normalize_heading_compare_text(text)
    normalized = re.sub(r"[.．·•…\-_ ]{2,}\s*[IVXLCDMivxlcdm\d]+\s*$", "", normalized)
    normalized = re.sub(r"\s+[IVXLCDMivxlcdm\d]+\s*$", "", normalized)
    return normalized.strip()


def _iter_nonempty_paragraphs(paragraphs):
    for paragraph in paragraphs or []:
        if paragraph is None:
            continue
        if _paragraph_text(paragraph):
            yield paragraph


def _nonempty_runs(paragraph):
    for run in getattr(paragraph, "runs", []) or []:
        if str(getattr(run, "text", "") or "").strip():
            yield run


def _append_issue(collector: list[Issue], *, rule_id: str, title: str, message: str, problem: str, section: str, content: str, severity: Severity = Severity.WARNING, metadata: dict | None = None) -> None:
    merged_metadata = {
        "section": section,
        "content": content,
        "problem": problem,
    }
    if metadata:
        merged_metadata.update(metadata)
    collector.append(
        Issue(
            rule_id=rule_id,
            title=title,
            message=message,
            severity=severity,
            source=Source.DOCX,
            fixable=False,
            metadata=merged_metadata,
        )
    )


def _extract_catalogue_entries(catalogue_content) -> list[str]:
    entries: list[str] = []
    for paragraph in catalogue_content or []:
        text = _strip_catalogue_trailing_page(_paragraph_text(paragraph))
        if not text:
            continue
        if text in {"目录", "目 录"}:
            continue
        if re.match(r"^(图|表)\s*[A-Z]?\d+", text):
            continue
        if not (
            re.match(r"^第[一二三四五六七八九十百千0-9]+章", text)
            or re.match(r"^\d+\.\d+\.\d+", text)
            or re.match(r"^\d+\.\d+", text)
            or re.match(r"^[一二三四五六七八九十]+、", text)
        ):
            continue
        entries.append(text)
    return entries


def _catalogue_number_prefix(text: str) -> str | None:
    raw = str(text or "")
    if not raw.strip():
        return None

    patterns = [
        r"^\s*(\d+(?:\.\d+){0,2})\s+",
        r"^\s*(第\s*\d+\s*章)\s*",
    ]
    for pattern in patterns:
        match = re.match(pattern, raw)
        if match:
            prefix = match.group(1)
            return prefix if re.search(r"\d", prefix) else None
    return None


def _runs_in_prefix(paragraph, prefix_text: str):
    if not prefix_text:
        return []

    remaining = len(prefix_text)
    selected = []
    for run in getattr(paragraph, "runs", []) or []:
        run_text = str(getattr(run, "text", "") or "")
        if not run_text:
            continue
        if remaining <= 0:
            break

        take_len = min(len(run_text), remaining)
        if run_text[:take_len]:
            selected.append((run, run_text[:take_len]))
        remaining -= take_len

    return selected


def _collect_body_headings(docx_sections: dict) -> list[str]:
    headings = docx_sections.get("headings", {}) if isinstance(docx_sections, dict) else {}
    if not isinstance(headings, dict):
        return []
    collected: list[str] = []
    for level in ("chapter", "level1", "level2"):
        for paragraph in headings.get(level, []) or []:
            text = _normalize_heading_compare_text(_paragraph_text(paragraph))
            if text:
                collected.append(text)
    return collected


def _extract_citation_numbers(text: str) -> list[int]:
    numbers: list[int] = []
    for match in re.finditer(r"\[\s*([0-9,\-，、\s]+)\s*\]", text or ""):
        body = match.group(1)
        parts = [item.strip() for item in re.split(r"[，,、]", body) if item.strip()]
        for part in parts:
            if "-" in part:
                bounds = [x.strip() for x in part.split("-", 1)]
                if len(bounds) == 2 and bounds[0].isdigit() and bounds[1].isdigit():
                    start = int(bounds[0])
                    end = int(bounds[1])
                    if start <= end:
                        numbers.extend(range(start, end + 1))
                continue
            if part.isdigit():
                numbers.append(int(part))
    return numbers


def _extract_reference_entry_numbers(reference_entries: list[str]) -> tuple[dict[int, str], list[str]]:
    numbered: dict[int, str] = {}
    unnumbered: list[str] = []
    for entry in reference_entries:
        match = re.match(r"^[\[［]\s*(\d+)\s*[\]］]\s*(.*)$", entry)
        if not match:
            unnumbered.append(entry)
            continue
        numbered[int(match.group(1))] = entry
    return numbered, unnumbered


def _create_format_checker(ctx: RuleContext) -> FormatChecker:
    checker = FormatChecker()
    checker.format_rules["catalogue"] = {
        "title": {
            "font": "黑体",
            "size": "小二 (18pt)",
            "alignment": "居中",
            "line_spacing": "固定值20pt",
        },
        "content": {
            "font": "宋体",
            "size": "小四 (12pt)",
            "alignment": "左对齐",
            "line_spacing": "固定值20pt",
        },
    }

    user_formats = ctx.extras.get("user_formats") or {}
    if isinstance(user_formats, dict):
        try:
            checker.update_formats(user_formats)
        except Exception:  # noqa: BLE001
            pass

        catalogue_settings = user_formats.get("目录")
        if isinstance(catalogue_settings, dict):
            title_rules = checker.format_rules["catalogue"]["title"]
            content_rules = checker.format_rules["catalogue"]["content"]
            title_rules["font"] = catalogue_settings.get("title_font", title_rules["font"])
            title_rules["size"] = catalogue_settings.get("title_size", title_rules["size"])
            title_rules["alignment"] = catalogue_settings.get("title_align", title_rules["alignment"])
            title_rules["line_spacing"] = catalogue_settings.get("title_line_spacing", title_rules["line_spacing"])
            content_rules["font"] = catalogue_settings.get("content_font", content_rules["font"])
            content_rules["size"] = catalogue_settings.get("content_size", content_rules["size"])
            content_rules["alignment"] = catalogue_settings.get("content_align", content_rules["alignment"])
            content_rules["line_spacing"] = catalogue_settings.get("content_line_spacing", content_rules["line_spacing"])

    return checker


def _iter_format_targets(ctx: RuleContext, format_rules: dict):
    sections = ctx.docx_sections or {}
    if not isinstance(sections, dict):
        return

    mapping = [
        ("cover", "school", "school"),
        ("cover", "title", "title"),
        ("cover", "personal_information", "personal_information"),
        ("statement", "title", "title"),
        ("statement", "content", "content"),
        ("abstract_keyword", "chinese_title", "chinese_title"),
        ("abstract_keyword", "chinese_content", "chinese_content"),
        ("abstract_keyword", "chinese_keyword_title", "chinese_keyword_title"),
        ("abstract_keyword", "chinese_keyword", "chinese_keyword"),
        ("abstract_keyword", "english_title", "english_title"),
        ("abstract_keyword", "english_content", "english_content"),
        ("abstract_keyword", "english_keyword_title", "english_keyword_title"),
        ("abstract_keyword", "english_keyword", "english_keyword"),
        ("catalogue", "title", "title"),
        ("catalogue", "content", "content"),
        ("headings", "chapter", "chapter"),
        ("headings", "level1", "level1"),
        ("headings", "level2", "level2"),
        ("headings", "level3", "level3"),
        ("main_text", None, None),
        ("figures_or_tables_title", None, None),
        ("references", "title", "title"),
        ("references", "content", "content"),
        ("acknowledgments", "title", "title"),
        ("acknowledgments", "content", "content"),
    ]

    for section_key, sub_key, rule_key in mapping:
        content = sections.get(section_key)
        if content is None:
            continue

        if rule_key is None:
            expected = format_rules.get(section_key)
            paragraphs = content if isinstance(content, list) else [content]
        else:
            section_rules = format_rules.get(section_key, {})
            expected = section_rules.get(rule_key)
            if isinstance(content, dict):
                value = content.get(sub_key)
            else:
                value = None
            if isinstance(value, list):
                paragraphs = value
            elif value is None:
                paragraphs = []
            else:
                paragraphs = [value]

        if not expected:
            continue

        for index, paragraph in enumerate(paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue
            yield {
                "section_label": _field_display_label(section_key, sub_key),
                "section_key": section_key,
                "sub_key": sub_key,
                "paragraph_index": index,
                "paragraph": paragraph,
                "text": text,
                "expected": expected,
            }


def _effective_alignment(paragraph):
    actual_alignment = getattr(paragraph, "alignment", None)
    if actual_alignment is not None:
        return actual_alignment
    style = getattr(paragraph, "style", None)
    style_format = getattr(style, "paragraph_format", None) if style else None
    return getattr(style_format, "alignment", None) if style_format else None


def _effective_spacing_value(paragraph, attr_name: str):
    paragraph_format = getattr(paragraph, "paragraph_format", None)
    value = getattr(paragraph_format, attr_name, None) if paragraph_format else None
    if value is not None:
        return value
    style = getattr(paragraph, "style", None)
    style_format = getattr(style, "paragraph_format", None) if style else None
    return getattr(style_format, attr_name, None) if style_format else None


def _space_matches_half_line(value, expected_size_pt: float | None) -> bool:
    if value is None or expected_size_pt is None:
        return True
    actual_pt = getattr(value, "pt", None)
    if actual_pt is None:
        return True
    target = expected_size_pt * 0.5
    return abs(actual_pt - target) <= 2.0


def _field_display_label(section_key: str, sub_key: str | None) -> str:
    mapping = {
        ("cover", "school"): "封面学校名称",
        ("cover", "title"): "封面题目",
        ("cover", "personal_information"): "封面个人信息",
        ("statement", "title"): "声明标题",
        ("statement", "content"): "声明内容",
        ("abstract_keyword", "chinese_title"): "中文摘要标题",
        ("abstract_keyword", "chinese_content"): "中文摘要内容",
        ("abstract_keyword", "chinese_keyword_title"): "中文关键词标题",
        ("abstract_keyword", "chinese_keyword"): "中文关键词内容",
        ("abstract_keyword", "english_title"): "英文摘要标题",
        ("abstract_keyword", "english_content"): "英文摘要内容",
        ("abstract_keyword", "english_keyword_title"): "英文关键词标题",
        ("abstract_keyword", "english_keyword"): "英文关键词内容",
        ("catalogue", "title"): "目录标题",
        ("catalogue", "content"): "目录内容",
        ("headings", "chapter"): "章节标题",
        ("headings", "level1"): "一级标题",
        ("headings", "level2"): "二级标题",
        ("headings", "level3"): "三级标题",
        ("main_text", None): "正文",
        ("figures_or_tables_title", None): "图表题",
        ("references", "title"): "参考文献标题",
        ("references", "content"): "参考文献内容",
        ("acknowledgments", "title"): "致谢标题",
        ("acknowledgments", "content"): "致谢内容",
    }
    return mapping.get((section_key, sub_key), section_key)


def _extract_keyword_parts(paragraph) -> dict | None:
    text = _paragraph_text(paragraph)
    if not text or "：" not in text and ":" not in text:
        return None

    match = re.match(r"^\s*(.+?[:：])\s*(.*)$", text)
    if not match:
        return None

    title_text = match.group(1).strip()
    content_text = match.group(2).strip()
    if not title_text:
        return None

    runs = list(getattr(paragraph, "runs", []) or [])
    if not runs:
        return {
            "title_text": title_text,
            "content_text": content_text,
            "title_runs": [],
            "content_runs": [],
        }

    title_end = len(title_text)
    consumed = 0
    title_runs = []
    content_runs = []
    for run in runs:
        run_text = str(getattr(run, "text", "") or "")
        if not run_text:
            continue
        run_start = consumed
        run_end = consumed + len(run_text)
        if run_start < title_end:
            title_runs.append(run)
        if run_end > title_end:
            content_runs.append(run)
        consumed = run_end

    return {
        "title_text": title_text,
        "content_text": content_text,
        "title_runs": title_runs,
        "content_runs": content_runs,
    }


def _check_runs_font(checker: FormatChecker, runs, expected_font: str | None) -> bool:
    if not expected_font:
        return True
    if not runs:
        return True
    for run in runs:
        if not str(getattr(run, "text", "") or "").strip():
            continue
        candidate_fonts = checker._get_run_font_candidates(run)
        if not candidate_fonts:
            continue
        if not any(checker._is_font_equivalent(expected_font, item) for item in candidate_fonts):
            return False
    return True


def _check_runs_size(checker: FormatChecker, runs, expected_size: str | None) -> bool:
    target_size = checker._get_font_size_pt(expected_size) if expected_size else None
    if target_size is None or not runs:
        return True
    for run in runs:
        if not str(getattr(run, "text", "") or "").strip():
            continue
        actual_size = run.font.size.pt if run.font.size else None
        if actual_size is None:
            continue
        if abs(actual_size - target_size) > 0.5:
            return False
    return True


def _chapter_number_from_heading(text: str) -> str | None:
    stripped = str(text or "").strip()
    match = re.match(r"^第([一二三四五六七八九十百千0-9]+)章", stripped)
    if not match:
        return None
    raw = match.group(1)
    if raw.isdigit():
        return raw

    digits = {
        "零": 0,
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    if raw == "十":
        return "10"
    if "十" in raw:
        parts = raw.split("十")
        tens = digits.get(parts[0], 1) if parts[0] else 1
        ones = digits.get(parts[1], 0) if len(parts) > 1 and parts[1] else 0
        return str(tens * 10 + ones)
    if raw in digits:
        return str(digits[raw])
    return None


def _appendix_letter_from_heading(text: str) -> str | None:
    match = re.match(r"^附录\s*([A-Z])(?:\s+.*)?$", text.strip())
    return match.group(1) if match else None


class SectionOrderRule(BaseRule):
    rule_id = "docx.stage2.section_order"
    display_name = "Section order"
    spec_ref = "撰写规范 资料整理要求（1）装订顺序"
    engine = "docx"

    _expected_order = [
        "cover",
        "statement1",
        "chinese_abstract",
        "english_abstract",
        "catalogue",
        "main_text",
        "references",
        "acknowledgments",
    ]
    _part_name = {
        "cover": "封面",
        "statement1": "原创声明",
        "chinese_abstract": "中文摘要",
        "english_abstract": "英文摘要",
        "catalogue": "目录",
        "main_text": "正文",
        "references": "参考文献",
        "acknowledgments": "致谢",
    }

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        parts_order = list(ctx.extras.get("docx_parts_order") or [])
        if not parts_order:
            return []

        issues: list[Issue] = []
        comparable_parts = [part for part in parts_order if part in self._expected_order]
        expected_positions = {part: idx for idx, part in enumerate(self._expected_order)}

        for idx in range(len(comparable_parts) - 1):
            current = comparable_parts[idx]
            nxt = comparable_parts[idx + 1]
            if expected_positions[current] > expected_positions[nxt]:
                content = " -> ".join(self._part_name.get(part, part) for part in comparable_parts)
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.sequence.{idx + 1}",
                    title="装订顺序",
                    message=content,
                    problem=f"{self._part_name.get(current, current)} 与 {self._part_name.get(nxt, nxt)} 的先后顺序可能有误",
                    section="整体结构",
                    content=content,
                    metadata={"parts_order": comparable_parts},
                )
                break

        return issues


class CatalogueHeadingConsistencyRule(BaseRule):
    rule_id = "docx.stage2.catalogue_heading_consistency"
    display_name = "Catalogue heading consistency"
    spec_ref = "撰写规范（4）目录、（8）目录"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        catalogue = sections.get("catalogue", {}) if isinstance(sections, dict) else {}
        catalogue_content = catalogue.get("content") if isinstance(catalogue, dict) else None
        catalogue_entries = _extract_catalogue_entries(catalogue_content)
        if not catalogue_entries:
            return []

        body_headings = _collect_body_headings(sections)
        if not body_headings:
            return []

        issues: list[Issue] = []
        body_set = set(body_headings)
        catalogue_set = set(catalogue_entries)

        missing_in_body = [entry for entry in catalogue_entries if entry not in body_set]
        missing_in_catalogue = [entry for entry in body_headings if entry not in catalogue_set]

        for index, entry in enumerate(missing_in_body, start=1):
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.catalogue_only.{index}",
                title="目录",
                message=entry,
                problem="目录标题在正文标题中未找到对应项",
                section="目录",
                content=entry,
            )

        for index, entry in enumerate(missing_in_catalogue[:10], start=1):
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.body_only.{index}",
                title="目录",
                message=entry,
                problem="正文标题可能未出现在目录中",
                section="目录",
                content=entry,
            )

        if len(body_headings) >= 3:
            level3_exists = any(re.match(r"^\d+\.\d+\.\d+", item) for item in body_headings)
            level3_in_catalogue = any(re.match(r"^\d+\.\d+\.\d+", item) for item in catalogue_entries)
            if level3_exists and not level3_in_catalogue:
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.level3_missing",
                    title="目录",
                    message="目录条目层级检查",
                    problem="正文存在三级标题，但目录中可能未体现三级标题",
                    section="目录",
                    content="目录条目层级检查",
                )

        return issues


class CatalogueNumberFontRule(BaseRule):
    rule_id = "docx.stage2.catalogue_number_font"
    display_name = "Catalogue numbering font"
    spec_ref = "撰写规范（8）目录"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        if not isinstance(sections, dict):
            return []

        catalogue = sections.get("catalogue", {})
        catalogue_content = catalogue.get("content") if isinstance(catalogue, dict) else None
        if not catalogue_content:
            return []

        checker = _create_format_checker(ctx)
        issues: list[Issue] = []
        for index, paragraph in enumerate(catalogue_content, start=1):
            text = _paragraph_text(paragraph)
            prefix = _catalogue_number_prefix(text)
            if not prefix:
                continue

            prefix_runs = _runs_in_prefix(paragraph, prefix)
            digit_runs = [run for run, piece in prefix_runs if re.search(r"\d", piece)]
            if not digit_runs:
                continue

            if _check_runs_font(checker, digit_runs, "Times New Roman"):
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{index}",
                title="目录",
                message=text.strip(),
                problem="目录题序中的阿拉伯数字字体可能不是 Times New Roman",
                section="目录",
                content=text.strip(),
                metadata={
                    "paragraph_index": index,
                    "prefix": prefix,
                    "expected_font": "Times New Roman",
                },
            )

        return issues


class CitationReferenceConsistencyRule(BaseRule):
    rule_id = "docx.stage2.citation_reference_consistency"
    display_name = "Citation and reference consistency"
    spec_ref = "撰写规范（9）中外文参考文献、（10）引文标识、（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        main_text = sections.get("main_text", []) if isinstance(sections, dict) else []
        references = sections.get("references", {}) if isinstance(sections, dict) else {}
        reference_paragraphs = references.get("content") if isinstance(references, dict) else None
        if not isinstance(reference_paragraphs, list) or not reference_paragraphs:
            return []

        reference_entries = [_paragraph_text(p) for p in reference_paragraphs if _paragraph_text(p)]
        if not reference_entries:
            return []

        reference_number_map, unnumbered_entries = _extract_reference_entry_numbers(reference_entries)
        cited_numbers: list[int] = []
        for paragraph in main_text or []:
            cited_numbers.extend(_extract_citation_numbers(_paragraph_text(paragraph)))

        issues: list[Issue] = []
        if unnumbered_entries:
            for index, entry in enumerate(unnumbered_entries[:5], start=1):
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.reference_number_format.{index}",
                    title="参考文献",
                    message=entry,
                    problem="参考文献条目可能未按 [序号] 形式编号",
                    section="参考文献",
                    content=entry,
                )

        if not cited_numbers:
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.no_citation",
                title="引文与参考文献",
                message="正文中未识别到 [1] 这类引文标示",
                problem="无法确认正文引用与参考文献是否一致",
                section="正文",
                content="未识别到方括号引文",
                severity=Severity.INFO,
            )
            return issues

        first_seen_citations: list[int] = []
        seen_citations: set[int] = set()
        for num in cited_numbers:
            if num in seen_citations:
                continue
            seen_citations.add(num)
            first_seen_citations.append(num)

        cited_unique = sorted(seen_citations)
        max_reference_no = max(reference_number_map) if reference_number_map else 0
        out_of_range = [num for num in cited_unique if num not in reference_number_map]
        if out_of_range:
            content = "、".join(f"[{num}]" for num in out_of_range)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.out_of_range",
                title="引文与参考文献",
                message=content,
                problem="正文中的部分引文编号在参考文献列表中未找到",
                section="正文",
                content=content,
                metadata={"out_of_range": out_of_range, "max_reference_no": max_reference_no},
            )

        uncited_references = [num for num in sorted(reference_number_map) if num not in seen_citations]
        if uncited_references:
            content = "、".join(f"[{num}]" for num in uncited_references[:10])
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.uncited_reference",
                title="参考文献",
                message=content,
                problem="部分参考文献可能未在正文中引用",
                section="参考文献",
                content=content,
                metadata={"uncited_references": uncited_references},
            )

        if first_seen_citations != sorted(first_seen_citations):
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.citation_order",
                title="引文与参考文献",
                message="、".join(f"[{num}]" for num in first_seen_citations[:12]),
                problem="正文中引文首次出现的顺序可能不是递增编号",
                section="正文",
                content="、".join(f"[{num}]" for num in first_seen_citations[:12]),
                metadata={"cited_numbers": first_seen_citations},
                severity=Severity.INFO,
            )

        return issues


class FontSizeFormatRule(BaseRule):
    rule_id = "docx.stage2.font_size_format"
    display_name = "Font and size format"
    spec_ref = "撰写规范（3）字体和字号"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        checker = _create_format_checker(ctx)
        issues: list[Issue] = []
        for item in _iter_format_targets(ctx, checker.format_rules):
            paragraph = item["paragraph"]
            text = item["text"]
            expected = item["expected"]
            expected_font = expected.get("font")
            expected_size = expected.get("size")

            font_ok = checker._check_font(paragraph, expected_font)
            size_ok = checker._check_size(paragraph, expected_size)
            display_text = text

            if item["section_key"] == "abstract_keyword" and item["sub_key"] in {
                "chinese_keyword_title",
                "chinese_keyword",
                "english_keyword_title",
                "english_keyword",
            }:
                parts = _extract_keyword_parts(paragraph)
                if parts:
                    if item["sub_key"].endswith("keyword_title"):
                        display_text = parts["title_text"]
                        font_ok = _check_runs_font(checker, parts["title_runs"], expected_font)
                        size_ok = _check_runs_size(checker, parts["title_runs"], expected_size)
                    else:
                        display_text = parts["content_text"] or text
                        font_ok = _check_runs_font(checker, parts["content_runs"], expected_font)
                        size_ok = _check_runs_size(checker, parts["content_runs"], expected_size)

            if font_ok and size_ok:
                continue

            problems: list[str] = []
            if not font_ok and expected_font:
                problems.append(f"字体可能不是{expected_font}")
            if not size_ok and expected_size:
                problems.append(f"字号可能不是{expected_size}")

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{item['section_key']}.{item['sub_key'] or 'content'}.{item['paragraph_index']}",
                title=item["section_label"],
                message=display_text,
                problem="；".join(problems),
                section=item["section_label"],
                content=display_text,
                metadata={
                    "section_key": item["section_key"],
                    "sub_key": item["sub_key"],
                    "paragraph_index": item["paragraph_index"],
                    "expected_font": expected_font,
                    "expected_size": expected_size,
                },
            )

        return issues


class AlignmentFormatRule(BaseRule):
    rule_id = "docx.stage2.alignment_format"
    display_name = "Alignment format"
    spec_ref = "撰写规范（3）字体和字号"
    engine = "docx"

    _allowed_targets = {
        ("cover", "school"),
        ("cover", "title"),
        ("figures_or_tables_title", None),
    }

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        checker = _create_format_checker(ctx)
        issues: list[Issue] = []
        for item in _iter_format_targets(ctx, checker.format_rules):
            if (item["section_key"], item["sub_key"]) not in self._allowed_targets:
                continue

            paragraph = item["paragraph"]
            text = item["text"]
            expected_alignment = item["expected"].get("alignment")
            if not expected_alignment:
                continue

            display_text = text
            if item["section_key"] == "abstract_keyword" and item["sub_key"] in {
                "chinese_keyword_title",
                "chinese_keyword",
                "english_keyword_title",
                "english_keyword",
            }:
                parts = _extract_keyword_parts(paragraph)
                if parts:
                    display_text = parts["title_text"] if item["sub_key"].endswith("keyword_title") else (parts["content_text"] or text)

            if checker._check_alignment(paragraph, expected_alignment):
                continue

            actual_alignment = _effective_alignment(paragraph)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{item['section_key']}.{item['sub_key'] or 'content'}.{item['paragraph_index']}",
                title=item["section_label"],
                message=display_text,
                problem=f"对齐方式可能不是“{expected_alignment}”",
                section=item["section_label"],
                content=display_text,
                metadata={
                    "section_key": item["section_key"],
                    "sub_key": item["sub_key"],
                    "paragraph_index": item["paragraph_index"],
                    "expected_alignment": expected_alignment,
                    "actual_alignment": str(actual_alignment) if actual_alignment is not None else None,
                },
            )

        return issues


class LineSpacingFormatRule(BaseRule):
    rule_id = "docx.stage2.line_spacing_format"
    display_name = "Line spacing and paragraph spacing format"
    spec_ref = "撰写规范（6）页面设置"
    engine = "docx"

    _title_keys = {("headings", "chapter"), ("headings", "level1"), ("headings", "level2"), ("headings", "level3")}
    _skip_targets = {
        ("cover", "school"),
        ("cover", "title"),
        ("cover", "personal_information"),
    }

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        checker = _create_format_checker(ctx)
        issues: list[Issue] = []
        for item in _iter_format_targets(ctx, checker.format_rules):
            if (item["section_key"], item["sub_key"]) in self._skip_targets:
                continue

            paragraph = item["paragraph"]
            text = item["text"]
            expected = item["expected"]
            expected_line_spacing = expected.get("line_spacing")
            line_spacing_ok = checker._check_line_spacing(paragraph, expected_line_spacing) if expected_line_spacing else True
            display_text = text

            if item["section_key"] == "abstract_keyword" and item["sub_key"] in {
                "chinese_keyword_title",
                "chinese_keyword",
                "english_keyword_title",
                "english_keyword",
            }:
                parts = _extract_keyword_parts(paragraph)
                if parts:
                    display_text = parts["title_text"] if item["sub_key"].endswith("keyword_title") else (parts["content_text"] or text)

            problems: list[str] = []
            if not line_spacing_ok and expected_line_spacing:
                problems.append(f"行距可能不是“{expected_line_spacing}”")

            if (item["section_key"], item["sub_key"]) in self._title_keys:
                expected_size_pt = checker._get_font_size_pt(expected.get("size"))
                actual_before = _effective_spacing_value(paragraph, "space_before")
                actual_after = _effective_spacing_value(paragraph, "space_after")
                if not _space_matches_half_line(actual_before, expected_size_pt):
                    problems.append("段前距可能不是 0.5 行")
                if not _space_matches_half_line(actual_after, expected_size_pt):
                    problems.append("段后距可能不是 0.5 行")

            if not problems:
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{item['section_key']}.{item['sub_key'] or 'content'}.{item['paragraph_index']}",
                title=item["section_label"],
                message=display_text,
                problem="；".join(problems),
                section=item["section_label"],
                content=display_text,
                metadata={
                    "section_key": item["section_key"],
                    "sub_key": item["sub_key"],
                    "paragraph_index": item["paragraph_index"],
                    "expected_line_spacing": expected_line_spacing,
                    "actual_space_before_pt": getattr(_effective_spacing_value(paragraph, "space_before"), "pt", None),
                    "actual_space_after_pt": getattr(_effective_spacing_value(paragraph, "space_after"), "pt", None),
                },
            )

        return issues


class PageSettingsRule(BaseRule):
    rule_id = "docx.stage2.page_settings"
    display_name = "Page settings"
    spec_ref = "撰写规范（6）页面设置及（2）打印规范"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        sections = getattr(doc, "sections", None) or []
        if not sections:
            return []

        section = sections[0]
        actual = {
            "纸张宽度": _cm_value(getattr(section, "page_width", None)),
            "纸张高度": _cm_value(getattr(section, "page_height", None)),
            "上边距": _cm_value(getattr(section, "top_margin", None)),
            "下边距": _cm_value(getattr(section, "bottom_margin", None)),
            "左边距": _cm_value(getattr(section, "left_margin", None)),
            "右边距": _cm_value(getattr(section, "right_margin", None)),
            "页眉距离": _cm_value(getattr(section, "header_distance", None)),
            "页脚距离": _cm_value(getattr(section, "footer_distance", None)),
        }
        expected = {
            "纸张宽度": 21.0,
            "纸张高度": 29.7,
            "上边距": 2.5,
            "下边距": 2.0,
            "左边距": 2.5,
            "右边距": 2.0,
            "页眉距离": 2.6,
            "页脚距离": 2.4,
        }

        failed_items: list[str] = []
        for key, target in expected.items():
            if not _approx_equal(actual.get(key), target):
                value = actual.get(key)
                if value is None:
                    failed_items.append(f"{key}（未读取到）")
                else:
                    failed_items.append(f"{key}（当前约 {value:.2f}cm，规范为 {target:.2f}cm）")

        if not failed_items:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="页面设置",
                message="；".join(failed_items),
                severity=Severity.WARNING,
                source=Source.DOCX,
                fixable=False,
                metadata={
                    "section": "页面设置",
                    "content": "首节页面属性",
                    "problem": "部分页面设置可能不符合规范",
                    "actual": actual,
                    "expected": expected,
                },
            )
        ]


class HeaderFormatRule(BaseRule):
    rule_id = "docx.stage2.header_format"
    display_name = "Header format"
    spec_ref = "撰写规范（3）字体和字号、（7）页眉和页码"
    engine = "docx"

    def __init__(self) -> None:
        self.format_checker = FormatChecker()

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        sections = getattr(doc, "sections", None) or []
        if not sections:
            return []

        issues: list[Issue] = []
        expected_text = "武汉理工大学毕业设计（论文）"
        for index, section in enumerate(sections, start=1):
            header = getattr(section, "header", None)
            header_paragraphs = getattr(header, "paragraphs", None) if header is not None else None
            meaningful = list(_iter_nonempty_paragraphs(header_paragraphs))
            if not meaningful:
                continue

            for para_index, paragraph in enumerate(meaningful, start=1):
                text = _paragraph_text(paragraph)
                if text != expected_text:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.text.{index}.{para_index}",
                        title="页眉",
                        message=text,
                        problem="页眉内容可能不符合规范",
                        section="页眉",
                        content=text,
                        metadata={"section_index": index, "paragraph_index": para_index},
                    )

                actual_alignment = getattr(paragraph, "alignment", None)
                if actual_alignment not in (None, WD_ALIGN_PARAGRAPH.CENTER):
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.align.{index}.{para_index}",
                        title="页眉",
                        message=text,
                        problem="页眉可能未居中",
                        section="页眉",
                        content=text,
                        metadata={"section_index": index, "paragraph_index": para_index},
                    )

                font_problem = False
                size_problem = False
                for run in _nonempty_runs(paragraph):
                    candidate_fonts = self.format_checker._get_run_font_candidates(run)
                    if candidate_fonts and not any(
                        self.format_checker._is_font_equivalent("宋体", candidate) for candidate in candidate_fonts
                    ):
                        font_problem = True
                    actual_size = run.font.size.pt if run.font.size else None
                    if actual_size is not None and abs(actual_size - 10.5) > 0.5:
                        size_problem = True

                if font_problem:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.font.{index}.{para_index}",
                        title="页眉",
                        message=text,
                        problem="页眉字体可能不是宋体五号",
                        section="页眉",
                        content=text,
                        metadata={"section_index": index, "paragraph_index": para_index},
                    )
                elif size_problem:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.size.{index}.{para_index}",
                        title="页眉",
                        message=text,
                        problem="页眉字号可能不是五号",
                        section="页眉",
                        content=text,
                        metadata={"section_index": index, "paragraph_index": para_index},
                    )

        return issues


class CaptionFormatRule(BaseRule):
    rule_id = "docx.stage2.caption_format"
    display_name = "Figure and table caption format"
    spec_ref = "撰写规范（16）（17）表格与图"
    engine = "docx"

    _caption_pattern = re.compile(r"^(图|表)\s*([A-Z]?\d+(?:[\.．]\d+)*)(\s*)(.*)$")
    _table_title_punctuation = re.compile(r"[，。！？；：、,.!?;:()\[\]【】《》<>“”\"'‘’·—\-]")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        current_chapter_no: str | None = None
        current_appendix_letter: str | None = None
        for index, paragraph in enumerate(paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            chapter_no = _chapter_number_from_heading(text)
            if chapter_no:
                current_chapter_no = chapter_no
                current_appendix_letter = None

            appendix_letter = _appendix_letter_from_heading(text)
            if appendix_letter:
                current_appendix_letter = appendix_letter

            match = self._caption_pattern.match(text)
            if not match:
                continue

            kind, number, spacing, title_text = match.groups()
            title_text = title_text.strip()
            if not title_text:
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.missing_title.{index}",
                    title=f"{kind}题",
                    message=text,
                    problem=f"{kind}序号后可能缺少标题内容",
                    section=f"{kind}题",
                    content=text,
                    metadata={"index": index, "kind": kind, "number": number},
                )
                continue

            if spacing != " ":
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.spacing.{index}",
                    title=f"{kind}题",
                    message=text,
                    problem=f"{kind}序号与标题之间可能未空一格",
                    section=f"{kind}题",
                    content=text,
                    metadata={"index": index, "kind": kind, "number": number},
                )

            if kind == "表" and self._table_title_punctuation.search(title_text):
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.table_punctuation.{index}",
                    title="表题",
                    message=text,
                    problem="表标题中可能包含标点符号",
                    section="表题",
                    content=text,
                    metadata={"index": index, "number": number, "title_text": title_text},
                )

            if current_appendix_letter:
                if not number.startswith(current_appendix_letter):
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.appendix_number.{index}",
                        title=f"{kind}题",
                        message=text,
                        problem=f"{kind}序号在附录中可能未按“{kind}{current_appendix_letter}1”这类格式编号",
                        section=f"{kind}题",
                        content=text,
                        metadata={"index": index, "kind": kind, "number": number, "appendix_letter": current_appendix_letter},
                    )
            elif current_chapter_no and not number.startswith(current_chapter_no):
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.chapter_number.{index}",
                    title=f"{kind}题",
                    message=text,
                    problem=f"{kind}序号可能未按当前章编号（当前章疑似为第{current_chapter_no}章）",
                    section=f"{kind}题",
                    content=text,
                    metadata={"index": index, "kind": kind, "number": number, "current_chapter_no": current_chapter_no},
                )

        return issues


class FormulaNumberFormatRule(BaseRule):
    rule_id = "docx.stage2.formula_number_format"
    display_name = "Formula number format"
    spec_ref = "撰写规范（15）公式"
    engine = "docx"

    _formula_number_pattern = re.compile(r"[（(]\s*(\d+)\.(\d+)\s*[)）]\s*$")
    _math_hint_pattern = re.compile(r"[=＝+\-－×÷/*^<>≤≥∑Σ∫√α-ωΑ-Ω]")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        current_chapter_no: str | None = None
        for index, paragraph in enumerate(paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            chapter_no = _chapter_number_from_heading(text)
            if chapter_no:
                current_chapter_no = chapter_no
                continue

            match = self._formula_number_pattern.search(text)
            if not match:
                continue

            if not self._math_hint_pattern.search(text):
                continue

            chapter_part = match.group(1)
            sequence_part = match.group(2)
            if sequence_part == "0":
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.sequence_zero.{index}",
                    title="公式",
                    message=text,
                    problem="公式编号序号可能不能为 0",
                    section="公式",
                    content=text,
                    metadata={"index": index, "chapter_part": chapter_part, "sequence_part": sequence_part},
                )

            if current_chapter_no and chapter_part != current_chapter_no:
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.chapter_mismatch.{index}",
                    title="公式",
                    message=text,
                    problem=f"公式编号可能未按当前章编号（当前章疑似为第{current_chapter_no}章）",
                    section="公式",
                    content=text,
                    metadata={"index": index, "chapter_part": chapter_part, "current_chapter_no": current_chapter_no},
                )

        return issues


class AppendixFormatRule(BaseRule):
    rule_id = "docx.stage2.appendix_format"
    display_name = "Appendix format"
    spec_ref = "撰写规范（20）附录"
    engine = "docx"

    _appendix_heading_pattern = re.compile(r"^附录\s*([A-Z])(?:\s+.*)?$")
    _appendix_number_pattern = re.compile(r"^附([A-Z]\d+(?:\.\d+)*)$")
    _appendix_item_pattern = re.compile(r"^(图|表)\s*([A-Z]\d+(?:\.\d+)*)\b")
    _appendix_formula_pattern = re.compile(r"^式[（(]([A-Z]\d+(?:\.\d+)*)[)）]$")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        appendix_letters: list[str] = []
        current_appendix_letter: str | None = None
        for index, paragraph in enumerate(paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            if text.startswith("附录"):
                heading_match = self._appendix_heading_pattern.match(text)
                if heading_match:
                    current_appendix_letter = heading_match.group(1)
                    appendix_letters.append(current_appendix_letter)
                else:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.heading.{index}",
                        title="附录",
                        message=text,
                        problem="附录标题格式可能不符合“附录A、附录B…”",
                        section="附录",
                        content=text,
                        metadata={"index": index},
                    )
                    current_appendix_letter = None

            if text.startswith("附") and not text.startswith("附录"):
                appendix_number_match = self._appendix_number_pattern.match(text)
                if appendix_number_match:
                    appendix_number = appendix_number_match.group(1)
                    if current_appendix_letter and not appendix_number.startswith(current_appendix_letter):
                        _append_issue(
                            issues,
                            rule_id=f"{self.rule_id}.internal_appendix.{index}",
                            title="附录",
                            message=text,
                            problem=f"附录内部编号可能未按“附{current_appendix_letter}1”这类格式编号",
                            section="附录",
                            content=text,
                            metadata={"index": index, "appendix_number": appendix_number, "appendix_letter": current_appendix_letter},
                        )

            item_match = self._appendix_item_pattern.match(text)
            if item_match:
                kind = item_match.group(1)
                item_number = item_match.group(2)
                if current_appendix_letter and not item_number.startswith(current_appendix_letter):
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.{kind}_number.{index}",
                        title="附录",
                        message=text,
                        problem=f"附录中的{kind}编号可能未按“{kind}{current_appendix_letter}1”这类格式编号",
                        section="附录",
                        content=text,
                        metadata={"index": index, "kind": kind, "item_number": item_number, "appendix_letter": current_appendix_letter},
                    )

            formula_match = self._appendix_formula_pattern.match(text)
            if formula_match:
                formula_number = formula_match.group(1)
                if current_appendix_letter and not formula_number.startswith(current_appendix_letter):
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.formula_number.{index}",
                        title="附录",
                        message=text,
                        problem=f"附录中的公式编号可能未按“式（{current_appendix_letter}1）”这类格式编号",
                        section="附录",
                        content=text,
                        metadata={"index": index, "formula_number": formula_number, "appendix_letter": current_appendix_letter},
                    )

        if appendix_letters:
            expected_ord = ord("A")
            for pos, letter in enumerate(appendix_letters, start=1):
                if ord(letter) != expected_ord:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.sequence.{pos}",
                        title="附录",
                        message="、".join(f"附录{item}" for item in appendix_letters),
                        problem="附录字母顺序可能不连续",
                        section="附录",
                        content="、".join(f"附录{item}" for item in appendix_letters),
                        metadata={"letters": appendix_letters},
                    )
                    break
                expected_ord += 1

        return issues
