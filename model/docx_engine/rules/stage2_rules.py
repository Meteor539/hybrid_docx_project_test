import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source
from model.format_checker import FormatChecker


_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_XML_NS = {
    "w": _WORD_NS,
    "m": _MATH_NS,
}


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


def _paragraph_xml_root(paragraph):
    element = getattr(paragraph, "_element", None)
    if element is None:
        return None
    try:
        return etree.fromstring(element.xml.encode("utf-8"))
    except Exception:  # noqa: BLE001
        return None


def _paragraph_has_math(paragraph) -> bool:
    root = _paragraph_xml_root(paragraph)
    if root is None:
        return False
    try:
        return bool(root.xpath(".//m:oMath | .//m:oMathPara", namespaces=_XML_NS))
    except Exception:  # noqa: BLE001
        return False


def _formula_paragraph_text(paragraph) -> str:
    text = _paragraph_text(paragraph)
    if text:
        return text

    root = _paragraph_xml_root(paragraph)
    if root is None:
        return ""

    try:
        tokens = root.xpath(".//m:t/text() | .//w:t/text()", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return ""

    return "".join(str(token) for token in tokens).strip()


def _formula_paragraph_raw_text(paragraph) -> str:
    root = _paragraph_xml_root(paragraph)
    if root is None:
        return ""

    try:
        tokens = root.xpath(".//m:t/text() | .//w:t/text()", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return ""

    return "".join(str(token) for token in tokens)


def _math_paragraph_justification(paragraph) -> str | None:
    root = _paragraph_xml_root(paragraph)
    if root is None:
        return None
    try:
        value = root.xpath("string(.//m:oMathParaPr/m:jc/@m:val)", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return None
    return value.strip() or None


def _has_right_tab_stop(paragraph) -> bool:
    root = _paragraph_xml_root(paragraph)
    if root is None:
        return False
    try:
        tab_values = root.xpath("./w:pPr/w:tabs/w:tab/@w:val", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return False
    return any(str(value).strip().lower() == "right" for value in tab_values)


def _run_spacing_value(run) -> int | None:
    element = getattr(run, "_element", None)
    if element is None:
        return None
    try:
        values = element.xpath("./w:rPr/w:spacing/@w:val", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return None
    if values:
        try:
            return int(values[0])
        except Exception:  # noqa: BLE001
            return None

    style = getattr(run, "style", None)
    style_element = getattr(style, "_element", None) if style is not None else None
    if style_element is None:
        return None
    try:
        values = style_element.xpath("./w:rPr/w:spacing/@w:val", namespaces=_XML_NS)
    except Exception:  # noqa: BLE001
        return None
    if not values:
        return None
    try:
        return int(values[0])
    except Exception:  # noqa: BLE001
        return None


def _paragraph_has_nonstandard_character_spacing(paragraph) -> tuple[bool, int | None]:
    for run in getattr(paragraph, "runs", []) or []:
        run_text = str(getattr(run, "text", "") or "")
        if not run_text.strip():
            continue
        spacing = _run_spacing_value(run)
        if spacing is None:
            continue
        if spacing != 0:
            return True, spacing
    return False, None


def _run_is_superscript(run) -> bool:
    font = getattr(run, "font", None)
    if font is not None and getattr(font, "superscript", None) is True:
        return True

    element = getattr(run, "_element", None)
    if element is not None:
        try:
            values = element.xpath("./w:rPr/w:vertAlign/@w:val", namespaces=_XML_NS)
        except Exception:  # noqa: BLE001
            values = []
        if any(str(value).strip().lower() == "superscript" for value in values):
            return True

    style = getattr(run, "style", None)
    style_element = getattr(style, "_element", None) if style is not None else None
    if style_element is not None:
        try:
            values = style_element.xpath("./w:rPr/w:vertAlign/@w:val", namespaces=_XML_NS)
        except Exception:  # noqa: BLE001
            values = []
        if any(str(value).strip().lower() == "superscript" for value in values):
            return True

    return False


def _iter_in_text_note_marker_runs(paragraph):
    text = _paragraph_text(paragraph)
    if not text:
        return

    trimmed = text.lstrip()
    leading_marker = trimmed[0] if trimmed and trimmed[0] in _NOTE_MARKER_SET else None
    leading_consumed = False

    for run in getattr(paragraph, "runs", []) or []:
        run_text = str(getattr(run, "text", "") or "")
        if not run_text:
            continue
        for marker in _extract_note_markers(run_text):
            if marker == leading_marker and not leading_consumed:
                leading_consumed = True
                continue
            yield marker, run, run_text


def _iter_citation_runs(paragraph):
    for run in getattr(paragraph, "runs", []) or []:
        run_text = str(getattr(run, "text", "") or "")
        if not run_text.strip():
            continue
        matches = list(re.finditer(r"\[\s*([0-9,\-，、\s]+)\s*\]", run_text))
        for match in matches:
            yield match.group(0), run, run_text


def _style_by_id(doc, style_id: str | None):
    if doc is None or not style_id:
        return None
    try:
        for style in getattr(doc, "styles", []) or []:
            if getattr(style, "style_id", None) == style_id:
                return style
    except Exception:  # noqa: BLE001
        return None
    return None


def _footer_page_number_nodes(section, doc) -> list[dict]:
    footer = getattr(section, "footer", None)
    paragraphs = getattr(footer, "paragraphs", None) if footer is not None else None
    if not paragraphs:
        return []

    nodes: list[dict] = []
    seen_keys: set[tuple] = set()
    for footer_para in paragraphs:
        root = _paragraph_xml_root(footer_para)
        if root is None:
            continue
        try:
            page_paras = root.xpath(
                ".//w:txbxContent//w:p[.//w:instrText[contains(translate(., 'page', 'PAGE'), 'PAGE')]]",
                namespaces=_XML_NS,
            )
        except Exception:  # noqa: BLE001
            continue

        for page_para in page_paras:
            style_id = None
            align = None
            text = ""
            font_candidates: list[str] = []
            size_half_points: int | None = None
            try:
                style_values = page_para.xpath("./w:pPr/w:pStyle/@w:val", namespaces=_XML_NS)
                if style_values:
                    style_id = str(style_values[0])
                align_values = page_para.xpath("./w:pPr/w:jc/@w:val", namespaces=_XML_NS)
                if align_values:
                    align = str(align_values[0]).lower()

                text_tokens = page_para.xpath(".//w:t/text()", namespaces=_XML_NS)
                text = "".join(str(token) for token in text_tokens).strip()

                run_nodes = page_para.xpath("./w:r", namespaces=_XML_NS)
                for run_node in run_nodes:
                    run_text = "".join(str(token) for token in run_node.xpath("./w:t/text()", namespaces=_XML_NS)).strip()
                    if not run_text:
                        continue
                    font_values = run_node.xpath("./w:rPr/w:rFonts/@w:ascii | ./w:rPr/w:rFonts/@w:hAnsi", namespaces=_XML_NS)
                    font_candidates.extend(str(value) for value in font_values if str(value).strip())
                    size_values = run_node.xpath("./w:rPr/w:sz/@w:val", namespaces=_XML_NS)
                    if size_values:
                        try:
                            size_half_points = int(size_values[0])
                        except Exception:  # noqa: BLE001
                            pass
            except Exception:  # noqa: BLE001
                continue

            style = _style_by_id(doc, style_id)
            if size_half_points is None and style is not None:
                style_size = getattr(getattr(style, "font", None), "size", None)
                if style_size is not None:
                    try:
                        size_half_points = int(round(float(style_size.pt) * 2))
                    except Exception:  # noqa: BLE001
                        pass

            if not font_candidates and style is not None:
                style_font = getattr(getattr(style, "font", None), "name", None)
                if style_font:
                    font_candidates.append(str(style_font))

            node = {
                "text": text or "页码域",
                "align": align,
                "style_id": style_id,
                "font_candidates": font_candidates,
                "size_half_points": size_half_points,
            }
            key = (
                node["text"],
                node["align"],
                node["style_id"],
                tuple(node["font_candidates"]),
                node["size_half_points"],
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)
            nodes.append(node)
    return nodes


def _footer_text_nodes(section, doc) -> list[dict]:
    footer = getattr(section, "footer", None)
    paragraphs = getattr(footer, "paragraphs", None) if footer is not None else None
    if not paragraphs:
        return []

    nodes: list[dict] = []
    seen_keys: set[tuple] = set()
    for footer_para in paragraphs:
        root = _paragraph_xml_root(footer_para)
        if root is None:
            continue
        try:
            content_paras = root.xpath(".//w:txbxContent//w:p | self::w:p", namespaces=_XML_NS)
        except Exception:  # noqa: BLE001
            continue

        for para_node in content_paras:
            try:
                instr = para_node.xpath(".//w:instrText/text()", namespaces=_XML_NS)
                if any("PAGE" in str(item).upper() for item in instr):
                    continue

                text_tokens = para_node.xpath(".//w:t/text()", namespaces=_XML_NS)
                text = "".join(str(token) for token in text_tokens).strip()
                if not text:
                    continue

                style_values = para_node.xpath("./w:pPr/w:pStyle/@w:val", namespaces=_XML_NS)
                style_id = str(style_values[0]) if style_values else None
                align_values = para_node.xpath("./w:pPr/w:jc/@w:val", namespaces=_XML_NS)
                align = str(align_values[0]).lower() if align_values else None

                font_candidates: list[str] = []
                size_half_points: int | None = None
                run_nodes = para_node.xpath("./w:r", namespaces=_XML_NS)
                for run_node in run_nodes:
                    run_text = "".join(str(token) for token in run_node.xpath("./w:t/text()", namespaces=_XML_NS)).strip()
                    if not run_text:
                        continue
                    font_values = run_node.xpath("./w:rPr/w:rFonts/@w:ascii | ./w:rPr/w:rFonts/@w:hAnsi | ./w:rPr/w:rFonts/@w:eastAsia", namespaces=_XML_NS)
                    font_candidates.extend(str(value) for value in font_values if str(value).strip())
                    size_values = run_node.xpath("./w:rPr/w:sz/@w:val", namespaces=_XML_NS)
                    if size_values:
                        try:
                            size_half_points = int(size_values[0])
                        except Exception:  # noqa: BLE001
                            pass
            except Exception:  # noqa: BLE001
                continue

            style = _style_by_id(doc, style_id)
            if size_half_points is None and style is not None:
                style_size = getattr(getattr(style, "font", None), "size", None)
                if style_size is not None:
                    try:
                        size_half_points = int(round(float(style_size.pt) * 2))
                    except Exception:  # noqa: BLE001
                        pass

            if not font_candidates and style is not None:
                style_font = getattr(getattr(style, "font", None), "name", None)
                if style_font:
                    font_candidates.append(str(style_font))

            node = {
                "text": text,
                "align": align,
                "style_id": style_id,
                "font_candidates": font_candidates,
                "size_half_points": size_half_points,
            }
            key = (
                node["text"],
                node["align"],
                node["style_id"],
                tuple(node["font_candidates"]),
                node["size_half_points"],
            )
            if key in seen_keys:
                continue
            seen_keys.add(key)
            nodes.append(node)
    return nodes


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


def _collect_citation_occurrences(paragraphs) -> list[dict]:
    occurrences: list[dict] = []
    for paragraph_index, paragraph in enumerate(paragraphs or [], start=1):
        text = _paragraph_text(paragraph)
        if not text:
            continue
        numbers = _extract_citation_numbers(text)
        for number in numbers:
            occurrences.append(
                {
                    "number": number,
                    "text": text,
                    "paragraph_index": paragraph_index,
                }
            )
    return occurrences


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


def _extract_reference_body(text: str) -> str:
    match = re.match(r"^[\[［]\s*\d+\s*[\]］]\s*(.*)$", text or "")
    return match.group(1).strip() if match else str(text or "").strip()


def _split_reference_authors(text: str) -> list[str]:
    body = _extract_reference_body(text)
    if not body:
        return []

    # 优先截取到文献类型标记前，降低把题名部分算成作者的概率。
    type_marker = re.search(r"\[[JMDCPS]\]|［[JMDCPS]］", body, flags=re.IGNORECASE)
    candidate = body[: type_marker.start()].strip() if type_marker else body

    # 中文作者一般在第一个句点前。
    chinese_prefix = re.match(r"^([^.]*)\.", candidate)
    if chinese_prefix:
        author_part = chinese_prefix.group(1).strip()
        if re.search(r"[\u4e00-\u9fff]", author_part):
            return [item.strip() for item in re.split(r"[，,、]", author_part) if item.strip()]

    # 英文作者常写为 Name A, Name B, Name C, Name D. Title...
    # 这里保守取到第一个类型标记前，再按逗号切分，只有明显是作者串时才返回。
    if re.search(r"[A-Za-z]", candidate):
        segments = [item.strip() for item in re.split(r",", candidate) if item.strip()]
        if len(segments) >= 4:
            return segments

    return []


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


def _is_left_flush(paragraph) -> bool:
    left_indent = _effective_spacing_value(paragraph, "left_indent")
    first_line_indent = _effective_spacing_value(paragraph, "first_line_indent")

    left_indent_pt = getattr(left_indent, "pt", None)
    first_line_indent_pt = getattr(first_line_indent, "pt", None)

    left_ok = left_indent_pt is None or abs(left_indent_pt) <= 1.0
    first_line_ok = first_line_indent_pt is None or abs(first_line_indent_pt) <= 1.0
    return left_ok and first_line_ok


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


def _extract_level_number_parts(text: str, level: int) -> tuple[int, ...] | None:
    stripped = str(text or "").strip()
    patterns = {
        1: r"^(\d+)\.(\d+)",
        2: r"^(\d+)\.(\d+)\.(\d+)",
        3: r"^(\d+)[\.\)]",
    }
    pattern = patterns.get(level)
    if not pattern:
        return None
    match = re.match(pattern, stripped)
    if not match:
        return None
    return tuple(int(item) for item in match.groups())


_NOTE_MARKERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_NOTE_MARKER_SET = set(_NOTE_MARKERS)
_NOTE_MARKER_ORDER = {marker: index + 1 for index, marker in enumerate(_NOTE_MARKERS)}


def _extract_note_markers(text: str) -> list[str]:
    return [char for char in str(text or "") if char in _NOTE_MARKER_SET]


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
        citation_occurrences = _collect_citation_occurrences(main_text)
        cited_numbers = [item["number"] for item in citation_occurrences]

        issues: list[Issue] = []
        if unnumbered_entries:
            for index, entry in enumerate(unnumbered_entries[:5], start=1):
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.reference_number_format.{index}",
                    title="参考文献",
                    message=entry,
                    problem="参考文献条目可能未按 [序号] 形式编号",
                    section="参考文献内容",
                    content=entry,
                )

        if not cited_numbers:
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.no_citation",
                title="引文与参考文献",
                message="正文中未识别到 [1] 这类引文标示",
                problem="无法确认正文引用与参考文献是否一致",
                section="正文内容",
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
            offending_contexts = []
            for item in citation_occurrences:
                if item["number"] in out_of_range:
                    offending_contexts.append(item["text"])
            content = " | ".join(offending_contexts[:3]) if offending_contexts else "、".join(
                f"[{num}]" for num in out_of_range
            )
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.out_of_range",
                title="引文与参考文献",
                message=content,
                problem="正文中的部分引文编号在参考文献列表中未找到",
                section="正文内容",
                content=content,
                metadata={
                    "out_of_range": out_of_range,
                    "max_reference_no": max_reference_no,
                    "original_content": content,
                    "problem_detail": "、".join(f"[{num}]" for num in out_of_range),
                },
            )

        uncited_references = [num for num in sorted(reference_number_map) if num not in seen_citations]
        if uncited_references:
            uncited_entries = [reference_number_map[num] for num in uncited_references[:3] if num in reference_number_map]
            content = " | ".join(uncited_entries) if uncited_entries else "、".join(
                f"[{num}]" for num in uncited_references[:10]
            )
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.uncited_reference",
                title="参考文献",
                message=content,
                problem="部分参考文献可能未在正文中引用",
                section="参考文献内容",
                content=content,
                metadata={
                    "uncited_references": uncited_references,
                    "original_content": content,
                    "problem_detail": "、".join(f"[{num}]" for num in uncited_references[:10]),
                },
            )

        if first_seen_citations != sorted(first_seen_citations):
            order_contexts = []
            seen_order_numbers: set[int] = set()
            for item in citation_occurrences:
                num = item["number"]
                if num in seen_order_numbers:
                    continue
                seen_order_numbers.add(num)
                order_contexts.append(item["text"])
            content = " | ".join(order_contexts[:3]) if order_contexts else "、".join(
                f"[{num}]" for num in first_seen_citations[:12]
            )
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.citation_order",
                title="引文与参考文献",
                message=content,
                problem="正文中引文首次出现的顺序可能不是递增编号",
                section="正文内容",
                content=content,
                metadata={
                    "cited_numbers": first_seen_citations,
                    "original_content": content,
                    "problem_detail": "、".join(f"[{num}]" for num in first_seen_citations[:12]),
                },
                severity=Severity.INFO,
            )

        return issues


class CitationSuperscriptRule(BaseRule):
    rule_id = "docx.stage2.citation_superscript"
    display_name = "Citation superscript"
    spec_ref = "撰写规范（10）引文标识"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        main_text = sections.get("main_text", []) if isinstance(sections, dict) else []
        if not isinstance(main_text, list) or not main_text:
            return []

        issues: list[Issue] = []
        for paragraph_index, paragraph in enumerate(main_text, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            for marker_text, run, run_text in _iter_citation_runs(paragraph):
                if _run_is_superscript(run):
                    continue

                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.{paragraph_index}.{marker_text}",
                    title="引文标识",
                    message=text,
                    problem=f"引文编号 {marker_text} 可能未采用方括号上标形式",
                    section="正文内容",
                    content=text,
                    metadata={
                        "paragraph_index": paragraph_index,
                        "marker": marker_text,
                        "run_text": run_text,
                        "original_content": text,
                    },
                )
                break

        return issues


class ReferenceEntryFormatRule(BaseRule):
    rule_id = "docx.stage2.reference_entry_format"
    display_name = "Reference entry numbering and indentation"
    spec_ref = "撰写规范（19）参考文献"
    engine = "docx"

    _entry_pattern = re.compile(r"^[\[［]\s*\d+\s*[\]］]")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        references = sections.get("references", {}) if isinstance(sections, dict) else {}
        reference_paragraphs = references.get("content") if isinstance(references, dict) else None
        if not isinstance(reference_paragraphs, list) or not reference_paragraphs:
            return []

        issues: list[Issue] = []
        for index, paragraph in enumerate(reference_paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            problems: list[str] = []
            if not self._entry_pattern.match(text):
                problems.append("参考文献序号可能未按数字加方括号表示")
            if not _is_left_flush(paragraph):
                problems.append("参考文献序号可能未左顶格")

            if not problems:
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{index}",
                title="参考文献",
                message=text,
                problem="；".join(problems),
                section="参考文献内容",
                content=text,
                metadata={
                    "paragraph_index": index,
                    "left_indent_pt": getattr(_effective_spacing_value(paragraph, "left_indent"), "pt", None),
                    "first_line_indent_pt": getattr(_effective_spacing_value(paragraph, "first_line_indent"), "pt", None),
                },
            )

        return issues


class ReferenceNumberSequenceRule(BaseRule):
    rule_id = "docx.stage2.reference_number_sequence"
    display_name = "Reference number sequence"
    spec_ref = "撰写规范（9）中外文参考文献、（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        references = sections.get("references", {}) if isinstance(sections, dict) else {}
        reference_paragraphs = references.get("content") if isinstance(references, dict) else None
        if not isinstance(reference_paragraphs, list) or not reference_paragraphs:
            return []

        numbered_items: list[tuple[int, str, int]] = []
        for index, paragraph in enumerate(reference_paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue
            match = re.match(r"^[\[［]\s*(\d+)\s*[\]］]", text)
            if not match:
                continue
            numbered_items.append((int(match.group(1)), text, index))

        if len(numbered_items) <= 1:
            return []

        issues: list[Issue] = []
        numbers = [item[0] for item in numbered_items]

        seen_positions: dict[int, int] = {}
        duplicate_numbers: list[int] = []
        for pos, number in enumerate(numbers):
            if number in seen_positions:
                duplicate_numbers.append(number)
            else:
                seen_positions[number] = pos

        if duplicate_numbers:
            duplicate_entries = [
                text
                for pos, (number, text, _) in enumerate(numbered_items)
                if pos > 0 and number in duplicate_numbers
            ]
            content = " | ".join(duplicate_entries[:3]) if duplicate_entries else "、".join(
                f"[{num}]" for num in sorted(set(duplicate_numbers))
            )
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.duplicate",
                title="参考文献",
                message=content,
                problem="参考文献编号可能存在重复",
                section="参考文献内容",
                content=content,
                metadata={
                    "duplicate_numbers": sorted(set(duplicate_numbers)),
                    "original_content": content,
                },
            )

        expected_numbers = list(range(1, len(numbered_items) + 1))
        if numbers != expected_numbers:
            gap_or_order = []
            mismatched_entries: list[str] = []
            for pos, (expected, actual) in enumerate(zip(expected_numbers, numbers)):
                if expected != actual:
                    gap_or_order.append((expected, actual))
                    if pos < len(numbered_items):
                        mismatched_entries.append(numbered_items[pos][1])
            if len(numbers) > len(expected_numbers):
                gap_or_order.extend((None, actual) for actual in numbers[len(expected_numbers):])

            if gap_or_order:
                display_pairs = "；".join(
                    f"应为[{expected}]，实际[{actual}]"
                    for expected, actual in gap_or_order[:5]
                    if expected is not None
                )
                if not display_pairs:
                    display_pairs = "、".join(f"[{num}]" for num in numbers[:10])
                original_content = " | ".join(mismatched_entries[:3]) if mismatched_entries else "、".join(
                    text for _, text, _ in numbered_items[:3]
                )
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.non_consecutive",
                    title="参考文献",
                    message=original_content,
                    problem="参考文献编号顺序可能不是从 [1] 开始连续递增",
                    section="参考文献内容",
                    content=original_content,
                    metadata={"numbers": numbers, "original_content": original_content, "problem_detail": display_pairs},
                )

        return issues


class ReferenceAuthorCountRule(BaseRule):
    rule_id = "docx.stage2.reference_author_count"
    display_name = "Reference author count abbreviation"
    spec_ref = "撰写规范（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        references = sections.get("references", {}) if isinstance(sections, dict) else {}
        reference_paragraphs = references.get("content") if isinstance(references, dict) else None
        if not isinstance(reference_paragraphs, list) or not reference_paragraphs:
            return []

        issues: list[Issue] = []
        for index, paragraph in enumerate(reference_paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            if re.search(r"(，等\b|,?\s*et\s+al\.?)", text, flags=re.IGNORECASE):
                continue

            authors = _split_reference_authors(text)
            if len(authors) <= 3:
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{index}",
                title="参考文献",
                message=text,
                problem="作者人数可能超过 3 位，但未写“等”或“et al.”",
                section="参考文献内容",
                content=text,
                severity=Severity.INFO,
                metadata={
                    "paragraph_index": index,
                    "author_count_guess": len(authors),
                    "original_content": text,
                },
            )

        return issues


class HeadingNumberHierarchyRule(BaseRule):
    rule_id = "docx.stage2.heading_number_hierarchy"
    display_name = "Heading number hierarchy and sequence"
    spec_ref = "撰写规范（4）目录、（9）正文"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        headings = sections.get("headings", {}) if isinstance(sections, dict) else {}
        if not isinstance(headings, dict):
            return []

        issues: list[Issue] = []
        chapter_numbers: list[tuple[int, str]] = []
        for paragraph in headings.get("chapter", []) or []:
            text = _paragraph_text(paragraph)
            chapter_no = _chapter_number_from_heading(text)
            if chapter_no and chapter_no.isdigit():
                chapter_numbers.append((int(chapter_no), text))

        if chapter_numbers:
            expected = 1
            for actual, text in chapter_numbers:
                if actual != expected:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.chapter.{expected}",
                        title="章节标题",
                        message=text,
                        problem=f"章节编号可能未按顺序递增（期望第{expected}章，实际第{actual}章）",
                        section="章节标题",
                        content=text,
                        metadata={"original_content": text},
                    )
                    break
                expected += 1

        chapter_set = {number for number, _ in chapter_numbers}

        level1_items: list[tuple[tuple[int, int], str]] = []
        for paragraph in headings.get("level1", []) or []:
            text = _paragraph_text(paragraph)
            parts = _extract_level_number_parts(text, 1)
            if parts:
                level1_items.append((parts, text))

        level1_by_chapter: dict[int, list[tuple[int, str]]] = {}
        for (chapter_no, seq_no), text in level1_items:
            level1_by_chapter.setdefault(chapter_no, []).append((seq_no, text))
            if chapter_set and chapter_no not in chapter_set:
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.level1.chapter.{chapter_no}.{seq_no}",
                    title="一级标题",
                    message=text,
                    problem="一级标题编号中的章号可能与正文章节不一致",
                    section="一级标题",
                    content=text,
                    metadata={"original_content": text},
                )

        for chapter_no, items in level1_by_chapter.items():
            expected = 1
            for seq_no, text in items:
                if seq_no != expected:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.level1.sequence.{chapter_no}.{expected}",
                        title="一级标题",
                        message=text,
                        problem=f"一级标题编号可能未连续（当前章期望 {chapter_no}.{expected}，实际 {chapter_no}.{seq_no}）",
                        section="一级标题",
                        content=text,
                        metadata={"original_content": text},
                    )
                    break
                expected += 1

        level2_items: list[tuple[tuple[int, int, int], str]] = []
        for paragraph in headings.get("level2", []) or []:
            text = _paragraph_text(paragraph)
            parts = _extract_level_number_parts(text, 2)
            if parts:
                level2_items.append((parts, text))

        level1_parent_set = {parts for parts, _ in level1_items}
        level2_by_parent: dict[tuple[int, int], list[tuple[int, str]]] = {}
        for (chapter_no, parent_no, seq_no), text in level2_items:
            parent_key = (chapter_no, parent_no)
            level2_by_parent.setdefault(parent_key, []).append((seq_no, text))
            if level1_parent_set and parent_key not in level1_parent_set:
                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.level2.parent.{chapter_no}.{parent_no}.{seq_no}",
                    title="二级标题",
                    message=text,
                    problem="二级标题编号对应的上级标题可能不存在",
                    section="二级标题",
                    content=text,
                    metadata={"original_content": text},
                )

        for (chapter_no, parent_no), items in level2_by_parent.items():
            expected = 1
            for seq_no, text in items:
                if seq_no != expected:
                    _append_issue(
                        issues,
                        rule_id=f"{self.rule_id}.level2.sequence.{chapter_no}.{parent_no}.{expected}",
                        title="二级标题",
                        message=text,
                        problem=f"二级标题编号可能未连续（当前小节期望 {chapter_no}.{parent_no}.{expected}，实际 {chapter_no}.{parent_no}.{seq_no}）",
                        section="二级标题",
                        content=text,
                        metadata={"original_content": text},
                    )
                    break
                expected += 1

        return issues


class NoteMarkerConsistencyRule(BaseRule):
    rule_id = "docx.stage2.note_marker_consistency"
    display_name = "Note marker consistency"
    spec_ref = "撰写规范（18）注释"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        paragraphs = getattr(doc, "paragraphs", None) or []
        if not paragraphs:
            return []

        in_text_markers: list[tuple[str, str]] = []
        note_entries: list[tuple[str, str]] = []

        for paragraph in paragraphs:
            text = _paragraph_text(paragraph)
            if not text:
                continue

            markers = _extract_note_markers(text)
            if not markers:
                continue

            leading_marker = text[0] if text and text[0] in _NOTE_MARKER_SET else None
            if leading_marker is not None:
                note_entries.append((leading_marker, text))
                markers = markers[1:]

            for marker in markers:
                in_text_markers.append((marker, text))

        if not in_text_markers and not note_entries:
            return []

        issues: list[Issue] = []
        in_text_first_seen: list[str] = []
        in_text_seen_set: set[str] = set()
        for marker, text in in_text_markers:
            if marker in in_text_seen_set:
                continue
            in_text_seen_set.add(marker)
            in_text_first_seen.append(marker)

        note_order = [marker for marker, _ in note_entries]
        note_set = {marker for marker, _ in note_entries}

        missing_entries = [marker for marker in in_text_first_seen if marker not in note_set]
        if missing_entries:
            contexts = [text for marker, text in in_text_markers if marker in missing_entries]
            content = " | ".join(contexts[:3]) if contexts else "、".join(missing_entries)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.missing_entry",
                title="注释",
                message=content,
                problem="正文中的部分注释标记可能未找到对应注释条目",
                section="正文内容",
                content=content,
                metadata={
                    "original_content": content,
                    "problem_detail": "、".join(missing_entries),
                },
                severity=Severity.INFO,
            )

        extra_entries = [marker for marker in note_order if marker not in in_text_seen_set]
        if extra_entries:
            contexts = [text for marker, text in note_entries if marker in extra_entries]
            content = " | ".join(contexts[:3]) if contexts else "、".join(extra_entries)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.extra_entry",
                title="注释",
                message=content,
                problem="部分注释条目可能未在正文中找到对应标记",
                section="注释内容",
                content=content,
                metadata={
                    "original_content": content,
                    "problem_detail": "、".join(extra_entries),
                },
                severity=Severity.INFO,
            )

        comparable_note_order = [marker for marker in note_order if marker in in_text_seen_set]
        if in_text_first_seen and comparable_note_order and comparable_note_order != in_text_first_seen[: len(comparable_note_order)]:
            contexts = [text for _, text in note_entries[:3]]
            content = " | ".join(contexts) if contexts else "、".join(comparable_note_order)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.order",
                title="注释",
                message=content,
                problem="注释条目序号与正文标记顺序可能不一致",
                section="注释内容",
                content=content,
                metadata={
                    "original_content": content,
                    "problem_detail": f"正文标记顺序：{'、'.join(in_text_first_seen)}；注释条目顺序：{'、'.join(comparable_note_order)}",
                },
                severity=Severity.INFO,
            )

        sorted_markers = sorted(in_text_seen_set | note_set, key=lambda item: _NOTE_MARKER_ORDER.get(item, 999))
        expected_markers = _NOTE_MARKERS[: len(sorted_markers)]
        if sorted_markers and list(sorted_markers) != list(expected_markers):
            content = "、".join(sorted_markers)
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.sequence",
                title="注释",
                message=content,
                problem="注释序号可能未按 ①②③ 顺序连续使用",
                section="注释内容",
                content=content,
                metadata={
                    "original_content": content,
                    "problem_detail": f"当前识别到：{content}",
                },
                severity=Severity.INFO,
            )

        return issues


class NoteMarkerSuperscriptRule(BaseRule):
    rule_id = "docx.stage2.note_marker_superscript"
    display_name = "Note marker superscript"
    spec_ref = "撰写规范（18）注释"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        for paragraph_index, paragraph in enumerate(paragraphs, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue

            for marker, run, run_text in _iter_in_text_note_marker_runs(paragraph):
                if _run_is_superscript(run):
                    continue

                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.{paragraph_index}.{ord(marker)}",
                    title="注释",
                    message=text,
                    problem="正文中的注释标记可能未以上标形式显示",
                    section="正文内容",
                    content=text,
                    metadata={
                        "paragraph_index": paragraph_index,
                        "marker": marker,
                        "run_text": run_text,
                        "original_content": text,
                    },
                )
                break

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
    _fixed_20pt_targets = {
        ("statement", "content"),
        ("abstract_keyword", "chinese_content"),
        ("abstract_keyword", "chinese_keyword_title"),
        ("abstract_keyword", "chinese_keyword"),
        ("abstract_keyword", "english_content"),
        ("abstract_keyword", "english_keyword_title"),
        ("abstract_keyword", "english_keyword"),
        ("catalogue", "content"),
        ("main_text", None),
        ("figures_or_tables_title", None),
        ("references", "content"),
        ("acknowledgments", "content"),
    }
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
            expected_line_spacing = expected.get("line_spacing") if (item["section_key"], item["sub_key"]) in self._fixed_20pt_targets else None
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


class CharacterSpacingFormatRule(BaseRule):
    rule_id = "docx.stage2.character_spacing_format"
    display_name = "Character spacing format"
    spec_ref = "撰写规范（6）页面设置"
    engine = "docx"

    _skip_targets = {
        ("cover", "school"),
        ("cover", "title"),
        ("headings", "chapter"),
        ("headings", "level1"),
        ("headings", "level2"),
        ("headings", "level3"),
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

            has_nonstandard_spacing, spacing_value = _paragraph_has_nonstandard_character_spacing(paragraph)
            if not has_nonstandard_spacing:
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{item['section_key']}.{item['sub_key'] or 'content'}.{item['paragraph_index']}",
                title=item["section_label"],
                message=display_text,
                problem="字符间距可能不是“标准”",
                section=item["section_label"],
                content=display_text,
                metadata={
                    "section_key": item["section_key"],
                    "sub_key": item["sub_key"],
                    "paragraph_index": item["paragraph_index"],
                    "spacing_value": spacing_value,
                },
            )

        return issues


class PageNumberFormatRule(BaseRule):
    rule_id = "docx.stage2.page_number_format"
    display_name = "Page number format"
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

        grouped: dict[tuple[str, tuple[str, ...], tuple[str, ...], int | None], dict] = {}
        for index, section in enumerate(sections, start=1):
            page_nodes = _footer_page_number_nodes(section, doc)
            if not page_nodes:
                continue

            for node_index, node in enumerate(page_nodes, start=1):
                problems: list[str] = []
                if node.get("align") not in {None, "center"}:
                    problems.append("页码所在页脚段落可能未居中")

                size_half_points = node.get("size_half_points")
                if size_half_points is not None:
                    actual_pt = size_half_points / 2.0
                    if abs(actual_pt - 10.5) > 0.5:
                        problems.append("页码字号可能不是五号")

                font_candidates = node.get("font_candidates") or []
                if font_candidates:
                    if not any(self.format_checker._is_font_equivalent("Times New Roman", candidate) for candidate in font_candidates):
                        problems.append("页码字体可能不是Times New Roman")

                if not problems:
                    continue

                key = (
                    "；".join(problems),
                    tuple(font_candidates),
                    (node.get("align"),),
                    size_half_points,
                )
                bucket = grouped.setdefault(
                    key,
                    {
                        "sections": [],
                        "node_indices": [],
                        "font_candidates": font_candidates,
                        "size_half_points": size_half_points,
                        "problem": "；".join(problems),
                    },
                )
                bucket["sections"].append(index)
                bucket["node_indices"].append(node_index)

        issues: list[Issue] = []
        for order, bucket in enumerate(grouped.values(), start=1):
            sections_text = "、".join(f"第{item}节" for item in bucket["sections"])
            content = f"{sections_text}页码域"
            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{order}",
                title="页码",
                message=content,
                problem=bucket["problem"],
                section="页码",
                content=content,
                metadata={
                    "section_indices": bucket["sections"],
                    "node_indices": bucket["node_indices"],
                    "font_candidates": bucket["font_candidates"],
                    "size_half_points": bucket["size_half_points"],
                },
            )

        return issues


class FooterFormatRule(BaseRule):
    rule_id = "docx.stage2.footer_format"
    display_name = "Footer format"
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
        for index, section in enumerate(sections, start=1):
            footer_nodes = _footer_text_nodes(section, doc)
            if not footer_nodes:
                continue

            for node_index, node in enumerate(footer_nodes, start=1):
                problems: list[str] = []
                if node.get("align") not in {None, "center"}:
                    problems.append("页脚可能未居中")

                size_half_points = node.get("size_half_points")
                if size_half_points is not None:
                    actual_pt = size_half_points / 2.0
                    if abs(actual_pt - 10.5) > 0.5:
                        problems.append("页脚字号可能不是五号")

                font_candidates = node.get("font_candidates") or []
                if font_candidates:
                    if not any(self.format_checker._is_font_equivalent("宋体", candidate) for candidate in font_candidates):
                        problems.append("页脚字体可能不是宋体")

                if not problems:
                    continue

                _append_issue(
                    issues,
                    rule_id=f"{self.rule_id}.{index}.{node_index}",
                    title="页脚",
                    message=str(node.get("text") or "页脚内容"),
                    problem="；".join(problems),
                    section="页脚",
                    content=str(node.get("text") or "页脚内容"),
                    metadata={
                        "section_index": index,
                        "node_index": node_index,
                        "font_candidates": font_candidates,
                        "size_half_points": size_half_points,
                        "style_id": node.get("style_id"),
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
            text = _formula_paragraph_text(paragraph)
            has_math = _paragraph_has_math(paragraph)
            if not text:
                continue

            chapter_no = _chapter_number_from_heading(text)
            if chapter_no:
                current_chapter_no = chapter_no
                continue

            match = self._formula_number_pattern.search(text)
            if not match:
                continue

            if not has_math and not self._math_hint_pattern.search(text):
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
                    section="公式编号",
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
                    section="公式编号",
                    content=text,
                    metadata={"index": index, "chapter_part": chapter_part, "current_chapter_no": current_chapter_no},
                )

        return issues


class FormulaAlignmentRule(BaseRule):
    rule_id = "docx.stage2.formula_alignment"
    display_name = "Formula alignment"
    spec_ref = "撰写规范（15）公式"
    engine = "docx"

    _formula_number_pattern = re.compile(r"[（(]\s*(\d+)\.(\d+)\s*[)）]\s*$")
    _math_hint_pattern = re.compile(r"[=＝+\-－×÷/*^<>≤≥∑Σ∫√α-ωΑ-Ω]")

    def __init__(self) -> None:
        self.format_checker = FormatChecker()

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        for index, paragraph in enumerate(paragraphs, start=1):
            text = _formula_paragraph_text(paragraph)
            has_math = _paragraph_has_math(paragraph)
            if not text:
                continue

            if not self._formula_number_pattern.search(text):
                continue
            if not has_math and not self._math_hint_pattern.search(text):
                continue
            if self.format_checker._check_alignment(paragraph, "居中"):
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{index}",
                title="公式",
                message=text,
                problem="公式所在段落可能未居中",
                section="公式内容",
                content=text,
                metadata={"index": index, "original_content": text},
            )

        return issues


class FormulaNumberRightAlignedRule(BaseRule):
    rule_id = "docx.stage2.formula_number_right_aligned"
    display_name = "Formula number right-end alignment"
    spec_ref = "撰写规范（15）公式"
    engine = "docx"

    _formula_number_pattern = re.compile(r"[（(]\s*([A-Z]?\d+(?:\.\d+)*)\s*[)）]\s*$")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        issues: list[Issue] = []
        paragraphs = getattr(doc, "paragraphs", None) or []
        for index, paragraph in enumerate(paragraphs, start=1):
            if not _paragraph_has_math(paragraph):
                continue

            raw_text = _formula_paragraph_raw_text(paragraph)
            display_text = _formula_paragraph_text(paragraph)
            if not raw_text or not display_text:
                continue

            match = self._formula_number_pattern.search(raw_text)
            if not match:
                continue

            if _has_right_tab_stop(paragraph) or "\t" in raw_text:
                continue

            math_jc = _math_paragraph_justification(paragraph)
            gap_text = raw_text[: match.start()]
            trailing_gap = re.search(r"(\s+)$", gap_text)
            gap_len = len(trailing_gap.group(1)) if trailing_gap else 0

            problems: list[str] = []
            if gap_len >= 1:
                problems.append("公式编号前疑似仅用空格与公式分隔")
            if math_jc == "left":
                problems.append("公式对象内部对齐方式疑似为左对齐")
            elif math_jc is None and gap_len < 1:
                problems.append("未识别到将公式编号稳定定位到右侧行末的结构")

            if not problems:
                continue

            _append_issue(
                issues,
                rule_id=f"{self.rule_id}.{index}",
                title="公式编号",
                message=display_text,
                problem="；".join(problems),
                section="公式编号",
                content=display_text,
                metadata={
                    "index": index,
                    "original_content": display_text,
                    "math_justification": math_jc,
                    "space_before_number": gap_len,
                },
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
