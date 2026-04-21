import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


PAGE_ROLE_OTHER = "other"
PAGE_ROLE_CN_ABSTRACT = "cn_abstract"
PAGE_ROLE_EN_ABSTRACT = "en_abstract"
PAGE_ROLE_CATALOGUE = "catalogue"
PAGE_ROLE_MAIN = "main"


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _normalize_page_number_text(text: str) -> str:
    normalized = _normalize_text(text)
    normalized = normalized.strip("()（）[]【】")
    return normalized


def _page_number_candidates(page):
    pattern = re.compile(r"^(?:\d+|[IVXLCDMivxlcdm]+)$")
    page_height = float(getattr(page, "height", 0.0) or 0.0)
    page_width = float(getattr(page, "width", 0.0) or 0.0)
    candidates = []
    for span in getattr(page, "spans", []):
        bbox = getattr(span, "bbox", None) or []
        if len(bbox) != 4:
            continue

        normalized = _normalize_page_number_text(getattr(span, "text", ""))
        if not normalized or not pattern.fullmatch(normalized):
            continue

        x0, y0, x1, y1 = [float(x) for x in bbox]
        if page_height > 0 and y1 < page_height * 0.78:
            continue

        center_x = (x0 + x1) / 2
        candidates.append(
            {
                "text": normalized,
                "bbox": [x0, y0, x1, y1],
                "kind": "arabic" if normalized.isdigit() else "roman",
                "center_offset_ratio": abs(center_x - (page_width / 2)) / page_width if page_width > 0 else 1.0,
            }
        )
    return candidates


def _best_page_number_candidate(page):
    candidates = _page_number_candidates(page)
    if not candidates:
        return None
    return min(candidates, key=lambda item: (item["center_offset_ratio"], -item["bbox"][3]))


def _roman_to_int(text: str) -> int | None:
    values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    candidate = (text or "").upper()
    if not candidate or not re.fullmatch(r"[IVXLCDM]+", candidate):
        return None
    total = 0
    prev = 0
    for char in reversed(candidate):
        value = values[char]
        if value < prev:
            total -= value
        else:
            total += value
            prev = value
    return total


def _page_has_heading(page, keywords: tuple[str, ...]) -> bool:
    normalized = _normalize_text(getattr(page, "text", ""))
    return any(keyword in normalized for keyword in keywords)


def _is_catalogue_page(page) -> bool:
    normalized = _normalize_text(getattr(page, "text", ""))
    if "目录" in normalized or "目錄" in normalized:
        return True
    dotted_entries = len(re.findall(r"\d+\.\d+|\d+$", getattr(page, "text", ""), flags=re.MULTILINE))
    return dotted_entries >= 8 and "第1章" in normalized


def _build_page_roles(pages) -> dict[int, str]:
    roles = {getattr(page, "page_no", idx + 1): PAGE_ROLE_OTHER for idx, page in enumerate(pages)}
    if not pages:
        return roles

    for idx, page in enumerate(pages):
        if _is_catalogue_page(page):
            roles[getattr(page, "page_no", idx + 1)] = PAGE_ROLE_CATALOGUE

    first_main_idx = None
    for idx, page in enumerate(pages):
        if roles.get(getattr(page, "page_no", idx + 1)) == PAGE_ROLE_CATALOGUE:
            continue
        if _page_has_heading(page, ("第1章", "第一章", "绪论", "緒論")):
            first_main_idx = idx
            break

    if first_main_idx is None:
        return roles

    backmatter_indices = []
    for idx in range(first_main_idx + 1, len(pages)):
        if _page_has_heading(pages[idx], ("参考文献", "附录", "附錄", "致谢", "致謝")):
            backmatter_indices.append(idx)
    main_end_idx = min(backmatter_indices) if backmatter_indices else len(pages)

    for idx in range(first_main_idx, main_end_idx):
        roles[getattr(pages[idx], "page_no", idx + 1)] = PAGE_ROLE_MAIN

    for idx in range(first_main_idx):
        page = pages[idx]
        page_no = getattr(page, "page_no", idx + 1)
        if roles.get(page_no) == PAGE_ROLE_CATALOGUE:
            continue
        if _page_has_heading(page, ("摘要",)) and not _page_has_heading(page, ("ABSTRACT", "Abstract")):
            roles[page_no] = PAGE_ROLE_CN_ABSTRACT
        elif _page_has_heading(page, ("ABSTRACT", "Abstract")):
            roles[page_no] = PAGE_ROLE_EN_ABSTRACT

    return roles


def _expected_page_number_kind(role: str) -> str | None:
    if role in {PAGE_ROLE_CN_ABSTRACT, PAGE_ROLE_EN_ABSTRACT}:
        return "roman"
    if role == PAGE_ROLE_MAIN:
        return "arabic"
    return None


def _top_area_texts(page) -> list[str]:
    page_height = float(getattr(page, "height", 0.0) or 0.0)
    if page_height <= 0:
        return []

    texts = []
    for span in getattr(page, "spans", []):
        text = getattr(span, "text", "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4:
            continue
        if float(bbox[3]) > page_height * 0.15:
            continue
        texts.append(text)
    return texts


def _top_area_has_expected_header(page, expected_header: str) -> tuple[bool, str]:
    top_texts = _top_area_texts(page)
    merged = _normalize_text("".join(top_texts))
    expected = _normalize_text(expected_header)
    content = " ".join(top_texts[:5]).strip() or "未识别到顶部页眉文本"
    return (bool(expected and expected in merged), content)


class PageNumberPresencePdfRule(BaseRule):
    rule_id = "page_number.presence"
    display_name = "Page number presence (pdf)"
    spec_ref = "撰写规范（7）页眉和页码"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        roles = _build_page_roles(pages)
        issues: list[Issue] = []
        for page in pages:
            page_no = getattr(page, "page_no", None)
            role = roles.get(page_no, PAGE_ROLE_OTHER)
            expected_kind = _expected_page_number_kind(role)
            if expected_kind is None:
                continue

            best = _best_page_number_candidate(page)
            if best is not None:
                continue

            role_name = "中文摘要" if role == PAGE_ROLE_CN_ABSTRACT else "英文摘要" if role == PAGE_ROLE_EN_ABSTRACT else "正文"
            issues.append(
                Issue(
                    rule_id=self.rule_id,
                    title="页码可能缺失",
                    message=f"第{page_no}页属于{role_name}部分，但页面底端未识别到页码。",
                    severity=Severity.INFO,
                    source=Source.PDF,
                    page=page_no,
                    fixable=False,
                    metadata={
                        "section": "页码",
                        "content": role_name,
                        "problem": "页面底端未识别到页码",
                    },
                )
            )
        return issues


class PageNumberBottomCenterPdfRule(BaseRule):
    rule_id = "page_number.bottom_center"
    display_name = "Page number bottom-center check (pdf)"
    spec_ref = "撰写规范（7）页眉和页码"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        roles = _build_page_roles(pages)
        issues: list[Issue] = []
        for page in pages:
            page_no = getattr(page, "page_no", None)
            if _expected_page_number_kind(roles.get(page_no, PAGE_ROLE_OTHER)) is None:
                continue

            best = _best_page_number_candidate(page)
            if best is None or best["center_offset_ratio"] <= 0.12:
                continue

            issues.append(
                Issue(
                    rule_id=self.rule_id,
                    title="页码",
                    message=f"第{page_no}页页码“{best['text']}”可能未位于页面底端居中。",
                    severity=Severity.WARNING,
                    source=Source.PDF,
                    page=page_no,
                    bbox=[int(x) for x in best["bbox"]],
                    fixable=False,
                    metadata={
                        "section": "页码",
                        "content": best["text"],
                        "problem": "页码可能未位于页面底端居中",
                        "center_offset_ratio": best["center_offset_ratio"],
                    },
                )
            )

        return issues


class HeaderTopContentPdfRule(BaseRule):
    rule_id = "header.top_content"
    display_name = "Header top-content check (pdf)"
    spec_ref = "撰写规范（7）页眉和页码"
    engine = "pdf"

    expected_header = "武汉理工大学毕业设计（论文）"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        roles = _build_page_roles(pages)
        expected = _normalize_text(self.expected_header)
        issues: list[Issue] = []
        for page in pages:
            page_no = getattr(page, "page_no", None)
            if roles.get(page_no, PAGE_ROLE_OTHER) != PAGE_ROLE_MAIN:
                continue

            matched, content = _top_area_has_expected_header(page, self.expected_header)
            if matched:
                continue

            issues.append(
                Issue(
                    rule_id=self.rule_id,
                    title="页眉",
                    message=f"第{page_no}页顶部未识别到规范页眉内容。",
                    severity=Severity.WARNING,
                    source=Source.PDF,
                    page=page_no,
                    fixable=False,
                    metadata={
                        "section": "页眉",
                        "content": content,
                        "problem": "正文页顶部未识别到规范页眉内容",
                    },
                )
            )

        return issues


class HeaderStartBoundaryPdfRule(BaseRule):
    rule_id = "header.start_boundary"
    display_name = "Header starts from chapter one boundary check (pdf)"
    spec_ref = "撰写规范（7）页眉和页码"
    engine = "pdf"

    expected_header = "武汉理工大学毕业设计（论文）"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        roles = _build_page_roles(pages)
        main_pages = [getattr(page, "page_no", idx + 1) for idx, page in enumerate(pages) if roles.get(getattr(page, "page_no", idx + 1)) == PAGE_ROLE_MAIN]
        if not main_pages:
            return []

        first_main_page_no = min(main_pages)
        issues: list[Issue] = []
        for idx, page in enumerate(pages):
            page_no = getattr(page, "page_no", idx + 1)
            if page_no >= first_main_page_no:
                continue

            matched, content = _top_area_has_expected_header(page, self.expected_header)
            if not matched:
                continue

            role = roles.get(page_no, PAGE_ROLE_OTHER)
            role_name_map = {
                PAGE_ROLE_CN_ABSTRACT: "中文摘要",
                PAGE_ROLE_EN_ABSTRACT: "英文摘要",
                PAGE_ROLE_CATALOGUE: "目录",
                PAGE_ROLE_OTHER: "第1章前置部分",
            }
            role_name = role_name_map.get(role, "第1章前置部分")
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.{page_no}",
                    title="页眉",
                    message=f"第{page_no}页属于{role_name}，但顶部已识别到规范页眉内容，页眉可能早于第1章启用。",
                    severity=Severity.WARNING,
                    source=Source.PDF,
                    page=page_no,
                    fixable=False,
                    metadata={
                        "section": "页眉",
                        "content": content,
                        "problem": "页眉可能早于第1章启用",
                    },
                )
            )

        return issues


class PageNumberStyleSequencePdfRule(BaseRule):
    rule_id = "page_number.style_sequence"
    display_name = "Page number style and sequence check (pdf)"
    spec_ref = "撰写规范（7）页眉和页码"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        roles = _build_page_roles(pages)
        role_items = {
            PAGE_ROLE_CN_ABSTRACT: [],
            PAGE_ROLE_EN_ABSTRACT: [],
            PAGE_ROLE_MAIN: [],
        }

        for page in pages:
            page_no = getattr(page, "page_no", None)
            role = roles.get(page_no, PAGE_ROLE_OTHER)
            if role not in role_items:
                continue
            best = _best_page_number_candidate(page)
            if best is None:
                continue
            role_items[role].append(
                {
                    "page_no": page_no,
                    "text": best["text"],
                    "kind": best["kind"],
                    "bbox": best["bbox"],
                }
            )

        issues: list[Issue] = []
        issues.extend(self._check_roman_section(role_items[PAGE_ROLE_CN_ABSTRACT], "中文摘要"))
        issues.extend(self._check_roman_section(role_items[PAGE_ROLE_EN_ABSTRACT], "英文摘要"))
        issues.extend(self._check_arabic_section(role_items[PAGE_ROLE_MAIN]))
        return issues

    def _check_roman_section(self, items: list[dict], section_name: str) -> list[Issue]:
        if not items:
            return []

        issues: list[Issue] = []
        for item in items:
            if item["kind"] != "roman":
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.{section_name}.style.{item['page_no']}",
                        title="页码",
                        message=f"第{item['page_no']}页{section_name}页码“{item['text']}”可能不是罗马数字。",
                        severity=Severity.INFO,
                        source=Source.PDF,
                        page=item["page_no"],
                        bbox=[int(x) for x in item["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "页码",
                            "content": item["text"],
                            "problem": f"{section_name}页码可能不是罗马数字",
                        },
                    )
                )
                return issues

        values = [(_roman_to_int(item["text"]), item) for item in items]
        values = [pair for pair in values if pair[0] is not None]
        if not values:
            return issues

        first_value, first_item = values[0]
        if first_value != 1:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.{section_name}.start.{first_item['page_no']}",
                    title="页码",
                    message=f"第{first_item['page_no']}页{section_name}页码起始值识别为“{first_item['text']}”，可能不是 I。",
                    severity=Severity.INFO,
                    source=Source.PDF,
                    page=first_item["page_no"],
                    bbox=[int(x) for x in first_item["bbox"]],
                    fixable=False,
                    metadata={
                        "section": "页码",
                        "content": first_item["text"],
                        "problem": f"{section_name}页码起始值可能不是 I",
                    },
                )
            )

        for idx in range(len(values) - 1):
            current_value, current_item = values[idx]
            next_value, next_item = values[idx + 1]
            if next_value != current_value + 1:
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.{section_name}.gap.{next_item['page_no']}",
                        title="页码",
                        message=f"第{current_item['page_no']}页到第{next_item['page_no']}页的{section_name}罗马页码可能不连续。",
                        severity=Severity.INFO,
                        source=Source.PDF,
                        page=next_item["page_no"],
                        bbox=[int(x) for x in next_item["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "页码",
                            "content": f"{current_item['text']} -> {next_item['text']}",
                            "problem": f"{section_name}罗马页码可能不连续",
                        },
                    )
                )
                break

        return issues

    def _check_arabic_section(self, items: list[dict]) -> list[Issue]:
        if not items:
            return []

        issues: list[Issue] = []
        for item in items:
            if item["kind"] != "arabic":
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.main.style.{item['page_no']}",
                        title="页码",
                        message=f"第{item['page_no']}页正文页码“{item['text']}”可能不是阿拉伯数字。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=item["page_no"],
                        bbox=[int(x) for x in item["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "页码",
                            "content": item["text"],
                            "problem": "正文页码可能不是阿拉伯数字",
                        },
                    )
                )
                return issues

        first_item = items[0]
        if first_item["text"] != "1":
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.main.start.{first_item['page_no']}",
                    title="页码",
                    message=f"第{first_item['page_no']}页正文页码起始值识别为“{first_item['text']}”，可能不是 1。",
                    severity=Severity.INFO,
                    source=Source.PDF,
                    page=first_item["page_no"],
                    bbox=[int(x) for x in first_item["bbox"]],
                    fixable=False,
                    metadata={
                        "section": "页码",
                        "content": first_item["text"],
                        "problem": "正文页码起始值可能不是 1",
                    },
                )
            )

        for idx in range(len(items) - 1):
            current_item = items[idx]
            next_item = items[idx + 1]
            if int(next_item["text"]) != int(current_item["text"]) + 1:
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.main.gap.{next_item['page_no']}",
                        title="页码",
                        message=f"第{current_item['page_no']}页到第{next_item['page_no']}页的正文页码可能不连续。",
                        severity=Severity.INFO,
                        source=Source.PDF,
                        page=next_item["page_no"],
                        bbox=[int(x) for x in next_item["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "页码",
                            "content": f"{current_item['text']} -> {next_item['text']}",
                            "problem": "正文页码可能不连续",
                        },
                    )
                )
                break

        return issues
