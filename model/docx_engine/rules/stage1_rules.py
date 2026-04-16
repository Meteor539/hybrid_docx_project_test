import re
from collections.abc import Iterable

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _paragraph_text(paragraph) -> str:
    return str(getattr(paragraph, "text", "") or "").strip()


def _flatten_paragraphs(items) -> list:
    if items is None:
        return []
    if isinstance(items, list):
        return [x for x in items if x is not None]
    return [items]


def _doc_paragraph_texts(ctx: RuleContext) -> list[str]:
    doc = ctx.docx_obj
    if doc is None:
        return []
    paragraphs = getattr(doc, "paragraphs", None) or []
    return [text for text in (_paragraph_text(p) for p in paragraphs) if text]


def _contains_keywords(texts: Iterable[str], keywords: Iterable[str], *, lower: bool = False) -> bool:
    keys = [k.lower() for k in keywords] if lower else list(keywords)
    for text in texts:
        candidate = text.lower() if lower else text
        if any(key in candidate for key in keys):
            return True
    return False


def _extract_keyword_count(paragraph) -> int:
    text = _paragraph_text(paragraph)
    if not text:
        return 0

    # 常见写法：关键词：A；B；C / Key Words: A; B; C
    text = re.sub(r"^\s*(关键词|关 键 词|key\s*words?)\s*[:：]\s*", "", text, flags=re.IGNORECASE)
    parts = [x.strip() for x in re.split(r"[;；,，、]+", text) if x.strip()]
    return len(parts)


def _strip_heading_prefix(text: str) -> str:
    if not text:
        return ""
    stripped = text.strip()
    patterns = [
        r"^第[一二三四五六七八九十百千0-9]+章\s*",
        r"^chapter\s+\d+\s+",
        r"^\d+\.\d+\.\d+\s*",
        r"^\d+\.\d+\s*",
        r"^\d+\.\s*",
        r"^\d+\)\s*",
        r"^[一二三四五六七八九十]+、\s*",
        r"^（[一二三四五六七八九十]+）\s*",
    ]
    for pattern in patterns:
        new_text = re.sub(pattern, "", stripped, flags=re.IGNORECASE)
        if new_text != stripped:
            return new_text.strip()
    return stripped


def _count_non_whitespace_chars(text: str) -> int:
    return len(re.sub(r"\s+", "", text or ""))


def _join_paragraph_texts(paragraphs) -> str:
    joined = "\n".join(_paragraph_text(p) for p in _flatten_paragraphs(paragraphs) if _paragraph_text(p))
    return joined.strip()


def _looks_like_foreign_reference(text: str) -> bool:
    if not text:
        return False
    ascii_letters = len(re.findall(r"[A-Za-z]", text))
    cjk_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
    return ascii_letters >= 8 and ascii_letters > cjk_chars


def _title_paragraphs(sections: dict) -> list:
    collected: list = []
    if not isinstance(sections, dict):
        return collected

    cover = sections.get("cover", {})
    if isinstance(cover, dict):
        collected.extend(_flatten_paragraphs(cover.get("title")))

    catalogue = sections.get("catalogue", {})
    if isinstance(catalogue, dict):
        collected.extend(_flatten_paragraphs(catalogue.get("title")))

    references = sections.get("references", {})
    if isinstance(references, dict):
        collected.extend(_flatten_paragraphs(references.get("title")))

    acknowledgments = sections.get("acknowledgments", {})
    if isinstance(acknowledgments, dict):
        collected.extend(_flatten_paragraphs(acknowledgments.get("title")))

    abstract = sections.get("abstract_keyword", {})
    if isinstance(abstract, dict):
        for key in ("chinese_title", "english_title", "chinese_keyword_title", "english_keyword_title"):
            collected.extend(_flatten_paragraphs(abstract.get(key)))

    statement = sections.get("statement", {})
    if isinstance(statement, dict):
        collected.extend(_flatten_paragraphs(statement.get("title")))

    headings = sections.get("headings", {})
    if isinstance(headings, dict):
        for key in ("chapter", "level1", "level2", "level3"):
            collected.extend(_flatten_paragraphs(headings.get(key)))

    return [p for p in collected if p is not None]


class FirstStageSectionPresenceRule(BaseRule):
    rule_id = "docx.stage1.section_presence"
    display_name = "First-stage key section presence"
    spec_ref = "撰写规范（2）（4）（5）（9）及资料整理要求（1）"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        doc_texts = _doc_paragraph_texts(ctx)

        missing: list[str] = []

        cover = sections.get("cover", {})
        if not isinstance(cover, dict) or cover.get("title") is None:
            missing.append("封面题目")

        statement = sections.get("statement", {})
        has_statement = (
            isinstance(statement, dict)
            and bool(statement.get("title"))
            or _contains_keywords(doc_texts, ["原创性声明", "使用授权"])
        )
        if not has_statement:
            missing.append("原创性声明")

        abstract = sections.get("abstract_keyword", {})
        if not isinstance(abstract, dict) or abstract.get("chinese_title") is None:
            missing.append("中文摘要")
        if not isinstance(abstract, dict) or abstract.get("english_title") is None:
            missing.append("英文摘要")

        catalogue = sections.get("catalogue", {})
        has_catalogue = (
            isinstance(catalogue, dict)
            and catalogue.get("title") is not None
            or _contains_keywords(doc_texts, ["目录"])
        )
        if not has_catalogue:
            missing.append("目录")

        main_text = sections.get("main_text", [])
        if not isinstance(main_text, list) or len(main_text) == 0:
            missing.append("正文")

        references = sections.get("references", {})
        has_references = (
            isinstance(references, dict)
            and references.get("title") is not None
            and isinstance(references.get("content"), list)
            and len(references.get("content")) > 0
        )
        if not has_references:
            missing.append("参考文献")

        if not missing:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="章节完整性",
                message=f"未稳定识别到以下内容：{', '.join(missing)}。",
                severity=Severity.WARNING,
                source=Source.DOCX,
                page=1,
                fixable=False,
                metadata={
                    "section": "整体结构",
                    "content": ", ".join(missing),
                    "problem": "可能缺失",
                    "missing_sections": missing,
                },
            )
        ]


class CoverTitleLengthRule(BaseRule):
    rule_id = "docx.stage1.cover_title_length"
    display_name = "Cover title length"
    spec_ref = "撰写规范（1）题目"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        cover = sections.get("cover", {})
        title_paragraph = cover.get("title") if isinstance(cover, dict) else None
        title_text = _paragraph_text(title_paragraph)
        if not title_text:
            return []

        # 以非空白字符计数，适配中文题目场景。
        char_count = len(re.sub(r"\s+", "", title_text))
        if char_count <= 25:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="封面",
                message=f"{title_text}；长度约为 {char_count} 个非空白字符。",
                severity=Severity.WARNING,
                source=Source.DOCX,
                page=1,
                fixable=False,
                metadata={
                    "section": "封面",
                    "content": title_text,
                    "problem": "题目可能过长",
                    "char_count": char_count,
                },
            )
        ]


class KeywordCountRule(BaseRule):
    rule_id = "docx.stage1.keyword_count"
    display_name = "Keyword count"
    spec_ref = "撰写规范（3）中、英文关键词"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        abstract = sections.get("abstract_keyword", {})
        if not isinstance(abstract, dict):
            return []

        issues: list[Issue] = []
        for lang, key in (("中文关键词", "chinese_keyword"), ("英文关键词", "english_keyword")):
            count = _extract_keyword_count(abstract.get(key))
            if count == 0:
                continue
            if 3 <= count <= 5:
                continue

            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.{key}",
                    title="摘要与关键词",
                    message=f"{lang}数量为 {count}。",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    page=1,
                    fixable=False,
                    metadata={
                        "section": "摘要与关键词",
                        "content": _paragraph_text(abstract.get(key)),
                        "problem": f"{lang}数量可能不符合规范",
                        "keyword_count": count,
                        "field": key,
                    },
                )
            )

        return issues


class ChineseAbstractLengthRule(BaseRule):
    rule_id = "docx.stage1.abstract_length"
    display_name = "Chinese abstract length"
    spec_ref = "撰写规范（2）中、英文摘要"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        abstract = sections.get("abstract_keyword", {})
        if not isinstance(abstract, dict):
            return []

        content_text = _join_paragraph_texts(abstract.get("chinese_content"))
        if not content_text:
            return []

        char_count = _count_non_whitespace_chars(content_text)
        if 200 <= char_count <= 400:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="摘要与关键词",
                message=f"中文摘要长度约为 {char_count} 个非空白字符。",
                severity=Severity.INFO,
                source=Source.DOCX,
                page=1,
                fixable=False,
                metadata={
                    "section": "摘要与关键词",
                    "content": content_text,
                    "problem": "中文摘要长度可能偏离“300字左右”",
                    "char_count": char_count,
                },
            )
        ]


class HeadingPunctuationRule(BaseRule):
    rule_id = "docx.stage1.heading_punctuation"
    display_name = "Heading punctuation"
    spec_ref = "撰写规范（9）正文"
    engine = "docx"

    _punctuation_pattern = re.compile(r"[，。！？；：、,.!?;:()\[\]【】《》<>“”\"'‘’·—\-]")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        headings = sections.get("headings", {})
        if not isinstance(headings, dict):
            return []

        issues: list[Issue] = []
        for level in ("chapter", "level1", "level2", "level3"):
            for paragraph in _flatten_paragraphs(headings.get(level)):
                text = _paragraph_text(paragraph)
                if not text:
                    continue
                title_body = _strip_heading_prefix(text)
                if not title_body:
                    continue
                if not self._punctuation_pattern.search(title_body):
                    continue
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.{level}",
                        title="标题",
                        message=text,
                        severity=Severity.WARNING,
                        source=Source.DOCX,
                        fixable=False,
                        metadata={
                            "section": "标题",
                            "content": text,
                            "problem": "可能包含标点符号",
                            "level": level,
                            "title_body": title_body,
                        },
                    )
                )

        return issues


class HeadingCitationRule(BaseRule):
    rule_id = "docx.stage1.heading_citation"
    display_name = "Heading citation marker"
    spec_ref = "撰写规范（10）引文标识"
    engine = "docx"

    _citation_pattern = re.compile(r"\[\s*\d+\s*\]")

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        headings = sections.get("headings", {})
        if not isinstance(headings, dict):
            return []

        issues: list[Issue] = []
        for level in ("chapter", "level1", "level2", "level3"):
            for paragraph in _flatten_paragraphs(headings.get(level)):
                text = _paragraph_text(paragraph)
                if not text:
                    continue
                if not self._citation_pattern.search(text):
                    continue
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.{level}",
                        title="标题",
                        message=text,
                        severity=Severity.WARNING,
                        source=Source.DOCX,
                        fixable=False,
                        metadata={
                            "section": "标题",
                            "content": text,
                            "problem": "可能包含引文标示",
                            "level": level,
                        },
                    )
                )

        return issues


class ReferenceTerminalPeriodRule(BaseRule):
    rule_id = "docx.stage1.reference_terminal_period"
    display_name = "Reference terminal period"
    spec_ref = "撰写规范（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        references = sections.get("references", {})
        content = references.get("content") if isinstance(references, dict) else None
        if not isinstance(content, list):
            return []

        issues: list[Issue] = []
        for index, paragraph in enumerate(content, start=1):
            text = _paragraph_text(paragraph)
            if not text:
                continue
            if text.endswith("."):
                continue
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.{index}",
                    title="参考文献",
                    message=text,
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={
                        "section": "参考文献",
                        "content": text,
                        "problem": "可能缺少句点",
                        "index": index,
                    },
                )
            )

        return issues


class ReferenceCountRule(BaseRule):
    rule_id = "docx.stage2.reference_count"
    display_name = "Reference count"
    spec_ref = "撰写规范（9）（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        sections = ctx.docx_sections or {}
        references = sections.get("references", {})
        content = references.get("content") if isinstance(references, dict) else None
        if not isinstance(content, list) or not content:
            return []

        entries = [_paragraph_text(p) for p in content if _paragraph_text(p)]
        total_count = len(entries)
        foreign_count = sum(1 for text in entries if _looks_like_foreign_reference(text))

        issues: list[Issue] = []
        if total_count < 10:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.total",
                    title="参考文献",
                    message=f"当前共识别到 {total_count} 条参考文献。",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={
                        "section": "参考文献",
                        "content": f"共 {total_count} 条",
                        "problem": "参考文献数量可能不足（软件类通常不少于10篇）",
                        "total_count": total_count,
                    },
                )
            )
        elif total_count < 15:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.total_info",
                    title="参考文献",
                    message=f"当前共识别到 {total_count} 条参考文献。",
                    severity=Severity.INFO,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={
                        "section": "参考文献",
                        "content": f"共 {total_count} 条",
                        "problem": "若按一般论文要求，参考文献数量可能不足15篇；软件类通常不少于10篇",
                        "total_count": total_count,
                    },
                )
            )

        if foreign_count < 3:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.foreign",
                    title="参考文献",
                    message=f"当前推测外文参考文献约为 {foreign_count} 条。",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={
                        "section": "参考文献",
                        "content": f"推测外文参考文献约 {foreign_count} 条",
                        "problem": "外文参考文献数量可能不足3篇",
                        "foreign_count": foreign_count,
                    },
                )
            )

        return issues
