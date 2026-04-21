import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _page_has_heading(page, keywords: tuple[str, ...]) -> bool:
    normalized = _normalize_text(getattr(page, "text", ""))
    return any(keyword in normalized for keyword in keywords)


def _looks_like_catalogue_page(page) -> bool:
    normalized = _normalize_text(getattr(page, "text", ""))
    if "目录" in normalized or "目錄" in normalized:
        return True

    raw_text = getattr(page, "text", "") or ""
    dotted_entries = len(re.findall(r"[.．·•…\-_ ]{2,}\s*[IVXLCDMivxlcdm\d]+\s*$", raw_text, flags=re.MULTILINE))
    numbered_entries = len(re.findall(r"^\s*(?:第\s*\d+\s*章|\d+\.\d+(?:\.\d+)?)", raw_text, flags=re.MULTILINE))
    return dotted_entries >= 5 and numbered_entries >= 3


def _find_first_main_index(pages) -> int | None:
    for idx, page in enumerate(pages):
        if _looks_like_catalogue_page(page):
            continue
        if _page_has_heading(page, ("第1章", "第一章", "绪论", "緒論")):
            return idx
    return None


def _collect_catalogue_pages(pages):
    if not pages:
        return []

    title_index = None
    for idx, page in enumerate(pages):
        if "目录" in _normalize_text(getattr(page, "text", "")):
            title_index = idx
            break

    if title_index is None:
        return []

    first_main_index = _find_first_main_index(pages)
    end_index = first_main_index if first_main_index is not None else len(pages)
    collected = []
    for idx in range(title_index, end_index):
        page = pages[idx]
        if idx == title_index or _looks_like_catalogue_page(page):
            collected.append(page)
            continue
        break
    return collected


def _body_has_level3_headings(docx_sections) -> bool:
    sections = docx_sections or {}
    if not isinstance(sections, dict):
        return False

    headings = sections.get("headings", {})
    if not isinstance(headings, dict):
        return False

    for paragraph in headings.get("level2", []) or []:
        text = str(getattr(paragraph, "text", "") or "").strip()
        if re.match(r"^\d+\.\d+\.\d+", text):
            return True
    return False


class TocPresencePdfRule(BaseRule):
    rule_id = "toc.exists_and_match"
    display_name = "TOC presence (pdf)"
    spec_ref = "撰写规范（8）目录"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []
        all_text = "\n".join(getattr(p, "text", "") for p in pages)
        if "目录" in all_text:
            return []
        return [
            Issue(
                rule_id=self.rule_id,
                title="目录标题未识别",
                message="PDF 文本中未识别到“目录”关键词，建议人工确认目录页。",
                severity=Severity.INFO,
                source=Source.PDF,
                fixable=False,
            )
        ]


class TocLevelPresentationPdfRule(BaseRule):
    rule_id = "toc.level_presentation"
    display_name = "TOC third-level presentation (pdf)"
    spec_ref = "撰写规范（4）目录、（8）目录"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        if not _body_has_level3_headings(ctx.docx_sections):
            return []

        catalogue_pages = _collect_catalogue_pages(pages)
        if not catalogue_pages:
            return []

        catalogue_text = "\n".join(getattr(page, "text", "") for page in catalogue_pages)
        normalized_catalogue_text = re.sub(r"[\u3000\t ]+", "", catalogue_text)
        if re.search(r"^\s*\d+\.\d+\.\d+", normalized_catalogue_text, flags=re.MULTILINE):
            return []

        page_nos = [str(getattr(page, "page_no", "")) for page in catalogue_pages if getattr(page, "page_no", None) is not None]
        page_desc = "、".join(page_nos) if page_nos else "目录页"
        return [
            Issue(
                rule_id=self.rule_id,
                title="目录",
                message=f"目录页（第{page_desc}页）中未明显识别到三级标题条目。",
                severity=Severity.INFO,
                source=Source.PDF,
                fixable=False,
                metadata={
                    "section": "目录",
                    "content": f"第{page_desc}页",
                    "problem": "正文存在三级标题，但目录页中未明显识别到三级标题条目",
                },
            )
        ]
