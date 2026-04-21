import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _is_catalogue_page(page) -> bool:
    normalized = _normalize_text(getattr(page, "text", ""))
    if "目录" in normalized or "目錄" in normalized:
        return True
    dotted_entries = len(re.findall(r"\d+\.\d+|\d+$", getattr(page, "text", ""), flags=re.MULTILINE))
    return dotted_entries >= 8 and "第1章" in normalized


def _iter_text_lines(page):
    rows: list[dict] = []
    spans = sorted(
        [span for span in getattr(page, "spans", []) if getattr(span, "text", "").strip()],
        key=lambda item: ((item.bbox[1] + item.bbox[3]) / 2, item.bbox[0]),
    )

    for span in spans:
        bbox = getattr(span, "bbox", None) or []
        if len(bbox) != 4:
            continue

        x0, y0, x1, y1 = [float(x) for x in bbox]
        center_y = (y0 + y1) / 2
        target_row = None
        for row in rows:
            if abs(center_y - row["center_y"]) <= 4.0:
                target_row = row
                break

        if target_row is None:
            target_row = {"center_y": center_y, "items": []}
            rows.append(target_row)

        target_row["items"].append({"text": span.text.strip(), "bbox": [x0, y0, x1, y1]})

    for row in rows:
        items = sorted(row["items"], key=lambda item: item["bbox"][0])
        text = "".join(item["text"] for item in items).strip()
        if not text:
            continue
        x0 = min(item["bbox"][0] for item in items)
        y0 = min(item["bbox"][1] for item in items)
        x1 = max(item["bbox"][2] for item in items)
        y1 = max(item["bbox"][3] for item in items)
        yield {"text": text, "bbox": [x0, y0, x1, y1]}


def _is_chapter_heading_text(text: str) -> bool:
    original = (text or "").strip()
    if not original:
        return False
    if re.search(r"[，。；：！？!?]", original):
        return False
    if len(original) > 40:
        return False

    compact = _normalize_text(original)
    if compact.lower().startswith("chapter"):
        return bool(re.fullmatch(r"chapter\d+.{1,30}", compact, flags=re.IGNORECASE))

    if not compact.startswith("第") or "章" not in compact:
        return False

    chapter_mark_index = compact.find("章")
    if chapter_mark_index <= 1:
        return False

    chapter_no = compact[1:chapter_mark_index]
    title = compact[chapter_mark_index + 1 :]
    if not title or len(title) > 30:
        return False

    valid_chapter_chars = set("一二三四五六七八九十百千0123456789")
    if any(char not in valid_chapter_chars for char in chapter_no):
        return False

    if re.search(r"[，。；：！？!?]", title):
        return False

    return True


def _chapter_heading_lines(page):
    for line in _iter_text_lines(page):
        if _is_chapter_heading_text(line["text"]):
            yield line


def _has_substantial_text_before_heading(page, heading_top: float, page_height: float) -> bool:
    header_limit = page_height * 0.14
    for line in _iter_text_lines(page):
        bbox = line["bbox"]
        text = line["text"].strip()
        if not text:
            continue
        if bbox[3] <= header_limit:
            continue
        if bbox[3] >= heading_top:
            continue
        if re.fullmatch(r"(?:\d+|[IVXLCDMivxlcdm]+)", text):
            continue
        return True
    return False


class ChapterStartsNewPagePdfRule(BaseRule):
    rule_id = "chapter.starts_new_page"
    display_name = "Chapter starts on a new page check (pdf)"
    spec_ref = "撰写规范（9）正文"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        for page in pages:
            if _is_catalogue_page(page):
                continue

            page_no = getattr(page, "page_no", None)
            page_height = float(getattr(page, "height", 0.0) or 0.0)
            if page_height <= 0:
                continue

            heading_lines = list(_chapter_heading_lines(page))
            if not heading_lines:
                continue

            if len(heading_lines) >= 2:
                for heading in heading_lines[1:]:
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="章节标题",
                            message=f"第{page_no}页出现多个章标题，其中“{heading['text']}”可能未另起一页。",
                            severity=Severity.WARNING,
                            source=Source.PDF,
                            page=page_no,
                            bbox=[int(x) for x in heading["bbox"]],
                            fixable=False,
                            metadata={
                                "section": "章节标题",
                                "content": heading["text"],
                                "problem": "同一页出现多个章标题，后续章标题可能未另起一页",
                            },
                        )
                    )

            for heading in heading_lines:
                heading_top = heading["bbox"][1]
                if heading_top <= page_height * 0.28 and not _has_substantial_text_before_heading(page, heading_top, page_height):
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="章节标题",
                        message=f"第{page_no}页章标题“{heading['text']}”未出现在新页页首区域，可能未另起一页。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=page_no,
                        bbox=[int(x) for x in heading["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "章节标题",
                            "content": heading["text"],
                            "problem": "章标题未出现在新页页首区域，可能未另起一页",
                        },
                    )
                )

        return issues
