import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source
from model.pdf_engine.page_roles import is_catalogue_page, normalize_text, page_has_top_heading


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

    compact = normalize_text(original)
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
    lines = list(_iter_text_lines(page))
    for idx, line in enumerate(lines):
        if _is_chapter_heading_text(line["text"]):
            yield line
            continue

        compact = normalize_text(line["text"])
        if not re.fullmatch(r"第[一二三四五六七八九十百千0123456789]+章", compact):
            continue

        if idx + 1 >= len(lines):
            continue

        next_line = lines[idx + 1]
        next_compact = normalize_text(next_line["text"])
        if not next_compact or len(next_compact) > 30:
            continue
        if re.search(r"[，。；：！？!?]", next_compact):
            continue

        current_box = line["bbox"]
        next_box = next_line["bbox"]
        gap = float(next_box[1]) - float(current_box[3])
        page_height = float(getattr(page, "height", 0.0) or 0.0)
        max_gap = max(18.0, page_height * 0.025) if page_height > 0 else 18.0
        if gap < -2.0 or gap > max_gap:
            continue

        merged = {
            "text": f"{compact}{next_compact}",
            "bbox": [
                min(current_box[0], next_box[0]),
                min(current_box[1], next_box[1]),
                max(current_box[2], next_box[2]),
                max(current_box[3], next_box[3]),
            ],
        }
        if _is_chapter_heading_text(merged["text"]):
            yield merged


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
            if page_has_top_heading(page, ("目录", "目次")) or is_catalogue_page(page):
                if not list(_chapter_heading_lines(page)):
                    continue
            if page_has_top_heading(page, ("目录", "目次")):
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
