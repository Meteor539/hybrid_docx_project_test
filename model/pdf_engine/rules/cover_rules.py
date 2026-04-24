import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _group_page_lines(page):
    rows = []
    for span in getattr(page, "spans", []):
        text = (getattr(span, "text", "") or "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4:
            continue

        x0, y0, x1, y1 = [float(value) for value in bbox]
        center_y = (y0 + y1) / 2
        target_row = None
        for row in rows:
            if abs(center_y - row["center_y"]) <= 5.0:
                target_row = row
                break
        if target_row is None:
            target_row = {"center_y": center_y, "items": []}
            rows.append(target_row)
        target_row["items"].append({"text": text, "bbox": [x0, y0, x1, y1]})

    merged = []
    for row in sorted(rows, key=lambda item: item["center_y"]):
        items = sorted(row["items"], key=lambda item: item["bbox"][0])
        text = "".join(item["text"] for item in items).strip()
        if not text:
            continue
        bbox = [
            min(item["bbox"][0] for item in items),
            min(item["bbox"][1] for item in items),
            max(item["bbox"][2] for item in items),
            max(item["bbox"][3] for item in items),
        ]
        merged.append({"text": text, "bbox": bbox})
    return merged


def _center_offset_ratio(box, page_width: float) -> float:
    if page_width <= 0:
        return 1.0
    center_x = (box[0] + box[2]) / 2
    return abs(center_x - (page_width / 2)) / page_width


def _school_line_candidate(lines):
    for line in lines:
        normalized = _normalize_text(line["text"])
        if "武汉理工大学毕业设计" in normalized or "武汉理工大学毕业论文" in normalized:
            return line
    return None


def _title_line_candidates(lines):
    info_keywords = ("院", "系", "专业", "班级", "学生姓名", "指导教师", "学号")
    candidates = []
    for line in lines:
        text = line["text"].strip()
        normalized = _normalize_text(text)
        if not text or len(normalized) < 4:
            continue
        if any(keyword in text for keyword in info_keywords):
            continue
        if "武汉理工大学" in normalized:
            continue
        if re.search(r"[A-Za-z]{5,}", text):
            continue
        if len(normalized) > 40:
            continue
        candidates.append(line)
    return candidates


class CoverTitleCenterPdfRule(BaseRule):
    rule_id = "cover.title_center"
    display_name = "Cover title centered check (pdf)"
    spec_ref = "撰写规范（4）论文封面"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        first_page = pages[0]
        page_width = float(getattr(first_page, "width", 0.0) or 0.0)
        page_height = float(getattr(first_page, "height", 0.0) or 0.0)
        if page_width <= 0 or page_height <= 0:
            return []

        lines = [
            line
            for line in _group_page_lines(first_page)
            if line["bbox"][1] <= page_height * 0.7
        ]
        if not lines:
            return []

        issues: list[Issue] = []

        school_line = _school_line_candidate(lines)
        if school_line is not None:
            offset_ratio = _center_offset_ratio(school_line["bbox"], page_width)
            if offset_ratio > 0.08:
                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.school",
                        title="封面学校名称",
                        message=f"第1页封面学校名称“{school_line['text']}”可能未视觉居中。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=1,
                        bbox=[int(value) for value in school_line["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "封面学校名称",
                            "content": school_line["text"],
                            "problem": "封面学校名称可能未视觉居中",
                            "center_offset_ratio": offset_ratio,
                        },
                    )
                )

        title_candidates = _title_line_candidates(lines)
        if not title_candidates:
            return issues

        title_line = max(
            title_candidates,
            key=lambda item: ((item["bbox"][3] - item["bbox"][1]), len(_normalize_text(item["text"]))),
        )
        offset_ratio = _center_offset_ratio(title_line["bbox"], page_width)
        if offset_ratio > 0.08:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.paper_title",
                    title="封面题目",
                    message=f"第1页封面题目“{title_line['text']}”可能未视觉居中。",
                    severity=Severity.WARNING,
                    source=Source.PDF,
                    page=1,
                    bbox=[int(value) for value in title_line["bbox"]],
                    fixable=False,
                    metadata={
                        "section": "封面题目",
                        "content": title_line["text"],
                        "problem": "封面题目可能未视觉居中",
                        "center_offset_ratio": offset_ratio,
                    },
                )
            )

        return issues
