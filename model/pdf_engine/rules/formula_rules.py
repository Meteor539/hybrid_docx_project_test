import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _formula_number_candidates(page):
    pattern = re.compile(r"^[（(][A-Z]?\d+(?:\.\d+)?[)）]$")
    for span in getattr(page, "spans", []):
        text = (getattr(span, "text", "") or "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4:
            continue
        if not pattern.fullmatch(text):
            continue
        yield {"text": text, "bbox": [float(x) for x in bbox]}


def _same_line_spans(page, target_box):
    target_center_y = (target_box[1] + target_box[3]) / 2
    same_line = []
    for span in getattr(page, "spans", []):
        text = (getattr(span, "text", "") or "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4:
            continue
        box = [float(x) for x in bbox]
        center_y = (box[1] + box[3]) / 2
        if abs(center_y - target_center_y) <= 4.0:
            same_line.append({"text": text, "bbox": box})
    return sorted(same_line, key=lambda item: item["bbox"][0])


def _is_formula_like_line(text: str) -> bool:
    stripped = re.sub(r"[\s\u3000]+", "", text or "")
    if not stripped:
        return False
    if any(symbol in stripped for symbol in ("=", "＝", "+", "-", "−", "×", "÷", "/", "*", "^", "∑", "∫", "√", "≤", "≥")):
        return True

    ascii_math = len(re.findall(r"[A-Za-z0-9]", stripped))
    cjk_chars = len(re.findall(r"[\u4e00-\u9fff]", stripped))
    return ascii_math >= 5 and ascii_math > cjk_chars


class FormulaNumberRightAlignedPdfRule(BaseRule):
    rule_id = "formula.number_right_aligned"
    display_name = "Formula number right-end alignment check (pdf)"
    spec_ref = "撰写规范（15）公式"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        for page in pages:
            page_width = float(getattr(page, "width", 0.0) or 0.0)
            if page_width <= 0:
                continue

            for candidate in _formula_number_candidates(page):
                line_spans = _same_line_spans(page, candidate["bbox"])
                left_spans = [item for item in line_spans if item["bbox"][2] <= candidate["bbox"][0] + 1]
                if not left_spans:
                    continue

                line_text = " ".join(item["text"] for item in left_spans)
                if not _is_formula_like_line(line_text):
                    continue

                # 右侧行末：编号右边界应接近页面右侧区域，且应为本行最右文本。
                if any(item["bbox"][0] > candidate["bbox"][2] + 2 for item in line_spans):
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="公式编号",
                            message=f"第{getattr(page, 'page_no', '?')}页公式编号“{candidate['text']}”右侧仍存在文本，可能未位于行末。",
                            severity=Severity.WARNING,
                            source=Source.PDF,
                            page=getattr(page, "page_no", None),
                            bbox=[int(x) for x in candidate["bbox"]],
                            fixable=False,
                            metadata={
                                "section": "公式编号",
                                "content": candidate["text"],
                                "problem": "公式编号右侧仍存在文本，可能未位于右侧行末",
                            },
                        )
                    )
                    continue

                if candidate["bbox"][2] >= page_width * 0.82:
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="公式编号",
                        message=f"第{getattr(page, 'page_no', '?')}页公式编号“{candidate['text']}”可能未位于右侧行末。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in candidate["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "公式编号",
                            "content": candidate["text"],
                            "problem": "公式编号可能未位于右侧行末",
                        },
                    )
                )

        return issues
