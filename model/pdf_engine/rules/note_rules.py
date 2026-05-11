import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


_NOTE_MARKERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_NOTE_MARKER_SET = set(_NOTE_MARKERS)


def _is_note_marker_span(text: str) -> bool:
    normalized = re.sub(r"[\s\u3000]+", "", text or "")
    return normalized in _NOTE_MARKER_SET


def _same_line_left_span(page, marker_box):
    target_center_y = (marker_box[1] + marker_box[3]) / 2
    candidates = []
    for span in getattr(page, "spans", []):
        text = (getattr(span, "text", "") or "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4:
            continue
        if _is_note_marker_span(text):
            continue
        box = [float(x) for x in bbox]
        center_y = (box[1] + box[3]) / 2
        if abs(center_y - target_center_y) > 6.0:
            continue
        if box[2] > marker_box[0] + 1:
            continue
        candidates.append({"text": text, "bbox": box})
    if not candidates:
        return None
    return max(candidates, key=lambda item: item["bbox"][2])


def _looks_like_note_entry_line(left_span, marker_box) -> bool:
    if left_span is None:
        return False
    text = str(left_span["text"] or "").strip()
    if not text:
        return False
    if left_span["bbox"][0] > 80:
        return False
    gap = marker_box[0] - left_span["bbox"][2]
    return gap <= 8 and len(text) <= 4


class NoteMarkerPositionPdfRule(BaseRule):
    rule_id = "note.marker_position"
    display_name = "Note marker position check (pdf)"
    spec_ref = "撰写规范（18）注释"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        for page in pages:
            for span_index, span in enumerate(getattr(page, "spans", []) or [], start=1):
                text = (getattr(span, "text", "") or "").strip()
                bbox = getattr(span, "bbox", None) or []
                if not _is_note_marker_span(text) or len(bbox) != 4:
                    continue

                marker_box = [float(x) for x in bbox]
                left_span = _same_line_left_span(page, marker_box)
                if left_span is None or _looks_like_note_entry_line(left_span, marker_box):
                    continue

                left_box = left_span["bbox"]
                horizontal_gap = marker_box[0] - left_box[2]
                raised = marker_box[3] <= left_box[3] - 1.5
                near_right = -1.0 <= horizontal_gap <= 18.0
                if raised and near_right:
                    continue

                issues.append(
                    Issue(
                        rule_id=f"{self.rule_id}.{getattr(page, 'page_no', 0)}.{span_index}",
                        title="注释",
                        message=f"第{getattr(page, 'page_no', '?')}页注释标记“{text}”附近正文为“{left_span['text']}”。",
                        severity=Severity.INFO,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in marker_box],
                        fixable=False,
                        metadata={
                            "section": "正文内容",
                            "content": text,
                            "problem": "注释标记可能未位于被注释词条右上角",
                            "context_text": left_span["text"],
                            "horizontal_gap": round(horizontal_gap, 2),
                        },
                    )
                )

        return issues
