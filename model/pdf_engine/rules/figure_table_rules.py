import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _caption_candidates(page, *, kind: str):
    prefix = "图" if kind == "figure" else "表"
    pattern = re.compile(rf"^{prefix}\s*[A-Z]?\d+(?:[\.．\-－]\d+)*(?:\s+.+)?$")

    for span in getattr(page, "spans", []):
        text = (getattr(span, "text", "") or "").strip()
        bbox = getattr(span, "bbox", None) or []
        if not text or len(bbox) != 4 or len(text) > 80:
            continue
        if not pattern.match(text):
            continue
        yield {"text": text, "bbox": [float(x) for x in bbox]}


def _regions(page, *, kind: str):
    for region in getattr(page, "regions", []):
        if getattr(region, "kind", "") != kind:
            continue
        bbox = getattr(region, "bbox", None) or []
        if len(bbox) != 4:
            continue
        yield [float(x) for x in bbox]


def _horizontal_overlap_ratio(box1, box2) -> float:
    left = max(box1[0], box2[0])
    right = min(box1[2], box2[2])
    overlap = max(0.0, right - left)
    base = min(box1[2] - box1[0], box2[2] - box2[0])
    if base <= 0:
        return 0.0
    return overlap / base


def _center_offset_ratio(box, page_width: float) -> float:
    if page_width <= 0:
        return 1.0
    center_x = (box[0] + box[2]) / 2
    return abs(center_x - (page_width / 2)) / page_width


def _vertical_gap_to_region(caption_box, region_box, *, caption_expected: str) -> float:
    if caption_expected == "below":
        return caption_box[1] - region_box[3]
    return region_box[1] - caption_box[3]


def _is_split_pair(caption_box, region_box, *, caption_expected: str, page_height: float) -> bool:
    if page_height <= 0:
        return False

    overlap = _horizontal_overlap_ratio(caption_box, region_box)
    if overlap < 0.3:
        return False

    if caption_expected == "below":
        caption_near_top = caption_box[1] <= page_height * 0.2
        region_near_bottom = region_box[3] >= page_height * 0.8
        return caption_near_top and region_near_bottom

    caption_near_bottom = caption_box[3] >= page_height * 0.8
    region_near_top = region_box[1] <= page_height * 0.2
    return caption_near_bottom and region_near_top


def _find_nearby_region(page, caption_box, *, kind: str, caption_expected: str):
    for region_box in _regions(page, kind=kind):
        overlap = _horizontal_overlap_ratio(caption_box, region_box)
        gap = _vertical_gap_to_region(caption_box, region_box, caption_expected=caption_expected)
        if overlap >= 0.3 and 0 <= gap <= 80:
            return region_box
    return None


def _find_split_region(pages, index: int, caption_box, *, kind: str, caption_expected: str):
    for neighbor_index in (index - 1, index + 1):
        if neighbor_index < 0 or neighbor_index >= len(pages):
            continue
        neighbor_page = pages[neighbor_index]
        page_height = float(getattr(neighbor_page, "height", 0.0) or 0.0)
        for region_box in _regions(neighbor_page, kind=kind):
            if _is_split_pair(caption_box, region_box, caption_expected=caption_expected, page_height=page_height):
                return neighbor_page, region_box
    return None, None


def _is_probable_caption_hint(page, span_text: str, bbox, *, kind: str) -> bool:
    text = (span_text or "").strip()
    if not text or len(text) > 60:
        return False

    normalized = _normalize_text(text)
    prefix = "图" if kind == "figure" else "表"
    if not normalized.startswith(prefix):
        return False

    if not re.match(rf"^{prefix}[A-Z]?\d", normalized):
        return False

    if normalized.startswith(("图示", "图像", "图中", "表明", "表示", "表达", "表征", "表面")):
        return False

    region_kind = "image" if kind == "figure" else "table"
    near_region = False
    for region_box in _regions(page, kind=region_kind):
        overlap = _horizontal_overlap_ratio(bbox, region_box)
        vertical_gap = min(abs(bbox[1] - region_box[3]), abs(region_box[1] - bbox[3]))
        if overlap >= 0.2 and vertical_gap <= 120:
            near_region = True
            break

    if near_region:
        return True

    page_width = float(getattr(page, "width", 0.0) or 0.0)
    return _center_offset_ratio(bbox, page_width) <= 0.18


class FigureTableCaptionHintRule(BaseRule):
    rule_id = "figure_table.caption_hint"
    display_name = "Figure/table caption format hint (pdf)"
    spec_ref = "撰写规范（16）（17）表题与图题"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        figure_pattern = re.compile(r"^图\s*[A-Z]?\d+(?:[\.．\-－]\d+)*\s+.+$")
        table_pattern = re.compile(r"^表\s*[A-Z]?\d+(?:[\.．\-－]\d+)*\s+.+$")

        issues: list[Issue] = []
        for page in pages:
            page_no = getattr(page, "page_no", None)
            for span in getattr(page, "spans", []):
                text = (getattr(span, "text", "") or "").strip()
                bbox = getattr(span, "bbox", None) or []
                if len(bbox) != 4:
                    continue

                if _is_probable_caption_hint(page, text, bbox, kind="figure") and not figure_pattern.match(text):
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="图题编号格式可疑",
                            message=f"第{page_no}页存在疑似图题“{text}”，但未匹配常见图题编号格式。",
                            severity=Severity.INFO,
                            source=Source.PDF,
                            page=page_no,
                            bbox=[int(x) for x in bbox],
                            fixable=False,
                            metadata={
                                "section": "图题",
                                "content": text,
                                "problem": "疑似图题未匹配常见编号与标题格式",
                            },
                        )
                    )

                if _is_probable_caption_hint(page, text, bbox, kind="table") and not table_pattern.match(text):
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="表题编号格式可疑",
                            message=f"第{page_no}页存在疑似表题“{text}”，但未匹配常见表题编号格式。",
                            severity=Severity.INFO,
                            source=Source.PDF,
                            page=page_no,
                            bbox=[int(x) for x in bbox],
                            fixable=False,
                            metadata={
                                "section": "表题",
                                "content": text,
                                "problem": "疑似表题未匹配常见编号与标题格式",
                            },
                        )
                    )

        return issues


class FigureCaptionBelowPdfRule(BaseRule):
    rule_id = "figure.caption_below"
    display_name = "Figure caption below check (pdf)"
    spec_ref = "撰写规范（17）图"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        for page_index, page in enumerate(pages):
            image_boxes = list(_regions(page, kind="image"))
            if not image_boxes:
                continue

            for caption in _caption_candidates(page, kind="figure"):
                caption_box = caption["bbox"]
                if _find_nearby_region(page, caption_box, kind="image", caption_expected="below") is not None:
                    continue

                split_page, _ = _find_split_region(
                    pages,
                    page_index,
                    caption_box,
                    kind="image",
                    caption_expected="below",
                )
                if split_page is not None:
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="图题",
                        message=f"第{getattr(page, 'page_no', '?')}页图题“{caption['text']}”上方未识别到紧邻插图。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in caption_box],
                        fixable=False,
                        metadata={
                            "section": "图题",
                            "content": caption["text"],
                            "problem": "图题上方未识别到对应插图，可能不在图下方",
                        },
                    )
                )

        return issues


class FigureCaptionCenterPdfRule(BaseRule):
    rule_id = "figure.caption_center"
    display_name = "Figure caption centered check (pdf)"
    spec_ref = "撰写规范（17）图"
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

            for caption in _caption_candidates(page, kind="figure"):
                offset_ratio = _center_offset_ratio(caption["bbox"], page_width)
                if offset_ratio <= 0.12:
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="图题",
                        message=f"第{getattr(page, 'page_no', '?')}页图题“{caption['text']}”可能未居中。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in caption["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "图题",
                            "content": caption["text"],
                            "problem": "图题可能未居中",
                            "center_offset_ratio": offset_ratio,
                        },
                    )
                )

        return issues


class TableCaptionAbovePdfRule(BaseRule):
    rule_id = "table.caption_above"
    display_name = "Table caption above check (pdf)"
    spec_ref = "撰写规范（16）表格"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        for page_index, page in enumerate(pages):
            table_boxes = list(_regions(page, kind="table"))
            if not table_boxes:
                continue

            for caption in _caption_candidates(page, kind="table"):
                caption_box = caption["bbox"]
                if _find_nearby_region(page, caption_box, kind="table", caption_expected="above") is not None:
                    continue

                split_page, _ = _find_split_region(
                    pages,
                    page_index,
                    caption_box,
                    kind="table",
                    caption_expected="above",
                )
                if split_page is not None:
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="表题",
                        message=f"第{getattr(page, 'page_no', '?')}页表题“{caption['text']}”下方未识别到紧邻表格。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in caption_box],
                        fixable=False,
                        metadata={
                            "section": "表题",
                            "content": caption["text"],
                            "problem": "表题下方未识别到对应表格，可能不在表上方",
                        },
                    )
                )

        return issues


class FigureTableSplitAcrossPagesPdfRule(BaseRule):
    rule_id = "figure_table.split_across_pages"
    display_name = "Figure/table split across pages check (pdf)"
    spec_ref = "撰写规范（16）（17）表格与图"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        issues: list[Issue] = []
        seen_keys: set[tuple[str, int, str]] = set()

        for page_index, page in enumerate(pages):
            page_no = getattr(page, "page_no", None)

            for caption in _caption_candidates(page, kind="figure"):
                caption_box = caption["bbox"]
                if _find_nearby_region(page, caption_box, kind="image", caption_expected="below") is not None:
                    continue

                split_page, _ = _find_split_region(
                    pages,
                    page_index,
                    caption_box,
                    kind="image",
                    caption_expected="below",
                )
                if split_page is None:
                    continue

                key = ("figure", page_no or -1, caption["text"])
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                split_page_no = getattr(split_page, "page_no", None)
                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="图题",
                        message=f"第{page_no}页图题“{caption['text']}”与对应插图疑似跨页分离（关联页：第{split_page_no}页）。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=page_no,
                        bbox=[int(x) for x in caption_box],
                        fixable=False,
                        metadata={
                            "section": "图题",
                            "content": caption["text"],
                            "problem": "图与图题疑似被拆开排版为两页",
                            "related_page": split_page_no,
                        },
                    )
                )

            for caption in _caption_candidates(page, kind="table"):
                caption_box = caption["bbox"]
                if _find_nearby_region(page, caption_box, kind="table", caption_expected="above") is not None:
                    continue

                split_page, _ = _find_split_region(
                    pages,
                    page_index,
                    caption_box,
                    kind="table",
                    caption_expected="above",
                )
                if split_page is None:
                    continue

                key = ("table", page_no or -1, caption["text"])
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                split_page_no = getattr(split_page, "page_no", None)
                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="表题",
                        message=f"第{page_no}页表题“{caption['text']}”与对应表格疑似跨页分离（关联页：第{split_page_no}页）。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=page_no,
                        bbox=[int(x) for x in caption_box],
                        fixable=False,
                        metadata={
                            "section": "表题",
                            "content": caption["text"],
                            "problem": "表与表题疑似被拆开排版为两页",
                            "related_page": split_page_no,
                        },
                    )
                )

        return issues


class TableCaptionCenterPdfRule(BaseRule):
    rule_id = "table.caption_center"
    display_name = "Table caption centered check (pdf)"
    spec_ref = "撰写规范（16）表格"
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

            for caption in _caption_candidates(page, kind="table"):
                offset_ratio = _center_offset_ratio(caption["bbox"], page_width)
                if offset_ratio <= 0.12:
                    continue

                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="表题",
                        message=f"第{getattr(page, 'page_no', '?')}页表题“{caption['text']}”可能未居中。",
                        severity=Severity.WARNING,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        bbox=[int(x) for x in caption["bbox"]],
                        fixable=False,
                        metadata={
                            "section": "表题",
                            "content": caption["text"],
                            "problem": "表题可能未居中",
                            "center_offset_ratio": offset_ratio,
                        },
                    )
                )

        return issues
