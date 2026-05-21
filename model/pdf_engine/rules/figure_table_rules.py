import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def _page_text_lines(page) -> list[dict[str, object]]:
    raw_spans: list[tuple[str, list[float]]] = []
    for span in getattr(page, "spans", []):
        text = str(getattr(span, "text", "") or "").strip()
        bbox = list(getattr(span, "bbox", None) or [])
        if not text or len(bbox) != 4:
            continue
        raw_spans.append((text, [float(x) for x in bbox]))

    if not raw_spans:
        return []

    raw_spans.sort(key=lambda item: (round((item[1][1] + item[1][3]) / 2, 1), item[1][0]))
    groups: list[list[tuple[str, list[float]]]] = []
    for text, bbox in raw_spans:
        cy = (bbox[1] + bbox[3]) / 2
        placed = False
        for group in groups:
            sample_bbox = group[0][1]
            sample_cy = (sample_bbox[1] + sample_bbox[3]) / 2
            if abs(cy - sample_cy) <= 3.0:
                group.append((text, bbox))
                placed = True
                break
        if not placed:
            groups.append([(text, bbox)])

    lines: list[dict[str, object]] = []
    for group in groups:
        group.sort(key=lambda item: item[1][0])
        parts: list[str] = []
        merged_bbox = [group[0][1][0], group[0][1][1], group[0][1][2], group[0][1][3]]
        prev_bbox: list[float] | None = None
        for text, bbox in group:
            if prev_bbox is not None:
                gap = bbox[0] - prev_bbox[2]
                if gap > 2:
                    parts.append(" ")
            parts.append(text)
            merged_bbox[0] = min(merged_bbox[0], bbox[0])
            merged_bbox[1] = min(merged_bbox[1], bbox[1])
            merged_bbox[2] = max(merged_bbox[2], bbox[2])
            merged_bbox[3] = max(merged_bbox[3], bbox[3])
            prev_bbox = bbox
        merged_text = "".join(parts).strip()
        if merged_text:
            lines.append({"text": merged_text, "bbox": merged_bbox})

    lines.sort(key=lambda item: (item["bbox"][1], item["bbox"][0]))  # type: ignore[index]
    return lines


def _page_span_rows(page) -> list[dict[str, object]]:
    raw_spans: list[tuple[str, list[float]]] = []
    for span in getattr(page, "spans", []):
        text = str(getattr(span, "text", "") or "").strip()
        bbox = list(getattr(span, "bbox", None) or [])
        if not text or len(bbox) != 4:
            continue
        raw_spans.append((text, [float(x) for x in bbox]))

    if not raw_spans:
        return []

    raw_spans.sort(key=lambda item: (round((item[1][1] + item[1][3]) / 2, 1), item[1][0]))
    groups: list[dict[str, object]] = []
    for text, bbox in raw_spans:
        cy = (bbox[1] + bbox[3]) / 2
        target = None
        for group in groups:
            if abs(cy - float(group["center_y"])) <= 3.0:
                target = group
                break
        if target is None:
            target = {"center_y": cy, "items": []}
            groups.append(target)
        target["items"].append({"text": text, "bbox": bbox})  # type: ignore[index]

    rows: list[dict[str, object]] = []
    for group in groups:
        items = sorted(group["items"], key=lambda item: item["bbox"][0])  # type: ignore[index]
        if not items:
            continue
        bbox = [
            min(item["bbox"][0] for item in items),
            min(item["bbox"][1] for item in items),
            max(item["bbox"][2] for item in items),
            max(item["bbox"][3] for item in items),
        ]
        rows.append({"items": items, "bbox": bbox})
    return rows


def _row_total_text_len(row: dict[str, object]) -> int:
    items = row["items"]  # type: ignore[index]
    return sum(len(str(item["text"]).strip()) for item in items)


def _row_large_gap_count(row: dict[str, object]) -> int:
    items = row["items"]  # type: ignore[index]
    large_gaps = 0
    for idx in range(len(items) - 1):
        gap = float(items[idx + 1]["bbox"][0] - items[idx]["bbox"][2])
        if gap >= 12:
            large_gaps += 1
    return large_gaps


def _row_looks_table_primary(row: dict[str, object], page_width: float) -> bool:
    items = row["items"]  # type: ignore[index]
    bbox = row["bbox"]  # type: ignore[index]
    if len(items) < 2:
        return False
    row_width = float(bbox[2] - bbox[0])
    if row_width < page_width * 0.35:
        return False
    return _row_large_gap_count(row) >= 1


def _row_can_extend_table_cluster(row: dict[str, object], cluster_bbox: list[float], page_width: float) -> bool:
    bbox = row["bbox"]  # type: ignore[index]
    items = row["items"]  # type: ignore[index]
    overlap = _horizontal_overlap_ratio(bbox, cluster_bbox)
    if overlap < 0.2:
        return False
    if _row_looks_table_primary(row, page_width):
        return True
    if len(items) >= 2 and _row_total_text_len(row) <= 40:
        return True
    return len(items) == 1 and _row_total_text_len(row) <= 16


def _probable_table_regions(page) -> list[list[float]]:
    page_width = float(getattr(page, "width", 0.0) or 0.0)
    if page_width <= 0:
        return []

    rows = _page_span_rows(page)
    if not rows:
        return []

    regions: list[list[float]] = []
    cluster: list[dict[str, object]] = []
    cluster_primary_count = 0
    for row in rows:
        is_primary = _row_looks_table_primary(row, page_width)
        if not cluster:
            if is_primary:
                cluster = [row]
                cluster_primary_count = 1
            continue

        prev_bbox = cluster[-1]["bbox"]  # type: ignore[index]
        bbox = row["bbox"]  # type: ignore[index]
        vertical_gap = float(bbox[1] - prev_bbox[3])
        cluster_bbox = [
            min(item["bbox"][0] for item in cluster),  # type: ignore[index]
            min(item["bbox"][1] for item in cluster),  # type: ignore[index]
            max(item["bbox"][2] for item in cluster),  # type: ignore[index]
            max(item["bbox"][3] for item in cluster),  # type: ignore[index]
        ]
        if vertical_gap <= 28 and _row_can_extend_table_cluster(row, cluster_bbox, page_width):
            cluster.append(row)
            if is_primary:
                cluster_primary_count += 1
            continue

        if cluster_primary_count >= 2 and len(cluster) >= 3:
            regions.append(
                [
                    min(item["bbox"][0] for item in cluster),  # type: ignore[index]
                    min(item["bbox"][1] for item in cluster),  # type: ignore[index]
                    max(item["bbox"][2] for item in cluster),  # type: ignore[index]
                    max(item["bbox"][3] for item in cluster),  # type: ignore[index]
                ]
            )
        cluster = [row] if is_primary else []
        cluster_primary_count = 1 if is_primary else 0

    if cluster_primary_count >= 2 and len(cluster) >= 3:
        regions.append(
            [
                min(item["bbox"][0] for item in cluster),  # type: ignore[index]
                min(item["bbox"][1] for item in cluster),  # type: ignore[index]
                max(item["bbox"][2] for item in cluster),  # type: ignore[index]
                max(item["bbox"][3] for item in cluster),  # type: ignore[index]
            ]
        )

    return regions


def _looks_like_caption_text(text: str) -> bool:
    raw = (text or "").strip()
    match = re.match(r"^(图|表)\s*([A-Z]?\d+(?:[\.．\-－]\d+)*)(.*)$", raw)
    if not match:
        return False

    title_text = (match.group(3) or "").strip()
    if not title_text:
        return True

    normalized = _normalize_text(title_text)
    if not normalized:
        return False

    if len(normalized) > 40:
        return False

    if re.search(r"[。！？!?]", normalized):
        return False

    if re.match(r"^(列出|列出了|给出|给出了|展示|展示了|说明|说明了|反映|反映了|表示|表示了|体现|体现了|验证|验证了|用于|可以|通过|采用|实现|描述)", normalized):
        return False

    if any(token in normalized for token in ("可以", "通过", "用于说明", "进行了", "如图", "如表")):
        return False

    return True


def _caption_candidates(page, *, kind: str):
    prefix = "图" if kind == "figure" else "表"
    pattern = re.compile(rf"^{prefix}\s*[A-Z]?\d+(?:[\.．\-－]\d+)*(?:\s+.+)?$")

    for line in _page_text_lines(page):
        text = str(line["text"]).strip()
        bbox = list(line["bbox"])  # type: ignore[index]
        if not text or len(bbox) != 4 or len(text) > 80:
            continue
        if not pattern.match(text):
            continue
        if not _looks_like_caption_text(text):
            continue
        yield {"text": text, "bbox": [float(x) for x in bbox]}


def _regions(page, *, kind: str):
    yielded: list[list[float]] = []
    for region in getattr(page, "regions", []):
        if getattr(region, "kind", "") != kind:
            continue
        bbox = getattr(region, "bbox", None) or []
        if len(bbox) != 4:
            continue
        box = [float(x) for x in bbox]
        yielded.append(box)
        yield box

    if kind == "table":
        for box in _probable_table_regions(page):
            duplicate = False
            for existing in yielded:
                if (
                    abs(existing[0] - box[0]) <= 6
                    and abs(existing[1] - box[1]) <= 6
                    and abs(existing[2] - box[2]) <= 6
                    and abs(existing[3] - box[3]) <= 6
                ):
                    duplicate = True
                    break
            if not duplicate:
                yield box


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
    if overlap < 0.2:
        return False

    if caption_expected == "below":
        caption_near_top = caption_box[1] <= page_height * 0.45
        region_near_bottom = region_box[3] >= page_height * 0.55
        return caption_near_top and region_near_bottom

    caption_near_bottom = caption_box[3] >= page_height * 0.55
    region_near_top = region_box[1] <= page_height * 0.45
    return caption_near_bottom and region_near_top


def _find_nearby_region(page, caption_box, *, kind: str, caption_expected: str):
    for region_box in _regions(page, kind=kind):
        overlap = _horizontal_overlap_ratio(caption_box, region_box)
        gap = _vertical_gap_to_region(caption_box, region_box, caption_expected=caption_expected)
        if overlap >= 0.3 and 0 <= gap <= 80:
            return region_box
    return None


def _find_continuation_region(next_page, current_region_box, *, kind: str):
    page_height = float(getattr(next_page, "height", 0.0) or 0.0)
    if page_height <= 0:
        return None

    best_region = None
    best_score = -1.0
    current_width = max(1.0, current_region_box[2] - current_region_box[0])
    for region_box in _regions(next_page, kind=kind):
        if region_box[1] > page_height * 0.35:
            continue
        overlap = _horizontal_overlap_ratio(current_region_box, region_box)
        if overlap < 0.2:
            continue
        next_width = max(1.0, region_box[2] - region_box[0])
        width_ratio = min(current_width, next_width) / max(current_width, next_width)
        score = overlap * 0.7 + width_ratio * 0.3
        if score > best_score:
            best_score = score
            best_region = region_box
    return best_region


def _find_previous_continuation_region(prev_page, current_region_box, *, kind: str):
    page_height = float(getattr(prev_page, "height", 0.0) or 0.0)
    if page_height <= 0:
        return None

    best_region = None
    best_score = -1.0
    current_width = max(1.0, current_region_box[2] - current_region_box[0])
    for region_box in _regions(prev_page, kind=kind):
        if region_box[3] < page_height * 0.55:
            continue
        overlap = _horizontal_overlap_ratio(current_region_box, region_box)
        if overlap < 0.2:
            continue
        prev_width = max(1.0, region_box[2] - region_box[0])
        width_ratio = min(current_width, prev_width) / max(current_width, prev_width)
        score = overlap * 0.7 + width_ratio * 0.3
        if score > best_score:
            best_score = score
            best_region = region_box
    return best_region


def _has_caption_before_region(page, region_box, *, kind: str):
    prefix = "图" if kind == "figure" else "表"
    pattern = re.compile(rf"^{prefix}\s*[A-Z]?\d+(?:[\.．\-－]\d+)*(?:\s+.+)?$")
    for line in _page_text_lines(page):
        text = str(line["text"]).strip()
        bbox = list(line["bbox"])  # type: ignore[index]
        if len(bbox) != 4 or not text or not pattern.match(text):
            continue
        span_box = [float(x) for x in bbox]
        if span_box[3] > region_box[1]:
            continue
        if region_box[1] - span_box[3] > 120:
            continue
        if _horizontal_overlap_ratio(span_box, region_box) < 0.2:
            continue
        return True
    return False


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
    if not _looks_like_caption_text(text):
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
            for line in _page_text_lines(page):
                text = str(line["text"]).strip()
                bbox = list(line["bbox"])  # type: ignore[index]
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
        seen_keys: set[tuple[str, int, str, str]] = set()
        for page_index, page in enumerate(pages):
            page_no = getattr(page, "page_no", None)
            table_boxes = list(_regions(page, kind="table"))
            if not table_boxes:
                continue

            for caption in _caption_candidates(page, kind="table"):
                caption_box = caption["bbox"]
                nearby_region = _find_nearby_region(page, caption_box, kind="table", caption_expected="above")
                if nearby_region is not None:
                    if page_index + 1 < len(pages):
                        next_page = pages[page_index + 1]
                        page_height = float(getattr(page, "height", 0.0) or 0.0)
                        if page_height > 0 and nearby_region[3] >= page_height * 0.7:
                            continuation_region = _find_continuation_region(next_page, nearby_region, kind="table")
                            if continuation_region is not None and not _has_caption_before_region(next_page, continuation_region, kind="table"):
                                key = ("table", page_no or -1, caption["text"], "body")
                                if key not in seen_keys:
                                    seen_keys.add(key)
                                    split_page_no = getattr(next_page, "page_no", None)
                                    issues.append(
                                        Issue(
                                            rule_id=self.rule_id,
                                            title="表题",
                                            message=f"第{page_no}页表题“{caption['text']}”对应表格主体疑似延续到第{split_page_no}页。",
                                            severity=Severity.WARNING,
                                            source=Source.PDF,
                                            page=page_no,
                                            bbox=[int(x) for x in caption_box],
                                            fixable=False,
                                            metadata={
                                                "section": "表题",
                                                "content": caption["text"],
                                                "problem": "表格主体疑似被拆开排版为两页",
                                                "related_page": split_page_no,
                                            },
                                        )
                                    )
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
            page_height = float(getattr(page, "height", 0.0) or 0.0)

            for caption in _caption_candidates(page, kind="figure"):
                caption_box = caption["bbox"]
                nearby_region = _find_nearby_region(page, caption_box, kind="image", caption_expected="below")
                if nearby_region is not None:
                    if page_index + 1 < len(pages) and page_height > 0 and nearby_region[3] >= page_height * 0.7:
                        next_page = pages[page_index + 1]
                        continuation_region = _find_continuation_region(next_page, nearby_region, kind="image")
                        if continuation_region is not None and not _has_caption_before_region(next_page, continuation_region, kind="figure"):
                            key = ("figure", page_no or -1, caption["text"], "body_next")
                            if key not in seen_keys:
                                seen_keys.add(key)
                                split_page_no = getattr(next_page, "page_no", None)
                                issues.append(
                                    Issue(
                                        rule_id=self.rule_id,
                                        title="图题",
                                        message=f"第{page_no}页图题“{caption['text']}”对应插图主体疑似延续到第{split_page_no}页。",
                                        severity=Severity.WARNING,
                                        source=Source.PDF,
                                        page=page_no,
                                        bbox=[int(x) for x in caption_box],
                                        fixable=False,
                                        metadata={
                                            "section": "图题",
                                            "content": caption["text"],
                                            "problem": "插图主体疑似被拆开排版为两页",
                                            "related_page": split_page_no,
                                        },
                                    )
                                )
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
                nearby_region = _find_nearby_region(page, caption_box, kind="table", caption_expected="above")
                if nearby_region is not None:
                    if page_index > 0 and page_height > 0 and nearby_region[1] <= page_height * 0.3:
                        prev_page = pages[page_index - 1]
                        continuation_region = _find_previous_continuation_region(prev_page, nearby_region, kind="table")
                        if continuation_region is not None and not _has_caption_before_region(page, nearby_region, kind="table"):
                            key = ("table", page_no or -1, caption["text"], "body_prev")
                            if key not in seen_keys:
                                seen_keys.add(key)
                                split_page_no = getattr(prev_page, "page_no", None)
                                issues.append(
                                    Issue(
                                        rule_id=self.rule_id,
                                        title="表题",
                                        message=f"第{page_no}页表题“{caption['text']}”对应表格主体疑似从第{split_page_no}页延续而来。",
                                        severity=Severity.WARNING,
                                        source=Source.PDF,
                                        page=page_no,
                                        bbox=[int(x) for x in caption_box],
                                        fixable=False,
                                        metadata={
                                            "section": "表题",
                                            "content": caption["text"],
                                            "problem": "表格主体疑似被拆开排版为两页",
                                            "related_page": split_page_no,
                                        },
                                    )
                                )
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
