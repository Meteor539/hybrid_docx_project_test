import re


PAGE_ROLE_OTHER = "other"
PAGE_ROLE_CN_ABSTRACT = "cn_abstract"
PAGE_ROLE_EN_ABSTRACT = "en_abstract"
PAGE_ROLE_CATALOGUE = "catalogue"
PAGE_ROLE_MAIN = "main"
PAGE_ROLE_BACKMATTER = "backmatter"


def normalize_text(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "")


def page_has_heading(page, keywords: tuple[str, ...]) -> bool:
    normalized = normalize_text(getattr(page, "text", ""))
    return any(keyword in normalized for keyword in keywords)


def is_catalogue_page(page) -> bool:
    normalized = normalize_text(getattr(page, "text", ""))
    if "目录" in normalized or "目次" in normalized:
        return True

    raw_text = getattr(page, "text", "") or ""
    dotted_entries = len(
        re.findall(r"[.\-·•_ ]{2,}\s*[IVXLCDMivxlcdm\d]+\s*$", raw_text, flags=re.MULTILINE)
    )
    numbered_entries = len(
        re.findall(r"^\s*(?:第\s*\d+\s*章|\d+\.\d+(?:\.\d+)?)", raw_text, flags=re.MULTILINE)
    )
    return dotted_entries >= 5 and numbered_entries >= 3


def looks_like_first_main_page(page) -> bool:
    raw_text = getattr(page, "text", "") or ""
    normalized = normalize_text(raw_text)
    if "第1章" in normalized or "第一章" in normalized:
        return True
    return bool(
        re.search(
            r"^\s*(?:1[.、\s]+(?:绪论|引言)|第\s*1\s*章(?:\s|$)|第一章(?:\s|$)|(?:绪论|引言)\s*$)",
            raw_text,
            flags=re.MULTILINE,
        )
    )


def top_area_texts(page) -> list[str]:
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


def top_area_lines(page, top_ratio: float = 0.25) -> list[str]:
    page_height = float(getattr(page, "height", 0.0) or 0.0)
    if page_height <= 0:
        return []

    rows = []
    spans = sorted(
        [span for span in getattr(page, "spans", []) if getattr(span, "text", "").strip()],
        key=lambda item: ((item.bbox[1] + item.bbox[3]) / 2, item.bbox[0]),
    )
    for span in spans:
        bbox = getattr(span, "bbox", None) or []
        if len(bbox) != 4:
            continue
        x0, y0, x1, y1 = [float(value) for value in bbox]
        if y1 > page_height * top_ratio:
            continue

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

    lines = []
    for row in rows:
        items = sorted(row["items"], key=lambda item: item["bbox"][0])
        text = "".join(item["text"] for item in items).strip()
        if text:
            lines.append(text)
    return lines


def page_has_top_heading(page, keywords: tuple[str, ...]) -> bool:
    normalized_keywords = {normalize_text(keyword) for keyword in keywords if keyword}
    if not normalized_keywords:
        return False

    for line in top_area_lines(page):
        compact = normalize_text(line)
        if not compact:
            continue
        for keyword in normalized_keywords:
            if compact == keyword:
                return True
            if compact.startswith(keyword) and len(compact) <= max(len(keyword) + 8, 20):
                return True
    return False


def looks_like_backmatter_start(page) -> bool:
    texts = top_area_texts(page)
    if not texts:
        return False

    for text in texts[:6]:
        compact = normalize_text(text)
        if compact in {"参考文献", "致谢", "致謝"}:
            return True
        if re.fullmatch(r"附录[A-ZＡ-Ｚ]?", compact):
            return True
    return False


def build_page_roles(pages) -> dict[int, str]:
    roles = {getattr(page, "page_no", idx + 1): PAGE_ROLE_OTHER for idx, page in enumerate(pages)}
    if not pages:
        return roles

    for idx, page in enumerate(pages):
        if is_catalogue_page(page):
            roles[getattr(page, "page_no", idx + 1)] = PAGE_ROLE_CATALOGUE

    first_main_idx = None
    for idx, page in enumerate(pages):
        page_no = getattr(page, "page_no", idx + 1)
        if roles.get(page_no) == PAGE_ROLE_CATALOGUE:
            continue
        if looks_like_first_main_page(page):
            first_main_idx = idx
            break

    if first_main_idx is None:
        return roles

    backmatter_indices = []
    for idx in range(first_main_idx + 1, len(pages)):
        if looks_like_backmatter_start(pages[idx]):
            backmatter_indices.append(idx)
    main_end_idx = min(backmatter_indices) if backmatter_indices else len(pages)

    for idx in range(first_main_idx, main_end_idx):
        roles[getattr(pages[idx], "page_no", idx + 1)] = PAGE_ROLE_MAIN

    for idx in range(main_end_idx, len(pages)):
        page_no = getattr(pages[idx], "page_no", idx + 1)
        if roles.get(page_no) == PAGE_ROLE_CATALOGUE:
            continue
        roles[page_no] = PAGE_ROLE_BACKMATTER

    for idx in range(first_main_idx):
        page = pages[idx]
        page_no = getattr(page, "page_no", idx + 1)
        if roles.get(page_no) == PAGE_ROLE_CATALOGUE:
            continue
        if page_has_top_heading(page, ("摘要",)) and not page_has_top_heading(page, ("ABSTRACT", "Abstract")):
            roles[page_no] = PAGE_ROLE_CN_ABSTRACT
        elif page_has_top_heading(page, ("ABSTRACT", "Abstract")):
            roles[page_no] = PAGE_ROLE_EN_ABSTRACT

    return roles


def expected_page_number_kind(role: str) -> str | None:
    if role in {PAGE_ROLE_CN_ABSTRACT, PAGE_ROLE_EN_ABSTRACT}:
        return "roman"
    if role in {PAGE_ROLE_MAIN, PAGE_ROLE_BACKMATTER}:
        return "arabic"
    return None
