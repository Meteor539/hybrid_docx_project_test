from pathlib import Path

from model.pdf_engine.models import PdfPage, PdfRegion, PdfSpan


class PdfExtractor:
    """
    轻量级 PDF 文本与坐标提取器。
    可选依赖：PyMuPDF（fitz）。
    """

    def extract(self, pdf_path: str) -> tuple[list[PdfPage], str | None]:
        try:
            import fitz  # type: ignore
        except Exception:
            return [], "PyMuPDF is not installed"

        path = Path(pdf_path)
        if not path.exists():
            return [], f"PDF not found: {pdf_path}"

        pages: list[PdfPage] = []
        try:
            doc = fitz.open(pdf_path)
            for idx, page in enumerate(doc, start=1):
                text_dict = page.get_text("dict")
                spans: list[PdfSpan] = []
                regions: list[PdfRegion] = []
                for block in text_dict.get("blocks", []):
                    block_type = int(block.get("type", 0) or 0)
                    block_bbox = list(block.get("bbox", []))
                    if block_type == 1 and len(block_bbox) == 4:
                        regions.append(PdfRegion(kind="image", bbox=block_bbox))
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = str(span.get("text", "")).strip()
                            bbox = list(span.get("bbox", []))
                            if text:
                                spans.append(PdfSpan(text=text, bbox=bbox))

                try:
                    finder = getattr(page, "find_tables", None)
                    if callable(finder):
                        tables = finder()
                        for table in getattr(tables, "tables", []) or []:
                            bbox = list(getattr(table, "bbox", []) or [])
                            if len(bbox) == 4:
                                regions.append(PdfRegion(kind="table", bbox=bbox))
                except Exception:
                    pass

                rect = page.rect
                pages.append(
                    PdfPage(
                        page_no=idx,
                        width=float(rect.width),
                        height=float(rect.height),
                        spans=spans,
                        regions=regions,
                    )
                )
            return pages, None
        except Exception as exc:  # noqa: BLE001
            return [], str(exc)
