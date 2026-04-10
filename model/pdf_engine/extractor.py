from pathlib import Path

from model.pdf_engine.models import PdfPage, PdfSpan


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
                for block in text_dict.get("blocks", []):
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            text = str(span.get("text", "")).strip()
                            bbox = list(span.get("bbox", []))
                            if text:
                                spans.append(PdfSpan(text=text, bbox=bbox))
                pages.append(PdfPage(page_no=idx, spans=spans))
            return pages, None
        except Exception as exc:  # noqa: BLE001
            return [], str(exc)
