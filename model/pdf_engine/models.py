from dataclasses import dataclass, field


@dataclass
class PdfSpan:
    text: str
    bbox: list[float]


@dataclass
class PdfRegion:
    kind: str
    bbox: list[float]


@dataclass
class PdfPage:
    page_no: int
    width: float = 0.0
    height: float = 0.0
    spans: list[PdfSpan] = field(default_factory=list)
    regions: list[PdfRegion] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(span.text for span in self.spans if span.text)
