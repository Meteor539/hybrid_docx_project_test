from dataclasses import dataclass, field
from typing import Any, Optional


@dataclass
class RuleContext:
    file_path: str
    docx_obj: Optional[Any] = None
    docx_sections: Optional[dict[str, Any]] = None
    pdf_pages: Optional[list[Any]] = None
    ocr_pages: Optional[list[dict[str, Any]]] = None
    extras: dict[str, Any] = field(default_factory=dict)

