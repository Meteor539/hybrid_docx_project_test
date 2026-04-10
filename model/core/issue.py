from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class Severity(str, Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


class Source(str, Enum):
    DOCX = "docx"
    PDF = "pdf"
    OCR = "ocr"
    HYBRID = "hybrid"


@dataclass
class Issue:
    rule_id: str
    title: str
    message: str
    severity: Severity
    source: Source
    page: Optional[int] = None
    bbox: Optional[list[int]] = None
    confidence: float = 1.0
    fixable: bool = False
    fix_action: Optional[str] = None
    metadata: dict[str, Any] = field(default_factory=dict)

