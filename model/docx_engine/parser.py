from model.compat.docx_structure_adapter import DocxStructureAdapter
from model.core.context import RuleContext


class DocxContextBuilder:
    def __init__(self) -> None:
        self.adapter = DocxStructureAdapter()

    def build(self, file_path: str) -> RuleContext:
        doc_obj = None
        sections = None
        parse_error = None
        parts_order = None
        try:
            doc_obj, sections = self.adapter.parse(file_path)
            parts_order = list(getattr(self.adapter, "parts_order", []) or [])
        except Exception as exc:  # noqa: BLE001
            parse_error = str(exc)

        return RuleContext(
            file_path=file_path,
            docx_obj=doc_obj,
            docx_sections=sections,
            extras={
                "docx_parse_error": parse_error,
                "docx_parts_order": parts_order,
            },
        )
