"""Compatibility layer exports.

Only the active DOCX structure adapter is exported from the package root.
Legacy format / vision adapters are kept on disk for reference and optional
direct imports, but they are not part of the default mixed-check pipeline.
"""

from model.compat.docx_structure_adapter import DocxStructureAdapter

__all__ = ["DocxStructureAdapter"]
