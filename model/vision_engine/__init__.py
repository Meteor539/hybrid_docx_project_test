"""Vision-side utilities.

The current default mixed-check pipeline relies on PDF page extraction rather
than the placeholder OCR chain. Keep OCR-related modules available via direct
imports, but expose only the rendering helper from the package root.
"""

from model.vision_engine.renderer import DocumentRenderer

__all__ = ["DocumentRenderer"]
