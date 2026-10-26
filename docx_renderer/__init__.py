"""Offline Pure-Python OpenXML (.docx) to PNG Layout and Rendering Engine."""

from .fonts import FontManager
from .layout import LayoutEngine
from .model import PageLayout, ParsedDocument
from .parser import DocxParser
from .renderer import CanvasRenderer, parse_page_spec, render_docx_to_png

__all__ = [
    "FontManager",
    "DocxParser",
    "LayoutEngine",
    "CanvasRenderer",
    "PageLayout",
    "ParsedDocument",
    "parse_page_spec",
    "render_docx_to_png",
]
