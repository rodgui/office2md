"""office2md - Convert Microsoft Office files to Markdown."""

from office2md.converter_factory import ConverterFactory
from office2md.converters.base_converter import BaseConverter
from office2md.converters.docx_converter import DocxConverter
from office2md.converters.xlsx_converter import XlsxConverter
from office2md.converters.pptx_converter import PptxConverter

# Optional converters (may not be available)
try:
    from office2md.converters.pandoc_converter import PandocConverter, PANDOC_AVAILABLE
except ImportError:
    PandocConverter = None
    PANDOC_AVAILABLE = False

try:
    from office2md.converters.docling_converter import DoclingConverter, DOCLING_AVAILABLE
except ImportError:
    DoclingConverter = None
    DOCLING_AVAILABLE = False

__version__ = "0.1.0"

__all__ = [
    "ConverterFactory",
    "BaseConverter",
    "DocxConverter",
    "XlsxConverter",
    "PptxConverter",
    "PandocConverter",
    "DoclingConverter",
    "PANDOC_AVAILABLE",
    "DOCLING_AVAILABLE",
]
