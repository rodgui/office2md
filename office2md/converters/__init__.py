"""Converters package for office2md."""

from office2md.converters.base_converter import BaseConverter
from office2md.converters.docx_converter import DocxConverter
from office2md.converters.xlsx_converter import XlsxConverter
from office2md.converters.pptx_converter import PptxConverter

# Optional converters
try:
    from office2md.converters.pandoc_converter import PandocConverter
except ImportError:
    PandocConverter = None

try:
    from office2md.converters.mammoth_converter import MammothConverter
except ImportError:
    MammothConverter = None

try:
    from office2md.converters.basic_docx_converter import BasicDocxConverter
except ImportError:
    BasicDocxConverter = None

try:
    from office2md.converters.docling_converter import DoclingConverter
except ImportError:
    DoclingConverter = None

__all__ = [
    "BaseConverter",
    "DocxConverter",
    "XlsxConverter",
    "PptxConverter",
    "PandocConverter",
    "MammothConverter",
    "BasicDocxConverter",
    "DoclingConverter",
]
