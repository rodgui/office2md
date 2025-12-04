"""office2md - Convert Office files to Markdown."""

try:
    from office2md._version import version as __version__
except ImportError:
    __version__ = "0.1.0"  # fallback version

__author__ = "Your Name"
__license__ = "MIT"

from office2md.converter_factory import ConverterFactory
from office2md.converters.base_converter import BaseConverter
from office2md.converters.docx_converter import DocxConverter
from office2md.converters.xlsx_converter import XlsxConverter
from office2md.converters.pptx_converter import PptxConverter

__all__ = [
    "ConverterFactory",
    "BaseConverter",
    "DocxConverter",
    "XlsxConverter",
    "PptxConverter",
    "__version__",
]
