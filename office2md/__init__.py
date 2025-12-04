"""
office2md - Convert Office files to Markdown

This package provides converters for various Office file formats:
- DOCX (Word documents)
- XLSX/XLS (Excel spreadsheets)
- PPTX/PPT (PowerPoint presentations)
"""

__version__ = "0.1.0"

from office2md.converter_factory import ConverterFactory
from office2md.converters.docx_converter import DocxConverter
from office2md.converters.pptx_converter import PptxConverter
from office2md.converters.xlsx_converter import XlsxConverter

__all__ = [
    "DocxConverter",
    "XlsxConverter",
    "PptxConverter",
    "ConverterFactory",
]
