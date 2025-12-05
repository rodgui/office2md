"""
office2md - Convert Microsoft Office documents to Markdown.

Supports:
- DOCX (Word) with Pandoc/Mammoth/python-docx
- XLSX (Excel) with openpyxl
- PPTX (PowerPoint) with python-pptx
- PDF with Docling (optional)
"""

from office2md.__version__ import __version__, __version_info__
from office2md.converter_factory import ConverterFactory
from office2md.converters.base_converter import BaseConverter

__all__ = [
    "__version__",
    "__version_info__",
    "ConverterFactory",
    "BaseConverter",
]
