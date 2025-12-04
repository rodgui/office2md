"""Factory for creating appropriate converters based on file type."""

from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter
from office2md.converters.docx_converter import DocxConverter
from office2md.converters.pptx_converter import PptxConverter
from office2md.converters.xlsx_converter import XlsxConverter


class ConverterFactory:
    """Factory class for creating appropriate converters."""

    SUPPORTED_EXTENSIONS = {
        ".docx": DocxConverter,
        ".xlsx": XlsxConverter,
        ".xls": XlsxConverter,
        ".pptx": PptxConverter,
        ".ppt": PptxConverter,
    }

    @classmethod
    def create_converter(
        cls, input_path: str, output_path: Optional[str] = None, **kwargs
    ) -> BaseConverter:
        """
        Create an appropriate converter based on file extension.

        Args:
            input_path: Path to the input file
            output_path: Optional path for the output Markdown file
            **kwargs: Additional arguments to pass to the converter

        Returns:
            An instance of the appropriate converter

        Raises:
            ValueError: If file extension is not supported
        """
        file_path = Path(input_path)
        extension = file_path.suffix.lower()

        converter_class = cls.SUPPORTED_EXTENSIONS.get(extension)
        if not converter_class:
            supported = ", ".join(cls.SUPPORTED_EXTENSIONS.keys())
            raise ValueError(
                f"Unsupported file type: {extension}. " f"Supported types: {supported}"
            )

        return converter_class(input_path, output_path, **kwargs)

    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        """
        Check if a file type is supported.

        Args:
            file_path: Path to the file

        Returns:
            True if the file type is supported, False otherwise
        """
        extension = Path(file_path).suffix.lower()
        return extension in cls.SUPPORTED_EXTENSIONS
