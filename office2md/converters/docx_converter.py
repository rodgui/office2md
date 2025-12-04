"""Converter for DOCX files using mammoth (primary) and python-docx (fallback)."""

import logging
from typing import Optional

from office2md.converters.base_converter import BaseConverter

try:
    import mammoth

    MAMMOTH_AVAILABLE = True
except ImportError:
    MAMMOTH_AVAILABLE = False
    logging.warning("mammoth not available, using python-docx fallback")

try:
    import docx

    PYTHON_DOCX_AVAILABLE = True
except ImportError:
    PYTHON_DOCX_AVAILABLE = False


class DocxConverter(BaseConverter):
    """Converter for DOCX files to Markdown."""

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        use_mammoth: bool = True,
    ):
        """
        Initialize the DOCX converter.

        Args:
            input_path: Path to the input DOCX file
            output_path: Optional path for the output Markdown file
            use_mammoth: If True, use mammoth for high-quality conversion.
                        If False or mammoth unavailable, use python-docx
        """
        super().__init__(input_path, output_path)
        self.use_mammoth = use_mammoth and MAMMOTH_AVAILABLE

    def convert(self) -> str:
        """
        Convert DOCX to Markdown.

        Returns:
            The Markdown content as a string
        """
        if self.use_mammoth:
            return self._convert_with_mammoth()
        elif PYTHON_DOCX_AVAILABLE:
            return self._convert_with_python_docx()
        else:
            raise RuntimeError(
                "No DOCX converter available. Install mammoth or python-docx."
            )

    def _convert_with_mammoth(self) -> str:
        """Convert using mammoth library for high-quality conversion."""
        self.logger.info("Using mammoth for DOCX conversion")
        with open(self.input_path, "rb") as docx_file:
            result = mammoth.convert_to_markdown(docx_file)
            if result.messages:
                for message in result.messages:
                    self.logger.debug(f"Mammoth message: {message}")
            return result.value

    def _convert_with_python_docx(self) -> str:
        """Convert using python-docx as fallback."""
        self.logger.info("Using python-docx for DOCX conversion")
        doc = docx.Document(self.input_path)
        markdown_lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                markdown_lines.append("")
                continue

            # Detect heading styles
            if para.style.name.startswith("Heading"):
                try:
                    level = int(para.style.name.split()[-1])
                    markdown_lines.append(f"{'#' * level} {text}")
                except (ValueError, IndexError):
                    markdown_lines.append(text)
            else:
                markdown_lines.append(text)

        # Process tables
        for table in doc.tables:
            markdown_lines.append("")
            for i, row in enumerate(table.rows):
                cells = [cell.text.strip() for cell in row.cells]
                markdown_lines.append("| " + " | ".join(cells) + " |")
                if i == 0:  # Add header separator
                    markdown_lines.append(
                        "| " + " | ".join(["---"] * len(cells)) + " |"
                    )
            markdown_lines.append("")

        return "\n".join(markdown_lines)
