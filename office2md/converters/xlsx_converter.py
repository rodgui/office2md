"""Converter for XLSX/XLS files using openpyxl."""

from typing import Optional

import openpyxl

from office2md.converters.base_converter import BaseConverter


class XlsxConverter(BaseConverter):
    """Converter for XLSX/XLS files to Markdown."""

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        include_all_sheets: bool = True,
    ):
        """
        Initialize the XLSX converter.

        Args:
            input_path: Path to the input XLSX file
            output_path: Optional path for the output Markdown file
            include_all_sheets: If True, include all sheets. If False, only first sheet
        """
        super().__init__(input_path, output_path)
        self.include_all_sheets = include_all_sheets

    def convert(self) -> str:
        """
        Convert XLSX to Markdown.

        Returns:
            The Markdown content as a string
        """
        self.logger.info(f"Converting XLSX file: {self.input_path}")
        workbook = openpyxl.load_workbook(self.input_path, data_only=True)
        markdown_lines = []

        sheets = workbook.worksheets if self.include_all_sheets else [workbook.active]

        for sheet in sheets:
            markdown_lines.append(f"## {sheet.title}")
            markdown_lines.append("")

            # Get all rows
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                markdown_lines.append("*Empty sheet*")
                markdown_lines.append("")
                continue

            # Filter out completely empty rows
            non_empty_rows = [
                row for row in rows if any(cell is not None for cell in row)
            ]

            if not non_empty_rows:
                markdown_lines.append("*Empty sheet*")
                markdown_lines.append("")
                continue

            # Find the maximum number of columns
            max_cols = max(len(row) for row in non_empty_rows)

            # Format as markdown table
            for i, row in enumerate(non_empty_rows):
                # Pad row to max_cols and convert None to empty string
                cells = [str(cell) if cell is not None else "" for cell in row]
                cells.extend([""] * (max_cols - len(cells)))
                markdown_lines.append("| " + " | ".join(cells) + " |")

                # Add separator after first row (header)
                if i == 0:
                    markdown_lines.append("| " + " | ".join(["---"] * max_cols) + " |")

            markdown_lines.append("")

        return "\n".join(markdown_lines)
