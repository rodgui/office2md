"""Converter for DOCX files."""

import logging
import re
from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

try:
    import mammoth

    MAMMOTH_AVAILABLE = True
except ImportError:
    MAMMOTH_AVAILABLE = False

try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    PYTHON_DOCX_AVAILABLE = True
except ImportError:
    PYTHON_DOCX_AVAILABLE = False


class DocxConverter(BaseConverter):
    """Converter for DOCX files using mammoth with python-docx fallback."""

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        use_mammoth: bool = True,
        **kwargs
    ):
        """
        Initialize DOCX converter.

        Args:
            input_path: Path to input DOCX file
            output_path: Optional output path
            use_mammoth: Use mammoth if available (default: True)
            **kwargs: Additional options (extract_images, embed_images, skip_images)
        """
        super().__init__(input_path, output_path, **kwargs)
        self.use_mammoth = use_mammoth and MAMMOTH_AVAILABLE

    def convert(self) -> str:
        """Convert DOCX to Markdown with image processing."""
        try:
            if self.use_mammoth:
                markdown = self._convert_with_mammoth()
                # Post-process mammoth output for better formatting
                markdown = self._clean_mammoth_output(markdown)
            else:
                if not PYTHON_DOCX_AVAILABLE:
                    raise RuntimeError(
                        "Neither mammoth nor python-docx available for DOCX conversion"
                    )
                markdown = self._convert_with_python_docx()

            # Process images based on mode (extract/embed/skip)
            markdown = self._replace_base64_images(markdown)
            return markdown

        except Exception as e:
            logger.error(f"Error converting DOCX: {e}")
            raise

    def _convert_with_mammoth(self) -> str:
        """
        Convert using mammoth library with enhanced style mappings.

        Returns:
            Markdown string
        """
        logger.info("Using mammoth for DOCX conversion")

        try:
            # Enhanced style map for better markdown conversion
            style_map = """
p[style-name='Heading 1'] => h1:fresh
p[style-name='Heading 2'] => h2:fresh
p[style-name='Heading 3'] => h3:fresh
p[style-name='Heading 4'] => h4:fresh
p[style-name='Heading 5'] => h5:fresh
p[style-name='Heading 6'] => h6:fresh
p[style-name='Título'] => h1:fresh
p[style-name='Subtítulo'] => h2:fresh
p[style-name='Title'] => h1:fresh
p[style-name='Subtitle'] => h2:fresh
r[style-name='Strong'] => strong
r[style-name='Emphasis'] => em
            """

            with open(self.input_path, "rb") as docx_file:
                result = mammoth.convert_to_markdown(
                    docx_file,
                    style_map=style_map,
                    convert_image=mammoth.images.inline(mammoth.images.data_uri),
                )

                if result.messages:
                    for msg in result.messages:
                        if msg.type == "warning":
                            logger.debug(f"Mammoth: {msg.message}")

                return result.value

        except Exception as e:
            logger.error(f"Mammoth conversion failed: {e}")
            if PYTHON_DOCX_AVAILABLE:
                logger.info("Falling back to python-docx")
                return self._convert_with_python_docx()
            raise

    def _clean_mammoth_output(self, markdown: str) -> str:
        """
        Post-process mammoth output to fix common formatting issues.

        Args:
            markdown: Raw mammoth markdown output

        Returns:
            Cleaned markdown string
        """
        # Remove unnecessary escapes
        markdown = re.sub(r"\\\.", ".", markdown)
        markdown = re.sub(r"\\\(", "(", markdown)
        markdown = re.sub(r"\\\)", ")", markdown)

        # Fix heading levels (mammoth sometimes uses ### for h1)
        lines = markdown.split("\n")
        cleaned_lines = []

        for line in lines:
            # Convert bold headings to proper markdown headings
            if line.strip().startswith("**") and line.strip().endswith("**"):
                # Check if it looks like a heading (short, no punctuation at end except ?)
                content = line.strip()[2:-2].strip()
                if len(content) < 100 and not content.endswith("."):
                    line = f"# {content}"

            # Fix heading level nesting (ensure proper hierarchy)
            if line.startswith("###"):
                # Keep as h3
                pass
            elif line.startswith("##"):
                # Keep as h2
                pass
            elif line.startswith("#"):
                # Keep as h1
                pass

            cleaned_lines.append(line)

        return "\n".join(cleaned_lines)

    def _convert_with_python_docx(self) -> str:
        """
        Convert using python-docx library with comprehensive formatting support.

        Returns:
            Markdown string
        """
        logger.info("Using python-docx for DOCX conversion")

        if not PYTHON_DOCX_AVAILABLE:
            raise RuntimeError("python-docx is not available")

        try:
            doc = Document(self.input_path)
            markdown_lines = []

            in_list = False
            prev_num_id = None

            for element in doc.element.body:
                # Handle paragraphs
                if isinstance(element, CT_P):
                    para = Paragraph(element, doc)

                    # Check if paragraph is in a list
                    num_id = self._get_numbering_id(para)

                    if num_id is not None:
                        # List item
                        if not in_list:
                            in_list = True
                            if markdown_lines and markdown_lines[-1] != "":
                                markdown_lines.append("")

                        md_line = self._paragraph_to_markdown(para, in_list=True)
                        if md_line:
                            markdown_lines.append(md_line)

                        prev_num_id = num_id
                    else:
                        # Regular paragraph
                        if in_list:
                            in_list = False
                            markdown_lines.append("")

                        md_line = self._paragraph_to_markdown(para, in_list=False)
                        if md_line:
                            markdown_lines.append(md_line)
                        elif not in_list:
                            markdown_lines.append("")

                # Handle tables
                elif hasattr(element, "tag") and "tbl" in element.tag:
                    for table in doc.tables:
                        if table._element == element:
                            if in_list:
                                in_list = False
                            if markdown_lines and markdown_lines[-1] != "":
                                markdown_lines.append("")
                            markdown_lines.append(self._table_to_markdown(table))
                            markdown_lines.append("")
                            break

            return "\n".join(markdown_lines)

        except Exception as e:
            logger.error(f"python-docx conversion failed: {e}")
            raise

    def _get_numbering_id(self, para: "Paragraph") -> Optional[int]:
        """
        Get numbering ID if paragraph is part of a list.

        Args:
            para: python-docx Paragraph object

        Returns:
            Numbering ID or None
        """
        try:
            pPr = para._element.pPr
            if pPr is not None:
                numPr = pPr.numPr
                if numPr is not None:
                    numId = numPr.numId
                    if numId is not None:
                        return numId.val
        except Exception:
            pass
        return None

    def _paragraph_to_markdown(self, para: "Paragraph", in_list: bool = False) -> str:
        """
        Convert paragraph to Markdown with proper formatting.

        Args:
            para: python-docx Paragraph object
            in_list: Whether paragraph is part of a list

        Returns:
            Markdown formatted string
        """
        text = para.text.strip()
        if not text:
            return ""

        # Handle headings
        style_name = para.style.name.lower()

        if "heading" in style_name or "título" in style_name:
            try:
                if "heading" in style_name:
                    level = int(para.style.name.split()[-1])
                else:
                    level = 1
                return f"{'#' * level} {text}"
            except (ValueError, IndexError):
                return f"# {text}"

        if style_name in ["title", "título"]:
            return f"# {text}"
        if style_name in ["subtitle", "subtítulo"]:
            return f"## {text}"

        # Handle lists
        if in_list:
            # Get indentation level
            indent_level = 0
            if para.paragraph_format.left_indent:
                indent_level = int(para.paragraph_format.left_indent / 360000)

            indent = "  " * indent_level

            # Check if numbered (simple heuristic: starts with digit)
            if text and text[0].isdigit():
                return f"{indent}1. {text}"
            else:
                return f"{indent}- {text}"

        # Apply run-level formatting (bold/italic)
        formatted_text = self._apply_run_formatting(para)

        return formatted_text

    def _apply_run_formatting(self, para: "Paragraph") -> str:
        """
        Apply run-level formatting (bold, italic) to paragraph text.

        Args:
            para: python-docx Paragraph object

        Returns:
            Markdown formatted string with inline styles
        """
        result = []

        for run in para.runs:
            text = run.text
            if not text:
                continue

            # Apply bold
            if run.bold:
                text = f"**{text}**"
            # Apply italic
            elif run.italic:
                text = f"*{text}*"
            # Apply underline (use italic as markdown alternative)
            elif run.underline:
                text = f"_{text}_"

            result.append(text)

        return "".join(result)

    def _table_to_markdown(self, table: "Table") -> str:
        """
        Convert a table to Markdown format with proper alignment.

        Args:
            table: python-docx Table object

        Returns:
            Markdown table string
        """
        if not table.rows:
            return ""

        rows = []

        for i, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                # Get cell text, handle merged cells and newlines
                cell_text = cell.text.strip().replace("\n", " ").replace("|", "\\|")
                cells.append(cell_text if cell_text else " ")

            # Format row
            rows.append("| " + " | ".join(cells) + " |")

            # Add separator after header row
            if i == 0:
                separator = "|" + "|".join([" --- " for _ in cells]) + "|"
                rows.append(separator)

        return "\n".join(rows)
