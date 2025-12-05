"""Converter for DOCX files using hybrid mammoth + python-docx approach."""

import base64
import logging
import re
from pathlib import Path
from typing import Optional, List, Tuple

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
    """
    Hybrid DOCX converter using both mammoth and python-docx.
    
    Strategy:
    - Use mammoth to extract ALL images (including from tables)
    - Use python-docx for structure (tables, headings, lists)
    - Track image positions and insert references correctly
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        use_mammoth: bool = True,
        **kwargs
    ):
        """Initialize DOCX converter."""
        super().__init__(input_path, output_path, **kwargs)
        self.use_mammoth = use_mammoth and MAMMOTH_AVAILABLE
        self._mammoth_images: List[str] = []
        self._image_index: int = 0  # Global image counter

    def convert(self) -> str:
        """Convert DOCX to Markdown using hybrid approach."""
        try:
            # STEP 1: Extract ALL images with mammoth (including from tables)
            if self.use_mammoth and MAMMOTH_AVAILABLE:
                self._extract_images_with_mammoth()
            
            # STEP 2: Convert structure with python-docx
            markdown = self._convert_with_python_docx_structured()
            
            # STEP 3: Fix formatting issues with smart spacing
            markdown = self._fix_bold_formatting(markdown)
            
            # STEP 4: Final cleanup
            markdown = self._cleanup_markdown(markdown)
            
            return markdown

        except Exception as e:
            logger.error(f"Error converting DOCX: {e}", exc_info=True)
            raise

    def _extract_images_with_mammoth(self) -> None:
        """Extract all images from DOCX using mammoth."""
        logger.info("Extracting images with mammoth")
        
        try:
            with open(self.input_path, "rb") as docx_file:
                result = mammoth.convert_to_markdown(
                    docx_file,
                    convert_image=mammoth.images.data_uri
                )
                
                mammoth_md = result.value
                
                pattern = r"!\[([^\]]*)\]\(data:image/([a-zA-Z0-9+]+);base64,([A-Za-z0-9+/=\s]+)\)"
                matches = re.finditer(pattern, mammoth_md)
                
                for match in matches:
                    image_format = match.group(2).lower().replace("+xml", "")
                    b64_data = match.group(3).replace(" ", "").replace("\n", "").replace("\r", "")
                    
                    try:
                        image_data = base64.b64decode(b64_data)
                        ref = self._process_image(image_data, image_format)
                        if ref:
                            self._mammoth_images.append(ref)
                    except Exception as e:
                        logger.warning(f"Failed to decode image: {e}")
                
                logger.info(f"Extracted {len(self._mammoth_images)} images from mammoth")
                
                if hasattr(result, 'messages'):
                    try:
                        for msg in list(result.messages or []):
                            msg_type = getattr(msg, 'type', 'unknown')
                            if msg_type == "warning":
                                logger.debug(f"Mammoth: {getattr(msg, 'message', str(msg))}")
                    except (TypeError, AttributeError):
                        pass
                        
        except Exception as e:
            logger.warning(f"Mammoth image extraction failed: {e}")

    def _get_next_image_ref(self) -> str:
        """Get the next image reference from extracted images."""
        if self._image_index < len(self._mammoth_images):
            ref = self._mammoth_images[self._image_index]
            self._image_index += 1
            return ref
        return ""

    def _count_images_in_element(self, element) -> int:
        """Count how many images are in an XML element."""
        try:
            xml = element.xml if hasattr(element, 'xml') else str(element)
            # Count drawing and pict elements (images)
            drawings = xml.count('w:drawing') + xml.count('w:pict')
            return drawings
        except Exception:
            return 0

    def _convert_with_python_docx_structured(self) -> str:
        """Convert using python-docx with proper structure support."""
        logger.info("Converting structure with python-docx")
        
        if not PYTHON_DOCX_AVAILABLE:
            raise RuntimeError("python-docx is not available")

        try:
            doc = Document(self.input_path)
            markdown_lines = []
            in_list = False
            
            for element in doc.element.body:
                if isinstance(element, CT_P):
                    para = Paragraph(element, doc)
                    
                    # Check for inline images in paragraph
                    image_count = self._count_images_in_element(para._element)
                    if image_count > 0:
                        # Insert image references
                        for _ in range(image_count):
                            img_ref = self._get_next_image_ref()
                            if img_ref:
                                markdown_lines.append("")
                                markdown_lines.append(img_ref)
                                markdown_lines.append("")
                        
                        # Also include any text in the paragraph
                        text = para.text.strip()
                        if text:
                            markdown_lines.append(text)
                        continue
                    
                    md_line = self._paragraph_to_markdown(para)
                    
                    if self._is_list_paragraph(para):
                        if not in_list:
                            in_list = True
                            if markdown_lines and markdown_lines[-1] != "":
                                markdown_lines.append("")
                        if md_line:
                            markdown_lines.append(md_line)
                    else:
                        if in_list:
                            in_list = False
                            markdown_lines.append("")
                        if md_line:
                            markdown_lines.append(md_line)
                
                elif hasattr(element, 'tag') and 'tbl' in element.tag:
                    if in_list:
                        in_list = False
                    
                    for table in doc.tables:
                        if table._element == element:
                            if markdown_lines and markdown_lines[-1] != "":
                                markdown_lines.append("")
                            
                            # Convert table WITH image support
                            table_md = self._table_to_markdown_with_images(table)
                            markdown_lines.append(table_md)
                            markdown_lines.append("")
                            break

            return "\n".join(markdown_lines)

        except Exception as e:
            logger.error(f"python-docx conversion failed: {e}")
            raise

    def _table_to_markdown_with_images(self, table: "Table") -> str:
        """
        Convert table to Markdown, handling images in cells.
        
        For cells with images:
        - If cell has ONLY image: insert image reference
        - If cell has image + text: insert image reference + text
        """
        if not table.rows:
            return ""
        
        rows = []
        
        for i, row in enumerate(table.rows):
            cells = []
            for cell in row.cells:
                cell_content = self._process_table_cell(cell)
                cells.append(cell_content if cell_content else " ")
            
            rows.append("| " + " | ".join(cells) + " |")
            
            if i == 0:
                separator = "|" + "|".join([" --- " for _ in cells]) + "|"
                rows.append(separator)

        return "\n".join(rows)

    def _process_table_cell(self, cell) -> str:
        """
        Process a table cell, extracting text and handling images.
        
        Returns cell content as markdown string.
        """
        parts = []
        
        # Check for images in cell
        image_count = self._count_images_in_element(cell._element)
        
        if image_count > 0:
            # Cell contains image(s)
            for _ in range(image_count):
                img_ref = self._get_next_image_ref()
                if img_ref:
                    parts.append(img_ref)
        
        # Get text content
        text = cell.text.strip().replace("\n", " ").replace("|", "\\|").replace("\r", "")
        if text:
            parts.append(text)
        
        return " ".join(parts)

    def _paragraph_has_image(self, para: "Paragraph") -> bool:
        """Check if paragraph contains an inline image."""
        return self._count_images_in_element(para._element) > 0

    def _paragraph_to_markdown(self, para: "Paragraph") -> str:
        """Convert paragraph to Markdown."""
        text = para.text.strip()
        if not text:
            return ""

        style_name = para.style.name if para.style else ""
        style_lower = style_name.lower()
        
        # Heading detection
        if 'heading' in style_lower:
            try:
                level = int(style_name.split()[-1])
            except (ValueError, IndexError):
                level = 1
            return f"{'#' * level} {text}"
        
        if 'título' in style_lower or 'titulo' in style_lower:
            return f"# {text}"
        if 'subtítulo' in style_lower or 'subtitle' in style_lower:
            return f"## {text}"
        if style_lower == 'title':
            return f"# {text}"

        # List detection
        if self._is_list_paragraph(para):
            indent_level = self._get_list_indent_level(para)
            indent = "  " * indent_level
            
            if self._is_numbered_list(para):
                return f"{indent}1. {text}"
            else:
                return f"{indent}- {text}"

        # Apply inline formatting
        formatted_text = self._apply_run_formatting(para)
        return formatted_text if formatted_text else text

    def _is_list_paragraph(self, para: "Paragraph") -> bool:
        """Check if paragraph is part of a list."""
        style_name = para.style.name if para.style else ""
        
        list_styles = [
            "List", "List Paragraph", "Lista", "PargrafodaLista",
            "List Bullet", "List Number", "Bullet", "Numbered"
        ]
        
        if any(style in style_name for style in list_styles):
            return True
        
        try:
            pPr = para._element.pPr
            if pPr is not None:
                return pPr.numPr is not None
        except Exception:
            pass
        
        return False

    def _is_numbered_list(self, para: "Paragraph") -> bool:
        """Check if list paragraph is numbered."""
        try:
            pPr = para._element.pPr
            if pPr is not None and pPr.numPr is not None:
                return True
        except Exception:
            pass
        return False

    def _get_list_indent_level(self, para: "Paragraph") -> int:
        """Get indentation level for list items."""
        try:
            if para.paragraph_format.left_indent:
                return int(para.paragraph_format.left_indent / 360000)
        except Exception:
            pass
        return 0

    def _apply_run_formatting(self, para: "Paragraph") -> str:
        """Apply run-level formatting with proper spacing."""
        result = []
        prev_bold = False
        prev_italic = False
        
        for run in para.runs:
            text = run.text
            if not text:
                continue
            
            curr_bold = bool(run.bold)
            curr_italic = bool(run.italic)
            
            # Close previous formatting if style changed
            if prev_bold and not curr_bold:
                if result:
                    result[-1] = result[-1].rstrip() + "**"
            if prev_italic and not curr_italic:
                if result:
                    result[-1] = result[-1].rstrip() + "*"
            
            # Open new formatting if style changed
            if curr_bold and not prev_bold:
                text = f"**{text}"
            if curr_italic and not prev_italic:
                text = f"*{text}"
            
            result.append(text)
            prev_bold = curr_bold
            prev_italic = curr_italic
        
        # Close any open formatting at the end
        final = "".join(result)
        if prev_bold:
            final = final.rstrip() + "**"
        if prev_italic:
            final = final.rstrip() + "*"
        
        return final

    def _is_section_title(self, bold_text: str) -> bool:
        """Determine if bold text is a section title or inline emphasis."""
        text = bold_text.strip()
        
        if len(text) > 60:
            return False
        
        if text and text[-1] in '.,:;!?':
            return False
        
        if text and text[0].isupper() and len(text) < 40:
            return True
        
        return len(text) < 50

    def _get_separator(self, before_char: str, after_char: str, bold_text: str) -> Tuple[str, str]:
        """Determine what separator to use before and after bold text."""
        is_title = self._is_section_title(bold_text)
        
        if is_title:
            before_sep = "\n\n" if before_char and before_char not in '\n' else ""
            after_sep = "\n\n" if after_char and after_char not in '\n' else ""
        else:
            before_sep = " " if before_char and before_char not in ' \n' else ""
            after_sep = " " if after_char and after_char not in ' \n' else ""
        
        return before_sep, after_sep

    def _fix_bold_formatting(self, markdown: str) -> str:
        """Fix bold formatting with intelligent spacing decisions."""
        # Step 1: Fix multiple asterisks
        markdown = re.sub(r'\*{4,}', '**', markdown)
        
        # Step 2: Find all bold patterns and fix spacing intelligently
        pattern = r'(.?)\*\*([^\*]+)\*\*(.?)'
        
        def smart_replace(match):
            before_char = match.group(1)
            bold_text = match.group(2).strip()
            after_char = match.group(3)
            
            if not bold_text:
                return before_char + after_char
            
            before_sep, after_sep = self._get_separator(before_char, after_char, bold_text)
            
            if before_char in ' \n':
                before_char = ''
            if after_char in ' \n':
                after_char = ''
            
            return f"{before_char}{before_sep}**{bold_text}**{after_sep}{after_char}"
        
        markdown = re.sub(pattern, smart_replace, markdown)
        
        # Step 3: Clean up
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        markdown = re.sub(r'  +', ' ', markdown)
        markdown = re.sub(r'\n +', '\n', markdown)
        
        return markdown

    def _cleanup_markdown(self, markdown: str) -> str:
        """Final cleanup pass on markdown."""
        lines = markdown.split('\n')
        cleaned = []
        
        for line in lines:
            if line.strip() == "":
                if not cleaned or cleaned[-1].strip() != "":
                    cleaned.append("")
            else:
                cleaned.append(line)
        
        while cleaned and cleaned[0] == "":
            cleaned.pop(0)
        while cleaned and cleaned[-1] == "":
            cleaned.pop()
        
        return "\n".join(cleaned)
