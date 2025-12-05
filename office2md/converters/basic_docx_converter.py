"""Basic DOCX converter using python-docx (fallback when Pandoc/Mammoth unavailable)."""

import logging
import re
from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

try:
    import docx
    from docx.document import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    PYTHON_DOCX_AVAILABLE = True
except ImportError:
    PYTHON_DOCX_AVAILABLE = False


class BasicDocxConverter(BaseConverter):
    """
    Basic DOCX converter using python-docx.
    
    This is the fallback converter when Pandoc and Mammoth are not available.
    It provides basic text and table extraction but limited formatting support.
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        **kwargs
    ):
        """Initialize basic DOCX converter."""
        super().__init__(input_path, output_path, **kwargs)
        
        if not PYTHON_DOCX_AVAILABLE:
            raise RuntimeError("python-docx not available. Install with: pip install python-docx")

    def convert(self) -> str:
        """Convert DOCX to Markdown using python-docx."""
        logger.info(f"Converting with python-docx (basic): {self.input_path}")
        
        doc = docx.Document(self.input_path)
        
        # Create images directory
        if self.images_dir and self.extract_images:
            self.images_dir.mkdir(parents=True, exist_ok=True)
        
        markdown_parts = []
        
        # Process document body
        for element in doc.element.body:
            tag = element.tag.split('}')[-1]  # Remove namespace
            
            if tag == 'p':
                # Paragraph
                para = Paragraph(element, doc)
                md = self._paragraph_to_markdown(para)
                if md:
                    markdown_parts.append(md)
                    
            elif tag == 'tbl':
                # Table
                table = Table(element, doc)
                md = self._table_to_markdown(table)
                if md:
                    markdown_parts.append(md)
        
        # Extract images from document
        if self.extract_images and not self.skip_images:
            self._extract_images(doc)
        
        markdown = '\n\n'.join(markdown_parts)
        
        # Clean up
        markdown = self._cleanup_markdown(markdown)
        
        logger.info("Basic python-docx conversion completed")
        return markdown

    def _paragraph_to_markdown(self, para: 'Paragraph') -> str:
        """Convert paragraph to Markdown."""
        text = para.text.strip()
        if not text:
            return ""
        
        # Check for heading style
        style_name = para.style.name if para.style else ""
        
        if style_name.startswith('Heading'):
            try:
                level = int(style_name.replace('Heading', '').strip())
                return '#' * level + ' ' + text
            except ValueError:
                pass
        
        if style_name == 'Title':
            return '# ' + text
        
        # Check for list
        if style_name.startswith('List'):
            return '- ' + text
        
        # Apply inline formatting
        formatted = self._apply_inline_formatting(para)
        
        return formatted

    def _apply_inline_formatting(self, para: 'Paragraph') -> str:
        """Apply bold/italic formatting to paragraph text."""
        parts = []
        
        for run in para.runs:
            text = run.text
            if not text:
                continue
            
            if run.bold and run.italic:
                text = f"***{text}***"
            elif run.bold:
                text = f"**{text}**"
            elif run.italic:
                text = f"*{text}*"
            
            parts.append(text)
        
        return ''.join(parts)

    def _table_to_markdown(self, table: 'Table') -> str:
        """Convert table to Markdown."""
        rows = []
        
        for row in table.rows:
            cells = []
            for cell in row.cells:
                # Get cell text, handling merged cells
                cell_text = cell.text.replace('\n', ' ').replace('|', '\\|').strip()
                cells.append(cell_text)
            rows.append(cells)
        
        if not rows:
            return ""
        
        # Normalize column count
        max_cols = max(len(row) for row in rows)
        normalized = [row + [''] * (max_cols - len(row)) for row in rows]
        
        lines = []
        
        # Header row
        lines.append('| ' + ' | '.join(normalized[0]) + ' |')
        
        # Separator
        lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
        
        # Data rows
        for row in normalized[1:]:
            lines.append('| ' + ' | '.join(row) + ' |')
        
        return '\n'.join(lines)

    def _extract_images(self, doc) -> None:
        """Extract images from document."""
        try:
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    try:
                        image_data = rel.target_part.blob
                        ext = rel.target_ref.split('.')[-1].lower()
                        if ext == 'jpeg':
                            ext = 'jpg'
                        
                        self._process_image(image_data, ext)
                        
                    except Exception as e:
                        logger.debug(f"Failed to extract image: {e}")
                        
        except Exception as e:
            logger.warning(f"Image extraction failed: {e}")

    def _cleanup_markdown(self, markdown: str) -> str:
        """Clean up generated Markdown."""
        # Remove multiple blank lines
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        
        # Fix broken bold/italic
        markdown = re.sub(r'\*{4,}', '**', markdown)
        
        # Clean trailing whitespace
        lines = markdown.split('\n')
        lines = [line.rstrip() for line in lines]
        
        return '\n'.join(lines).strip()