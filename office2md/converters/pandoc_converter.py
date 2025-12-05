"""Pandoc-based converter for DOCX files."""

import logging
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
from html.parser import HTMLParser

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)


def is_pandoc_available() -> bool:
    """Check if Pandoc is installed on the system."""
    return shutil.which('pandoc') is not None


PANDOC_AVAILABLE = is_pandoc_available()


class HTMLTableParser(HTMLParser):
    """Parser to extract table data from HTML."""
    
    def __init__(self):
        super().__init__()
        self.tables = []
        self.current_table = []
        self.current_row = []
        self.current_cell = ""
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.is_header = False
    
    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.in_table = True
            self.current_table = []
        elif tag == 'tr':
            self.in_row = True
            self.current_row = []
        elif tag in ['td', 'th']:
            self.in_cell = True
            self.current_cell = ""
            self.is_header = (tag == 'th')
        elif tag == 'img' and self.in_cell:
            # Extract image src
            for attr_name, attr_value in attrs:
                if attr_name == 'src':
                    self.current_cell += f"![image]({attr_value})"
        elif tag == 'br':
            self.current_cell += " "
    
    def handle_endtag(self, tag):
        if tag == 'table':
            if self.current_table:
                self.tables.append(self.current_table)
            self.in_table = False
            self.current_table = []
        elif tag == 'tr':
            if self.current_row:
                self.current_table.append(self.current_row)
            self.in_row = False
            self.current_row = []
        elif tag in ['td', 'th']:
            self.current_row.append(self.current_cell.strip())
            self.in_cell = False
            self.current_cell = ""
    
    def handle_data(self, data):
        if self.in_cell:
            self.current_cell += data


class PandocConverter(BaseConverter):
    """
    DOCX converter using Pandoc (external binary).
    
    Pandoc provides high-quality direct DOCX → Markdown conversion
    with excellent support for tables, headings, lists, and images.
    
    Requires: `brew install pandoc` (macOS) or equivalent.
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        **kwargs
    ):
        """Initialize Pandoc converter."""
        super().__init__(input_path, output_path, **kwargs)
        
        if not PANDOC_AVAILABLE:
            raise RuntimeError(
                "Pandoc is not installed. Install with:\n"
                "  macOS: brew install pandoc\n"
                "  Ubuntu: sudo apt-get install pandoc\n"
                "  Windows: choco install pandoc"
            )

    def convert(self) -> str:
        """Convert DOCX to Markdown using Pandoc."""
        logger.info(f"Converting with Pandoc: {self.input_path}")
        
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                media_dir = Path(tmpdir) / "media"
                
                # Build Pandoc command - force pipe tables
                cmd = [
                    'pandoc',
                    str(self.input_path),
                    '-f', 'docx',
                    '-t', 'markdown+pipe_tables-simple_tables-multiline_tables-grid_tables',
                    '--wrap=none',
                    '--markdown-headings=atx',
                    f'--extract-media={tmpdir}',
                ]
                
                # Run Pandoc
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=120
                )
                
                if result.returncode != 0:
                    raise RuntimeError(f"Pandoc error: {result.stderr}")
                
                markdown = result.stdout
                
                # Extract images and build path mapping
                path_mapping = {}
                if self.extract_images and media_dir.exists():
                    path_mapping = self._extract_and_map_images(media_dir)
                
                # Replace image paths BEFORE converting HTML tables
                if path_mapping:
                    markdown = self._replace_image_paths(markdown, path_mapping)
                elif self.skip_images:
                    markdown = re.sub(r'!\[.*?\]\(.*?\)', '', markdown)
                
                # Convert any HTML tables to Markdown pipe tables
                markdown = self._convert_html_tables_to_markdown(markdown)
                
                # Clean up Pandoc-specific artifacts
                markdown = self._cleanup_pandoc_output(markdown)
                
                # Final pass: update any remaining temp image paths
                if path_mapping:
                    markdown = self._replace_image_paths(markdown, path_mapping)
                
                logger.info("Pandoc conversion completed successfully")
                return markdown
                
        except subprocess.TimeoutExpired:
            raise RuntimeError("Pandoc conversion timed out")
        except Exception as e:
            logger.error(f"Pandoc conversion failed: {e}")
            raise

    def _extract_and_map_images(self, media_dir: Path) -> Dict[str, str]:
        """
        Extract images from Pandoc media directory and create path mapping.
        """
        path_mapping = {}
        
        image_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.emf', '.wmf', '.tiff', '.bmp']
        image_files = []
        
        if media_dir.exists():
            for ext in image_extensions:
                image_files.extend(media_dir.glob(f"*{ext}"))
                image_files.extend(media_dir.glob(f"*{ext.upper()}"))
        
        if not image_files:
            logger.debug("No images found in Pandoc media directory")
            return path_mapping
        
        if self.images_dir:
            self.images_dir.mkdir(parents=True, exist_ok=True)
        
        for img_file in sorted(image_files, key=lambda x: x.name):
            try:
                with open(img_file, 'rb') as f:
                    image_data = f.read()
                
                ext = img_file.suffix[1:].lower()
                if ext in ['emf', 'wmf']:
                    ext = 'png'
                
                new_ref = self._process_image(image_data, ext)
                
                if new_ref:
                    # Map ALL possible path variations
                    path_mapping[str(img_file)] = new_ref
                    path_mapping[img_file.name] = new_ref
                    path_mapping[f"media/{img_file.name}"] = new_ref
                    
                    try:
                        rel_path = str(img_file.relative_to(img_file.parent.parent))
                        path_mapping[rel_path] = new_ref
                    except ValueError:
                        pass
                    
                    logger.debug(f"Mapped image: {img_file.name} -> {new_ref}")
                    
            except Exception as e:
                logger.warning(f"Failed to process image {img_file}: {e}")
        
        return path_mapping

    def _replace_image_paths(self, markdown: str, path_mapping: Dict[str, str]) -> str:
        """Replace all image paths in markdown with new references."""
        
        def replace_md_image(match):
            alt_text = match.group(1) or ''
            img_path = match.group(2)
            
            new_ref = self._find_image_mapping(img_path, path_mapping)
            
            if new_ref:
                if alt_text:
                    return new_ref.replace('![]', f'![{alt_text}]')
                return new_ref
            
            # Keep original if not in temp path
            if '/tmp' not in img_path and '/var/folders' not in img_path:
                return match.group(0)
            
            logger.warning(f"No mapping found for temp image: {img_path}")
            return ''
        
        markdown = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', replace_md_image, markdown)
        
        return markdown

    def _find_image_mapping(self, img_path: str, path_mapping: Dict[str, str]) -> Optional[str]:
        """Find the new reference for an image path."""
        if img_path in path_mapping:
            return path_mapping[img_path]
        
        img_name = Path(img_path).name
        if img_name in path_mapping:
            return path_mapping[img_name]
        
        for old_path, ref in path_mapping.items():
            if img_name in old_path or old_path.endswith(img_name):
                return ref
            if old_path in img_path:
                return ref
        
        return None

    def _convert_html_tables_to_markdown(self, markdown: str) -> str:
        """
        Convert all HTML tables to Markdown pipe tables.
        
        This ensures consistent output format regardless of 
        what Pandoc generates.
        """
        # Find all HTML tables
        table_pattern = re.compile(r'<table[^>]*>.*?</table>', re.DOTALL | re.IGNORECASE)
        
        def convert_table(match):
            html_table = match.group(0)
            return self._html_table_to_markdown(html_table)
        
        markdown = table_pattern.sub(convert_table, markdown)
        
        # Also handle tables with colgroup
        colgroup_table_pattern = re.compile(r'<colgroup>.*?</colgroup>\s*<thead>.*?</tbody>', re.DOTALL | re.IGNORECASE)
        
        # Handle orphaned table parts (colgroup, thead, tbody without table wrapper)
        if '<colgroup>' in markdown or '<thead>' in markdown or '<tbody>' in markdown:
            markdown = self._convert_orphaned_html_table_parts(markdown)
        
        return markdown

    def _html_table_to_markdown(self, html_table: str) -> str:
        """Convert a single HTML table to Markdown pipe table."""
        parser = HTMLTableParser()
        
        try:
            parser.feed(html_table)
        except Exception as e:
            logger.warning(f"Failed to parse HTML table: {e}")
            return html_table
        
        if not parser.tables:
            return html_table
        
        # Use the first (and usually only) table
        table_data = parser.tables[0]
        
        if not table_data:
            return ""
        
        return self._format_markdown_table(table_data)

    def _convert_orphaned_html_table_parts(self, markdown: str) -> str:
        """
        Convert orphaned HTML table parts (without <table> wrapper) to Markdown.
        
        Sometimes Pandoc outputs just colgroup/thead/tbody without table tags.
        """
        lines = markdown.split('\n')
        result = []
        table_html = []
        in_table_part = False
        
        for line in lines:
            # Check if line contains table HTML elements
            has_table_element = any(tag in line.lower() for tag in [
                '<colgroup', '<thead', '<tbody', '<tr', '<td', '<th',
                '</colgroup', '</thead', '</tbody', '</tr', '</td', '</th'
            ])
            
            if has_table_element:
                if not in_table_part:
                    in_table_part = True
                    table_html = []
                table_html.append(line)
            else:
                if in_table_part:
                    # End of table part, convert it
                    full_html = f"<table>{''.join(table_html)}</table>"
                    md_table = self._html_table_to_markdown(full_html)
                    result.append(md_table)
                    in_table_part = False
                    table_html = []
                result.append(line)
        
        # Handle table at end of document
        if in_table_part and table_html:
            full_html = f"<table>{''.join(table_html)}</table>"
            md_table = self._html_table_to_markdown(full_html)
            result.append(md_table)
        
        return '\n'.join(result)

    def _format_markdown_table(self, table_data: List[List[str]]) -> str:
        """
        Format table data as Markdown pipe table.
        
        Args:
            table_data: List of rows, each row is a list of cell contents
            
        Returns:
            Markdown formatted table string
        """
        if not table_data:
            return ""
        
        # Normalize column count
        max_cols = max(len(row) for row in table_data)
        normalized = []
        for row in table_data:
            normalized_row = row + [''] * (max_cols - len(row))
            # Clean cell content
            cleaned_row = [self._clean_cell_content(cell) for cell in normalized_row]
            normalized.append(cleaned_row)
        
        lines = []
        
        # Header row
        header = normalized[0]
        lines.append('| ' + ' | '.join(header) + ' |')
        
        # Separator row
        lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
        
        # Data rows
        for row in normalized[1:]:
            lines.append('| ' + ' | '.join(row) + ' |')
        
        return '\n'.join(lines)

    def _clean_cell_content(self, content: str) -> str:
        """Clean cell content for Markdown table."""
        # Remove newlines
        content = content.replace('\n', ' ').replace('\r', '')
        
        # Remove multiple spaces
        content = re.sub(r'\s+', ' ', content)
        
        # Escape pipe characters
        content = content.replace('|', '\\|')
        
        # Remove HTML entities
        content = content.replace('&amp;', '&')
        content = content.replace('&lt;', '<')
        content = content.replace('&gt;', '>')
        content = content.replace('&nbsp;', ' ')
        
        return content.strip()

    def _cleanup_pandoc_output(self, markdown: str) -> str:
        """
        Clean up Pandoc-specific artifacts from markdown.
        
        Removes:
        - Image dimension metadata {width="..." height="..."}
        - Orphaned HTML tags
        - Extra backslashes
        """
        # Remove image dimension metadata
        markdown = re.sub(r'\{width="[^"]*"(\s*height="[^"]*")?\}', '', markdown)
        markdown = re.sub(r'\{height="[^"]*"\}', '', markdown)
        markdown = re.sub(r'\{style="[^"]*"\}', '', markdown)
        
        # Remove other Pandoc attributes
        markdown = re.sub(r'\{[#.][^}]*\}', '', markdown)
        
        # Remove any remaining HTML img tags with temp paths
        markdown = re.sub(r'<img[^>]*src="[^"]*tmp[^"]*"[^>]*/?\s*>', '', markdown)
        markdown = re.sub(r'<img[^>]*src="[^"]*var/folders[^"]*"[^>]*/?\s*>', '', markdown)
        
        # Remove empty paragraph tags
        markdown = re.sub(r'<p>\s*</p>', '', markdown)
        
        # Remove orphaned HTML tags that aren't part of tables
        markdown = re.sub(r'</?colgroup[^>]*>', '', markdown)
        markdown = re.sub(r'<col[^>]*/?\s*>', '', markdown)
        markdown = re.sub(r'</?thead[^>]*>', '', markdown)
        markdown = re.sub(r'</?tbody[^>]*>', '', markdown)
        
        # Remove excessive backslashes
        markdown = re.sub(r'\\([.,:;!?])', r'\1', markdown)
        
        # Clean up multiple blank lines
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        
        # Clean up trailing whitespace
        lines = markdown.split('\n')
        lines = [line.rstrip() for line in lines]
        markdown = '\n'.join(lines)
        
        return markdown