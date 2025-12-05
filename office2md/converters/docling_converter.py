"""Docling-based converter for PDF files."""

import base64
import io
import logging
import re
from pathlib import Path
from typing import Optional, List, Any

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

try:
    from docling.document_converter import DocumentConverter
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False

# Docling works best with PDFs - DOCX support is limited
DOCLING_SUPPORTED_EXTENSIONS = ['.pdf']


class DoclingConverter(BaseConverter):
    """
    Document converter using Docling (IBM Research).
    
    Docling provides ML-based document parsing with excellent
    support for PDFs, including scanned documents and complex layouts.
    
    **Note**: Docling is recommended only for PDF files.
    For DOCX files, use the default converter or Pandoc.
    
    Requires: `pip install docling`
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        **kwargs
    ):
        """Initialize Docling converter."""
        super().__init__(input_path, output_path, **kwargs)
        
        if not DOCLING_AVAILABLE:
            raise RuntimeError(
                "Docling is not installed. Install with:\n"
                "  pip install docling\n\n"
                "Note: Docling requires additional system dependencies.\n"
                "See: https://github.com/DS4SD/docling"
            )
        
        # Validate file type
        file_ext = Path(input_path).suffix.lower()
        if file_ext not in DOCLING_SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Docling is optimized for PDF files only.\n"
                f"File '{input_path}' has extension '{file_ext}'.\n"
                f"For DOCX files, use:\n"
                f"  - Default converter: office2md document.docx\n"
                f"  - Pandoc: office2md document.docx --use-pandoc"
            )
        
        self._image_counter = 0
        self._extracted_images: List[str] = []

    def convert(self) -> str:
        """Convert PDF to Markdown using Docling."""
        logger.info(f"Converting PDF with Docling: {self.input_path}")
        
        try:
            # Initialize Docling converter
            converter = DocumentConverter()
            
            # Convert document
            result = converter.convert(str(self.input_path))
            doc = result.document
            
            # Extract all images from the document
            if self.extract_images and not self.skip_images:
                self._extract_all_images(doc)
            
            # Export to Markdown
            markdown = doc.export_to_markdown()
            
            logger.debug(f"Original markdown length: {len(markdown)} chars")
            
            # Replace image placeholders with actual references
            if self._extracted_images:
                markdown = self._replace_image_placeholders(markdown)
            
            # Clean up markdown
            markdown = self._cleanup_docling_output(markdown)
            
            logger.info(f"Docling conversion completed ({len(markdown)} chars, {len(self._extracted_images)} images)")
            return markdown
            
        except Exception as e:
            logger.error(f"Docling conversion failed: {e}")
            raise

    def _extract_all_images(self, doc: Any) -> None:
        """Extract all images from Docling document."""
        if self.images_dir:
            self.images_dir.mkdir(parents=True, exist_ok=True)
        
        # From document pictures
        if hasattr(doc, 'pictures'):
            for picture in (doc.pictures or []):
                self._extract_picture(picture)
        
        # From document pages
        if hasattr(doc, 'pages'):
            for page in (doc.pages or []):
                if hasattr(page, 'image') and page.image:
                    self._extract_image_object(page.image)
                
                if hasattr(page, 'pictures'):
                    for picture in (page.pictures or []):
                        self._extract_picture(picture)
        
        # Iterate through all items
        if hasattr(doc, 'iterate_items'):
            try:
                for item, level in doc.iterate_items():
                    if hasattr(item, 'image') and item.image:
                        self._extract_image_object(item.image)
                    if hasattr(item, 'pil_image') and item.pil_image:
                        self._save_pil_image(item.pil_image)
            except Exception as e:
                logger.debug(f"iterate_items failed: {e}")
        
        logger.debug(f"Extracted {len(self._extracted_images)} images")

    def _extract_picture(self, picture: Any) -> Optional[str]:
        """Extract image from Docling picture object."""
        try:
            if hasattr(picture, 'pil_image') and picture.pil_image:
                return self._save_pil_image(picture.pil_image)
            
            if hasattr(picture, 'image'):
                return self._extract_image_object(picture.image)
            
            if hasattr(picture, 'data') and picture.data:
                ref = self._process_image(picture.data, 'png')
                if ref:
                    self._extracted_images.append(ref)
                    return ref
                    
        except Exception as e:
            logger.debug(f"Failed to extract picture: {e}")
        
        return None

    def _extract_image_object(self, image_obj: Any) -> Optional[str]:
        """Extract image from image object."""
        try:
            if hasattr(image_obj, 'pil_image') and image_obj.pil_image:
                return self._save_pil_image(image_obj.pil_image)
            
            if hasattr(image_obj, 'data') and image_obj.data:
                ref = self._process_image(image_obj.data, 'png')
                if ref:
                    self._extracted_images.append(ref)
                    return ref
            
            if hasattr(image_obj, 'uri') and image_obj.uri:
                if image_obj.uri.startswith('data:'):
                    match = re.match(r'data:image/(\w+);base64,(.+)', image_obj.uri)
                    if match:
                        ext = match.group(1)
                        data = base64.b64decode(match.group(2))
                        ref = self._process_image(data, ext)
                        if ref:
                            self._extracted_images.append(ref)
                            return ref
                            
        except Exception as e:
            logger.debug(f"Failed to extract image object: {e}")
        
        return None

    def _save_pil_image(self, pil_image: Any) -> Optional[str]:
        """Save PIL image and return markdown reference."""
        try:
            img_buffer = io.BytesIO()
            pil_image.save(img_buffer, format='PNG')
            ref = self._process_image(img_buffer.getvalue(), 'png')
            if ref:
                self._extracted_images.append(ref)
                return ref
        except Exception as e:
            logger.debug(f"Failed to save PIL image: {e}")
        return None

    def _replace_image_placeholders(self, markdown: str) -> str:
        """Replace image placeholders with actual image references."""
        if not self._extracted_images:
            return markdown
        
        ref_index = 0
        
        patterns = [
            r'<!--\s*image\s*-->',
            r'\[image\d*\]',
            r'\{image\d*\}',
            r'!\[\]\(\s*\)',
            r'\[\[image\]\]',
        ]
        
        for pattern in patterns:
            def make_replacer(idx_holder):
                def replacer(match):
                    if idx_holder[0] < len(self._extracted_images):
                        ref = self._extracted_images[idx_holder[0]]
                        idx_holder[0] += 1
                        return ref
                    return match.group(0)
                return replacer
            
            idx_holder = [ref_index]
            markdown = re.sub(pattern, make_replacer(idx_holder), markdown, flags=re.IGNORECASE)
            ref_index = idx_holder[0]
        
        # Append unused images at the end
        if ref_index < len(self._extracted_images):
            remaining = self._extracted_images[ref_index:]
            if remaining:
                markdown += "\n\n"
                for ref in remaining:
                    markdown += f"\n{ref}\n"
        
        return markdown

    def _cleanup_docling_output(self, markdown: str) -> str:
        """Clean up Docling markdown output."""
        markdown = re.sub(r'!\[\]\(\s*\)', '', markdown)
        markdown = re.sub(r'<!--\s*image\s*-->', '', markdown)
        markdown = re.sub(r'\[image\d*\]', '', markdown)
        markdown = re.sub(r'\{image\d*\}', '', markdown)
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        
        lines = markdown.split('\n')
        lines = [line.rstrip() for line in lines]
        markdown = '\n'.join(lines)
        
        return markdown