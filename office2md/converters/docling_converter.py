"""Docling-based converter for DOCX and PDF files."""

import base64
import io
import logging
import re
from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

try:
    from docling.document_converter import DocumentConverter
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.document_converter import PdfFormatOption
    DOCLING_AVAILABLE = True
except ImportError:
    DOCLING_AVAILABLE = False


class DoclingConverter(BaseConverter):
    """
    Document converter using Docling (IBM Research).
    
    Docling provides ML-based document parsing with excellent
    support for complex layouts, tables, and images.
    
    Supports: DOCX, PDF, PPTX, and more.
    
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
        
        self._image_counter = 0

    def convert(self) -> str:
        """Convert document to Markdown using Docling."""
        logger.info(f"Converting with Docling: {self.input_path}")
        
        try:
            # Initialize Docling converter with image extraction enabled
            pipeline_options = PdfPipelineOptions()
            pipeline_options.generate_picture_images = True
            
            converter = DocumentConverter(
                format_options={
                    InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
                }
            )
            
            # Convert document
            result = converter.convert(str(self.input_path))
            doc = result.document
            
            # Export to Markdown
            markdown = doc.export_to_markdown()
            
            # Process images
            if self.extract_images and not self.skip_images:
                markdown = self._process_docling_images(markdown, doc)
            elif self.skip_images:
                # Remove image placeholders
                markdown = re.sub(r'!\[.*?\]\(.*?\)', '', markdown)
                markdown = re.sub(r'<!--\s*image\s*-->', '', markdown)
            
            # Clean up markdown
            markdown = self._cleanup_docling_output(markdown)
            
            logger.info("Docling conversion completed successfully")
            return markdown
            
        except Exception as e:
            logger.error(f"Docling conversion failed: {e}")
            raise

    def _process_docling_images(self, markdown: str, doc) -> str:
        """
        Process images from Docling document.
        
        Docling stores images in document.pictures or as page images.
        """
        try:
            # Create images directory
            if self.images_dir:
                self.images_dir.mkdir(parents=True, exist_ok=True)
            
            image_refs = []
            
            # Method 1: Extract from pictures collection
            if hasattr(doc, 'pictures') and doc.pictures:
                for picture in doc.pictures:
                    ref = self._extract_docling_picture(picture)
                    if ref:
                        image_refs.append(ref)
            
            # Method 2: Extract from pages
            if hasattr(doc, 'pages') and doc.pages:
                for page in doc.pages:
                    if hasattr(page, 'image') and page.image:
                        ref = self._extract_page_image(page.image)
                        if ref:
                            image_refs.append(ref)
                    
                    # Check for pictures in page
                    if hasattr(page, 'pictures'):
                        for picture in page.pictures:
                            ref = self._extract_docling_picture(picture)
                            if ref:
                                image_refs.append(ref)
            
            # Method 3: Check document items
            if hasattr(doc, 'items'):
                for item in doc.items:
                    if hasattr(item, 'image'):
                        ref = self._extract_item_image(item)
                        if ref:
                            image_refs.append(ref)
            
            # Replace placeholders with actual image references
            if image_refs:
                markdown = self._replace_image_placeholders(markdown, image_refs)
            
            return markdown
            
        except Exception as e:
            logger.warning(f"Failed to process Docling images: {e}")
            return markdown

    def _extract_docling_picture(self, picture) -> Optional[str]:
        """Extract image from Docling picture object."""
        try:
            image_data = None
            ext = 'png'
            
            # Try different ways to get image data
            if hasattr(picture, 'image'):
                if hasattr(picture.image, 'pil_image') and picture.image.pil_image:
                    # PIL Image
                    img_buffer = io.BytesIO()
                    picture.image.pil_image.save(img_buffer, format='PNG')
                    image_data = img_buffer.getvalue()
                elif hasattr(picture.image, 'data'):
                    image_data = picture.image.data
                elif hasattr(picture.image, 'uri') and picture.image.uri:
                    if picture.image.uri.startswith('data:'):
                        match = re.match(r'data:image/(\w+);base64,(.+)', picture.image.uri)
                        if match:
                            ext = match.group(1)
                            image_data = base64.b64decode(match.group(2))
            
            elif hasattr(picture, 'data'):
                image_data = picture.data
            
            elif hasattr(picture, 'pil_image') and picture.pil_image:
                img_buffer = io.BytesIO()
                picture.pil_image.save(img_buffer, format='PNG')
                image_data = img_buffer.getvalue()
            
            if image_data:
                return self._process_image(image_data, ext)
            
        except Exception as e:
            logger.debug(f"Failed to extract picture: {e}")
        
        return None

    def _extract_page_image(self, page_image) -> Optional[str]:
        """Extract image from page image object."""
        try:
            if hasattr(page_image, 'pil_image') and page_image.pil_image:
                img_buffer = io.BytesIO()
                page_image.pil_image.save(img_buffer, format='PNG')
                return self._process_image(img_buffer.getvalue(), 'png')
        except Exception as e:
            logger.debug(f"Failed to extract page image: {e}")
        return None

    def _extract_item_image(self, item) -> Optional[str]:
        """Extract image from document item."""
        try:
            if hasattr(item, 'image') and item.image:
                if hasattr(item.image, 'pil_image') and item.image.pil_image:
                    img_buffer = io.BytesIO()
                    item.image.pil_image.save(img_buffer, format='PNG')
                    return self._process_image(img_buffer.getvalue(), 'png')
        except Exception as e:
            logger.debug(f"Failed to extract item image: {e}")
        return None

    def _replace_image_placeholders(self, markdown: str, image_refs: list) -> str:
        """
        Replace image placeholders with actual image references.
        
        Docling may use various placeholder formats:
        - <!-- image -->
        - [image]
        - {image}
        - ![](...)
        """
        ref_index = 0
        
        # Pattern for various image placeholders
        patterns = [
            r'<!--\s*image\s*-->',
            r'\[image\d*\]',
            r'\{image\d*\}',
            r'!\[\]\(\s*\)',  # Empty image reference
        ]
        
        for pattern in patterns:
            def replace_placeholder(match):
                nonlocal ref_index
                if ref_index < len(image_refs):
                    ref = image_refs[ref_index]
                    ref_index += 1
                    return ref
                return match.group(0)
            
            markdown = re.sub(pattern, replace_placeholder, markdown)
        
        # If we still have unused images, append them at positions where they might belong
        # Look for lines that might be image captions
        if ref_index < len(image_refs):
            remaining_refs = image_refs[ref_index:]
            
            # Append remaining images at the end with a note
            if remaining_refs:
                markdown += "\n\n<!-- Additional extracted images -->\n"
                for ref in remaining_refs:
                    markdown += f"\n{ref}\n"
        
        return markdown

    def _cleanup_docling_output(self, markdown: str) -> str:
        """Clean up Docling markdown output."""
        # Remove empty image references
        markdown = re.sub(r'!\[\]\(\s*\)', '', markdown)
        
        # Remove orphaned image placeholders
        markdown = re.sub(r'<!--\s*image\s*-->', '', markdown)
        
        # Clean up multiple blank lines
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        
        # Clean up trailing whitespace
        lines = markdown.split('\n')
        lines = [line.rstrip() for line in lines]
        markdown = '\n'.join(lines)
        
        return markdown