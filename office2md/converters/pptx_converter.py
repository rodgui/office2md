"""Converter for PPTX/PPT files."""

import logging
from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


class PptxConverter(BaseConverter):
    """Converter for PPTX/PPT files."""

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        include_notes: bool = True,
        **kwargs
    ):
        """
        Initialize PPTX converter.

        Args:
            input_path: Path to input PPTX file
            output_path: Optional output path
            include_notes: Include speaker notes (default: True)
            **kwargs: Additional options (extract_images, embed_images, skip_images)
        """
        super().__init__(input_path, output_path, **kwargs)
        self.include_notes = include_notes

    def convert(self) -> str:
        """Convert PPTX to Markdown."""
        if not PPTX_AVAILABLE:
            raise RuntimeError("python-pptx is not available for PPTX conversion")

        try:
            logger.info("Converting PPTX using python-pptx")
            presentation = Presentation(self.input_path)

            markdown_lines = []

            for slide_num, slide in enumerate(presentation.slides, 1):
                # Slide heading
                markdown_lines.append(f"## Slide {slide_num}")
                markdown_lines.append("")

                # Extract text from shapes
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        # Simple bullet detection (basic implementation)
                        text = shape.text.strip()
                        if len(text) > 0:
                            markdown_lines.append(f"- {text}")

                # Add speaker notes if requested
                if self.include_notes and slide.has_notes_slide:
                    notes_slide = slide.notes_slide
                    notes_text = notes_slide.notes_text_frame.text.strip()
                    if notes_text:
                        markdown_lines.append("")
                        markdown_lines.append("**Notes:**")
                        markdown_lines.append(f"> {notes_text}")

                markdown_lines.append("")

            return "\n".join(markdown_lines)

        except Exception as e:
            logger.error(f"Error converting PPTX: {e}")
            raise
