"""Abstract base class for all format converters."""

import base64
import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class BaseConverter(ABC):
    """Abstract base class for all format converters."""

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        extract_images: bool = True,
        embed_images: bool = False,
        skip_images: bool = False,
        **kwargs
    ):
        """
        Initialize converter with image handling modes.

        Args:
            input_path: Path to input file
            output_path: Optional output path (defaults to input with .md extension)
            extract_images: Extract images to subdirectory (default: True)
            embed_images: Embed images as base64 (overrides extract_images)
            skip_images: Skip images entirely
            **kwargs: Additional converter-specific options
        """
        self.input_path = Path(input_path)
        self.output_path = (
            Path(output_path) if output_path else self._default_output_path()
        )

        # Image handling modes (mutually exclusive)
        self.skip_images = skip_images
        self.embed_images = embed_images
        self.extract_images = (
            extract_images and not embed_images and not skip_images
        )

        # Image directory for extraction mode
        if self.extract_images:
            self.images_dir = (
                self.output_path.parent / f"{self.output_path.stem}_images"
            )
        else:
            self.images_dir = None

        self.extracted_images = {}
        self.kwargs = kwargs

    def _default_output_path(self) -> Path:
        """Generate default output path (same dir, .md extension)."""
        return self.input_path.with_suffix(".md")

    @abstractmethod
    def convert(self) -> str:
        """
        Convert file to Markdown string.

        Returns:
            Markdown string representation of the file
        """
        pass

    def save(self) -> None:
        """Save converted Markdown to disk."""
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.output_path, "w", encoding="utf-8") as f:
            f.write(self.convert())
        logger.info(f"Saved: {self.output_path}")

    def convert_and_save(self) -> None:
        """Convert and save in one operation."""
        self.save()
        if self.extract_images and self.extracted_images:
            logger.info(
                f"Extracted {len(self.extracted_images)} images to {self.images_dir}"
            )

    def _process_image(self, image_data: bytes, image_format: str = "png") -> str:
        """
        Process image based on mode (extract/embed/skip).

        Args:
            image_data: Raw image bytes
            image_format: Image format (png, jpg, gif, etc.)

        Returns:
            Markdown image reference
        """
        if self.skip_images:
            return ""

        if self.embed_images:
            b64_data = base64.b64encode(image_data).decode("utf-8")
            return f"![](data:image/{image_format};base64,{b64_data})"

        if self.extract_images:
            self.images_dir.mkdir(parents=True, exist_ok=True)

            image_count = len(self.extracted_images) + 1
            filename = f"image_{image_count}.{image_format}"
            image_path = self.images_dir / filename

            with open(image_path, "wb") as f:
                f.write(image_data)

            image_key = hash(image_data)
            self.extracted_images[image_key] = filename

            relative_path = f"{self.output_path.stem}_images/{filename}"
            return f"![](./{relative_path})"

        return ""

    def _replace_base64_images(self, markdown: str) -> str:
        """
        Replace base64 images in markdown based on mode.

        Args:
            markdown: Markdown content potentially containing base64 images

        Returns:
            Processed markdown with images handled per mode
        """
        if self.skip_images:
            return re.sub(r"!\[.*?\]\(.*?\)", "", markdown)

        if self.embed_images:
            logger.info("Keeping images as base64 inline")
            return markdown

        if self.extract_images:
            # IMPROVED: More flexible pattern for base64 detection
            # Handles: png, jpg, jpeg, gif, svg+xml, webp with optional whitespace
            pattern = r"!\[([^\]]*)\]\(data:image/([a-zA-Z0-9+]+);base64,([A-Za-z0-9+/=\s]+)\)"

            def replace_func(match):
                alt_text = match.group(1)
                image_format = match.group(2).lower().replace("+xml", "")  # svg+xml -> svg
                b64_data = match.group(3).replace(" ", "").replace("\n", "").replace("\r", "")

                try:
                    image_data = base64.b64decode(b64_data)
                    ref = self._process_image(image_data, image_format)
                    logger.debug(f"Extracted image: format={image_format}, size={len(image_data)} bytes, alt='{alt_text}'")
                    return ref
                except Exception as e:
                    logger.warning(f"Failed to decode base64 image: {e}")
                    return ""

            processed = re.sub(pattern, replace_func, markdown)
            
            # Log extraction summary
            if self.extracted_images:
                logger.info(f"Successfully extracted {len(self.extracted_images)} images")
            else:
                logger.warning("No images extracted - check if markdown contains base64 images")
            
            return processed

        return markdown
