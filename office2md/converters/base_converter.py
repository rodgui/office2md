"""Base converter class for office2md."""

import base64
import hashlib
import logging
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class BaseConverter(ABC):
    """
    Abstract base class for all document converters.
    
    All format-specific converters inherit from this class and implement
    the convert() method.
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        extract_images: bool = True,
        skip_images: bool = False,
        images_dir: Optional[Path] = None,
        **kwargs
    ):
        """
        Initialize the converter.

        Args:
            input_path: Path to the input file.
            output_path: Optional path for the output Markdown file.
            extract_images: Whether to extract images from the document.
            skip_images: Whether to skip image processing entirely.
            images_dir: Custom directory for extracted images.
            **kwargs: Additional format-specific options.
        """
        self.input_path = Path(input_path)
        
        if output_path:
            self.output_path = Path(output_path)
        else:
            self.output_path = self.input_path.with_suffix('.md')
        
        self.extract_images = extract_images and not skip_images
        self.skip_images = skip_images
        
        # Set images directory
        if images_dir:
            self.images_dir = Path(images_dir)
        else:
            # Default: {output_name}_images/
            self.images_dir = self.output_path.parent / f"{self.output_path.stem}_images"
        
        self._image_counter = 0
        self._kwargs = kwargs

    @abstractmethod
    def convert(self) -> str:
        """
        Convert the document to Markdown.

        Returns:
            Markdown formatted string.
        """
        pass

    def save(self, content: str) -> None:
        """
        Save the converted content to the output file.

        Args:
            content: Markdown string to save.
        """
        if not self.output_path:
            raise ValueError("No output path specified for saving.")
        
        # Ensure parent directory exists
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(self.output_path, "w", encoding="utf-8") as f:
            f.write(content)
        
        logger.info(f"Saved output to: {self.output_path}")

    def convert_and_save(self) -> None:
        """Convert and save in one operation."""
        self.save()
        if self.extract_images and self.extracted_images:
            logger.info(
                f"Extracted {len(self.extracted_images)} images to {self.images_dir}"
            )

    def _process_image(self, image_data: bytes, extension: str = "png") -> str:
        """
        Process and save an image, returning a Markdown reference.

        Args:
            image_data: Raw image bytes.
            extension: Image file extension (png, jpg, etc.).

        Returns:
            Markdown image reference string.
        """
        if self.skip_images or not self.extract_images:
            return ""
        
        self._image_counter += 1
        
        # Ensure images directory exists
        self.images_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate filename
        image_filename = f"image_{self._image_counter}.{extension}"
        image_path = self.images_dir / image_filename
        
        # Save image
        with open(image_path, "wb") as f:
            f.write(image_data)
        
        # Return relative markdown reference
        try:
            rel_path = image_path.relative_to(self.output_path.parent)
        except ValueError:
            rel_path = image_path
        
        return f"![](./{rel_path})"

    def _generate_image_hash(self, image_data: bytes) -> str:
        """Generate a hash for image data to detect duplicates."""
        return hashlib.md5(image_data).hexdigest()[:8]

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
