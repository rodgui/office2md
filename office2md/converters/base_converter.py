"""Base converter class for all Office file converters."""

import logging
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional


class BaseConverter(ABC):
    """Abstract base class for Office file converters."""

    def __init__(self, input_path: str, output_path: Optional[str] = None):
        """
        Initialize the converter.

        Args:
            input_path: Path to the input Office file
            output_path: Optional path for the output Markdown file.
                        If not provided, will use input filename with .md extension
        """
        self.input_path = Path(input_path)
        if not self.input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")

        if output_path:
            self.output_path = Path(output_path)
        else:
            self.output_path = self.input_path.with_suffix(".md")

        self.logger = logging.getLogger(self.__class__.__name__)

    @abstractmethod
    def convert(self) -> str:
        """
        Convert the Office file to Markdown format.

        Returns:
            The Markdown content as a string
        """
        pass

    def save(self, markdown_content: str) -> None:
        """
        Save the Markdown content to a file.

        Args:
            markdown_content: The Markdown content to save
        """
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.output_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)
        self.logger.info(f"Saved Markdown to: {self.output_path}")

    def convert_and_save(self) -> str:
        """
        Convert the file and save the result.

        Returns:
            The Markdown content as a string
        """
        self.logger.info(f"Converting {self.input_path}")
        markdown_content = self.convert()
        self.save(markdown_content)
        return markdown_content
