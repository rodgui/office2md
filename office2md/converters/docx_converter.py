"""
DOCX to Markdown converter with intelligent fallback chain.

Converter Priority:
1. Pandoc (best tables, requires external binary)
2. Mammoth (good formatting, pure Python)
3. python-docx (basic, always available)
"""

import logging
import shutil
from pathlib import Path
from typing import Optional

from office2md.converters.base_converter import BaseConverter

logger = logging.getLogger(__name__)

# Check availability of converters
PANDOC_AVAILABLE = shutil.which('pandoc') is not None

try:
    import mammoth
    MAMMOTH_AVAILABLE = True
except ImportError:
    MAMMOTH_AVAILABLE = False

try:
    import docx
    PYTHON_DOCX_AVAILABLE = True
except ImportError:
    PYTHON_DOCX_AVAILABLE = False


class DocxConverter(BaseConverter):
    """
    DOCX to Markdown converter with automatic fallback chain.
    
    Converter priority (best to basic):
    1. **Pandoc**: Best for tables, requires `brew install pandoc`
    2. **Mammoth**: Good formatting, pure Python
    3. **python-docx**: Basic extraction, always available
    
    Use flags to force a specific converter:
    - `use_pandoc=True`: Force Pandoc (error if unavailable)
    - `use_mammoth=True`: Force Mammoth, skip Pandoc
    - `use_basic=True`: Force python-docx only (skip Pandoc and Mammoth)
    """

    def __init__(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        use_pandoc: bool = False,
        use_mammoth: bool = False,
        use_basic: bool = False,
        **kwargs
    ):
        """
        Initialize DOCX converter.
        
        Args:
            input_path: Path to DOCX file
            output_path: Optional output path for Markdown
            use_pandoc: Force Pandoc converter
            use_mammoth: Force Mammoth converter (skip Pandoc)
            use_basic: Force python-docx only (skip Pandoc and Mammoth)
            **kwargs: Additional options passed to base converter
        """
        super().__init__(input_path, output_path, **kwargs)
        
        self.use_pandoc = use_pandoc
        self.use_mammoth = use_mammoth
        self.use_basic = use_basic
        
        # Store converter used for logging
        self._converter_used = None

    def convert(self) -> str:
        """
        Convert DOCX to Markdown using the best available converter.
        
        Returns:
            Markdown formatted string
        """
        # Determine which converter to use
        converter_func = self._select_converter()
        
        # Run conversion
        result = converter_func()
        
        logger.info(f"Conversion completed using: {self._converter_used}")
        return result

    def _select_converter(self):
        """Select the best available converter based on flags and availability."""
        
        # Force basic (python-docx only)
        if self.use_basic:
            if not PYTHON_DOCX_AVAILABLE:
                raise RuntimeError("python-docx not available. Install with: pip install python-docx")
            self._converter_used = "python-docx"
            return self._convert_with_python_docx
        
        # Force Mammoth (skip Pandoc)
        if self.use_mammoth:
            if not MAMMOTH_AVAILABLE:
                raise RuntimeError("Mammoth not available. Install with: pip install mammoth")
            self._converter_used = "mammoth"
            return self._convert_with_mammoth
        
        # Force Pandoc
        if self.use_pandoc:
            if not PANDOC_AVAILABLE:
                raise RuntimeError(
                    "Pandoc not available. Install with:\n"
                    "  macOS: brew install pandoc\n"
                    "  Ubuntu: sudo apt-get install pandoc\n"
                    "  Windows: choco install pandoc"
                )
            self._converter_used = "pandoc"
            return self._convert_with_pandoc
        
        # Auto-select: try in order of quality
        if PANDOC_AVAILABLE:
            self._converter_used = "pandoc"
            return self._convert_with_pandoc
        
        if MAMMOTH_AVAILABLE:
            logger.info("Pandoc not available, using Mammoth")
            self._converter_used = "mammoth"
            return self._convert_with_mammoth
        
        if PYTHON_DOCX_AVAILABLE:
            logger.warning("Using basic python-docx converter (limited formatting)")
            self._converter_used = "python-docx"
            return self._convert_with_python_docx
        
        raise RuntimeError(
            "No DOCX converter available. Install one of:\n"
            "  - Pandoc: brew install pandoc (recommended)\n"
            "  - Mammoth: pip install mammoth\n"
            "  - python-docx: pip install python-docx"
        )

    def _convert_with_pandoc(self) -> str:
        """Convert using Pandoc (best quality)."""
        from office2md.converters.pandoc_converter import PandocConverter
        
        converter = PandocConverter(
            str(self.input_path),
            str(self.output_path) if self.output_path else None,
            extract_images=self.extract_images,
            skip_images=self.skip_images,
            images_dir=self.images_dir,
        )
        return converter.convert()

    def _convert_with_mammoth(self) -> str:
        """Convert using Mammoth (good quality)."""
        from office2md.converters.mammoth_converter import MammothConverter
        
        converter = MammothConverter(
            str(self.input_path),
            str(self.output_path) if self.output_path else None,
            extract_images=self.extract_images,
            skip_images=self.skip_images,
            images_dir=self.images_dir,
        )
        return converter.convert()

    def _convert_with_python_docx(self) -> str:
        """Convert using python-docx (basic)."""
        from office2md.converters.basic_docx_converter import BasicDocxConverter
        
        converter = BasicDocxConverter(
            str(self.input_path),
            str(self.output_path) if self.output_path else None,
            extract_images=self.extract_images,
            skip_images=self.skip_images,
            images_dir=self.images_dir,
        )
        return converter.convert()

    @property
    def converter_used(self) -> Optional[str]:
        """Return the name of the converter that was used."""
        return self._converter_used
