"""Command-line interface for office2md."""

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional, Tuple

from office2md.__version__ import __version__
from office2md.converter_factory import ConverterFactory

logger = logging.getLogger(__name__)


def parse_args(args=None) -> argparse.Namespace:
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert Office documents to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Converter Priority (DOCX):
  By default, converters are tried in order: Pandoc → Mammoth → python-docx
  Use flags to force a specific converter.

Examples:
    # Auto-select best available converter
    office2md document.docx
    
    # Force specific converter
    office2md document.docx --use-pandoc      # Best tables
    office2md document.docx --use-mammoth     # Good formatting
    office2md document.docx --use-basic       # Fallback only
    
    # PDF with Docling
    office2md document.pdf --use-docling
    
    # Batch conversion
    office2md --batch ./input -o ./output --recursive
        """
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}"
    )
    
    parser.add_argument(
        "input",
        nargs="?",
        help="Input file or directory (with --batch)"
    )
    
    parser.add_argument(
        "-o", "--output",
        help="Output file or directory"
    )
    
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output"
    )
    
    # Converter selection for DOCX
    converter_group = parser.add_argument_group("DOCX Converter Selection")
    
    converter_group.add_argument(
        "--use-pandoc",
        action="store_true",
        help="Force Pandoc converter (best tables, requires external binary)"
    )
    
    converter_group.add_argument(
        "--use-mammoth",
        action="store_true",
        help="Force Mammoth converter (good formatting, skip Pandoc)"
    )
    
    converter_group.add_argument(
        "--use-basic",
        action="store_true",
        help="Force basic python-docx converter (skip Pandoc and Mammoth)"
    )
    
    converter_group.add_argument(
        "--use-docling",
        action="store_true",
        help="Use Docling converter (PDF only, ML-based)"
    )
    
    # Image options
    image_group = parser.add_argument_group("Image Options")
    
    image_group.add_argument(
        "--skip-images",
        action="store_true",
        help="Skip image extraction"
    )
    
    image_group.add_argument(
        "--images-dir",
        help="Custom directory for extracted images"
    )
    
    # Batch options
    batch_group = parser.add_argument_group("Batch Processing")
    
    batch_group.add_argument(
        "--batch",
        action="store_true",
        help="Process directory of files"
    )
    
    batch_group.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="Process subdirectories recursively"
    )
    
    # Format-specific options
    format_group = parser.add_argument_group("Format-Specific Options")
    
    format_group.add_argument(
        "--first-sheet-only",
        action="store_true",
        help="XLSX: Convert only the first sheet"
    )
    
    format_group.add_argument(
        "--no-notes",
        action="store_true",
        help="PPTX: Skip speaker notes"
    )
    
    return parser.parse_args(args)


def setup_logging(verbose: bool = False):
    """Setup logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )


def convert_file(
    input_path: str,
    output_path: Optional[str] = None,
    use_pandoc: bool = False,
    use_mammoth: bool = False,
    use_basic: bool = False,
    use_docling: bool = False,
    **kwargs
) -> bool:
    """
    Convert a single file to Markdown.

    Returns:
        True if conversion was successful, False otherwise.
    """
    try:
        input_file = Path(input_path)
        
        if not input_file.exists():
            logger.error(f"File not found: {input_path}")
            return False
        
        ext = input_file.suffix.lower()
        
        # Validate Docling usage (PDF only)
        if use_docling:
            if ext != '.pdf':
                logger.error(
                    f"Docling is optimized for PDF files only. "
                    f"For '{ext}' files, use default or --use-pandoc"
                )
                return False
            
            from office2md.converters.docling_converter import DoclingConverter
            converter = DoclingConverter(input_path, output_path, **kwargs)
        
        # DOCX with specific converter
        elif ext == '.docx':
            from office2md.converters.docx_converter import DocxConverter
            converter = DocxConverter(
                input_path,
                output_path,
                use_pandoc=use_pandoc,
                use_mammoth=use_mammoth,
                use_basic=use_basic,
                **kwargs
            )
        
        # Other formats use factory
        else:
            converter = ConverterFactory.create(input_path, output_path, **kwargs)
        
        result = converter.convert()
        converter.save(result)
        
        # Log which converter was used
        if hasattr(converter, 'converter_used') and converter.converter_used:
            logger.info(f"Converted with {converter.converter_used}: {input_path}")
        else:
            logger.info(f"Successfully converted: {input_path}")
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to convert {input_path}: {e}")
        return False


def batch_convert(
    input_dir: str,
    output_dir: Optional[str] = None,
    recursive: bool = False,
    **kwargs
) -> Tuple[int, int]:
    """
    Convert multiple files in a directory.

    Returns:
        Tuple of (success_count, failure_count)
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir) if output_dir else input_path
    
    if not input_path.is_dir():
        logger.error(f"Not a directory: {input_dir}")
        return 0, 1
    
    # Find all supported files
    pattern = "**/*" if recursive else "*"
    files = [
        f for f in input_path.glob(pattern)
        if f.is_file() and ConverterFactory.is_supported(str(f))
    ]
    
    if not files:
        logger.warning(f"No supported files found in {input_dir}")
        return 0, 0
    
    success = 0
    failure = 0
    
    for file in files:
        # Calculate output path
        if recursive:
            rel_path = file.relative_to(input_path)
            out_file = output_path / rel_path.with_suffix('.md')
        else:
            out_file = output_path / file.with_suffix('.md').name
        
        # Ensure output directory exists
        out_file.parent.mkdir(parents=True, exist_ok=True)
        
        if convert_file(str(file), str(out_file), **kwargs):
            success += 1
        else:
            failure += 1
    
    return success, failure


def main(args=None) -> int:
    """Main entry point."""
    parsed = parse_args(args)
    
    setup_logging(parsed.verbose)
    
    # Handle legacy --no-mammoth flag
    if hasattr(parsed, 'no_mammoth') and parsed.no_mammoth:
        logger.warning("--no-mammoth is deprecated, use --use-basic instead")
        parsed.use_basic = True
    
    if not parsed.input:
        logger.error("No input file specified. Use --help for usage.")
        return 1
    
    # Build kwargs
    kwargs = {
        "skip_images": parsed.skip_images,
        "extract_images": not parsed.skip_images,
    }
    
    if parsed.images_dir:
        kwargs["images_dir"] = Path(parsed.images_dir)
    
    # Format-specific options
    if parsed.first_sheet_only:
        kwargs["include_all_sheets"] = False
    
    if parsed.no_notes:
        kwargs["include_notes"] = False
    
    # Converter flags
    converter_kwargs = {
        "use_pandoc": parsed.use_pandoc,
        "use_mammoth": parsed.use_mammoth,
        "use_basic": parsed.use_basic,
        "use_docling": parsed.use_docling,
    }
    
    if parsed.batch:
        success, failure = batch_convert(
            parsed.input,
            parsed.output,
            parsed.recursive,
            **kwargs,
            **converter_kwargs
        )
        
        logger.info(f"Batch complete: {success} succeeded, {failure} failed")
        return 0 if failure == 0 else 1
    else:
        success = convert_file(
            parsed.input,
            parsed.output,
            **kwargs,
            **converter_kwargs
        )
        return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
