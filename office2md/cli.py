"""Command-line interface for office2md."""

import argparse
import logging
import sys
from pathlib import Path
from typing import Tuple

from office2md.converter_factory import ConverterFactory

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """Configure logging based on verbosity."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def parse_args(args=None) -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        prog="office2md",
        description="Convert Microsoft Office files to Markdown",
        epilog="Examples:\n"
               "  office2md document.docx\n"
               "  office2md document.docx -o output.md\n"
               "  office2md --batch ./input -o ./output\n"
               "  office2md document.docx --use-pandoc\n"
               "  office2md document.pdf --use-docling\n",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # Positional argument
    parser.add_argument(
        "input",
        nargs="?",
        help="Input file or directory (with --batch)",
    )

    # Output options
    parser.add_argument(
        "-o", "--output",
        help="Output file or directory",
    )

    # Batch mode
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Batch convert all supported files in directory",
    )
    parser.add_argument(
        "--recursive", "-r",
        action="store_true",
        help="Recursively process subdirectories (with --batch)",
    )

    # Converter selection
    converter_group = parser.add_argument_group("Converter Options")
    converter_group.add_argument(
        "--use-pandoc",
        action="store_true",
        help="Use Pandoc for DOCX conversion (requires pandoc installed)",
    )
    converter_group.add_argument(
        "--use-docling",
        action="store_true",
        help="Use Docling for conversion (requires docling installed, supports PDF)",
    )
    converter_group.add_argument(
        "--no-mammoth",
        action="store_true",
        help="Use python-docx instead of mammoth for DOCX",
    )

    # Image handling
    image_group = parser.add_argument_group("Image Options")
    image_group.add_argument(
        "--embed-images",
        action="store_true",
        help="Embed images as base64 in markdown",
    )
    image_group.add_argument(
        "--skip-images",
        action="store_true",
        help="Skip image extraction entirely",
    )

    # Format-specific options
    format_group = parser.add_argument_group("Format-Specific Options")
    format_group.add_argument(
        "--first-sheet-only",
        action="store_true",
        help="XLSX: Convert only the first sheet",
    )
    format_group.add_argument(
        "--no-notes",
        action="store_true",
        help="PPTX: Exclude speaker notes",
    )

    # Verbosity
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 0.1.0",
    )

    return parser.parse_args(args)


def convert_file(
    input_path: str,
    output_path: str = None,
    use_pandoc: bool = False,
    use_docling: bool = False,
    **kwargs
) -> bool:
    """
    Convert a single file to Markdown.

    Args:
        input_path: Path to input file
        output_path: Optional output path
        use_pandoc: Use Pandoc converter
        use_docling: Use Docling converter
        **kwargs: Additional converter options

    Returns:
        True if successful, False otherwise
    """
    try:
        input_file = Path(input_path)
        
        if not input_file.exists():
            logger.error(f"File not found: {input_path}")
            return False

        # Select converter based on options
        if use_pandoc:
            from office2md.converters.pandoc_converter import PandocConverter, PANDOC_AVAILABLE
            if not PANDOC_AVAILABLE:
                logger.error("Pandoc is not installed. Install with: brew install pandoc")
                return False
            converter = PandocConverter(input_path, output_path, **kwargs)
            
        elif use_docling:
            from office2md.converters.docling_converter import DoclingConverter, DOCLING_AVAILABLE
            if not DOCLING_AVAILABLE:
                logger.error("Docling is not installed. Install with: pip install docling")
                return False
            converter = DoclingConverter(input_path, output_path, **kwargs)
            
        else:
            # Use factory for standard conversion
            converter = ConverterFactory.create_converter(
                input_path, output_path, **kwargs
            )

        converter.convert_and_save()
        logger.info(f"Successfully converted: {input_path}")
        return True

    except Exception as e:
        logger.error(f"Failed to convert {input_path}: {e}")
        return False


def batch_convert(
    input_dir: str,
    output_dir: str = None,
    recursive: bool = False,
    **kwargs
) -> Tuple[int, int]:
    """
    Batch convert all supported files in a directory.

    Args:
        input_dir: Input directory path
        output_dir: Optional output directory
        recursive: Process subdirectories recursively
        **kwargs: Converter options

    Returns:
        Tuple of (success_count, failure_count)
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir) if output_dir else input_path

    if not input_path.exists():
        logger.error(f"Directory not found: {input_dir}")
        return 0, 1

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
        logger.warning(f"No supported files found in: {input_dir}")
        return 0, 0

    logger.info(f"Found {len(files)} files to convert")

    success_count = 0
    failure_count = 0

    for file in files:
        # Calculate output path preserving directory structure
        if recursive:
            relative = file.relative_to(input_path)
            out_file = output_path / relative.with_suffix(".md")
        else:
            out_file = output_path / file.with_suffix(".md").name

        # Ensure output directory exists
        out_file.parent.mkdir(parents=True, exist_ok=True)

        if convert_file(str(file), str(out_file), **kwargs):
            success_count += 1
        else:
            failure_count += 1

    return success_count, failure_count


def main(args=None) -> int:
    """Main entry point for CLI."""
    parsed_args = parse_args(args)
    setup_logging(parsed_args.verbose)

    if not parsed_args.input:
        logger.error("No input file or directory specified")
        return 1

    # Validate mutually exclusive converter options
    if parsed_args.use_pandoc and parsed_args.use_docling:
        logger.error("Cannot use --use-pandoc and --use-docling together. Choose one.")
        return 1

    # Build kwargs for converters
    kwargs = {
        "use_mammoth": not parsed_args.no_mammoth,
        "include_all_sheets": not parsed_args.first_sheet_only,
        "include_notes": not parsed_args.no_notes,
        "embed_images": parsed_args.embed_images,
        "skip_images": parsed_args.skip_images,
        "extract_images": not parsed_args.embed_images and not parsed_args.skip_images,
    }

    if parsed_args.batch:
        success, failure = batch_convert(
            parsed_args.input,
            parsed_args.output,
            recursive=parsed_args.recursive,
            use_pandoc=parsed_args.use_pandoc,
            use_docling=parsed_args.use_docling,
            **kwargs
        )
        logger.info(f"Batch complete: {success} succeeded, {failure} failed")
        return 0 if failure == 0 else 1
    else:
        success = convert_file(
            parsed_args.input,
            parsed_args.output,
            use_pandoc=parsed_args.use_pandoc,
            use_docling=parsed_args.use_docling,
            **kwargs
        )
        return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
