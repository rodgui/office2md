"""Command-line interface for office2md."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Tuple

from office2md import __version__
from office2md.converter_factory import ConverterFactory

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """
    Setup logging configuration.

    Args:
        verbose: If True, set logging level to DEBUG
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def convert_file(input_path: str, output_path: str = None, **kwargs) -> bool:
    """
    Convert a single file.

    Args:
        input_path: Path to input file
        output_path: Optional path for output file
        **kwargs: Additional converter options (image handling, format-specific)

    Returns:
        True if successful, False otherwise
    """
    try:
        if not ConverterFactory.is_supported(input_path):
            logger.error(f"Unsupported file type: {input_path}")
            return False

        # Filter kwargs based on file type (per pattern in instructions)
        file_path = Path(input_path)
        extension = file_path.suffix.lower()

        filtered_kwargs = {}

        # Image handling (applies to all formats)
        if "extract_images" in kwargs:
            filtered_kwargs["extract_images"] = kwargs["extract_images"]
        if "embed_images" in kwargs:
            filtered_kwargs["embed_images"] = kwargs["embed_images"]
        if "skip_images" in kwargs:
            filtered_kwargs["skip_images"] = kwargs["skip_images"]

        if extension == ".docx":
            if "use_mammoth" in kwargs:
                filtered_kwargs["use_mammoth"] = kwargs["use_mammoth"]

        elif extension in [".xlsx", ".xls"]:
            if "include_all_sheets" in kwargs:
                filtered_kwargs["include_all_sheets"] = kwargs["include_all_sheets"]

        elif extension in [".pptx", ".ppt"]:
            if "include_notes" in kwargs:
                filtered_kwargs["include_notes"] = kwargs["include_notes"]

        converter = ConverterFactory.create_converter(
            input_path, output_path, **filtered_kwargs
        )
        converter.convert_and_save()
        logger.info(f"Successfully converted: {input_path}")
        return True

    except Exception as e:
        logger.error(f"Error converting {input_path}: {e}", exc_info=True)
        return False


def batch_convert(
    input_dir: str, output_dir: str = None, recursive: bool = False, **kwargs
) -> Tuple[int, int]:
    """
    Convert all supported files in a directory.

    Args:
        input_dir: Input directory path
        output_dir: Output directory path (defaults to input_dir)
        recursive: If True, process subdirectories recursively
        **kwargs: Additional converter options

    Returns:
        Tuple of (success_count, failure_count)
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir) if output_dir else input_path

    if not input_path.is_dir():
        logger.error(f"Input path is not a directory: {input_dir}")
        return 0, 1

    # Find all files to convert
    glob_pattern = "**/*" if recursive else "*"
    files_to_convert = [
        f for f in input_path.glob(glob_pattern)
        if f.is_file() and ConverterFactory.is_supported(f)
    ]

    if not files_to_convert:
        logger.warning(f"No supported files found in {input_dir}")
        return 0, 0

    success_count = 0
    failure_count = 0

    for file_path in files_to_convert:
        # Preserve directory structure if recursive
        if recursive:
            rel_path = file_path.relative_to(input_path)
            output_file = output_path / rel_path.with_suffix(".md")
        else:
            output_file = output_path / file_path.with_suffix(".md").name

        if convert_file(str(file_path), str(output_file), **kwargs):
            success_count += 1
        else:
            failure_count += 1

    logger.info(
        f"Batch conversion complete: {success_count} succeeded, {failure_count} failed"
    )
    return success_count, failure_count


def main(argv: List[str] = None) -> int:
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="Convert Office files (docx, xlsx, pptx) to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Default: extract images to subdirectory
  office2md document.docx

  # Embed images as base64
  office2md document.docx --embed-images

  # Skip images entirely
  office2md document.docx --skip-images

  # Batch with recursive structure preservation
  office2md --batch ./input -o ./output --recursive
        """,
    )

    parser.add_argument("input", help="Input file or directory (with --batch)")
    parser.add_argument(
        "-o",
        "--output",
        help="Output file or directory (with --batch). "
        "If not specified, uses input filename with .md extension",
    )
    parser.add_argument(
        "-b",
        "--batch",
        action="store_true",
        help="Batch mode: convert all supported files in input directory",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Process directories recursively (only with --batch)",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"office2md {__version__}",
        help="Show version and exit",
    )

    # Image handling options
    parser.add_argument(
        "--embed-images",
        action="store_true",
        help="Embed images as base64 in markdown (default: extract to images/)",
    )
    parser.add_argument(
        "--skip-images",
        action="store_true",
        help="Skip/remove images entirely",
    )

    # DOCX options
    parser.add_argument(
        "--no-mammoth",
        action="store_true",
        help="Don't use mammoth for DOCX conversion (use python-docx instead)",
    )

    # XLSX options
    parser.add_argument(
        "--first-sheet-only",
        action="store_true",
        help="Only convert first sheet of XLSX files",
    )

    # PPTX options
    parser.add_argument(
        "--no-notes",
        action="store_true",
        help="Don't include speaker notes in PPTX conversion",
    )

    args = parser.parse_args(argv)
    setup_logging(args.verbose)

    # Prepare converter options
    converter_kwargs = {
        "use_mammoth": not args.no_mammoth,
        "include_all_sheets": not args.first_sheet_only,
        "include_notes": not args.no_notes,
        "embed_images": args.embed_images,
        "skip_images": args.skip_images,
        "extract_images": not args.embed_images and not args.skip_images,
    }

    try:
        if args.batch:
            success, failure = batch_convert(
                args.input, args.output, args.recursive, **converter_kwargs
            )
            return 0 if failure == 0 else 1
        else:
            success = convert_file(args.input, args.output, **converter_kwargs)
            return 0 if success else 1

    except KeyboardInterrupt:
        print("\nConversion interrupted by user", file=sys.stderr)
        return 130
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
