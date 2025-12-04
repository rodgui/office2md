"""Command-line interface for office2md."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List

from office2md.converter_factory import ConverterFactory


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
        **kwargs: Additional converter options

    Returns:
        True if successful, False otherwise
    """
    logger = logging.getLogger(__name__)

    try:
        if not ConverterFactory.is_supported(input_path):
            logger.error(f"Unsupported file type: {input_path}")
            return False

        # Filter kwargs based on file type
        file_path = Path(input_path)
        extension = file_path.suffix.lower()

        filtered_kwargs = {}
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
) -> tuple:
    """
    Convert multiple files in a directory.

    Args:
        input_dir: Directory containing input files
        output_dir: Optional directory for output files
        recursive: If True, process subdirectories recursively
        **kwargs: Additional converter options

    Returns:
        Tuple of (success_count, failure_count)
    """
    logger = logging.getLogger(__name__)
    input_path = Path(input_dir)

    if not input_path.is_dir():
        logger.error(f"Input directory not found: {input_dir}")
        return 0, 0

    # Find all supported files
    pattern = "**/*" if recursive else "*"
    all_files = input_path.glob(pattern)
    supported_files = [f for f in all_files if ConverterFactory.is_supported(str(f))]

    if not supported_files:
        logger.warning(f"No supported files found in: {input_dir}")
        return 0, 0

    logger.info(f"Found {len(supported_files)} files to convert")

    success_count = 0
    failure_count = 0

    for file_path in supported_files:
        # Calculate output path
        if output_dir:
            output_path_obj = Path(output_dir)
            if recursive:
                # Preserve directory structure
                rel_path = file_path.relative_to(input_path)
                output_path_obj = output_path_obj / rel_path.parent
            output_path_obj.mkdir(parents=True, exist_ok=True)
            output_file = str(output_path_obj / file_path.with_suffix(".md").name)
        else:
            output_file = None

        if convert_file(str(file_path), output_file, **kwargs):
            success_count += 1
        else:
            failure_count += 1

    logger.info(
        f"Batch conversion complete: {success_count} successful, {failure_count} failed"
    )
    return success_count, failure_count


def main(argv: List[str] = None) -> int:
    """
    Main entry point for the CLI.

    Args:
        argv: Command line arguments (defaults to sys.argv[1:])

    Returns:
        Exit code (0 for success, non-zero for failure)
    """
    parser = argparse.ArgumentParser(
        description="Convert Office files (docx, xlsx, pptx) to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert a single file
  office2md document.docx

  # Convert with custom output path
  office2md document.docx -o output.md

  # Batch convert all files in a directory
  office2md --batch ./input_dir -o ./output_dir

  # Recursive batch conversion
  office2md --batch ./input_dir -o ./output_dir --recursive

  # Enable verbose logging
  office2md document.docx -v
        """,
    )

    parser.add_argument(
        "input",
        help="Input file or directory (with --batch)",
    )

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

    # Converter-specific options
    parser.add_argument(
        "--no-mammoth",
        action="store_true",
        help="Don't use mammoth for DOCX conversion (use python-docx instead)",
    )

    parser.add_argument(
        "--first-sheet-only",
        action="store_true",
        help="Only convert first sheet of XLSX files",
    )

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
        logging.error(f"Unexpected error: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
