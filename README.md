# office2md

Transform Office Documents to Markdown with ease.

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## Overview

**office2md** is a modern Python tool for converting Microsoft Office files to Markdown format. It supports:

- **Word documents** (.docx) - Using mammoth (recommended) for high-quality conversion with python-docx as fallback
- **Excel spreadsheets** (.xlsx, .xls) - Using openpyxl
- **PowerPoint presentations** (.pptx, .ppt) - Using python-pptx

## Features

✨ **High-Quality Conversion**: Uses mammoth for superior DOCX to Markdown conversion  
📦 **Batch Processing**: Convert entire directories of files at once  
🔄 **Recursive Support**: Process nested directory structures  
📝 **Rich Logging**: Detailed logging for debugging and monitoring  
🎯 **CLI Interface**: Easy-to-use command-line interface  
🧩 **Modular Design**: Clean, extensible architecture  
✅ **Well Tested**: Comprehensive test suite included  

## Installation

### From source

```bash
git clone https://github.com/rodgui/office2md.git
cd office2md
pip install -e .
```

### For development

```bash
git clone https://github.com/rodgui/office2md.git
cd office2md
pip install -e ".[dev]"
```

## Quick Start

### Convert a single file

```bash
office2md document.docx
```

This creates `document.md` in the same directory.

### Specify output location

```bash
office2md document.docx -o output/result.md
```

### Batch convert a directory

```bash
office2md --batch ./input_dir -o ./output_dir
```

### Recursive batch conversion

```bash
office2md --batch ./input_dir -o ./output_dir --recursive
```

## Usage

### Command Line Interface

```bash
office2md [OPTIONS] INPUT
```

#### Arguments

- `INPUT` - Input file or directory (with --batch)

#### Options

- `-o, --output` - Output file or directory path
- `-b, --batch` - Enable batch mode for directory processing
- `-r, --recursive` - Process directories recursively (with --batch)
- `-v, --verbose` - Enable verbose logging
- `--no-mammoth` - Use python-docx instead of mammoth for DOCX files
- `--first-sheet-only` - Only convert first sheet of Excel files
- `--no-notes` - Exclude speaker notes from PowerPoint conversion

### Python API

```python
from office2md import ConverterFactory

# Convert a single file
converter = ConverterFactory.create_converter("document.docx")
markdown_content = converter.convert_and_save()

# Or get content without saving
markdown_content = converter.convert()

# Specify output path
converter = ConverterFactory.create_converter("document.docx", "output.md")
converter.convert_and_save()

# Use specific converters
from office2md import DocxConverter, XlsxConverter, PptxConverter

# DOCX with options
docx_converter = DocxConverter("document.docx", use_mammoth=True)
md_content = docx_converter.convert_and_save()

# XLSX with options
xlsx_converter = XlsxConverter("spreadsheet.xlsx", include_all_sheets=True)
md_content = xlsx_converter.convert_and_save()

# PPTX with options
pptx_converter = PptxConverter("presentation.pptx", include_notes=True)
md_content = pptx_converter.convert_and_save()
```

## Examples

See the [examples](examples/) directory for sample files and outputs.

To generate sample Office files:

```bash
python examples/create_samples.py
```

Then convert them:

```bash
office2md --batch examples/input -o examples/output
```

## Architecture

The project follows a modular architecture:

```
office2md/
├── converters/
│   ├── base_converter.py    # Abstract base class
│   ├── docx_converter.py    # Word document converter
│   ├── xlsx_converter.py    # Excel spreadsheet converter
│   └── pptx_converter.py    # PowerPoint converter
├── converter_factory.py     # Factory for creating converters
└── cli.py                   # Command-line interface
```

### Adding New Converters

To add support for a new file format:

1. Create a new converter class inheriting from `BaseConverter`
2. Implement the `convert()` method
3. Register the file extension in `ConverterFactory.SUPPORTED_EXTENSIONS`

```python
from office2md.converters.base_converter import BaseConverter

class MyConverter(BaseConverter):
    def convert(self) -> str:
        # Your conversion logic here
        return markdown_content
```

## Development

### Setup Development Environment

```bash
# Clone the repository
git clone https://github.com/rodgui/office2md.git
cd office2md

# Install with development dependencies
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=office2md --cov-report=html

# Run specific test file
pytest tests/test_cli.py
```

### Code Quality

```bash
# Format code with black
black office2md tests

# Lint with ruff
ruff check office2md tests

# Type checking (if mypy is installed)
mypy office2md
```

## Conversion Quality

### DOCX Files

- **Primary**: Uses [mammoth](https://github.com/mwilliamson/python-mammoth) for high-quality conversion that preserves document structure
- **Fallback**: Uses [python-docx](https://github.com/python-openxml/python-docx) if mammoth is unavailable
- Supports: headings, paragraphs, tables, basic formatting

### XLSX Files

- Uses [openpyxl](https://openpyxl.readthedocs.io/) for Excel file parsing
- Converts sheets to Markdown tables
- Handles multiple sheets
- Preserves cell values (not formulas)

### PPTX Files

- Uses [python-pptx](https://github.com/scanny/python-pptx) for PowerPoint parsing
- Extracts text from slides
- Includes speaker notes (optional)
- Handles tables in slides

## Troubleshooting

### ImportError for mammoth

Mammoth is the recommended library for DOCX conversion. If it's not installed:

```bash
pip install mammoth
```

Or use the fallback:

```bash
office2md document.docx --no-mammoth
```

### Empty output for complex files

Some complex formatting may not convert perfectly. Try:

1. Simplifying the source document
2. Using different converter options
3. Reporting issues on GitHub

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [mammoth](https://github.com/mwilliamson/python-mammoth) - High-quality DOCX to Markdown conversion
- [python-docx](https://github.com/python-openxml/python-docx) - DOCX file manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling
- [python-pptx](https://github.com/scanny/python-pptx) - PowerPoint file parsing

## Support

If you encounter any issues or have questions:

- Check the [examples](examples/) directory
- Review existing [issues](https://github.com/rodgui/office2md/issues)
- Open a new issue with detailed information
