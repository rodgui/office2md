# office2md

Convert Microsoft Office files (DOCX, XLSX, PPTX) to Markdown format.

## Features

- **DOCX → Markdown**: Full support for headings, lists, tables, images, and formatting
- **XLSX → Markdown**: Convert spreadsheets to Markdown tables
- **PPTX → Markdown**: Extract slides with optional speaker notes
- **Image Extraction**: Automatically extract and reference images
- **Multiple Converters**: Choose between mammoth, Pandoc, or Docling

## Installation

```bash
pip install office2md
```

### Optional Dependencies

For Pandoc support (recommended for complex documents):
```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt-get install pandoc

# Windows
choco install pandoc
```

For Docling support (ML-based, supports PDF):
```bash
pip install docling
```

## Quick Start

### Single File Conversion

```bash
# Basic conversion (auto-generates output.md)
office2md document.docx

# Specify output file
office2md document.docx -o output.md

# Use Pandoc for better quality
office2md document.docx --use-pandoc

# Use Docling for complex documents or PDFs
office2md document.pdf --use-docling
```

### Batch Conversion

```bash
# Convert all files in directory
office2md --batch ./input -o ./output

# Recursive conversion (preserves directory structure)
office2md --batch ./input -o ./output --recursive
```

## Converter Options

### DOCX Converters

| Converter             | Flag            | Pros                  | Cons                      |
| --------------------- | --------------- | --------------------- | ------------------------- |
| **Mammoth** (default) | -               | Good formatting, fast | Tables can be imperfect   |
| **Python-docx**       | `--no-mammoth`  | Direct parsing        | Less formatting           |
| **Pandoc**            | `--use-pandoc`  | Best quality          | Requires external binary  |
| **Docling**           | `--use-docling` | ML-based, handles PDF | Slower, more dependencies |

### Image Handling

```bash
# Extract images to subdirectory (default)
office2md document.docx

# Embed images as base64 in markdown
office2md document.docx --embed-images

# Skip images entirely
office2md document.docx --skip-images
```

### Format-Specific Options

```bash
# XLSX: Convert only first sheet
office2md data.xlsx --first-sheet-only

# PPTX: Exclude speaker notes
office2md slides.pptx --no-notes
```

## Examples

### Basic DOCX Conversion

```bash
office2md report.docx -o report.md -v
```

Output structure:
```
report.md
report_images/
  ├── image_1.png
  ├── image_2.png
  └── image_3.png
```

### High-Quality Conversion with Pandoc

```bash
office2md complex_document.docx --use-pandoc -o output.md
```

### PDF Conversion with Docling

```bash
office2md scanned_document.pdf --use-docling -o output.md
```

### Batch Processing

```bash
# Convert entire project documentation
office2md --batch ./docs -o ./markdown --recursive -v
```

## Python API

```python
from office2md import ConverterFactory

# Create converter
converter = ConverterFactory.create_converter(
    "document.docx",
    "output.md",
    extract_images=True
)

# Convert and save
converter.convert_and_save()

# Or just get markdown string
markdown = converter.convert()
print(markdown)
```

### Using Specific Converters

```python
from office2md.converters.pandoc_converter import PandocConverter
from office2md.converters.docling_converter import DoclingConverter

# Pandoc (requires pandoc installed)
converter = PandocConverter("document.docx", "output.md")
markdown = converter.convert()

# Docling (requires pip install docling)
converter = DoclingConverter("document.pdf", "output.md")
markdown = converter.convert()
```

## Supported Formats

| Format     | Extension       | Default Converter     |
| ---------- | --------------- | --------------------- |
| Word       | `.docx`         | Mammoth + python-docx |
| Excel      | `.xlsx`, `.xls` | openpyxl              |
| PowerPoint | `.pptx`         | python-pptx           |
| PDF        | `.pdf`          | Docling (optional)    |

## Troubleshooting

### Images Not Extracting

Ensure you're not using `--skip-images` or `--embed-images`:
```bash
office2md document.docx -v
```

### Tables Not Rendering Correctly

Try using Pandoc for better table support:
```bash
office2md document.docx --use-pandoc
```

### Complex Layouts

For documents with complex layouts, use Docling:
```bash
pip install docling
office2md document.docx --use-docling
```

### Verbose Logging

Enable verbose mode to see detailed conversion logs:
```bash
office2md document.docx -v
```

## Development

### Setup

```bash
git clone https://github.com/yourusername/office2md.git
cd office2md
pip install -e ".[dev]"
```

### Running Tests

```bash
# All tests
pytest

# With coverage
pytest --cov=office2md

# Specific test file
pytest tests/converters/test_docx_converter.py -v
```

### Project Structure

```
office2md/
├── __init__.py
├── cli.py                      # Command-line interface
├── converter_factory.py        # Factory for creating converters
└── converters/
    ├── base_converter.py       # Abstract base class
    ├── docx_converter.py       # DOCX converter (mammoth/python-docx)
    ├── xlsx_converter.py       # Excel converter
    ├── pptx_converter.py       # PowerPoint converter
    ├── pandoc_converter.py     # Pandoc-based converter
    └── docling_converter.py    # Docling-based converter
```

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.
