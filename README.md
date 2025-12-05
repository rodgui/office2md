# office2md

Convert Microsoft Office documents to Markdown with intelligent converter selection.

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **DOCX** → Markdown with images, tables, and formatting
- **XLSX** → Markdown tables (single or all sheets)
- **PPTX** → Markdown with slide content and speaker notes
- **PDF** → Markdown using ML-based extraction (via Docling)
- **Batch processing** with recursive directory support
- **Automatic image extraction** to separate folder

## Installation

```bash
pip install office2md
```

### Optional Dependencies

For best results with DOCX files, install Pandoc:

```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt-get install pandoc

# Windows
choco install pandoc
```

For PDF support:

```bash
pip install docling
```

## Quick Start

```bash
# Convert a single file (auto-selects best converter)
office2md document.docx

# Specify output path
office2md document.docx -o output/result.md

# Convert all files in a directory
office2md --batch ./input -o ./output

# Recursive batch processing
office2md --batch ./input -o ./output --recursive
```

## Converter Selection

### DOCX Converters (Priority Order)

| Converter       | Quality  | Tables      | Requirements    |
| --------------- | -------- | ----------- | --------------- |
| **Pandoc**      | ⭐⭐⭐ Best | ✅ Excellent | External binary |
| **Mammoth**     | ⭐⭐ Good  | ⚠️ Partial   | Pure Python     |
| **python-docx** | ⭐ Basic  | ✅ Basic     | Pure Python     |

By default, office2md automatically selects the best available converter:

```
Pandoc (if installed) → Mammoth → python-docx
```

### Force a Specific Converter

```bash
# Force Pandoc (best for complex tables)
office2md document.docx --use-pandoc

# Force Mammoth (skip Pandoc)
office2md document.docx --use-mammoth

# Force basic python-docx only
office2md document.docx --use-basic
```

### PDF Converter

```bash
# Use Docling for PDF (ML-based, excellent for scanned docs)
office2md document.pdf --use-docling
```

> **Note**: Docling is optimized for PDF files only. For DOCX, use the default converter or Pandoc.

## CLI Reference

```
office2md [OPTIONS] INPUT

Arguments:
  INPUT                     Input file or directory (with --batch)

Options:
  -o, --output PATH         Output file or directory
  -v, --verbose             Enable verbose output

DOCX Converter Selection:
  --use-pandoc              Force Pandoc (best tables)
  --use-mammoth             Force Mammoth (good formatting)
  --use-basic               Force python-docx (fallback)
  --use-docling             Use Docling (PDF only)

Image Options:
  --skip-images             Skip image extraction
  --images-dir PATH         Custom directory for images

Batch Processing:
  --batch                   Process directory of files
  -r, --recursive           Process subdirectories

Format-Specific Options:
  --first-sheet-only        XLSX: Convert only first sheet
  --no-notes                PPTX: Skip speaker notes
```

## Examples

### Single File Conversion

```bash
# Auto-select converter
office2md report.docx

# With custom output and image directory
office2md report.docx -o docs/report.md --images-dir docs/images

# Skip images
office2md report.docx --skip-images

# Verbose output to see which converter is used
office2md report.docx -v
```

### Batch Conversion

```bash
# Convert all supported files in directory
office2md --batch ./documents -o ./markdown

# Recursive with structure preserved
office2md --batch ./documents -o ./markdown --recursive

# Force Pandoc for all DOCX in batch
office2md --batch ./documents -o ./markdown --use-pandoc
```

### Format-Specific

```bash
# Excel: first sheet only
office2md data.xlsx --first-sheet-only

# PowerPoint: without speaker notes
office2md presentation.pptx --no-notes

# PDF with Docling
office2md scanned-document.pdf --use-docling
```

## Python API

```python
from office2md import ConverterFactory

# Auto-select converter
converter = ConverterFactory.create("document.docx")
markdown = converter.convert()
converter.save(markdown)

# With options
converter = ConverterFactory.create(
    "document.docx",
    output_path="output.md",
    extract_images=True,
    images_dir="./images"
)
markdown = converter.convert()
converter.save(markdown)
```

### Using Specific Converters

```python
from office2md.converters import DocxConverter, PandocConverter

# DOCX with automatic fallback
converter = DocxConverter(
    "document.docx",
    use_pandoc=True  # Force Pandoc, error if unavailable
)

# Or use Pandoc directly
converter = PandocConverter("document.docx")
markdown = converter.convert()
```

### Check Available Converters

```python
from office2md.converters.pandoc_converter import PANDOC_AVAILABLE
from office2md.converters.mammoth_converter import MAMMOTH_AVAILABLE

print(f"Pandoc: {'✅' if PANDOC_AVAILABLE else '❌'}")
print(f"Mammoth: {'✅' if MAMMOTH_AVAILABLE else '❌'}")
```

## Image Handling

Images are automatically extracted to a folder named `{output_name}_images/`:

```
document.docx
  ↓
document.md
document_images/
  ├── image_1.png
  ├── image_2.jpg
  └── image_3.png
```

Markdown references use relative paths:

```markdown
![](./document_images/image_1.png)
```

### Custom Image Directory

```bash
office2md document.docx --images-dir ./assets/images
```

## Quality Comparison

Use the quality check script to compare converters:

```bash
python scripts/quality_check.py document.docx
```

Output:

```
============================================================
QUALITY REPORT
============================================================
Input: document.docx

====================================================================================================
Converter    Status   Time     Size       Images   Headings   Tables   Issues    
====================================================================================================
default      ✅        4.19s    24,338     30       1          0        2         
pandoc       ✅        3.25s    62,678     30       1          0        0         
====================================================================================================

🏆 BEST CONVERTER: pandoc (score: 100/100)
```

## Supported Formats

| Format     | Extension | Converter                  | Notes                          |
| ---------- | --------- | -------------------------- | ------------------------------ |
| Word       | `.docx`   | Pandoc/Mammoth/python-docx | Auto-fallback                  |
| Excel      | `.xlsx`   | openpyxl                   | All sheets or first only       |
| PowerPoint | `.pptx`   | python-pptx                | With/without notes             |
| PDF        | `.pdf`    | Docling                    | Requires `pip install docling` |

## Troubleshooting

### "Pandoc not available"

Install Pandoc for best DOCX conversion:

```bash
# macOS
brew install pandoc

# Ubuntu
sudo apt-get install pandoc
```

### Tables not rendering correctly

Use Pandoc for complex tables:

```bash
office2md document.docx --use-pandoc
```

### Images have wrong paths

Ensure the images directory is created:

```bash
office2md document.docx -v  # Check verbose output for image paths
```

### PDF conversion fails

Install Docling:

```bash
pip install docling
```

## Development

```bash
# Clone repository
git clone https://github.com/yourusername/office2md.git
cd office2md

# Install in development mode
pip install -e ".[dev]"

# Run tests
pytest

# Run with coverage
pytest --cov=office2md
```

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request
