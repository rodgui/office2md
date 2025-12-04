# Examples

This directory contains example Office files and their converted Markdown outputs.

## Structure

- `input/` - Example Office files (docx, xlsx, pptx)
- `output/` - Converted Markdown files

## Usage

To convert the example files:

```bash
# Convert a single file
office2md examples/input/sample.docx -o examples/output/sample.md

# Batch convert all example files
office2md --batch examples/input -o examples/output
```

## Creating Test Files

Since example Office files are binary, they need to be created manually. You can:

1. Create simple Office files using Microsoft Office, LibreOffice, or Google Docs
2. Use the provided Python scripts to generate test files programmatically
3. Download sample files from the internet

### Minimal Test File Creation

You can use Python to create minimal test files:

```python
# Create a simple DOCX file
from docx import Document
doc = Document()
doc.add_heading('Sample Document', 0)
doc.add_paragraph('This is a sample paragraph.')
doc.save('examples/input/sample.docx')

# Create a simple XLSX file
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws['A1'] = 'Name'
ws['B1'] = 'Value'
ws['A2'] = 'Item 1'
ws['B2'] = 100
wb.save('examples/input/sample.xlsx')

# Create a simple PPTX file
from pptx import Presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
title = slide.shapes.title
title.text = 'Sample Presentation'
prs.save('examples/input/sample.pptx')
```
