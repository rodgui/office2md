# office2md Copilot Instructions

## Project Overview

**office2md** is a Python converter toolkit transforming Microsoft Office files (DOCX, XLSX, PPTX) and PDFs into Markdown format. The architecture uses a **Factory + Strategy pattern** with format-specific converters inheriting from `BaseConverter`.

## Architecture & Design Patterns

### Core Components

1. **ConverterFactory** (`office2md/converter_factory.py`): Routes file types to appropriate converters
   - Maps file extensions to converter classes
   - Validates file support via `SUPPORTED_EXTENSIONS` dict
   - Creates converter instances with optional kwargs

2. **BaseConverter** (abstract base): Defines the conversion contract
   - `convert()`: Transforms file to Markdown string
   - `save()`: Persists Markdown to disk
   - `_process_image()`: Handles image extraction and reference generation
   - All converters inherit this class

3. **Format-Specific Converters**:
   - **DocxConverter**: Auto-selects Pandoc → Mammoth → python-docx
   - **PandocConverter**: Best DOCX tables, requires external binary
   - **MammothConverter**: Good DOCX formatting, pure Python
   - **BasicDocxConverter**: Fallback using python-docx
   - **XlsxConverter**: openpyxl; supports `include_all_sheets`
   - **PptxConverter**: python-pptx; supports `include_notes`
   - **DoclingConverter**: PDF only, ML-based extraction

### DOCX Converter Hierarchy

```
DocxConverter (orchestrator)
  ├── PandocConverter    (priority 1 - best tables)
  ├── MammothConverter   (priority 2 - good formatting)
  └── BasicDocxConverter (priority 3 - fallback)
```

Flags: `use_pandoc`, `use_mammoth`, `use_basic`

### Adding a New Converter

1. Create `office2md/converters/new_converter.py` inheriting `BaseConverter`
2. Implement `convert()` returning markdown string
3. Add extension → class mapping to `SUPPORTED_EXTENSIONS` in factory
4. Add test file `tests/converters/test_new_converter.py` 
5. Export from `office2md/converters/__init__.py`

### Key Design Decisions

- **Pandoc as primary**: Best table support for DOCX
- **Graceful fallback**: Pandoc → Mammoth → python-docx
- **Docling for PDF only**: DOCX support is limited in Docling
- **Per-format options as kwargs**: Passed through factory
- **Batch mode preserves directory structure**: With `--recursive`

## Testing Strategy

Run all tests: `pytest`
Run specific module: `pytest tests/test_cli.py`
Run with coverage: `pytest --cov=office2md`

Quality check script: `python scripts/quality_check.py document.docx`

## CLI Workflows

```bash
# Single file: auto-selects best converter
office2md input.docx

# Force specific DOCX converter
office2md input.docx --use-pandoc     # Best tables
office2md input.docx --use-mammoth    # Good formatting  
office2md input.docx --use-basic      # python-docx only

# PDF with Docling
office2md input.pdf --use-docling

# Batch mode
office2md --batch ./input -o ./output --recursive

# Format-specific options
office2md --first-sheet-only data.xlsx
office2md --no-notes presentation.pptx
```

## Dependencies & Fallbacks

- **pandoc** (external): DOCX primary converter
- **mammoth** (≥1.6.0): DOCX secondary
- **python-docx** (≥1.1.0): DOCX fallback
- **openpyxl** (≥3.1.0): XLSX
- **python-pptx** (≥0.6.23): PPTX
- **docling**: PDF (optional)
- **markdownify**: HTML→MD conversion

Check availability via `PANDOC_AVAILABLE`, `MAMMOTH_AVAILABLE` flags.

## Common Patterns

### Image Processing

```python
# In converter
ref = self._process_image(image_data, extension)
# Returns: ![](./output_images/image_1.png)
```

### Table Generation

```python
lines = []
lines.append('| ' + ' | '.join(headers) + ' |')
lines.append('| ' + ' | '.join(['---'] * len(headers)) + ' |')
for row in data:
    lines.append('| ' + ' | '.join(row) + ' |')
return '\n'.join(lines)
```

### Error Handling

- `convert_file()` returns bool, logs exceptions
- `batch_convert()` returns (success, failure) tuple
- Main CLI returns exit code 0/1

## Debugging

Enable verbose logging: `office2md -v input.docx`

Check converter used in logs:
```
Converted with pandoc: input.docx
```
