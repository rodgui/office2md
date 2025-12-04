# office2md Copilot Instructions

## Project Overview

**office2md** is a Python converter toolkit transforming Microsoft Office files (DOCX, XLSX, PPTX) into Markdown format. The architecture uses a **Factory + Strategy pattern** with format-specific converters inheriting from `BaseConverter`.

## Architecture & Design Patterns

### Core Components

1. **ConverterFactory** (`office2md/converter_factory.py`): Routes file types to appropriate converters
   - Maps file extensions to converter classes
   - Validates file support via `SUPPORTED_EXTENSIONS` dict
   - Creates converter instances with optional kwargs

2. **BaseConverter** (abstract base): Defines the conversion contract
   - `convert()`: Transforms file to Markdown string
   - `save()`: Persists Markdown to disk
   - Handles path resolution (input → output.md)
   - All converters inherit this class

3. **Format-Specific Converters**: Four formats supported
   - **DocxConverter**: Primary=mammoth (best fidelity), fallback=python-docx
   - **XlsxConverter**: openpyxl; supports `include_all_sheets` parameter
   - **PptxConverter**: python-pptx; supports `include_notes` parameter

### Adding a New Converter

1. Create `office2md/converters/new_converter.py` inheriting `BaseConverter`
2. Implement `convert()` returning markdown string
3. Add extension → class mapping to `SUPPORTED_EXTENSIONS` in factory
4. Add test file `tests/converters/test_new_converter.py` 
5. Export from `office2md/__init__.py`

### Key Design Decisions

- **Mammoth → python-docx fallback**: DOCX uses mammoth for superior HTML-to-MD conversion, but gracefully falls back if unavailable
- **Per-format options as kwargs**: `use_mammoth`, `include_all_sheets`, `include_notes` passed through factory to allow fine-grained control
- **Batch mode preserves directory structure**: When using `--batch --recursive`, relative paths are maintained in output directory

## Testing Strategy

Run all tests: `pytest`
Run specific module: `pytest tests/test_cli.py`
Run with coverage: `pytest --cov=office2md`

Tests validate:
- Factory routes to correct converter classes
- Each converter handles its format correctly
- CLI argument parsing and batch processing logic
- Error handling for missing/unsupported files

## CLI Workflows

```bash
# Single file: auto-generates output.md in same directory
office2md input.docx

# Custom output: explicit -o path
office2md input.docx -o output/result.md

# Batch mode: flatten or preserve structure
office2md --batch ./input -o ./output                    # flat
office2md --batch ./input -o ./output --recursive        # structure preserved

# Format-specific options
office2md --no-mammoth doc.docx                          # use python-docx
office2md --first-sheet-only data.xlsx                   # XLSX only
office2md --no-notes presentation.pptx                   # PPTX only
```

## Dependencies & Fallbacks

- **mammoth** (≥1.6.0): DOCX conversion; if missing, python-docx used
- **python-docx** (≥1.1.0): DOCX fallback + sample creation
- **openpyxl** (≥3.1.0): XLSX/XLS reading
- **python-pptx** (≥0.6.23): PPTX reading

All deps imported with try/except to log warnings but not crash. Check `AVAILABLE` flags in converters.

## Common Patterns

### Markdown Generation

Converters use lists and string concatenation:
```python
markdown_lines = ["# Title"]
for item in content:
    markdown_lines.append(f"- {item}")
return "\n".join(markdown_lines)
```

### Error Handling in CLI

- `convert_file()` returns bool (success/failure), logs exceptions
- `batch_convert()` returns tuple (success_count, failure_count), continues on errors
- Main CLI returns exit code 0 (success) or 1 (failure)

### Batch Processing

CLI `batch_convert()` uses `Path.glob()` with `**/*` (recursive) or `*` (flat):
- Filters results via `ConverterFactory.is_supported()`
- Preserves directory structure when `--recursive` is used
- Creates output directories with `mkdir(parents=True)`

## Debugging

Enable verbose logging: `office2md -v input.docx`
Logging format: `TIMESTAMP - MODULE_NAME - LEVEL - MESSAGE`

Check fallback usage: "mammoth not available" warning indicates python-docx in use.
