# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.1] - 2024-12-05

### Fixed
- Correção de caminhos de imagens no Pandoc
- Tabelas HTML convertidas para Markdown puro
- Docling limitado apenas para PDF
- Melhorias na extração de texto do python-docx
  
## [0.1.0] - 2024-12-05

### Added
- Initial release of office2md
- **DOCX Conversion**
  - Pandoc converter (primary, best tables)
  - Mammoth converter (secondary, good formatting)
  - python-docx converter (fallback, basic)
  - Automatic fallback chain: Pandoc → Mammoth → python-docx
  - CLI flags: `--use-pandoc`, `--use-mammoth`, `--use-basic`
- **XLSX Conversion**
  - Full sheet support with openpyxl
  - Option: `--first-sheet-only`
- **PPTX Conversion**
  - Slide content and speaker notes
  - Option: `--no-notes`
- **PDF Conversion** (optional)
  - Docling-based ML extraction
  - Flag: `--use-docling`
- **Image Extraction**
  - Automatic extraction to `{output}_images/` folder
  - Option: `--skip-images`, `--images-dir`
- **Batch Processing**
  - Directory conversion with `--batch`
  - Recursive with `--recursive`
- **CLI**
  - Verbose mode with `-v`
  - Quality check script

### Notes
- Docling is recommended for PDF files only
- Pandoc requires external binary installation
- All other converters are pure Python

[Unreleased]: https://github.com/yourusername/office2md/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/yourusername/office2md/releases/tag/v0.1.0