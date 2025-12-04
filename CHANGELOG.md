# Changelog

Todas as mudanças notáveis neste projeto estão documentadas neste arquivo.

## [0.2.0] - 2025-12-04

### Added
- Versionamento automático com setuptools_scm
- Suporte a extração de imagens em subdiretório (padrão)
- Opções `--embed-images` e `--skip-images` para controlar manipulação de imagens
- Comando `--version` no CLI
- Suporte a imagens base64 em conversão DOCX

### Changed
- Comportamento padrão: imagens extraídas para `{output_stem}_images/`
- Refatoração de BaseConverter para suportar múltiplos modos de imagem

## [0.1.0] - 2025-12-01

### Added
- Suporte inicial para DOCX, XLSX, PPTX
- CLI com modo batch e recursivo
- Factory pattern para roteamento de conversores
- Fallback mammoth → python-docx para DOCX
- Opções por formato: `--no-mammoth`, `--first-sheet-only`, `--no-notes`