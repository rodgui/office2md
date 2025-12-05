"""Version information for office2md."""

__version__ = "0.1.2"
__version_info__ = tuple(int(x) for x in __version__.split("."))

# Version history:
# 0.1.0 - Initial release
#   - DOCX conversion with Pandoc/Mammoth/python-docx fallback chain
#   - XLSX conversion with openpyxl
#   - PPTX conversion with python-pptx
#   - PDF conversion with Docling (optional)
#   - CLI with batch processing
#   - Image extraction