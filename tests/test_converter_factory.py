"""Tests for ConverterFactory."""


import pytest

from office2md.converter_factory import ConverterFactory


class TestConverterFactory:
    """Test ConverterFactory class."""

    def test_is_supported_docx(self):
        """Test that DOCX files are supported."""
        assert ConverterFactory.is_supported("test.docx")

    def test_is_supported_xlsx(self):
        """Test that XLSX files are supported."""
        assert ConverterFactory.is_supported("test.xlsx")

    def test_is_supported_xls(self):
        """Test that XLS files are supported."""
        assert ConverterFactory.is_supported("test.xls")

    def test_is_supported_pptx(self):
        """Test that PPTX files are supported."""
        assert ConverterFactory.is_supported("test.pptx")

    def test_is_supported_ppt(self):
        """Test that PPT files are supported."""
        assert ConverterFactory.is_supported("test.ppt")

    def test_is_not_supported(self):
        """Test that unsupported files are correctly identified."""
        assert not ConverterFactory.is_supported("test.pdf")
        assert not ConverterFactory.is_supported("test.txt")
        assert not ConverterFactory.is_supported("test.doc")

    def test_create_converter_unsupported_type(self, tmp_path):
        """Test that ValueError is raised for unsupported file types."""
        test_file = tmp_path / "test.pdf"
        test_file.write_text("test")
        with pytest.raises(ValueError, match="Unsupported file type"):
            ConverterFactory.create_converter(str(test_file))
