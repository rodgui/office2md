"""Tests for DocxConverter."""


import pytest

from office2md.converters.docx_converter import DocxConverter


class TestDocxConverter:
    """Test DocxConverter class."""

    def test_file_not_found(self):
        """Test that FileNotFoundError is raised for missing files."""
        with pytest.raises(FileNotFoundError):
            DocxConverter("nonexistent.docx")

    def test_output_path_default(self, tmp_path):
        """Test that default output path is correct."""
        input_file = tmp_path / "test.docx"
        input_file.write_text("test")
        converter = DocxConverter(str(input_file))
        assert converter.output_path == tmp_path / "test.md"

    def test_output_path_custom(self, tmp_path):
        """Test that custom output path is used."""
        input_file = tmp_path / "test.docx"
        input_file.write_text("test")
        output_file = tmp_path / "custom.md"
        converter = DocxConverter(str(input_file), str(output_file))
        assert converter.output_path == output_file
