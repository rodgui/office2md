"""Tests for XlsxConverter."""


import pytest

from office2md.converters.xlsx_converter import XlsxConverter


class TestXlsxConverter:
    """Test XlsxConverter class."""

    def test_file_not_found(self):
        """Test that FileNotFoundError is raised for missing files."""
        with pytest.raises(FileNotFoundError):
            XlsxConverter("nonexistent.xlsx")

    def test_output_path_default(self, tmp_path):
        """Test that default output path is correct."""
        input_file = tmp_path / "test.xlsx"
        input_file.write_text("test")
        converter = XlsxConverter(str(input_file))
        assert converter.output_path == tmp_path / "test.md"

    def test_output_path_custom(self, tmp_path):
        """Test that custom output path is used."""
        input_file = tmp_path / "test.xlsx"
        input_file.write_text("test")
        output_file = tmp_path / "custom.md"
        converter = XlsxConverter(str(input_file), str(output_file))
        assert converter.output_path == output_file
