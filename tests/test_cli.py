"""Tests for CLI module."""

from unittest.mock import patch

import pytest

from office2md.cli import main, setup_logging


class TestCLI:
    """Test CLI functionality."""

    def test_setup_logging_default(self):
        """Test that logging is setup with INFO level by default."""
        with patch("logging.basicConfig") as mock_basic_config:
            setup_logging(verbose=False)
            mock_basic_config.assert_called_once()
            call_kwargs = mock_basic_config.call_args[1]
            assert call_kwargs["level"] == 20  # logging.INFO

    def test_setup_logging_verbose(self):
        """Test that logging is setup with DEBUG level when verbose."""
        with patch("logging.basicConfig") as mock_basic_config:
            setup_logging(verbose=True)
            mock_basic_config.assert_called_once()
            call_kwargs = mock_basic_config.call_args[1]
            assert call_kwargs["level"] == 10  # logging.DEBUG

    def test_main_no_args(self):
        """Test that main exits with error when no arguments provided."""
        with pytest.raises(SystemExit) as exc_info:
            main([])
        assert exc_info.value.code == 2

    def test_main_help(self):
        """Test that help text is displayed."""
        with pytest.raises(SystemExit) as exc_info:
            main(["--help"])
        assert exc_info.value.code == 0

    @patch("office2md.cli.convert_file")
    def test_main_single_file(self, mock_convert):
        """Test single file conversion through CLI."""
        mock_convert.return_value = True
        exit_code = main(["test.docx"])
        assert exit_code == 0
        mock_convert.assert_called_once()

    @patch("office2md.cli.batch_convert")
    def test_main_batch_mode(self, mock_batch):
        """Test batch mode through CLI."""
        mock_batch.return_value = (5, 0)
        exit_code = main(["--batch", "input_dir"])
        assert exit_code == 0
        mock_batch.assert_called_once()
