"""Unit tests for console encoding configuration (R1)."""

import sys
import unittest
from unittest.mock import MagicMock

from synthesize import configure_console_encoding


class ConsoleEncodingTest(unittest.TestCase):
    def test_reconfigures_non_utf8_stream(self):
        mock_stdout = MagicMock()
        mock_stdout.encoding = "cp936"
        mock_stderr = MagicMock()
        mock_stderr.encoding = "gbk"

        original_stdout, original_stderr = sys.stdout, sys.stderr
        try:
            sys.stdout, sys.stderr = mock_stdout, mock_stderr
            configure_console_encoding()
            mock_stdout.reconfigure.assert_called_once_with(encoding="utf-8", errors="replace")
            mock_stderr.reconfigure.assert_called_once_with(encoding="utf-8", errors="replace")
        finally:
            sys.stdout, sys.stderr = original_stdout, original_stderr

    def test_noop_when_already_utf8(self):
        mock_stdout = MagicMock()
        mock_stdout.encoding = "UTF-8"
        mock_stderr = MagicMock()
        mock_stderr.encoding = "utf-8"

        original_stdout, original_stderr = sys.stdout, sys.stderr
        try:
            sys.stdout, sys.stderr = mock_stdout, mock_stderr
            configure_console_encoding()
            mock_stdout.reconfigure.assert_not_called()
            mock_stderr.reconfigure.assert_not_called()
        finally:
            sys.stdout, sys.stderr = original_stdout, original_stderr

    def test_tolerates_missing_reconfigure_method(self):
        # E.g. io.StringIO or custom objects that don't implement reconfigure
        mock_stdout = object()
        original_stdout = sys.stdout
        try:
            sys.stdout = mock_stdout
            # Must not raise
            configure_console_encoding()
        finally:
            sys.stdout = original_stdout
