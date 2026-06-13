import pytest
import sys
import os
import subprocess
from unittest.mock import patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.docx_tool import _convert_to_pdf

@patch('subprocess.run')
@patch('shutil.which')
@patch('os.rename')
def test_convert_to_pdf_fallback(mock_rename, mock_which, mock_run):
    mock_which.return_value = '/usr/bin/soffice'
    mock_run.return_value.returncode = 0
    
    with patch.dict('sys.modules', {'docx2pdf': None}):
        res = _convert_to_pdf("test.docx", "test.pdf")
        assert res == "test.pdf"
        mock_run.assert_called_once()
