import pytest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.docx_tool import _validate_content

def test_validate_valid_content(sample_content):
    errors, warnings = _validate_content(sample_content)
    assert not errors
    assert not warnings

def test_validate_missing_type():
    content = {"content": [{"text": "no type"}]}
    errors, warnings = _validate_content(content)
    assert len(errors) == 1
    assert "missing 'type'" in errors[0]

def test_validate_unknown_type():
    content = {"content": [{"type": "magic_block"}]}
    errors, warnings = _validate_content(content)
    assert len(errors) == 1
    assert "unknown type" in errors[0]

def test_validate_table_no_headers():
    content = {"content": [{"type": "table", "rows": [["1"]]}]}
    errors, warnings = _validate_content(content)
    assert len(errors) == 1
    assert "no headers" in errors[0]
