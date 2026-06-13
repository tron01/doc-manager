import pytest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.docx_tool import _substitute_variables

def test_variable_substitution_simple():
    data = {"text": "Hello {{name}}"}
    vars = {"name": "World"}
    res = _substitute_variables(data, vars)
    assert res["text"] == "Hello World"

def test_variable_substitution_nested():
    data = {"content": [{"text": "Value: {{val}}"}]}
    vars = {"val": "42"}
    res = _substitute_variables(data, vars)
    assert res["content"][0]["text"] == "Value: 42"
