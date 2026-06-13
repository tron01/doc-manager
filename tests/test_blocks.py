import pytest
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from scripts.docx_tool import _build_document

def test_build_document_basic(sample_content):
    doc = _build_document(sample_content)
    paragraphs = [p.text for p in doc.paragraphs if p.text]
    assert doc.core_properties.title == "Test Doc"
    assert "Heading 1" in paragraphs
    assert "A paragraph" in paragraphs
