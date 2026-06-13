import pytest
import os
import json

@pytest.fixture
def sample_content():
    return {
        "metadata": {"title": "Test Doc"},
        "content": [
            {"type": "heading", "level": 1, "text": "Heading 1"},
            {"type": "paragraph", "text": "A paragraph"}
        ]
    }

@pytest.fixture
def tmp_output_dir(tmp_path):
    d = tmp_path / "output"
    d.mkdir()
    return str(d)
