import pytest
from docx import Document
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx_to_md import convert_document


def test_placeholder():
    assert True
