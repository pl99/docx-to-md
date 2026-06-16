import pytest
from docx import Document
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx_to_md import convert_document


def test_plain_paragraph():
    doc = Document()
    doc.add_paragraph("Hello world")
    result = convert_document(doc)
    assert result == "Hello world\n"


def test_heading():
    doc = Document()
    doc.add_heading("Title", level=1)
    doc.add_heading("Subtitle", level=2)
    result = convert_document(doc)
    assert result == "# Title\n\n## Subtitle\n"


def test_unordered_list():
    doc = Document()
    doc.add_paragraph("Item 1", style="List Bullet")
    result = convert_document(doc)
    assert result == "- Item 1\n"


def test_ordered_list():
    doc = Document()
    doc.add_paragraph("First", style="List Number")
    result = convert_document(doc)
    assert result == "1. First\n"


def test_bold_and_italic():
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run("bold text")
    run.bold = True
    run2 = p.add_run(" and ")
    run3 = p.add_run("italic text")
    run3.italic = True
    result = convert_document(doc)
    assert result == "**bold text** and _italic text_\n"
