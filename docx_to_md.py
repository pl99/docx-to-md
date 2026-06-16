from docx import Document
from docx.oxml.ns import qn


def convert_document(doc: Document) -> str:
    paragraphs = [_convert_paragraph(p) for p in doc.paragraphs]
    return "\n".join(paragraphs)


def _convert_paragraph(p) -> str:
    style = p.style.name if p.style else "Normal"

    # Heading
    if style.startswith("Heading"):
        level = style.split()[-1]
        if level.isdigit():
            prefix = "#" * int(level)
        else:
            prefix = ""
        return f"{prefix} {_extract_inline_text(p)}\n"

    # Unordered list
    if style.startswith("List Bullet"):
        return f"- {_extract_inline_text(p)}\n"

    # Ordered list
    if style.startswith("List Number"):
        return f"1. {_extract_inline_text(p)}\n"

    # Normal paragraph
    return f"{_extract_inline_text(p)}\n"


def _extract_inline_text(p) -> str:
    parts = []
    for run in p.runs:
        text = run.text
        if run.bold:
            text = f"**{text}**"
        if run.italic:
            text = f"_{text}_"
        parts.append(text)
    return "".join(parts)
