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


def _is_simple_table(table) -> bool:
    """Table is simple if it has no colspan/rowspan in any cell."""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tc_pr = tc.find(qn('w:tcPr'))
            if tc_pr is not None:
                # colspan
                if tc_pr.find(qn('w:gridSpan')) is not None:
                    return False
                # rowspan (vMerge)
                if tc_pr.find(qn('w:vMerge')) is not None:
                    return False
    return True


def _convert_table_pipe(table) -> str:
    rows_data = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            text = _cell_text(cell)
            cells.append(text)
        rows_data.append(cells)

    if not rows_data:
        return ""

    # Column widths for alignment padding
    col_widths = []
    for col_idx in range(len(rows_data[0])):
        widths = []
        for row in rows_data:
            if col_idx < len(row):
                widths.append(len(row[col_idx]))
        col_widths.append(max(widths) if widths else 3)

    lines = []

    # Header row
    header = rows_data[0]
    line = "| " + " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(header)) + " |"
    lines.append(line)

    # Separator row
    sep = "| " + " | ".join("-" * max(w, 3) for w in col_widths) + " |"
    lines.append(sep)

    # Data rows
    for row in rows_data[1:]:
        line = "| " + " | ".join(
            (row[i] if i < len(row) else "").ljust(col_widths[i])
            for i in range(len(col_widths))
        ) + " |"
        lines.append(line)

    return "\n".join(lines) + "\n"


def _cell_text(cell) -> str:
    return " ".join(p.text or "" for p in cell.paragraphs).strip()
