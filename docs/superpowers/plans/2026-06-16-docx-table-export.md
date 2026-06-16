# docx_to_md: Table Export as Markdown Tables — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace `pypandoc` with `python-docx` so that Word tables are exported as pipe tables (simple) or HTML `<table>` (complex/merged cells).

**Architecture:** Single script `docx_to_md.py` with separate functions for paragraph and table conversion. Table complexity is detected by checking for `gridSpan`/`vMerge` in XML. Output is assembled by iterating `document.element.body` in document order.

**Tech Stack:** `python-docx` 1.1.2, standard library only beyond that, `pytest` for testing.

---

## File Structure

| File | Status | Purpose |
|------|--------|---------|
| `docx_to_md.py` | Modify | Rewrite to use python-docx directly |
| `requirements.txt` | Create | List python-docx as dependency |
| `tests/test_docx_to_md.py` | Create | All unit/integration tests |

No test fixtures on disk — each test creates its own `.docx` in memory via `Document()`.

---

### Task 1: Update dependencies and create test harness

**Files:**
- Create: `requirements.txt`
- Create: `tests/test_docx_to_md.py`
- Modify: `docx_to_md.py` (imports only)

- [ ] **Step 1: Create requirements.txt**

```txt
python-docx>=1.1.2
pytest>=8.0
```

- [ ] **Step 2: Create `tests/test_docx_to_md.py` with imports and an empty placeholder test**

```python
import pytest
from docx import Document
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx_to_md import convert_document


def test_placeholder():
    assert True
```

- [ ] **Step 3: Add python-docx import to `docx_to_md.py` and a stub `convert_document`**

```python
from docx import Document
from docx.oxml.ns import qn


def convert_document(doc: Document) -> str:
    return ""
```

- [ ] **Step 4: Verify test runs**

Run:
```bash
cd C:\wrk\python\utils\docx
python -m pytest tests\test_docx_to_md.py -v
```
Expected: 1 passed

- [ ] **Step 5: Install python-docx and pytest if needed**

Run:
```bash
python -m pip install python-docx pytest
```

- [ ] **Step 6: Commit**

```bash
git init
git add .
git commit -m "chore: set up project structure and test harness"
```

---

### Task 2: Implement paragraph conversion

**Files:**
- Modify: `docx_to_md.py` (add `_convert_paragraph`)
- Modify: `tests/test_docx_to_md.py`

- [ ] **Step 1: Write tests for paragraph conversion**

Add to `test_docx_to_md.py`:

```python
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
    p = doc.add_paragraph("Item 1", style="List Bullet")
    result = convert_document(doc)
    assert result == "- Item 1\n"


def test_ordered_list():
    doc = Document()
    p = doc.add_paragraph("First", style="List Number")
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v
```
Expected: 5 failures (functions not implemented)

- [ ] **Step 3: Implement `_convert_paragraph` in `docx_to_md.py`**

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v
```
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: add paragraph, heading, and list conversion"
```

---

### Task 3: Implement table complexity detection

**Files:**
- Modify: `docx_to_md.py`
- Modify: `tests/test_docx_to_md.py`

- [ ] **Step 1: Write tests for table complexity detection**

```python
def _make_simple_table(doc):
    """Helper to create a 3x2 table with a header row, no merged cells."""
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    table.cell(2, 0).text = "3"
    table.cell(2, 1).text = "4"
    return table


def _make_complex_table_with_colspan(doc):
    """Helper to create a table with colspan."""
    from docx.oxml.ns import qn
    table = doc.add_table(rows=2, cols=3)
    # merge first two cells in row 0
    cell = table.cell(0, 0)
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = tc.makeelement(qn('w:tcPr'), {})
        tc.insert(0, tcPr)
    gridSpan = tcPr.makeelement(qn('w:gridSpan'), {qn('w:val'): '2'})
    tcPr.append(gridSpan)
    # remove the merged-away cell
    table.cell(0, 1)._tc.getparent().remove(table.cell(0, 1)._tc)
    table.cell(0, 2).text = "C"
    table.cell(1, 0).text = "D"
    table.cell(1, 1).text = "E"
    table.cell(1, 2).text = "F"
    return table


def test_table_is_simple():
    doc = Document()
    _make_simple_table(doc)
    table = doc.tables[0]
    from docx_to_md import _is_simple_table
    assert _is_simple_table(table) is True


def test_table_with_colspan_is_complex():
    doc = Document()
    _make_complex_table_with_colspan(doc)
    table = doc.tables[0]
    from docx_to_md import _is_simple_table
    assert _is_simple_table(table) is False
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "table"
```
Expected: 2 failures

- [ ] **Step 3: Implement `_is_simple_table`**

```python
from docx.oxml.ns import qn


def _is_simple_table(table) -> bool:
    """Table is simple if it has a header row and no colspan/rowspan."""
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "table"
```
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: add table complexity detection"
```

---

### Task 4: Implement pipe-table conversion

**Files:**
- Modify: `docx_to_md.py` (add `_convert_table_pipe`)
- Modify: `tests/test_docx_to_md.py`

- [ ] **Step 1: Write tests for pipe table conversion**

```python
def test_pipe_table():
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Age"
    table.cell(1, 0).text = "Alice"
    table.cell(1, 1).text = "30"
    table.cell(2, 0).text = "Bob"
    table.cell(2, 1).text = "25"
    from docx_to_md import _convert_table_pipe
    result = _convert_table_pipe(table)
    expected = (
        "| Name  | Age |\n"
        "| ----- | --- |\n"
        "| Alice | 30  |\n"
        "| Bob   | 25  |\n"
    )
    assert result == expected


def test_pipe_table_empty_cell():
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = ""
    table.cell(0, 2).text = "C"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    table.cell(1, 2).text = "3"
    from docx_to_md import _convert_table_pipe
    result = _convert_table_pipe(table)
    assert "| A |  | C |" in result
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "pipe_table"
```
Expected: 2 failures

- [ ] **Step 3: Implement `_convert_table_pipe`**

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "pipe_table"
```
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add .
git commit -m "feat: add pipe table conversion for simple tables"
```

---

### Task 5: Implement HTML-table conversion (complex tables)

**Files:**
- Modify: `docx_to_md.py` (add `_convert_table_html`, `_get_colspan`, `_get_rowspan`)
- Modify: `tests/test_docx_to_md.py`

- [ ] **Step 1: Write tests for HTML table conversion**

```python
def test_html_table_with_colspan():
    from docx.oxml.ns import qn
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    # Merge first two cells in row 0
    cell = table.cell(0, 0)
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = tc.makeelement(qn('w:tcPr'), {})
        tc.insert(0, tcPr)
    gridSpan = tcPr.makeelement(qn('w:gridSpan'), {qn('w:val'): '2'})
    tcPr.append(gridSpan)
    # remove merged cell
    table.cell(0, 1)._tc.getparent().remove(table.cell(0, 1)._tc)
    table.cell(0, 2).text = "C"
    table.cell(1, 0).text = "D"
    table.cell(1, 1).text = "E"
    table.cell(1, 2).text = "F"

    from docx_to_md import _convert_table_html
    result = _convert_table_html(table)

    assert "<table>" in result
    assert "colspan=\"2\"" in result
    assert result.strip().endswith("</table>")


def test_html_table_no_header():
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    table.cell(1, 0).text = "C"
    table.cell(1, 1).text = "D"
    from docx_to_md import _convert_table_html
    result = _convert_table_html(table)
    assert "<table>" in result
    assert result.count("<tr>") == 2
    assert result.count("<td>") == 4
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "html_table"
```
Expected: 2 failures

- [ ] **Step 3: Implement `_convert_table_html` and helpers**

```python
def _get_colspan(cell) -> int:
    tc = cell._tc
    tc_pr = tc.find(qn('w:tcPr'))
    if tc_pr is None:
        return 1
    gs = tc_pr.find(qn('w:gridSpan'))
    if gs is None:
        return 1
    return int(gs.get(qn('w:val'), 1))


def _has_rowspan_restart(cell) -> bool:
    tc = cell._tc
    tc_pr = tc.find(qn('w:tcPr'))
    if tc_pr is None:
        return False
    vm = tc_pr.find(qn('w:vMerge'))
    if vm is None:
        return False
    val = vm.get(qn('w:val'))
    return val == "restart"


def _is_rowspan_continue(cell) -> bool:
    tc = cell._tc
    tc_pr = tc.find(qn('w:tcPr'))
    if tc_pr is None:
        return False
    vm = tc_pr.find(qn('w:vMerge'))
    if vm is None:
        return False
    val = vm.get(qn('w:val'))
    return val is None  # no val attribute = continue


def _count_rowspan(table, start_row_idx, col_idx) -> int:
    """Count how many rows a vMerge-restart cell spans."""
    count = 1
    for row_idx in range(start_row_idx + 1, len(table.rows)):
        row = table.rows[row_idx]
        cells = row.cells
        if col_idx < len(cells) and _is_rowspan_continue(cells[col_idx]):
            count += 1
        else:
            break
    return count


def _convert_table_html(table) -> str:
    lines = ["<table>"]
    # Track which columns in the current row are consumed by rowspan
    rowspan_skip = {}  # col_idx -> rows_remaining

    for row_idx, row in enumerate(table.rows):
        lines.append("<tr>")
        col_idx = 0
        # Decrement rowspan counters
        skip_cols = set()
        for c in list(rowspan_skip.keys()):
            rowspan_skip[c] -= 1
            if rowspan_skip[c] <= 0:
                del rowspan_skip[c]
            else:
                skip_cols.add(c)
        col_idx = 0
        cell_idx = 0
        while cell_idx < len(row.cells):
            if col_idx in skip_cols:
                col_idx += 1
                continue
            cell = row.cells[cell_idx]
            colspan = _get_colspan(cell)
            text = _cell_text(cell)
            attrs = ""
            if colspan > 1:
                attrs += f' colspan="{colspan}"'
            if _has_rowspan_restart(cell):
                rs = _count_rowspan(table, row_idx, cell_idx)
                if rs > 1:
                    attrs += f' rowspan="{rs}"'
                    rowspan_skip[col_idx] = rs
            lines.append(f"  <td{attrs}>{_escape_html(text)}</td>")
            col_idx += colspan
            cell_idx += 1
        lines.append("</tr>")
    lines.append("</table>")
    return "\n".join(lines) + "\n"


def _escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "html_table"
```
Expected: 2 passed

- [ ] **Step 5: Run all tests**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v
```
Expected: all previous tests still pass

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: add HTML table conversion for complex tables"
```

---

### Task 6: Wire up the main `convert_document` iterator

**Files:**
- Modify: `docx_to_md.py` (rewrite `convert_document`)
- Modify: `tests/test_docx_to_md.py`

- [ ] **Step 1: Write tests for document iteration order**

```python
def test_mixed_content():
    doc = Document()
    doc.add_heading("Report", level=1)
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "X"
    table.cell(0, 1).text = "Y"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    doc.add_paragraph("Footer text")
    result = convert_document(doc)
    assert result.startswith("# Report\n")
    assert "| X | Y |" in result
    assert result.strip().endswith("Footer text")


def test_complex_table_uses_html():
    from docx.oxml.ns import qn
    doc = Document()
    doc.add_paragraph("Before")
    table = doc.add_table(rows=2, cols=2)
    cell = table.cell(0, 0)
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = tc.makeelement(qn('w:tcPr'), {})
        tc.insert(0, tcPr)
    gs = tcPr.makeelement(qn('w:gridSpan'), {qn('w:val'): '2'})
    tcPr.append(gs)
    table.cell(0, 1)._tc.getparent().remove(table.cell(0, 1)._tc)
    # cell(0,1) XML was removed; the merged content is in cell(0,0)
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "B"
    doc.add_paragraph("After")
    result = convert_document(doc)
    assert "Before" in result
    assert "<table>" in result
    assert "After" in result
```

- [ ] **Step 2: Run tests to verify they fail**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "mixed_content or complex_table_uses"
```
Expected: 2 failures

- [ ] **Step 3: Implement `convert_document` with body element iteration**

```python
def convert_document(doc: Document) -> str:
    lines = []
    body = doc.element.body

    # Map XML elements to python-docx objects
    para_by_elem = {p._element: p for p in doc.paragraphs}
    table_by_elem = {t._tbl: t for t in doc.tables}

    for child in body:
        if child.tag == qn('w:p'):
            p = para_by_elem.get(child)
            if p is not None:
                lines.append(_convert_paragraph(p))
        elif child.tag == qn('w:tbl'):
            table = table_by_elem.get(child)
            if table is not None:
                if _is_simple_table(table):
                    lines.append(_convert_table_pipe(table))
                else:
                    lines.append(_convert_table_html(table))

    return "\n".join(lines)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v -k "mixed_content or complex_table_uses"
```
Expected: 2 passed

- [ ] **Step 5: Run all tests**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v
```
Expected: all ~12 tests pass

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: wire up document body iterator with table dispatch"
```

---

### Task 7: Update CLI entry point

**Files:**
- Modify: `docx_to_md.py` (rewrite `convert_docx_to_markdown` and `main`)

- [ ] **Step 1: Implement the updated CLI**

Replace the old `convert_docx_to_markdown` and `main`:

```python
import sys
from pathlib import Path
from docx import Document


def convert_docx_to_markdown(input_path: Path, output_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError("Поддерживается только формат .docx")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(input_path))
    md_content = convert_document(doc)

    output_path.write_text(md_content, encoding="utf-8")


def main():
    if len(sys.argv) != 3:
        print("Usage: python docx_to_md.py input.docx output.md")
        sys.exit(1)

    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2])

    try:
        convert_docx_to_markdown(input_file, output_file)
        print(f"Успешно сконвертировано: {output_file}")
    except Exception as e:
        print(f"Ошибка: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Write integration test**

```python
def test_cli_conversion(tmp_path):
    from docx_to_md import convert_docx_to_markdown

    doc = Document()
    doc.add_heading("Test", level=1)
    doc.add_paragraph("Hello world")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Key"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "1"

    input_path = tmp_path / "test.docx"
    doc.save(str(input_path))
    output_path = tmp_path / "test.md"

    convert_docx_to_markdown(input_path, output_path)

    content = output_path.read_text(encoding="utf-8")
    assert "# Test" in content
    assert "Hello world" in content
    assert "| Key   | Value |" in content
```

- [ ] **Step 3: Run all tests including the new integration test**

Run:
```bash
python -m pytest tests\test_docx_to_md.py -v
```
Expected: all ~13 tests pass

- [ ] **Step 4: Commit**

```bash
git add .
git commit -m "feat: update CLI entry point to use python-docx"
```

---

### Task 8: Final integration smoke test with real file

**Files:** (no file changes, manual verification)

- [ ] **Step 1: Create a test .docx with a real Word document containing mixed content and tables**

Run:
```python
python -c "
from docx import Document
doc = Document()
doc.add_heading('Документация', level=1)
doc.add_paragraph('Ниже представлена таблица функций:')
table = doc.add_table(rows=3, cols=3)
table.cell(0,0).text = 'Код'
table.cell(0,1).text = 'Функция'
table.cell(0,2).text = 'Описание'
table.cell(1,0).text = 'Ф-01.1'
table.cell(1,1).text = 'Создание заявки'
table.cell(1,2).text = 'Создание заявки типа онлайн просмотр'
table.cell(2,0).text = 'Ф-01.2'
table.cell(2,1).text = 'Приём заявки'
table.cell(2,2).text = 'Приём заявки от внешней системы'
doc.save('test_input.docx')
print('Created test_input.docx')
"
```

- [ ] **Step 2: Run conversion**

```bash
python docx_to_md.py test_input.docx test_output.md
```
Expected: `Успешно сконвертировано: test_output.md`

- [ ] **Step 3: Verify output**

```bash
type test_output.md
```

Expected:
```markdown
# Документация

Ниже представлена таблица функций:

| Код    | Функция           | Описание                                |
| ------ | ----------------- | --------------------------------------- |
| Ф-01.1 | Создание заявки   | Создание заявки типа онлайн просмотр    |
| Ф-01.2 | Приём заявки      | Приём заявки от внешней системы         |
```

- [ ] **Step 4: Commit**

```bash
git add .
git commit -m "feat: final working version with table export"
```

---

### Task 9: Clean up temporary files

- [ ] **Step 1: Remove test artifacts**

```bash
Remove-Item -LiteralPath test_input.docx, test_output.md -Force
```

- [ ] **Step 2: Final commit**

```bash
git add -A
git commit -m "chore: clean up test artifacts"
```
