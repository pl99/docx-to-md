# AGENTS.md — docx-to-md

Single-script Python utility converting `.docx` → `.md` using `python-docx`.
No package structure (no `pyproject.toml` / `setup.py`), raw script only.

## Commands

```bash
# run all tests
pytest -v

# run one test
pytest -v -k "test_heading"

# run conversion
python docx_to_md.py input.docx output.md
```

## Project structure

| Path | Purpose |
|------|---------|
| `docx_to_md.py` | Single entry point with CLI (`main()`) + library API (`convert_docx_to_markdown`) |
| `tests/test_docx_to_md.py` | All tests — uses `sys.path.insert(0, …)` hack to import from parent |
| `docs/superpowers/` | Design spec and implementation plan |
| `requirements.txt` | `python-docx>=1.1.2`, `pytest>=8.0` |

## Quirks to know

- **No fixtures on disk** — every test creates a `.docx` in memory via `Document()`.
- **`tests/test_docx_to_md.py`** uses `sys.path.insert(0, str(Path(__file__).parent.parent))` before `from docx_to_md import …`. Breaking this import path will break all tests.
- **Table dispatch**: simple tables (no colspan/rowspan in any cell) → GFM pipe tables. Complex tables → HTML `<table>` with `colspan`/`rowspan` attributes.
- **First table row always treated as header** in pipe tables; all cells rendered as `<td>` (never `<th>`) in HTML tables.
- **`_cell_text()`** joins cell paragraph texts with spaces: `" ".join(p.text or "" for p in cell.paragraphs).strip()`.
- **Russian CLI messages** — error/success strings are in Russian.
- No linter, formatter, type checker, or CI configured.
