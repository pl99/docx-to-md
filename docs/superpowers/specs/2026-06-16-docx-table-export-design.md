# docx_to_md: Table Export as Markdown Tables

Date: 2026-06-16

## Overview

Replace `pypandoc` with `python-docx` for converting Word (`.docx`) files to Markdown (`.md`). The primary driver is that pandoc outputs grid tables that are not portable, whereas standard GFM pipe tables or HTML `<table>` elements are needed.

## Architecture

A single script `docx_to_md.py`. The function `convert_docx_to_markdown` iterates over `document.body` elements and dispatches each to one of two converters:

- `convert_paragraph(p)` — paragraphs, headings, lists
- `convert_table(t)` — tables (pipe or HTML)

### Dependency change

| Before | After |
|--------|-------|
| `pypandoc` | `python-docx` |

## Table complexity criterion

| Format | Condition |
|--------|-----------|
| Pipe table | Exactly one header row (first row) AND no `colspan`/`rowspan` in any cell |
| HTML `<table>` | Any `colspan`/`rowspan` present OR no header row |

Multi-paragraph cells, empty cells, and variable column counts always fall back to HTML.

## Pipe table conversion

```
| Header1 | Header2 | Header3 |
| ------- | ------- | ------- |
| Cell1   | Cell2   | Cell3   |
```

- Empty cells → empty string between pipes
- Minimum 3 dashes in separator row
- Columns padded for readable alignment in output

## HTML table conversion

```html
<table>
<tr>
  <td colspan="2">Merged cell</td>
  <td>Normal</td>
</tr>
</table>
```

- All cells rendered as `<td>` (no `<th>`) — header row detection is unreliable with merged cells
- `colspan` and `rowspan` attributes preserved
- Multi-paragraph cell content → `<p>` tags inside `<td>`

## Paragraph, heading, list conversion

| Element | Markdown |
|---------|----------|
| Heading 1–6 | `#` … `######` |
| Unordered list | `- ` prefix |
| Ordered list | `1. ` prefix |
| Bold | `**text**` |
| Italic | `_text_` |
| Monospace | `` `text` `` |
| Empty line separator | added between blocks |

## Testing

- pytest-based tests
- Fixture `.docx` files with simple tables, merged cells, no-header tables
- Assert output matches expected pipe or HTML tables
- Test mixed content (headings + tables + lists)
