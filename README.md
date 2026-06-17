# docx → md

Конвертация Word (.docx) в Markdown (.md). Замена `pypandoc` на `python-docx` для полного контроля над таблицами.

## Установка

```bash
pip install python-docx
```

## Использование

```bash
python docx_to_md.py input.docx output.md
```

## Таблицы

- **Простые** (без объединённых ячеек) → pipe-таблицы GFM
- **Сложные** (colspan/rowspan) → HTML `<table>`

## Тесты

```bash
pip install pytest
pytest tests/
```
