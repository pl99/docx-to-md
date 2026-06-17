# docx-to-md

Конвертер документов `.docx` в формат Markdown.

## Описание

Утилита для преобразования файлов Microsoft Word (`.docx`) в Markdown с сохранением структуры:

- Заголовки (уровни 1-6)
- Абзацы и форматирование текста (**жирный**, *курсив*)
- Маркированные и нумерованные списки
- Таблицы: простые (pipe-формат) и сложные с объединением ячеек (HTML `<table>`)

## Установка

```bash
pip install -r requirements.txt
```

## Использование

### Командная строка

```bash
python docx_to_md.py input.docx output.md
```

### Как библиотека

```python
from pathlib import Path
from docx_to_md import convert_docx_to_markdown

input_path = Path("document.docx")
output_path = Path("document.md")
convert_docx_to_markdown(input_path, output_path)
```

### Конвертация документа в строку

```python
from docx import Document
from docx_to_md import convert_document

doc = Document("document.docx")
markdown = convert_document(doc)
print(markdown)
```

## Тесты

```bash
pytest tests/
```

## Лицензия

MIT
