# PDF to Word Converter

Конвертер PDF в редактируемые DOCX-файлы.

## Как установить

Требуется Python 3.10+.

```powershell
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## Как пользоваться

1. Положите PDF-файлы в папку `pdf_to_convert`.
2. Запустите:

```powershell
python main.py
```

3. Готовые Word-файлы появятся в папке `converted_word`.

После успешной конвертации исходные PDF удаляются из `pdf_to_convert`.

## Важно

Качество конвертации зависит от PDF. Если PDF состоит из сканов, обычного редактируемого текста внутри него нет. Для таких файлов нужен OCR.
