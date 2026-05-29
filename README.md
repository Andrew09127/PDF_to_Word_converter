# PDF to Word Converter

Конвертер PDF в редактируемые DOCX-файлы.

Поддерживает два режима:

- обычные PDF с текстовым слоем конвертируются через `pdf2docx`;
- сканированные PDF обрабатываются через `PaddleOCR`.

## Установка На Linux

Рекомендуется Python 3.10 или 3.11.

```bash
cd /path/to/Convert
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install -r requirements-ocr.txt
python -m pip install paddlepaddle -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
```

Проверка:

```bash
python -c "import paddle; print(paddle.__version__); from paddleocr import PaddleOCR, PPStructure; print('PaddleOCR OK')"
```

## Установка На Windows

Обычная PDF-конвертация работает через `requirements.txt`.

PaddleOCR на Windows CPU может падать внутри PaddlePaddle (`OneDnnContext... fused_conv2d`). Для OCR-режима лучше использовать Linux, WSL или Docker.

## Использование

1. Положите PDF-файлы в папку `pdf_to_convert`.
2. Запустите конвертацию:

```bash
python main.py
```

3. Готовые Word-файлы появятся в папке `converted_word`.

После успешной конвертации исходные PDF удаляются из `pdf_to_convert`.

## Режимы

Обычный автоматический режим:

```bash
export CONVERT_MODE=auto
python main.py
```

Принудительный OCR через PaddleOCR:

```bash
export CONVERT_MODE=paddleocr
export PADDLE_LANG=en
export PADDLE_OCR_LANG=ru
export OCR_GPU=0
export PADDLE_CPU_THREADS=4
export PADDLE_FALLBACK_TO_OCR=1
python main.py
```

Только обычные PDF без OCR:

```bash
export CONVERT_MODE=pdf2docx
python main.py
```

## Настройки OCR

`PADDLE_LANG` используется для PP-Structure layout recovery. В PaddleOCR layout-модели доступны только `en` и `ch`.

`PADDLE_OCR_LANG` используется fallback-режимом обычного PaddleOCR. Для русского текста:

```bash
export PADDLE_OCR_LANG=ru
```

Разрешение рендера страниц:

```bash
export OCR_DPI=300
```

Если PP-Structure падает, код автоматически откатывается на обычный PaddleOCR с координатной сборкой DOCX. Отключить fallback:

```bash
export PADDLE_FALLBACK_TO_OCR=0
```

## Ограничения

Сканированный PDF является изображением. OCR не гарантирует Word один-в-один с оригиналом. PaddleOCR может восстановить часть структуры лучше, чем простой OCR, но для максимально точного результата обычно нужен ABBYY, Adobe Acrobat OCR или корпоративный OCR-конвертер.

## Git

Не коммитьте входные PDF и результаты:

- `pdf_to_convert/*.pdf`
- `converted_word/`
- `pdf_conversion.log`
- `.venv/`
