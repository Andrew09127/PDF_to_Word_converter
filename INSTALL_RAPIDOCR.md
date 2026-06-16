# Подключение RapidOCR (второй OCR-движок) — офлайн

RapidOCR — альтернативный OCR-движок (как EasyOCR), python-библиотека, работает в
процессе Python **без exe/bat**. Под капотом — ONNX-модели + onnxruntime (C++ внутри
wheel). Подключается полностью офлайн, как LLM: wheel в git + `.onnx` модели в git.

Docling **нативно** поддерживает RapidOCR — интеграция уже написана
(`docling_dev/pipeline.py`, флаг `--ocr-engine rapidocr`). Нужно только привезти
пакеты и кириллические модели.

---

## Шаг 1. Скачать на машине С ИНТЕРНЕТОМ

### 1а. Wheel-пакеты (движок)
```bash
pip download rapidocr onnxruntime -d vendor/wheels --only-binary :all:
```
Это скачает `rapidocr`, `onnxruntime` и их зависимости (`opencv-python`,
`pyclipper`, `shapely`, `pyyaml`, `tqdm`, …) — все `.whl` лягут в `vendor/wheels/`.

> Важно: качать на ТОЙ ЖЕ ОС/версии Python (Windows, Python 3.12), что и целевой ПК —
> wheel бинарные. Проверить версию: `python --version` на целевом ПК.

### 1б. Кириллические модели (.onnx)
Стандартные модели RapidOCR — китайские/английские. Для русского нужна
**кириллическая rec-модель**. Источник — RapidOCR model zoo / PaddleOCR (PP-OCRv4/v5):

| Файл (положить как) | Что это | Где взять |
|---|---|---|
| `models/rapidocr/det.onnx` | детекция текста (языконезависима) | RapidOCR: `ch_PP-OCRv4_det` (onnx) |
| `models/rapidocr/cls.onnx` | классиф. поворота (опц., языконезависима) | RapidOCR: `ch_ppocr_mobile_v2.0_cls` (onnx) |
| `models/rapidocr/rec.onnx` | **РАСПОЗНАВАНИЕ — кириллица** | PaddleOCR `cyrillic_PP-OCRv3/v4_rec` → onnx |
| `models/rapidocr/keys.txt` | словарь символов кириллицы | идёт с rec-моделью (`cyrillic_dict.txt`) |

Источники моделей:
- RapidOCR: https://github.com/RapidAI/RapidOCR (раздел Model / ModelScope);
- PaddleOCR multilingual: https://github.com/PaddlePaddle/PaddleOCR (cyrillic). Если
  модель в формате PaddlePaddle (`.pdmodel`), конвертировать в onnx через `paddle2onnx`.

> rec-модель и keys.txt должны быть ПАРОЙ (словарь соответствует модели).

---

## Шаг 2. Перенести на офлайн-ПК и установить
```bash
# wheel-пакеты (в .venv проекта)
.\.venv\Scripts\pip install --no-index --find-links vendor/wheels rapidocr onnxruntime
```
Модели уже лежат в `models/rapidocr/` (перенесены вместе с репозиторием/архивом).

Проверка установки:
```bash
.\.venv\Scripts\python -c "import rapidocr; print('rapidocr OK')"
```

---

## Шаг 3. Запуск
```bash
# один файл
.\.venv\Scripts\python convert_docling_dev.py "pdf_backup\file.pdf" -o out --ocr-engine rapidocr

# батч
.\.venv\Scripts\python convert_docling_dev.py --ocr-engine rapidocr
```
Если пакет/модели не найдены — конвертер автоматически откатится на EasyOCR и
напишет предупреждение (документ всё равно сконвертируется).

---

## Шаг 4. Сравнить с EasyOCR (нужно ОБЯЗАТЕЛЬНО)
RapidOCR не гарантированно точнее EasyOCR на ваших сканах — это надо измерить.
Скажите, когда привезёте wheel+модели, — я прогоню A/B на Батлере/Alfa/ПСБ
(число искажённых слов до/после) и сравню, как делал для предобработки.

---

## Что уже готово в коде
- `docling_dev/pipeline.py`: `build_converter(ocr_engine="rapidocr")` →
  `RapidOcrOptions` с путями к `models/rapidocr/*.onnx`; при отсутствии — откат на EasyOCR;
- CLI: `--ocr-engine {easyocr,rapidocr}` (по умолчанию `easyocr`);
- папки `vendor/wheels/` и `models/rapidocr/` созданы (с `.gitkeep`).
