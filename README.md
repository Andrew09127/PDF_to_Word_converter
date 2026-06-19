# PDF → Word конвертер

Превращает PDF (отсканированные и нативные) в редактируемые DOCX, максимально
близкие по вёрстке к оригиналу. Заточен под юридические документы (заявления о
включении в РТК и т.п.). **Работает полностью локально** — никакие данные не
уходят в интернет.

Есть два способа использования:
- **Веб-интерфейс** (браузер, локальный сервер) — основной, удобный.
- **Командная строка** (батч/один файл) — для автоматизации.

---

## Что делает приложение

Принимает PDF и выдаёт DOCX. Два режима конвертации под разные типы документов:

| Режим | Для чего | Движок | Скорость |
|---|---|---|---|
| **Сканы (LLM)** | страницы-изображения, текст не выделяется | Docling + RapidOCR + LLM | медленнее |
| **Нативный PDF** | есть текстовый слой (текст копируется) | pdf2docx | быстро |

В скан-режиме распознаётся текст (RapidOCR, серверный детектор — высокий recall,
не теряет строки), восстанавливается структура (заголовки, абзацы, таблицы
метка|значение, шапка сторон, штамп ЭП), чистятся типичные OCR-ошибки. Нативный
режим переносит текст/таблицы/картинки напрямую; если pdf2docx теряет часть строк,
они дописываются в конец документа — **полнота текста гарантируется**.

---

## Как это устроено

```
Браузер (React + MUI, frontend/)
   │  HTTP (только 127.0.0.1)
   ▼
FastAPI (main.py → docling_dev/api.py)
   ├─ /convert/scan    → docling_dev (Docling + RapidOCR + LLM)
   ├─ /convert/native  → convert_no_scan.py (pdf2docx)
   ├─ /convert/status/{id}   (опрос прогресса)
   └─ /convert/download/{id} (готовый DOCX)
```

- **Бэкенд** — Python/FastAPI. HTTP-слой вынесен в `docling_dev/api.py`
  (`APIRouter`), `main.py` — локальный сервер, который ещё и раздаёт собранный
  фронт статикой. Конвертация идёт в фоне, клиент опрашивает статус.
- **Фронт** — Vite + React + TypeScript + MUI (`frontend/`). Собранная версия
  лежит в `frontend/dist/` (коммитится) — поэтому **на конечной машине Node не
  нужен**.
- **Карта модулей** бэкенда — см. `CLAUDE.md` и комментарии в `docling_dev/`.

Полностью офлайн: сервер слушает только `127.0.0.1`, модели и пакеты — локально.

---

## Установка

Нужен **Python 3.12, только 64-битный**. Дальше — выберите сценарий.

> **32-бит не поддерживается.** Это ограничение ML-стека, а не нашей упаковки:
> torch (нужен Docling), onnxruntime (нужен RapidOCR), paddlepaddle и
> llama-cpp-python **не выпускают 32-битных wheel'ов под Windows** в принципе,
> плюс 32-битный процесс упёрся бы в лимит ~2–4 ГБ памяти. Требуется 64-битная
> ОС и 64-битный Python. Если на машине стоит 32-битный Python — поставьте рядом
> 64-битный Python 3.12 (отдельная ОС не нужна).

### A. Онлайн (любая ОС: Windows / Linux / macOS)

```bash
git clone <репозиторий> convert
cd convert
python -m venv .venv
# Windows:        .venv\Scripts\activate
# Linux / macOS:  source .venv/bin/activate

# Полный точный набор (рекомендуется, Windows x64):
pip install -r requirements-lock.txt ^
  --extra-index-url https://download.pytorch.org/whl/cpu ^
  --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu

# Либо читаемый список прямых зависимостей (любая ОС):
pip install -r requirements.txt ^
  --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu
```
Доп. индексы нужны: `download.pytorch.org/whl/cpu` — CPU-сборка torch (без
гигантского CUDA), `abetlen.github.io/...` — готовый бинарный `llama-cpp-python`
(иначе pip полезет собирать из исходника).

> `requirements-lock.txt` пинит `torch==2.12.0+cpu` и Windows-wheel'ы — он точен
> для **Windows x64**. На Linux/macOS ставьте `requirements.txt` (pip подберёт
> wheel'ы под вашу платформу).

### B. Офлайн через `git clone` (Windows x64 — готовый комплект)

Все пакеты уже лежат в репозитории (`offline_kit/wheels/`), поэтому после
клонирования интернет не нужен:

```bat
git clone <репозиторий> convert
cd convert
tools\install_offline.bat
```
Скрипт: склеит крупные wheel из кусков → создаст `.venv` → поставит все 152
пакета из `offline_kit\wheels` (`pip install --no-index`). Ни интернета, ни
Node, ни компилятора.

Крупные wheel (torch ~117 МБ) превышают лимит GitHub 100 МБ, поэтому лежат
кусками `*.whl.partNNN` и склеиваются при установке (`tools\reassemble_wheels.py`)
— тот же приём, что и для LLM-модели.

### C. Офлайн на другой ОС (Linux / macOS)

Готовый комплект собран под Windows x64. Под другую ОС его надо **пересобрать**
на машине с интернетом этой же ОС:

```bash
python -m pip wheel -r requirements-lock.txt -w offline_kit/wheels \
  --only-binary=llama-cpp-python,torch,torchvision \
  --extra-index-url https://download.pytorch.org/whl/cpu \
  --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu
python tools/split_wheels.py            # разбить крупные на куски
```
На офлайн-машине той же ОС:
```bash
python tools/reassemble_wheels.py
python -m venv .venv && source .venv/bin/activate
pip install --no-index --find-links offline_kit/wheels -r requirements-lock.txt
```
(`.bat`-скрипты — для Windows; на Linux/macOS выполняйте эти команды напрямую.)

---

## Запуск

### Веб-интерфейс
```bash
.venv\Scripts\python.exe main.py      # Windows  (Linux/Mac: .venv/bin/python main.py)
```
Откройте **http://127.0.0.1:8000**. Загрузите PDF, выберите режим
(«Отсканированный PDF» или «Неотсканированный PDF»), скачайте DOCX.

Разработка фронта с hot-reload (нужен Node, только для правок интерфейса):
```bash
cd frontend && npm install && npm run dev   # http://127.0.0.1:5173 (прокси на :8000)
npm run build                               # пересобрать dist/ перед коммитом
```

### Командная строка
```bash
# батч из pdf_to_convert/ → converted_word_docling/
python convert_docling_dev.py

# один файл
python convert_docling_dev.py "путь/к/файлу.pdf" --output out_dir
```
Полезные флаги:
- `--word-order` — EasyOCR-переупорядочивание блоков (ВЫКЛ по умолчанию: нативный
  порядок Docling точнее).
- `--no-highlight` — не подсвечивать жёлтым подозрительные OCR-слова.
- `--no-llm` — отключить LLM-доочистку остатка (по умолчанию **ВКЛ**; см. ниже).
- `--ocr-preprocess` — предобработка плохих сканов (deskew + контраст).

---

## Очистка OCR-ошибок (скан-режим)

1. **Детерминированная** (всегда вкл): нормализация помеченных слов — гомоглифы
   латиница→кириллица (`co→со`), случайные заглавные (`должНЫ→должны`).
2. **Подсветка**: что не починилось автоматически — выделяется жёлтым для
   быстрой ручной вычитки.
3. **LLM-доочистка (опц.)**: добивает оставшиеся подсвеченные слова — только по
   ним, не по всему документу → быстро даже на CPU.

### Локальная LLM (llama-cpp-python)
Движок — pip-пакет `llama-cpp-python`, модель — `.gguf` в `models/`. Инференс
идёт в самом процессе Python, офлайн, на CPU, без прав администратора. И
библиотека, и модель **вендорятся в git** — работает в закрытой корпсети.

Правки LLM проходят жёсткие предохранители: не меняются цифры/номера/ФИО/римские
цифры/чисто-латинские токены; правка должна быть мелкой (иначе отклоняется как
галлюцинация). Используется 1.5B (`qwen2.5-1.5b-instruct-q4_k_m`) — безопасно
добивает остаток (модель 3B протаскивает неверные правки вроде `VII→VIII`).

#### Модель в git без Git LFS (куски + авто-склейка)
GitHub не принимает файлы >100 МБ. Поэтому `.gguf` (~1 ГБ) не коммитится, а в git
лежат куски `models/*.gguf.partNNN` (~95 МБ, без LFS). При первом запуске с
`--llm` модуль сам склеивает их в полный `.gguf` (`llm_correct._resolve_model`).

Перерезать модель на куски (если заменили `.gguf`):
```bash
python - <<'PY'
src="models/qwen2.5-1.5b-instruct-q4_k_m.gguf"; C=95*1024*1024; i=0
with open(src,"rb") as f:
    while (b:=f.read(C)):
        open(f"{src}.part{i:03d}","wb").write(b); i+=1
print("кусков:", i)
PY
```

---

## Встраивание в приложение sberAct

Конвертер спроектирован переносимым: HTTP-логика в `docling_dev/api.py`
(`APIRouter`), `main.py` — лишь локальный dev-запуск. В sberAct (бэкенд тоже
FastAPI) встраивание сводится к подключению того же роутера.

**Бэкенд** (`python-backend/`):
1. Скопировать рядом: `docling_dev/`, `convert_no_scan.py`, `vendor/`, `models/`.
2. В их `app/main.py` добавить:
   ```python
   from docling_dev.api import router as converter_router
   app.include_router(converter_router)
   ```
   Их CORS (`allow_origins=["*"]`) и привязка к `127.0.0.1:8000` уже подходят.

**Фронтенд** (`electron-app/`, навигация через состояние `currentStep`):
1. Добавить шаг, напр. `'convert'`, в `App.tsx`.
2. Перенести `frontend/src/ScanConverter.tsx` в их `components/` (он уже на MUI;
   убрать собственную `ThemeProvider` — возьмёт их тему). Пути API относительные
   (`/convert/...`), менять не нужно — тот же origin.

Эндпоинты API:

| Метод | Путь | Назначение |
|---|---|---|
| POST | `/convert/scan` | PDF + флаги → `{job_id}` |
| POST | `/convert/native` | PDF (без флагов) → `{job_id}` |
| GET | `/convert/status/{job_id}` | `{status, stage, progress, filename, error}` |
| GET | `/convert/download/{job_id}` | готовый DOCX (после `status==done`) |

---

## Тесты вёрстки (golden)
```bash
python tests/golden.py save     # снять эталоны из converted_word_docling/
python tests/golden.py check    # сверить текущий вывод с эталонами (diff)
```
Ловит регрессии состава/порядка текста при изменениях кода.

---

## Структура файлов зависимостей и документации

- `requirements.txt` — читаемый список прямых зависимостей (онлайн-установка).
- `requirements-lock.txt` — полный точный слепок (152 пакета) для офлайн/воспроизводимой установки.
- `offline_kit/wheels/` — сами wheel'ы для офлайн-установки (крупные — кусками).
- `tools/` — скрипты сборки/установки офлайн-комплекта и разбиения/склейки wheel.
- `CLAUDE.md` — рабочий конспект проекта (карта модулей, канон проверки).
