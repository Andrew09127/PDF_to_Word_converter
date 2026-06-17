# Встраивание конвертера в sberAct

Конвертер спроектирован как **переносимый модуль**: HTTP-логика вынесена в
`docling_dev/api.py` (FastAPI `APIRouter`), а `main.py` — лишь локальный dev-запуск.
В sberAct (бэкенд тоже FastAPI) встраивание сводится к подключению того же роутера.

Всё работает **полностью локально** — сервер слушает только `127.0.0.1`,
модели вендорены, сетевых вызовов нет.

## Архитектура

```
React (Electron-фронт sberAct)
  └─ экран ScanConverter ──HTTP──► FastAPI (python-backend/main.py)
                                     └─ app.include_router(converter_router)
                                          └─ docling_dev/api.py
                                               ├─ /convert/scan    → Docling+RapidOCR+LLM
                                               ├─ /convert/native  → pdf2docx
                                               ├─ /convert/status/{id}
                                               └─ /convert/download/{id}
```

## Бэкенд (python-backend)

1. Скопировать в `python-backend/` рядом с их кодом:
   - пакет `docling_dev/`
   - `convert_no_scan.py` (нативный режим)
   - `vendor/` (wheels), `models/` (RapidOCR + LLM .gguf)
2. Влить зависимости в их `requirements.txt` (наш веб-слой у них уже есть —
   FastAPI/uvicorn — поэтому добавить только конвертерные пакеты + docling/rapidocr).
3. В их `python-backend/app/main.py` добавить две строки:
   ```python
   from docling_dev.api import router as converter_router
   app.include_router(converter_router)
   ```
   Их CORS (`allow_origins=["*"]`) и привязка к `127.0.0.1:8000` уже подходят.

## Фронтенд (electron-app)

Навигация в sberAct — переключение через состояние `currentStep` (не react-router).

1. Добавить новый шаг, напр. `'convert'`, в их `App.tsx`.
2. Перенести компонент `frontend/src/ScanConverter.tsx` в их `electron-app/src/components/`
   (он уже на MUI). При переносе:
   - убрать собственную `ThemeProvider` из `main.tsx` — компонент возьмёт их тему;
   - пути API относительные (`/convert/...`) — менять не нужно, бэкенд тот же origin.
3. Добавить пункт входа на экран конвертера в их навигацию.

## Эндпоинты API

| Метод | Путь | Назначение |
|---|---|---|
| POST | `/convert/scan` | PDF + флаги (`no_highlight`,`word_order`,`iim`,`ink_bold`,`ocr_preprocess`) → `{job_id}` |
| POST | `/convert/native` | PDF (без флагов) → `{job_id}` |
| GET | `/convert/status/{job_id}` | `{status, stage, progress, filename, error}` |
| GET | `/convert/download/{job_id}` | готовый DOCX (после `status==done`) |

Задачи выполняются в фоне (один воркер, сериализация), клиент опрашивает статус
раз в секунду. Прогресс — грубые стадии (конвертация идёт одним блокирующим вызовом).
