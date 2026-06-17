"""
api.py
──────
HTTP-слой конвертера: FastAPI APIRouter.

Не запускает сервер сам по себе — только описывает эндпоинты. Подключается
в любой FastAPI-приложение строкой `app.include_router(router)`:
  • локально для разработки — из main.py (см. корень проекта);
  • в бою — из python-backend/main.py приложения sberAct.

Полностью локально, без сетевых вызовов. Логику конвертации не дублирует —
оборачивает существующие точки входа:
  • сканы          → docling_dev.convert_pdf (Docling + RapidOCR + LLM)
  • нативные PDF    → convert_no_scan.PDFPipelineConverter.convert_single_pdf (pdf2docx)

Конвертация тяжёлая и блокирующая, поэтому выполняется в фоновом потоке, а
клиент опрашивает статус. Задачи сериализованы (один воркер) — приложение
однопользовательское локальное, параллельный прогон сканов только съел бы память.
"""
from __future__ import annotations

import logging
import shutil
import tempfile
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from fastapi import APIRouter, BackgroundTasks, Form, HTTPException, UploadFile, File
from fastapi.responses import FileResponse

log = logging.getLogger(__name__)

router = APIRouter(prefix="/convert", tags=["convert"])

# ── Хранилище задач (в памяти) ───────────────────────────────────────────────
# job_id → {status, stage, progress, mode, filename, docx_path, error}
#   status:   queued | running | done | error
#   progress: 0.0 … 1.0 (грубые отметки — convert_pdf одним блокирующим вызовом)
_JOBS: dict[str, dict] = {}
_JOBS_LOCK = threading.Lock()

# Один воркер — задачи идут по очереди (сериализация прогонов).
_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="convert")

# Рабочая папка задач (вход PDF + выход DOCX). В системном temp, чистится при старте.
_WORK_DIR = Path(tempfile.gettempdir()) / "convert_api_jobs"
_WORK_DIR.mkdir(parents=True, exist_ok=True)

# Кэш Docling-конвертера: построение тяжёлое (грузит модели), поэтому держим
# по одному экземпляру на значение ocr_preprocess (оно меняет pipeline).
_SCAN_CONVERTERS: dict[bool, object] = {}
_SCAN_LOCK = threading.Lock()

# Кэш EasyOCR-reader: нужен только для флага word_order (переупорядочивание
# блоков). Инициализация тяжёлая (грузит модели), поэтому строим лениво один раз.
_OCR_READER: list = []          # [reader] или [] ; None-маркер недоступности — [None]
_OCR_READER_LOCK = threading.Lock()


def _set(job_id: str, **changes) -> None:
    """Атомарно обновить запись задачи."""
    with _JOBS_LOCK:
        if job_id in _JOBS:
            _JOBS[job_id].update(changes)


def _get_scan_converter(ocr_preprocess: bool):
    """Лениво строит и кэширует Docling-конвертер под флаг предобработки."""
    with _SCAN_LOCK:
        conv = _SCAN_CONVERTERS.get(ocr_preprocess)
        if conv is None:
            from . import build_converter
            log.info("api: строю Docling-конвертер (ocr_preprocess=%s)", ocr_preprocess)
            conv = build_converter(ocr_preprocess=ocr_preprocess, ocr_engine="rapidocr")
            _SCAN_CONVERTERS[ocr_preprocess] = conv
    return conv


def _get_ocr_reader():
    """Лениво строит EasyOCR-reader для word_order. None — если недоступен."""
    with _OCR_READER_LOCK:
        if not _OCR_READER:
            try:
                import easyocr
                log.info("api: инициализация EasyOCR для word_order…")
                _OCR_READER.append(easyocr.Reader(["ru", "en"], gpu=False, verbose=False))
                log.info("api: EasyOCR готов.")
            except Exception as exc:                 # noqa: BLE001
                log.warning("api: EasyOCR недоступен, word_order отключён: %s", exc)
                _OCR_READER.append(None)
        return _OCR_READER[0]


# ── Фоновые воркеры ──────────────────────────────────────────────────────────

def _run_scan(job_id: str, pdf_path: Path, docx_path: Path, flags: dict) -> None:
    """Прогон скана через Docling-пайплайн (медленный путь)."""
    try:
        _set(job_id, status="running", stage="Анализ скана…", progress=0.1)
        converter = _get_scan_converter(flags["ocr_preprocess"])

        # EasyOCR-reader нужен только для word_order — иначе не грузим (он тяжёлый).
        ocr_reader = _get_ocr_reader() if flags["word_order"] else None

        _set(job_id, stage="Распознавание текста и сборка DOCX…", progress=0.35)
        from . import convert_pdf
        ok = convert_pdf(
            pdf_path, docx_path, converter,
            ocr_reader=ocr_reader,
            use_word_order=flags["word_order"],
            highlight=flags["highlight"],
            llm=flags["llm"],
            ink_bold=flags["ink_bold"],
        )
        if ok and docx_path.exists():
            _set(job_id, status="done", stage="Готово", progress=1.0)
        else:
            _set(job_id, status="error", stage="Ошибка", error="Конвертация не удалась")
    except Exception as e:                       # noqa: BLE001 — отдаём текст клиенту
        log.exception("api: ошибка прогона скана %s", job_id)
        _set(job_id, status="error", stage="Ошибка", error=str(e))


def _run_native(job_id: str, pdf_path: Path, docx_path: Path) -> None:
    """Прогон нативного PDF через pdf2docx (быстрый путь, без флагов)."""
    try:
        _set(job_id, status="running", stage="Конвертация PDF…", progress=0.2)
        # Берём только одиночную конвертацию — без батч-пайплайна, который удаляет исходники.
        from convert_no_scan import PDFPipelineConverter
        conv = PDFPipelineConverter(
            input_folder=str(pdf_path.parent),
            output_folder=str(docx_path.parent),
        )
        ok, _ = conv.convert_single_pdf(pdf_path, docx_path)
        if ok and docx_path.exists():
            _set(job_id, status="done", stage="Готово", progress=1.0)
        else:
            _set(job_id, status="error", stage="Ошибка", error="Конвертация не удалась")
    except Exception as e:                       # noqa: BLE001
        log.exception("api: ошибка прогона нативного PDF %s", job_id)
        _set(job_id, status="error", stage="Ошибка", error=str(e))


# ── Приём загруженного файла ─────────────────────────────────────────────────

def _accept_upload(file: UploadFile) -> tuple[str, Path, Path]:
    """Сохраняет загруженный PDF в папку задачи. Возвращает (job_id, pdf, docx)."""
    name = Path(file.filename or "document.pdf").name
    if not name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Принимаются только PDF-файлы.")

    job_id = uuid.uuid4().hex
    job_dir = _WORK_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = job_dir / name
    with pdf_path.open("wb") as out:
        shutil.copyfileobj(file.file, out)

    docx_path = job_dir / f"{Path(name).stem}.docx"
    return job_id, pdf_path, docx_path


# ── Эндпоинты ────────────────────────────────────────────────────────────────

@router.post("/scan")
async def convert_scan(
    background: BackgroundTasks,
    file: UploadFile = File(...),
    no_highlight: bool = Form(True),
    word_order: bool = Form(False),
    iim: bool = Form(True),
    ink_bold: bool = Form(False),
    ocr_preprocess: bool = Form(False),
):
    """Скан-режим (Docling + RapidOCR + LLM). Возвращает job_id для опроса статуса."""
    job_id, pdf_path, docx_path = _accept_upload(file)
    flags = {
        "highlight": not no_highlight,   # фронтовый --no-highlight инвертирует подсветку
        "word_order": word_order,
        "llm": iim,                      # «iim» в интерфейсе = LLM-доочистка
        "ink_bold": ink_bold,
        "ocr_preprocess": ocr_preprocess,
    }
    with _JOBS_LOCK:
        _JOBS[job_id] = {"status": "queued", "stage": "В очереди…", "progress": 0.0,
                         "mode": "scan", "filename": docx_path.name,
                         "docx_path": str(docx_path), "error": None}
    _EXECUTOR.submit(_run_scan, job_id, pdf_path, docx_path, flags)
    return {"job_id": job_id}


@router.post("/native")
async def convert_native(file: UploadFile = File(...)):
    """Нативный PDF (pdf2docx, без флагов). Возвращает job_id для опроса статуса."""
    job_id, pdf_path, docx_path = _accept_upload(file)
    with _JOBS_LOCK:
        _JOBS[job_id] = {"status": "queued", "stage": "В очереди…", "progress": 0.0,
                         "mode": "native", "filename": docx_path.name,
                         "docx_path": str(docx_path), "error": None}
    _EXECUTOR.submit(_run_native, job_id, pdf_path, docx_path)
    return {"job_id": job_id}


@router.get("/status/{job_id}")
async def convert_status(job_id: str):
    """Текущий статус задачи: queued | running | done | error (+ стадия и прогресс)."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
        if job is None:
            raise HTTPException(status_code=404, detail="Задача не найдена.")
        return {"job_id": job_id, "status": job["status"], "stage": job["stage"],
                "progress": job["progress"], "filename": job["filename"],
                "error": job["error"]}


@router.get("/download/{job_id}")
async def convert_download(job_id: str):
    """Отдаёт готовый DOCX. Доступно только после status == done."""
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Задача не найдена.")
    if job["status"] != "done":
        raise HTTPException(status_code=409, detail="Конвертация ещё не завершена.")
    docx_path = Path(job["docx_path"])
    if not docx_path.exists():
        raise HTTPException(status_code=410, detail="Результат недоступен.")
    return FileResponse(
        docx_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=job["filename"],
    )
