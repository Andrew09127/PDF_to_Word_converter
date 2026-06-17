"""Настройка Docling pipeline для отсканированных PDF."""
from __future__ import annotations

import logging
import os

from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend
from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import (
    EasyOcrOptions,
    PdfPipelineOptions,
    TableFormerMode,
)
from docling.document_converter import DocumentConverter, PdfFormatOption

log = logging.getLogger(__name__)

# Вендоренные модели RapidOCR (.onnx) — кладутся в models/rapidocr/.
# det.onnx (детекция, языконезависима) — ОБЯЗАТЕЛЬНО СЕРВЕРНАЯ ch_PP-OCRv5_det_server
#   (~88МБ): mobile-детектор терял ~половину текста на части документов (низкий
#   recall — не находил текстовые области), серверный находит почти всё. Mobile —
#   дефолт пакета rapidocr (ch_PP-OCRv5_det_mobile), при нужде качается оттуда; в git
#   НЕ держим (теряет текст). Recall дополнительно поднимает rapidocr_merge.py.
# cls.onnx (поворот строки, опц.), rec.onnx (РАСПОЗНАВАНИЕ — eslav PP-OCRv5,
#   КИРИЛЛИЦА), keys.txt (словарь символов; rec и keys — ПАРА).
_RAPIDOCR_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "models", "rapidocr")


def _easyocr_options(langs: list[str]):
    """EasyOCR-опции (движок по умолчанию)."""
    opts = EasyOcrOptions(lang=langs, force_full_page_ocr=True)
    # confidence_threshold ниже дефолтного 0.5: при низком DPI EasyOCR помечает
    # реальные слова как «неуверенные» и Docling их отбрасывает. 0.2 сохраняет.
    try:
        opts.confidence_threshold = 0.2
    except Exception:
        pass
    return opts


def _rapidocr_options(langs: list[str]):
    """RapidOcrOptions с ВЕНДОРЕННЫМИ кириллическими моделями. None если пакет
    или модели недоступны (тогда вызывающий откатывается на EasyOCR)."""
    try:
        from docling.datamodel.pipeline_options import RapidOcrOptions
        import rapidocr  # noqa: F401  — проверка наличия пакета
    except Exception as exc:
        log.warning("RapidOCR: пакет недоступен (%s) — остаюсь на EasyOCR. "
                    "Установите офлайн из vendor/wheels (см. INSTALL_RAPIDOCR.md)", exc)
        return None

    det  = os.path.join(_RAPIDOCR_DIR, "det.onnx")
    cls  = os.path.join(_RAPIDOCR_DIR, "cls.onnx")
    rec  = os.path.join(_RAPIDOCR_DIR, "rec.onnx")
    keys = os.path.join(_RAPIDOCR_DIR, "keys.txt")
    missing = [os.path.basename(p) for p in (det, rec, keys) if not os.path.isfile(p)]
    if missing:
        log.warning("RapidOCR: нет моделей %s в %s — остаюсь на EasyOCR "
                    "(см. INSTALL_RAPIDOCR.md)", missing, _RAPIDOCR_DIR)
        return None

    opts = RapidOcrOptions(
        lang=langs, force_full_page_ocr=True, backend="onnxruntime",
        det_model_path=det, rec_model_path=rec, rec_keys_path=keys,
    )
    if os.path.isfile(cls):
        opts.cls_model_path = cls
    try:
        opts.confidence_threshold = 0.2
    except Exception:
        pass
    log.info("RapidOCR: движок включён (кириллические модели из %s)", _RAPIDOCR_DIR)
    return opts


def build_converter(
    langs: list[str] | None = None,
    images_scale: float = 2.0,
    ocr_preprocess: bool = False,
    ocr_engine: str = "rapidocr",
) -> DocumentConverter:
    """
    Создаёт DocumentConverter для сканов:
      - EasyOCR с принудительным полностраничным OCR
      - TableFormer ACCURATE для таблиц
      - PyPdfium backend
      - images_scale=2.0 → 144 DPI (по умолчанию; 1.0 = 72 DPI)

    images_scale влияет на ВСЁ качество обработки:
      - При 72 DPI (scale=1.0) символы 1-2 пикселя → путаница I/l/1, U/0, Z/2
      - При 144 DPI (scale=2.0) символы 3-4 пикселя → значительно меньше ошибок
      - При 216 DPI (scale=3.0) выше распознавание, НО OCR дробит строки мельче →
        конвертер оценивает более мелкий кегль → вёрстка пакуется плотнее и пагинация
        расходится с оригиналом (Alfa 4→3 страницы). Поэтому по умолчанию 144.
    """
    if langs is None:
        langs = ["ru", "en"]

    # Предобработка изображения перед EasyOCR (deskew + CLAHE-контраст). ОПЦИЯ
    # (по умолчанию ВЫКЛ): помогает плохим сканам (Батлер 102→90, ПСБ 71→66, Doc1
    # 23→14 искажённых), НО может портить мелкий текст хороших сканов — CLAHE даёт
    # артефакты, EasyOCR путает кириллицу с латиницей («регистрации»→«pezucmpauuu»
    # в шапке Alfa). Включать флагом --ocr-preprocess для проблемных сканов.
    if ocr_preprocess:
        try:
            from .ocr_preprocess import install_easyocr_preprocess
            install_easyocr_preprocess()
        except Exception as _exc:
            import logging
            logging.getLogger(__name__).debug("ocr_preprocess недоступен: %s", _exc)

    # Выбор OCR-движка: «rapidocr» (если привезён пакет+модели) или «easyocr»
    # (по умолчанию). Если RapidOCR недоступен — тихий откат на EasyOCR.
    ocr_opts = _rapidocr_options(langs) if ocr_engine == "rapidocr" else None
    if ocr_opts is None:
        ocr_opts = _easyocr_options(langs)
    else:
        # Объединённая детекция на нескольких масштабах — повышает recall RapidOCR
        # (часть строк находится только на одном из масштабов). Чистоту не теряем —
        # распознаёт всё та же eslav rec-модель. Патч идемпотентен.
        try:
            from .rapidocr_merge import install_merged_detection
            install_merged_detection()
        except Exception as _exc:
            import logging
            logging.getLogger(__name__).debug("merged detection недоступен: %s", _exc)

    pipeline_opts = PdfPipelineOptions()
    pipeline_opts.do_ocr                                   = True
    pipeline_opts.ocr_options                              = ocr_opts
    pipeline_opts.do_table_structure                       = True
    pipeline_opts.table_structure_options.do_cell_matching = True
    pipeline_opts.table_structure_options.mode             = TableFormerMode.ACCURATE
    pipeline_opts.generate_page_images                     = True
    pipeline_opts.images_scale                             = images_scale

    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_opts,
                backend=PyPdfiumDocumentBackend,
            )
        }
    )
