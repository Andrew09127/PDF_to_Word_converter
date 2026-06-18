"""
renderer.py
Публичный API рендеринга документа.

Оркестрирует pipeline:
  1. page_analyser  — анализ страниц (тип, колонки, медианы)
  2. doc_type       — определение типа документа
  3. legal_fixes    — document-specific исправления (только для DOC_LEGAL)
  4. build_docx     — рендер в DOCX (существующая логика, не изменена)

build_docx внутри по-прежнему использует свой _fix_reading_order для
обратной совместимости. Этот модуль — точка входа для новой архитектуры.
"""
from __future__ import annotations

import logging

from .converter import build_docx
from .doc_type  import detect_doc_type, DOC_LEGAL
from .page_analyser import analyse_pages as analyse_pages_new
from .block_sorter  import sort_items_by_reading_order
from .geometry      import detect_pdf_native, reading_order_key

log = logging.getLogger(__name__)


def render_document(
    dl_doc,
    page_sizes: dict[int, tuple[float, float]],
    ocr_reader=None,
    use_word_order: bool = True,
) -> object:
    """
    Главная точка входа: DoclingDocument - python-docx Document.

    Параметры:
        dl_doc         — Docling DoclingDocument
        page_sizes     — {page_no: (width_pt, height_pt)}
        ocr_reader     — EasyOCR Reader (None = word_order отключён)
        use_word_order — использовать ли EasyOCR для уточнения порядка

    Возвращает python-docx Document.
    """
    all_items = list(dl_doc.iterate_items())

    # Определяем тип документа по содержимому
    doc_type = detect_doc_type(all_items)
    log.info("renderer: doc_type=%s", doc_type)

    # Анализ страниц — новый модуль (дополняет, не заменяет analyse_pages в build_docx)
    pdf_native = detect_pdf_native(all_items)
    page_infos = analyse_pages_new(all_items)

    for pn, info in sorted(page_infos.items()):
        log.info("  стр.%d: тип=%s колонок=%d median_h=%.1f left_min=%.1f",
                 pn, info.page_type, len(info.columns), info.median_h, info.left_min)

    # Делегируем build_docx — он содержит полную рабочую логику
    # включая letterhead, fix_reading_order (legal), рендер параграфов и т.д.
    doc = build_docx(
        dl_doc,
        page_sizes,
        ocr_reader=ocr_reader,
        use_word_order=use_word_order,
        doc_type=doc_type,
        page_infos=page_infos,
    )

    return doc
