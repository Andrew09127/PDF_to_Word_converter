"""Настройка Docling pipeline для отсканированных PDF."""
from __future__ import annotations

from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend
from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import (
    EasyOcrOptions,
    PdfPipelineOptions,
    TableFormerMode,
)
from docling.document_converter import DocumentConverter, PdfFormatOption


def build_converter(langs: list[str] | None = None) -> DocumentConverter:
    """
    Создаёт DocumentConverter для сканов:
      - EasyOCR с принудительным полностраничным OCR
      - TableFormer ACCURATE для таблиц
      - PyPdfium backend
    """
    if langs is None:
        langs = ["ru", "en"]

    # confidence_threshold ниже дефолтного 0.5: при 72 DPI EasyOCR помечал
    # часть реальных слов как «неуверенные» и Docling их ОТБРАСЫВАЛ (терялись
    # «В связи», предлоги). 0.2 сохраняет такие слова.
    ocr_opts = EasyOcrOptions(lang=langs, force_full_page_ocr=True)
    try:
        ocr_opts.confidence_threshold = 0.2
    except Exception:
        pass

    pipeline_opts = PdfPipelineOptions()
    pipeline_opts.do_ocr                                   = True
    pipeline_opts.ocr_options                              = ocr_opts
    pipeline_opts.do_table_structure                       = True
    pipeline_opts.table_structure_options.do_cell_matching = True
    pipeline_opts.table_structure_options.mode             = TableFormerMode.ACCURATE
    pipeline_opts.generate_page_images                     = True

    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_opts,
                backend=PyPdfiumDocumentBackend,
            )
        }
    )
