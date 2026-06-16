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


def build_converter(
    langs: list[str] | None = None,
    images_scale: float = 2.0,
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
      - При 216 DPI (scale=3.0) ещё лучше, но в 2× медленнее и требует RAM
    """
    if langs is None:
        langs = ["ru", "en"]

    # confidence_threshold ниже дефолтного 0.5: при низком DPI EasyOCR помечает
    # реальные слова как «неуверенные» и Docling их отбрасывает. 0.2 сохраняет.
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
    pipeline_opts.images_scale                             = images_scale

    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_opts,
                backend=PyPdfiumDocumentBackend,
            )
        }
    )
