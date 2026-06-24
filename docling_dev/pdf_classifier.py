"""
pdf_classifier.py
─────────────────
Быстрое определение типа PDF: отсканированный (нужен OCR) или нативный
(есть текстовый слой). Используется веб-интерфейсом, чтобы подсветить
рекомендованный режим конвертации при загрузке файла.

Анализ синхронный и лёгкий — только чтение текстового слоя через PyMuPDF
(fitz), без OCR. Логика порогов портирована из проекта analizator_to_PDF
(функция is_scanned_pdf), но на PyMuPDF, который уже есть в зависимостях
(pdf2docx), чтобы не тащить новый пакет в офлайн-комплект.

Полностью локально — ничего в сеть не уходит.
"""
from __future__ import annotations

import logging
from pathlib import Path

import fitz  # PyMuPDF

log = logging.getLogger(__name__)


def _is_scanned(pdf_path: Path) -> bool:
    """
    Эвристика «скан или текст» по текстовому слою и размеру файла.
    True — отсканированный (нужен OCR), False — нативный (есть текст).
    Пороги перенесены 1-в-1 из analizator_to_PDF.is_scanned_pdf.
    """
    with fitz.open(pdf_path) as doc:
        total_pages = doc.page_count
        if total_pages == 0:
            return True

        pages_with_real_text = 0.0
        total_text_length = 0
        pages_with_only_digits = 0

        for page in doc:
            text = page.get_text()
            if text:
                clean_text = ''.join(c for c in text if c.isalnum() or c in '.,!?-;:')
                clean_text = clean_text.strip()

                if len(clean_text) > 30:
                    pages_with_real_text += 1
                    total_text_length += len(clean_text)
                elif len(clean_text) > 0:
                    # Не только цифры (номера страниц) — считаем за половину «текстовой».
                    if not clean_text.replace('.', '').replace('-', '').isdigit():
                        pages_with_real_text += 0.5
                        total_text_length += len(clean_text)
                    else:
                        pages_with_only_digits += 1

        text_page_ratio = pages_with_real_text / total_pages
        has_substantial_text = total_text_length > 150
        file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
        avg_size_per_page = file_size_mb / total_pages
        # Страницы только с цифрами (номера страниц в скан-копиях).
        has_many_page_numbers = pages_with_only_digits > total_pages * 0.3

        # Явно текстовый: большинство страниц содержат реальный текст.
        if text_page_ratio >= 0.6 and has_substantial_text:
            return False

        # Явно отсканированный: мало текста + большой размер (данные изображений).
        if text_page_ratio < 0.2 and avg_size_per_page > 0.3:
            return True

        # Скан с OCR: только номера страниц, почти нет текста.
        if has_many_page_numbers and total_text_length < 100:
            return True

        # Большой объём изображений на страницу, почти нет текста.
        if avg_size_per_page > 0.8 and total_text_length < 200:
            return True

        # Очень мало текста в ненулевом файле.
        if total_text_length < 50 and file_size_mb > 0.5:
            return True

        # Есть значимый текст — скорее текстовый.
        if text_page_ratio > 0.3 and total_text_length > 80:
            return False

        return True


def classify_pdf(pdf_path: Path) -> dict:
    """
    Определяет рекомендуемый режим конвертации для PDF.

    Возвращает словарь:
      suggested  — "scan" (отсканированный) | "native" (нативный);
      confidence — "high" | "low" (low — если файл не прочитался, дефолт «скан»);
      reason     — короткое пояснение для пользователя.
    """
    try:
        scanned = _is_scanned(pdf_path)
    except Exception as exc:  # noqa: BLE001 — при сбое чтения безопасный дефолт «скан»
        log.warning("classify_pdf: не удалось прочитать %s: %s", pdf_path.name, exc)
        return {
            "suggested": "scan",
            "confidence": "low",
            "reason": "Не удалось проанализировать файл — на всякий случай рекомендуем "
                      "режим распознавания скана.",
        }

    if scanned:
        return {
            "suggested": "scan",
            "confidence": "high",
            "reason": "Похоже на отсканированный документ (текст не выделяется) — "
                      "рекомендуем режим распознавания скана.",
        }
    return {
        "suggested": "native",
        "confidence": "high",
        "reason": "В файле есть текстовый слой — рекомендуем быстрый режим для "
                  "неотсканированного PDF.",
    }
