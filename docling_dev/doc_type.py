"""
doc_type.py
Детектор типа документа по содержимому.

Тип документа определяет какие дополнительные fix-модули применяются
поверх базового рендера. Базовый рендер одинаков для всех типов.
"""
from __future__ import annotations

import re

# Типы документов

DOC_LEGAL   = "legal"    # арбитражные заявления, исковые требования
DOC_GENERIC = "generic"  # всё остальное — только базовый рендер

# Паттерны для детекции судебных/арбитражных документов
_LEGAL_PATTERNS = [
    re.compile(r'\bарбитражный\s+суд\b',        re.IGNORECASE),
    re.compile(r'\bпо\s+делу\s+№',              re.IGNORECASE),
    re.compile(r'\bпросит\s+суд\b',             re.IGNORECASE),
    re.compile(r'\bкредитор\b.*\bдолжник\b',    re.IGNORECASE | re.DOTALL),
    re.compile(r'\bисковое\s+заявление\b',       re.IGNORECASE),
    re.compile(r'\bзаявление\b.*\bреестр\b',    re.IGNORECASE | re.DOTALL),
    re.compile(r'\bА\d{2}-\d+/\d{4}\b'),        
]

# Достаточно совпадения хотя бы 2 паттернов из 7
_LEGAL_THRESHOLD = 2


def detect_doc_type(items: list) -> str:
    """
    Определяет тип документа по первым N элементам.

    Анализирует текст первых 30 элементов (обычно это первые 1-2 страницы).
    Возвращает одну из констант DOC_*.
    """
    # Собираем текст первых 30 элементов
    texts: list[str] = []
    for item, _ in items[:30]:
        raw = getattr(item, "text", None)
        if raw:
            texts.append(raw.strip())

    combined = " ".join(texts)

    # Считаем совпавшие паттерны
    legal_hits = sum(1 for pat in _LEGAL_PATTERNS if pat.search(combined))

    if legal_hits >= _LEGAL_THRESHOLD:
        return DOC_LEGAL

    return DOC_GENERIC
