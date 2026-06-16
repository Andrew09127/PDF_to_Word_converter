"""Анализ насыщенности штриха по изображению — определение ЖИРНОСТИ на сканах.

На отсканированных PDF Docling `formatting.bold` недоступен (OCR не даёт стиль),
поэтому жирность приходится оценивать по картинке: у жирного шрифта штрих толще
→ выше доля тёмных пикселей в зоне текста.

Метрики (block_ink_stats):
  ink_ratio      — доля тёмных пикселей во всём кропе (грубая).
  stroke_density — доля тёмных среди ТЕКСТОВЫХ строк пикселей (нормирует на
                   межстрочье); путает жирность с плотностью символов.
  stroke_w       — МЕДИАНА длин горизонтальных тёмных пробегов (в пикселях),
                   исключая слишком длинные (подчёркивания, рамки, «_», «@»,
                   длинная латиница дают линии-выбросы). Это устойчивая оценка
                   ТОЛЩИНЫ штриха: у жирного шрифта вертикальные штрихи толще →
                   медиана пробегов выше; «плотные, но тонкие» цифры дают много
                   КОРОТКИХ пробегов → медиана низкая. Отделяет жирность от
                   плотности И устойчива к линиям-артефактам — ИСПОЛЬЗУЕМ её.
  mean_run       — среднее тех же пробегов (для отладки; чувствительно к выбросам).

Решение «жирный» принимает converter.build_docx: сравнивает stroke_w блока с
МЕДИАНОЙ тела страницы (см. _estimate_ink_medians). Здесь — только измерение.
"""
from __future__ import annotations

import logging

log = logging.getLogger(__name__)

try:
    import numpy as _np
except Exception:                       # numpy всегда есть (зависимость docling), но на всякий
    _np = None


def block_ink_stats(
    pil_crop,
    dark_threshold: int = 160,
    row_text_frac: float = 0.04,
    max_run_px: int = 12,
) -> dict | None:
    """Считает насыщенность/толщину штриха для кропа текстового блока.

    Возвращает {"ink_ratio", "stroke_density", "stroke_w", "mean_run"} или None,
    если изображение недоступно/слишком мелкое.

    dark_threshold — пиксель «тёмный» (штрих) при яркости < порога (0–255).
    row_text_frac  — строка считается «текстовой», если доля тёмных ≥ этого (отсекает межстрочье).
    max_run_px     — пробеги длиннее (подчёркивания/рамки/линии) НЕ учитываются в
                     толщине штриха (это не буквы). Калибровано под ~144 DPI.
    """
    if pil_crop is None or _np is None:
        return None
    try:
        arr = _np.asarray(pil_crop.convert("L"))
    except Exception:
        return None
    if arr.ndim != 2:
        return None
    h, w = arr.shape
    if h < 3 or w < 3:
        return None

    dark = arr < dark_threshold                 # bool H×W: штрих
    total_px = h * w
    total_dark = int(dark.sum())
    ink_ratio = total_dark / total_px
    if total_dark == 0:
        return {"ink_ratio": 0.0, "stroke_density": 0.0,
                "stroke_w": 0.0, "mean_run": 0.0}

    row_frac = dark.mean(axis=1)                 # доля тёмных в каждой строке
    text_rows = row_frac >= row_text_frac
    n_text_rows = int(text_rows.sum())
    if n_text_rows > 0:
        stroke_density = float(dark[text_rows].sum()) / (n_text_rows * w)
    else:
        stroke_density = ink_ratio              # текстовых строк не нашли — грубая оценка

    # Длины горизонтальных тёмных пробегов (run-length по строкам). Между строками
    # вставляем столбец False, чтобы конец одной строки не сливался с началом другой.
    sep    = _np.zeros((h, 1), dtype=bool)
    padded = _np.hstack([dark, sep]).ravel()
    diff   = _np.diff(padded.astype(_np.int8))
    starts = _np.where(diff == 1)[0] + 1
    ends   = _np.where(diff == -1)[0] + 1
    if padded[0]:                               # пробег, начинающийся в самом начале
        starts = _np.concatenate(([0], starts))
    runs = ends - starts
    if runs.size == 0:
        return {"ink_ratio": ink_ratio, "stroke_density": stroke_density,
                "stroke_w": 0.0, "mean_run": 0.0}
    short = runs[runs <= max_run_px]            # штрихи букв (без линий-артефактов)
    use   = short if short.size > 0 else runs
    stroke_w = float(_np.median(use))           # устойчивая толщина штриха
    mean_run = float(use.mean())
    return {"ink_ratio": ink_ratio, "stroke_density": stroke_density,
            "stroke_w": stroke_w, "mean_run": mean_run}
