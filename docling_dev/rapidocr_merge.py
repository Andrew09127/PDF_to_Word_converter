"""Объединённая детекция RapidOCR на нескольких масштабах (повышение recall).

Проблема: det-модель RapidOCR на части документов НЕ находит
часть текстовых строк — текст теряется. Параметры детекции это не лечат: одни строки 
находятся на одном масштабе, другие — на другом. Объединение проходов на 2 масштабах с
нижает потери, оставаясь ЧИСТЫМ (распознаёт всё та же eslav rec-модель, без шума EasyOCR).

Механизм: оборачиваем `reader` (экземпляр RapidOCR), который Docling зовёт как
`reader(im, use_det, use_cls, use_rec)`. Обёртка гоняет полный проход (det+cls+rec)
на исходном изображении И на уменьшенной копии, переводит координаты найденных строк
обратно в систему исходного изображения и добавляет те строки, которых на первом
проходе НЕ было (центр бокса не попал ни в один уже найденный бокс). Возвращает
объект с тем же интерфейсом (.boxes/.txts/.scores), что ждёт Docling.

Полностью локально, без сети. Цена — проход OCR кратно числу масштабов (медленнее),
но приоритет проекта: полный чистый текст, скорость не критична.
"""
from __future__ import annotations

import logging

import numpy as np

log = logging.getLogger(__name__)

# Масштабы относительно изображения, которое передаёт Docling (у него scale=3 - 216
# DPI). 1.0 — как есть; 0.667 ≈ эквивалент scale=2 (144 DPI). Эти два прохода в
# замерах давали максимум прироста (добавление 4-го масштаба уже ничего не давало).
_MERGE_SCALES: tuple[float, ...] = (1.0, 0.667)


class _MergedResult:
    """Минимальный результат в формате RapidOCR: .boxes (Nx4x2), .txts, .scores."""

    def __init__(self, boxes: list, txts: list, scores: list):
        self.boxes = np.array(boxes, dtype=float) if boxes else None
        self.txts = txts
        self.scores = scores


def _box_center(box) -> tuple[float, float]:
    pts = np.asarray(box, dtype=float)
    return float(pts[:, 0].mean()), float(pts[:, 1].mean())


def _box_bounds(box) -> tuple[float, float, float, float]:
    pts = np.asarray(box, dtype=float)
    return (float(pts[:, 0].min()), float(pts[:, 1].min()),
            float(pts[:, 0].max()), float(pts[:, 1].max()))


def _center_inside(cx: float, cy: float, bounds: tuple[float, float, float, float]) -> bool:
    x0, y0, x1, y1 = bounds
    return x0 <= cx <= x1 and y0 <= cy <= y1


class MergedReader:
    """Обёртка над RapidOCR: объединяет детекцию на нескольких масштабах."""

    def __init__(self, base_reader, scales: tuple[float, ...] = _MERGE_SCALES):
        self._base = base_reader
        self._scales = scales

    def __call__(self, im, use_det=None, use_cls=None, use_rec=None):
        try:
            import cv2
        except Exception:
            cv2 = None

        kept_boxes: list = []
        kept_txts: list = []
        kept_scores: list = []
        kept_bounds: list = []

        for k, s in enumerate(self._scales):
            if s == 1.0 or cv2 is None:
                img = im
                inv = 1.0
            else:
                img = cv2.resize(im, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
                inv = 1.0 / s
            try:
                r = self._base(img, use_det=use_det, use_cls=use_cls, use_rec=use_rec)
            except Exception as exc:
                log.debug("merge: проход scale=%.3f упал: %s", s, exc)
                continue
            if r is None or getattr(r, "boxes", None) is None:
                continue
            for box, txt, score in zip(r.boxes.tolist(), r.txts, r.scores):
                # координаты найденного бокса - в систему ИСХОДНОГО изображения
                box = (np.asarray(box, dtype=float) * inv).tolist()
                if k == 0:
                    # первый (полноразмерный) проход — берём всё как базу
                    kept_boxes.append(box); kept_txts.append(txt); kept_scores.append(score)
                    kept_bounds.append(_box_bounds(box))
                else:
                    # последующие проходы — добавляем только НОВЫЕ строки (центр бокса
                    # не попал ни в один уже найденный) — то, что первый проход пропустил
                    cx, cy = _box_center(box)
                    if any(_center_inside(cx, cy, b) for b in kept_bounds):
                        continue
                    kept_boxes.append(box); kept_txts.append(txt); kept_scores.append(score)
                    kept_bounds.append(_box_bounds(box))

        if not kept_boxes:
            return _MergedResult([], [], [])
        log.debug("merge: итог строк=%d (масштабы %s)", len(kept_boxes), self._scales)
        return _MergedResult(kept_boxes, kept_txts, kept_scores)


def install_merged_detection() -> None:
    """Патчит docling RapidOcrModel: после создания оборачивает self.reader в
    MergedReader. Идемпотентно. Если docling/rapidocr недоступны — тихо пропускает."""
    try:
        from docling.models.stages.ocr.rapid_ocr_model import RapidOcrModel
    except Exception as exc:
        log.debug("merge: docling RapidOcrModel недоступен (%s) — пропуск", exc)
        return

    if getattr(RapidOcrModel, "_merged_detection_installed", False):
        return

    _orig_init = RapidOcrModel.__init__

    def _patched_init(self, *args, **kwargs):
        _orig_init(self, *args, **kwargs)
        reader = getattr(self, "reader", None)
        if reader is not None and not isinstance(reader, MergedReader):
            self.reader = MergedReader(reader)
            log.info("merge: объединённая детекция RapidOCR включена (масштабы %s)",
                     _MERGE_SCALES)

    setattr(RapidOcrModel, "__init__", _patched_init)
    setattr(RapidOcrModel, "_merged_detection_installed", True)
