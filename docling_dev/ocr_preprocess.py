"""Предобработка изображения перед EasyOCR — поднимает качество распознавания
плохих сканов (наклон, серый фон, шум, бледность).

КАК ВСТРАИВАЕТСЯ: EasyOCR вызывается ВНУТРИ Docling (EasyOcrModel.__call__ -
reader.readtext(im)). Перехватываем `easyocr.Reader.readtext` (monkeypatch) и
предобрабатываем numpy-изображение перед распознаванием. Одна точка покрывает и
Docling-OCR, и word_order-проход.

ВАЖНО: EasyOCR обучен на полутоновых/RGB изображениях — ЖЁСТКАЯ бинаризация может
УХУДШИТЬ распознавание (детектор CRAFT использует градации). Поэтому по умолчанию
предобработка ЛЁГКАЯ: deskew + мягкий denoise + CLAHE-контраст; бинаризация —
опция (binarize=True), включать только если проверено, что помогает.

Полностью офлайн (только cv2/numpy, уже в зависимостях). Если cv2 недоступен —
no-op (OCR работает как раньше).
"""
from __future__ import annotations

import logging
from typing import TYPE_CHECKING

log = logging.getLogger(__name__)

if TYPE_CHECKING:
    # Для статического анализатора — реальные модули (без union с None),
    # чтобы cv2.*/np.* разрешались. В рантайме работает try/except ниже.
    import cv2
    import numpy as np
else:
    try:
        import cv2
        import numpy as np
    except Exception:                   # cv2/numpy всегда есть, но на всякий
        cv2 = None
        np = None

_PATCHED = False


def _estimate_skew_deg(gray) -> float:
    """Оценивает угол наклона текста в градусах (через minAreaRect тёмных точек).
    Возвращает 0.0, если оценить нельзя."""
    try:
        thr = cv2.threshold(gray, 0, 255,
                            cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        coords = np.column_stack(np.where(thr > 0))
        if coords.shape[0] < 200:
            return 0.0
        angle = cv2.minAreaRect(coords)[-1]
        # minAreaRect возвращает угол в (-90, 0]; нормируем к (-45, 45]
        if angle < -45:
            angle = 90.0 + angle
        elif angle > 45:
            angle = angle - 90.0
        return float(angle)
    except Exception:
        return 0.0


def _rotate(img, angle: float):
    """Поворот изображения на angle градусов с белым фоном."""
    h, w = img.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), angle, 1.0)
    border = 255 if img.ndim == 2 else (255, 255, 255)
    return cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=border)


def preprocess_for_ocr(
    image,
    deskew: bool = True,
    denoise: bool = False,   # medianBlur размывает текст и СНИЖАЕТ confidence на всех
    clahe: bool = True,      # сканах (A/B-тест) — по умолчанию выкл.
    binarize: bool = False,
    max_skew_deg: float = 7.0,
):
    """Предобрабатывает numpy-изображение для OCR. Возвращает RGB numpy (3 канала).

    Не-numpy вход (путь/None) или отсутствие cv2 - возвращаем как есть (no-op).
    max_skew_deg — углы больше считаем ошибкой детекции (не вращаем).
    """
    if cv2 is None or np is None or not isinstance(image, np.ndarray):
        return image
    try:
        img = image
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY) if img.ndim == 3 else img.copy()

        # 1) Deskew — выравниваем наклон (только небольшие углы 0.2..max_skew_deg)
        if deskew:
            angle = _estimate_skew_deg(gray)
            if 0.2 < abs(angle) <= max_skew_deg:
                gray = _rotate(gray, angle)
                log.debug("ocr_preprocess: deskew на %.2f°", angle)

        # 2) Denoise — медианный фильтр 3×3: быстро убирает точки «соль-перец»
        # (которые читаются как «[», «]», «1»), сохраняя штрихи текста ≥3px.
        # NL-means точнее, но ~1s/страницу — для батча неприемлемо.
        if denoise:
            gray = cv2.medianBlur(gray, 3)

        # 3) CLAHE — адаптивный контраст (вытягивает бледный текст)
        if clahe:
            gray = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)

        # 4) Бинаризация — ОПЦИЯ (по умолчанию выкл; может навредить EasyOCR)
        if binarize:
            gray = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 31, 15)

        return cv2.cvtColor(gray, cv2.COLOR_GRAY2RGB)
    except Exception as exc:
        log.debug("ocr_preprocess: пропуск (ошибка: %s)", exc)
        return image


def install_easyocr_preprocess(
    deskew: bool = True,
    denoise: bool = False,
    clahe: bool = True,
    binarize: bool = False,
) -> bool:
    """Monkeypatch easyocr.Reader.readtext: предобрабатывает изображение перед OCR.
    Идемпотентно. Возвращает True если патч установлен."""
    global _PATCHED
    if _PATCHED:
        return True
    if cv2 is None:
        log.warning("ocr_preprocess: cv2 недоступен — предобработка отключена")
        return False
    try:
        import easyocr
    except Exception as exc:
        log.warning("ocr_preprocess: easyocr недоступен (%s)", exc)
        return False

    _orig_readtext = easyocr.Reader.readtext

    def _readtext_pre(self, image, *args, **kwargs):
        image = preprocess_for_ocr(image, deskew=deskew, denoise=denoise,
                                   clahe=clahe, binarize=binarize)
        return _orig_readtext(self, image, *args, **kwargs)

    easyocr.Reader.readtext = _readtext_pre
    _PATCHED = True
    log.info("ocr_preprocess: предобработка включена (deskew=%s denoise=%s clahe=%s binarize=%s)",
             deskew, denoise, clahe, binarize)
    return True
