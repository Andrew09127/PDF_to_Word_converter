"""Геометрические утилиты для работы с bbox Docling."""
from __future__ import annotations


def bbox_h(bbox) -> float:
    """Высота bbox в PDF-точках, независимо от системы координат."""
    return abs(float(getattr(bbox, "t", 0)) - float(getattr(bbox, "b", 0)))


def bbox_mid_y(bbox) -> float:
    """Вертикальная середина bbox."""
    return (float(getattr(bbox, "t", 0)) + float(getattr(bbox, "b", 0))) / 2.0


def bbox_x0(bbox) -> float:
    return float(getattr(bbox, "l", 0))


def bbox_x1(bbox) -> float:
    return float(getattr(bbox, "r", 0))


def coplanar(bbox_a, bbox_b, tolerance: float = 60.0) -> bool:
    """True если два bbox перекрываются по вертикали в пределах tolerance pt."""
    if bbox_a is None or bbox_b is None:
        return False
    a_vals = (float(getattr(bbox_a, "t", 0)), float(getattr(bbox_a, "b", 0)))
    b_vals = (float(getattr(bbox_b, "t", 0)), float(getattr(bbox_b, "b", 0)))
    a_top, a_bot = min(a_vals), max(a_vals)
    b_top, b_bot = min(b_vals), max(b_vals)
    return (a_top - tolerance) < b_bot and (b_top - tolerance) < a_bot


def detect_pdf_native(items: list) -> bool:
    """
    Определяет систему координат Docling.
    PDF-native: y=0 снизу, значит t > b для обычных bbox.
    Screen: y=0 сверху, значит t < b.
    """
    for item, _ in items[:10]:
        pv = (getattr(item, "prov", None) or [None])[0]
        if pv:
            bx = getattr(pv, "bbox", None)
            if bx:
                t = float(getattr(bx, "t", 0))
                b = float(getattr(bx, "b", 0))
                if abs(t - b) > 2:
                    return t > b
    return True


def reading_order_key(item_level: tuple, pdf_native: bool) -> tuple:
    """
    Ключ сортировки для порядка чтения: страница ↑, сверху вниз, слева направо.
    Y квантуется в бины по 40 pt — items на одной строке сортируются по X.
    """
    item, _ = item_level
    prov = (getattr(item, "prov", None) or [None])[0]
    if prov is None:
        return (9999, 0.0, 0.0)
    pg   = int(getattr(prov, "page_no", 9999))
    bx   = getattr(prov, "bbox", None)
    if bx is None:
        return (pg, 0.0, 0.0)
    mid = bbox_mid_y(bx)
    x0  = bbox_x0(bx)
    y_bin = round((-mid if pdf_native else mid) / 40) * 40
    return (pg, y_bin, x0)
