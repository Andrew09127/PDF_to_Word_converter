"""
block_sorter.py
Сортировка блоков Docling в правильный порядок чтения.

Поддерживает:
- одноколоночные страницы (сверху вниз)
- многоколоночные страницы (колонка слева → колонка справа, внутри каждой сверху вниз)
- страницы-картинки (блоки не сортируются)

Не содержит document-specific логики — универсальный модуль.
"""
from __future__ import annotations

import logging

from .geometry import bbox_mid_y, bbox_x0
from .page_analyser import PageInfo, PAGE_IMAGE

log = logging.getLogger(__name__)


def sort_items_by_reading_order(
    items: list,
    page_infos: dict[int, PageInfo],
    pdf_native: bool = True,
) -> list:
    """
    Сортирует все элементы документа в правильный порядок чтения.

    Алгоритм:
    1. Группирует элементы по страницам
    2. Для каждой страницы применяет нужную стратегию сортировки:
       - PAGE_IMAGE: порядок не меняется
       - PAGE_MULTICOLUMN: сначала левая колонка сверху вниз, потом правая
       - Остальные: сверху вниз (с квантованием Y для элементов одной строки)
    3. Сохраняет межстраничный порядок (страница 1 → страница 2 → ...)

    pdf_native=True: y=0 снизу (Docling PDF-режим), инвертируем Y для сортировки.
    """
    if not items:
        return items

    # Группируем по страницам, сохраняя исходные индексы
    pages: dict[int, list[tuple[int, object, object]]] = {}
    no_page: list[tuple[int, object, object]] = []

    for orig_idx, (item, level) in enumerate(items):
        prov_list = getattr(item, "prov", None) or []
        if not prov_list:
            no_page.append((orig_idx, item, level))
            continue
        page_no = int(getattr(prov_list[0], "page_no", 9999))
        pages.setdefault(page_no, []).append((orig_idx, item, level))

    result: list = []

    for page_no in sorted(pages.keys()):
        page_items = pages[page_no]
        info = page_infos.get(page_no)

        if info is None or info.is_image_only:
            # Страница-картинка или нет метаданных — оставляем как есть
            result.extend((item, level) for _, item, level in page_items)
            continue

        if info.is_multicolumn and len(info.columns) >= 2:
            sorted_page = _sort_multicolumn(page_items, info, pdf_native)
        else:
            sorted_page = _sort_single_column(page_items, info, pdf_native)

        result.extend(sorted_page)

    # Элементы без страницы — в конец
    result.extend((item, level) for _, item, level in no_page)
    return result


def _get_y_top(item, page_height: float, pdf_native: bool) -> float:
    """Возвращает Y координату верхнего края блока в screen-координатах (0=верх)."""
    prov_list = getattr(item, "prov", None) or []
    if not prov_list:
        return 0.0
    bbox = getattr(prov_list[0], "bbox", None)
    if bbox is None:
        return 0.0
    mid = bbox_mid_y(bbox)
    if pdf_native:
        return page_height - mid
    return mid


def _sort_single_column(
    page_items: list[tuple[int, object, object]],
    info: PageInfo,
    pdf_native: bool,
) -> list[tuple[object, object]]:
    """Сортировка одноколоночной страницы: сверху вниз, с квантованием Y."""
    BIN_SIZE = 15.0  # pt — блоки в пределах одного бина считаются на одной строке

    def _key(entry):
        _, item, _ = entry
        y = _get_y_top(item, info.height, pdf_native)
        x = bbox_x0(getattr((getattr(item, "prov", None) or [None])[0], "bbox", None) or type('', (), {'l': 0})())
        y_bin = round(y / BIN_SIZE) * BIN_SIZE
        return (y_bin, x)

    sorted_items = sorted(page_items, key=_key)
    return [(item, level) for _, item, level in sorted_items]


def _sort_multicolumn(
    page_items: list[tuple[int, object, object]],
    info: PageInfo,
    pdf_native: bool,
) -> list[tuple[object, object]]:
    """
    Сортировка многоколоночной страницы.

    Каждый блок назначается в колонку по x0 координатe.
    Внутри колонки — сверху вниз.
    Колонки выводятся слева направо.
    """
    BIN_SIZE = 15.0

    columns: dict[int, list[tuple[float, float, object, object]]] = {
        i: [] for i in range(len(info.columns))
    }
    unassigned: list[tuple[float, float, object, object]] = []

    for _, item, level in page_items:
        prov_list = getattr(item, "prov", None) or []
        if not prov_list:
            unassigned.append((0.0, 0.0, item, level))
            continue
        bbox = getattr(prov_list[0], "bbox", None)
        if bbox is None:
            unassigned.append((0.0, 0.0, item, level))
            continue
        x0 = bbox_x0(bbox)
        y  = _get_y_top(item, info.height, pdf_native)
        col_idx = info.column_for_x(x0)
        columns[col_idx].append((y, x0, item, level))

    log.debug("block_sorter multicolumn стр.%d: %s",
              info.page_no,
              {i: len(v) for i, v in columns.items()})

    result: list[tuple[object, object]] = []
    for col_idx in sorted(columns.keys()):
        col_items = sorted(columns[col_idx],
                           key=lambda e: (round(e[0] / BIN_SIZE) * BIN_SIZE, e[1]))
        result.extend((item, level) for _, _, item, level in col_items)

    # Неназначенные элементы — в конец страницы
    result.extend((item, level) for _, _, item, level in unassigned)
    return result
