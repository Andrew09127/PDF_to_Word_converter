"""
page_analyser.py
Анализ страниц документа: тип страницы, колонки, медианы.
Не зависит от типа документа — универсальный модуль.
"""
from __future__ import annotations

import statistics
from dataclasses import dataclass, field

from .geometry import bbox_h, bbox_mid_y, bbox_x0, bbox_x1


# Типы страниц 

PAGE_TEXT       = "text"        # обычная текстовая страница
PAGE_IMAGE      = "image"       # страница целиком как картинка (нет текст-блоков)
PAGE_MULTICOLUMN = "multicolumn" # два или более столбцов текста
PAGE_TABLE_HEAVY = "table_heavy" # преимущественно таблицы


@dataclass
class ColumnBounds:
    """Границы одной колонки на странице."""
    x0: float
    x1: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2.0

    def contains(self, x: float) -> bool:
        return self.x0 - 10 <= x <= self.x1 + 10


@dataclass
class PageInfo:
    """Метаданные одной страницы."""
    page_no:    int
    width:      float
    height:     float
    page_type:  str               = PAGE_TEXT
    columns:    list[ColumnBounds] = field(default_factory=list)
    median_h:   float             = 0.0
    left_min:   float             = 0.0

    @property
    def is_multicolumn(self) -> bool:
        return self.page_type == PAGE_MULTICOLUMN

    @property
    def is_image_only(self) -> bool:
        return self.page_type == PAGE_IMAGE

    def column_for_x(self, x: float) -> int:
        """Возвращает индекс колонки для координаты x. -1 если не найдено."""
        for i, col in enumerate(self.columns):
            if col.contains(x):
                return i
        # Фолбэк: ближайшая колонка
        if self.columns:
            return min(range(len(self.columns)),
                       key=lambda i: abs(self.columns[i].cx - x))
        return 0


_BODY_LABELS = frozenset({"paragraph", "text", "list_item"})
_TABLE_LABELS = frozenset({"table"})


def analyse_pages(items: list) -> dict[int, PageInfo]:
    """
    Анализирует все элементы документа и возвращает PageInfo для каждой страницы.

    Определяет:
    - median_h: медианная высота текстовых bbox
    - left_min: 10-й перцентиль левого края (эффективное левое поле)
    - page_type: text / image / multicolumn / table_heavy
    - columns: границы колонок если multicolumn
    """
    # Сбор сырых данных по страницам
    heights_by_page: dict[int, list[float]] = {}
    lefts_by_page:   dict[int, list[float]] = {}
    rights_by_page:  dict[int, list[float]] = {}
    x0s_by_page:     dict[int, list[float]] = {}
    table_count:     dict[int, int]         = {}
    text_count:      dict[int, int]         = {}
    page_sizes:      dict[int, tuple[float, float]] = {}

    for item, _ in items:
        raw = getattr(item, "label", None)
        if raw is None:
            continue
        label = (raw.value if hasattr(raw, "value") else str(raw)).lower()

        prov_list = getattr(item, "prov", None) or []
        if not prov_list:
            continue
        prov = prov_list[0]
        bbox = getattr(prov, "bbox", None)
        page = int(getattr(prov, "page_no", 1))

        # Размер страницы из prov
        pw = float(getattr(prov, "page_w", 0) or 0)
        ph = float(getattr(prov, "page_h", 0) or 0)
        if pw > 0 and ph > 0:
            page_sizes[page] = (pw, ph)

        if label in _TABLE_LABELS:
            table_count[page] = table_count.get(page, 0) + 1
            continue

        if label not in _BODY_LABELS:
            continue

        if bbox is None:
            continue

        h  = bbox_h(bbox)
        x0 = bbox_x0(bbox)
        x1 = bbox_x1(bbox)
        if h < 2:
            continue

        text_count[page] = text_count.get(page, 0) + 1
        heights_by_page.setdefault(page, []).append(h)
        if x0 >= 0:
            lefts_by_page.setdefault(page, []).append(x0)
        if x1 > 0:
            rights_by_page.setdefault(page, []).append(x1)
        x0s_by_page.setdefault(page, []).append(x0)

    # Строим PageInfo для каждой страницы
    all_pages = set(heights_by_page) | set(table_count) | set(text_count)
    result: dict[int, PageInfo] = {}

    for page in all_pages:
        pw, ph = page_sizes.get(page, (595.0, 842.0))
        info = PageInfo(page_no=page, width=pw, height=ph)

        # Медиана высоты
        hs = heights_by_page.get(page, [])
        info.median_h = statistics.median(hs) if hs else 0.0

        # Левый минимум (10-й перцентиль)
        ls = lefts_by_page.get(page, [])
        if ls:
            ls_sorted = sorted(ls)
            info.left_min = ls_sorted[max(0, len(ls_sorted) // 10)]

        # Тип страницы
        n_text  = text_count.get(page, 0)
        n_table = table_count.get(page, 0)

        if n_text == 0 and n_table == 0:
            info.page_type = PAGE_IMAGE
        elif n_table > 2 and n_text < 5:
            info.page_type = PAGE_TABLE_HEAVY
        else:
            # Детекция колонок по распределению x0
            x0s = x0s_by_page.get(page, [])
            cols = _detect_columns(x0s, pw)
            if len(cols) >= 2:
                info.page_type = PAGE_MULTICOLUMN
                info.columns   = cols
            else:
                info.page_type = PAGE_TEXT
                # Одна колонка — от left_min до правого края
                if pw > 0:
                    info.columns = [ColumnBounds(x0=info.left_min, x1=pw)]

        result[page] = info

    return result


def _detect_columns(x0s: list[float], page_width: float) -> list[ColumnBounds]:
    """
    Определяет колонки по распределению x0 координат блоков.

    Алгоритм:
    1. Строит гистограмму x0 с шагом 5% ширины страницы
    2. Ищет «пики» (зоны концентрации блоков) разделённые «долинами»
    3. Если пиков >= 2 и они достаточно разделены → multicolumn

    Минимальное разделение колонок: 20% ширины страницы.
    """
    if len(x0s) < 6 or page_width <= 0:
        return []

    # Нормализуем x0 к [0, 1]
    norm = [x / page_width for x in x0s]

    # Гистограмма с шагом 0.05 (20 бинов)
    N_BINS = 20
    bins   = [0] * N_BINS
    for v in norm:
        b = min(int(v * N_BINS), N_BINS - 1)
        bins[b] += 1

    # Находим пики — бины с числом > 15% от максимума
    max_count = max(bins)
    if max_count < 2:
        return []
    threshold = max(max_count * 0.15, 1)

    # Объединяем соседние ненулевые бины в группы
    groups: list[list[int]] = []
    current: list[int] = []
    for i, cnt in enumerate(bins):
        if cnt >= threshold:
            current.append(i)
        else:
            if current:
                groups.append(current)
                current = []
    if current:
        groups.append(current)

    if len(groups) < 2:
        return []

    # Минимальное расстояние между центрами групп: 20% ширины
    MIN_SEP = 0.20
    group_centers = [(sum(g) / len(g)) / N_BINS for g in groups]
    # Проверяем что группы достаточно разделены
    for i in range(len(group_centers) - 1):
        if group_centers[i + 1] - group_centers[i] < MIN_SEP:
            return []

    # Строим ColumnBounds: от начала группы до начала следующей
    cols: list[ColumnBounds] = []
    for gi, group in enumerate(groups):
        col_x0 = (group[0] / N_BINS) * page_width - 5
        if gi + 1 < len(groups):
            col_x1 = (groups[gi + 1][0] / N_BINS) * page_width - 5
        else:
            col_x1 = page_width
        cols.append(ColumnBounds(x0=max(col_x0, 0), x1=col_x1))

    return cols
