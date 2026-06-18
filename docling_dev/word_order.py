"""
word_order.py
Реконструкция правильного порядка слов из результатов EasyOCR.

Проблема: Docling/EasyOCR на скане иногда возвращает текстовые регионы
в неправильном порядке (нечётные строки в один регион, чётные — в другой).
Это модуль решает проблему, работая на уровне отдельных слов с их bbox.

Алгоритм:
  1. Получаем слова с координатами от EasyOCR (формат readtext detail=1)
  2. Сортируем слова: сверху вниз (Y), внутри строки — слева направо (X)
  3. Группируем в визуальные строки по Y-близости
  4. Группируем строки в параграфы по Y-разрыву
  5. Возвращаем список TextBlock с правильным порядком текста
"""
from __future__ import annotations

import statistics
from dataclasses import dataclass, field


# Структуры данных

@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    conf: float = 1.0

    @property
    def mid_y(self) -> float:
        return (self.y0 + self.y1) / 2.0

    @property
    def mid_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def height(self) -> float:
        return abs(self.y1 - self.y0)


@dataclass
class VisualLine:
    words: list[Word] = field(default_factory=list)

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def mid_y(self) -> float:
        if not self.words:
            return 0.0
        return statistics.mean(w.mid_y for w in self.words)

    @property
    def x0(self) -> float:
        return min(w.x0 for w in self.words) if self.words else 0.0

    @property
    def x1(self) -> float:
        return max(w.x1 for w in self.words) if self.words else 0.0

    @property
    def median_height(self) -> float:
        hs = [w.height for w in self.words if w.height > 2]
        return statistics.median(hs) if hs else 10.0


@dataclass
class TextBlock:
    lines: list[VisualLine] = field(default_factory=list)

    @property
    def text(self) -> str:
        return " ".join(line.text for line in self.lines)

    @property
    def x0(self) -> float:
        return min(l.x0 for l in self.lines) if self.lines else 0.0

    @property
    def x1(self) -> float:
        return max(l.x1 for l in self.lines) if self.lines else 0.0

    @property
    def mid_y(self) -> float:
        if not self.lines:
            return 0.0
        return statistics.mean(l.mid_y for l in self.lines)

    @property
    def font_height(self) -> float:
        hs = [l.median_height for l in self.lines]
        return statistics.median(hs) if hs else 10.0

    @property
    def line_count(self) -> int:
        return len(self.lines)


# Основная функция

def reconstruct_blocks(
    ocr_results: list[tuple],
    pdf_native: bool = True,
    min_conf: float = 0.25,
    line_merge_ratio: float = 0.55,
    para_gap_ratio: float = 1.8,
) -> list[TextBlock]:
    """
    Конвертирует результаты EasyOCR readtext(detail=1) в список TextBlock
    с правильным порядком слов.

    Параметры:
        ocr_results   — вывод reader.readtext(detail=1): [(pts, text, conf), ...]
        pdf_native    — True если y=0 снизу страницы (PDF-координаты)
        min_conf      — минимальная уверенность EasyOCR для включения слова
        line_merge_ratio — слова с |ΔmidY| < ratio * median_h попадают в одну строку
        para_gap_ratio   — разрыв > ratio * median_spacing → новый параграф
    """
    # 1. Фильтрация и создание объектов Word
    words: list[Word] = []
    for pts, text, conf in ocr_results:
        if conf < min_conf or not text.strip():
            continue
        xs = [float(p[0]) for p in pts]
        ys = [float(p[1]) for p in pts]
        words.append(Word(
            text=text.strip(),
            x0=min(xs), y0=min(ys),
            x1=max(xs), y1=max(ys),
            conf=conf,
        ))

    if not words:
        return []

    # 2. Медианная высота слова — порог для группировки строк
    heights = [w.height for w in words if w.height > 2]
    if not heights:
        return []
    median_h = statistics.median(heights)
    line_tol  = median_h * line_merge_ratio

    # 3. Сортировка: сверху вниз - слева направо
    #    PDF-native (y=0 снизу): большой mid_y = верх страницы - убывающий порядок
    #    Screen-coords (y=0 сверху): малый mid_y = верх - возрастающий порядок
    words.sort(key=lambda w: (-w.mid_y if pdf_native else w.mid_y))

    # 4. Группировка слов в визуальные строки
    # Используем running mean O(n) вместо statistics.mean O(n²)
    lines: list[VisualLine] = []
    cur_line  = VisualLine(words=[words[0]])
    cur_y     = words[0].mid_y
    cur_count = 1

    for w in words[1:]:
        if abs(w.mid_y - cur_y) <= line_tol:
            cur_line.words.append(w)
            cur_count += 1
            cur_y = cur_y + (w.mid_y - cur_y) / cur_count  # O(1) running mean
        else:
            cur_line.words.sort(key=lambda ww: ww.x0)
            lines.append(cur_line)
            cur_line  = VisualLine(words=[w])
            cur_y     = w.mid_y
            cur_count = 1

    cur_line.words.sort(key=lambda ww: ww.x0)
    lines.append(cur_line)

    if not lines:
        return []

    # 5. Медиана межстрочного интервала для порога параграфа
    spacings = []
    for i in range(1, len(lines)):
        gap = abs(lines[i].mid_y - lines[i - 1].mid_y)
        if gap < median_h * 4:
            spacings.append(gap)

    median_spacing = statistics.median(spacings) if spacings else median_h * 1.4
    para_threshold = median_spacing * para_gap_ratio

    # 6. Группировка строк в параграфы
    blocks: list[TextBlock] = []
    cur_block = TextBlock(lines=[lines[0]])

    for i in range(1, len(lines)):
        gap = abs(lines[i].mid_y - lines[i - 1].mid_y)
        if gap > para_threshold:
            blocks.append(cur_block)
            cur_block = TextBlock()
        cur_block.lines.append(lines[i])

    blocks.append(cur_block)
    return [b for b in blocks if b.lines]


# Вспомогательные функции 

def get_page_ocr(dl_doc, page_no: int, ocr_reader) -> tuple[list[tuple], float]:
    """
    Запускает EasyOCR на изображении страницы из DoclingDocument.
    Возвращает (список (pts, text, conf), высота изображения в пикселях).
    Возвращает ([], 0.0) если изображение недоступно.
    """
    import numpy as np

    # Попытка получить изображение из Docling
    page = dl_doc.pages.get(page_no)
    pil_img = None
    if page is not None:
        img_ref = getattr(page, "image", None)
        if img_ref is not None:
            pil_img = getattr(img_ref, "pil_image", None)

    if pil_img is None:
        return [], 0.0

    img_h   = float(pil_img.height)
    img_np  = np.array(pil_img)
    results = ocr_reader.readtext(img_np, detail=1, paragraph=False)
    return results, img_h


def blocks_for_page(
    dl_doc,
    page_no: int,
    ocr_reader,
    min_conf: float = 0.25,
) -> tuple[list[TextBlock], float]:
    """
    Удобная обёртка: получает OCR страницы и возвращает (TextBlock-и, высота изображения px).
    EasyOCR всегда работает в экранных координатах (y=0 сверху), поэтому pdf_native=False.
    """
    raw, img_h = get_page_ocr(dl_doc, page_no, ocr_reader)
    blocks = reconstruct_blocks(raw, pdf_native=False, min_conf=min_conf)
    return blocks, img_h
