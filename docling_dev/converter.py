"""
converter.py
────────────
Основная логика конвертации DoclingDocument → DOCX.
word_order.py используется только для ПЕРЕУПОРЯДОЧИВАНИЯ блоков Docling —
текст из Docling остаётся нетронутым (качество OCR Docling лучше сырого EasyOCR).
"""
from __future__ import annotations

import gc
import logging
import re
import shutil
import time
from collections import defaultdict
from io import BytesIO
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from .config import (
    BODY_PT, FONT_NAME, LABEL_HEADING, LABEL_MARGIN_THRESHOLD, LABEL_PT,
    MARGIN_INCH, SKIP_LABELS,
)
from .docx_builder import (
    add_label_content_table, add_sidebyside,
    add_header_row,
    add_table_from_cells, add_table_from_grid,
    analyse_pages, detect_alignment, init_document, split_label_content,
)
from .geometry import (
    bbox_h, bbox_mid_y, bbox_x0, coplanar,
    detect_pdf_native, reading_order_key,
)
from .ocr_fixes import postprocess
from .pipeline import build_converter
from .word_order import TextBlock, blocks_for_page

log = logging.getLogger(__name__)


# ── Переупорядочивание через word_order ───────────────────────────────────────

def _reorder_by_word_order(
    all_items: list,
    word_blocks_map: dict[int, tuple[list, float]],
    page_sizes: dict[int, tuple[float, float]],
) -> list:
    """
    Переставляет ТОЛЬКО body-текст (paragraph/text) по позициям word_order-блоков.

    Картинки, таблицы, заголовки остаются на своих местах — это критично для
    корректной работы add_sidebyside (логотип + текст реквизитов рядом).

    Алгоритм:
      1. Собираем paragraph/text элементы каждой страницы, пропуская:
         - элементы в верхней зоне (топ 15% страницы) — шапка с логотипом
         - страницы без word_order данных
      2. Сортируем по word_order-блокам (block_idx, x0).
      3. Вставляем отсортированные элементы обратно на те же позиции.
         Нетекстовые элементы и элементы шапки не трогаем.

    Константа HEADER_ZONE: топ 15% страницы не переупорядочивается, чтобы
    не сломать логику add_sidebyside (логотип + текст реквизитов рядом).
    """
    _BODY        = frozenset({"paragraph", "text"})
    HEADER_ZONE  = 0.15   # топ 15% — шапка страницы (логотип, реквизиты)
    MIN_MOVE     = 2      # минимум элементов должны сдвинуться (снижено с 3→2: ловит swap двух блоков)
    MAX_Y_DIFF   = 0.04   # макс. разница Y [0–1] для надёжного совпадения блока

    result: list = list(all_items)
    page_body: defaultdict[int, list] = defaultdict(list)

    for i, (item, _level) in enumerate(all_items):
        raw = getattr(item, "label", None)
        if raw is None:
            continue
        label = (raw.value if hasattr(raw, "value") else str(raw)).lower()
        if label not in _BODY:
            continue

        prov_list = getattr(item, "prov", None) or []
        if not prov_list:
            continue
        page_no = int(getattr(prov_list[0], "page_no", 9999))
        if page_no not in word_blocks_map:
            continue

        bbox = getattr(prov_list[0], "bbox", None)
        if bbox is None:
            continue

        blocks, img_h = word_blocks_map[page_no]
        _, ph = page_sizes.get(page_no, (595.0, 842.0))
        if not blocks or img_h <= 0 or ph <= 0:
            continue

        # Нормализуем Y к [0,1]: Docling (y=0 снизу) → (0=верх, 1=низ)
        y_norm = (ph - bbox_mid_y(bbox)) / ph

        # Пропускаем элементы из шапки страницы (логотип, реквизиты)
        if y_norm < HEADER_ZONE:
            continue

        best_bi  = min(range(len(blocks)),
                       key=lambda bi: abs(blocks[bi].mid_y / img_h - y_norm))
        y_diff = abs(blocks[best_bi].mid_y / img_h - y_norm)
        if y_diff > MAX_Y_DIFF:
            raw_text = (getattr(all_items[i][0], "text", "") or "")[:40]
            log.debug("word_order стр.%d: нет совпадения (y_diff=%.3f>%.3f) %r",
                      page_no, y_diff, MAX_Y_DIFF, raw_text)
            continue
        page_body[page_no].append((i, best_bi, bbox_x0(bbox)))

    for page_no, body in page_body.items():
        if len(body) < 2:
            log.debug("word_order стр.%d: только %d блоков — пропуск", page_no, len(body))
            continue

        slots     = sorted(x[0] for x in body)
        wo_sorted = sorted(body, key=lambda x: (x[1], x[2]))

        if [x[0] for x in body] == [x[0] for x in wo_sorted]:
            log.debug("word_order стр.%d: %d блоков — порядок уже верный", page_no, len(body))
            continue

        moved = sum(1 for s, (oi, _, _) in zip(slots, wo_sorted) if s != oi)
        log.info("word_order стр.%d: %d/%d блоков требуют перестановки (MIN_MOVE=%d)",
                 page_no, moved, len(body), MIN_MOVE)
        if moved < MIN_MOVE:
            for s, (oi, bi, x0) in zip(slots, wo_sorted):
                if s != oi:
                    txt = (getattr(all_items[oi][0], "text", "") or "")[:50]
                    log.info("  word_order пропущен (moved=%d<MIN=%d): slot=%d←orig=%d bi=%d %r",
                             moved, MIN_MOVE, s, oi, bi, txt)
            continue

        # Если все перемещаемые блоки указывают на один и тот же bi,
        # сортировка по x0 ненадёжна — пропускаем (оба блока на одной Y-позиции)
        moved_bis = {bi for s, (oi, bi, _) in zip(slots, wo_sorted) if s != oi}
        if len(moved_bis) == 1:
            sole_bi = next(iter(moved_bis))
            log.info("word_order стр.%d: пропуск — %d блоков конкурируют за bi=%d (x0-сортировка ненадёжна)",
                     page_no, moved, sole_bi)
            continue

        log.info("word_order стр.%d: применяем перестановку %d блоков:", page_no, moved)
        for s, (oi, bi, _) in zip(slots, wo_sorted):
            if s != oi:
                txt = (getattr(all_items[oi][0], "text", "") or "")[:50]
                log.info("  слот %d ← orig=%d (bi=%d) %r", s, oi, bi, txt)

        for slot, (orig_idx, _bi, _x0) in zip(slots, wo_sorted):
            result[slot] = all_items[orig_idx]

        log.info("word_order стр.%d: скорректировано %d блоков", page_no, moved)

    return result


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _label_str(item) -> str:
    raw = getattr(item, "label", None)
    if raw is None:
        return ""
    return (raw.value if hasattr(raw, "value") else str(raw)).lower()


def _item_bbox_page(item, current_page: int):
    prov_list = getattr(item, "prov", None) or []
    if not prov_list:
        return None, current_page
    bbox   = getattr(prov_list[0], "bbox", None)
    page_no = int(getattr(prov_list[0], "page_no", current_page))
    return bbox, page_no


# ── Исправление порядка чтения ───────────────────────────────────────────────

_NUM_ITEM_RE  = re.compile(r'^(\d+)\s')
_DATE_START_RE = re.compile(r'^\d{2}\.\d{2}\.\d{4}')  # DD.MM.YYYY в начале блока


def _fix_reading_order(all_items: list) -> list:
    """
    Пост-обработка порядка элементов после word_order:
    1. ALL-CAPS section_header ставится перед своим строчным подзаголовком
    2. Dated text block preceded by mid-sentence block → swap
       (Docling иногда переставляет начало и середину одного абзаца)
    3. Numbered list_items сортируются по ведущей цифре; пробелы в нумерации
       заполняются ненумерованными элементами
    """
    result = list(all_items)
    n      = len(result)

    def _lbl(i):
        return _label_str(result[i][0])

    def _pg(i):
        pv = (getattr(result[i][0], "prov", None) or [None])[0]
        return int(getattr(pv, "page_no", -1)) if pv else -1

    def _txt(i):
        return (getattr(result[i][0], "text", None) or "").strip()

    # Fix 1: ALL-CAPS section_header перед строчным subtitle
    for i in range(n - 1):
        if _lbl(i) != "section_header" or _lbl(i + 1) != "section_header":
            continue
        if _pg(i) != _pg(i + 1):
            continue
        ti, tnext  = _txt(i), _txt(i + 1)
        ai    = [c for c in ti    if c.isalpha()]
        anext = [c for c in tnext if c.isalpha()]
        if ai and ai[0].islower() and anext and all(c.isupper() for c in anext):
            result[i], result[i + 1] = result[i + 1], result[i]
            log.info("fix_order: swap '%s' ↔ '%s'", ti[:40], tnext[:40])

    # Fix 2: текстовый блок с датой DD.MM.YYYY стоит ПОСЛЕ незаконченного блока
    # Docling иногда разбивает один абзац на части и переставляет их:
    #   "Игоревичем (Заемщик) Заемщику..." → "20.12.2019 между..."
    # Корректный порядок: дата → продолжение.
    # Признак неправильного порядка: предыдущий блок не заканчивается на «.»/«!»/«?»
    _TEXT_LABELS = frozenset({"text", "paragraph"})
    for i in range(1, n):
        if _lbl(i) not in _TEXT_LABELS or _lbl(i - 1) not in _TEXT_LABELS:
            continue
        if _pg(i) != _pg(i - 1):
            continue
        curr = _txt(i)
        prev = _txt(i - 1)
        if not _DATE_START_RE.match(curr):
            continue
        # Предыдущий блок не завершает предложение → переставляем
        if prev.endswith(('.', '!', '?', ':', '»')):
            continue
        result[i - 1], result[i] = result[i], result[i - 1]
        log.info("fix_order: date-start swap %r ↔ %r", prev[:50], curr[:50])

    # Fix 3: сортировка numbered list_items + заполнение пробелов ненумерованными
    i = 0
    while i < n:
        if _lbl(i) != "list_item":
            i += 1
            continue
        j, pg = i + 1, _pg(i)
        while j < n and _lbl(j) == "list_item" and _pg(j) == pg:
            j += 1
        run_len = j - i
        if run_len < 3:
            i = j
            continue

        txts       = [_txt(k) for k in range(i, j)]
        numbered   = [(int(m.group(1)), k)
                      for k, t in enumerate(txts)
                      for m in [_NUM_ITEM_RE.match(t)] if m]
        unnumbered = [k for k in range(run_len) if not _NUM_ITEM_RE.match(txts[k])]

        if len(numbered) < max(3, run_len // 2):
            i = j
            continue

        numbered_sorted = sorted(numbered)
        nums_only       = [num for num, _ in numbered]
        sorted_nums     = sorted(nums_only)

        # Есть ли пробелы в нумерации?
        has_gap = any(sorted_nums[k + 1] > sorted_nums[k] + 1
                      for k in range(len(sorted_nums) - 1))

        already_sorted = nums_only == sorted_nums
        if already_sorted and not (has_gap and unnumbered):
            i = j
            continue  # порядок верный, пробелов нет

        # Строим новый порядок: вставляем ненумерованные в пробелы нумерации
        final_order: list[int] = []
        unnum_used  = 0
        prev_num    = 0
        for num, k in numbered_sorted:
            gap_size = num - prev_num - 1
            for _ in range(gap_size):
                if unnum_used < len(unnumbered):
                    gap_k = unnumbered[unnum_used]
                    final_order.append(gap_k)
                    log.info("fix_order: пункт %d пропущен OCR, заполняем %r",
                             prev_num + 1, txts[gap_k][:60])
                    unnum_used += 1
            final_order.append(k)
            prev_num = num
        # Оставшиеся ненумерованные — в конец
        final_order.extend(unnumbered[unnum_used:])

        reordered = [result[i + k] for k in final_order]
        for k, item_lvl in enumerate(reordered):
            result[i + k] = item_lvl

        if not already_sorted:
            log.info("fix_order: sorted %d numbered list items on page %d",
                     len(numbered), pg)
        if has_gap and unnum_used > 0:
            log.info("fix_order: filled %d gap(s) with unnumbered items on page %d",
                     unnum_used, pg)
        i = j

    return result


# ── Основная функция построения DOCX ─────────────────────────────────────────

_LETTERHEAD_STOP_RE = re.compile(
    r"\b("
    r"кредитор|должник|заявлени[ея]|исковое|госпошлин\w*|"
    r"арбитражн\w*\s+управляющ|арбитражн\w*\s+суд|по\s+делу"
    r")\b",
    re.IGNORECASE,
)


def _bbox_top_bottom(bbox, page_height: float, pdf_native: bool) -> tuple[float, float]:
    a = float(getattr(bbox, "t", 0.0))
    b = float(getattr(bbox, "b", 0.0))
    if pdf_native:
        top = page_height - max(a, b)
        bottom = page_height - min(a, b)
    else:
        top = min(a, b)
        bottom = max(a, b)
    return top, bottom


def _is_letterhead_stop(text: str) -> bool:
    return bool(_LETTERHEAD_STOP_RE.search(text.casefold()))


def _render_first_page_letterhead(
    doc,
    dl_doc,
    all_items: list,
    page_sizes: dict[int, tuple[float, float]],
    pdf_native: bool,
    ocr_blocks: tuple | None = None,   # (list[TextBlock], img_h) из word_order
) -> set[int]:
    """Render the top first-page image/text row without assuming logo position."""
    # Первая страница — минимальный ключ в page_sizes (может быть 0 или 1)
    first_page = min(page_sizes.keys()) if page_sizes else 1
    pw, ph = page_sizes.get(first_page, (595.0, 842.0))
    log.info("letterhead: first_page=%s pdf_native=%s pw=%.0f ph=%.0f",
             first_page, pdf_native, pw, ph)
    if ph <= 0 or pw <= 0:
        return set()

    pictures: list[tuple[int, object, object, float, float]] = []
    for idx, (item, _level) in enumerate(all_items):
        if _label_str(item) not in {"picture", "figure", "image"}:
            continue
        bbox, page_no = _item_bbox_page(item, first_page)
        if page_no != first_page or bbox is None:
            continue
        # Принимаем картинку если она в верхних 40% в ЛЮБОЙ системе координат.
        # pdf_native может быть определён неверно — проверяем оба варианта.
        t_val = float(getattr(bbox, "t", 0.0))
        b_val = float(getattr(bbox, "b", 0.0))
        top_if_native = (ph - max(t_val, b_val)) / ph   # если y=0 снизу
        top_if_screen = min(t_val, b_val) / ph          # если y=0 сверху
        if top_if_native > 0.40 and top_if_screen > 0.40:
            continue  # картинка не в верхней части страницы
        top, bottom = _bbox_top_bottom(bbox, ph, pdf_native)
        pictures.append((idx, item, bbox, top, bottom))

    if not pictures:
        log.info("letterhead: картинки не найдены — используется старый код add_sidebyside")
        return set()

    pic_idx, pic_item, pic_bbox, pic_top, pic_bottom = min(pictures, key=lambda x: x[3])
    pic_h = max(pic_bottom - pic_top, 1.0)
    row_top = max(0.0, pic_top - max(pic_h * 0.35, ph * 0.025))
    default_bottom = min(
        ph,
        max(pic_bottom + max(pic_h * 1.5, ph * 0.10), ph * 0.24),
    )
    stop_tops: list[float] = []
    for item, _level in all_items:
        lbl = _label_str(item)
        if lbl not in {"paragraph", "text", "list_item", "caption", "title", "section_header"}:
            continue
        bbox, page_no = _item_bbox_page(item, first_page)
        if page_no != first_page or bbox is None:
            continue
        top, _bottom = _bbox_top_bottom(bbox, ph, pdf_native)
        if top <= pic_bottom:
            continue
        text = postprocess((getattr(item, "text", None) or "").strip())
        if text and _is_letterhead_stop(text):
            stop_tops.append(top)
    if stop_tops:
        # Граница = первый стоп-маркер (без буфера): сам маркер уже исключён
        # через _is_letterhead_stop в цикле сбора text_blocks.
        # Буфер ph*0.01 (~8pt) отрезал строки вплотную к маркеру (тел./факс).
        row_bottom = min(stop_tops)
        row_bottom = max(row_bottom, pic_bottom)
    else:
        # Стоп-маркер не найден — безопасный максимум: верхняя четверть страницы
        row_bottom = min(default_bottom, ph * 0.25)

    text_indices: list[int] = []
    text_blocks: list[tuple[int, object, object, str]] = []
    for idx, (item, _level) in enumerate(all_items):
        if idx == pic_idx:
            continue
        lbl = _label_str(item)
        if lbl not in {"paragraph", "text", "list_item", "caption"}:
            continue
        bbox, page_no = _item_bbox_page(item, first_page)
        if page_no != first_page or bbox is None:
            continue
        top, bottom = _bbox_top_bottom(bbox, ph, pdf_native)
        if top > row_bottom or bottom < row_top:
            continue
        text = postprocess((getattr(item, "text", None) or "").strip())
        if not text or _is_letterhead_stop(text):
            continue
        text_indices.append(idx)
        text_blocks.append((idx, item, bbox, text))

    try:
        pil_img = pic_item.get_image(dl_doc)
    except Exception:
        pil_img = None
    if pil_img is None:
        return set()

    text_w_inch = (pw - 2 * MARGIN_INCH * 72) / 72
    pic_l = float(getattr(pic_bbox, "l", 0.0))
    pic_r = float(getattr(pic_bbox, "r", pic_l))
    pic_w_pt = max(pic_r - pic_l, 1.0)
    img_w_inch = min(max(pic_w_pt / 72, 0.5), text_w_inch)

    left_blocks = []
    right_blocks = []
    # Сортировка в экранных координатах (top=0 сверху) для правильного порядка.
    # bbox_mid_y для PDF-native (y=0 снизу) даёт обратный порядок — используем
    # _bbox_top_bottom, который нормализует к screen coords (top возрастает вниз).
    for _idx, _item, bbox, text in sorted(
        text_blocks,
        key=lambda x: (_bbox_top_bottom(x[2], ph, pdf_native)[0], bbox_x0(x[2]))
    ):
        block = {"text": text, "font_pt": 8.5, "bold": False, "italic": False}
        if float(getattr(bbox, "r", 0.0)) <= pic_l:
            left_blocks.append(block)
        elif bbox_x0(bbox) >= pic_r:
            right_blocks.append(block)
        elif bbox_x0(bbox) < pic_l:
            left_blocks.append(block)
        else:
            right_blocks.append(block)

    cells: list[dict] = []
    if left_blocks:
        left_w = max(pic_l - min(float(getattr(x[2], "l", pic_l)) for x in text_blocks), 72.0)
        cells.append({
            "kind": "text",
            "blocks": left_blocks,
            "width_pt": left_w,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        })

    cells.append({
        "kind": "image",
        "image": pil_img,
        "width_pt": pic_w_pt,
        "image_width_inch": img_w_inch,
        "alignment": detect_alignment(pic_bbox, pw),
    })

    if right_blocks:
        # Ширина правой колонки = от правого края логотипа до правого поля страницы.
        # Предыдущий вариант брал max bbox.r из text_blocks — мог недооценивать,
        # если Docling давал неправильные правые края для строк реквизитов.
        right_w = max(pw - pic_r, 72.0)
        cells.append({
            "kind": "text",
            "blocks": right_blocks,
            "width_pt": right_w,
            "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        })

    # ── OCR-дополнение шапки ─────────────────────────────────────────────────
    # Docling иногда пропускает строки в шапке (напр. тел./факс) из-за слияния
    # нескольких визуальных строк в один элемент. EasyOCR из word_order работает
    # на уровне визуальных строк и может восстановить пропущенные.
    if ocr_blocks is not None and right_blocks:
        ocr_blk_list, img_h = ocr_blocks
        img_w = img_h * pw / ph if ph > 0 else img_h
        x_thresh = (pic_r / pw) * 0.75 if pw > 0 else 0
        y_top_frac = pic_bottom / ph if ph > 0 else 0
        y_bot_frac = row_bottom / ph if ph > 0 else 1

        ocr_lines: list[tuple[float, str]] = []
        for blk in sorted(ocr_blk_list, key=lambda b: b.mid_y):
            y_frac = blk.mid_y / img_h if img_h > 0 else 0
            x_frac = blk.x0 / img_w if img_w > 0 else 0
            if not (y_top_frac <= y_frac <= y_bot_frac):
                continue
            if x_frac < x_thresh:
                continue
            blk_text = postprocess(blk.text)
            if blk_text:
                ocr_lines.append((y_frac, blk_text))
                log.info("letterhead OCR-line y=%.2f: %r", y_frac, blk_text[:70])

        if ocr_lines:
            # Заменяем Docling-блоки EasyOCR-строками: более полный список строк
            right_blocks.clear()
            for _, line_text in ocr_lines:
                right_blocks.append({
                    "text": line_text, "font_pt": 8.5,
                    "bold": False, "italic": False,
                })
            # Пересчитываем ячейку правой колонки (right_w не меняется)
            for cell in cells:
                if cell.get("kind") == "text" and cell.get("alignment") == WD_ALIGN_PARAGRAPH.LEFT:
                    cell["blocks"] = right_blocks
                    break

    log.info("letterhead: left_blocks=%d right_blocks=%d",
             len(left_blocks), len(right_blocks))
    for i, rb in enumerate(right_blocks):
        full_text = rb.get("text", "")
        log.info("  right_block[%d] (%d chars): %r", i, len(full_text), full_text)
    add_header_row(doc, cells, text_w_inch)
    return {pic_idx, *text_indices}


def build_docx(
    dl_doc,
    page_sizes: dict[int, tuple[float, float]],
    ocr_reader=None,
    use_word_order: bool = True,
) -> object:
    """
    Конвертирует DoclingDocument в python-docx Document.

    Параметры:
        dl_doc         — Docling DoclingDocument
        page_sizes     — {page_no: (width_pt, height_pt)}
        ocr_reader     — EasyOCR Reader для word_order (None = отключить)
        use_word_order — True: пересортировать блоки через word_order (текст Docling сохраняется)
    """
    all_items = list(dl_doc.iterate_items())
    log.info("build_docx: %d элементов из Docling, %d страниц",
             len(all_items), len(page_sizes))

    # Базовая сортировка в порядке чтения по Docling-координатам
    pdf_native = detect_pdf_native(all_items)
    log.info("build_docx: pdf_native=%s (система координат Docling)", pdf_native)
    all_items.sort(key=lambda x: reading_order_key(x, pdf_native))

    page_medians, page_left_min = analyse_pages(all_items)
    for pn in sorted(page_sizes):
        log.info("  стр.%d: %.0f×%.0f pt, median_h=%.1f px, left_min=%.1f pt",
                 pn, page_sizes[pn][0], page_sizes[pn][1],
                 page_medians.get(pn, 0.0), page_left_min.get(pn, 0.0))

    # word_order: строим карту блоков {page_no: (blocks, img_h_px)},
    # затем пересортировываем all_items по позициям этих блоков.
    # Текст Docling НЕ заменяем — только меняем порядок элементов.
    word_blocks_map: dict[int, tuple[list[TextBlock], float]] = {}
    if use_word_order and ocr_reader is not None:
        for page_no in page_sizes:
            try:
                blocks, img_h = blocks_for_page(dl_doc, page_no, ocr_reader)
                if blocks:
                    word_blocks_map[page_no] = (blocks, img_h)
            except Exception as exc:
                log.warning("word_order page %d failed: %s", page_no, exc)

        if word_blocks_map:
            all_items = _reorder_by_word_order(all_items, word_blocks_map, page_sizes)
            log.info("word_order: переупорядочено на %d стр.", len(word_blocks_map))

    # Пост-обработка: ALL-CAPS заголовки перед subtitle, numbered list в порядке
    all_items = _fix_reading_order(all_items)

    doc               = init_document()
    last_content_page = -1
    current_page      = -1
    prev_midY: dict[int, float] = {}
    prev_h:    dict[int, float] = {}
    skip_indices: set[int] = set()

    def _page_break(target: int) -> None:
        nonlocal last_content_page
        if last_content_page >= 0 and target > last_content_page:
            pb = doc.add_page_break()
            pb.paragraph_format.space_before  = Pt(0)
            pb.paragraph_format.space_after   = Pt(0)
            pb.paragraph_format.widow_control = False
        last_content_page = target

    # Передаём OCR-блоки первой страницы для дополнения шапки
    first_page_no = min(page_sizes.keys()) if page_sizes else 1
    letterhead_indices = _render_first_page_letterhead(
        doc, dl_doc, all_items, page_sizes, pdf_native,
        ocr_blocks=word_blocks_map.get(first_page_no),
    )
    if letterhead_indices:
        skip_indices.update(letterhead_indices)
        last_content_page = first_page_no
        log.info("letterhead: пропускаем %d элементов шапки", len(letterhead_indices))

    _align_names = {
        WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
        WD_ALIGN_PARAGRAPH.CENTER:  "CENTER",
        WD_ALIGN_PARAGRAPH.LEFT:    "LEFT",
        WD_ALIGN_PARAGRAPH.RIGHT:   "RIGHT",
    }

    for idx, (item, _level) in enumerate(all_items):
        if idx in skip_indices:
            continue

        lbl = _label_str(item)
        if not lbl or lbl in SKIP_LABELS:
            log.debug("[%d] skip lbl=%r", idx, lbl)
            continue

        bbox, page_no = _item_bbox_page(item, max(current_page, 1))
        current_page  = page_no

        pw, ph    = page_sizes.get(page_no, (595.0, 842.0))
        median_h  = page_medians.get(page_no, 0.0)
        item_h    = bbox_h(bbox) if bbox is not None else 0.0

        font_pt = LABEL_PT.get(lbl, BODY_PT)
        if lbl in ("paragraph", "text") and item_h > 2 and median_h > 0:
            ratio = item_h / median_h
            # Масштабируем только если явно меньше 0.55 медианы.
            # НЕ масштабируем если это будет label:content — у них bbox часто мал
            # из-за того что Docling даёт bbox только на метку, а не на всё содержимое.
            raw_x0 = float(getattr(bbox, "l", 0)) if bbox is not None else 0.0
            is_potential_label = raw_x0 <= pw * LABEL_MARGIN_THRESHOLD
            if ratio < 0.55 and not is_potential_label:
                font_pt = max(round(BODY_PT * ratio * 2) / 2, 9.0)

        # ── Таблицы ──────────────────────────────────────────────────────────
        if lbl == "table":
            try:
                data = getattr(item, "data", None)
                if data is None:
                    log.debug("[стр%d] table: нет data — пропуск", page_no)
                    continue
                nr = getattr(data, "num_rows", len(getattr(data, "grid", [])))
                nc = getattr(data, "num_cols",
                             max((len(r) for r in getattr(data, "grid", [[]])), default=0))
                log.info("[стр%d] table: %d строк × %d столбцов", page_no, nr, nc)
                _page_break(page_no)
                if hasattr(data, "grid") and data.grid:
                    add_table_from_grid(doc, data.grid)
                elif getattr(data, "num_rows", 0) > 0:
                    add_table_from_cells(doc, item)
            except Exception as exc:
                log.warning("[стр%d] table: пропущена — %s", page_no, exc)
            continue

        # ── Картинки / логотипы ───────────────────────────────────────────────
        if lbl in ("picture", "figure", "image"):
            try:
                pil_img = item.get_image(dl_doc)
                if pil_img is None:
                    log.debug("[стр%d] %s: get_image вернул None — пропуск", page_no, lbl)
                    continue
                log.info("[стр%d] %s: %dx%d px", page_no, lbl, pil_img.width, pil_img.height)
                text_w_inch = (pw - 2 * MARGIN_INCH * 72) / 72
                text_h_inch = (ph - 2 * MARGIN_INCH * 72) / 72

                if bbox is not None:
                    bbox_w_pt     = float(getattr(bbox, "r", 0)) - float(getattr(bbox, "l", 0))
                    target_w_inch = min(bbox_w_pt / 72, text_w_inch) if bbox_w_pt > 10 else text_w_inch
                    img_align     = detect_alignment(bbox, pw)
                else:
                    target_w_inch = text_w_inch
                    img_align     = WD_ALIGN_PARAGRAPH.CENTER

                pil_w, pil_h_px = pil_img.size
                aspect = pil_h_px / pil_w if pil_w > 0 else 1.0
                if target_w_inch * aspect > text_h_inch:
                    target_w_inch = text_h_inch / aspect
                target_w_inch = max(target_w_inch, 0.5)

                pic_right   = float(getattr(bbox, "r", 0)) if bbox is not None else pw
                right_blocks: list[dict] = []
                right_skip:   list[int]  = []
                _stop = {"table", "picture", "figure", "image"}

                for j in range(idx + 1, min(idx + 15, len(all_items))):
                    nxt_item, _ = all_items[j]
                    nxt_lbl     = _label_str(nxt_item)
                    if not nxt_lbl or nxt_lbl in SKIP_LABELS:
                        right_skip.append(j)
                        continue
                    if nxt_lbl in _stop:
                        break
                    nxt_prov = getattr(nxt_item, "prov", None) or []
                    if not nxt_prov:
                        continue
                    if int(getattr(nxt_prov[0], "page_no", -1)) != page_no:
                        break
                    nxt_bbox = getattr(nxt_prov[0], "bbox", None)
                    nxt_x0   = bbox_x0(nxt_bbox) if nxt_bbox else 0.0
                    if not (coplanar(bbox, nxt_bbox, tolerance=250.0)
                            and nxt_x0 >= pic_right * 0.80):
                        break
                    nxt_text = postprocess(
                        (getattr(nxt_item, "text", None) or "").strip()
                    )
                    if _is_letterhead_stop(nxt_text):
                        break
                    right_skip.append(j)
                    if not nxt_text:
                        continue
                    nxt_h  = bbox_h(nxt_bbox) if nxt_bbox else 0.0
                    if nxt_h > 2 and median_h > 0 and nxt_h / median_h < 0.85:
                        nxt_pt = max(round(BODY_PT * (nxt_h / median_h) * 2) / 2, 7.0)
                    else:
                        nxt_pt = 8.5
                    _ha = [c for c in nxt_text if c.isalpha()]
                    right_blocks.append({
                        "text":      nxt_text,
                        "font_pt":   nxt_pt,
                        "bold":      bool(_ha) and all(c.isupper() for c in _ha) and len(nxt_text) <= 60,
                        "italic":    nxt_lbl in ("caption", "footnote"),
                        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
                    })

                _page_break(page_no)
                if right_blocks:
                    skip_indices.update(right_skip)
                    add_sidebyside(doc, pil_img, target_w_inch, right_blocks, text_w_inch)
                else:
                    buf = BytesIO()
                    pil_img.save(buf, format="PNG")
                    buf.seek(0)
                    para = doc.add_paragraph()
                    para.alignment                     = img_align
                    para.paragraph_format.space_before = Pt(4)
                    para.paragraph_format.space_after  = Pt(4)
                    para.add_run().add_picture(buf, width=Inches(target_w_inch))
            except Exception as exc:
                log.debug("Picture skipped: %s", exc)
            continue

        # ── Получаем текст из Docling (word_order меняет только порядок блоков) ──
        raw_text = (getattr(item, "text", None) or "").strip()
        text = postprocess(raw_text)
        if not text:
            continue

        # ── space_before, indent, alignment ──────────────────────────────────
        space_before = 0.0
        if bbox is not None and page_no in prev_midY and median_h > 0:
            gap   = abs(bbox_mid_y(bbox) - prev_midY[page_no]) \
                    - (item_h / 2 + prev_h.get(page_no, item_h) / 2)
            extra = gap - median_h * 1.2
            if extra > 2:
                space_before = min(extra * 0.5, 8.0)

        if bbox is not None:
            prev_midY[page_no] = bbox_mid_y(bbox)
            prev_h[page_no]    = item_h

        indent_pt = 0.0
        if bbox is not None:
            raw_indent = float(getattr(bbox, "l", 0)) - page_left_min.get(page_no, 0.0)
            if 0 < raw_indent < pw * 0.25:
                indent_pt = round(raw_indent, 1)

        alignment = (
            detect_alignment(bbox, pw)
            if bbox is not None and pw > 0
            else WD_ALIGN_PARAGRAPH.JUSTIFY
        )

        # Абзацы шире 50% страницы не центрируем: короткие обёрнутые строки
        # («Срок возврата кредита – не позднее 10.05.2028.») ложно детектируются
        # как CENTER, хотя должны быть выровнены по ширине.
        if lbl in ("paragraph", "text") and alignment == WD_ALIGN_PARAGRAPH.CENTER \
                and bbox is not None:
            bw = float(getattr(bbox, "r", pw)) - float(getattr(bbox, "l", 0))
            if bw / pw >= 0.50:
                alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        _alpha       = [c for c in text if c.isalpha()]
        _is_all_caps = bool(_alpha) and all(c.isupper() for c in _alpha) and len(text.strip()) <= 60
        bold   = lbl in ("title", "section_header") or _is_all_caps
        italic = lbl in ("caption", "footnote")

        log.info("[стр%d] lbl=%-16s align=%-8s bold=%s font=%.1f  %r",
                 page_no, lbl,
                 _align_names.get(alignment, str(alignment)),
                 "Y" if bold else "N",
                 font_pt,
                 text[:80])

        _page_break(page_no)

        # ── Блоки МЕТКА:содержимое ────────────────────────────────────────────
        _item_x0 = float(getattr(bbox, "l", 0)) if bbox is not None else 0.0
        lc = split_label_content(text) if _item_x0 <= pw * LABEL_MARGIN_THRESHOLD else None
        if lc is not None:
            log.info("  → label:content  label=%r  content=%r  x0=%.1fpt (thr=%.1fpt)",
                     lc[0], lc[1][:50], _item_x0, pw * LABEL_MARGIN_THRESHOLD)
            label_txt, first_content = lc
            text_w_inch = (pw - 2 * MARGIN_INCH * 72) / 72
            content_items: list[dict] = []
            if first_content:
                content_items.append({"text": first_content, "font_pt": font_pt, "bold": True, "italic": False})
            last_y = bbox_mid_y(bbox) if bbox is not None else 0.0
            for j in range(idx + 1, min(idx + 8, len(all_items))):
                if j in skip_indices:
                    continue
                nxt_item, _ = all_items[j]
                nxt_lbl     = _label_str(nxt_item)
                if nxt_lbl in (SKIP_LABELS | {"table", "picture", "figure", "image", "list_item"}):
                    break
                nxt_prov = getattr(nxt_item, "prov", None) or []
                if not nxt_prov or int(getattr(nxt_prov[0], "page_no", -1)) != page_no:
                    break
                nxt_bbox = getattr(nxt_prov[0], "bbox", None)
                if nxt_bbox and median_h > 0:
                    nxt_y = bbox_mid_y(nxt_bbox)
                    if abs(nxt_y - last_y) > median_h * 4:
                        break
                    last_y = nxt_y
                nxt_text = postprocess((getattr(nxt_item, "text", None) or "").strip())
                if not nxt_text:
                    skip_indices.add(j)
                    continue
                if split_label_content(nxt_text) is not None:
                    break
                nxt_h  = bbox_h(nxt_bbox) if nxt_bbox else 0.0
                nxt_pt = font_pt
                if nxt_h > 2 and median_h > 0 and nxt_h / median_h < 0.78:
                    nxt_pt = max(round(BODY_PT * (nxt_h / median_h) * 2) / 2, 7.0)
                content_items.append({
                    "text": nxt_text, "font_pt": nxt_pt,
                    "bold": (len(content_items) == 0), "italic": False,
                })
                skip_indices.add(j)
            if bbox is not None:
                prev_midY[page_no] = bbox_mid_y(bbox)
                prev_h[page_no]    = item_h
            add_label_content_table(doc, label_txt, content_items, text_w_inch, space_before)
            continue

        # ── Заголовки ─────────────────────────────────────────────────────────
        if lbl in LABEL_HEADING:
            para = doc.add_heading("", level=min(LABEL_HEADING[lbl], 9))
            para.alignment                       = alignment
            para.paragraph_format.space_before   = Pt(max(space_before, 4.0))
            para.paragraph_format.space_after    = Pt(2.0)
            para.paragraph_format.widow_control  = False
            para.paragraph_format.keep_with_next = False
            para.paragraph_format.keep_together  = False
            run                = para.add_run(text)
            run.font.name      = FONT_NAME
            run.font.size      = Pt(font_pt)
            run.font.bold      = True
            run.font.color.rgb = RGBColor(0, 0, 0)
            continue

        # ── Списки ───────────────────────────────────────────────────────────
        if lbl == "list_item":
            try:
                para = doc.add_paragraph(style="List Bullet")
            except KeyError:
                para = doc.add_paragraph()
            para.paragraph_format.space_before = Pt(space_before)
            para.paragraph_format.space_after  = Pt(1)
            para.paragraph_format.left_indent  = Pt(indent_pt)
            run           = para.add_run(text)
            run.font.name = FONT_NAME
            run.font.size = Pt(font_pt)
            continue

        # ── Обычные параграфы ─────────────────────────────────────────────────
        para = doc.add_paragraph()
        para.alignment                      = alignment
        para.paragraph_format.space_before  = Pt(space_before)
        para.paragraph_format.space_after   = Pt(0)
        para.paragraph_format.widow_control = False

        if lbl in ("paragraph", "text") and alignment == WD_ALIGN_PARAGRAPH.JUSTIFY and indent_pt < 8.0:
            para.paragraph_format.first_line_indent = Pt(35.4)
            para.paragraph_format.left_indent       = Pt(0)
        else:
            para.paragraph_format.first_line_indent = Pt(0)
            para.paragraph_format.left_indent       = Pt(indent_pt)

        run           = para.add_run(text)
        run.font.name = FONT_NAME
        run.font.size = Pt(font_pt)
        run.bold      = bold
        run.italic    = italic

    # ── Итоговая статистика ───────────────────────────────────────────────────
    total   = len(all_items)
    skipped = len(skip_indices)
    labels_seen: dict[str, int] = {}
    for item, _ in all_items:
        lbl = _label_str(item)
        if lbl:
            labels_seen[lbl] = labels_seen.get(lbl, 0) + 1
    log.info("build_docx итог: обработано %d из %d элементов (пропущено %d)",
             total - skipped, total, skipped)
    for lbl, cnt in sorted(labels_seen.items()):
        log.info("  %-20s %d", lbl, cnt)

    return doc


# ── Публичные функции ─────────────────────────────────────────────────────────

def convert_pdf(
    pdf_path: Path,
    docx_path: Path,
    converter,
    ocr_reader=None,
    use_word_order: bool = True,
) -> bool:
    log.info("  Конвертация: %s", pdf_path.name)
    try:
        result = converter.convert(str(pdf_path))
        dl_doc = result.document

        page_sizes: dict[int, tuple[float, float]] = {}
        for page_no, page in dl_doc.pages.items():
            size = getattr(page, "size", None)
            if size:
                page_sizes[int(page_no)] = (
                    float(getattr(size, "width",  0.0)),
                    float(getattr(size, "height", 0.0)),
                )

        doc = build_docx(dl_doc, page_sizes, ocr_reader=ocr_reader,
                         use_word_order=use_word_order)
        doc.save(str(docx_path))
        log.info("  ✓ %s", docx_path.name)
        return True
    except Exception as exc:
        log.error("  ✗ %s: %s", pdf_path.name, exc, exc_info=True)
        return False


class DoclingBatchConverter:
    def __init__(
        self,
        input_folder: str,
        output_folder: str,
        backup_folder: str | None = None,
        langs: list[str] | None = None,
        use_word_order: bool = True,
    ) -> None:
        self.input_folder   = Path(input_folder)
        self.output_folder  = Path(output_folder)
        self.backup_folder  = Path(backup_folder) if backup_folder else None
        self.langs          = langs or ["ru", "en"]
        self.use_word_order = use_word_order
        self.output_folder.mkdir(parents=True, exist_ok=True)
        if self.backup_folder:
            self.backup_folder.mkdir(parents=True, exist_ok=True)
        self._converter  = None
        self._ocr_reader = None

    @property
    def converter(self):
        if self._converter is None:
            log.info("Инициализация Docling pipeline (DocLayNet + TableFormer + EasyOCR)...")
            self._converter = build_converter(self.langs)
            log.info("Pipeline готов.")
        return self._converter

    @property
    def ocr_reader(self):
        if self._ocr_reader is None and self.use_word_order:
            try:
                import easyocr
                log.info("Инициализация EasyOCR reader для word_order...")
                self._ocr_reader = easyocr.Reader(self.langs, gpu=False, verbose=False)
                log.info("EasyOCR готов.")
            except Exception as exc:
                log.warning("EasyOCR недоступен, word_order отключён: %s", exc)
        return self._ocr_reader

    def process(self, move_to_backup: bool = True) -> None:
        pdf_files = sorted(self.input_folder.glob("*.pdf"))
        total     = len(pdf_files)
        log.info("Найдено %d PDF-файлов.", total)
        stats = {"ok": 0, "fail": 0, "skip": 0}
        t0    = time.monotonic()

        for i, pdf_path in enumerate(pdf_files, 1):
            docx_path = self.output_folder / f"{pdf_path.stem}.docx"
            if docx_path.exists():
                log.info("[%d/%d] Пропуск: %s", i, total, pdf_path.name)
                stats["skip"] += 1
                continue
            log.info("[%d/%d] %s (%.2f MB)", i, total, pdf_path.name,
                     pdf_path.stat().st_size / 1_048_576)
            ok = convert_pdf(pdf_path, docx_path, self.converter,
                             ocr_reader=self.ocr_reader,
                             use_word_order=self.use_word_order)
            if ok:
                stats["ok"] += 1
                if self.backup_folder and move_to_backup:
                    shutil.move(str(pdf_path), str(self.backup_folder / pdf_path.name))
            else:
                stats["fail"] += 1
            if i % 10 == 0:
                gc.collect()

        elapsed = time.monotonic() - t0
        sep = "=" * 55
        log.info("\n%s", sep)
        log.info("Готово за %.1f мин | OK: %d | Ошибок: %d | Пропущено: %d",
                 elapsed / 60, stats["ok"], stats["fail"], stats["skip"])
        log.info(sep)
