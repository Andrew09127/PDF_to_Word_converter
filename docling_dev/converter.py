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


class _TextPrefixItem:
    """Обёртка элемента Docling: добавляет числовой префикс к тексту при рендере.

    Используется в _fix_reading_order когда ненумерованный item вставляется
    в пробел нумерованного списка (OCR пропустил число).
    """
    __slots__ = ("_item", "_prefix")

    def __init__(self, item, prefix: str) -> None:
        self._item   = item
        self._prefix = prefix

    @property
    def text(self) -> str:
        return self._prefix + (getattr(self._item, "text", None) or "")

    def __getattr__(self, name: str):
        if name in ("_item", "_prefix"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextAppendItem:
    """Добавляет суффикс к тексту элемента (для патроним-фикса)."""
    __slots__ = ("_item", "_suffix")

    def __init__(self, item, suffix: str) -> None:
        self._item   = item
        self._suffix = suffix

    @property
    def text(self) -> str:
        return (getattr(self._item, "text", None) or "") + self._suffix

    def __getattr__(self, name: str):
        if name in ("_item", "_suffix"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextSliceItem:
    """Возвращает текст элемента начиная с указанной позиции (для патроним-фикса)."""
    __slots__ = ("_item", "_start")

    def __init__(self, item, start: int) -> None:
        self._item  = item
        self._start = start

    @property
    def text(self) -> str:
        return (getattr(self._item, "text", None) or "")[self._start:]

    def __getattr__(self, name: str):
        if name in ("_item", "_start"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextInsertItem:
    """Вставляет строку в позицию pos, опционально усекает суффикс (truncate_at >= 0)."""
    __slots__ = ("_item", "_pos", "_insertion", "_truncate_at")

    def __init__(self, item, pos: int, insertion: str, truncate_at: int = -1) -> None:
        self._item        = item
        self._pos         = pos
        self._insertion   = insertion
        self._truncate_at = truncate_at

    @property
    def text(self) -> str:
        t = getattr(self._item, "text", None) or ""
        if 0 <= self._truncate_at < len(t):
            t = t[:self._truncate_at]
        return t[:self._pos] + self._insertion + t[self._pos:]

    def __getattr__(self, name: str):
        if name in ("_item", "_pos", "_insertion", "_truncate_at"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextSliceAndInsertItem:
    """Slice от start + вставка insertion после insert_pos (в sliced-координатах)."""
    __slots__ = ("_item", "_start", "_insert_pos", "_insertion")

    def __init__(self, item, start: int, insert_pos: int, insertion: str) -> None:
        self._item       = item
        self._start      = start
        self._insert_pos = insert_pos
        self._insertion  = insertion

    @property
    def text(self) -> str:
        sliced = (getattr(self._item, "text", None) or "")[self._start:]
        p = self._insert_pos
        return sliced[:p] + self._insertion + sliced[p:]

    def __getattr__(self, name: str):
        if name in ("_item", "_start", "_insert_pos", "_insertion"):
            raise AttributeError(name)
        return getattr(self._item, name)


# Паттерн: блок начинается с отчества в тв. падеже (Игоревичем, Петровичем).
# Используем \S+? вместо [А-ЯЁ][а-яё]*? — EasyOCR иногда подмешивает Latin-символы
# (Latin 'e' вместо Cyrillic 'е'), поэтому кириллический char-class ненадёжен.
# [еe] в суффиксе покрывает оба варианта написания «е».
_PATRON_CONT_RE = re.compile(
    r'^(\S+?(?:ич[еe]м|вич[еe]м)(?:\s*\([^)]{2,30}\))?)\s+(.+)',
    re.DOTALL,
)

# Паттерн для поиска позиции вставки патронима в блоке A:
# имя в тв. падеже (оканчивается на «ем»/«еем» с возможными Latin-символами)
# непосредственно перед глаголом «был»/«была»/«были»/«заключен».
# Захват: group(1) = слово с именем, позиция вставки = m.end(1).
_FN_BEFORE_VERB_RE = re.compile(
    r'(\S+?[еe]{1,2}м)\s+(?=(?:был[аи]?\b|заключен\b))',
    re.IGNORECASE,
)

# Паттерн: блок A заканчивается "лимитом выдачи в размере" — OCR разорвал фразу.
# Используем \S*keyword\S* для устойчивости к OCR-артефактам (Latin B вместо Cyrillic в и т.п.).
_LIMIT_SUFFIX_RE  = re.compile(
    r'\s+\S*лимит\S*\s+\S*выдач\S*\s+\S+\s+\S*размер\S*\s*$',
    re.IGNORECASE,
)
_CREDIT_LINE_RE   = re.compile(r'\S*кредит\S*\s+\S*лини\S+', re.IGNORECASE)


def _fix_reading_order(all_items: list) -> tuple[list, set]:
    """
    Пост-обработка порядка элементов после word_order.

    Возвращает (result, continuation_ids):
      result          — переупорядоченный список элементов
      continuation_ids — set id() элементов, которые нужно безусловно слить
                         с предыдущим параграфом при рендере (Fix 2 date-start)
    """
    result           = list(all_items)
    n                = len(result)
    continuation_ids: set = set()

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

    # Fix 4: section_header «по делу №» стоит ПОСЛЕ label:content (КРЕДИТОР:/ДОЛЖНИК:)
    # В судебных заявлениях «по делу» всегда предшествует блокам сторон.
    # Docling иногда путает порядок из-за близости по Y-координате.
    _DELO_RE = re.compile(r'^\s*по\s+делу\b', re.IGNORECASE)
    for i in range(n - 1):
        if _lbl(i) != "text" or _lbl(i + 1) not in ("section_header", "text"):
            continue
        if _pg(i) != _pg(i + 1):
            continue
        if not _DELO_RE.match(_txt(i + 1)):
            continue
        if split_label_content(_txt(i)) is None:
            continue
        result[i], result[i + 1] = result[i + 1], result[i]
        log.info("fix_order: по_делу swap %r ↔ %r", _txt(i)[:40], _txt(i + 1)[:40])

    # Fix 2: текстовый блок с датой DD.MM.YYYY стоит ПОСЛЕ незаконченного блока
    # Docling иногда разбивает один абзац на части и переставляет их:
    #   "Игоревичем (Заемщик) Заемщику..." → "20.12.2019 между..."
    # Корректный порядок: дата → продолжение.
    # Признак неправильного порядка: предыдущий ТЕКСТОВЫЙ блок не заканчивается на «.»/«!»/«?»
    # Важно: между двумя текстовыми блоками могут быть section_header / picture —
    # ищем не строго i-1, а ближайший предшествующий текстовый блок на той же странице.
    _TEXT_LABELS = frozenset({"text", "paragraph"})
    for i in range(1, n):
        if _lbl(i) not in _TEXT_LABELS:
            continue
        curr = _txt(i)
        if not _DATE_START_RE.match(curr):
            continue
        # Ищем ближайший предшествующий текстовый блок на той же странице
        prev_i = None
        for k in range(i - 1, max(i - 6, -1), -1):
            if _lbl(k) in _TEXT_LABELS and _pg(k) == _pg(i):
                prev_i = k
                break
            if _pg(k) != _pg(i):
                break
        if prev_i is None:
            log.debug("fix_order: date-start не нашёл текстового предшественника для %r", curr[:40])
            continue
        prev = _txt(prev_i)

        # «17.09.1987, место рождения:» — дата рождения, за датой стоит запятая.
        # Это НЕ начало предложения о договоре → не переставляем.
        date_match_end = _DATE_START_RE.match(curr).end()
        rest = curr[date_match_end:].lstrip()
        if rest.startswith(','):
            log.debug("fix_order: date-start пропуск (дата рождения) %r", curr[:50])
            continue

        # Предыдущий блок завершает ПОЛНОЕ предложение → не переставляем.
        # «:» НЕ включаем: OCR часто ставит «:» вместо «.», а «средств:» перед
        # датой договора — это как раз наш случай (нужно переставить).
        if prev.endswith(('.', '!', '?', '»', ';')):
            log.debug("fix_order: date-start пропуск (предш. заканчивается на терминатор) %r", prev[-20:])
            continue

        log.debug("fix_order: date-start кандидат: prev=%r curr=%r (prev_i=%d, i=%d)",
                  prev[:50], curr[:50], prev_i, i)

        # Переставляем: все элементы между prev_i и i сдвигаются на 1 вперёд
        saved = result[i]
        for k in range(i, prev_i, -1):
            result[k] = result[k - 1]
        result[prev_i] = saved
        log.info("fix_order: date-start swap %r ↔ %r", prev[:50], curr[:50])
        # Элемент на позиции prev_i+1 — прямое продолжение вставленного блока
        # (это и есть тот «Игоревичем...», что шёл ДО даты в Docling).
        # Помечаем его для безусловного слияния в renderer.
        if prev_i + 1 < n and _lbl(prev_i + 1) in _TEXT_LABELS:
            b_raw = _txt(prev_i + 1)   # raw text из Docling (до OCR-фиксов)
            pm = _PATRON_CONT_RE.match(b_raw)
            if pm:
                # Блок B начинается с отчества в тв. падеже (Игоревичем (Заемщик)).
                # Переносим его в конец блока A — восстанавливаем полное ФИО.
                patron     = pm.group(1)   # "Игоревичем (Заемщик)"
                b_start    = pm.start(2)   # позиция в b_raw (stripped)
                a_orig     = result[prev_i]
                b_orig     = result[prev_i + 1]
                # Корректируем b_start до позиции в сыром тексте оригинального элемента
                # (b_raw = strip(original.text), поэтому start может сдвинуться).
                orig_text  = getattr(b_orig[0], "text", None) or ""
                g2_prefix  = pm.group(2)[:20]   # первые 20 символов группы 2
                real_start = orig_text.find(g2_prefix)
                if real_start >= 0:
                    b_start = real_start
                a_text    = getattr(a_orig[0], "text", None) or ""

                # Шаг 1: ищем имя перед «был» → вставляем патроним после имени.
                fn_match   = _FN_BEFORE_VERB_RE.search(a_text)
                # Шаг 2: ищем «лимитом выдачи в размере» в конце блока A →
                # переносим в блок B после «кредитной линии».
                lim_match  = _LIMIT_SUFFIX_RE.search(a_text)
                lim_text   = a_text[lim_match.start():].strip() if lim_match else None
                truncate_a = lim_match.start() if lim_match else -1

                if fn_match:
                    ins_pos = fn_match.end(1)
                    result[prev_i] = (
                        _TextInsertItem(a_orig[0], ins_pos, " " + patron, truncate_at=truncate_a),
                        a_orig[1],
                    )
                    log.info("fix_order: patron-insert после %r (pos=%d) truncate=%d в %r",
                             fn_match.group(1), ins_pos, truncate_a, a_text[:60])
                else:
                    result[prev_i] = (_TextAppendItem(a_orig[0], " " + patron), a_orig[1])
                    log.info("fix_order: patron-append (имя не найдено) в %r", a_text[:60])

                # Шаг 3: если нашли «лимитом» → вставляем его в блок B после «кредитной линии»
                b_sliced_text = orig_text[b_start:]
                cl_match = _CREDIT_LINE_RE.search(b_sliced_text) if lim_text else None
                if lim_text and cl_match:
                    result[prev_i + 1] = (
                        _TextSliceAndInsertItem(
                            b_orig[0], b_start,
                            cl_match.end(),
                            " с " + lim_text,
                        ),
                        b_orig[1],
                    )
                    log.info("fix_order: лимит-move %r → в блок B после 'кредитной линии'", lim_text)
                else:
                    result[prev_i + 1] = (_TextSliceItem(b_orig[0], b_start), b_orig[1])
                log.info("fix_order: B теперь %r", b_sliced_text[:50])
            else:
                log.debug("fix_order: mark continuation %r", b_raw[:40])
            continuation_ids.add(id(result[prev_i + 1][0]))

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
        final_order:  list[int]      = []
        gap_prefixes: dict[int, str] = {}   # gap_k → "N " для заполнителей пробелов
        unnum_used  = 0
        prev_num    = 0
        for num, k in numbered_sorted:
            gap_size = num - prev_num - 1
            for _ in range(gap_size):
                if unnum_used < len(unnumbered):
                    gap_k   = unnumbered[unnum_used]
                    gap_num = prev_num + 1
                    final_order.append(gap_k)
                    gap_prefixes[gap_k] = f"{gap_num} "
                    log.info("fix_order: пункт %d пропущен OCR, заполняем %r",
                             gap_num, txts[gap_k][:60])
                    unnum_used += 1
            final_order.append(k)
            prev_num = num
        # Оставшиеся ненумерованные — в конец
        final_order.extend(unnumbered[unnum_used:])

        reordered = [
            (_TextPrefixItem(result[i + k][0], gap_prefixes[k]), result[i + k][1])
            if k in gap_prefixes else result[i + k]
            for k in final_order
        ]
        for k, item_lvl in enumerate(reordered):
            result[i + k] = item_lvl

        if not already_sorted:
            log.info("fix_order: sorted %d numbered list items on page %d",
                     len(numbered), pg)
        if has_gap and unnum_used > 0:
            log.info("fix_order: filled %d gap(s) with unnumbered items on page %d",
                     unnum_used, pg)
        i = j

    return result, continuation_ids


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
    skip_set = {pic_idx, *text_indices}
    log.info("letterhead: skip_indices добавлены — pic_idx=%d, text_idx=%s",
             pic_idx, sorted(text_indices))
    return skip_set


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
    all_items, _continuation_ids = _fix_reading_order(all_items)

    doc               = init_document()
    last_content_page = -1
    current_page      = -1
    prev_midY: dict[int, float] = {}
    prev_h:    dict[int, float] = {}

    # ── Трекинг продолжений (orphan-блоки Docling) ────────────────────────────
    # Docling иногда разрывает абзац: конец одного блока + начало следующего
    # теряют связь. Признак продолжения: блок начинается с союза/частицы
    # строчными буквами И предыдущий text-абзац не завершён.
    # Список консервативный — только слова, которые НЕ начинают новых предложений.
    _CONT_RE = re.compile(
        r'^(как\b|а\s+также\b|в\s+том\s+числе\b|при\s+этом\b|при\s+условии\b'
        r'|которы[еийх]\b|котор\w{2,5}\b'
        r'|обеспеченн\w+|предусмотренн\w+)',
        re.IGNORECASE,
    )
    # Только настоящие концы предложений: . ! ?
    # «»; НЕ включаем: в русских юридических текстах «» встречается внутри предложений
    # (названия законов, цитаты), а «;» — внутри перечислений.
    _SENT_END_RE = re.compile(r'[.!?]\s*$')

    _last_body_para:  object = None  # последний Word-параграф тела документа
    _last_body_text:  str    = ""    # его текст (для проверки завершённости)
    _seen_text_pages: set    = set() # страницы, на которых уже был text-блок
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

    # ── Детальный дамп стр.1 для диагностики рендера ─────────────────────────
    _pg1 = min(page_sizes.keys()) if page_sizes else 1
    _pw1, _ph1 = page_sizes.get(_pg1, (595.0, 842.0))
    log.info("=== ДАМП СТРАНИЦЫ %d  pw=%.0f ph=%.0f pdf_native=%s ===",
             _pg1, _pw1, _ph1, pdf_native)
    for _di, (_ditem, _dlvl) in enumerate(all_items):
        _dprov = getattr(_ditem, "prov", None) or []
        if not _dprov:
            continue
        _dpg = int(getattr(_dprov[0], "page_no", -1))
        if _dpg != _pg1:
            continue
        _dbbox  = getattr(_dprov[0], "bbox", None)
        _dlbl   = _label_str(_ditem)
        _draw   = (getattr(_ditem, "text", None) or "")
        _dflags: list[str] = []
        if _di in skip_indices:
            _dflags.append("SKIP")
        if id(_ditem) in _continuation_ids:
            _dflags.append("CONT")
        if _dbbox is not None:
            _dl = float(getattr(_dbbox, "l", 0));  _dt = float(getattr(_dbbox, "t", 0))
            _dr = float(getattr(_dbbox, "r", 0));  _db = float(getattr(_dbbox, "b", 0))
            _dh = abs(_db - _dt)
            # Нормированная вертикальная позиция (0=верх, 1=низ) в обеих системах
            _top_screen  = min(_dt, _db) / _ph1
            _top_native  = 1.0 - max(_dt, _db) / _ph1
            log.info(
                "  [%2d] %-16s lvl=%d  l=%5.1f t=%5.1f r=%5.1f b=%5.1f  h=%5.1f"
                "  top_sc=%.3f top_nat=%.3f  %s | %r",
                _di, _dlbl, _dlvl,
                _dl, _dt, _dr, _db, _dh,
                _top_screen, _top_native,
                ",".join(_dflags) if _dflags else "-",
                _draw[:80],
            )
        else:
            log.info("  [%2d] %-16s lvl=%d  (no bbox)  %s | %r",
                     _di, _dlbl, _dlvl,
                     ",".join(_dflags) if _dflags else "-", _draw[:80])
    log.info("=== КОНЕЦ ДАМПА СТРАНИЦЫ %d ===", _pg1)

    _align_names = {
        WD_ALIGN_PARAGRAPH.JUSTIFY: "JUSTIFY",
        WD_ALIGN_PARAGRAPH.CENTER:  "CENTER",
        WD_ALIGN_PARAGRAPH.LEFT:    "LEFT",
        WD_ALIGN_PARAGRAPH.RIGHT:   "RIGHT",
    }

    for idx, (item, _level) in enumerate(all_items):
        if idx in skip_indices:
            raw = (getattr(item, "text", None) or "").strip()[:50]
            log.debug("[%d] пропуск (skip_indices): lbl=%s  %r",
                      idx, _label_str(item), raw)
            continue

        lbl = _label_str(item)
        if not lbl or lbl in SKIP_LABELS:
            log.debug("[%d] пропуск (SKIP_LABELS): lbl=%r", idx, lbl)
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
            # CENTER-блоки (Госпошлина, колонтитулы) — шрифт не уменьшаем:
            # они либо намеренно мелкие (9pt), либо ложно сжаты по bbox.
            # Примечание: alignment вычисляется позже — определяем здесь отдельно.
            _pre_align = (detect_alignment(bbox, pw)
                          if bbox is not None and pw > 0
                          else WD_ALIGN_PARAGRAPH.JUSTIFY)
            if ratio < 0.55 and not is_potential_label \
                    and _pre_align != WD_ALIGN_PARAGRAPH.CENTER:
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
                        nxt_pt = max(round(BODY_PT * (nxt_h / median_h) * 2) / 2, 9.0)
                    else:
                        nxt_pt = 9.0
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
            log.debug("[стр%d] idx=%d lbl=%s: пустой текст — пропуск", page_no, idx, lbl)
            continue

        # Логируем значимые исправления OCR (изменение ≥ 3 символов или 5% текста)
        if raw_text != text:
            n_changed = sum(a != b for a, b in zip(raw_text, text)) + abs(len(raw_text) - len(text))
            if n_changed >= 3 or n_changed / max(len(raw_text), 1) >= 0.05:
                log.info("  ocr_fix [стр%d]: %r → %r", page_no, raw_text[:70], text[:70])

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

        # Корректировка ложного выравнивания для body text (paragraph/text):
        # – CENTER для блоков шире 50% страницы → JUSTIFY (короткие строки с переносом)
        # – RIGHT для блоков, начинающихся в левых 30% страницы → JUSTIFY
        #   (Docling даёт RIGHT когда bbox сдвинут вправо, но блок — обычный абзац)
        if lbl in ("paragraph", "text") and bbox is not None:
            bw    = float(getattr(bbox, "r", pw)) - float(getattr(bbox, "l", 0))
            raw_x0 = float(getattr(bbox, "l", 0))
            if alignment == WD_ALIGN_PARAGRAPH.CENTER and bw / pw >= 0.50:
                alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif alignment == WD_ALIGN_PARAGRAPH.RIGHT and raw_x0 < pw * 0.30:
                alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        _alpha       = [c for c in text if c.isalpha()]
        _is_all_caps = bool(_alpha) and all(c.isupper() for c in _alpha) and len(text.strip()) <= 60
        bold   = lbl in ("title", "section_header") or _is_all_caps
        italic = lbl in ("caption", "footnote")

        # Предупреждение при нетипичном размере шрифта
        if font_pt < 8.0:
            log.warning("[стр%d] lbl=%s: font=%.1fpt СЛИШКОМ МАЛО — возможна ошибка bbox",
                        page_no, lbl, font_pt)
        elif font_pt > 20.0 and lbl not in ("title", "section_header"):
            log.warning("[стр%d] lbl=%s: font=%.1fpt СЛИШКОМ ВЕЛИК для тела — возможна ошибка",
                        page_no, lbl, font_pt)

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
            # Вычисляем отступ для правоколоночных блоков (КРЕДИТОР:, ДОЛЖНИК: и т.п.)
            # Они располагаются в правой части страницы (x0 > 30% ширины) и должны
            # начинаться там же в Word, а не от левого поля.
            _full_tw     = (pw - 2 * MARGIN_INCH * 72) / 72
            _lc_indent   = 0.0
            _lc_col_ratio = None
            if alignment == WD_ALIGN_PARAGRAPH.RIGHT and _item_x0 > pw * 0.30:
                _lc_indent    = max((_item_x0 - MARGIN_INCH * 72) / 72, 0.0)
                _lc_col_ratio = 0.22   # компактная колонка для коротких меток
            text_w_inch = _full_tw - _lc_indent

            log.info("  → label:content  label=%r  content=%r  x0=%.1fpt (thr=%.1fpt) "
                     "indent=%.2fin col_ratio=%s",
                     lc[0], lc[1][:50], _item_x0, pw * LABEL_MARGIN_THRESHOLD,
                     _lc_indent, f"{_lc_col_ratio:.2f}" if _lc_col_ratio else "default")

            label_txt, first_content = lc
            content_items: list[dict] = []
            if first_content:
                content_items.append({"text": first_content, "font_pt": BODY_PT, "bold": True, "italic": False})
            last_y = bbox_mid_y(bbox) if bbox is not None else 0.0
            for j in range(idx + 1, min(idx + 8, len(all_items))):
                if j in skip_indices:
                    continue
                nxt_item, _ = all_items[j]
                nxt_lbl     = _label_str(nxt_item)
                # section_header и title — самостоятельные блоки, не сливаем в label:content
                if nxt_lbl in (SKIP_LABELS | {"table", "picture", "figure", "image",
                                               "list_item", "section_header", "title"}):
                    log.debug("  label:content остановлен на lbl=%s %r", nxt_lbl,
                              (getattr(nxt_item, "text", "") or "")[:40])
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
                # Стоп-слова: самостоятельные секции документа (не продолжение блока)
                if re.match(r'^(Госпошлин|ПРОСИТ\s+СУД|Приложени)', nxt_text, re.IGNORECASE):
                    log.debug("  label:content стоп (standalone) %r", nxt_text[:30])
                    break
                nxt_h  = bbox_h(nxt_bbox) if nxt_bbox else 0.0
                # Базируемся на BODY_PT (11pt), а не font_pt метки (может быть 13pt
                # для section_header → тогда персданные ДОЛЖНИКА рендерятся 13pt).
                nxt_pt = BODY_PT
                if nxt_h > 2 and median_h > 0 and nxt_h / median_h < 0.78:
                    nxt_pt = max(round(BODY_PT * (nxt_h / median_h) * 2) / 2, 9.0)
                content_items.append({
                    "text": nxt_text, "font_pt": nxt_pt,
                    "bold": (len(content_items) == 0), "italic": False,
                })
                log.info("    content[%d] font=%.1fpt bold=%s %r",
                         len(content_items) - 1, nxt_pt,
                         "Y" if len(content_items) == 1 else "N", nxt_text[:60])
                skip_indices.add(j)
            if bbox is not None:
                prev_midY[page_no] = bbox_mid_y(bbox)
                prev_h[page_no]    = item_h
            add_label_content_table(doc, label_txt, content_items, text_w_inch, space_before,
                                    indent_inch=_lc_indent, col_ratio=_lc_col_ratio)
            continue

        # ── Заголовки ─────────────────────────────────────────────────────────
        if lbl in LABEL_HEADING:
            # Подзаголовки (начинаются со строчной буквы) — центрируем.
            # «о включении в реестр» / «требований кредиторов должника» стоят
            # под «ЗАЯВЛЕНИЕ» и должны быть по центру, как в оригинале.
            alpha = [c for c in text if c.isalpha()]
            if alpha and alpha[0].islower():
                alignment = WD_ALIGN_PARAGRAPH.CENTER
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

        # ── Списки ────────────────────────────────────────────────────────────
        # Нумерованные пункты («N текст»)   → стиль Word «List Number»
        # Подпункты (первая буква строчная) → стиль Word «List Bullet» + отступ
        # Остальные list_item               → абзац с красной строкой (ГОСТ 7.32)
        if lbl == "list_item":
            if alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            _li_alpha  = [c for c in text if c.isalpha()]
            num_match  = _NUM_ITEM_RE.match(text)
            is_sub_item = bool(_li_alpha) and _li_alpha[0].islower() and not num_match

            if num_match:
                text_body = text[num_match.end():]
                para = doc.add_paragraph(style="List Number")
                para.alignment                      = alignment
                para.paragraph_format.space_before  = Pt(space_before)
                para.paragraph_format.space_after   = Pt(0)
                para.paragraph_format.widow_control = False
                run           = para.add_run(text_body)
                run.font.name = FONT_NAME
                run.font.size = Pt(font_pt)
                log.info("[стр%d] list_item → List Number  font=%.1fpt align=%-8s %r",
                         page_no, font_pt,
                         _align_names.get(alignment, str(alignment)), text[:60])
            elif is_sub_item:
                # Подпункт (строчная буква): маркированный список с отступом
                para = doc.add_paragraph(style="List Bullet")
                para.alignment                      = WD_ALIGN_PARAGRAPH.JUSTIFY
                para.paragraph_format.space_before  = Pt(space_before)
                para.paragraph_format.space_after   = Pt(0)
                para.paragraph_format.widow_control = False
                run           = para.add_run(text)
                run.font.name = FONT_NAME
                run.font.size = Pt(font_pt)
                log.info("[стр%d] list_item → List Bullet   font=%.1fpt %r",
                         page_no, font_pt, text[:60])
            else:
                para = doc.add_paragraph()
                para.alignment                      = alignment
                para.paragraph_format.space_before  = Pt(space_before)
                para.paragraph_format.space_after   = Pt(0)
                para.paragraph_format.widow_control = False
                # Красная строка: 35.4 pt ≈ 1.25 cm (ГОСТ 7.32)
                para.paragraph_format.first_line_indent = Pt(35.4)
                para.paragraph_format.left_indent       = Pt(0)
                run           = para.add_run(text)
                run.font.name = FONT_NAME
                run.font.size = Pt(font_pt)
                log.info("[стр%d] list_item → красная строка font=%.1fpt align=%-8s %r",
                         page_no, font_pt,
                         _align_names.get(alignment, str(alignment)), text[:60])
            # Обновляем трекинг для последующей проверки continuation
            _last_body_para = para
            _last_body_text = text
            _seen_text_pages.add(page_no)
            continue

        # ── Обычные параграфы ─────────────────────────────────────────────────

        # Детектирование продолжения (orphan-блок): блок начинается с союза/
        # частицы строчными буквами, а предыдущий параграф не завершён.
        # Две ситуации:
        #   1) Межстраничное: первый text-блок страницы с маленькой буквы
        #   2) Внутри страницы: блок начинается с явного союза
        _text_alpha = [c for c in text if c.isalpha()]
        _is_lowercase_start = bool(_text_alpha) and _text_alpha[0].islower()
        _prev_unfinished = (
            _last_body_para is not None and
            not _SENT_END_RE.search(_last_body_text)
        )

        _is_crosspage = (
            _is_lowercase_start and
            page_no not in _seen_text_pages and   # первый text-блок этой страницы
            _prev_unfinished
        )
        _is_conjunction = (
            _is_lowercase_start and
            bool(_CONT_RE.match(text)) and
            _prev_unfinished
        )
        # Третий случай: JUSTIFY-блок с маленькой буквы после блока, заканчивающегося
        # на двоеточие «:» — явный признак начатого перечисления/продолжения.
        # «возврата кредита не позднее...» после «...оборотных средств:».
        # Ограничения делают правило безопасным:
        #   • предыдущий блок ДОЛЖЕН заканчиваться на ':'
        #   • текущий блок НЕ является датой DD.MM.YYYY (дата всегда начинает новый абзац)
        _is_justify_cont = (
            _is_lowercase_start and
            not bool(_CONT_RE.match(text)) and
            not bool(_DATE_START_RE.match(text)) and
            lbl in ("text", "paragraph") and
            alignment == WD_ALIGN_PARAGRAPH.JUSTIFY and
            _prev_unfinished and
            _last_body_text.rstrip().endswith(':')
        )

        # Четвёртый случай: блок помечен в Fix 2 (date-start swap) как прямое
        # продолжение вставленного датой блока — сливаем независимо от регистра.
        # «Игоревичем (Заемщик) Заемщику предоставлялся кредит...» после «20.12.2019 между...»
        _is_marked_cont = (
            id(item) in _continuation_ids and
            lbl in ("text", "paragraph") and
            _prev_unfinished
        )

        if lbl in ("text", "paragraph") and (_is_crosspage or _is_conjunction
                                              or _is_justify_cont or _is_marked_cont):
            # Присоединяем к предыдущему параграфу вместо создания нового
            kind = ("cross-page"   if _is_crosspage   else
                    "conjunction"  if _is_conjunction  else
                    "marked-cont"  if _is_marked_cont  else "justify-cont")
            log.info("[стр%d] continuation (%s): %r → merged into previous",
                     page_no, kind, text[:60])
            run           = _last_body_para.add_run(" " + text)
            run.font.name = FONT_NAME
            run.font.size = Pt(font_pt)
            _last_body_text += " " + text
            if page_no not in _seen_text_pages:
                _seen_text_pages.add(page_no)
        else:
            para = doc.add_paragraph()
            para.alignment                      = alignment
            para.paragraph_format.space_before  = Pt(space_before)
            para.paragraph_format.space_after   = Pt(0)
            para.paragraph_format.widow_control = False

            # Красная строка: paragraph/text с JUSTIFY-выравниванием.
            # indent_pt < 8  — блок у левого поля → красная строка.
            # 20 < indent_pt < 50 — OCR bbox захватил первую строку с отступом
            #   (Docling даёт l-координату ПЕРВОГО символа = начало красной строки).
            #   В этом случае bbox.l ≈ 70pt + 35pt = 105pt → indent_pt ≈ 35pt.
            _is_red_line = (
                lbl in ("paragraph", "text")
                and alignment == WD_ALIGN_PARAGRAPH.JUSTIFY
                and (indent_pt < 8.0 or 20.0 < indent_pt < 50.0)
            )
            if _is_red_line:
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

            if lbl in ("text", "paragraph"):
                _last_body_para  = para
                _last_body_text  = text
                _seen_text_pages.add(page_no)

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
