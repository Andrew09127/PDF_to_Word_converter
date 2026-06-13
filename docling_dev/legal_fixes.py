"""
legal_fixes.py
──────────────
Исправление порядка блоков специфичное для судебных/арбитражных документов.

Активируется только когда doc_type.detect_doc_type() == DOC_LEGAL.
Содержит бывшую функцию _fix_reading_order из converter.py без изменений.

Паттерны:
  Fix 1: ALL-CAPS section_header перед строчным subtitle
  Fix 2: блок с датой DD.MM.YYYY стоит ПОСЛЕ незаконченного блока (date-start swap)
  Fix 3: сортировка numbered list_items + заполнение пробелов
  Fix 4: «по делу №» стоит ПОСЛЕ label:content (КРЕДИТОР:/ДОЛЖНИК:)
  Fix 5: list_item стоит ДО heading «ПРОСИТ СУД»
  Fix 6: разбивка блоков где OCR объединил несколько абзацев
"""
from __future__ import annotations

import logging
import re

from .docx_builder import split_label_content

log = logging.getLogger(__name__)

# ── Регулярные выражения ──────────────────────────────────────────────────────

_NUM_ITEM_RE   = re.compile(r'^(\d+)\s')
_DATE_START_RE = re.compile(r'^\d{2}\.\d{2}\.\d{4}')

_PATRON_CONT_RE = re.compile(
    r'^(\S+?(?:ич[еe]м|вич[еe]м)(?:\s*\([^)]{2,30}\))?)\s+(.+)',
    re.DOTALL,
)
_FN_BEFORE_VERB_RE = re.compile(
    r'(\S+?[еe]{1,2}м)\s+(?=(?:был[аи]?\b|заключен\b))',
    re.IGNORECASE,
)
_LIMIT_SUFFIX_RE = re.compile(
    r'\s+\S*лимит\S*\s+\S*выдач\S*\s+\S+\s+\S*размер\S*\s*$',
    re.IGNORECASE,
)
_CREDIT_LINE_RE = re.compile(r'\S*кредит\S*\s+\S*лини\S+', re.IGNORECASE)
_DELO_RE        = re.compile(r'^\s*по\s+делу\b', re.IGNORECASE)
_PROSIT_RE      = re.compile(r'\bПРОСИТ\b', re.IGNORECASE)

_TEXT_LABELS  = frozenset({"text", "paragraph"})
_PROSIT_LBLS  = frozenset({"list_item", "text", "paragraph"})

_PARA_SPLIT_RES = [
    re.compile(r'[;:,.]\s+((?:Срок\s+)?возврата\s+кредита)', re.IGNORECASE),
    re.compile(r'\.\s+(За\s+пользование)', re.IGNORECASE),
]


# ── Вспомогательные обёртки элементов ────────────────────────────────────────

class _TextPrefixItem:
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


class _TextTruncateItem:
    __slots__ = ("_item", "_end")

    def __init__(self, item, end: int) -> None:
        self._item = item
        self._end  = end

    @property
    def text(self) -> str:
        return (getattr(self._item, "text", None) or "")[:self._end]

    def __getattr__(self, name: str):
        if name in ("_item", "_end"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextSliceItem:
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


class _ForceListItem:
    __slots__ = ("_item", "_prefix")

    def __init__(self, item, prefix: str) -> None:
        self._item   = item
        self._prefix = prefix

    @property
    def text(self) -> str:
        return self._prefix + (getattr(self._item, "text", None) or "")

    @property
    def label(self):
        class _L:
            value = "list_item"
        return _L()

    def __getattr__(self, name: str):
        if name in ("_item", "_prefix"):
            raise AttributeError(name)
        return getattr(self._item, name)


class _TextSliceAndInsertItem:
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


class _RenumberItem:
    """Перенумеровывает пункт списка: убирает старый номер OCR, ставит новый.

    Текст становится «<num> <содержимое>» (число + пробел, без точки — как в
    исходном OCR), а label принудительно = list_item.
    """
    __slots__ = ("_item", "_num")
    _STRIP_NUM = re.compile(r'^\s*\d+\s*[.)]?\s+')

    def __init__(self, item, num: int) -> None:
        self._item = item
        self._num  = num

    @property
    def text(self) -> str:
        raw = (getattr(self._item, "text", None) or "").lstrip()
        body = self._STRIP_NUM.sub("", raw)
        return f"{self._num} {body}"

    @property
    def label(self):
        class _L:
            value = "list_item"
        return _L()

    def __getattr__(self, name: str):
        if name in ("_item", "_num"):
            raise AttributeError(name)
        return getattr(self._item, name)


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _label_str(item) -> str:
    raw = getattr(item, "label", None)
    if raw is None:
        return ""
    return (raw.value if hasattr(raw, "value") else str(raw)).lower()


# ── Основная функция ──────────────────────────────────────────────────────────

def apply_legal_fixes(all_items: list) -> tuple[list, set]:
    """
    Пост-обработка порядка элементов для судебных документов.

    Возвращает (result, continuation_ids):
      result           — переупорядоченный список элементов
      continuation_ids — set id() элементов для безусловного слияния с предыдущим
    """
    result           = list(all_items)
    n                = len(result)
    continuation_ids: set = set()

    def _lbl(i):  return _label_str(result[i][0])
    def _pg(i):
        pv = (getattr(result[i][0], "prov", None) or [None])[0]
        return int(getattr(pv, "page_no", -1)) if pv else -1
    def _txt(i):  return (getattr(result[i][0], "text", None) or "").strip()

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

    # Fix 4: section_header «по делу №» стоит ПОСЛЕ label:content
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
    for i in range(1, n):
        if _lbl(i) not in _TEXT_LABELS:
            continue
        curr = _txt(i)
        if not _DATE_START_RE.match(curr):
            continue
        prev_i = None
        for k in range(i - 1, max(i - 6, -1), -1):
            if _lbl(k) in _TEXT_LABELS and _pg(k) == _pg(i):
                prev_i = k
                break
            if _pg(k) != _pg(i):
                break
        if prev_i is None:
            log.debug("fix_order: date-start не нашёл предшественника для %r", curr[:40])
            continue
        prev = _txt(prev_i)

        date_match_end = _DATE_START_RE.match(curr).end()
        rest = curr[date_match_end:].lstrip()
        if rest.startswith(','):
            log.debug("fix_order: date-start пропуск (дата рождения) %r", curr[:50])
            continue
        if prev.endswith(('.', '!', '?', '»', ';')):
            log.debug("fix_order: date-start пропуск (предш. заканчивается терминатором) %r", prev[-20:])
            continue

        # Предыдущий блок — предложение-ВВЕДЕНИЕ списка/изложения: «…подтверждается
        # нижеследующим», «…в следующем размере:», «…следующим образом:». Блок с датой
        # и есть это «последующее» содержимое, он должен идти ПОСЛЕ введения — не
        # переставляем. Двоеточие необязательно (OCR часто его теряет).
        if re.search(r'(следующ\w*|образом)\s*[:.]?\s*$', prev, re.IGNORECASE):
            log.debug("fix_order: date-start пропуск (введение списка) %r", prev[-40:])
            continue

        # Предыдущий блок — короткое ALL-CAPS ФИО (2-4 слова, только буквы/пробел/дефис):
        # это значение поля («Должник: ХЛИЯН ЕЛЕНА ОГАНОВНА»), а не незаконченная фраза.
        _prev_words = prev.split()
        _prev_alpha = [c for c in prev if c.isalpha()]
        if (2 <= len(_prev_words) <= 4
                and _prev_alpha
                and all(c.isupper() for c in _prev_alpha)
                and all(c.isalpha() or c.isspace() or c == '-' for c in prev)):
            log.debug("fix_order: date-start пропуск (ALL-CAPS ФИО) %r", prev[:50])
            continue

        log.debug("fix_order: date-start кандидат: prev=%r curr=%r (prev_i=%d, i=%d)",
                  prev[:50], curr[:50], prev_i, i)

        saved = result[i]
        for k in range(i, prev_i, -1):
            result[k] = result[k - 1]
        result[prev_i] = saved
        log.info("fix_order: date-start swap %r ↔ %r", prev[:50], curr[:50])

        if prev_i + 1 < n and _lbl(prev_i + 1) in _TEXT_LABELS:
            b_raw = _txt(prev_i + 1)
            pm = _PATRON_CONT_RE.match(b_raw)
            if pm:
                patron     = pm.group(1)
                b_start    = pm.start(2)
                a_orig     = result[prev_i]
                b_orig     = result[prev_i + 1]
                orig_text  = getattr(b_orig[0], "text", None) or ""
                g2_prefix  = pm.group(2)[:20]
                real_start = orig_text.find(g2_prefix)
                if real_start >= 0:
                    b_start = real_start
                a_text    = getattr(a_orig[0], "text", None) or ""

                fn_match   = _FN_BEFORE_VERB_RE.search(a_text)
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

    # Fix 3: сортировка numbered list_items + заполнение пробелов
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

        has_gap = any(sorted_nums[k + 1] > sorted_nums[k] + 1
                      for k in range(len(sorted_nums) - 1))
        already_sorted = nums_only == sorted_nums
        if already_sorted and not (has_gap and unnumbered):
            i = j
            continue

        final_order:  list[int]      = []
        gap_prefixes: dict[int, str] = {}
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

        _remaining_unnum = unnumbered[unnum_used:]
        if _remaining_unnum and numbered:
            _next_auto = max(num for num, _ in numbered) + 1
            for _gap_k in _remaining_unnum:
                final_order.append(_gap_k)
                gap_prefixes[_gap_k] = f"{_next_auto} "
                log.info("fix_order: автонумер %d %r", _next_auto, txts[_gap_k][:60])
                _next_auto += 1
        else:
            final_order.extend(_remaining_unnum)

        reordered = [
            (_TextPrefixItem(result[i + k][0], gap_prefixes[k]), result[i + k][1])
            if k in gap_prefixes else result[i + k]
            for k in final_order
        ]
        for k, item_lvl in enumerate(reordered):
            result[i + k] = item_lvl

        if not already_sorted:
            log.info("fix_order: sorted %d numbered list items on page %d", len(numbered), pg)
        if has_gap and unnum_used > 0:
            log.info("fix_order: filled %d gap(s) with unnumbered items on page %d", unnum_used, pg)
        i = j

    # Fix 5: list_item стоит ДО heading «ПРОСИТ СУД»
    for i in range(n - 1):
        if _lbl(i) != "list_item":
            continue
        if _lbl(i + 1) not in ("section_header", "text", "title"):
            continue
        if _pg(i) != _pg(i + 1):
            continue
        if not _PROSIT_RE.search(_txt(i + 1)):
            continue
        result[i], result[i + 1] = result[i + 1], result[i]
        log.info("fix_order: ПРОСИТ_СУД swap %r ↔ %r", _txt(i)[:40], _txt(i + 1)[:40])
        num_counter = 1
        for j in range(i + 1, min(i + 10, n)):
            cur_lbl = _lbl(j)
            if cur_lbl not in _PROSIT_LBLS:
                break
            t_j = _txt(j)
            a_j = [c for c in t_j if c.isalpha()]
            if not a_j or a_j[0].islower():
                break
            if not _NUM_ITEM_RE.match(t_j):
                if cur_lbl == "list_item":
                    result[j] = (_TextPrefixItem(result[j][0], f"{num_counter} "), result[j][1])
                else:
                    result[j] = (_ForceListItem(result[j][0], f"{num_counter} "), result[j][1])
                log.info("fix_order: ПРОСИТ_СУД нумерует %d %r (lbl=%s)",
                         num_counter, t_j[:40], cur_lbl)
                num_counter += 1
        break

    # Fix 7: «Приложения:» / «Приложение:» стоит ПОСЛЕ первого пункта приложений.
    # Docling иногда переставляет заголовок раздела вниз.
    # Признак: section_header «Приложени» идёт после list_item на той же странице.
    _PRILOG_RE = re.compile(r'^\s*Приложени', re.IGNORECASE)
    for i in range(1, n):
        if _lbl(i) not in ("section_header", "text"):
            continue
        if not _PRILOG_RE.match(_txt(i)):
            continue
        # Ищем первый list_item на той же странице перед i
        pg_i = _pg(i)
        first_li = None
        for k in range(i - 1, max(i - 10, -1), -1):
            if _pg(k) != pg_i:
                break
            if _lbl(k) == "list_item":
                first_li = k
            elif _lbl(k) in ("section_header", "title"):
                break
        if first_li is not None and first_li < i:
            # Переставляем: заголовок перед первым list_item
            saved = result[i]
            for k in range(i, first_li, -1):
                result[k] = result[k - 1]
            result[first_li] = saved
            log.info("fix_order: Приложения swap → перед list_item[%d] %r",
                     first_li, _txt(first_li)[:40])
            break

    # Fix 8: «Место рождения:» / «Дата рождения:» стоят ПЕРЕД «Должник:»
    # Причина: reading_order_key квантует Y в бины 40pt, и эти три метки попадают
    # в один бин, сортируясь по X. Место(x=172) < Должник(x=207) → неверный порядок.
    # Исправление: находим «Должник:» и переставляем перед ним все «Место/Дата рождения:»
    _DOLZHNIK_RE    = re.compile(r'^\s*Должник\s*:', re.IGNORECASE)
    _ROZHD_RE       = re.compile(r'^\s*(Место|Дата)\s+рождения\s*:', re.IGNORECASE)
    for i in range(n):
        if _lbl(i) not in _TEXT_LABELS:
            continue
        if not _DOLZHNIK_RE.match(_txt(i)):
            continue
        pg_i = _pg(i)
        # Собираем индексы «Место/Дата рождения:» которые идут ПЕРЕД «Должник:» на той же стр.
        rozhd_indices = []
        for k in range(i - 1, max(i - 8, -1), -1):
            if _pg(k) != pg_i:
                break
            if _lbl(k) in _TEXT_LABELS and _ROZHD_RE.match(_txt(k)):
                rozhd_indices.append(k)
        if not rozhd_indices:
            break
        rozhd_indices.sort()  # по возрастанию индекса
        # Получаем Y-координаты для сортировки по реальной позиции
        def _mid_y(idx):
            bx = getattr((getattr(result[idx][0], "prov", None) or [None])[0], "bbox", None)
            if bx is None:
                return 0.0
            return (float(getattr(bx, "t", 0)) + float(getattr(bx, "b", 0))) / 2.0
        # В pdf_native: больший Y = выше (Должник y≈248 > Место y≈233 > Дата y≈221)
        # Нужный порядок: Должник(i), Место(248→233), Дата(227→221)
        # Вставляем rozhd_indices ПОСЛЕ «Должник:»
        # Сортируем rozhd по убыванию mid_y (выше = раньше в pdf_native)
        rozhd_sorted = sorted(rozhd_indices, key=_mid_y, reverse=True)
        # Извлекаем элементы из текущих позиций и вставляем после i
        saved_items = [result[k] for k in rozhd_sorted]
        # Удаляем из оригинальных позиций (с конца чтобы не сдвигать)
        for k in sorted(rozhd_indices, reverse=True):
            result.pop(k)
        # Находим новую позицию i (она сдвинулась если rozhd_indices < i)
        shift = sum(1 for k in rozhd_indices if k < i)
        new_i = i - shift
        # Вставляем после Должника
        for offset, item_lvl in enumerate(saved_items, 1):
            result.insert(new_i + offset, item_lvl)
        n = len(result)
        log.info("fix_order: Должник-first: переставлены %s после позиции %d",
                 [_txt_raw for _txt_raw in [_txt_item[0].text[:30]
                  if hasattr(_txt_item[0], 'text') else '' for _txt_item in saved_items]], new_i)
        break

    # Fix 6: разбивка блоков где OCR объединил несколько абзацев
    new_result = []
    for _item, _level in result:
        segs = [_item]
        for _pat in _PARA_SPLIT_RES:
            new_segs: list = []
            for _seg in segs:
                _raw = getattr(_seg, "text", None) or ""
                _m   = _pat.search(_raw)
                if _m:
                    _p1 = _m.start(0) + 1
                    _p2 = _m.start(1)
                    new_segs.append(_TextTruncateItem(_seg, _p1))
                    new_segs.append(_TextSliceItem(_seg, _p2))
                    log.info("fix_order: split at %r (pos %d/%d)",
                             _pat.pattern[:35], _p1, _p2)
                else:
                    new_segs.append(_seg)
            segs = new_segs
        for _seg in segs:
            new_result.append((_seg, _level))
    result = new_result

    # Fix 7: сортировка пунктов разбивки задолженности по реальной Y-позиции.
    # OCR/Docling иногда переставляет пункты разбивки («комиссия», «штрафы»,
    # «несанкционированный перерасход») и вставляет между ними заключительную фразу
    # «Вышеуказанная задолженность образовалась…». Берём последовательный прогон
    # из ≥3 таких пунктов и сортируем по (страница ↑, mid_y ↓) — это восстанавливает
    # естественный порядок чтения сверху вниз. Универсально: если прогон уже
    # отсортирован (как в блоке «ПРОШУ»), порядок не меняется.
    _BRK_RE = re.compile(
        r'(просроченн\w*\s+основн\w*\s+долг|начисленн\w*\s+процент|штраф|неустойк|'
        r'комисси|несанкционированн\w*\s+перерасход|пени|Вышеуказанн\w*\s+задолж)',
        re.IGNORECASE)
    _CONCL_RE = re.compile(r'Вышеуказанн\w*\s+задолж', re.IGNORECASE)

    def _brk_mid_y(idx):
        bx = getattr((getattr(result[idx][0], "prov", None) or [None])[0], "bbox", None)
        if bx is None:
            return 0.0
        return (float(getattr(bx, "t", 0)) + float(getattr(bx, "b", 0))) / 2.0

    i, n = 0, len(result)
    while i < n:
        if not (_BRK_RE.search(_txt(i)) and _lbl(i) in (_PROSIT_LBLS | {"list_item"})):
            i += 1
            continue
        j = i
        while j < n and _BRK_RE.search(_txt(j)) and _lbl(j) in (_PROSIT_LBLS | {"list_item"}):
            j += 1
        run = list(range(i, j))
        # Считаем именно категории (не заключительную фразу)
        cat = sum(1 for k in run if not _CONCL_RE.search(_txt(k)))
        if cat >= 3:
            order = sorted(run, key=lambda k: (_pg(k), -_brk_mid_y(k)))
            if order != run:
                saved = [result[k] for k in order]
                for off, it in zip(run, saved):
                    result[off] = it
                log.info("fix_order: breakdown-sort %d..%d → %s",
                         i, j - 1, [_txt(k)[:20] for k in range(i, j)])
        i = j

    # Fix 8: блок «Приложения» — восстановить порядок по реальной Y-позиции и
    # перенумеровать пункты приложений. OCR/Docling ставит «Приложения:» раньше
    # пунктов «ПРОШУ» (4–6), путает порядок приложений и дублирует номера. Берём
    # прогон от «Приложения:» до подписи и сортируем по (страница ↑, mid_y ↓) —
    # это и ставит «Приложения:» на место (между п.6 и списком), и упорядочивает
    # приложения; затем перенумеровываем пункты ПОСЛЕ «Приложения:» как 1..N.
    _PRILOZH_RE = re.compile(r'^\s*Приложени[яе]\s*:?\s*$', re.IGNORECASE)
    _SIG_STOP_RE = re.compile(
        r'^\s*(Представитель\b|Электронн\w*\s+подпис)', re.IGNORECASE)
    pril_i = next((k for k in range(len(result)) if _PRILOZH_RE.match(_txt(k))), None)
    if pril_i is not None:
        pg_pr = _pg(pril_i)
        run = [pril_i]
        for k in range(pril_i + 1, len(result)):
            if _pg(k) != pg_pr:
                break
            if _SIG_STOP_RE.match(_txt(k)):
                break
            if _lbl(k) in (_PROSIT_LBLS | {"list_item"}) and _txt(k):
                run.append(k)
            else:
                break
        # Сортируем весь прогон по реальной позиции (выше = раньше)
        if len(run) >= 4:
            order = sorted(run, key=lambda k: (_pg(k), -_brk_mid_y(k)))
            saved = [result[k] for k in order]
            for off, it in zip(run, saved):
                result[off] = it
            # Находим «Приложения:» в новом порядке и нумеруем приложения после него
            new_pril = next((off for off in run if _PRILOZH_RE.match(_txt(off))), None)
            if new_pril is not None:
                num = 0
                for off in run:
                    if off <= new_pril:
                        continue
                    if _lbl(off) != "list_item":
                        continue
                    num += 1
                    result[off] = (_RenumberItem(result[off][0], num), result[off][1])
                log.info("fix_order: приложения отсортированы (%d пунктов перенумерованы)", num)

    # Fix 9: склейка оторванного имени организации. OCR разрезает строку
    # «Получатель: АО «АЛЬФА-БАНК»» на два блока: «…АО» и «АЛЬФА-БАНК»». Если блок
    # оканчивается голой орг.-формой (АО/ПАО/ООО/ОАО/ЗАО), а следующий — короткое
    # ЗАГЛАВНОЕ имя, приклеиваем имя в кавычках к предыдущему блоку.
    # [ОO] — кириллическая или ЛАТИНСКАЯ O (OCR на сыром тексте даёт «АO»).
    _ORGFORM_END_RE = re.compile(r'\b(?:А|ПА|ОО|ОА|ЗА)[ОO]\s*$')
    _ORGNAME_ONLY_RE = re.compile(
        r'^\s*[«"(<]?\s*([А-ЯЁ][А-ЯЁ0-9\-]+(?:\s+[А-ЯЁ0-9\-]+)?)\s*[»"эжх)>]?\s*$')
    _i = 0
    while _i < len(result) - 1:
        if (_lbl(_i) in _PROSIT_LBLS and _lbl(_i + 1) in _PROSIT_LBLS
                and _pg(_i) == _pg(_i + 1)
                and _ORGFORM_END_RE.search(_txt(_i))):
            _m = _ORGNAME_ONLY_RE.match(_txt(_i + 1))
            if _m:
                _name = _m.group(1).strip()
                result[_i] = (_TextAppendItem(result[_i][0], f' «{_name}»'),
                              result[_i][1])
                result.pop(_i + 1)
                log.info("fix_order: склейка орг-имени → «%s»", _name)
                continue
        _i += 1

    return result, continuation_ids
