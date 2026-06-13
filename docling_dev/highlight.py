"""Подсветка «подозрительных» (вероятно искажённых OCR) слов в DOCX.

Детерминированно, без LLM/GPU: помечаем то, что безопасно автоматически НЕ
исправить, чтобы человек-проверяющий сразу видел сомнительные места.

Эвристики «подозрительного» слова (низкий процент ложных срабатываний):
  1) смешанный регистр-внутри-слова: строчная, затем заглавная («должНЫ»,
     «тысЯЧ», «ПОЧтОВЫм») — типичный OCR-капс-шум;
  2) смешанные алфавиты в одном слове (кириллица + латиница: «Ng», «д0говор»);
  3) одиночный КОРОТКИЙ латинский токен (≤3 буквы) среди кириллицы и без цифр
     («co», «CO», «CT», «B») — почти всегда гомоглиф;
  4) остаточные кавычки-гомоглифы, прилипшие к слову («Коллектэ», «банкэ»).
Числа, ИНН/коды (с цифрами), email/URL не помечаются.
"""
from __future__ import annotations

import copy
import logging
import re

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

log = logging.getLogger(__name__)

_CYR = "А-Яа-яЁё"
_LAT = "A-Za-z"

# Токен = последовательность букв/цифр/дефисов/подчёркиваний/собак/точек внутри.
_TOKEN_RE = re.compile(r"[^\s]+")

_CAPS_MID_RE   = re.compile(rf"[{_CYR}]*[а-яё][А-ЯЁ][{_CYR}]*")   # строчная→заглавная
_HAS_CYR       = re.compile(rf"[{_CYR}]")
_HAS_LAT       = re.compile(rf"[{_LAT}]")
_HAS_DIGIT     = re.compile(r"\d")
_PURE_LAT_SHORT = re.compile(rf"^[{_LAT}]{{1,3}}$")
_URL_EMAIL_RE  = re.compile(r"[@./]|https?|www|\.ru|\.com", re.IGNORECASE)
# «слово» с прилипшим гомоглифом-кавычкой на конце (э/ж/х/» после строчной буквы)
_TRAIL_QUOTE_RE = re.compile(rf"[{_CYR}]{{3,}}[эжхъ]$")

_STRIP = " \t.,;:!?()«»\"'–—-"


def _is_suspicious(core: str) -> bool:
    if not core:
        return False
    if _HAS_DIGIT.search(core):
        return False                      # коды/числа/ИНН — не трогаем
    if _URL_EMAIL_RE.search(core):
        return False                      # email/URL — не трогаем
    has_cyr = bool(_HAS_CYR.search(core))
    has_lat = bool(_HAS_LAT.search(core))
    # 2) смешанные алфавиты
    if has_cyr and has_lat:
        return True
    # 3) короткий чисто-латинский токен (гомоглиф среди кириллицы)
    if has_lat and not has_cyr and _PURE_LAT_SHORT.match(core):
        return True
    # 1) капс в середине слова
    if has_cyr and _CAPS_MID_RE.fullmatch(core):
        return True
    return False


def suspicious_spans(text: str) -> list[tuple[int, int]]:
    """Возвращает список (start, end) символьных диапазонов подозрительных слов."""
    spans: list[tuple[int, int]] = []
    for m in _TOKEN_RE.finditer(text):
        tok = m.group(0)
        # обрезаем обрамляющую пунктуацию/кавычки для анализа «ядра»
        core = tok.strip(_STRIP)
        if not core:
            continue
        if _is_suspicious(core):
            off = m.start() + tok.find(core)
            spans.append((off, off + len(core)))
    return spans


# Латиница → кириллица (визуальные двойники) — для очистки гомоглифов
_LAT2CYR = {
    "a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М",
    "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У",
}


def _autofix_word(core: str) -> str | None:
    """Пытается ДЕТЕРМИНИРОВАННО починить подозрительное слово.

    1) случайные заглавные OCR (если строчных больше заглавных) → строчим;
    2) латинские буквы-двойники → кириллица (co→со, B→В, OOО→ООО, КB→КВ).
    Возвращает исправленное слово, если оно стало «чистым» (не подозрительным),
    иначе None (оставляем подсветку).
    """
    w = core
    n_low = sum(1 for c in w if c.islower())
    n_up  = sum(1 for c in w if c.isupper())
    if n_up and n_low > n_up:           # слово в нижнем регистре со «скачущими» CAPS
        w = w.lower()
    if _HAS_CYR.search(w) or _PURE_LAT_SHORT.match(w):
        w2 = "".join(_LAT2CYR.get(ch, ch) for ch in w)
        if not _HAS_LAT.search(w2):     # после замены латиницы не осталось
            w = w2
    if w != core and not _is_suspicious(w):
        return w
    return None


_ROMAN_RE = re.compile(r"^[IVXLCDM]{1,7}$", re.IGNORECASE)


def _is_garbled_cyrillic(core: str) -> bool:
    """Стоит ли подсвечивать слово, которое автофикс НЕ починил.

    Да — если это искажённое РУССКОЕ слово (есть кириллица): человек поправит.
    Нет — если чисто-латинский токен (URL-обрывок «ru»/«pro», код), римская
    цифра (IV, VII) или мусор: это не «слово с опечаткой», подсветка только шумит
    (и LLM на таком только галлюцинирует, напр. VII→VIII).
    """
    if _ROMAN_RE.match(core):
        return False
    return bool(_HAS_CYR.search(core))


def _iter_paragraphs(parent):
    """Все абзацы документа, включая вложенные в ячейки таблиц (рекурсивно)."""
    body = parent.element.body if hasattr(parent, "element") else parent._element
    yield from _iter_in(body, parent)


def _iter_in(element, doc):
    for child in element.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            tbl = Table(child, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    yield from _iter_in(cell._tc, doc)


def _has_complex_content(r) -> bool:
    """True если run содержит не только текст (переносы, картинки и т.п.)."""
    for ch in r.iterchildren():
        t = ch.tag.split("}")[-1]
        if t not in ("rPr", "t"):
            return True
    return False


def _process_run(run) -> tuple[int, int]:
    """Для каждого подозрительного слова: либо ДЕТЕРМИНИРОВАННО чиним (без
    подсветки), либо подсвечиваем жёлтым (если автопочинить нельзя).
    Возвращает (подсвечено, исправлено)."""
    spans = suspicious_spans(run.text)
    if not spans or _has_complex_content(run._r):
        return (0, 0)
    text = run.text
    segs: list[tuple[str, bool]] = []   # (текст, подсвечивать?)
    pos, n_hl, n_fix = 0, 0, 0
    for s, e in spans:
        if s > pos:
            segs.append((text[pos:s], False))
        word  = text[s:e]
        fixed = _autofix_word(word)
        if fixed is not None:
            segs.append((fixed, False)); n_fix += 1     # очищено — без подсветки
        elif _is_garbled_cyrillic(word):
            segs.append((word, True));  n_hl += 1       # искажённое рус. слово — подсветка
        else:
            segs.append((word, False))                  # URL-обрывок/римское/код — не трогаем
        pos = e
    if pos < len(text):
        segs.append((text[pos:], False))

    orig_r = copy.deepcopy(run._r)          # шаблон форматирования
    run.text = segs[0][0]
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW if segs[0][1] else None
    anchor = run._r
    for seg_text, hl in segs[1:]:
        new_r = copy.deepcopy(orig_r)
        anchor.addnext(new_r)
        nr = Run(new_r, run._parent)
        nr.text = seg_text
        nr.font.highlight_color = WD_COLOR_INDEX.YELLOW if hl else None
        anchor = new_r
    return (n_hl, n_fix)


def highlight_suspicious(doc) -> int:
    """Чистит детерминируемые OCR-ошибки в помеченных словах и подсвечивает
    остаток жёлтым. Возвращает число подсвеченных (оставшихся) фрагментов."""
    n_hl = n_fix = 0
    for para in _iter_paragraphs(doc):
        for run in list(para.runs):
            try:
                hl, fix = _process_run(run)
                n_hl += hl; n_fix += fix
            except Exception as exc:               # один run не должен ломать документ
                log.debug("highlight: пропуск run: %s", exc)
    log.info("highlight: исправлено %d, подсвечено %d (осталось) фрагментов",
             n_fix, n_hl)
    return n_hl
