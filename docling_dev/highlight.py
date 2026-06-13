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


def _segments(text: str, spans: list[tuple[int, int]]):
    segs, pos = [], 0
    for s, e in spans:
        if s > pos:
            segs.append((text[pos:s], False))
        segs.append((text[s:e], True))
        pos = e
    if pos < len(text):
        segs.append((text[pos:], False))
    return segs


def _has_complex_content(r) -> bool:
    """True если run содержит не только текст (переносы, картинки и т.п.)."""
    for ch in r.iterchildren():
        t = ch.tag.split("}")[-1]
        if t not in ("rPr", "t"):
            return True
    return False


def _highlight_run(run) -> int:
    spans = suspicious_spans(run.text)
    if not spans or _has_complex_content(run._r):
        return 0
    segs = _segments(run.text, spans)
    orig_r = copy.deepcopy(run._r)          # шаблон форматирования
    run.text = segs[0][0]
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW if segs[0][1] else None
    anchor = run._r
    for seg_text, susp in segs[1:]:
        new_r = copy.deepcopy(orig_r)
        anchor.addnext(new_r)
        nr = Run(new_r, run._parent)
        nr.text = seg_text
        nr.font.highlight_color = WD_COLOR_INDEX.YELLOW if susp else None
        anchor = new_r
    return sum(1 for _, s in segs if s)


def highlight_suspicious(doc) -> int:
    """Подсвечивает жёлтым все подозрительные слова в документе. Возвращает счётчик."""
    n = 0
    for para in _iter_paragraphs(doc):
        for run in list(para.runs):
            try:
                n += _highlight_run(run)
            except Exception as exc:               # один run не должен ломать документ
                log.debug("highlight: пропуск run: %s", exc)
    log.info("highlight: помечено %d подозрительных фрагментов", n)
    return n
