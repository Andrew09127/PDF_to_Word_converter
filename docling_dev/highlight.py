"""Подсветка «подозрительных» (вероятно искажённых OCR) слов в DOCX.

Детерминированно, без LLM/GPU: помечаем то, что безопасно автоматически НЕ
исправить, чтобы человек-проверяющий сразу видел сомнительные места.

Эвристики «подозрительного» слова (низкий процент ложных срабатываний):
  1) смешанный регистр-внутри-слова: строчная, затем заглавная («должНЫ»,
     «тысЯЧ», «ПОЧтОВЫм») — типичный OCR-капс-шум;
  2) смешанные алфавиты в одном слове (кириллица + латиница: «Ng», «д0говор»);
  3) одиночный латинский токен (≤5 букв) среди кириллицы и без цифр
     («co», «CO», «CT», «Bank», «Ne», «Hi») — почти всегда гомоглиф;
  4) остаточные кавычки-гомоглифы, прилипшие к слову («Коллектэ», «банкэ»);
  5) если доступна pymorphy3/pymorphy2: слово ≥6 букв из чистой кириллицы, не
     распознанное морфологически («нанченования», «Крецитного» и т.п.).
Числа, ИНН/коды (с цифрами), email/URL не помечаются.

Часть помеченных слов чинится ДЕТЕРМИНИРОВАННО ещё до подсветки (_autofix_word):
латиница→кириллица, регистр, и словарный спелл-фикс кириллических OCR-путаниц
(_try_cyr_spell_fix: н↔и, с↔е, о↔а … — берём правку только при ЕДИНСТВЕННОМ
словарном кандидате, иначе слово уходит на подсветку/LLM).
"""
from __future__ import annotations

import copy
import logging
import re
from functools import lru_cache

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

log = logging.getLogger(__name__)

# ── Опциональный морфологический словарь (pymorphy3, фолбэк pymorphy2) ────────
# Если пакет не установлен — продолжаем без него (только эвристики).
try:
    import pymorphy3 as _pm2
    _morph = _pm2.MorphAnalyzer()
    log.debug("highlight: pymorphy3 доступен — включена словарная проверка")
except Exception:
    try:
        import pymorphy2 as _pm2
        _morph = _pm2.MorphAnalyzer()
        log.debug("highlight: pymorphy2 доступен — включена словарная проверка")
    except Exception:
        _morph = None

_CYR = "А-Яа-яЁё"
_LAT = "A-Za-z"

# Токен = последовательность букв/цифр/дефисов/подчёркиваний/собак/точек внутри.
_TOKEN_RE = re.compile(r"[^\s]+")

_CAPS_MID_RE    = re.compile(rf"[{_CYR}]*[а-яё][А-ЯЁ][{_CYR}]*")   # строчная→заглавная
_HAS_CYR        = re.compile(rf"[{_CYR}]")
_HAS_LAT        = re.compile(rf"[{_LAT}]")
_HAS_DIGIT      = re.compile(r"\d")
# Был {1,3} — расширён до {1,5}: «Hi», «Ne», «Bank», «Name» тоже гомоглифы
_PURE_LAT_SHORT = re.compile(rf"^[{_LAT}]{{1,5}}$")
_URL_EMAIL_RE   = re.compile(r"[@./]|https?|www|\.ru|\.com", re.IGNORECASE)
# «слово» с прилипшим гомоглифом-кавычкой на конце (э/ж/х/» после строчной буквы)
_TRAIL_QUOTE_RE = re.compile(rf"[{_CYR}]{{3,}}[эжхъ]$")

_STRIP = " \t.,;:!?()«»\"'–—-"

# Аббревиатуры и ALL-CAPS токены ≤6 — ИНН, ОГРН, РФ, АО — не проверяем словарём.
_ABBREV_RE = re.compile(r"^[А-ЯЁ]{2,6}$")

# Доменные юр./фин. термины, которых НЕТ в словаре pymorphy, но они ВЕРНЫЕ.
# Без этого списка «взыскателя» чинилось бы на «изыскателя», «займодавец» на
# архаичное «заимодавец», а правильные «микрофинансовый»/«коллекторская»
# подсвечивались бы как ошибки. Сравниваем по началу слова (стем), без регистра.
_DOMAIN_OK_STEMS: tuple[str, ...] = (
    "взыскател", "займодав", "заимодав", "залогодержател", "залогодател",
    "цессион", "цедент", "цессионар", "микрофинанс", "коллекторск",
    "созаёмщик", "созаемщик", "поручительств", "неустойк", "потребительск",
    "правоустанавлива", "правопреемник", "правопреемств", "реструктуризац",
    "досудебн", "внесудебн", "подведомственн", "подсудн",
)


def _domain_known(word: str) -> bool:
    """True если слово — известный доменный термин (нет в pymorphy, но верный)."""
    w = word.lower()
    return any(w.startswith(s) for s in _DOMAIN_OK_STEMS)

# Кластеры визуально похожих кириллических букв — OCR путает их между собой.
# Пары ДВУНАПРАВЛЕННЫЕ: «и↔н» порождает и подстановку и→н, и обратную н→и.
# Только высоко-визуальные путаницы (мин. ложных): вертикальные штрихи (и/н/п/й/м),
# открытые дуги (с/е/о), диакритика (е/ё, и/й), хвосты (ц/и), засечки (г/т).
_CYR_OCR_PAIRS: tuple[tuple[str, str], ...] = (
    ("и", "н"), ("н", "п"), ("и", "й"), ("н", "й"), ("н", "м"),
    ("с", "е"), ("о", "а"), ("о", "с"), ("е", "ё"),
    ("ц", "и"), ("л", "п"), ("в", "н"), ("г", "т"), ("ь", "ы"),
)
# Разворачиваем пары в словарь: буква → кортеж её визуальных двойников.
_CYR_CONFUSIONS: dict[str, list[str]] = {}
for _a, _b in _CYR_OCR_PAIRS:
    _CYR_CONFUSIONS.setdefault(_a, []).append(_b)
    _CYR_CONFUSIONS.setdefault(_b, []).append(_a)

_CYR_ONLY      = re.compile(r"^[а-яё]+$")
_SPELL_MIN_LEN = 5       # слова короче — не трогаем (риск исказить «дом», «код»)
_SPELL_DEPTH   = 2       # максимум исправляемых ошибок на слово (1–2)
_SPELL_BUDGET  = 6000    # потолок перебора: слишком ветвистое слово → подсветка


@lru_cache(maxsize=20000)
def _pymorphy_known(word: str) -> bool:
    """True если pymorphy знает хотя бы одно осмысленное прочтение слова.

    Кэшируется: BFS-перебор спелл-фиксера проверяет одни и те же кандидаты
    многократно, а словарный разбор — самая дорогая операция в модуле.
    """
    if _morph is None:
        return True          # без словаря считаем все слова «известными»
    parses = _morph.parse(word.lower())
    # pymorphy3: is_known=True → слово в словаре
    # pymorphy2: UNKN POS → не в словаре (fallback)
    if hasattr(parses[0], "is_known"):
        return any(p.is_known for p in parses)
    return any(p.tag.POS not in (None, "UNKN") for p in parses)


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
    # 3) короткий (≤5 букв) чисто-латинский токен (гомоглиф среди кириллицы)
    if has_lat and not has_cyr and _PURE_LAT_SHORT.match(core):
        return True
    # 1) капс в середине слова
    if has_cyr and _CAPS_MID_RE.fullmatch(core):
        return True
    # 5) словарная проверка: только кириллица, ≥6 букв, не аббревиатура
    if has_cyr and not has_lat and len(core) >= 6 and not _ABBREV_RE.match(core):
        if _domain_known(core):
            return False                  # верный доменный термин — не шумим
        if not _pymorphy_known(core):
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


# Частые короткие предлоги/союзы — для снятия капс-шума (пО→по, cO→со).
# НЕ включаем «кв» и т.п. единицы, чтобы не сломать «кВ» (киловольт).
_COMMON_SHORT = {"по", "со", "из", "от", "на", "не", "до", "за", "об", "во",
                 "ко", "но", "да", "и", "в", "с", "к", "у", "о", "а"}


# Латиница → кириллица (визуальные двойники) — для очистки гомоглифов
_LAT2CYR = {
    "a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М",
    "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У",
}


def _apply_case_pattern(fixed: str, original: str) -> str:
    """Переносит регистровый рисунок оригинала на исправленное слово:
    ВЕСЬ ВЕРХНИЙ → ВЕРХНИЙ, Первая-заглавная → Первая-заглавная, иначе строчный."""
    if len(original) > 1 and original.isupper():
        return fixed.upper()
    if original[:1].isupper():
        return fixed[:1].upper() + fixed[1:]
    return fixed


def _try_cyr_spell_fix(word: str) -> str | None:
    """Исправляет чисто-кириллические OCR-ошибки словарным перебором (BFS).

    На каждом шаге глубины подставляем визуальные двойники букв (_CYR_CONFUSIONS),
    pymorphy выступает оракулом «это настоящее слово». Возвращаем правку ТОЛЬКО
    если на минимальной глубине нашёлся РОВНО ОДИН словарный кандидат — иначе
    (0 или ≥2) оставляем слово на подсветку/LLM, чтобы не выдумать неверное
    написание. Примеры: нмущества→имущества, абластн→области, Федсрацин→Федерации.
    """
    if _morph is None or _HAS_LAT.search(word):
        return None
    # Имена собственные/фамилии (Первая-Заглавная, но не ВЕСЬ-ВЕРХНИЙ) словарным
    # перебором НЕ трогаем: «Заречнев»→«Заречней», «Аракслян»→«Аракелян» —
    # недопустимый риск. ALL-CAPS (заголовки) и строчные слова — разрешены.
    if word[:1].isupper() and not word.isupper():
        return None
    core = word.lower()
    if len(core) < _SPELL_MIN_LEN or not _CYR_ONLY.match(core):
        return None
    if _pymorphy_known(core) or _domain_known(core):
        return None          # уже нормальное/доменное слово — не выдумываем правку

    seen: set[str] = {core}
    frontier: list[str] = [core]
    budget = _SPELL_BUDGET

    for _ in range(_SPELL_DEPTH):
        found: set[str] = set()
        nxt: list[str] = []
        for w in frontier:
            for i, ch in enumerate(w):
                for repl in _CYR_CONFUSIONS.get(ch, ()):
                    cand = w[:i] + repl + w[i + 1:]
                    if cand in seen:
                        continue
                    seen.add(cand)
                    budget -= 1
                    if budget <= 0:
                        return None          # слишком ветвисто — безопаснее подсветить
                    if _pymorphy_known(cand):
                        found.add(cand)
                    else:
                        nxt.append(cand)
        if found:
            # единственный словарный кандидат на этой глубине → принимаем;
            # несколько (неоднозначность) → отдаём LLM/человеку с контекстом.
            return _apply_case_pattern(next(iter(found)), word) if len(found) == 1 else None
        frontier = nxt
        if not frontier:
            break

    return None


def _autofix_word(core: str) -> str | None:
    """Пытается ДЕТЕРМИНИРОВАННО починить подозрительное слово.

    1) случайные заглавные OCR (если строчных больше заглавных) → строчим;
    2) латинские буквы-двойники → кириллица (co→со, B→В, OOО→ООО, КB→КВ).
    Возвращает исправленное слово, если оно стало «чистым» (не подозрительным),
    иначе None (оставляем подсветку).
    """
    w = core
    # 1) латинские буквы-двойники → кириллица (регистр сохраняем): co→со, тоM→тоМ
    if _HAS_CYR.search(w) or _PURE_LAT_SHORT.match(w):
        w2 = "".join(_LAT2CYR.get(ch, ch) for ch in w)
        if not _HAS_LAT.search(w2):     # после замены латиницы не осталось
            w = w2
    # 2) случайные заглавные OCR (строчных больше заглавных) → строчим; но
    #    сохраняем возможную легитимную ПЕРВУЮ заглавную (Сбер, Москва, тоМ→том)
    n_low = sum(1 for c in w if c.islower())
    n_up  = sum(1 for c in w if c.isupper())
    if n_up and n_low > n_up:
        w = (w[0] + w[1:].lower()) if w[0].isupper() else w.lower()
    # 3) короткое слово со «скачущим» регистром, дающее частый предлог/союз:
    #    пО→по, cO→сО→со. Белый список защищает единицы (кВ→кв НЕ трогаем).
    if len(w) <= 3 and w.lower() in _COMMON_SHORT:
        w = w.lower()
    if w != core and not _is_suspicious(w):
        return w
    # 4) кириллические OCR-путаницы (н↔и, а↔о, с↔е, ё↔е) — словарный BFS
    spell = _try_cyr_spell_fix(w)
    if spell is not None and not _is_suspicious(spell):
        return spell
    return None


_ROMAN_RE = re.compile(r"^[IVXLCDM]{1,7}$", re.IGNORECASE)


def _is_garbled_cyrillic(core: str) -> bool:
    """Стоит ли подсвечивать слово, которое автофикс НЕ починил.

    Да — если это искажённое РУССКОЕ слово (есть кириллица): человек поправит.
    Нет — если чисто-латинский токен (URL-обрывок «ru»/«pro», код), римская
    цифра (IV, VII) или мусор: это не «слово с опечаткой», подсветка только шумит
    (и LLM на таком только галлюцинирует, напр. VII→VIII).

    Чисто-латинные токены ≤5 букв — подсвечиваем (гомоглифы в русском тексте).
    Длинные латинские слова (>5) — скорее код/URL/имя — не трогаем.
    """
    if _ROMAN_RE.match(core):
        return False                         # римские цифры (IV, VII) — валидны
    if _PURE_LAT_SHORT.match(core):
        return True                          # короткая/средняя латиница → ошибка
    n_cyr = len(re.findall(rf"[{_CYR}]", core))
    n_lat = len(re.findall(rf"[{_LAT}]", core))
    # для смешанных: подсвечиваем, если слово ПРЕИМУЩЕСТВЕННО кириллическое;
    # латиницы больше → обрывок URL/латинского слова (лnalog) — не подсветка
    return n_cyr > 0 and n_cyr >= n_lat


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
