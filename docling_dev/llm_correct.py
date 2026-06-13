"""Локальная LLM-доочистка подсвеченных (не исправленных детерминированно) слов.

ПОЛНОСТЬЮ ЛОКАЛЬНО и ОПЦИОНАЛЬНО:
  • обращение к Ollama по HTTP на localhost:11434 — наружу ничего не уходит;
  • без новых pip-зависимостей (только стандартная библиотека urllib/json);
  • если Ollama не установлена/не запущена — модуль ничего не делает (no-op),
    программа продолжает работать как обычно (автоочистка + подсветка).

Обрабатываются ТОЛЬКО уже подсвеченные жёлтым слова (остаток после
детерминированной автоочистки) — это ~10 коротких фрагментов на документ,
поэтому быстро даже на CPU. Жёсткие предохранители не дают модели менять
цифры/номера и «галлюцинировать» длинные фразы.
"""
from __future__ import annotations

import json
import logging
import re
import urllib.request

from docx.enum.text import WD_COLOR_INDEX

from .highlight import _iter_paragraphs

log = logging.getLogger(__name__)

_BASE      = "http://localhost:11434"
_TAGS_URL  = _BASE + "/api/tags"
_GEN_URL   = _BASE + "/api/generate"
_DIGIT_RE  = re.compile(r"\d")
_DEFAULT_MODEL = "qwen2.5:3b"

_PROMPT = (
    "Ты исправляешь ошибки распознавания (OCR) в русском юридическом тексте.\n"
    "В предложении одно слово распознано неверно (смешаны латиница/кириллица, "
    "случайные заглавные и т.п.).\n"
    "Предложение: «{ctx}»\n"
    "Искажённое слово: «{word}»\n"
    "Верни ТОЛЬКО одно правильное русское слово — без кавычек, без пояснений. "
    "НЕ меняй цифры, номера, ФИО. Если слово уже верное — верни его без изменений."
)


def ollama_available(timeout: float = 2.0) -> bool:
    """True если локальный сервер Ollama отвечает."""
    try:
        urllib.request.urlopen(_TAGS_URL, timeout=timeout)
        return True
    except Exception:
        return False


def _ask(word: str, ctx: str, model: str, timeout: float = 60.0) -> str | None:
    payload = json.dumps({
        "model": model,
        "prompt": _PROMPT.format(ctx=ctx[:400], word=word),
        "stream": False,
        "options": {"temperature": 0.0, "num_predict": 16},
    }).encode("utf-8")
    try:
        req = urllib.request.Request(
            _GEN_URL, data=payload, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return (json.loads(r.read().decode("utf-8")).get("response") or "").strip()
    except Exception as exc:
        log.debug("LLM: запрос не удался: %s", exc)
        return None


def _digits(s: str) -> str:
    return "".join(_DIGIT_RE.findall(s))


def _edit_distance(a: str, b: str) -> int:
    """Расстояние Левенштейна (для сравнения исходного и исправленного слова)."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[-1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _accept(orig: str, cand: str | None) -> str | None:
    """Предохранители: принимаем правку, только если она безопасна и МЕЛКАЯ."""
    if not cand:
        return None
    cand = cand.splitlines()[0].strip().strip("«»\"'.,;:()")
    if not cand or cand == orig:
        return None
    if " " in cand or "\t" in cand:          # должно остаться ОДНО слово
        return None
    if _digits(cand) != _digits(orig):       # цифры менять запрещено
        return None
    # Правка должна быть МЕЛКОЙ (починка OCR), а не подмена другим словом.
    # Иначе модель «галлюцинирует» (пО → пОТОМУ). Порог растёт с длиной слова.
    if _edit_distance(orig.lower(), cand.lower()) > max(2, len(orig) // 3):
        return None
    return cand


def correct_highlighted(doc, model: str = _DEFAULT_MODEL) -> int:
    """Прогоняет подсвеченные слова через локальную LLM; что удалось безопасно
    исправить — заменяет и снимает подсветку. Возвращает число исправленных."""
    if not ollama_available():
        log.info("LLM-доочистка: Ollama недоступна (localhost:11434) — пропуск")
        return 0
    fixed = 0
    for para in _iter_paragraphs(doc):
        ctx = para.text
        for run in para.runs:
            if run.font.highlight_color != WD_COLOR_INDEX.YELLOW:
                continue
            word = run.text.strip()
            if not word:
                continue
            cand = _accept(word, _ask(word, ctx, model))
            if cand:
                run.text = run.text.replace(word, cand)
                run.font.highlight_color = None
                fixed += 1
                log.info("LLM: %r → %r", word, cand)
    log.info("LLM-доочистка: исправлено %d фрагментов", fixed)
    return fixed
