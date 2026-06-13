"""Локальная LLM-доочистка подсвеченных (не исправленных детерминированно) слов.

ПОЛНОСТЬЮ ЛОКАЛЬНО, БЕЗ ОТДЕЛЬНЫХ ПРОГРАММ:
  • движок — pip-пакет `llama-cpp-python` (вендорится в git, как torch);
  • модель — файл `.gguf` в папке `models/` (тоже в git);
  • инференс идёт В САМОМ процессе Python — ни сети, ни службы, ни установки
    на целевом ПК (работает в закрытой корпоративной сети офлайн).
  • ОПЦИОНАЛЬНО: если пакета или файла модели нет — модуль ничего не делает
    (no-op), программа продолжает работать как обычно (автоочистка + подсветка).

Обрабатываются ТОЛЬКО уже подсвеченные жёлтым слова (остаток после
детерминированной автоочистки) — это ~10 коротких фрагментов на документ,
поэтому быстро даже на CPU. Жёсткие предохранители не дают модели менять
цифры/номера и «галлюцинировать» (подменять слово другим).
"""
from __future__ import annotations

import glob
import logging
import os
import re
import shutil

from docx.enum.text import WD_COLOR_INDEX

from .highlight import _iter_paragraphs

log = logging.getLogger(__name__)

# Модель по умолчанию — вендоренный .gguf в папке models/ рядом с проектом.
_DEFAULT_MODEL = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "models", "qwen2.5-1.5b-instruct-q4_k_m.gguf",
)

_DIGIT_RE = re.compile(r"\d")
_HAS_CYR  = re.compile(r"[А-Яа-яЁё]")
_ROMAN_RE = re.compile(r"^[IVXLCDM]{1,7}$", re.IGNORECASE)
_SYSTEM = ("Ты исправляешь ошибки распознавания (OCR) в русском юридическом "
           "тексте. В ответе — ТОЛЬКО одно правильное русское слово, без кавычек "
           "и пояснений. Не меняй цифры, номера, ФИО.")

_llm = None          # ленивый singleton модели (грузим один раз на весь батч)
_llm_failed = False


def _resolve_model(model_path: str | None) -> str | None:
    """Путь к .gguf. Если собранного файла нет, но рядом лежат куски
    `<имя>.gguf.partNNN` (модель режут на <100 МБ ради лимита GitHub) —
    склеиваем их один раз в полный файл и далее используем его."""
    p = model_path or _DEFAULT_MODEL
    if os.path.isfile(p):
        return p
    parts = sorted(glob.glob(p + ".part*"))
    if parts:
        try:
            log.info("LLM: собираю модель из %d кусков…", len(parts))
            with open(p, "wb") as out:
                for part in parts:
                    with open(part, "rb") as f:
                        shutil.copyfileobj(f, out, 1024 * 1024)
            return p
        except Exception as exc:
            log.warning("LLM: не удалось собрать модель из кусков: %s", exc)
            try:
                if os.path.isfile(p):
                    os.remove(p)              # убираем частично записанный файл
            except OSError:
                pass
    return None


def llm_available(model_path: str | None = None) -> bool:
    """True если есть и пакет llama-cpp-python, и файл модели."""
    if _resolve_model(model_path) is None:
        return False
    try:
        import llama_cpp  # noqa: F401
        return True
    except Exception:
        return False


def _get_llm(model_path: str):
    global _llm, _llm_failed
    if _llm is None and not _llm_failed:
        try:
            from llama_cpp import Llama
            _llm = Llama(model_path=model_path, n_ctx=512,
                         n_threads=os.cpu_count() or 4, verbose=False)
            log.info("LLM: модель загружена (%s)", os.path.basename(model_path))
        except Exception as exc:
            _llm_failed = True
            log.warning("LLM: не удалось загрузить модель: %s", exc)
    return _llm


def _ask(llm, word: str, ctx: str) -> str | None:
    try:
        out = llm.create_chat_completion(
            messages=[
                {"role": "system", "content": _SYSTEM},
                {"role": "user", "content":
                    f"Предложение: «{ctx[:400]}»\n"
                    f"Искажённое слово: «{word}»\nПравильное слово:"},
            ],
            temperature=0.0, max_tokens=16,
        )
        return (out["choices"][0]["message"]["content"] or "").strip()
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
    if _ROMAN_RE.match(orig) or not _HAS_CYR.search(orig):
        return None                          # римские цифры/коды/латиница — не правим
    if " " in cand or "\t" in cand:          # должно остаться ОДНО слово
        return None
    if _digits(cand) != _digits(orig):       # цифры менять запрещено
        return None
    # Правка должна быть МЕЛКОЙ (починка OCR), а не подмена другим словом.
    if _edit_distance(orig.lower(), cand.lower()) > max(2, len(orig) // 3):
        return None
    return cand


def correct_highlighted(doc, model_path: str | None = None) -> int:
    """Прогоняет подсвеченные слова через локальную модель; что удалось безопасно
    исправить — заменяет и снимает подсветку. Возвращает число исправленных."""
    resolved = _resolve_model(model_path)
    if resolved is None:
        log.info("LLM-доочистка: модель не найдена (%s) — пропуск",
                 model_path or _DEFAULT_MODEL)
        return 0
    llm = _get_llm(resolved)
    if llm is None:
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
            cand = _accept(word, _ask(llm, word, ctx))
            if cand:
                run.text = run.text.replace(word, cand)
                run.font.highlight_color = None
                fixed += 1
                log.info("LLM: %r → %r", word, cand)
    log.info("LLM-доочистка: исправлено %d фрагментов", fixed)
    return fixed
