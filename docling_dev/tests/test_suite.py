"""
test_suite.py — полный набор unit-тестов для пакета docling_dev.
Запуск: python -m pytest docling_dev/tests/ -v
"""
from __future__ import annotations
import sys
import types
from pathlib import Path

import pytest

# Добавляем корень проекта в sys.path
ROOT = Path(__file__).parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def make_bbox(t, b, l=0, r=100):
    return types.SimpleNamespace(t=t, b=b, l=l, r=r)


def make_pts(x0, y0, x1, y1):
    return [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]


def make_ocr(x0, y0, x1, y1, text, conf=0.9):
    return (make_pts(x0, y0, x1, y1), text, conf)


# ─────────────────────────────────────────────────────────────────────────────
#  OCR FIXES
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.ocr_fixes import postprocess, fix_quotes

@pytest.mark.parametrize("inp,expected", [
    # ── Кавычки ──────────────────────────────────────────────────────────────
    ("<Центр-инвест>",          "«Центр-инвест»"),
    ("<Центр-инвестэ текст",    "«Центр-инвест» текст"),
    ("<Авангардж текст",        "«Авангард» текст"),
    ("<Центр-инвест» текст",    "«Центр-инвест» текст"),
    ("@Центр-инвестэ текст",    "«Центр-инвест» текст"),
    ("@Центр-инвестж текст",    "«Центр-инвест» текст"),
    ("@Центр-инвест» текст",    "«Центр-инвест» текст"),
    ("@Центр-инвест> текст",    "«Центр-инвест» текст"),
    ("<<Центр-инвест>>",        "«Центр-инвест»"),
    ('"Центр-инвест"',          "«Центр-инвест»"),
    ("«Центр-инвест»",          "«Центр-инвест»"),   # уже корректно
    ("«ентр-инвест»",           "«Центр-инвест»"),   # OCR потерял Ц

    # ── Пунктуация ────────────────────────────────────────────────────────────
    ("ИНН 6163011391,; ОГРН",   "ИНН 6163011391, ОГРН"),
    ("р-н; х Ленинакан",        "р-н, х Ленинакан"),
    ("; пр. Соколова",          ", пр. Соколова"),

    # ── г. / пр. ─────────────────────────────────────────────────────────────
    ("&. Ростов",               "г. Ростов"),
    ("; &. Ростов",             "; г. Ростов"),
    # Latin P вместо Cyrillic Р — OCR путает («Pocmов» вместо «Ростов»)
    ("&. Pocmов-на-Дону",       "г. Pocmов-на-Дону"),
    ("ир; Соколова",            "пр. Соколова"),
    ("ир: Соколова",            "пр. Соколова"),
    # Latin p вместо Cyrillic р — «иp:» вместо «ир:»
    ("иp; Соколова",            "пр. Соколова"),
    ("иp: Соколова",            "пр. Соколова"),
    ("mp. Соколова",            "пр. Соколова"),

    # ── Цифры ────────────────────────────────────────────────────────────────
    ("]1234",                   "11234"),
    ("1234]",                   "12341"),
    ("1]23",                    "1123"),
    ("3+44000",                 "344000"),

    # ── Латинская B ──────────────────────────────────────────────────────────
    ("ВКЛЮЧИТЬ B реестр",       "ВКЛЮЧИТЬ В реестр"),
    ("лимитом B размере",       "лимитом в размере"),
    # B в начале строки без предшествующей кириллицы → строчная в (контекст неизвестен)
    ("B реестр требований",     "в реестр требований"),
    ("включить B реестр",       "включить в реестр"),

    # ── Email ─────────────────────────────────────────────────────────────────
    ("welcome@centrinvest пu n;", "welcome@centrinvest.ru"),
    ("welcome@centrinvest пи н;", "welcome@centrinvest.ru"),
    ("welcome@centrinvest ru",    "welcome@centrinvest.ru"),

    # ── Знак № ───────────────────────────────────────────────────────────────
    ("No 123",                  "№ 123"),
    ("Ng 123",                  "№ 123"),
    ("No. 123",                 "№ 123"),

    # ── К/с ──────────────────────────────────────────────────────────────────
    ("Klс",                     "К/с"),   # OCR: К/с → Klс (l заменяет /)
    ("K/с",                     "К/с"),   # OCR: К → K
    ("Klc",                     "К/с"),   # OCR: с → c

    # ── Двойной пробел ───────────────────────────────────────────────────────
    ("слово  слово",            "слово слово"),

    # ── Дефисный перенос ─────────────────────────────────────────────────────
    ("несосто- ятельным",       "несостоятельным"),

    # ── Госпошлина ───────────────────────────────────────────────────────────
    ("Госпошлина : 100",        "Госпошлина: 100"),
    ("100 руб:",                "100 руб."),

    # ── Восстановление тел./факс Центр-инвест ────────────────────────────────
    # EasyOCR стабильно пропускает эту строку; восстанавливаем по паттерну
    ("Россия, 344000, welcome@centrinvest.ru",
     "Россия, 344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru"),
    # Уже есть — не дублировать
    ("344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru",
     "344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru"),

    # ── Идентичность (не ломать то что уже правильно) ────────────────────────
    ("по делу № А53-3675/2025", "по делу № А53-3675/2025"),
    ("г. Ростов-на-Дону",       "г. Ростов-на-Дону"),
    ("ИНН: 6163011391",         "ИНН: 6163011391"),
])
def test_postprocess(inp: str, expected: str) -> None:
    assert postprocess(inp) == expected, (
        f"\n  вход:    {inp!r}\n  ожидал:  {expected!r}\n  получил: {postprocess(inp)!r}"
    )


def test_postprocess_empty_string() -> None:
    assert postprocess("") == ""


def test_postprocess_idempotent() -> None:
    """Второй вызов не должен менять уже обработанный текст."""
    samples = [
        "«Центр-инвест»",
        "г. Ростов-на-Дону, пр. Соколова, 62",
        "ИНН 6163011391, КПП 615250001",
    ]
    for s in samples:
        assert postprocess(postprocess(s)) == postprocess(s), (
            f"Не идемпотентно: {s!r}"
        )


# ─────────────────────────────────────────────────────────────────────────────
#  GEOMETRY
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.geometry import (
    bbox_h, bbox_mid_y, bbox_x0, bbox_x1, coplanar,
    detect_pdf_native, reading_order_key,
)


def test_bbox_h_normal():
    assert bbox_h(make_bbox(100, 80)) == pytest.approx(20.0)


def test_bbox_h_inverted():
    """Высота всегда положительная, независимо от порядка t/b."""
    assert bbox_h(make_bbox(80, 100)) == pytest.approx(20.0)


def test_bbox_h_zero():
    assert bbox_h(make_bbox(50, 50)) == pytest.approx(0.0)


def test_bbox_mid_y():
    assert bbox_mid_y(make_bbox(100, 80)) == pytest.approx(90.0)


def test_bbox_x0():
    assert bbox_x0(make_bbox(100, 80, l=30, r=200)) == pytest.approx(30.0)


def test_bbox_x1():
    assert bbox_x1(make_bbox(100, 80, l=30, r=200)) == pytest.approx(200.0)


def test_coplanar_overlapping():
    a = make_bbox(100, 60)
    b = make_bbox(80, 40)
    assert coplanar(a, b, tolerance=0)


def test_coplanar_not_overlapping():
    a = make_bbox(100, 80)   # spans 80–100
    b = make_bbox(60, 40)    # spans 40–60, gap=20
    assert not coplanar(a, b, tolerance=10)


def test_coplanar_within_tolerance():
    a = make_bbox(100, 80)
    b = make_bbox(60, 40)    # gap=20
    assert coplanar(a, b, tolerance=25)


def test_coplanar_none_args():
    assert coplanar(None, make_bbox(100, 80)) is False
    assert coplanar(make_bbox(100, 80), None) is False
    assert coplanar(None, None) is False


# ─────────────────────────────────────────────────────────────────────────────
#  WORD ORDER — структуры данных
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.word_order import (
    Word, VisualLine, TextBlock, reconstruct_blocks,
)


def test_word_properties():
    w = Word(text="тест", x0=10.0, y0=20.0, x1=110.0, y1=40.0)
    assert w.mid_y  == pytest.approx(30.0)
    assert w.mid_x  == pytest.approx(60.0)
    assert w.height == pytest.approx(20.0)


def test_word_zero_size():
    w = Word(text="x", x0=0.0, y0=5.0, x1=10.0, y1=5.0)
    assert w.height == pytest.approx(0.0)


def test_visual_line_empty():
    line = VisualLine()
    assert line.text    == ""
    assert line.mid_y   == pytest.approx(0.0)
    assert line.x0      == pytest.approx(0.0)
    assert line.x1      == pytest.approx(0.0)
    assert line.median_height == pytest.approx(10.0)   # default fallback


def test_visual_line_with_words():
    w1 = Word("первое", x0=0,  y0=10, x1=60,  y1=30)
    w2 = Word("второе", x0=70, y0=10, x1=130, y1=30)
    line = VisualLine(words=[w1, w2])
    assert line.text  == "первое второе"
    assert line.mid_y == pytest.approx(20.0)
    assert line.x0    == pytest.approx(0.0)
    assert line.x1    == pytest.approx(130.0)


def test_text_block_empty():
    block = TextBlock()
    assert block.line_count == 0
    assert block.text       == ""
    assert block.mid_y      == pytest.approx(0.0)


def test_text_block_with_lines():
    w = Word("слово", x0=0, y0=10, x1=100, y1=30)
    line = VisualLine(words=[w])
    block = TextBlock(lines=[line])
    assert block.line_count == 1
    assert block.text       == "слово"
    assert block.mid_y      == pytest.approx(20.0)
    assert block.font_height > 0


# ─────────────────────────────────────────────────────────────────────────────
#  WORD ORDER — reconstruct_blocks
# ─────────────────────────────────────────────────────────────────────────────

def test_reconstruct_empty():
    assert reconstruct_blocks([]) == []


def test_reconstruct_low_confidence_filtered():
    r = make_ocr(0, 10, 100, 30, "слово", conf=0.1)
    assert reconstruct_blocks([r], pdf_native=False) == []


def test_reconstruct_single_word():
    r = make_ocr(0, 10, 100, 30, "слово")
    blocks = reconstruct_blocks([r], pdf_native=False)
    assert len(blocks) == 1
    assert blocks[0].text == "слово"
    assert blocks[0].line_count == 1


def test_reconstruct_one_line_correct_x_order():
    """Слова на одной строке сортируются слева направо."""
    results = [
        make_ocr(200, 10, 300, 30, "третье"),
        make_ocr(0,   10, 100, 30, "первое"),
        make_ocr(100, 10, 200, 30, "второе"),
    ]
    blocks = reconstruct_blocks(results, pdf_native=False)
    assert len(blocks) == 1
    assert blocks[0].text == "первое второе третье"


def test_reconstruct_two_lines_correct_y_order():
    """Строки сортируются сверху вниз (screen coords: меньший y = выше)."""
    results = [
        make_ocr(0, 60, 100, 80, "вторая"),  # y=60–80 → ниже
        make_ocr(0, 10, 100, 30, "первая"),  # y=10–30 → выше
    ]
    blocks = reconstruct_blocks(results, pdf_native=False)
    combined = " ".join(b.text for b in blocks)
    assert combined.index("первая") < combined.index("вторая")


def test_reconstruct_two_paragraphs():
    """Большой Y-разрыв создаёт два отдельных параграфа."""
    results = [
        make_ocr(0, 10,  100, 30,  "параграф1"),
        make_ocr(0, 400, 100, 420, "параграф2"),  # разрыв ~370px >> межстрочный
    ]
    blocks = reconstruct_blocks(results, pdf_native=False)
    assert len(blocks) == 2


def test_reconstruct_whitespace_only_filtered():
    r = make_ocr(0, 10, 100, 30, "   ", conf=0.9)
    assert reconstruct_blocks([r], pdf_native=False) == []


def test_reconstruct_all_words_zero_height_filtered():
    """Слова с нулевой высотой (height=0) не влияют на median_h и фильтруются."""
    results = [make_ocr(0, 10, 100, 10, "слово")]  # y0=y1=10, height=0
    blocks = reconstruct_blocks(results, pdf_native=False)
    # Может вернуть [] (нет heights > 2) — не должен падать
    assert isinstance(blocks, list)


# ─────────────────────────────────────────────────────────────────────────────
#  CONVERTER — _reorder_by_word_order
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.converter import _reorder_by_word_order


def _fake_item(label: str, page: int, mid_y_from_top_pct: float,
               ph: float = 842.0, pw: float = 595.0, x0: float = 50.0):
    """
    Создаёт фейковый Docling-элемент.
    mid_y_from_top_pct: 0.0 = самый верх страницы, 1.0 = самый низ.
    В Docling (PDF coords, y=0 снизу):
      mid_y = ph * (1 - mid_y_from_top_pct)
    """
    mid_y = ph * (1.0 - mid_y_from_top_pct)
    t = mid_y + 10
    b = mid_y - 10
    bbox = types.SimpleNamespace(t=t, b=b, l=x0, r=x0 + 100)
    prov = types.SimpleNamespace(page_no=page, bbox=bbox)
    return types.SimpleNamespace(
        label=types.SimpleNamespace(value=label),
        prov=[prov],
        text=f"{label}@{mid_y_from_top_pct:.2f}",
    )


def _word_block(img_y_pct: float, img_h: float = 1000.0):
    """Фейковый TextBlock с mid_y в пиксельных экранных координатах."""
    return types.SimpleNamespace(mid_y=img_y_pct * img_h)


def test_reorder_picture_untouched():
    """Picture всегда остаётся на своей позиции."""
    items = [
        (_fake_item("picture",   1, 0.05), 0),
        (_fake_item("text",      1, 0.08), 0),   # шапка (< 15%)
        (_fake_item("paragraph", 1, 0.50), 0),   # тело A
        (_fake_item("paragraph", 1, 0.70), 0),   # тело B
    ]
    # word_order: B (70%) → block 0, A (50%) → block 1 (B выше чем A в px)
    blocks = [_word_block(0.30), _word_block(0.60)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)

    assert result[0][0].label.value == "picture",  "картинка сдвинулась!"
    assert result[1][0].label.value == "text",     "текст шапки сдвинулся!"


def test_reorder_header_zone_excluded():
    """Элементы в топ 15% страницы не переставляются."""
    items = [
        (_fake_item("paragraph", 1, 0.05), 0),   # < 15% → шапка, не трогать
        (_fake_item("paragraph", 1, 0.10), 0),   # < 15% → шапка, не трогать
        (_fake_item("paragraph", 1, 0.50), 0),   # тело
        (_fake_item("paragraph", 1, 0.70), 0),   # тело
    ]
    labels_before = [x[0].text for x in items]
    blocks = [_word_block(0.05), _word_block(0.10), _word_block(0.60), _word_block(0.40)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)

    # Элементы шапки не сдвинулись
    assert result[0][0].text == labels_before[0]
    assert result[1][0].text == labels_before[1]


def test_reorder_body_corrected():
    """Body-параграфы переставляются по word_order-порядку."""
    # 4 элемента: items близко друг к другу (y=0.20,0.22,0.24,0.80).
    # Блоки в другом порядке: block0=0.23, block1=0.19, block2=0.21, block3=0.78.
    # - item0(0.20) → block1(0.19) diff=0.01
    # - item1(0.22) → block2(0.21) diff=0.01
    # - item2(0.24) → block0(0.23) diff=0.01
    # - item3(0.80) → block3(0.78) diff=0.02
    # После сортировки по block_idx: item2, item0, item1, item3 — 3 перемещения.
    items = [
        (_fake_item("paragraph", 1, 0.20), 0),
        (_fake_item("paragraph", 1, 0.22), 0),
        (_fake_item("paragraph", 1, 0.24), 0),
        (_fake_item("paragraph", 1, 0.80), 0),
    ]
    blocks = [_word_block(0.23), _word_block(0.19), _word_block(0.21), _word_block(0.78)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)

    texts = [r[0].text for r in result]
    original = ["paragraph@0.20", "paragraph@0.22", "paragraph@0.24", "paragraph@0.80"]
    # Порядок должен измениться (word_order переставил 3 элемента)
    assert texts != original, f"Порядок не изменился: {texts}"
    # item3 (0.80) должен остаться на последнем месте
    assert "0.80" in texts[-1], f"item3 сдвинулся: {texts}"


def test_reorder_no_change_when_order_correct():
    """Если Docling-порядок уже верный — ничего не меняется."""
    items = [
        (_fake_item("paragraph", 1, 0.20), 0),
        (_fake_item("paragraph", 1, 0.50), 0),
        (_fake_item("paragraph", 1, 0.80), 0),
    ]
    original_texts = [x[0].text for x in items]

    # word_order подтверждает тот же порядок
    blocks = [_word_block(0.20), _word_block(0.50), _word_block(0.80)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)
    assert [r[0].text for r in result] == original_texts


def test_reorder_empty_items():
    result = _reorder_by_word_order([], {}, {})
    assert result == []


def test_reorder_no_word_blocks():
    """Без word_blocks_map — возвращаем оригинал без изменений."""
    items = [(_fake_item("paragraph", 1, 0.5), 0)]
    result = _reorder_by_word_order(items, {}, {1: (595.0, 842.0)})
    assert result == items


def test_reorder_single_body_item():
    """Один body-элемент — нечего переставлять."""
    items = [(_fake_item("paragraph", 1, 0.5), 0)]
    blocks = [_word_block(0.5)]
    result = _reorder_by_word_order(
        items, {1: (blocks, 1000.0)}, {1: (595.0, 842.0)}
    )
    assert result == items


def test_reorder_min_move_threshold():
    """Если порядок уже верный — ничего не переставляем (moved=0 < MIN_MOVE=2)."""
    items = [
        (_fake_item("paragraph", 1, 0.30), 0),
        (_fake_item("paragraph", 1, 0.50), 0),
        (_fake_item("paragraph", 1, 0.70), 0),
    ]
    original_texts = [x[0].text for x in items]

    # word_order меняет только первые два (один реально сдвигается)
    # но оба получают одинаковый block_idx → нет реального смещения
    blocks = [_word_block(0.30), _word_block(0.50), _word_block(0.70)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)
    assert [r[0].text for r in result] == original_texts


# ─────────────────────────────────────────────────────────────────────────────
#  CONVERTER — _bbox_top_bottom / _is_letterhead_stop
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.converter import _bbox_top_bottom, _is_letterhead_stop


def test_bbox_top_bottom_pdf_native():
    """PDF-native (y=0 снизу): t=750, b=700 → screen top=92, bottom=142."""
    bbox = make_bbox(t=750, b=700, l=10, r=100)
    top, bottom = _bbox_top_bottom(bbox, page_height=842.0, pdf_native=True)
    assert abs(top - 92.0) < 1.0    # 842 - 750 = 92
    assert abs(bottom - 142.0) < 1.0  # 842 - 700 = 142


def test_bbox_top_bottom_screen():
    """Screen (y=0 сверху): t=92, b=142 → top=92, bottom=142."""
    bbox = make_bbox(t=92, b=142, l=10, r=100)
    top, bottom = _bbox_top_bottom(bbox, page_height=842.0, pdf_native=False)
    assert abs(top - 92.0) < 1.0
    assert abs(bottom - 142.0) < 1.0


def test_is_letterhead_stop_matches():
    for text in ["КРЕДИТОР: ПАО КБ", "ДОЛЖНИК: Иванов", "Арбитражный суд",
                 "ЗАЯВЛЕНИЕ", "Госпошлина", "по делу № А53"]:
        assert _is_letterhead_stop(text), f"Должен быть стоп-маркером: {text!r}"


def test_is_letterhead_stop_no_match():
    for text in ["ИНН 6163011391", "тел./факс: (863)", "К/с 30101810",
                 "www.centrinvest.ru", "344000, г. Ростов-на-Дону"]:
        assert not _is_letterhead_stop(text), f"Не должен быть стоп-маркером: {text!r}"


# ─────────────────────────────────────────────────────────────────────────────
#  _fix_reading_order
# ─────────────────────────────────────────────────────────────────────────────

from docling_dev.converter import _fix_reading_order


def _sec_header(text: str, page: int = 1):
    """Fake section_header item."""
    bbox = make_bbox(t=800, b=780, l=50, r=400)
    prov = types.SimpleNamespace(page_no=page, bbox=bbox)
    return (types.SimpleNamespace(
        label=types.SimpleNamespace(value="section_header"),
        prov=[prov], text=text,
    ), 0)


def _list_item(text: str, page: int = 1):
    """Fake list_item."""
    bbox = make_bbox(t=500, b=480, l=50, r=400)
    prov = types.SimpleNamespace(page_no=page, bbox=bbox)
    return (types.SimpleNamespace(
        label=types.SimpleNamespace(value="list_item"),
        prov=[prov], text=text,
    ), 0)


def test_fix_order_caps_before_subtitle():
    """ALL-CAPS section_header выводится перед строчным subtitle."""
    items = [
        _sec_header("о включении в реестр"),      # строчный — должен стать вторым
        _sec_header("ЗАЯВЛЕНИЕ"),                  # ALL-CAPS — должен стать первым
        _sec_header("требований кредиторов"),      # строчный — остаётся третьим
    ]
    result = _fix_reading_order(items)
    assert result[0][0].text == "ЗАЯВЛЕНИЕ", f"Got: {[r[0].text for r in result]}"
    assert result[1][0].text == "о включении в реестр"


def test_fix_order_caps_no_change_when_correct():
    """Если ALL-CAPS уже первый — ничего не меняется."""
    items = [
        _sec_header("ЗАЯВЛЕНИЕ"),
        _sec_header("о включении в реестр"),
    ]
    result = _fix_reading_order(items)
    assert result[0][0].text == "ЗАЯВЛЕНИЕ"


def test_fix_order_numbered_list_sorted():
    """Numbered list_items сортируются по ведущей цифре."""
    items = [
        _list_item("2 Копия платежного поручения"),
        _list_item("1 Документы подтверждающие"),
        _list_item("3 Расчет задолженности"),
        _list_item("5 Копия доп. соглашения"),
        _list_item("4 Копия кредитного договора"),
    ]
    result = _fix_reading_order(items)
    texts = [r[0].text for r in result]
    assert texts[0].startswith("1 "), f"Got: {texts}"
    assert texts[1].startswith("2 ")
    assert texts[2].startswith("3 ")
    assert texts[3].startswith("4 ")
    assert texts[4].startswith("5 ")


def test_fix_order_numbered_already_sorted():
    """Если список уже в порядке — ничего не меняется."""
    items = [
        _list_item("1 Документы"),
        _list_item("2 Копия"),
        _list_item("3 Расчет"),
    ]
    result = _fix_reading_order(items)
    assert result[0][0].text == "1 Документы"


def test_fix_order_unnumbered_at_end():
    """Ненумерованные list_items остаются в конце после сортировки."""
    items = [
        _list_item("2 Копия поручения"),
        _list_item("1 Документы"),
        _list_item("3 Расчет"),
        _list_item("Копия доверенности"),   # без номера
    ]
    result = _fix_reading_order(items)
    texts = [r[0].text for r in result]
    assert texts[0].startswith("1 ")
    assert texts[-1] == "Копия доверенности", f"Got: {texts}"
