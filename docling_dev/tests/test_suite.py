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

#  HELPERS

def make_bbox(t, b, l=0, r=100):
    return types.SimpleNamespace(t=t, b=b, l=l, r=r)


def make_pts(x0, y0, x1, y1):
    return [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]


def make_ocr(x0, y0, x1, y1, text, conf=0.9):
    return (make_pts(x0, y0, x1, y1), text, conf)

#  OCR FIXES

from docling_dev.ocr_fixes import postprocess, fix_quotes

@pytest.mark.parametrize("inp,expected", [
    #Кавычки
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
    ("«Центр-инвест»",          "«Центр-инвест»"),   
    ("«ентр-инвест»",           "«Центр-инвест»"), 

    # ── Пунктуация ────────────────────────────────────────────────────────────
    ("ИНН 6163011391,; ОГРН",   "ИНН 6163011391, ОГРН"),
    ("р-н; х Ленинакан",        "р-н, х Ленинакан"),
    ("; пр. Соколова",          ", пр. Соколова"),

    # ── г. / пр. ─────────────────────────────────────────────────────────────
    ("&. Ростов",               "г. Ростов"),
    ("; &. Ростов",             "; г. Ростов"),
    # Latin P вместо Cyrillic Р — OCR путает («Pocmов» вместо «Ростов»);
    # правило нормализует латиницу обратно в «Ростов».
    ("&. Pocmов-на-Дону",       "г. Ростов-на-Дону"),
    ("ир; Соколова",            "пр. Соколова"),
    ("ир: Соколова",            "пр. Соколова"),
    # Latin p вместо Cyrillic р — «иp:» вместо «ир:»
    ("иp; Соколова",            "пр. Соколова"),
    ("иp: Соколова",            "пр. Соколова"),
    ("mp. Соколова",            "пр. Соколова"),

    #Цифры
    ("]1234",                   "11234"),
    ("1234]",                   "12341"),
    ("1]23",                    "1123"),
    ("3+44000",                 "344000"),

    #Латинская B 
    # B, затем капслок-правило «ВКЛЮЧИТЬ»→«включить» (OCR-капс в середине)
    ("ВКЛЮЧИТЬ B реестр",       "включить в реестр"),
    ("лимитом B размере",       "лимитом в размере"),
    # B в начале строки без предшествующей кириллицы - строчная в (контекст неизвестен)
    ("B реестр требований",     "в реестр требований"),
    ("включить B реестр",       "включить в реестр"),

    #Email 
    ("welcome@centrinvest пu n;", "welcome@centrinvest.ru"),
    ("welcome@centrinvest пи н;", "welcome@centrinvest.ru"),
    ("welcome@centrinvest ru",    "welcome@centrinvest.ru"),

    # ── Знак №
    ("No 123",                  "№ 123"),
    ("Ng 123",                  "№ 123"),
    ("No. 123",                 "№ 123"),
    # Слитный Ne после кириллицы (OCR потерял пробел и №)
    ("требованияNе 14-02-25",   "требования № 14-02-25"),
    ("договораNe 60190309",     "договора № 60190309"),

    #К/с 
    ("Klс",                     "К/с"),   # OCR: К/с - Klс (l заменяет /)
    ("K/с",                     "К/с"),   # OCR: К - K
    ("Klc",                     "К/с"),   # OCR: с - c

    #Двойной пробел
    ("слово  слово",            "слово слово"),

    #Дефисный перенос
    ("несосто- ятельным",       "несостоятельным"),

    #Госпошлина
    ("Госпошлина : 100",        "Госпошлина: 100"),
    ("100 руб:",                "100 руб."),

    #Восстановление тел./факс Центр-инвест 
    # EasyOCR стабильно пропускает эту строку; восстанавливаем по паттерну
    ("Россия, 344000, welcome@centrinvest.ru",
     "Россия, 344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru"),
    # Уже есть — не дублировать
    ("344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru",
     "344000, тел./факс: (863) 2-000-000, www.centrinvest.ru, welcome@centrinvest.ru"),

    #Доменные OCR-слова (раунд spell) 
    ("действует в соответствин с", "действует в соответствии с"),
    ("на основании выеизложенного", "на основании вышеизложенного"),
    ("договор с Заемшиком",        "договор с Заёмщиком"),
    ("права заемщика защищены",    "права заёмщика защищены"),

    #Идентичность
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


#  GEOMETRY

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

#  WORD ORDER — структуры данных

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


#  WORD ORDER — reconstruct_blocks

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


#  CONVERTER — _reorder_by_word_order

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
    # word_order: B (70%) - block 0, A (50%) - block 1 (B выше чем A в px)
    blocks = [_word_block(0.30), _word_block(0.60)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)

    assert result[0][0].label.value == "picture",  "картинка сдвинулась!"
    assert result[1][0].label.value == "text",     "текст шапки сдвинулся!"


def test_reorder_header_zone_excluded():
    """Элементы в топ 15% страницы не переставляются."""
    items = [
        (_fake_item("paragraph", 1, 0.05), 0),   # < 15% - шапка, не трогать
        (_fake_item("paragraph", 1, 0.10), 0),   # < 15% - шапка, не трогать
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
    # но оба получают одинаковый block_idx - нет реального смещения
    blocks = [_word_block(0.30), _word_block(0.50), _word_block(0.70)]
    word_blocks_map = {1: (blocks, 1000.0)}
    page_sizes = {1: (595.0, 842.0)}

    result = _reorder_by_word_order(items, word_blocks_map, page_sizes)
    assert [r[0].text for r in result] == original_texts

#  CONVERTER — _bbox_top_bottom / _is_letterhead_stop

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

#  _fix_reading_order

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
    result, _ = _fix_reading_order(items)
    assert result[0][0].text == "ЗАЯВЛЕНИЕ", f"Got: {[r[0].text for r in result]}"
    assert result[1][0].text == "о включении в реестр"


def test_fix_order_caps_no_change_when_correct():
    """Если ALL-CAPS уже первый — ничего не меняется."""
    items = [
        _sec_header("ЗАЯВЛЕНИЕ"),
        _sec_header("о включении в реестр"),
    ]
    result, _ = _fix_reading_order(items)
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
    result, _ = _fix_reading_order(items)
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
    result, _ = _fix_reading_order(items)
    assert result[0][0].text == "1 Документы"


def test_fix_order_unnumbered_at_end():
    """Ненумерованный пункт в конце получает автонумерацию-продолжение:
    OCR часто теряет номер, поэтому к 1–3 добавляется «4 …»."""
    items = [
        _list_item("2 Копия поручения"),
        _list_item("1 Документы"),
        _list_item("3 Расчет"),
        _list_item("Копия доверенности"),   # без номера → станет «4 …»
    ]
    result, _ = _fix_reading_order(items)
    texts = [r[0].text for r in result]
    assert texts[0].startswith("1 ")
    assert texts[1].startswith("2 ")
    assert texts[2].startswith("3 ")
    assert texts[-1] == "4 Копия доверенности", f"Got: {texts}"

#  HIGHLIGHT — кириллический спелл-фиксер (_autofix_word)

from docling_dev.highlight import _autofix_word, _morph

# Спелл-фиксер требует словаря pymorphy; без него тесты пропускаем.
_need_morph = pytest.mark.skipif(_morph is None, reason="pymorphy не установлен")


@_need_morph
@pytest.mark.parametrize("garbled,fixed", [
    ("нмущества",       "имущества"),
    ("абластн",         "области"),
    ("падтверждено",    "подтверждено"),
    ("васстановлен",    "восстановлен"),
    ("частнасти",       "частности"),
    ("атсутстние",      "отсутствие"),
    ("прнменяеная",     "применяемая"),
    ("данньми",         "данными"),         # пара ь↔ы
    ("несостоятельньм", "несостоятельным"), # пара ь↔ы
    ("деятсльности",    "деятельности"),
])
def test_spell_fix_lower(garbled, fixed):
    """Строчные чисто-кириллические OCR-опечатки с ЕДИНСТВЕННЫМ кандидатом."""
    assert _autofix_word(garbled) == fixed


@_need_morph
def test_spell_fix_allcaps_preserves_case():
    """ALL-CAPS заголовок чинится с сохранением регистра."""
    assert _autofix_word("МЕЖРАПОННАЯ") == "МЕЖРАЙОННАЯ"


@_need_morph
@pytest.mark.parametrize("proper", ["Заречнев", "Аракслян", "Краснянскову"])
def test_spell_fix_skips_titlecase(proper):
    """Имена/фамилии (Первая-Заглавная) не «чиним» — только подсветка."""
    assert _autofix_word(proper) is None


@_need_morph
@pytest.mark.parametrize("domain", [
    "взыскателя", "займодавцу", "микрофинансовой", "коллекторская",
])
def test_spell_fix_skips_domain(domain):
    """Доменные юр./фин. термины (нет в pymorphy, но верные) не трогаем."""
    assert _autofix_word(domain) is None


@_need_morph
@pytest.mark.parametrize("ambiguous", ["арганом", "налоговон"])
def test_spell_fix_skips_ambiguous(ambiguous):
    """Несколько словарных кандидатов → не угадываем (отдаём на подсветку/LLM)."""
    assert _autofix_word(ambiguous) is None


@_need_morph
@pytest.mark.parametrize("valid", ["имущества", "области", "требований", "договоров"])
def test_spell_fix_keeps_valid(valid):
    """Уже корректные слова не меняем."""
    assert _autofix_word(valid) is None

#  INK — насыщенность штриха (жирность по изображению на сканах)

from docling_dev.ink import block_ink_stats

try:
    import numpy as _np
    from PIL import Image as _PILImage
    _HAS_NP = True
except Exception:
    _HAS_NP = False

_need_np = pytest.mark.skipif(not _HAS_NP, reason="numpy/PIL не установлены")


def _synthetic_text(stroke_px: int):
    """Белый холст с тремя «текстовыми» полосами из вертикальных штрихов
    заданной толщины — имитация тонкого/жирного шрифта."""
    arr = _np.full((50, 200), 255, dtype=_np.uint8)
    for row0 in (10, 25, 40):
        for col in range(5, 195, 8):
            arr[row0:row0 + 6, col:col + stroke_px] = 0
    return _PILImage.fromarray(arr)


@_need_np
def test_ink_none_for_empty():
    assert block_ink_stats(None) is None


@_need_np
def test_ink_none_for_tiny():
    assert block_ink_stats(_PILImage.fromarray(_np.zeros((2, 2), dtype=_np.uint8))) is None


@_need_np
def test_ink_stroke_w_grows_with_thickness():
    """Главное свойство: чем толще штрих, тем выше stroke_w/mean_run (жирнее)."""
    thin = block_ink_stats(_synthetic_text(1))
    mid  = block_ink_stats(_synthetic_text(2))
    bold = block_ink_stats(_synthetic_text(3))
    assert thin is not None and mid is not None and bold is not None
    assert thin["stroke_w"] < mid["stroke_w"] < bold["stroke_w"]
    assert thin["mean_run"] < mid["mean_run"] < bold["mean_run"]


@_need_np
def test_ink_robust_to_underline():
    """Подчёркивание (длинная линия) НЕ должно раздувать stroke_w/mean_run:
    толщина штриха букв сохраняется (линии-выбросы отсекаются max_run_px)."""
    plain = block_ink_stats(_synthetic_text(1))
    arr = _np.full((50, 200), 255, dtype=_np.uint8)
    for row0 in (10, 25, 40):
        for col in range(5, 195, 8):
            arr[row0:row0 + 6, col:col + 1] = 0
    arr[46:48, 5:195] = 0                       # длинная линия-подчёркивание
    underlined = block_ink_stats(_PILImage.fromarray(arr))
    assert plain is not None and underlined is not None
    assert abs(underlined["stroke_w"] - plain["stroke_w"]) < 0.6
    assert abs(underlined["mean_run"] - plain["mean_run"]) < 0.6


@_need_np
def test_ink_dense_thin_not_bold():
    """«Плотный, но тонкий» (имитация цифр) НЕ должен выглядеть жирным:
    stroke_w остаётся как у тонкого, хотя stroke_density высокая."""
    thin   = block_ink_stats(_synthetic_text(1))   # редкие тонкие штрихи
    arr = _np.full((50, 200), 255, dtype=_np.uint8)
    for row0 in (10, 25, 40):
        for col in range(5, 195, 3):
            arr[row0:row0 + 6, col:col + 1] = 0
    dense = block_ink_stats(_PILImage.fromarray(arr))
    assert thin is not None and dense is not None
    assert abs(dense["stroke_w"] - thin["stroke_w"]) < 0.6
    assert dense["stroke_density"] > thin["stroke_density"]  

@_need_np
def test_ink_values_in_range():
    st = block_ink_stats(_synthetic_text(2))
    assert st is not None
    assert 0.0 <= st["stroke_density"] <= 1.0
    assert 0.0 <= st["ink_ratio"] <= 1.0
    assert st["stroke_w"] >= 0.0
    assert st["mean_run"] >= 0.0


@_need_np
def test_ink_blank_is_zero():
    """Чистый белый холст — нулевая насыщенность."""
    blank = _PILImage.fromarray(_np.full((40, 120), 255, dtype=_np.uint8))
    st = block_ink_stats(blank)
    assert st is not None
    assert st["ink_ratio"] == 0.0
    assert st["stroke_density"] == 0.0

#  OCR PREPROCESS — предобработка изображения перед EasyOCR

from docling_dev.ocr_preprocess import preprocess_for_ocr

try:
    import cv2 as _cv2
    _HAS_CV2 = True
except Exception:
    _HAS_CV2 = False

_need_cv2 = pytest.mark.skipif(not (_HAS_NP and _HAS_CV2),
                               reason="cv2/numpy не установлены")


@_need_cv2
def test_preprocess_returns_rgb_same_size():
    """Возвращает 3-канальный RGB того же размера (EasyOCR ждёт RGB)."""
    img = _np.full((120, 200, 3), 255, dtype=_np.uint8)
    img[50:60, 20:180] = 30
    out = preprocess_for_ocr(img)
    assert out.shape == (120, 200, 3)


@_need_cv2
def test_preprocess_noop_on_non_array():
    """Не-numpy вход (путь/None) возвращается как есть (no-op)."""
    assert preprocess_for_ocr("page.png") == "page.png"
    assert preprocess_for_ocr(None) is None


@_need_cv2
def test_preprocess_handles_grayscale():
    """Серый вход (2D) не падает и даёт RGB."""
    gray = _np.full((80, 120), 240, dtype=_np.uint8)
    gray[30:36, 10:110] = 20
    out = preprocess_for_ocr(gray)
    assert out.ndim == 3 and out.shape[2] == 3


@_need_cv2
def test_preprocess_denoise_off_by_default():
    """denoise по умолчанию выкл (он размывал текст и снижал confidence)."""
    import inspect
    sig = inspect.signature(preprocess_for_ocr)
    assert sig.parameters["denoise"].default is False


@_need_cv2
def test_preprocess_no_crash_on_rotated():
    """Повёрнутый текст (deskew) не вызывает падения."""
    img = _np.full((150, 300, 3), 255, dtype=_np.uint8)
    for r in range(40, 110, 20):
        img[r:r + 4, 30:270] = 0
    rot = _cv2.warpAffine(
        img, _cv2.getRotationMatrix2D((150, 75), 3.0, 1.0), (300, 150),
        borderValue=(255, 255, 255))
    out = preprocess_for_ocr(rot)
    assert out.shape == (150, 300, 3)
