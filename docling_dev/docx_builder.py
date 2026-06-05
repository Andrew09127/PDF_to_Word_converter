"""
docx_builder.py
───────────────
Все функции для сборки Word-документа: стили, таблицы, шапка side-by-side,
блоки МЕТКА:содержимое, анализ страниц.
"""
from __future__ import annotations

import statistics
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .config import (
    BODY_PT, EMU_PER_INCH, FONT_NAME,
    LABEL_COL_RATIO, LABEL_LINE_RE,
    MARGIN_INCH, PAGE_H_INCH, PAGE_W_INCH,
)
from .geometry import bbox_h, bbox_mid_y, bbox_x0, coplanar
from .ocr_fixes import postprocess

_BODY_LABELS = frozenset({"paragraph", "text", "list_item"})


# ── Инициализация документа ───────────────────────────────────────────────────

def init_document() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width  = int(PAGE_W_INCH * EMU_PER_INCH)
    sec.page_height = int(PAGE_H_INCH * EMU_PER_INCH)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, attr, int(MARGIN_INCH * EMU_PER_INCH))

    normal           = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(BODY_PT)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    normal.paragraph_format.widow_control     = False
    normal.paragraph_format.keep_together     = False
    normal.paragraph_format.keep_with_next    = False

    for level in (1, 2, 3):
        h                = doc.styles[f"Heading {level}"]
        h.font.name      = FONT_NAME
        h.font.color.rgb = RGBColor(0, 0, 0)
        h.font.bold      = True
        h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        h.paragraph_format.keep_with_next    = False
        h.paragraph_format.keep_together     = False

    return doc


# ── Выравнивание ──────────────────────────────────────────────────────────────

def detect_alignment(bbox, page_width: float) -> WD_ALIGN_PARAGRAPH:
    if page_width <= 0:
        return WD_ALIGN_PARAGRAPH.LEFT
    x0      = float(getattr(bbox, "l", getattr(bbox, "x0", 0)))
    x1      = float(getattr(bbox, "r", getattr(bbox, "x1", page_width)))
    block_w = x1 - x0
    cx      = (x0 + x1) / 2
    ratio   = block_w / page_width
    if ratio > 0.75:
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    if abs(cx - page_width / 2) < page_width * 0.08 and ratio < 0.68:
        return WD_ALIGN_PARAGRAPH.CENTER
    if x0 > page_width * 0.12 and x1 > page_width * 0.72 and ratio < 0.82:
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


# ── Вспомогательные функции для таблиц ───────────────────────────────────────

def _set_cell_borders(cell) -> None:
    tc      = cell._tc
    tcPr    = tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)


def _make_borderless_table(doc: Document, n_cols: int) -> object:
    """Создаёт таблицу без видимых границ."""
    tbl       = doc.add_table(rows=1, cols=n_cols)
    tbl.style = "Normal Table"
    tbl_el    = tbl._tbl
    tbl_pr    = tbl_el.find(qn("w:tblPr"))
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tbl_pr)
    tbl_brd = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        tbl_brd.append(el)
    tbl_pr.append(tbl_brd)
    return tbl


def _set_cell_width(cell, width_inch: float) -> None:
    tc  = cell._tc
    tcp = tc.find(qn("w:tcPr"))
    if tcp is None:
        tcp = OxmlElement("w:tcPr")
        tc.insert(0, tcp)
    tcw = OxmlElement("w:tcW")
    tcw.set(qn("w:w"), str(int(width_inch * 1440)))
    tcw.set(qn("w:type"), "dxa")
    tcp.insert(0, tcw)
    tcb = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        tcb.append(el)
    tcp.append(tcb)


# ── Таблицы документа ─────────────────────────────────────────────────────────

def add_table_from_grid(doc: Document, grid: list) -> None:
    if not grid:
        return
    n_rows = len(grid)
    n_cols = max((len(row) for row in grid), default=0)
    if n_cols == 0:
        return
    tbl       = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.style = "Table Grid"
    for r_idx, row in enumerate(grid):
        for c_idx in range(n_cols):
            cell      = tbl.rows[r_idx].cells[c_idx]
            cell_data = row[c_idx] if c_idx < len(row) else None
            text      = postprocess((getattr(cell_data, "text", "") or "") if cell_data else "")
            _set_cell_borders(cell)
            for p in cell.paragraphs:
                p.clear()
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after  = Pt(0)
            run           = para.add_run(text)
            run.font.name = FONT_NAME
            run.font.size = Pt(9.0)
            run.bold      = (r_idx == 0)
    doc.add_paragraph()


def add_table_from_cells(doc: Document, table_item) -> None:
    data   = table_item.data
    n_rows = getattr(data, "num_rows", 0)
    n_cols = getattr(data, "num_cols", 0)
    if n_rows == 0 or n_cols == 0:
        return
    grid: list[list[str]] = [[""] * n_cols for _ in range(n_rows)]
    for cell in getattr(data, "table_cells", []):
        r = getattr(cell, "start_row_offset_idx", 0)
        c = getattr(cell, "start_col_offset_idx", 0)
        if 0 <= r < n_rows and 0 <= c < n_cols:
            grid[r][c] = postprocess(getattr(cell, "text", "") or "")
    tbl       = doc.add_table(rows=n_rows, cols=n_cols)
    tbl.style = "Table Grid"
    for r_idx, row in enumerate(grid):
        for c_idx, text in enumerate(row):
            cell = tbl.rows[r_idx].cells[c_idx]
            _set_cell_borders(cell)
            for p in cell.paragraphs:
                p.clear()
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after  = Pt(0)
            run           = para.add_run(text)
            run.font.name = FONT_NAME
            run.font.size = Pt(9.0)
            run.bold      = (r_idx == 0)
    doc.add_paragraph()


# ── Шапка: логотип + текст рядом ─────────────────────────────────────────────

def add_sidebyside(
    doc: Document,
    pil_img,
    img_w_inch: float,
    right_blocks: list[dict],
    text_zone_inch: float,
) -> None:
    """Безрамочная 2-колоночная таблица: логотип слева, текст реквизитов справа."""
    img_col_w = min(img_w_inch + 0.15, text_zone_inch * 0.48)
    txt_col_w = max(text_zone_inch - img_col_w, 1.0)

    tbl    = _make_borderless_table(doc, 2)
    cell_l = tbl.cell(0, 0)
    cell_r = tbl.cell(0, 1)
    _set_cell_width(cell_l, img_col_w)
    _set_cell_width(cell_r, txt_col_w)

    p_img = cell_l.paragraphs[0]
    p_img.alignment                     = WD_ALIGN_PARAGRAPH.LEFT
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after  = Pt(0)
    buf = BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    p_img.add_run().add_picture(buf, width=Inches(min(img_w_inch, img_col_w - 0.05)))

    first = True
    for blk in right_blocks:
        para = cell_r.paragraphs[0] if first else cell_r.add_paragraph()
        first = False
        para.alignment                     = blk.get("alignment", WD_ALIGN_PARAGRAPH.LEFT)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(1)
        run           = para.add_run(blk["text"])
        run.font.name = FONT_NAME
        run.font.size = Pt(blk.get("font_pt", BODY_PT))
        run.bold      = blk.get("bold", False)
        run.italic    = blk.get("italic", False)

    doc.add_paragraph()


# ── Блоки МЕТКА:содержимое ───────────────────────────────────────────────────

def add_header_row(
    doc: Document,
    cells: list[dict],
    text_zone_inch: float,
) -> None:
    """Render a first-page letterhead row with image/text cells in visual order."""
    if not cells:
        return

    tbl = _make_borderless_table(doc, len(cells))
    total_pt = sum(max(float(c.get("width_pt", 0.0)), 1.0) for c in cells)
    used_width = 0.0

    for idx, cell_data in enumerate(cells):
        cell = tbl.cell(0, idx)
        if idx == len(cells) - 1:
            col_w = max(text_zone_inch - used_width, 1.0)
        else:
            ratio = max(float(cell_data.get("width_pt", 0.0)), 1.0) / total_pt
            col_w = max(text_zone_inch * ratio, 1.0)
            used_width += col_w
        _set_cell_width(cell, col_w)

        para = cell.paragraphs[0]
        para.alignment = cell_data.get("alignment", WD_ALIGN_PARAGRAPH.LEFT)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)

        if cell_data.get("kind") == "image":
            pil_img = cell_data.get("image")
            if pil_img is None:
                continue
            buf = BytesIO()
            pil_img.save(buf, format="PNG")
            buf.seek(0)
            img_w = min(float(cell_data.get("image_width_inch", col_w)), col_w - 0.05)
            if img_w > 0:
                para.add_run().add_picture(buf, width=Inches(img_w))
            continue

        first = True
        for block in cell_data.get("blocks", []):
            p = para if first else cell.add_paragraph()
            first = False
            p.alignment = cell_data.get("alignment", WD_ALIGN_PARAGRAPH.LEFT)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(block.get("text", ""))
            run.font.name = FONT_NAME
            run.font.size = Pt(block.get("font_pt", 8.5))
            run.bold = block.get("bold", False)
            run.italic = block.get("italic", False)

    doc.add_paragraph()


def split_label_content(text: str) -> tuple[str, str] | None:
    """Возвращает (метка, содержимое) если текст начинается с ALL-CAPS метки."""
    m = LABEL_LINE_RE.match(text.strip())
    if not m:
        return None
    label = m.group(1).strip()
    alpha = [c for c in label if c.isalpha()]
    if not alpha or not all(c.isupper() for c in alpha) or len(alpha) < 3:
        return None
    return label, m.group(2).strip()


def add_label_content_table(
    doc: Document,
    label: str,
    content_items: list[dict],
    text_w_inch: float,
    space_before: float = 0.0,
    indent_inch: float = 0.0,
    col_ratio: float | None = None,
) -> None:
    """Безрамочная 2-колоночная строка: метка слева, содержимое справа.

    indent_inch — сдвиг таблицы от левого поля (для правоколоночных блоков).
    col_ratio   — доля ширины для колонки метки (None → LABEL_COL_RATIO).
    """
    ratio = col_ratio if col_ratio is not None else LABEL_COL_RATIO
    col_l = text_w_inch * ratio
    col_r = max(text_w_inch - col_l, 1.5)

    tbl    = _make_borderless_table(doc, 2)
    cl, cr = tbl.cell(0, 0), tbl.cell(0, 1)
    _set_cell_width(cl, col_l)
    _set_cell_width(cr, col_r)

    # Сдвиг таблицы от левого поля (tblInd, единица — twips = 1/1440 дюйма)
    if indent_inch > 0.01:
        tbl_el = tbl._tbl
        tbl_pr = tbl_el.find(qn("w:tblPr"))
        if tbl_pr is not None:
            tbl_ind = OxmlElement("w:tblInd")
            tbl_ind.set(qn("w:w"),    str(int(indent_inch * 1440)))
            tbl_ind.set(qn("w:type"), "dxa")
            tbl_pr.append(tbl_ind)

    p           = cl.paragraphs[0]
    p.alignment                     = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(0)
    r           = p.add_run(label)
    r.font.name = FONT_NAME
    r.font.size = Pt(BODY_PT)
    r.bold      = True

    first = True
    for blk in content_items:
        p = cr.paragraphs[0] if first else cr.add_paragraph()
        first = False
        p.alignment                     = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        r           = p.add_run(blk["text"])
        r.font.name = FONT_NAME
        r.font.size = Pt(blk.get("font_pt", BODY_PT))
        r.bold      = blk.get("bold", False)
        r.italic    = blk.get("italic", False)


# ── Анализ страниц ────────────────────────────────────────────────────────────

def analyse_pages(
    items: list,
) -> tuple[dict[int, float], dict[int, float]]:
    """
    Однопроходный анализ всех элементов:
      page_medians  — медианная высота bbox body-text по каждой странице
      page_left_min — 10-й перцентиль левого края (эффективное поле)
    """
    heights_by_page: dict[int, list[float]] = {}
    lefts_by_page:   dict[int, list[float]] = {}

    for item, _ in items:
        raw = getattr(item, "label", None)
        if raw is None:
            continue
        label = (raw.value if hasattr(raw, "value") else str(raw)).lower()
        if label not in _BODY_LABELS:
            continue
        prov_list = getattr(item, "prov", None) or []
        if not prov_list:
            continue
        bbox = getattr(prov_list[0], "bbox", None)
        if bbox is None:
            continue
        h = bbox_h(bbox)
        l = float(getattr(bbox, "l", 0))
        if h < 2:
            continue
        page = int(getattr(prov_list[0], "page_no", 1))
        heights_by_page.setdefault(page, []).append(h)
        if l >= 0:
            lefts_by_page.setdefault(page, []).append(l)

    page_medians: dict[int, float] = {
        p: statistics.median(hs) for p, hs in heights_by_page.items() if hs
    }
    page_left_min: dict[int, float] = {}
    for p, ls in lefts_by_page.items():
        if ls:
            ls_s = sorted(ls)
            page_left_min[p] = ls_s[max(0, len(ls_s) // 10)]

    return page_medians, page_left_min
