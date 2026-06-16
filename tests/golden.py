# -*- coding: utf-8 -*-
"""Golden-тесты вывода конвертера: снимок упорядоченного текста DOCX.

Использование:
  python tests/golden.py save   — снять/обновить эталоны из converted_word_docling/
  python tests/golden.py check  — сравнить текущий вывод с эталонами (diff)

Эталон — это упорядоченный текст документа (абзацы + ячейки таблиц) в том же
порядке, что и в DOCX. Так мы ловим РЕГРЕССИИ порядка/состава при переделке
reading-order, не перегоняя PDF заново.
"""
from __future__ import annotations
import sys, io, os, difflib, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

HERE      = os.path.dirname(os.path.abspath(__file__))
ROOT      = os.path.dirname(HERE)
OUT_DIR   = os.path.join(ROOT, "converted_word_docling")
GOLDEN_DIR = os.path.join(HERE, "golden")


def dump_text(path: str) -> str:
    """Упорядоченный текст: абзацы и ячейки таблиц по порядку DOCX."""
    d = Document(path)
    lines: list[str] = []
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            t = Paragraph(child, None).text.strip()
            if t:
                lines.append("P  | " + t)
        elif tag == "tbl":
            for row in Table(child, None).rows:
                cells = [c.text.strip().replace("\n", " / ") for c in row.cells]
                cells = [c for c in cells if c]
                if cells:
                    lines.append("TBL| " + "  ||  ".join(cells))
    return "\n".join(lines) + "\n"


def _golden_path(docx_path: str) -> str:
    base = os.path.splitext(os.path.basename(docx_path))[0]
    return os.path.join(GOLDEN_DIR, base + ".txt")


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    # необязательный 2-й аргумент — папка с DOCX (по умолч. converted_word_docling)
    src_dir = sys.argv[2] if len(sys.argv) > 2 else OUT_DIR
    if not os.path.isabs(src_dir):
        src_dir = os.path.join(ROOT, src_dir)
    os.makedirs(GOLDEN_DIR, exist_ok=True)
    # Исключаем временные lock-файлы Word (~$имя.docx), создаются при открытии в Word.
    docx_files = sorted(
        f for f in glob.glob(os.path.join(src_dir, "*.docx"))
        if not os.path.basename(f).startswith("~$")
    )
    if not docx_files:
        print("Нет DOCX в", OUT_DIR)
        return 1

    if mode == "save":
        for f in docx_files:
            text = dump_text(f)
            gp = _golden_path(f)
            with open(gp, "w", encoding="utf-8") as fh:
                fh.write(text)
            print(f"saved  {os.path.basename(gp)}  ({text.count(chr(10))} строк)")
        return 0

    # check
    rc = 0
    for f in docx_files:
        gp = _golden_path(f)
        cur = dump_text(f)
        if not os.path.exists(gp):
            print(f"[NEW]  {os.path.basename(f)} — нет эталона (запусти save)")
            rc = 1
            continue
        old = open(gp, encoding="utf-8").read()
        if old == cur:
            print(f"[ OK ] {os.path.basename(f)}")
        else:
            rc = 1
            print(f"[DIFF] {os.path.basename(f)}")
            diff = difflib.unified_diff(
                old.splitlines(), cur.splitlines(),
                fromfile="golden", tofile="current", lineterm="")
            for ln in diff:
                print("   " + ln)
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
