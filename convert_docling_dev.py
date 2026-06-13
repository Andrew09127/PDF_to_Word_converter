"""
convert_docling_dev.py
──────────────────────
Точка входа для Docling-конвертера (docling_dev/).

Использование:
  python convert_docling_dev.py                  # батч из pdf_to_convert/
  python convert_docling_dev.py file.pdf         # один файл
  python convert_docling_dev.py --no-word-order  # отключить автоопределение порядка

Папки по умолчанию:
  input:   ./pdf_to_convert/
  output:  ./converted_word_docling/
  backup:  ./pdf_backup/     (PDF перемещается сюда после конвертации)
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

# Консоль Windows часто в cp1251 — переключаем stdout на UTF-8, иначе логи с
# символами вроде «→»/«×» сыплют «Logging error» (на саму конвертацию не влияет).
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.FileHandler("docling_convert.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
# Подавляем DEBUG от внешних библиотек — только наш код
for _lib in ("docling", "httpx", "urllib3", "PIL", "torch", "easyocr",
             "transformers", "httpcore", "hpack"):
    logging.getLogger(_lib).setLevel(logging.WARNING)

_HERE = Path(__file__).parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

from docling_dev import DoclingBatchConverter, convert_pdf, build_converter


def _init_ocr_reader(langs: list[str]) -> object | None:
    """Инициализирует EasyOCR reader для word_order. None если недоступен."""
    try:
        import easyocr
        logging.info("Инициализация EasyOCR для word_order...")
        reader = easyocr.Reader(langs, gpu=False, verbose=False)
        logging.info("EasyOCR готов.")
        return reader
    except Exception as e:
        logging.warning("EasyOCR недоступен, word_order отключён: %s", e)
        return None


def _convert_single(pdf_path: Path, out_dir: Path, use_word_order: bool,
                    highlight: bool = True, llm: bool = False,
                    llm_model: str = "") -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    docx_path = out_dir / f"{pdf_path.stem}.docx"
    logging.info("Один файл: %s", pdf_path.name)

    converter  = build_converter()
    ocr_reader = _init_ocr_reader(["ru", "en"]) if use_word_order else None

    ok = convert_pdf(pdf_path, docx_path, converter,
                     ocr_reader=ocr_reader, use_word_order=use_word_order,
                     highlight=highlight, llm=llm, llm_model=llm_model)
    if ok:
        logging.info("Готово: %s", docx_path)
    else:
        logging.error("Ошибка конвертации: %s", pdf_path.name)
        sys.exit(1)


def _convert_batch(input_dir: Path, output_dir: Path, backup_dir: Path,
                   use_word_order: bool, highlight: bool = True,
                   llm: bool = False, llm_model: str = "") -> None:
    DoclingBatchConverter(
        input_folder=str(input_dir),
        output_folder=str(output_dir),
        backup_folder=str(backup_dir),
        use_word_order=use_word_order,
        highlight=highlight,
        llm=llm,
        llm_model=llm_model,
    ).process()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Конвертер отсканированных PDF → DOCX (Docling + EasyOCR)"
    )
    parser.add_argument(
        "pdf", nargs="?", type=Path,
        help="Путь к одному PDF (если не указан — батч-режим из pdf_to_convert/)"
    )
    parser.add_argument(
        "--input", "-i", type=Path, default=Path("pdf_to_convert"),
        help="Папка с PDF (по умолч. ./pdf_to_convert/)"
    )
    parser.add_argument(
        "--output", "-o", type=Path, default=Path("converted_word_docling"),
        help="Папка для DOCX (по умолч. ./converted_word_docling/)"
    )
    parser.add_argument(
        "--backup", "-b", type=Path, default=Path("pdf_backup"),
        help="Куда перемещать PDF после конвертации (по умолч. ./pdf_backup/)"
    )
    parser.add_argument(
        "--word-order", action="store_true",
        help="Включить EasyOCR-переупорядочивание блоков. По умолчанию ВЫКЛ: "
             "нативный порядок Docling точнее и не ломает вёрстку (штамп ЭП, "
             "реквизиты, колоночные блоки)."
    )
    parser.add_argument(
        "--no-highlight", action="store_true",
        help="Не подсвечивать жёлтым подозрительные (вероятно искажённые OCR) слова"
    )
    parser.add_argument(
        "--llm", action="store_true",
        help="Локальная LLM (llama-cpp-python + .gguf, в процессе) добивает "
             "оставшиеся подсвеченные слова. Полностью локально, без отдельных "
             "программ; если пакета/модели нет — шаг пропускается."
    )
    parser.add_argument(
        "--llm-model", default="",
        help="Путь к .gguf-модели для доочистки (по умолч. вендоренная в models/)"
    )
    args = parser.parse_args()

    use_word_order = args.word_order
    highlight      = not args.no_highlight

    if args.pdf:
        if not args.pdf.is_file():
            logging.error("Файл не найден: %s", args.pdf)
            sys.exit(1)
        _convert_single(args.pdf, args.output, use_word_order, highlight,
                        args.llm, args.llm_model)
    else:
        if not args.input.is_dir():
            logging.error(
                "Папка не найдена: %s\n"
                "Создайте её или укажите --input /путь/к/папке", args.input
            )
            sys.exit(1)
        _convert_batch(args.input, args.output, args.backup, use_word_order,
                       highlight, args.llm, args.llm_model)


if __name__ == "__main__":
    main()
