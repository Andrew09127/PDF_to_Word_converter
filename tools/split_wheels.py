"""
split_wheels.py
───────────────
Разбивает крупные wheel'ы (>95 МБ) на куски `<имя>.whl.partNNN` <90 МБ ради
лимита GitHub (100 MiB на файл). Куски коммитятся в git, оригинал удаляется.
При установке куски склеиваются обратно (tools/reassemble_wheels.py) — тот же
приём, что и для gguf-модели (см. llm_correct._resolve_model).

Запускается автоматически из build_offline_kit.bat после сборки wheels.
Идемпотентно: уже разбитые файлы пропускает.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Консоль Windows часто в cp1251 — переключаем вывод на UTF-8.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
except (AttributeError, ValueError):
    pass

THRESHOLD = 95 * 1024 * 1024   # делим всё, что крупнее 95 МБ
CHUNK = 90 * 1024 * 1024       # размер куска (< 100 MiB лимита GitHub)

WHEELS = Path(__file__).resolve().parent.parent / "offline_kit" / "wheels"


def main() -> int:
    if not WHEELS.is_dir():
        print(f"Нет папки {WHEELS}", file=sys.stderr)
        return 1

    for whl in sorted(WHEELS.glob("*.whl")):
        size = whl.stat().st_size
        if size <= THRESHOLD:
            continue
        nparts = (size + CHUNK - 1) // CHUNK
        print(f"Разбиваю {whl.name} ({size/1048576:.0f} МБ) на {nparts} кусков…")
        with whl.open("rb") as src:
            idx = 0
            while True:
                buf = src.read(CHUNK)
                if not buf:
                    break
                part = whl.with_name(f"{whl.name}.part{idx:03d}")
                part.write_bytes(buf)
                idx += 1
        whl.unlink()   # оригинал не коммитим — только куски
        print(f"  -> {idx} кусков, оригинал удалён")

    print("Готово.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
