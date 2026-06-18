"""
reassemble_wheels.py
────────────────────
Склеивает крупные wheel'ы из кусков `<имя>.whl.partNNN` обратно в `<имя>.whl`.
Запускается из install_offline.bat перед установкой пакетов. Идемпотентно:
если собранный .whl уже есть — пропускает.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Консоль Windows часто в cp1251 — переключаем вывод на UTF-8.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
except (AttributeError, ValueError):
    pass

WHEELS = Path(__file__).resolve().parent.parent / "offline_kit" / "wheels"


def main() -> int:
    if not WHEELS.is_dir():
        print(f"Нет папки {WHEELS}", file=sys.stderr)
        return 1

    # Каждый набор кусков опознаём по .part000
    for first in sorted(WHEELS.glob("*.whl.part000")):
        base = first.with_name(first.name[: -len(".part000")])  # <имя>.whl
        if base.exists():
            continue   # уже собран
        parts = sorted(WHEELS.glob(base.name + ".part*"))
        print(f"Склеиваю {base.name} из {len(parts)} кусков…")
        with base.open("wb") as out:
            for part in parts:
                out.write(part.read_bytes())
        print(f"  -> {base.stat().st_size/1048576:.0f} МБ")

    print("Готово.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
