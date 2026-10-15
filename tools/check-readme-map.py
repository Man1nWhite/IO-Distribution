#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сверка карты процедур в README с Module1.bas.

Карта («Карта этапов → процедур:») не обязана перечислять все 63
процедуры — служебные могут отсутствовать. Но каждое имя, названное
в карте, обязано существовать в коде: переименование или удаление
процедуры без правки README роняет эту проверку в CI.

Код возврата: 0 — карта сходится, 1 — есть расхождение.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main():
    bas = (ROOT / "Module1.bas").read_bytes().decode("cp1251")
    code = set(re.findall(r"(?m)^(?:Private\s+|Public\s+)?"
                          r"(?:Sub|Function)\s+([^\s(]+)", bas))
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    m = re.search(r"Карта этапов → процедур:\n(.*?)\n\n", readme, re.S)
    if not m:
        print("FAIL README: блок «Карта этапов → процедур» не найден")
        return 1
    named = []
    for segment in m.group(1).split("·"):
        if "—" not in segment:
            continue
        for name in segment.split("—", 1)[1].split(","):
            name = name.strip().rstrip(".")
            if name:
                named.append(name)
    missing = [n for n in named if n not in code]
    for n in named:
        print(("ok   " if n in code else "FAIL ") + n)
    # Обратная сторона не проверяется (служебные процедуры вне карты
    # допустимы), но список печатается: так видно, что именно не описано,
    # и расхождение после переименования не прячется в «служебные».
    outside = sorted(code - set(named))
    print(f"\nвне карты (служебные, допустимо): {len(outside)}")
    for n in outside:
        print("     " + n)
    print(f"\nкарта README: {len(named)} имён, "
          f"отсутствует в коде: {len(missing)}")
    return 1 if missing else 0


if __name__ == "__main__":
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main())
