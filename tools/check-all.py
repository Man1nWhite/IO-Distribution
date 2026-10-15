#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Все проверки репозитория одной командой — те же шаги, что в CI.

Запуск из корня или из любого каталога, на любой ОС:

    python3 tools/check-all.py          (Windows: py tools\\check-all.py)
    python3 tools/check-all.py --list   только перечислить шаги

Шаги — STEPS ниже. Их совпадение с шагами CI (.github/workflows/vba-lint.yml,
в том же порядке) сверяет tools/check-sync.py через `--list`: новый инструмент
обязан попасть и в CI, и сюда, иначе локальный прогон его молча пропустит.
Каждый шаг запускается тем же интерпретатором, что и этот скрипт
(sys.executable), поэтому `py`, `python` и venv работают одинаково.
Как и в CI, первый упавший шаг останавливает прогон.

`./tools/check-all.sh` — обёртка над этим файлом для привычной команды.

Код возврата: 0 — всё зелено, иначе код первой упавшей проверки.
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# (скрипт, заголовок) — порядок как в CI.
STEPS = [
    ("tools/vba-lint.py", "vba-lint: статика Module1.bas"),
    ("tools/parse-spec.py", "parse-spec: разбор чисел из текста"),
    ("tools/plan-spec.py", "plan-spec: укладчик реек"),
    ("tools/check-readme-map.py", "check-readme-map: карта процедур в README"),
    ("tools/check-example.py", "check-example: Example.xlsm против README и Service"),
    ("tools/check-sync.py", "check-sync: зеркала, константы и обещания README"),
]


def main(argv):
    args = argv[1:]
    if args == ["--list"]:
        # Формат как в `run:` шагов CI — его сверяет check-sync.py.
        for script, _title in STEPS:
            print("python3 " + script)
        return 0
    if args:
        print(__doc__)
        return 2

    total = len(STEPS)
    for n, (script, title) in enumerate(STEPS, 1):
        if n > 1:
            print()
        # flush: заголовок должен попасть в pipe раньше вывода дочернего шага.
        print("== %d/%d %s ==" % (n, total, title), flush=True)
        rc = subprocess.call([sys.executable, str(ROOT / script)], cwd=str(ROOT))
        if rc != 0:
            print("\nFAIL %s: код %d — остальные шаги не запускались" % (script, rc))
            return rc

    print("\nвсе проверки пройдены")
    return 0


if __name__ == "__main__":
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv))
