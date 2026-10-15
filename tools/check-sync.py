#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сверка зеркал: Module1.bas ⇄ tools/*-spec.py ⇄ README.md ⇄ CI.

Правило репозитория — «меняешь VBA, меняешь зеркала»: `tools/parse-spec.py`,
`tools/plan-spec.py` и README переносятся из `Module1.bas` **вручную**.
Этот скрипт проверяет то, что проверяемо статически, — числа и таблицы,
продублированные в нескольких файлах:

  1. `PLAN_EPS` / `PLAN_NODE_LIMIT` / `PLAN_SEARCH_BLOCK_LIMIT` / `MAX_LONG`
     (Module1.bas) ↔ `EPS` / `NODE_LIMIT` / `SEARCH_BLOCK_LIMIT`
     (tools/plan-spec.py) и `MAX_LONG` (tools/parse-spec.py);
  2. `defRows` (Module1.bas) ↔ набор строк из таблицы «Строки» в README;
  3. `defLensFallback` (Module1.bas) ↔ строка «Запасные ширины» в README;
  4. команды раздела «Проверки» README ↔ шаги CI: новый инструмент обязан
     попасть в оба файла, иначе он тихо не работает;
  5. счётчики, обещанные README («N тестовых вектора», «все N имён»), ↔
     фактические результаты parse-spec.py и check-readme-map.py.

Что НЕ проверяется: сами правила разбора и укладки (они переносятся по смыслу)
— только их параметры и обещанные числа.

Код возврата: 0 — всё сошлось, 1 — есть расхождение.
"""

import os
import re
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BAS = ROOT / "Module1.bas"
README = ROOT / "README.md"
WORKFLOW = ROOT / ".github" / "workflows" / "vba-lint.yml"
PLAN_SPEC = ROOT / "tools" / "plan-spec.py"
PARSE_SPEC = ROOT / "tools" / "parse-spec.py"

FAILS = []
CHECKS = 0


def check(label, got, want):
    """got — из кода, want — из доков/зеркала (в выводе это подписано)."""
    global CHECKS
    CHECKS += 1
    ok = got == want
    print(("ok   " if ok else "FAIL ") + label)
    if not ok:
        print("       код: %s" % (got,))
        print("       доки/зеркало: %s" % (want,))
        FAILS.append(label)


def bas_text():
    return BAS.read_bytes().decode("cp1251")


def bas_const(name):
    """Числовая константа из Module1.bas — строкой, как записана."""
    m = re.search(r"(?m)^Private Const %s As \w+ = ([0-9][0-9.]*)#?" %
                  re.escape(name), bas_text())
    if not m:
        return None
    return m.group(1)


def bas_array(name):
    """Array(...) из Module1.bas (переносы строк через «_») — списком чисел."""
    m = re.search(r"(?m)^\s*%s = Array\((.*?)\)" % re.escape(name),
                  bas_text(), re.S)
    if not m:
        return None
    return [int(x) for x in re.findall(r"\d+", m.group(1))]


def py_const(path, name):
    m = re.search(r"(?m)^%s\s*=\s*([0-9][0-9.]*)" % re.escape(name),
                  path.read_text(encoding="utf-8"))
    if not m:
        return None
    return m.group(1)


def readme():
    return README.read_text(encoding="utf-8")


def check_numbers():
    """Module1.bas ⇄ зеркала: параметры, продублированные вручную."""
    pairs = [
        ("PLAN_EPS", PLAN_SPEC, "EPS"),
        ("PLAN_NODE_LIMIT", PLAN_SPEC, "NODE_LIMIT"),
        ("PLAN_SEARCH_BLOCK_LIMIT", PLAN_SPEC, "SEARCH_BLOCK_LIMIT"),
        ("MAX_LONG", PARSE_SPEC, "MAX_LONG"),
    ]
    for bas_name, mirror, mirror_name in pairs:
        got = bas_const(bas_name)
        want = py_const(mirror, mirror_name)
        label = ("%s (Module1.bas) == %s (%s)" %
                 (bas_name, mirror_name, mirror.relative_to(ROOT)))
        if got is None:
            check(label, "<константа не найдена>", want)
        elif want is None:
            check(label, got, "<константа не найдена>")
        else:
            check(label, Decimal(got), Decimal(want))


def readme_rows():
    """Набор строк листа «1» из таблицы «| Строки | Что | …»."""
    lines = readme().splitlines()
    start = None
    for i, line in enumerate(lines):
        if line.startswith("| Строки |"):
            start = i
            break
    if start is None:
        return None
    rows = set()
    for line in lines[start + 2:]:
        if not line.startswith("|"):
            break
        cell = line.split("|")[1].strip()
        for part in re.split(r"[/,]", cell):
            part = part.strip()
            m = re.match(r"^(\d+)\s*(?:[–-]\s*(\d+))?$", part)
            if not m:
                return None
            lo, hi = int(m.group(1)), int(m.group(2) or m.group(1))
            rows.update(range(lo, hi + 1))
    return rows


def readme_fallback():
    """Таблица «строка → мм» из строки про defLensFallback: {row: width}."""
    m = re.search(r"\(`defLensFallback`\), строка → мм:(.*?)(?:\n\n|\Z)",
                  readme(), re.S)
    if not m:
        return None
    table = {}
    for lo, hi, width in re.findall(r"(\d+)(?:\.\.(\d+))?=(\d+)", m.group(1)):
        for row in range(int(lo), int(hi or lo) + 1):
            table[row] = int(width)
    return table


def check_rows_and_widths():
    rows = bas_array("defRows")
    fallback = bas_array("defLensFallback")
    doc_rows = readme_rows()
    doc_fallback = readme_fallback()

    if rows is None or doc_rows is None:
        check("defRows (Module1.bas) == таблица «Строки» в README",
              rows, doc_rows)
    else:
        check("defRows (Module1.bas) == таблица «Строки» в README",
              sorted(rows), sorted(doc_rows))

    if fallback is None or rows is None or doc_fallback is None:
        check("defLensFallback (Module1.bas) == «Запасные ширины» в README",
              fallback, doc_fallback)
        return
    got = dict(zip(rows, fallback))
    check("defLensFallback (Module1.bas) == «Запасные ширины» в README",
          got, doc_fallback)


def readme_commands():
    m = re.search(r"## Проверки\n\n```sh\n(.*?)```", readme(), re.S)
    if not m:
        return None
    return [line.strip() for line in m.group(1).splitlines() if line.strip()]


def ci_commands():
    text = WORKFLOW.read_text(encoding="utf-8")
    return re.findall(r"(?m)^\s*run:\s*(python3 tools/[\w.-]+\.py)", text)


def check_ci_matches_readme():
    doc = readme_commands()
    ci = ci_commands()
    if doc is None:
        check("раздел «Проверки» README == шаги CI", None, ci)
        return
    # got — всегда код (шаги CI), want — доки (README).
    check("раздел «Проверки» README == шаги CI",
          sorted(ci), sorted(doc))


def tool_output(args):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run([sys.executable] + args, cwd=ROOT, env=env,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace")
    return (proc.stdout or "") + (proc.stderr or "")


def check_readme_counts():
    """Числа, обещанные README, против фактических результатов инструментов."""
    doc = readme()

    m = re.search(r"(\d+)\s+тестовых вектора", doc)
    parse_out = tool_output(["tools/parse-spec.py"])
    need = int(m.group(1)) if m else None
    got = re.search(r"векторов:\s*(\d+)", parse_out)
    check("«N тестовых вектора» в README == счётчик parse-spec.py",
          int(got.group(1)) if got else None, need)

    m = re.search(r"все\s+(\d+)\s+имён", doc)
    map_out = tool_output(["tools/check-readme-map.py"])
    need = int(m.group(1)) if m else None
    got = re.search(r"карта README:\s*(\d+) имён", map_out)
    check("«все N имён» в README == счётчик check-readme-map.py",
          int(got.group(1)) if got else None, need)


def main():
    print("== check-sync: зеркала и обещания README ==")
    print("-- константы Module1.bas ↔ tools/*-spec.py")
    check_numbers()
    print("-- таблицы строк и ширин Module1.bas ↔ README")
    check_rows_and_widths()
    print("-- состав проверок README ↔ CI")
    check_ci_matches_readme()
    print("-- счётчики README ↔ результаты инструментов")
    check_readme_counts()

    print("\nпроверок: %d, расхождений: %d" % (CHECKS, len(FAILS)))
    if FAILS:
        print("FAIL check-sync: %s" % ", ".join(FAILS))
        return 1
    print("check-sync: зеркала и README сходятся")
    return 0


if __name__ == "__main__":
    sys.exit(main())
