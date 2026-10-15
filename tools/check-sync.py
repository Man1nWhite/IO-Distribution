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
  4. команды раздела «Проверки» README ↔ шаги джоба lint в CI, и они же ↔
     `tools/check-all.py --list` (тот же порядок): новый инструмент обязан
     попасть во все три места, иначе он тихо не работает; джоб windows
     запускает `python tools/check-all.py`;
  5. счётчики, обещанные README («N тестовых вектора», «все N имён»), ↔
     фактические результаты parse-spec.py и check-readme-map.py;
  6. вывод каждого `tools/*.py` переживает pipe в cp1252 — так Python пишет
     на Windows в Git Bash, в файл и в CI: без перевода вывода в UTF-8
     «↔», «−», «—» и кириллица роняют инструмент (UnicodeEncodeError);
  7. константы, названные в README как «`ИМЯ` = N» (`MAX_LONG`, `MAX_QTY`,
     `PLAN_NODE_LIMIT`, …), ↔ их значения в Module1.bas; имя, которого в
     коде нет, — тоже расхождение (константу переименовали, README забыли);
  8. номера `Err.Raise vbObjectError + N` ↔ таблица «Ошибки» в README, в обе
     стороны: новая ошибка описана, описанная — существует.

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
TOOLS = ROOT / "tools"
PLAN_SPEC = TOOLS / "plan-spec.py"
PARSE_SPEC = TOOLS / "parse-spec.py"
CHECK_ALL = TOOLS / "check-all.py"
SELF = Path(__file__).resolve()

# Признак вложенного запуска: проверка 6 запускает и сам check-sync.py —
# внутри он пропускает её, иначе рекурсия.
NESTED_ENV = "CHECK_SYNC_NESTED"

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


def check_readme_constants():
    """Числа констант, обещанные README («`ИМЯ` = N»), ↔ Module1.bas."""
    # Имя с «_» — константа модуля; так не цепляются ячейки (`D18`=4).
    found = sorted(set(re.findall(
        r"`([A-Z][A-Z0-9]*_[A-Z0-9_]+)`\s*=\s*([0-9][0-9.]*)", readme())))
    if not found:
        check("README называет константы как «`ИМЯ` = N»",
              "<ни одной>", "<хотя бы одна>")
        return
    for name, value in found:
        got = bas_const(name)
        check("README: `%s` = %s == Module1.bas" % (name, value),
              Decimal(got) if got is not None else "<нет такой константы>",
              Decimal(value))


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


def as_ranges(nums):
    """[10, 11, 12, 15] → «10–12, 15»: расхождение видно с одного взгляда."""
    out = []
    for n in sorted(nums):
        if out and out[-1][1] == n - 1:
            out[-1][1] = n
        else:
            out.append([n, n])
    return ", ".join(str(a) if a == b else "%d–%d" % (a, b) for a, b in out)


def readme_error_codes():
    """Номера из первого столбца таблицы раздела «Ошибки» README.

    Ячейки вида «10 / 11», «**28**», «20–27, 33–37, 40»: диапазоны через
    тире раскрываются, разделители — «,» и «/».
    """
    m = re.search(r"(?ms)^## Ошибки.*?\n(.*?)(?=^## |\Z)", readme())
    if not m:
        return None
    nums = set()
    for line in m.group(1).splitlines():
        if not line.startswith("| ") or line.startswith("| N "):
            continue
        for tok in re.split(r"[,/]", line.split("|")[1].replace("*", "")):
            tok = tok.strip()
            span = re.fullmatch(r"(\d+)\s*[–-]\s*(\d+)", tok)
            if span:
                nums.update(range(int(span.group(1)), int(span.group(2)) + 1))
            elif tok.isdigit():
                nums.add(int(tok))
    return nums


def check_error_codes():
    # Каждый Err.Raise vbObjectError + N описан в README, и README не
    # обещает номеров, которых в коде нет (диапазон «60–78» при отсутствии
    # 74 — тоже расхождение).
    code = {int(n) for n in re.findall(r"vbObjectError \+ (\d+)", bas_text())}
    doc = readme_error_codes()
    check("номера ошибок vbObjectError + N == таблица «Ошибки» в README",
          as_ranges(code), as_ranges(doc) if doc is not None else None)


def readme_commands():
    m = re.search(r"## Проверки\n\n```sh\n(.*?)```", readme(), re.S)
    if not m:
        return None
    return [line.strip() for line in m.group(1).splitlines() if line.strip()]


def ci_job(name):
    """Текст джоба `name` из workflow (до следующего джоба того же уровня)."""
    text = WORKFLOW.read_text(encoding="utf-8")
    m = re.search(r"(?ms)^  %s:[ \t]*\n(.*?)(?=^  \S|\Z)" % re.escape(name),
                  text)
    return m.group(1) if m else ""


def ci_commands():
    """Эталонные шаги CI — `run: python3 tools/…` джоба lint.

    Джоб windows гоняет их же одной командой (check-all.py) и в сверку
    состава не входит.
    """
    return re.findall(r"(?m)^\s*run:\s*(python3 tools/[\w.-]+\.py)",
                      ci_job("lint"))


def check_windows_job():
    # README обещает: джоб windows — те же шаги через tools/check-all.py.
    runs = re.findall(r"(?m)^\s*run:\s*(.+?)\s*$", ci_job("windows"))
    check("джоб windows в CI запускает tools/check-all.py",
          runs, ["python tools/check-all.py"])


def run_tool(path, args=(), stdout_encoding="utf-8", extra_env=None):
    """Запустить инструмент; PYTHONIOENCODING задаёт кодировку его вывода.

    Возвращает (код, stdout, stderr). Читается как UTF-8: инструменты сами
    переводят вывод в UTF-8, какую бы кодировку ни подсунули снаружи.
    """
    env = dict(os.environ, PYTHONIOENCODING=stdout_encoding)
    env.pop(NESTED_ENV, None)
    env.update(extra_env or {})
    proc = subprocess.run([sys.executable, str(path)] + list(args),
                          cwd=str(ROOT), env=env, capture_output=True)
    return (proc.returncode,
            proc.stdout.decode("utf-8", errors="replace"),
            proc.stderr.decode("utf-8", errors="replace"))


def check_ci_matches_readme():
    doc = readme_commands()
    ci = ci_commands()
    if doc is None:
        check("раздел «Проверки» README == шаги CI", None, ci)
        return
    # got — всегда код (шаги CI), want — доки (README).
    check("раздел «Проверки» README == шаги CI",
          sorted(ci), sorted(doc))


def check_all_commands():
    """Шаги tools/check-all.py — как он сам их перечисляет (`--list`)."""
    code, out, _err = run_tool(CHECK_ALL, ["--list"])
    if code != 0:
        return None
    return [line.strip() for line in out.splitlines() if line.strip()]


def check_ci_matches_check_all():
    # got — шаги CI, want — зеркало (check-all.py); порядок важен: «ровно те
    # же шаги, что в CI», и первый упавший останавливает прогон в обоих.
    check("шаги tools/check-all.py == шаги CI (тот же порядок)",
          ci_commands(), check_all_commands())


def check_output_encoding():
    """Каждый tools/*.py переживает stdout в cp1252, как на Windows-pipe."""
    broken = []
    for path in sorted(TOOLS.glob("*.py")):
        # check-all.py без аргументов запустил бы всё, в том числе этот
        # скрипт; его вывод проверяем на `--list`. Сам check-sync.py
        # запускается с NESTED_ENV и эту проверку внутри пропускает.
        args = ["--list"] if path.resolve() == CHECK_ALL.resolve() else []
        extra = {NESTED_ENV: "1"} if path.resolve() == SELF else None
        _code, _out, err = run_tool(path, args, "cp1252", extra)
        if "UnicodeEncodeError" in err:
            broken.append(path.name)
    check("вывод tools/*.py в pipe с cp1252 (как на Windows) — "
          "без UnicodeEncodeError", broken, [])


def tool_output(path):
    _code, out, err = run_tool(path)
    return out + err


def check_readme_counts():
    """Числа, обещанные README, против фактических результатов инструментов."""
    doc = readme()

    # «34 тестовых вектора» / «36 тестовых векторов» — форма по числу.
    m = re.search(r"(\d+)\s+тестовых\s+вектор\w*", doc)
    parse_out = tool_output(PARSE_SPEC)
    need = int(m.group(1)) if m else None
    got = re.search(r"векторов:\s*(\d+)", parse_out)
    check("«N тестовых вектора» в README == счётчик parse-spec.py",
          int(got.group(1)) if got else None, need)

    m = re.search(r"все\s+(\d+)\s+имён", doc)
    map_out = tool_output(TOOLS / "check-readme-map.py")
    need = int(m.group(1)) if m else None
    got = re.search(r"карта README:\s*(\d+) имён", map_out)
    check("«все N имён» в README == счётчик check-readme-map.py",
          int(got.group(1)) if got else None, need)


def main():
    print("== check-sync: зеркала и обещания README ==")
    print("-- константы Module1.bas ↔ tools/*-spec.py")
    check_numbers()
    print("-- константы, названные в README ↔ Module1.bas")
    check_readme_constants()
    print("-- таблицы строк и ширин Module1.bas ↔ README")
    check_rows_and_widths()
    print("-- номера ошибок Module1.bas ↔ README")
    check_error_codes()
    print("-- состав проверок README ↔ CI ↔ tools/check-all.py")
    check_ci_matches_readme()
    check_ci_matches_check_all()
    check_windows_job()
    print("-- счётчики README ↔ результаты инструментов")
    check_readme_counts()
    if not os.environ.get(NESTED_ENV):
        print("-- вывод инструментов в не-UTF-8 pipe (Windows)")
        check_output_encoding()

    print("\nпроверок: %d, расхождений: %d" % (CHECKS, len(FAILS)))
    if FAILS:
        print("FAIL check-sync: %s" % ", ".join(FAILS))
        return 1
    print("check-sync: зеркала и README сходятся")
    return 0


if __name__ == "__main__":
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main())
