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
     запускает `python tools/check-all.py`. Обратная сторона: каждый
     `tools/*.py` обязан быть чьим-то шагом (кроме самого check-all.py), а
     шаг — ссылаться на существующий файл;
  5. счётчики, обещанные README («N векторов» — каждое вхождение, «все N
     имён», «N логических строк», если оно названо), ↔ фактические
     результаты parse-spec.py, check-readme-map.py и vba-lint.py;
     каталог правил линтера («E01–EN / W01–WN») ↔ его докстринг;
  6. вывод каждого `tools/*.py` переживает pipe в cp1252 — так Python пишет
     на Windows в Git Bash, в файл и в CI: без перевода вывода в UTF-8
     «↔», «−», «—» и кириллица роняют инструмент (UnicodeEncodeError);
  7. константы, названные в README как «`ИМЯ` = N» (`MAX_LONG`, `MAX_QTY`,
     `PLAN_NODE_LIMIT`, …), ↔ их значения в Module1.bas; имя, которого в
     коде нет, — тоже расхождение (константу переименовали, README забыли);
  8. номера `Err.Raise vbObjectError + N` ↔ таблица «Ошибки» в README, в обе
     стороны: новая ошибка описана, описанная — существует;
  9. векторы самопроверки ТестРазбораЧисел (Module1.bas, гоняется в Excel)
     ↔ VECTORS / WRAP / KOL / REZERV из tools/parse-spec.py;
 10. выпуск отложен — README и AGENTS.md не упоминают выпуски (без версий,
     «релизов»/«тегов» и ссылок compare|releases/tag);
 11. локальный хук `.githooks/pre-commit` запускает `tools/vba-lint.py`,
     а README описывает его установку (`core.hooksPath .githooks`);
 12. AGENTS.md — точка входа для агентов: существует, зовёт на
     `tools/check-all.py` и описывает чтение cp1251 через iconv.

Что НЕ проверяется: сами правила разбора и укладки (они переносятся по смыслу)
— только их параметры и обещанные числа.

Код возврата: 0 — всё сошлось, 1 — есть расхождение.
"""

import os
import re
import runpy
import subprocess
import sys
from collections import Counter
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BAS = ROOT / "Module1.bas"
README = ROOT / "README.md"
AGENTS = ROOT / "AGENTS.md"
WORKFLOW = ROOT / ".github" / "workflows" / "vba-lint.yml"
TOOLS = ROOT / "tools"
PLAN_SPEC = TOOLS / "plan-spec.py"
PARSE_SPEC = TOOLS / "parse-spec.py"
CHECK_ALL = TOOLS / "check-all.py"
HOOK = ROOT / ".githooks" / "pre-commit"
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
    """Константа из Module1.bas — строкой, как записана (число или True/False)."""
    m = re.search(r"(?m)^Private Const %s As \w+ = (True|False|[0-9][0-9.]*)#?" %
                  re.escape(name), bas_text())
    if not m:
        return None
    return m.group(1)


def const_value(text):
    """Значение константы для сравнения: bool для True/False, иначе Decimal."""
    if text in ("True", "False"):
        return text == "True"
    return Decimal(text)


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
    """Значения констант, обещанные README («`ИМЯ` = N»), ↔ Module1.bas.

    Булевы (`SWITCH_RESERVE_*` = True/False) сверяются наравне с числами:
    режим резерва зоны — такое же обещание README, как и её размер.
    """
    # Имя с «_» — константа модуля; так не цепляются ячейки (`D18`=4).
    found = sorted(set(re.findall(
        r"`([A-Z][A-Z0-9]*_[A-Z0-9_]+)`\s*=\s*(True|False|[0-9][0-9.]*)",
        readme())))
    if not found:
        check("README называет константы как «`ИМЯ` = N»",
              "<ни одной>", "<хотя бы одна>")
        return
    for name, value in found:
        got = bas_const(name)
        check("README: `%s` = %s == Module1.bas" % (name, value),
              const_value(got) if got is not None else "<нет такой константы>",
              const_value(value))


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


# Строковое выражение VBA из литералов и ChrW(N): "1" & ChrW(160) & "500".
VBA_STR = r'(?:"(?:[^"]|"")*"|ChrW\$?\(\d+\))'
VBA_TEST_CALL = re.compile(
    r"(?m)^\s*(ТестЧисла|ТестЦелого|ТестКоличества|ТестРезерва)\s+(%s(?:\s*&\s*%s)*)\s*,"
    r"\s*(-?MAX_LONG|-?\d+(?:\.\d+)?)\s*,\s*n\s*,\s*bad\s*$"
    % (VBA_STR, VBA_STR))


def vba_string(expr):
    parts = re.findall(r'"((?:[^"]|"")*)"|ChrW\$?\((\d+)\)', expr)
    return "".join(chr(int(code)) if code else lit.replace('""', '"')
                   for lit, code in parts)


def check_parse_vectors():
    # ТестРазбораЧисел гоняет векторы parse-spec.py в самом Excel; список в
    # VBA — копия, и она не должна отставать. got — только в VBA, want —
    # только в parse-spec.py: сходятся, когда оба пусты.
    spec = runpy.run_path(str(PARSE_SPEC))
    top = spec["MAX_LONG"]
    py = {"ТестЧисла": Counter((s, float(w)) for s, w in spec["VECTORS"]),
          "ТестЦелого": Counter((s, float(w)) for s, w in spec["WRAP"]),
          "ТестКоличества": Counter((s, float(w)) for s, w in spec["KOL"]),
          "ТестРезерва": Counter((s, float(w)) for s, w in spec["REZERV"])}
    vba = {"ТестЧисла": Counter(), "ТестЦелого": Counter(),
           "ТестКоличества": Counter(), "ТестРезерва": Counter()}
    for kind, expr, num in VBA_TEST_CALL.findall(bas_text()):
        want = {"MAX_LONG": top, "-MAX_LONG": -top}.get(num)
        vba[kind][(vba_string(expr), float(num if want is None else want))] += 1
    for kind, func, table in (("ТестЧисла", "ЧислоИзТекста", "VECTORS"),
                              ("ТестЦелого", "ИзвлечьЧисло", "WRAP"),
                              ("ТестКоличества", "КолИзЗначения", "KOL"),
                              ("ТестРезерва", "РезервДоля", "REZERV")):
        check("векторы %s: ТестРазбораЧисел (Module1.bas) == %s "
              "(tools/parse-spec.py)" % (func, table),
              sorted((vba[kind] - py[kind]).elements()),
              sorted((py[kind] - vba[kind]).elements()))


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


def check_tools_registered():
    """Каждый tools/*.py попал в шаги — иначе он молча не работает.

    AGENTS.md требует для нового инструмента три места (STEPS в
    check-all.py, шаг CI, раздел «Проверки» README), но сверялся только
    состав уже названных шагов: инструмент, не попавший никуда, лежал бы в
    репозитории и никогда не запускался. check-all.py — сам запускающий, он
    шагом не считается.
    """
    steps = ci_commands()
    registered = {Path(s.split()[-1]).name for s in steps}
    scripts = {p.name for p in TOOLS.glob("*.py")} - {CHECK_ALL.name}
    check("каждый tools/*.py — шаг CI (кроме самого check-all.py)",
          sorted(scripts - registered), [])
    check("шаги CI ссылаются на существующие tools/*.py",
          sorted(n for n in registered if not (TOOLS / n).is_file()), [])


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
    # got — всегда код (шаги CI), want — доки (README). Без sorted():
    # README обещает «ровно те же шаги … в том же порядке», и первый
    # упавший шаг останавливает прогон — порядок здесь часть обещания.
    check("раздел «Проверки» README == шаги CI (тот же порядок)",
          ci, doc)


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

    # «38 тестовых векторов» / «Разбор чисел: 38 векторов» — форма по числу.
    # Сверяется каждое вхождение: счётчик продублирован в README трижды, и
    # проверка только первого оставляла два остальных отставать молча.
    parse_out = tool_output(PARSE_SPEC)
    got = re.search(r"векторов:\s*(\d+)", parse_out)
    got = int(got.group(1)) if got else None
    lines = doc.splitlines()
    hits = [(i, m.group(0), int(m.group(1)))
            for i, line in enumerate(lines, 1)
            for m in re.finditer(r"(\d+)\s+(?:тестовых\s+)?вектор\w*", line)]
    if not hits:
        check("README называет число векторов («N векторов»)",
              "<ни одного>", "<хотя бы одно>")
    for no, phrase, need in hits:
        check("«%s» в README (строка %d) == счётчик parse-spec.py"
              % (phrase, no), got, need)

    m = re.search(r"все\s+(\d+)\s+имён", doc)
    map_out = tool_output(TOOLS / "check-readme-map.py")
    need = int(m.group(1)) if m else None
    got = re.search(r"карта README:\s*(\d+) имён", map_out)
    check("«все N имён» в README == счётчик check-readme-map.py",
          int(got.group(1)) if got else None, need)

    # Счётчик логических строк меняется любой правкой модуля и никакого
    # инварианта не выражает — README его не обещает. Если фраза вернётся,
    # она снова сверяется: обещанное число не должно отставать.
    m = re.search(r"(\d+)\s+логических\s+строк", doc)
    if m:
        lint_out = tool_output(TOOLS / "vba-lint.py")
        got = re.search(r"== Module1\.bas ==.*?\(логических строк (\d+)\)",
                        lint_out, re.S)
        check("«N логических строк» в README == счётчик vba-lint.py",
              int(got.group(1)) if got else None, int(m.group(1)))
    else:
        print("     «N логических строк» README не обещает — сверять нечего")


# Выпуск отложен: в README и AGENTS.md не должно быть упоминаний выпусков.
# Ловятся слова «релиз», «тег» (с начала слова — «категории» не попадает),
# «tag», SemVer/Keep a Changelog, ссылки releases/tag и compare, версии vN.N.
RELEASE_MENTION = re.compile(
    r"(?<![А-Яа-яЁёA-Za-z])релиз\w*"
    r"|(?<![А-Яа-яЁёA-Za-z])тег\w*"
    r"|\btag\b|SemVer|keepachangelog|Keep\s+a\s+Changelog"
    r"|releases/tag|/compare/|v\d+\.\d+", re.I)


def check_no_release_mentions():
    """Выпуск отложен: README и AGENTS.md не упоминают выпуски."""
    texts = [("README.md", readme())]
    if AGENTS.is_file():
        texts.append(("AGENTS.md", AGENTS.read_text(encoding="utf-8")))
    hits = ["%s: «%s»" % (name, m.group(0))
            for name, text in texts
            for m in RELEASE_MENTION.finditer(text)]
    check("выпуск отложен: README и AGENTS.md без упоминаний выпусков", hits, [])


def check_lint_catalog():
    """«Каталог E01–EN / W01–WN» в README ↔ правила в докстринге линтера.

    Новый номер правила в линтере без правки README — расхождение: каталог
    в README переставал описывать то, что линтер проверяет (E18 → E20).
    """
    lint = (TOOLS / "vba-lint.py").read_text(encoding="utf-8")
    parts = lint.split('"""')
    doc = parts[1] if len(parts) > 1 else ""
    top = {}
    for letter in ("E", "W"):
        nums = [int(n) for n in re.findall(r"\b%s(\d{2})\b" % letter, doc)]
        top[letter] = max(nums) if nums else 0
    got = "E01–E%02d / W01–W%02d" % (top["E"], top["W"])
    m = re.search(r"Каталог\s*(E\d+–E\d+\s*/\s*W\d+–W\d+)", readme())
    check("каталог правил линтера == «Каталог …» в README",
          got, m.group(1) if m else "<в README не назван>")


def check_precommit_hook():
    """Локальный хук ↔ README: vba-lint перед коммитом, установка описана."""
    hook = HOOK.read_text(encoding="utf-8") if HOOK.is_file() else ""
    check(".githooks/pre-commit запускает tools/vba-lint.py",
          "tools/vba-lint.py" in hook, True)
    check("README описывает установку хука (`core.hooksPath .githooks`)",
          bool(re.search(r"core\.hooksPath\s+\.githooks", readme())), True)


def check_agents_doc():
    """AGENTS.md — точка входа для агентов: есть и не разошёлся с реальностью."""
    text = AGENTS.read_text(encoding="utf-8") if AGENTS.is_file() else ""
    check("AGENTS.md существует и зовёт на tools/check-all.py",
          bool(text) and "tools/check-all.py" in text, True)
    check("AGENTS.md описывает чтение Module1.bas в cp1251 через iconv",
          "iconv" in text and "cp1251" in text, True)


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
    print("-- векторы разбора: ТестРазбораЧисел (VBA) ↔ tools/parse-spec.py")
    check_parse_vectors()
    print("-- состав проверок README ↔ CI ↔ tools/check-all.py")
    check_ci_matches_readme()
    check_ci_matches_check_all()
    check_tools_registered()
    check_windows_job()
    print("-- счётчики README ↔ результаты инструментов")
    check_readme_counts()
    check_lint_catalog()
    print("-- отложенный выпуск, AGENTS.md и локальный хук")
    check_no_release_mentions()
    check_agents_doc()
    check_precommit_hook()
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
