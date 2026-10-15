#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Example.xlsm против README и против самого себя.

Книга читается стандартной библиотекой (zipfile + xml.etree), без openpyxl:
берутся кэшированные значения ячеек — то, что макрос получает через Value2
(формулы здесь не вычисляются, Excel сохраняет их результат в книге).

Проверки:
  1. «Значения примера» из README (`B18`=1800, `D18`=4, …) == ячейки листа
     «1» — README описывает именно эту книгу;
  2. каждое наименование из строк defRows (Module1.bas) с количеством > 0
     есть в Service!E с положительной шириной в Service!H — так ищет
     ШиринаИзService: UCase + Trim + NBSP → пробел, первая строка с шириной
     > 0. Иначе макрос возьмёт запасную ширину с предупреждением, а README
     обещает прогон примера без блока «ВНИМАНИЕ»; по той же причине
     количество должно быть числом, а не текстом;
  3. одно наименование в Service — одна ширина: при разных H макрос берёт
     первую подходящую строку, а выпадающий список в столбце C листа «1»
     может ссылаться на другую (так было с AIO.2.xx/2: 164 и 160 мм).

Запуск: python3 tools/check-example.py [книга.xlsm]  (по умолчанию —
Example.xlsm в корне репозитория).
Код возврата: 0 — расхождений нет, 1 — есть.
"""

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "Example.xlsm"
BAS = ROOT / "Module1.bas"
README = ROOT / "README.md"

SHEET_DATA = "1"
SHEET_SERVICE = "Service"
COL_NAME, COL_QTY = "C", "D"          # лист «1»
SVC_NAME, SVC_WIDTH = "E", "H"        # Service (SVC_COL_NAME / SVC_COL_WIDTH)

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# Числоподобный текст: его потребители макроса берут через IsNumeric → CDbl.
NUMERIC = re.compile(r"^\s*[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)\s*$")

FAILS = []
CHECKS = 0


def check(ok, label, detail=""):
    global CHECKS
    CHECKS += 1
    if ok:
        print("ok   " + label)
    else:
        FAILS.append(label)
        print("FAIL " + label + (("\n       " + detail) if detail else ""))


class XlError(str):
    """Значение ячейки-ошибки (#N/A, #REF! …) — не путать с текстом."""


class Book:
    def __init__(self, path):
        self.zip = zipfile.ZipFile(str(path))
        self.strings = self._shared_strings()
        self.paths = self._sheet_paths()

    def _shared_strings(self):
        try:
            root = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        except KeyError:
            return []
        out = []
        for si in root.findall("m:si", NS):
            # Простой текст — <t>, форматированный — прогоны <r><t>;
            # фонетика <rPh> в значение не входит.
            parts = si.findall("m:t", NS) + si.findall("m:r/m:t", NS)
            out.append("".join(t.text or "" for t in parts))
        return out

    def _sheet_paths(self):
        wb = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        target = {r.get("Id"): r.get("Target")
                  for r in rels.findall("rel:Relationship", NS)}
        out = {}
        for sheet in wb.findall("m:sheets/m:sheet", NS):
            t = target[sheet.get(R_ID)]
            out[sheet.get("name")] = t.lstrip("/") if t.startswith("/") \
                else "xl/" + t
        return out

    def cells(self, name):
        """{адрес: значение} листа — как Value2: число, текст, bool, XlError."""
        root = ET.fromstring(self.zip.read(self.paths[name]))
        out = {}
        for c in root.iterfind("m:sheetData/m:row/m:c", NS):
            kind = c.get("t", "n")
            v = c.find("m:v", NS)
            if kind == "inlineStr":
                value = "".join(t.text or "" for t in c.iter(
                    "{%s}t" % NS["m"]))
            elif v is None or v.text is None:
                continue
            elif kind == "s":
                value = self.strings[int(v.text)]
            elif kind == "str":
                value = v.text
            elif kind == "e":
                value = XlError(v.text)
            elif kind == "b":
                value = v.text == "1"
            else:
                value = float(v.text)
            out[c.get("r")] = value
        return out


def as_text(v):
    """CStr(Value2): целое число без «.0»."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def as_number(v):
    """Число или числоподобный текст → float, иначе None."""
    if isinstance(v, bool) or isinstance(v, XlError):
        return None
    if isinstance(v, float):
        return v
    if isinstance(v, str) and NUMERIC.match(v):
        return float(v.replace(",", "."))
    return None


def norm_service(name):
    """НормИмяService: UCase$(Trim$(Replace(nm, NBSP, " ")))."""
    return name.replace("\u00a0", " ").strip(" ").upper()


def def_rows():
    text = BAS.read_bytes().decode("cp1251")
    m = re.search(r"(?ms)^\s*defRows = Array\((.*?)\)", text)
    return [int(x) for x in re.findall(r"\d+", m.group(1))] if m else []


def readme_example_values():
    """«Значения примера: `B18`=1800, `D18`=4, `D19`=0,2, `D12`=`Yes`, …»."""
    doc = README.read_text(encoding="utf-8")
    m = re.search(r"Значения примера:(.*?)`D26`", doc, re.S)
    if not m:
        return None
    out = {}
    for addr, raw in re.findall(r"`([A-Z]+[0-9]+)`=(`[^`]*`|[0-9]+(?:[.,][0-9]+)?)",
                                m.group(1)):
        out[addr] = raw.strip("`") if raw.startswith("`") \
            else float(raw.replace(",", "."))
    return out


def check_readme_values(data):
    promised = readme_example_values()
    check(bool(promised), "README: «Значения примера: `ЯЧЕЙКА`=…» найдены")
    for addr, want in sorted((promised or {}).items()):
        got = data.get(addr)
        if isinstance(want, float):
            num = as_number(got)
            ok = num is not None and abs(num - want) < 1e-9
        else:
            ok = isinstance(got, str) and not isinstance(got, XlError) \
                and got == want
        check(ok, "README `%s`=%s == «%s»!%s"
              % (addr, as_text(want), SHEET_DATA, addr),
              "в книге: %r — поправьте README или книгу" % (got,))


def service_rows(svc):
    """[(строка, ключ, имя, ширина или None)] в порядке строк Service."""
    rows = []
    for ref, value in svc.items():
        col, row = re.match(r"([A-Z]+)(\d+)$", ref).groups()
        if col != SVC_NAME or isinstance(value, XlError):
            continue
        name = as_text(value)
        key = norm_service(name)
        if not key:
            continue
        width = as_number(svc.get(SVC_WIDTH + row))
        rows.append((int(row), key, name.strip(), width))
    rows.sort()
    return rows


def check_names_in_service(data, rows):
    for r in def_rows():
        qty_raw = data.get(COL_QTY + str(r))
        qty = as_number(qty_raw)
        if qty is None:
            text = "" if qty_raw is None else as_text(qty_raw).strip()
            if text:
                check(False, "строка %d: количество — число" % r,
                      "в D%d текст %r — макрос выдаст предупреждение"
                      % (r, text))
            continue
        if int(qty) <= 0:
            continue
        name = as_text(data.get(COL_NAME + str(r)))
        key = norm_service(name)
        hit = next((row for row, k, _n, w in rows
                    if k == key and w is not None and w > 0), None) \
            if key else None
        width = next((w for row, k, _n, w in rows if row == hit), None)
        check(hit is not None,
              "строка %d: %s ×%d → %s"
              % (r, name.strip() or "<пусто>", int(qty),
                 "%s мм (Service!%s%d)" % (as_text(width), SVC_WIDTH, hit)
                 if hit else "нет в Service с шириной > 0"),
              "макрос возьмёт запасную ширину с предупреждением — "
              "добавьте наименование или ширину в Service")


def check_width_conflicts(rows):
    widths = {}
    for row, key, name, w in rows:
        if w is not None and w > 0:
            widths.setdefault(key, (name, {}))[1].setdefault(w, []).append(row)
    conflicts = [(name, by_w) for name, by_w in widths.values()
                 if len(by_w) > 1]
    for name, by_w in conflicts:
        check(False, "Service: у «%s» разные ширины" % name,
              ", ".join("%s мм (%s %s)" % (as_text(w),
                                          "строка" if len(rs) == 1
                                          else "строки",
                                          ", ".join(map(str, rs)))
                        for w, rs in sorted(by_w.items()))
              + " — макрос берёт первую")
    if not conflicts:
        check(True, "Service: одно наименование — одна ширина (%d имён)"
              % len(widths))


def main(argv):
    book_path = Path(argv[1]) if len(argv) > 1 else BOOK
    print("== check-example: %s ==" % book_path.name)
    book = Book(book_path)
    missing = [s for s in (SHEET_DATA, SHEET_SERVICE) if s not in book.paths]
    check(not missing, "листы «%s» и «%s» есть" % (SHEET_DATA, SHEET_SERVICE),
          "нет: %s" % ", ".join(missing))
    if missing:
        return 1

    data = book.cells(SHEET_DATA)
    rows = service_rows(book.cells(SHEET_SERVICE))

    print("-- README «Значения примера» ↔ лист «1»")
    check_readme_values(data)
    print("-- наименования строк defRows ↔ Service (ширина)")
    check_names_in_service(data, rows)
    print("-- Service: ширины без противоречий")
    check_width_conflicts(rows)

    print("\nпроверок: %d, расхождений: %d" % (CHECKS, len(FAILS)))
    if FAILS:
        print("FAIL check-example")
        return 1
    print("check-example: пример сходится с README и Service")
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
