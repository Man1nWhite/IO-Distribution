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
     может ссылаться на другую (так было с AIO.2.xx/2: 164 и 160 мм);
  4. `Module1` внутри книги (`xl/vbaProject.bin`) == `Module1.bas`: Excel
     запускает копию из книги, а не файл из репозитория, и после правки
     `Module1.bas` книга молча отставала. vbaProject.bin читается как
     MS-CFB, поток VBA/dir и исходники модулей — как MS-OVBA (Excel не
     нужен); хвостовые переводы строк при сверке не считаются — их дописывает
     сам VBE при импорте и сохранении.

Ячейка-ошибка (#N/A, #REF!) в столбце D — не расхождение: макрос читает её
как 0 и молча пропускает строку (IsError → Exit Function в ЧитатьКол).
Проверка ведёт себя так же, иначе CI падал бы на книге, которую макрос
обрабатывает без замечаний.

`D26` макрос не пишет, поэтому обещание README «формула ROUNDUP по AE10 и
Service» проверяется. Если ячейка окажется неведущей в общей формуле
(t="shared"), текст по XML не восстановить — это говорится прямо, а не
выдаётся за потерянную формулу.

Отдельно печатается, что лежит в `D21` «Шкафов» — формула или число:
`ЗаписатьЧислоШкафов` заменяет формулу числом, и от `D21` считает `D45`.
Ни значение, ни наличие формулы не сверяются: и то и другое меняется
законно — до первого прогона формула-оценка, после — число раскладки.

Запуск: python3 tools/check-example.py [--sync] [книга.xlsm]  (по умолчанию —
Example.xlsm в корне репозитория).
Код возврата: 0 — расхождений нет, 1 — есть, 2 — неверные аргументы или
--sync не смог обновить книгу.

--sync сначала кладёт `Module1.bas` в `xl/vbaProject.bin` (если копия в книге
отличается), затем гонит обычные проверки по обновлённой книге. Остальные
части архива и метаданные его записей не меняются, листы и значения —
тоже; описание того, что именно делается с контейнером VBA, — у `sync_book`.
"""

import difflib
import os
import random
import re
import struct
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
# Заглушка для неведущей ячейки общей формулы: <f> есть, текста нет.
SHARED_FORMULA = "<общая формула>"

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

    def formulas(self, name):
        """{адрес: текст формулы} листа — только ячейки с <f>."""
        root = ET.fromstring(self.zip.read(self.paths[name]))
        out = {}
        for c in root.iterfind("m:sheetData/m:row/m:c", NS):
            f = c.find("m:f", NS)
            if f is not None:
                # Общая формула (t="shared", si=N) хранится один раз —
                # у остальных ячеек группы <f> пустой.
                out[c.get("r")] = (f.text or "").strip() or SHARED_FORMULA
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


def bas_text():
    return BAS.read_bytes().decode("cp1251")


def def_rows():
    m = re.search(r"(?ms)^\s*defRows = Array\((.*?)\)", bas_text())
    return [int(x) for x in re.findall(r"\d+", m.group(1))] if m else []


def bas_rows(const_name):
    """Номера строк из константы вида Private Const FAM_… As String = "28,29"."""
    m = re.search(r'(?m)^Private Const %s As String = "([\d,]+)"'
                  % re.escape(const_name), bas_text())
    return {int(x) for x in m.group(1).split(",")} if m else set()


def bas_number(const_name):
    m = re.search(r"(?m)^Private Const %s As \w+ = ([0-9][0-9.]*)#?"
                  % re.escape(const_name), bas_text())
    return float(m.group(1)) if m else None


def norm_category(s):
    """НормКатегория: НормИмяService плюс схлопывание двойных пробелов."""
    return " ".join(norm_service(s).split())


def channels_of(value):
    """Каналы из Service!F, как их читает макрос: целое в диапазоне Long.

    None — ячейка не годится (ошибка, пусто, текст, дробь, вне диапазона);
    такую строку КаналыИзService/КаналыПоКатегории пропускают и ищут дальше.
    """
    n = as_number(value)
    limit = bas_number("MAX_LONG")
    if n is None or n != int(n) or limit is None or n > limit:
        return None
    return int(n)


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
        if isinstance(qty_raw, XlError):
            # ЧитатьКол: IsError(v) → 0, строка не участвует и молчит.
            print("     строка %d: в %s%d %s — макрос читает 0, строка не участвует"
                  % (r, COL_QTY, r, qty_raw))
            continue
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


def check_channels(data, svc):
    """Каналы Service!F для используемых наименований — как их читает макрос.

    README обещает прогон примера без блока «ВНИМАНИЕ», а предупреждение
    даёт не только отсутствующая ширина, но и отсутствующие каналы:
    КаналыИзService (реле, барьеры, мультиплексоры) требует целое > 0,
    КаналыПоКатегории (модули 28–32) — целое ≥ 0 в своей категории, и
    предупреждает, только если нет ни AI, ни AO. Проверка ширины этого не
    покрывала: книга могла разойтись с обещанием незаметно.

    Берётся ПЕРВАЯ годная строка — как в макросе. Прочие строки с тем же
    наименованием и категорией, но другим F, печатаются справочно: в
    справочнике это режимы модуля («0-10 В поддержан не…»), и макрос их
    молча игнорирует, поэтому перестановка строк Service меняет ёмкость.
    """
    by_name = (bas_rows("FAM_DIO_REL") | bas_rows("ROWS_MUX")
               | bas_rows("FAM_DIO_BAR") | bas_rows("FAM_TC_BAR")
               | bas_rows("FAM_AI_BAR"))
    analog_mod = bas_rows("FAM_TC_MOD") | bas_rows("FAM_AI_MOD")
    cat_ai = norm_category("module AI")
    cat_ao = norm_category("module AO")
    # (ключ, категория) → [(строка, каналы)] в порядке строк Service.
    index = {}
    for ref, value in svc.items():
        col, row = re.match(r"([A-Z]+)(\d+)$", ref).groups()
        if col != SVC_NAME or isinstance(value, XlError):
            continue
        key = norm_service(as_text(value))
        if not key:
            continue
        cat = norm_category(as_text(svc.get("D" + row) or ""))
        index.setdefault((key, cat), []).append(
            (int(row), channels_of(svc.get("F" + row))))

    def first_valid(entries, minimum):
        return next(((r, ch) for r, ch in sorted(entries)
                     if ch is not None and ch >= minimum), (None, None))

    for r in def_rows():
        if isinstance(data.get(COL_QTY + str(r)), XlError):
            continue
        qty = as_number(data.get(COL_QTY + str(r)))
        if qty is None or int(qty) <= 0:
            continue
        name = as_text(data.get(COL_NAME + str(r))).strip()
        key = norm_service(name)
        if r in analog_mod:
            ai = index.get((key, cat_ai), [])
            ao = index.get((key, cat_ao), [])
            r_ai, ch_ai = first_valid(ai, 0)
            r_ao, ch_ao = first_valid(ao, 0)
            check(r_ai is not None or r_ao is not None,
                  "строка %d: %s — каналы AI/AO из Service" % (r, name),
                  "нет ни module AI, ни module AO с целым числом каналов — "
                  "макрос возьмёт AI×8 с предупреждением")
            if r_ai is not None or r_ao is not None:
                print("     строка %d: AI %s (Service!F%s), AO %s (Service!F%s)"
                      % (r, ch_ai if ch_ai is not None else "—",
                         r_ai if r_ai is not None else "—",
                         ch_ao if ch_ao is not None else "—",
                         r_ao if r_ao is not None else "—"))
            for cat, entries, used_row in (("module AI", ai, r_ai),
                                           ("module AO", ao, r_ao)):
                other = [(rr, ch) for rr, ch in sorted(entries)
                         if rr != used_row and ch is not None]
                if other and used_row is not None:
                    print("     %s: взята первая годная строка %d; прочие %s"
                          % (cat, used_row,
                             ", ".join("F%d=%d" % (rr, ch) for rr, ch in other)))
        elif r in by_name:
            entries = [e for (k, _c), lst in index.items() if k == key
                       for e in lst]
            row, ch = first_valid(entries, 1)
            check(row is not None,
                  "строка %d: %s — каналы %s (Service!F%s)"
                  % (r, name, ch, row) if row is not None else
                  "строка %d: %s — каналы из Service" % (r, name),
                  "нет целого положительного числа каналов — макрос возьмёт "
                  "запас с предупреждением")


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


# ---------------------------------------------------------------------------
# xl/vbaProject.bin стандартной библиотекой: контейнер MS-CFB, сжатие и каталог
# MS-OVBA. Нужны, чтобы сверить Module1 в книге с Module1.bas и (--sync) без
# Excel положить в книгу Module1.bas.
# ---------------------------------------------------------------------------

VBA_PROJECT = "xl/vbaProject.bin"

CFB_MAGIC = bytes.fromhex("d0cf11e0a1b11ae1")
CFB_MAXREG = 0xFFFFFFFA                # выше — служебные значения FAT
CFB_FATSECT, CFB_ENDOFCHAIN, CFB_FREESECT = 0xFFFFFFFD, 0xFFFFFFFE, 0xFFFFFFFF
CFB_NOSTREAM = 0xFFFFFFFF
CFB_SECTOR, CFB_MINI, CFB_CUTOFF, CFB_ENTRY = 512, 64, 4096, 128
CFB_HEADER_FAT = 109                   # секторов FAT, помещающихся в заголовок

# _VBA_PROJECT без кэша компиляции (MS-OVBA 2.3.4.1): Reserved1 = 0x61CC,
# Version = 0xFFFF, Reserved2 = 0x00, Reserved3 = 0x0000, и всё.
PURGED_VBA_PROJECT = bytes.fromhex("cc61ffff000000")
SRP_STREAM = re.compile(r"^VBA/__SRP_[0-9]+$", re.I)

OVBA_CHUNK = 4096
OVBA_CANDIDATES = 128                  # сколько прошлых совпадений пробовать


class FormatError(Exception):
    """vbaProject.bin не по MS-CFB / MS-OVBA; текст — что именно не так."""


def cfb_read(blob):
    """Контейнер MS-CFB (v3 и v4) → {путь: байты} всех потоков.

    Путь — имена хранилищ и потока через «/» без корневой записи:
    «PROJECT», «VBA/dir». Порядок — обход каталога в порядке имён.
    """
    if len(blob) < CFB_SECTOR or blob[:8] != CFB_MAGIC:
        raise FormatError("нет сигнатуры MS-CFB (D0 CF 11 E0 A1 B1 1A E1)")
    minor, major, order, shift, mini_shift = struct.unpack_from(
        "<5H", blob, 0x18)
    if order != 0xFFFE or mini_shift != 6 \
            or (major, shift) not in ((3, 9), (4, 12)):
        raise FormatError("заголовок MS-CFB не поддержан: версия %d.%d, "
                          "порядок байт %#06x, сдвиги %d/%d"
                          % (major, minor, order, shift, mini_shift))
    size = 1 << shift
    n_fat, dir_start, _tx, cutoff, mf_start, n_mf, dif_start, n_dif = \
        struct.unpack_from("<8I", blob, 0x2C)

    def sector(i):
        at = (i + 1) * size
        if i > CFB_MAXREG or at + size > len(blob):
            raise FormatError("сектор %d за концом файла" % i)
        return blob[at:at + size]

    def words(raw):
        return struct.unpack("<%dI" % (len(raw) // 4), raw)

    def chain(first, table):
        out, seen, i = [], set(), first
        while i != CFB_ENDOFCHAIN:
            if i >= len(table) or i in seen:
                raise FormatError("цепочка секторов с %d оборвана или "
                                  "зациклена" % first)
            seen.add(i)
            out.append(i)
            i = table[i]
        return out

    # FAT: первые 109 секторов — в заголовке, остальные — в цепочке DIFAT.
    fat_at = [i for i in struct.unpack_from("<%dI" % CFB_HEADER_FAT, blob,
                                            0x4C) if i != CFB_FREESECT]
    nxt = dif_start
    for _ in range(n_dif):
        ids = words(sector(nxt))
        fat_at += [i for i in ids[:-1] if i != CFB_FREESECT]
        nxt = ids[-1]
    if len(fat_at) != n_fat:
        raise FormatError("в DIFAT %d секторов FAT, в заголовке — %d"
                          % (len(fat_at), n_fat))
    fat = [w for i in fat_at for w in words(sector(i))]

    raw_dir = b"".join(sector(i) for i in chain(dir_start, fat))
    # Запись каталога: (имя, тип, левый, правый, ребёнок, начало, размер).
    entries = []
    for at in range(0, len(raw_dir), CFB_ENTRY):
        rec = raw_dir[at:at + CFB_ENTRY]
        nlen, kind, _color, left, right, child = struct.unpack_from(
            "<HBBIII", rec, 64)
        start, length = struct.unpack_from("<IQ", rec, 0x74)
        if kind and (nlen < 2 or nlen > 64 or nlen % 2):
            raise FormatError("запись каталога %d: длина имени %d"
                              % (len(entries), nlen))
        entries.append((rec[:max(nlen - 2, 0)].decode("utf-16-le"), kind,
                        left, right, child, start,
                        length & 0xFFFFFFFF if major == 3 else length))
    if not entries or entries[0][1] != 5:
        raise FormatError("в каталоге нет корневой записи")

    mini_fat = [w for i in (chain(mf_start, fat) if n_mf else [])
                for w in words(sector(i))]
    _name, _kind, _l, _r, root_child, root_start, root_len = entries[0]
    mini = (b"".join(sector(i) for i in chain(root_start, fat))[:root_len]
            if root_len else b"")

    def body(start, length):
        if not length:
            return b""
        if length < cutoff:
            data = b"".join(mini[i * CFB_MINI:(i + 1) * CFB_MINI]
                            for i in chain(start, mini_fat))
        else:
            data = b"".join(sector(i) for i in chain(start, fat))
        if len(data) < length:
            raise FormatError("поток короче заявленных %d байт" % length)
        return data[:length]

    out, seen = {}, {0}

    def walk(first, prefix):
        # Братья хранилища — двоичное дерево: левое, узел, правое.
        stack, node = [], first
        while stack or node != CFB_NOSTREAM:
            while node != CFB_NOSTREAM:
                if node >= len(entries) or node in seen:
                    raise FormatError("запись каталога %d повторяется или "
                                      "не существует" % node)
                seen.add(node)
                stack.append(node)
                node = entries[node][2]
            node = stack.pop()
            name, kind, _l, right, child, start, length = entries[node]
            if kind == 1:
                walk(child, prefix + name + "/")
            elif kind == 2:
                out[prefix + name] = body(start, length)
            else:
                raise FormatError("запись каталога %d («%s»): тип %d"
                                  % (node, name, kind))
            node = right

    walk(root_child, "")
    return out


def cfb_name_key(name):
    """Порядок записей-братьев (MS-CFB 2.6.4): сначала длина имени, затем
    коды UTF-16 имени в верхнем регистре."""
    up = "".join(c.upper() if len(c.upper()) == 1 else c for c in name)
    raw = up.encode("utf-16-le")
    return len(raw) // 2, struct.unpack("<%dH" % (len(raw) // 2), raw)


def cfb_write(streams):
    """{путь: байты} → контейнер MS-CFB v3 (сектор 512 Б, мини-сектор 64 Б,
    порог мини-потока 4096 Б).

    Раскладка: каталог, мини-FAT, мини-поток, большие потоки, в конце FAT.
    Братья каждого хранилища — сбалансированное красно-чёрное дерево: нижний
    ярус красный, остальные чёрные (так цвета верны при любом числе записей),
    имена упорядочены по cfb_name_key. Свободные записи в хвосте каталога —
    нули и NOSTREAM у ссылок, как требует MS-CFB.
    """
    root = {"name": "Root Entry", "kind": 5, "kids": {}, "color": 1}
    for path, data in streams.items():
        *folders, leaf = path.split("/")
        node = root
        for name in folders + [leaf]:
            if len(name) > 31:
                raise FormatError("имя «%s» длиннее 31 символа" % name)
        for folder in folders:
            node = node["kids"].setdefault(
                cfb_name_key(folder),
                {"name": folder, "kind": 1, "kids": {}})
        if cfb_name_key(leaf) in node["kids"]:
            raise FormatError("поток %s записан дважды" % path)
        node["kids"][cfb_name_key(leaf)] = {
            "name": leaf, "kind": 2, "data": bytes(data)}

    nodes, queue = [root], [root]
    for storage in queue:              # очередь растёт по ходу обхода
        storage["sorted"] = [storage["kids"][k]
                             for k in sorted(storage["kids"])]
        for kid in storage["sorted"]:
            kid["id"] = len(nodes)
            nodes.append(kid)
            if kid["kind"] == 1:
                queue.append(kid)
    root["id"] = 0

    def link(kids):
        """Дерево братьев по отсортированному списку → id его корня."""
        last = len(kids).bit_length() - 1          # нижний ярус

        def build(lo, hi, level):
            if lo >= hi:
                return CFB_NOSTREAM
            mid = (lo + hi) // 2
            kid = kids[mid]
            kid["left"] = build(lo, mid, level + 1)
            kid["right"] = build(mid + 1, hi, level + 1)
            kid["color"] = 0 if last > 0 and level == last else 1
            return kid["id"]
        return build(0, len(kids), 0)

    for node in nodes:
        node["child"] = link(node["sorted"]) if node["kind"] != 2 \
            else CFB_NOSTREAM

    leaves = [n for n in nodes if n["kind"] == 2]
    small = [n for n in leaves if 0 < len(n["data"]) < CFB_CUTOFF]
    large = [n for n in leaves if len(n["data"]) >= CFB_CUTOFF]
    mini, mini_fat = bytearray(), []
    for node in small:
        count = -(-len(node["data"]) // CFB_MINI)
        node["start"] = len(mini_fat)
        mini_fat += [len(mini_fat) + k + 1 for k in range(count - 1)]
        mini_fat.append(CFB_ENDOFCHAIN)
        mini += node["data"].ljust(count * CFB_MINI, b"\0")

    n_dir = -(-len(nodes) // (CFB_SECTOR // CFB_ENTRY))
    n_mf = -(-len(mini_fat) * 4 // CFB_SECTOR)
    n_ms = -(-len(mini) // CFB_SECTOR)
    at = 0
    dir_start, at = at, at + n_dir
    mf_start, at = (at if n_mf else CFB_ENDOFCHAIN), at + n_mf
    ms_start, at = (at if n_ms else CFB_ENDOFCHAIN), at + n_ms
    for node in large:
        node["start"], at = at, at + -(-len(node["data"]) // CFB_SECTOR)
    n_fat = -(-at // (CFB_SECTOR // 4 - 1))        # FAT описывает и себя
    if n_fat > CFB_HEADER_FAT:
        raise FormatError("vbaProject.bin больше 7 МБ — DIFAT не поддержан")

    fat = [CFB_FREESECT] * (n_fat * CFB_SECTOR // 4)

    def extend(first, count):
        for k in range(count):
            fat[first + k] = first + k + 1 if k < count - 1 \
                else CFB_ENDOFCHAIN
    extend(dir_start, n_dir)
    if n_mf:
        extend(mf_start, n_mf)
    if n_ms:
        extend(ms_start, n_ms)
    for node in large:
        extend(node["start"], -(-len(node["data"]) // CFB_SECTOR))
    for k in range(n_fat):
        fat[at + k] = CFB_FATSECT

    def entry(node):
        raw = node["name"].encode("utf-16-le")
        if node["kind"] == 5:
            start, length = ms_start, len(mini)
        elif node["kind"] == 1:
            start, length = 0, 0
        else:
            length = len(node["data"])
            start = node.get("start", CFB_ENDOFCHAIN) if length \
                else CFB_ENDOFCHAIN
        return (raw.ljust(64, b"\0")
                + struct.pack("<HBBIII", len(raw) + 2, node["kind"],
                              node["color"],
                              node.get("left", CFB_NOSTREAM),
                              node.get("right", CFB_NOSTREAM), node["child"])
                + bytes(16)            # CLSID
                + struct.pack("<IQQIQ", 0, 0, 0, start, length))
    free = (bytes(64) + struct.pack("<HBBIII", 0, 0, 0, CFB_NOSTREAM,
                                    CFB_NOSTREAM, CFB_NOSTREAM) + bytes(48))
    table = b"".join(entry(n) for n in nodes) \
        + free * (n_dir * CFB_SECTOR // CFB_ENTRY - len(nodes))

    header = bytearray(CFB_SECTOR)
    header[:8] = CFB_MAGIC
    struct.pack_into("<5H", header, 0x18, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<9I", header, 0x28, 0, n_fat, dir_start, 0, CFB_CUTOFF,
                     mf_start, n_mf, CFB_ENDOFCHAIN, 0)
    struct.pack_into("<%dI" % CFB_HEADER_FAT, header, 0x4C,
                     *([at + k for k in range(n_fat)]
                       + [CFB_FREESECT] * (CFB_HEADER_FAT - n_fat)))

    body = bytearray(table)
    body += struct.pack("<%dI" % (n_mf * CFB_SECTOR // 4), *(
        mini_fat + [CFB_FREESECT] * (n_mf * CFB_SECTOR // 4 - len(mini_fat))))
    body += mini.ljust(n_ms * CFB_SECTOR, b"\0")
    for node in large:
        body += node["data"].ljust(-(-len(node["data"]) // CFB_SECTOR)
                                   * CFB_SECTOR, b"\0")
    body += struct.pack("<%dI" % len(fat), *fat)
    return bytes(header) + bytes(body)


def ovba_decompress(data):
    """Сжатый контейнер MS-OVBA (2.4.1) → исходные байты."""
    if not data or data[0] != 1:
        raise FormatError("сжатый исходник: нет сигнатуры 0x01")
    out, pos = bytearray(), 1
    while pos < len(data):
        if pos + 2 > len(data):
            raise FormatError("сжатый исходник оборван в заголовке куска")
        head = struct.unpack_from("<H", data, pos)[0]
        end = pos + (head & 0x0FFF) + 3
        if (head >> 12) & 7 != 3 or end > len(data):
            raise FormatError("сжатый исходник: битый кусок на смещении %d"
                              % pos)
        pos += 2
        if not head & 0x8000:                      # сырой кусок, 4096 Б
            if end - pos != OVBA_CHUNK:
                raise FormatError("сырой кусок не 4096 байт")
            out += data[pos:end]
            pos = end
            continue
        first = len(out)
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not flags >> bit & 1:           # литерал
                    out.append(data[pos])
                    pos += 1
                    continue
                if pos + 2 > end:
                    raise FormatError("сжатый исходник: оборван токен")
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                done = len(out) - first
                bits = max((done - 1).bit_length(), 4)
                length = (token & (0xFFFF >> bits)) + 3
                offset = (token >> (16 - bits)) + 1
                if offset > done:
                    raise FormatError("сжатый исходник: ссылка за началом "
                                      "куска")
                for _ in range(length):            # возможен нахлёст
                    out.append(out[-offset])
        if len(out) - first > OVBA_CHUNK:
            raise FormatError("кусок распаковался больше чем в 4096 байт")
    return bytes(out)


def ovba_pack_chunk(chunk):
    """Токены одного куска (до 4096 байт): флаг-байт и до 8 токенов за ним.

    Токен — литерал (1 байт) или ссылка на уже выданное в этом куске (2
    байта): число бит под смещение растёт с позицией (MS-OVBA 2.4.1).
    Берётся самое длинное из ближайших совпадений.
    """
    out, seen, pos, size = bytearray(), {}, 0, len(chunk)
    while pos < size:
        flags, tokens = 0, bytearray()
        for bit in range(8):
            if pos >= size:
                break
            bits = max((pos - 1).bit_length(), 4)
            limit = min((0xFFFF >> bits) + 3, size - pos)
            best, offset = 0, 0
            if pos + 3 <= size:
                for cand in seen.get(chunk[pos:pos + 3], ())[
                        :-OVBA_CANDIDATES - 1:-1]:
                    length = 3
                    while length < limit \
                            and chunk[cand + length] == chunk[pos + length]:
                        length += 1
                    if length > best:
                        best, offset = length, pos - cand
                        if length == limit:
                            break
            if best >= 3:
                flags |= 1 << bit
                tokens += struct.pack("<H", ((offset - 1) << (16 - bits))
                                      | (best - 3))
                step = best
            else:
                tokens.append(chunk[pos])
                step = 1
            for k in range(pos, min(pos + step, size - 2)):
                seen.setdefault(chunk[k:k + 3], []).append(k)
            pos += step
        out.append(flags)
        out += tokens
    return bytes(out)


def ovba_compress(data):
    """Байты → сжатый контейнер MS-OVBA (2.4.1).

    Кусок, чьи токены не влезают в 4096 байт, пишется сырым: ровно 4096
    байт, добитых нулями. У текста модуля такого не бывает (он сжимается),
    но распаковка после сжатия всё равно сверяется — см. sync_book.
    """
    out = bytearray(b"\x01")
    for at in range(0, len(data), OVBA_CHUNK):
        chunk = data[at:at + OVBA_CHUNK]
        packed = ovba_pack_chunk(chunk)
        if len(packed) > OVBA_CHUNK:
            out += struct.pack("<H", 0x3FFF) + chunk.ljust(OVBA_CHUNK, b"\0")
        else:
            out += struct.pack("<H", 0xB000 | (len(packed) - 1)) + packed
    return bytes(out)


def ovba_records(raw):
    """Записи потока dir (MS-OVBA 2.3.4.2): [(Id, смещение данных, длина)].

    Запись — Id (2 Б), Size (4 Б), затем Size байт. Исключение одно —
    PROJECTVERSION (Id 0x0009): после Size = 4 идут ещё 6 байт версии.
    """
    out, pos = [], 0
    while pos < len(raw):
        rid, size = struct.unpack_from("<HI", raw, pos)
        size = 6 if rid == 0x0009 else size
        if pos + 6 + size > len(raw):
            raise FormatError("VBA/dir: запись %#06x оборвана на смещении %d"
                              % (rid, pos))
        out.append((rid, pos + 6, size))
        pos += 6 + size
    return out


def ovba_modules(raw):
    """Модули проекта из VBA/dir: [{name, stream, offset, offset_at}].

    offset — где в потоке модуля начинается сжатый исходник (MODULEOFFSET,
    до него — кэш компиляции), offset_at — где это число лежит в raw.
    """
    codepage, modules, cur = "cp1251", [], None
    for rid, at, size in ovba_records(raw):
        data = raw[at:at + size]
        if rid == 0x0003 and size == 2:            # PROJECTCODEPAGE
            codepage = "cp%d" % struct.unpack("<H", data)[0]
        elif rid == 0x0019:                        # MODULENAME
            cur = {"name": data.decode(codepage), "stream": None,
                   "offset": None, "offset_at": None}
            modules.append(cur)
        elif cur is None:
            continue
        elif rid == 0x0047:                        # MODULENAMEUNICODE
            cur["name"] = data.decode("utf-16-le")
        elif rid == 0x001A:                        # MODULESTREAMNAME
            if cur["stream"] is None:
                cur["stream"] = data.decode(codepage)
        elif rid == 0x0032:                        # …UNICODE
            cur["stream"] = data.decode("utf-16-le")
        elif rid == 0x0031 and size == 4:          # MODULEOFFSET
            cur["offset"] = struct.unpack("<I", data)[0]
            cur["offset_at"] = at
    for m in modules:
        if m["stream"] is None or m["offset"] is None:
            raise FormatError("VBA/dir: у модуля %s нет потока или смещения"
                              % m["name"])
    return modules


class Project:
    """Разобранный vbaProject.bin: потоки контейнера и модули из VBA/dir."""

    def __init__(self, blob):
        try:
            self.streams = cfb_read(blob)
            if "VBA/dir" not in self.streams:
                raise FormatError("в контейнере нет потока VBA/dir")
            self.dir = ovba_decompress(self.streams["VBA/dir"])
            self.modules = ovba_modules(self.dir)
        except (struct.error, IndexError, ValueError, RecursionError) as exc:
            raise FormatError("vbaProject.bin повреждён: %s: %s"
                              % (type(exc).__name__, exc))

    def find(self, name):
        return next((m for m in self.modules
                     if m["name"].upper() == name.upper()), None)

    def stream_path(self, module):
        want = ("VBA/" + module["stream"]).upper()
        for path in self.streams:
            if path.upper() == want:
                return path
        raise FormatError("нет потока VBA/%s модуля %s"
                          % (module["stream"], module["name"]))

    def payload(self, module):
        """Сжатый исходник модуля — поток без кэша компиляции в начале."""
        blob = self.streams[self.stream_path(module)]
        if module["offset"] > len(blob):
            raise FormatError("MODULEOFFSET модуля %s дальше конца потока"
                              % module["name"])
        return blob[module["offset"]:]

    def source(self, module):
        try:
            return ovba_decompress(self.payload(module))
        except (struct.error, IndexError, ValueError) as exc:
            raise FormatError("исходник модуля %s: %s" % (module["name"], exc))


def sync_book(book_path):
    """Module1.bas → xl/vbaProject.bin книги. True — книга изменена.

    Книга меняется, только если Module1 в ней расходится с Module1.bas
    (хвостовые переводы строк не в счёт). Что делается тогда — и почему:
      * Module1 заново сжимается из байтов Module1.bas как есть (MS-OVBA
        2.4.1), распаковка обратно обязана вернуть те же байты;
      * кэш компиляции убирается целиком: начало потоков модулей (в VBA/dir
        MODULEOFFSET = 0, остаётся один сжатый исходник), потоки __SRP_n и
        хвост _VBA_PROJECT (остаётся заголовок с Version = 0xFFFF — «кэша
        нет», MS-OVBA 2.3.4.1). Кэш собран Excel из прежнего текста; будь он
        оставлен, рядом с новым исходником лежал бы код старого. Excel
        соберёт кэш заново при открытии книги;
      * остальные модули (листы, ЭтаКнига, Module2), PROJECT и PROJECTwm
        переносятся байт в байт — сжатые исходники ровно те, что записал
        Excel;
      * контейнер пишется заново (MS-CFB v3), прочие части архива и
        метаданные его записей (метод, версии, атрибуты, время) не меняются.
    Запись атомарная: сначала соседний временный файл, затем замена.
    """
    want = BAS.read_bytes()
    name = bas_module_name()
    with zipfile.ZipFile(str(book_path)) as zf:
        if VBA_PROJECT not in zf.namelist():
            raise FormatError("в книге нет %s — макросов нет" % VBA_PROJECT)
        old_blob = zf.read(VBA_PROJECT)
        old = Project(old_blob)
    module = old.find(name)
    if module is None:
        raise FormatError("в книге нет модуля %s: один раз импортируйте "
                          "%s в VBE (File → Import File), дальше --sync "
                          "обновляет его сам" % (name, BAS.name))
    have = old.source(module)
    if have.rstrip(b"\r\n") == want.rstrip(b"\r\n"):
        print("sync: %s в книге уже совпадает с %s — книга не тронута"
              % (name, BAS.name))
        return False

    raw_dir = bytearray(old.dir)
    for m in old.modules:
        struct.pack_into("<I", raw_dir, m["offset_at"], 0)
    new = {path: blob for path, blob in old.streams.items()
           if not SRP_STREAM.match(path)}
    new["VBA/dir"] = ovba_compress(bytes(raw_dir))
    new["VBA/_VBA_PROJECT"] = PURGED_VBA_PROJECT
    for m in old.modules:
        if m is module:
            new[old.stream_path(m)] = ovba_compress(want)
        else:
            old.source(m)                  # сжатый исходник цел, иначе ошибка
            new[old.stream_path(m)] = old.payload(m)
    blob = cfb_write(new)

    # Сверка до записи: новый контейнер читается, Module1 — точно .bas,
    # остальные модули и прочие потоки — те же, что были.
    check_new = Project(blob)
    if check_new.source(check_new.find(name)) != want:
        raise FormatError("Module1.bas после сжатия и распаковки изменился")
    for m in old.modules:
        if m is not module and check_new.source(check_new.find(m["name"])) \
                != old.source(m):
            raise FormatError("модуль %s изменился при пересборке" % m["name"])
    for path, blob_old in old.streams.items():
        if path in ("PROJECT", "PROJECTwm") \
                and check_new.streams.get(path) != blob_old:
            raise FormatError("поток %s изменился при пересборке" % path)

    tmp = book_path.with_name(book_path.name + ".tmp")
    try:
        with zipfile.ZipFile(str(book_path)) as src, \
                zipfile.ZipFile(str(tmp), "w") as dst:
            dst.comment = src.comment
            for info in src.infolist():
                copy = zipfile.ZipInfo(info.filename, info.date_time)
                for attr in ("compress_type", "create_system",
                             "create_version", "extract_version",
                             "external_attr", "internal_attr", "comment"):
                    setattr(copy, attr, getattr(info, attr))
                dst.writestr(copy, blob if info.filename == VBA_PROJECT
                             else src.read(info), compresslevel=9)
        os.replace(str(tmp), str(book_path))
    finally:
        if tmp.exists():
            tmp.unlink()

    lines = have.rstrip(b"\r\n").count(b"\r\n") + 1
    print("sync: %s в книге (%d строк) заменён на %s (%d строк)"
          % (name, lines, BAS.name, want.rstrip(b"\r\n").count(b"\r\n") + 1))
    print("sync: vbaProject.bin %d → %d байт, кэш компиляции сброшен — "
          "Excel соберёт его при открытии книги" % (len(old_blob), len(blob)))
    return True


def codec_selftest():
    """MS-CFB и MS-OVBA: запись → чтение без потерь. Список расхождений."""
    bad = []
    rnd = random.Random(20260101)
    noise = bytes(rnd.randrange(256) for _ in range(OVBA_CHUNK + 900))
    code = ("Private Sub Дело(ByVal x As Long)\r\n    MsgBox \"x\"\r\n"
            "End Sub\r\n" * 70).encode("cp1251")
    for title, data in (("пусто", b""), ("один байт", b"a"),
                        ("повтор на кусок", b"a" * OVBA_CHUNK),
                        ("повтор на кусок и байт", b"a" * (OVBA_CHUNK + 1)),
                        ("шум: сырой кусок и хвост", noise),
                        ("текст в несколько кусков", code)):
        try:
            if ovba_decompress(ovba_compress(data)) != data:
                bad.append("MS-OVBA: %s — после сжатия и распаковки иначе"
                           % title)
        except FormatError as exc:
            bad.append("MS-OVBA: %s — %s" % (title, exc))
    streams = {"Первый": b"x" * 10, "Пустой": b"",
               "VBA/мини": noise[:CFB_CUTOFF - 1],
               "VBA/граница": noise[:CFB_CUTOFF],
               "VBA/большой": noise * 40, "VBA/Вложенный/ещё": b"yz" * 99}
    for k in range(40):                # много братьев — глубокое дерево
        streams["S%02d" % k] = bytes([k]) * (k * 37)
    try:
        if cfb_read(cfb_write(streams)) != streams:
            bad.append("MS-CFB: прочитано не то, что записано")
    except FormatError as exc:
        bad.append("MS-CFB: %s" % exc)
    return bad


def bas_module_name():
    """VB_Name из Module1.bas (если атрибута нет — Module1)."""
    m = re.search(r'(?m)^Attribute VB_Name = "([^"]+)"', bas_text())
    return m.group(1) if m else "Module1"


def source_diff(have, want, limit=10):
    """(число изменённых строк, первые строки unified diff): книга → .bas."""
    def lines(raw):
        return raw.rstrip(b"\r\n").decode("cp1251", "replace").split("\r\n")
    diff = list(difflib.unified_diff(lines(have), lines(want), "книга",
                                     BAS.name, n=0, lineterm=""))
    changed = sum(1 for d in diff
                  if d[:1] in "+-" and d[:3] not in ("+++", "---"))
    return changed, diff[:limit]


def check_vba_project(book):
    """Module1 в книге (xl/vbaProject.bin) == Module1.bas."""
    try:
        blob = book.zip.read(VBA_PROJECT)
    except KeyError:
        check(False, "в книге есть %s" % VBA_PROJECT,
              "макросов в книге нет — Module1 сверять не с чем")
        return
    try:
        project = Project(blob)
        sources = {m["name"]: project.source(m) for m in project.modules}
    except FormatError as exc:
        check(False, "%s читается как MS-CFB и MS-OVBA" % VBA_PROJECT,
              str(exc))
        return
    check(True, "%s читается как MS-CFB и MS-OVBA: модулей %d (%s)"
          % (VBA_PROJECT, len(project.modules),
             ", ".join(m["name"] for m in project.modules)))
    cached = len(project.streams.get("VBA/_VBA_PROJECT", b"")) \
        > len(PURGED_VBA_PROJECT)
    print("     " + ("кэш компиляции Excel есть" if cached else
                     "кэша компиляции нет — Excel соберёт его при открытии "
                     "книги"))
    name = bas_module_name()
    module = project.find(name)
    if module is None:
        check(False, "в книге есть модуль %s" % name,
              "%s сверять не с чем: --sync не поможет, пока модуль не "
              "импортирован в VBE (File → Import File)" % BAS.name)
        return
    have, want = sources[module["name"]], BAS.read_bytes()
    same = have.rstrip(b"\r\n") == want.rstrip(b"\r\n")
    changed, diff = source_diff(have, want)
    lines = want.rstrip(b"\r\n").count(b"\r\n") + 1
    check(same, "%s в книге == %s (%d строк)"
          % (module["name"], BAS.name, lines),
          "строк расходится: %d; книга → %s:\n         %s\n       "
          "положить %s в книгу: python3 tools/check-example.py --sync"
          % (changed, BAS.name, "\n         ".join(diff), BAS.name))


def main(argv):
    paths = [a for a in argv[1:] if a != "--sync"]
    if len(paths) > 1 or any(a.startswith("-") for a in paths):
        print("использование: check-example.py [--sync] [книга.xlsm]")
        return 2
    book_path = Path(paths[0]) if paths else BOOK
    print("== check-example: %s ==" % book_path.name)
    if "--sync" in argv[1:]:
        try:
            sync_book(book_path)
        except (FormatError, OSError, zipfile.BadZipFile) as exc:
            print("FAIL sync: %s" % exc)
            return 2
    book = Book(book_path)

    print("-- vbaProject.bin: Module1 в книге ↔ Module1.bas")
    bad = codec_selftest()
    check(not bad, "самопроверка чтения и записи MS-CFB / MS-OVBA",
          "; ".join(bad))
    check_vba_project(book)
    missing = [s for s in (SHEET_DATA, SHEET_SERVICE) if s not in book.paths]
    check(not missing, "листы «%s» и «%s» есть" % (SHEET_DATA, SHEET_SERVICE),
          "нет: %s" % ", ".join(missing))
    if missing:
        return 1

    data = book.cells(SHEET_DATA)
    svc = book.cells(SHEET_SERVICE)
    rows = service_rows(svc)

    print("-- README «Значения примера» ↔ лист «1»")
    check_readme_values(data)
    print("-- наименования строк defRows ↔ Service (ширина)")
    check_names_in_service(data, rows)
    print("-- каналы Service!F ↔ строки defRows (как читает макрос)")
    check_channels(data, svc)
    print("-- Service: ширины без противоречий")
    check_width_conflicts(rows)
    print("-- формулы листа «1»: D26 обещана README, D21 справочно")
    formulas = book.formulas(SHEET_DATA)
    f26 = formulas.get("D26")
    # D26 макрос не пишет (он пишет только ADDR_CAB_COUNT = D21), поэтому
    # обещание README «D26 — формула ROUNDUP по AE10 и Service» проверяемо.
    check(bool(f26), "README: `D26` — формула",
          "в книге число %s — формула потеряна (макрос её не трогает)"
          % as_text(data.get("D26")))
    if f26 == SHARED_FORMULA:
        # Общая формула хранится один раз — в ведущей ячейке группы.
        # Содержимое этой ячейки по XML не восстановить, а заглушка не
        # содержит ни ROUNDUP, ни AE10: проверка выдала бы ложное
        # «формула потеряна». Говорим прямо, что сверить нечем.
        print("     D26: общая формула (t=\"shared\") — текст хранится в"
              " ведущей ячейке группы, содержимое проверить вручную")
    elif f26:
        missing = [w for w in ("ROUNDUP", "AE10", "Service")
                   if w.upper() not in f26.upper()]
        check(not missing, "формула D26 считает по ROUNDUP, AE10 и Service",
              "в формуле нет: %s — %s" % (", ".join(missing), f26))
        print("     D26: %s" % f26)
    f21 = formulas.get("D21")
    if f21:
        print("     D21: формула %s — значение %s; макрос заменит её числом"
              % (f21, as_text(data.get("D21"))))
    else:
        print("     D21: число, формулы нет — значение %s; так и оставляет"
              " ЗаписатьЧислоШкафов, от D21 считает D45"
              % as_text(data.get("D21")))

    print("\nпроверок: %d, расхождений: %d" % (CHECKS, len(FAILS)))
    if FAILS:
        print("FAIL check-example")
        return 1
    print("check-example: пример сходится с Module1.bas, README и Service")
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
