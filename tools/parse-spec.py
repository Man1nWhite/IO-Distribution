#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Исполнимая спецификация разбора чисел из текста (`ЧислоИзТекста`).

Зеркало `Module1.bas` (функции `ЧислоИзТекста`, `ЕстьГруппаРазрядов`,
`ИзвлечьЧисло`, `КолИзЗначения`), перенесённое строка в строку: тот же
порядок накопления цифр, те же границы групп, локаль не участвует.
`КолИзЗначения` — количество столбца D (`ЧитатьКол` только читает ячейку):
тот же разбор плюс отрицательное → 0. `РезервДоля` — правило D19 (проценты
или доля, ограничение 90 %): оно намеренно сохранено как есть, и векторы
закрепляют в том числе ловушку «1 — это 100 %, а не 1 %». Это НЕ тест VBA (Excel/VBE
здесь нет), а канон поведения, зафиксированный таблицей векторов: правка слоя
чтения ячеек не должна уезжать незаметно.

МЕНЯЕШЬ ПРАВИЛА В VBA — МЕНЯЙ ЭТОТ ФАЙЛ. Алгоритм переводится вручную, но
векторы продублированы в VBA (`ТестРазбораЧисел`, запуск в Excel), и их
совпадение сверяет check-sync.py: если зеркало разошлось с настоящим VBA,
это покажет самопроверка в Excel.

Ветку `IsNumeric` → `CDbl`, которой потребители VBA опережают разбор, зеркало
не покрывает — она зависит от системной локали (`izvlech_chislo` приближает её
регуляркой `_FLOATLIKE` + `float()`: в en-US `IsNumeric("1,234")` = True, а
`float("1,234")` — ValueError). Голый `float()` на эту роль не годится: он
принимает "nan", "inf" и "1_000", а `IsNumeric` — нет; на "nan"/"inf" зеркало
падало в `int(d)`, на "1_000" возвращало 1000 вместо 1. Форма с экспонентой
«1D5» (VBA пишет Double и так) не покрыта — как и локаль, это граница
приближения. Булевы значения (`TRUE`/`FALSE` в ячейке) обе стороны трактуют
как 0 — в VBA `ИзвлечьЧисло` отсекает их до `IsNumeric`.

Код возврата: 0 — все векторы совпали, 1 — есть расхождение.
"""

import re
import sys

MAX_LONG = 2147483647

# Приближение IsNumeric для текста: только десятичная запись (знак, цифры,
# дробная часть, экспонента) и пробелы по краям. float() шире: "nan", "inf",
# "1_000" он берёт, а VBA — нет, и такие значения уходят в ЧислоИзТекста.
_FLOATLIKE = re.compile(r"^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$")


def is_digit(c):
    return '0' <= c <= '9'


def group_at(s, i, n, int_digits, groups):
    """Пробел s[i] разделяет группы разрядов (индексы 0-based)."""
    if groups == 0 and not (1 <= int_digits <= 3):
        return False
    if i + 4 > n:
        return False
    for k in (1, 2, 3):
        if not is_digit(s[i + k]):
            return False
    if i + 4 < n and is_digit(s[i + 4]):
        return False
    return True


def chislo_iz_teksta(s):
    """Зеркало Function ЧислоИзТекста(ByVal s As String) As Double."""
    for ch, to in (('\u00a0', ' '), ('\u202f', ' '), ('\u2009', ' '),
                   ('\u200a', ' '), ('\u2212', '-')):
        s = s.replace(ch, to)

    n = len(s)
    if n == 0:
        return 0.0

    i = 0
    neg = False
    while i < n:
        c = s[i]
        if is_digit(c):
            break
        if c in '+-' and (i == 0 or s[i - 1] == ' ') and i + 1 < n and is_digit(s[i + 1]):
            neg = (c == '-')
            i += 1
            break
        i += 1
    if i >= n:
        return 0.0

    d = 0.0
    int_digits = 0
    groups = 0
    dec_seen = False

    while i < n:
        c = s[i]
        if is_digit(c):
            d = d * 10.0 + float(ord(c) - 48)
            int_digits += 1
            i += 1
        elif c == ' ':
            if not group_at(s, i, n, int_digits, groups):
                break
            i += 1
            for _ in range(3):
                d = d * 10.0 + float(ord(s[i]) - 48)
                i += 1
            int_digits = 0
            groups += 1
        elif c in ',.':
            dec_seen = True
            i += 1
            break
        else:
            break

    if dec_seen:
        scale = 10.0
        while i < n:
            c = s[i]
            if not is_digit(c):
                break
            d = d + float(ord(c) - 48) / scale
            scale *= 10.0
            i += 1

    return -d if neg else d


def izvlech_chislo(v):
    """Зеркало Function ИзвлечьЧисло(ByVal v As Variant) As Long."""
    # Empty/ошибка и булево — 0 (в VBA булево отсекается vbBoolean до IsNumeric).
    if v is None or isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        d = float(v)
    elif isinstance(v, str) and _FLOATLIKE.match(v):
        d = float(v)
    else:
        d = chislo_iz_teksta(str(v))

    d = float(int(d))          # Fix: отбрасывание, а не округление
    d = min(d, float(MAX_LONG))
    d = max(d, -float(MAX_LONG))
    return int(d)


def kol_iz_znacheniya(v):
    """Зеркало Function КолИзЗначения(ByVal v As Variant) As Long.

    Количество столбца D: разбор — ИзвлечьЧисло, отрицательное → 0 (строка
    не участвует в расчёте). ЧитатьКол только читает ячейку и зовёт её,
    поэтому векторы гоняются здесь и в Excel без обращения к листу.
    """
    n = izvlech_chislo(v)
    return n if n > 0 else 0


def rezerv_dolya(v):
    """Зеркало Function РезервДоля(v, wasText, clamped) As Double.

    Возвращает (доля, был_текст, ограничен_90). Правило D19 намеренно
    сохранено, а не «исправлено»: число > 1 — проценты, <= 1 — доля, поэтому
    1 (и числом, и текстом «1 %») означает 100 % и даёт 0.9 после
    ограничения. Отрицательное → 0, всё >= 90 % → 0.9.
    """
    fv, was_text = 0.0, False
    if isinstance(v, bool):
        fv = -1.0 if v else 0.0      # IsNumeric(TRUE) = True, CDbl(TRUE) = -1
    elif isinstance(v, (int, float)):
        fv = float(v)
    elif isinstance(v, str):
        if _FLOATLIKE.match(v):
            fv = float(v)
        else:
            t = v.strip(" ")         # Trim$ — только пробелы, не NBSP
            was_text = len(t) > 0
            if was_text:
                fv = chislo_iz_teksta(t)
    frac = fv / 100.0 if fv > 1 else fv
    if frac < 0:
        frac = 0.0
    clamped = frac >= 0.9
    return (0.9 if clamped else frac), was_text, clamped


# Векторы: (исходная запись в ячейке, ожидаемый Double из ЧислоИзТекста).
# Копия VECTORS и WRAP — в ТестРазбораЧисел (Module1.bas); сверяет check-sync.py.
VECTORS = [
    ("2000", 2000.0),
    ("1 500", 1500.0),
    ("1 500 мм", 1500.0),
    ("-500", -500.0),
    ("\u2212500", -500.0),           # минус U+2212
    ("2000,5", 2000.5),
    ("2000.5", 2000.5),
    ("5 шт 160", 5.0),               # после первого числа текст не читается
    ("4", 4.0),
    ("D18 = 4", 18.0),               # цифры в наименовании — не значение
    ("нет", 0.0),
    ("", 0.0),
    ("1 500 000", 1500000.0),        # цепочка групп
    ("1  500", 1.0),                 # два пробела — не разделитель
    ("1,234", 1.234),                # запятая всегда десятичная
    ("12 345", 12345.0),
    ("1 000", 1000.0),
    ("2000,", 2000.0),               # висячий разделитель
    ("LPA-340-100", 340.0),          # дефис в имени — не знак
    ("1.234.567", 1.234),            # вторая точка — не число
    ("0,9", 0.9),
    ("5-3", 5.0),
    ("12345 678", 12345.0),          # первая группа длиннее 3 цифр
    ("\u00a01\u00a0500", 1500.0),    # неразрывные пробелы
    ("+6", 6.0),
    ("20 %", 20.0),                  # D19 текстом: >1 — проценты, затем /100
    ("резерв 15%", 15.0),
]

# Векторы целочисленной обёртки: (запись, ожидаемый Long)
WRAP = [
    ("2000", 2000),
    ("-500", -500),
    ("2000,5", 2000),                # Fix, не округление
    ("3,6", 3),
    ("99999999999", MAX_LONG),       # без ошибки Overflow
    ("-99999999999", -MAX_LONG),
    ("1 500", 1500),
    ("нет", 0),
    ("", 0),
    ("nan", 0),                      # float() берёт, IsNumeric — нет
    ("1_000", 1),                    # Python-разделитель разрядов, не число
]

# Векторы количества столбца D: (запись, ожидаемый Long). Отрицательное — 0.
# Все записи без десятичного разделителя: векторы не зависят от локали.
KOL = [
    ("4", 4),
    ("0", 0),
    ("-1", 0),                       # отрицательное количество — не количество
    ("-500", 0),
    ("нет", 0),
    ("", 0),
    ("5 шт", 5),                     # текст: первая группа цифр
    ("1_000", 1),
    ("99999999999", MAX_LONG),
]


# Векторы резерва D19: (запись, ожидаемая доля). Все записи целые или текст —
# от системной локали не зависят. «1» и «1 %» закрепляют ловушку: это 100 %,
# а не 1 %; если кто-то поменяет «fv > 1» на «fv >= 1», эти два вектора упадут.
REZERV = [
    ("1", 0.9),                      # 100 % → ограничение 90 %
    ("0", 0.0),
    ("20", 0.2),                     # проценты
    ("90", 0.9),
    ("91", 0.9),
    ("100", 0.9),
    ("150", 0.9),
    ("-5", 0.0),                     # отрицательное → 0
    ("1 %", 0.9),                    # текстом ловушка не обходится
    ("20 %", 0.2),
    ("нет", 0.0),
    ("", 0.0),
]


def main():
    bad = 0

    for src, want in VECTORS:
        got = chislo_iz_teksta(src)
        if abs(got - want) > 1e-9:
            bad += 1
            print("FAIL ЧислоИзТекста(%-16r) = %-12r ожидалось %r" % (src, got, want))
        else:
            print("ok   ЧислоИзТекста(%-16r) = %r" % (src, got))

    for src, want in WRAP:
        got = izvlech_chislo(src)
        if got != want:
            bad += 1
            print("FAIL ИзвлечьЧисло(%-16r) = %-12d ожидалось %d" % (src, got, want))
        else:
            print("ok   ИзвлечьЧисло(%-16r) = %d" % (src, got))

    for src, want in KOL:
        got = kol_iz_znacheniya(src)
        if got != want:
            bad += 1
            print("FAIL КолИзЗначения(%-14r) = %-12d ожидалось %d" % (src, got, want))
        else:
            print("ok   КолИзЗначения(%-14r) = %d" % (src, got))

    for src, want in REZERV:
        got, _was_text, _clamped = rezerv_dolya(src)
        if abs(got - want) > 1e-9:
            bad += 1
            print("FAIL РезервДоля(%-16r) = %-12r ожидалось %r" % (src, got, want))
        else:
            print("ok   РезервДоля(%-16r) = %r" % (src, got))

    print("\nвекторов: %d, расхождений: %d"
          % (len(VECTORS) + len(WRAP) + len(KOL) + len(REZERV), bad))
    return 1 if bad else 0


if __name__ == '__main__':
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main())
