#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vba-lint.py — статическая проверка VBA-модуля (Module1.bas) без Excel/VBE.

Запускается на Python 3, внешних зависимостей нет.

Проверки (ошибки = ненулевой код возврата):
  E01 кодировка: файл должен декодироваться в Windows-1251 и round-trip'иться
  E02 окончания строк: только CRLF, файл обязан заканчиваться переводом строки
  E03 несбалансированные блоки (If/ElseIf/End If, For/Next, Do/Loop, With,
      Select Case/End Select, While/Wend, Type/End Type) в пределах процедуры
  E04 скобки и кавычки в логической строке (с учётом переносов «_»)
  E05 висячее продолжение строки («_» в конце файла или перед комментарием)
  E06 отсутствие Option Explicit
  E07 дублирующиеся имена процедур (в VBA: «Ambiguous name detected»)
  E08 идентификатор не объявлен (при Option Explicit это ошибка компиляции)
  E09 вызов несуществующей процедуры
  E10 On Error GoTo <метка> — метка не найдена в этой процедуре
  E11 On Error Resume Next не закрыт до конца процедуры
  E12 функция нигде не присваивает имя себе (вернёт Empty/0/«»)
  E13 дублирующийся номер ошибки в Err.Raise vbObjectError + N
  E14 дубликат метки Case в одном Select Case
  E15 ReDim X(1 To N) с явной нулевой/отрицательной границей (Runtime error 9)
  E16 строка кода внутри процедуры на нулевом отступе (потерян отступ)
  E17 строка начинается с бинарного оператора, а предыдущая не продолжена «_»
  E18 ссылка на несуществующую процедуру (вызов с аргументами или «Голый»
      вызов строки) — ловит и случайно закомментированное определение
      (разорванное выражение — VBA: Compile error: Expected: expression)

Предупреждения (код возврата не меняют):
  W01 табуляция вместо пробелов
  W02 хвостовые пробелы
  W03 строка длиннее MAX_LINE
  W04 неиспользуемый параметр/локальная переменная
  W05 метка, на которую нет переходов

Использование:
    python3 tools/vba-lint.py [файл.bas ...]
Без аргументов — все *.bas/*.cls репозитория (обход без .git). Код возврата:
1, если есть ошибки E*; предупреждения (W*) его не меняют.

E08 опирается на списки встроенных имён VBA/Office ниже (KEYWORDS, BUILTINS,
CONSTANTS, TYPES). При легитимном ложном срабатывании — добавьте имя в нужный
список тем же коммитом, что и правку модуля.
"""

import os
import re
import sys
import collections

MAX_LINE = 100
ENCODING = "cp1251"

# ---------------------------------------------------------------------------
# Встроенные имена VBA / Excel / Office (не требуют объявления)
# ---------------------------------------------------------------------------
KEYWORDS = set("""
addressof alias attribute byref byval call case const declare default delete
dim do each else elseif end enum erase error event exit false for friend
function get global gosub goto if implements in input is let like load lock
lset next not object on optional paramarray preserve private public redim rem
resume raiseevent rset select set static sub then to true type unload variant
wend while with withevents write and or mod xor not is null empty missing
nothing next loop step error basic vba
""".split())

BUILTINS = set("""
abs access application array atn cbool cbyte ccur cdate cdbl cdec cint clng
clnglng clngptr cmdialog cobj cos csng cstr cvar date dateserial dateadd
datediff datepart datevalue day debug dir division exp fix format hex int
isempty iserror isnull isnumeric isarray isobject isdate ismissing lcase left
len lbound lcase loc lof ltrim mid minute month monthname now oct right rnd
round sgn second split sqr str strconv tan time timer timeserial
timevalue trim typename ucase ubound val variant weekday weekyear year
instr instrrev replace lcase ucase strcomp strconv filter join
thisworkbook workbook workbooks worksheet worksheets range ranges cells shape
shapes sheets columns rows selection activewindow windows application screen
font fill line shadow textframe textframe2 textrange paragraphformat
addshape addtextbox msofadeclear nothing
msgbox inputbox createobject getobject copyfile kill name chdir mkdir rmdir
filedatetime filesize freefile open close print input lock unlock seek eof
beep randomize sleep
err rgb iif chrw ascw curdir envar formatcurrency formatdatetime
isarray objecterror createline rangefromnotification
""".split())

CONSTANTS = set("""
vbnullstring vbcr vblf vbcrlf vbobjecterror vbcritical vbexclamation
vbinformation vbyes vbno vbok vbokcancel vbyesno vbretry vbcancel vbquestion
vbblue vbgreen vbred vbblack vbyellow vbwhite vbmagenta vbcyan vbtabs vbtab
vbformfeed vbnewline vbtrue vbfalse vbtextcompare vbbinary vbunicode
vbdirectory vbsystem vbsystemfile vbnormal vbmaximized vbminimized vbhide
vbshow vbfirstrow vbfirstcol vbgeneral vbext_ct_stdmodule
msoshaperectangle msoshapetextbox msotextorientationhorizontal msoalignleft
msoaligncenter msoalignright msoalignjustify msoanchor msoanchormiddle
msoanchorbottom msoanchorbottomright msoanchorbottomcenter msoline msolinedash
msotrue msofalse msoautofit msonormal msoedgeraise3d msofadeclear
xlnone xlup xldown xltoleft xlportrait xlexcelcurrentlocale
""".split())

TYPES = set("""
long integer double single string boolean variant object byte currency
decimal date datevariant autocolor collection
""".split())

# ---------------------------------------------------------------------------
# Утилиты разбора
# ---------------------------------------------------------------------------


def strip_comment(line):
    """Убрать хвостовой комментарий, не трогая строковые литералы."""
    out = []
    instr = False
    for ch in line:
        if ch == '"':
            instr = not instr
            out.append(ch)
        elif ch == "'" and not instr:
            break
        else:
            out.append(ch)
    return "".join(out)


def mask_strings(text):
    """Заменить содержимое строковых литералов заглушкой (для разбора)."""
    return re.sub(r'"(?:[^"]|"")*"', '""', text)


def split_top(text, sep=","):
    """Разбить по разделителю верхнего уровня (вне скобок и строк)."""
    depth = 0
    parts = []
    cur = ""
    instr = False
    for ch in text:
        if ch == '"':
            instr = not instr
        if not instr:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == sep and depth == 0:
                parts.append(cur)
                cur = ""
                continue
        cur += ch
    parts.append(cur)
    return [p.strip() for p in parts if p.strip()]


def build_logical(raw_lines):
    """Склеить строки с продолжением «_». Возвращает [(номер, текст, физ.строки)]."""
    logical = []
    buf = ""
    start = None
    span = []
    for n, line in enumerate(raw_lines, start=1):
        code = strip_comment(line).rstrip()
        if start is None:
            start = n
        span.append(n)
        if code.endswith("_") and code.count('"') % 2 == 0:
            buf += " " + code[:-1].strip()
            continue
        buf += " " + code
        logical.append((start, buf.strip(), list(span)))
        buf = ""
        start = None
        span = []
    if start is not None:
        logical.append((start, buf.strip(), list(span)))
    return logical


PROC_START = re.compile(
    r"^\s*(public |private |friend )?(sub|function|property (?:get|let|set))"
    r"\s+([^\s(]+)(\(.*\))?\s*(as\s+\w+)?\s*$",
    re.I,
)
PROC_END = re.compile(r"^\s*end (sub|function|property)\b", re.I)
NAME_RE = r"[A-Za-z_А-Яа-яЁё][\wА-Яа-яЁё]*"


def parse_procedures(logical):
    """Разбить модуль на процедуры + модульный уровень."""
    procs = []
    module = []
    cur = None
    for no, text, span in logical:
        if cur is None:
            m = PROC_START.match(text)
            if m:
                cur = {
                    "kind": m.group(2).lower().split()[0],
                    "name": m.group(3),
                    "sig": m.group(4) or "",
                    "ret": (m.group(5) or "").strip().lower(),
                    "start": no,
                    "body": [],
                    "span": span,
                }
                continue
            module.append((no, text, span))
            continue
        if PROC_END.match(text):
            cur["end"] = no
            procs.append(cur)
            cur = None
        else:
            cur["body"].append((no, text, span))
    if cur is not None:
        cur["end"] = logical[-1][0] if logical else cur["start"]
        cur["unterminated"] = True
        procs.append(cur)
    return procs, module


def collect_decls(procs, module):
    """Имена, объявленные на уровне модуля."""
    mod_names = set()
    for no, text, _ in module:
        low = text.lower()
        m = re.match(
            r"^\s*(private|public|friend)\s+(?:const|dim|static)?\s*(?:withevents\s+)?"
            r"(%s)" % NAME_RE,
            low,
        )
        if m:
            mod_names.add(m.group(2))
    return mod_names


def declared_in_proc(proc):
    """Локальные объявления + параметры процедуры (в нижнем регистре)."""
    names = set()
    sig = proc["sig"].lower()
    if sig.startswith("("):
        for arg in split_top(sig[1:-1]):
            arg = re.sub(r"^(byval|byref|optional|paramarray)\s+", "", arg).strip()
            while re.match(r"^(byval|byref|optional|paramarray)\s+", arg):
                arg = re.sub(r"^(byval|byref|optional|paramarray)\s+", "", arg).strip()
            m = re.match(r"(%s)" % NAME_RE, arg)
            if m:
                names.add(m.group(1))
    for no, text, _ in proc["body"]:
        low = text.lower()
        if re.match(r"^\s*(dim|static|const)\b", low):
            body = re.sub(r"^\s*(dim|static|const)\s+(withevents\s+)?(const\s+)?", "", low)
            for part in split_top(body):
                m = re.match(r"(%s)" % NAME_RE, part)
                if m:
                    names.add(m.group(1))
        # For / For Each счётчик обязан быть объявлен, но соберём для точности
    return names


def labels_in_proc(proc):
    out = {}
    for no, text, _ in proc["body"]:
        m = re.match(r"^(%s):\s*$" % NAME_RE, text)
        if m:
            out[m.group(1).lower()] = no
    return out


# ---------------------------------------------------------------------------
# Проверки
# ---------------------------------------------------------------------------


def check_file(path):
    errors = []
    warnings = []

    def err(rule, line, msg):
        errors.append((rule, line, msg))

    def warn(rule, line, msg):
        warnings.append((rule, line, msg))

    with open(path, "rb") as fh:
        raw = fh.read()

    # E01/E02
    try:
        text = raw.decode(ENCODING)
    except UnicodeDecodeError as exc:
        err("E01", 0, "файл не декодируется как Windows-1251: %s" % exc)
        return errors, warnings, []
    if text.encode(ENCODING) != raw:
        err("E01", 0, "Windows-1251 round-trih не идентичен (битые символы)")
    if b"\n" not in raw:
        err("E02", 0, "нет переводов строк")
    if re.search(rb"(?<!\r)\n", raw):
        err("E02", 0, "встречаются LF без CR (нужен единый CRLF)")
    if raw and not raw.endswith(b"\r\n"):
        err("E02", 0, "файл не заканчивается CRLF")

    raw_lines = text.splitlines()

    # W01/W02/W03
    for i, line in enumerate(raw_lines, start=1):
        if "\t" in line:
            warn("W01", i, "табуляция вместо пробелов")
        if line.rstrip() != line:
            warn("W02", i, "хвостовые пробелы")
        if len(line) > MAX_LINE:
            warn("W03", i, "строка %d симв. > %d" % (len(line), MAX_LINE))

    logical = build_logical(raw_lines)

    # E05
    for i, line in enumerate(raw_lines, start=1):
        code = strip_comment(line).rstrip()
        if code.endswith("_"):
            nxt = raw_lines[i] if i < len(raw_lines) else ""
            if not strip_comment(nxt).strip():
                err("E05", i, "продолжение «_», а следующая строка пуста/комментарий")

    # E06
    if not any(re.match(r"^\s*option\s+explicit\s*$", l, re.I) for l in raw_lines):
        err("E06", 0, "нет Option Explicit")

    procs, module = parse_procedures(logical)
    if not procs:
        err("E03", 0, "не найдено ни одной процедуры")
        return errors, warnings, logical

    # E07
    dups = [k for k, v in collections.Counter(p["name"].lower() for p in procs).items() if v > 1]
    for d in dups:
        err("E07", 0, "дублирующееся имя процедуры: %s" % d)

    mod_names = collect_decls(procs, module)
    proc_names = {p["name"].lower() for p in procs}
    known = KEYWORDS | BUILTINS | CONSTANTS | TYPES | mod_names | proc_names

    for p in procs:
        low_body = [(no, mask_strings(text).lower(), text) for no, text, _ in p["body"]]
        flat = "\n".join(t for _, t, _ in low_body)

        # E03 — баланс блоков
        stack = []
        for no, t, _orig in low_body:
            s = t.strip()
            if re.match(r"^if\b", s):
                m = re.match(r"^if\b.*?\bthen\b(.*)$", s)
                single = bool(m and m.group(1).strip() and not m.group(1).strip().startswith("else"))
                if not single:
                    stack.append(("if", no))
            elif re.match(r"^end if\b", s):
                _close(stack, "if", no, "End If без If", err, p)
            if re.match(r"^for\b", s) and (" to " in s or re.match(r"^for each\b", s)):
                stack.append(("for", no))
            elif re.match(r"^next\b", s):
                _close(stack, "for", no, "Next без For", err, p)
            if re.match(r"^do\b", s):
                stack.append(("do", no))
            elif re.match(r"^loop\b", s):
                _close(stack, "do", no, "Loop без Do", err, p)
            if re.match(r"^with\b", s) and not s.startswith("withevents"):
                stack.append(("with", no))
            elif re.match(r"^end with\b", s):
                _close(stack, "with", no, "End With без With", err, p)
            if re.match(r"^select case\b", s):
                stack.append(("select", no))
            elif re.match(r"^end select\b", s):
                _close(stack, "select", no, "End Select без Select Case", err, p)
            if re.match(r"^while\b", s):
                stack.append(("while", no))
            elif re.match(r"^wend\b", s):
                _close(stack, "while", no, "Wend без While", err, p)
        for kind, no in stack:
            err("E03", no, "%s не закрыт в %s" % (kind.capitalize(), p["name"]))
        if p.get("unterminated"):
            err("E03", p["start"], "процедура %s не закрыта (нет End %s)" % (p["name"], p["kind"]))

        # E04
        for no, t, orig in low_body:
            if orig.count('"') % 2:
                err("E04", no, "нечётное число кавычек")
            if t.count("(") != t.count(")"):
                err("E04", no, "скобки %d «(» / %d «)»" % (t.count("("), t.count(")")))

        # E08 — объявления/использование
        declared = declared_in_proc(p) | {p["name"].lower()}
        allowed = declared | known | set(labels_in_proc(p))
        for no, t, _orig in low_body:
            scan = re.sub(r"\.\s*%s" % NAME_RE, " ", t)
            scan = re.sub(r"^\s*(public|private|friend)?\s*(sub|function)\s+\S+", "", scan)
            scan = re.sub(r"%s\s*:=" % NAME_RE, " ", scan)
            scan = re.sub(r"\bas\s+%s" % NAME_RE, " ", scan)
            scan = re.sub(r"\bnew\s+%s" % NAME_RE, " ", scan)
            for m in re.finditer(r"(?<![\.\w])(%s)" % NAME_RE, scan):
                nm = m.group(1)
                if nm in allowed:
                    continue
                if nm in mod_names or nm in proc_names:
                    continue
                err("E08", no, "%s: «%s» не объявлен" % (p["name"], nm))

        # E09 — вызовы
        for no, t, _orig in low_body:
            scan = re.sub(r"\.\s*%s" % NAME_RE, " ", t)
            m = re.match(r"^\s*call\s+(%s)" % NAME_RE, scan)
            if m and m.group(1) not in proc_names:
                err("E09", no, "%s: Call %s — процедура не найдена" % (p["name"], m.group(1)))

        # E10/E11 — обработчики ошибок и метки
        labels = labels_in_proc(p)
        open_resume = None
        for no, t, _orig in low_body:
            m = re.match(r"^\s*on error goto ([a-z_]\w*)", t)
            if m and m.group(1) not in labels:
                err("E10", no, "%s: метка «%s» не определена" % (p["name"], m.group(1)))
            if re.match(r"^\s*on error resume next\s*$", t):
                if open_resume is None:
                    open_resume = no
            elif re.match(r"^\s*on error goto", t) or re.match(r"^\s*on error resume\b", t):
                if open_resume is not None and not re.match(r"^\s*on error resume next", t):
                    open_resume = None
            if re.match(r"^\s*(exit (sub|function)|end (sub|function))", t):
                pass
        if open_resume is not None:
            err("E11", open_resume, "%s: On Error Resume Next не восстановлен" % p["name"])

        # W05 — неиспользуемые метки
        used_labels = set()
        for no, t, _orig in low_body:
            for m in re.finditer(r"(?:goto|resume)\s+([a-z_]\w*)", t):
                used_labels.add(m.group(1))
        for lb, no in labels.items():
            if lb not in used_labels:
                warn("W05", no, "%s: метка «%s:» без переходов" % (p["name"], lb))

        # E12 — функция без присваивания результата
        if p["kind"] == "function":
            assigned = False
            for no, t, _orig in low_body:
                if re.match(r"^\s*%s\s*=" % re.escape(p["name"].lower()), t) or re.match(
                    r"^\s*set\s+%s\s*=" % re.escape(p["name"].lower()), t
                ):
                    assigned = True
                    break
            if not assigned:
                err("E12", p["start"], "%s: имя функции ни разу не присваивается" % p["name"])

        # E14 — дубликаты Case
        cases = None
        for no, t, _orig in low_body:
            s = t.strip()
            if s.startswith("select case"):
                cases = []
                continue
            if cases is not None:
                if s.startswith("case "):
                    if not s.startswith("case else"):
                        cases.extend(x.strip() for x in s[5:].split(","))
                elif s.startswith("end select"):
                    for k, v in collections.Counter(cases).items():
                        if v > 1:
                            err("E14", no, "%s: дубль Case %s (×%d)" % (p["name"], k, v))
                    cases = None

        # E15 — ReDim (1 To N) с риском N = 0
        for no, t, _orig in low_body:
            m = re.search(r"redim\s+(?:preserve\s+)?%s\s*\(\s*1\s+to\s*([^)]+)\)" % NAME_RE, t)
            if m:
                expr = m.group(1).strip()
                if re.match(r"^\d+$", expr) and int(expr) < 1:
                    err("E15", no, "%s: ReDim (1 To %s) — пустой массив" % (p["name"], expr))

    # E04b — строковый литерал не может продолжаться на следующей строке
    for i, line in enumerate(raw_lines, start=1):
        code = strip_comment(line)
        if code.count('"') % 2:
            err("E04", i, "нечётное число кавычек в физической строке "
                          "(литерал не может продолжаться через «_»): %s" % code.strip()[:60])

    # E18 — имя в позиции «Имя(...)» обязано быть объявлено (переменная или
    # массив → индексация) либо быть процедурой/встроенной функцией. Ловит в
    # том числе случайно закомментированное или удалённое определение.
    declared = set(mod_names) | proc_names
    for p2 in procs:
        declared |= declared_in_proc(p2)
    for no, text, _ in logical:
        m = re.match(
            r"^(private|public\s+)?(dim|const|static|redim|type|enum)\s+"
            r"(?:withevents\s+)?(%s)" % NAME_RE, text.strip(), re.I)
        if m:
            declared.add(m.group(3).lower())

    for p2 in procs:
        for no, text, _ in p2["body"]:
            scan = mask_strings(text)
            scan = re.sub(r"^\s*(public|private|friend)?\s*(sub|function)\s+\S+",
                          "", scan, flags=re.I)
            scan = re.sub(r"\.\s*%s" % NAME_RE, " ", scan)
            for m in re.finditer(r"(?<![.\w])(%s)\s*\(" % NAME_RE, scan):
                nm = m.group(1).lower()
                if nm in declared or nm in known:
                    continue
                err("E18", no, "%s: «%s(...)» — имя не объявлено и такой "
                               "процедуры нет" % (p2["name"], m.group(1)))

    # E17 — разорванная конкатенация/выражение
    cont_idx = set()
    for start_no, _t, span in logical:
        for x in span[:-1]:
            cont_idx.add(x)
    for i, line in enumerate(raw_lines, start=1):
        s = line.strip()
        if not s or s.startswith("'"):
            continue
        if i in cont_idx or (i > 1 and (i - 1) in cont_idx):
            pass
        if re.match(r"^(&|\+|=|<>|<|>|\*|/|\\\\)\s*\S", s):
            prev = strip_comment(raw_lines[i - 2]).rstrip() if i > 1 else ""
            if not prev.endswith("_"):
                err("E17", i, "строка начинается с «%s», но предыдущая не "
                              "продолжена знаком _" % s.split()[0])

    # E13 — дубликаты номеров ошибок
    nums = collections.OrderedDict()
    for i, line in enumerate(raw_lines, start=1):
        m = re.search(r"vbobjecterror\s*\+\s*(\d+)", line, re.I)
        if m:
            nums.setdefault(int(m.group(1)), []).append(i)
    for code, where in nums.items():
        if len(where) > 1:
            err("E13", where[1], "код ошибки %d reused на строках %s" % (code, where))

    # E16 — строка кода на нулевом отступе внутри процедуры
    in_proc = False
    for i, line in enumerate(raw_lines, start=1):
        s = line.strip()
        if not s:
            continue
        if PROC_START.match(line) and re.match(
            r"^\s*(public|private|friend)\s+(sub|function)\b", line, re.I
        ):
            in_proc = True
            continue
        if PROC_END.match(line):
            in_proc = False
            continue
        if not in_proc:
            continue
        if line[0] in " \t":
            continue
        if s.startswith("'") or re.match(r"^%s:$" % NAME_RE, s):
            continue
        err("E16", i, "код внутри процедуры на нулевом отступе: %s" % s[:60])

    # W04 — неиспользуемые параметры/локали
    for p in procs:
        flat = "\n".join(mask_strings(t).lower() for _, t, _ in p["body"])
        declared = declared_in_proc(p)
        for nm in sorted(declared):
            if nm == p["name"].lower() or nm in TYPES:
                continue
            if len(re.findall(r"(?<![\.\w])%s(?![\w])" % re.escape(nm), flat)) == 0:
                warn("W04", p["start"], "%s: «%s» объявлен и не используется" % (p["name"], nm))

    return errors, warnings, logical


def _close(stack, kind, no, msg, err, proc):
    if stack and stack[-1][0] == kind:
        stack.pop()
    else:
        err("E03", no, "%s в %s" % (msg, proc["name"]))


def main(argv):
    targets = argv[1:]
    if not targets:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        targets = sorted(
            os.path.join(dp, f)
            for dp, _d, fs in os.walk(root)
            if ".git" not in dp
            for f in fs
            if f.lower().endswith((".bas", ".cls"))
        )
    if not targets:
        print("нет файлов .bas/.cls для проверки")
        return 0

    total_err = 0
    total_warn = 0
    for path in targets:
        errors, warnings, logical = check_file(path)
        rel = os.path.relpath(path)
        print("== %s ==" % rel)
        for rule, line, msg in errors:
            print("  ERROR %s %s: %s" % (rule, ("строка %s" % line) if line else "", msg))
        for rule, line, msg in warnings:
            print("  warn  %s %s: %s" % (rule, ("строка %s" % line) if line else "", msg))
        print(
            "  итого: %d ошибок, %d предупреждений (логических строк %d)"
            % (len(errors), len(warnings), len(logical))
        )
        total_err += len(errors)
        total_warn += len(warnings)

    print("\nVBA-LINT: ошибок %d, предупреждений %d" % (total_err, total_warn))
    return 1 if total_err else 0


if __name__ == "__main__":
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv))
