#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
vba-lint.py — статическая проверка VBA-модуля (Module1.bas) без Excel/VBE.

Запускается на Python 3, внешних зависимостей нет.

Проверки (ошибки = ненулевой код возврата):
  E01 кодировка: файл должен декодироваться в Windows-1251 и round-trip'иться
  E02 окончания строк: только CRLF, файл обязан заканчиваться переводом строки
  E03 несбалансированные блоки (If/ElseIf/End If, For/Next, Do/Loop, With,
      Select Case/End Select, While/Wend) в пределах процедуры и
      Type/End Type, Enum/End Enum на уровне модуля
  E04 скобки и кавычки в логической строке (с учётом переносов «_»)
  E05 висячее продолжение строки («_» в конце файла или перед комментарием)
  E06 отсутствие Option Explicit
  E07 дублирующиеся имена процедур (в VBA: «Ambiguous name detected»)
  E08 идентификатор не объявлен (при Option Explicit это ошибка компиляции)
  E09 Call <имя> — такой процедуры нет. Только явный Call: голые вызовы
      (обычный стиль этого модуля) покрыты E08 — имя в них проверяется как
      любой идентификатор, — и E21/E22 по сигнатуре
  E10 On Error GoTo <метка> — метка не найдена в этой процедуре
  E11 On Error Resume Next не закрыт до конца процедуры
  E12 функция нигде не присваивает имя себе (вернёт Empty/0/«»)
  E13 дублирующийся номер ошибки в Err.Raise vbObjectError + N
  E14 дубликат метки Case в одном Select Case
  E15 ReDim X(1 To N) с явной нулевой/отрицательной границей (Runtime error 9)
  E16 строка кода внутри процедуры на нулевом отступе (потерян отступ);
      процедуры берутся из разбора модуля, многострочные сигнатуры учтены
  E17 строка начинается с бинарного оператора, а предыдущая не продолжена «_»
      (разорванное выражение — VBA: Compile error: Expected: expression)
  E18 снят: дубль E08 (E08 сверяет каждое имя со списком своей процедуры,
      E18 — с объединением списков всех процедур, то есть строго слабее).
      Номер зарезервирован, чтобы не переиспользовать его в чужих ссылках.
  E19 комментарий кончается на « _» — VBA продолжает комментарий, и следующая
      строка молча становится комментарием
  E20 больше 24 продолжений «_» в одной логической строке
      (VBA: Too many line continuations)
  E21 неверное число аргументов в вызове процедуры модуля — голым вызовом,
      в скобках или через Call; Optional и ParamArray учтены
      (VBA: Argument not optional / Wrong number of arguments)
  E22 именованный аргумент, которого нет в сигнатуре процедуры модуля
      (VBA: Named argument not found)
  E23 тип в «As <Тип>» не известен — ни встроенный, ни объектный, ни
      объявленный в модуле (VBA: User-defined type not defined)

Предупреждения (код возврата не меняют):
  W01 табуляция вместо пробелов
  W02 хвостовые пробелы
  W03 строка длиннее MAX_LINE
  W04 неиспользуемый параметр/локальная переменная
  W05 метка, на которую нет переходов
  W06 Private-объявление уровня модуля нигде не используется (Public/Friend
      может использовать другой модуль проекта — не проверяется)

Использование:
    python3 tools/vba-lint.py [файл.bas ...]
Без аргументов — все *.bas/*.cls репозитория (обход без .git). Код возврата:
1, если есть ошибки E*; предупреждения (W*) его не меняют.

Перед обходом файлов линтер гоняет самопроверку правил (SELFTEST ниже):
каждое правило обязано сработать на своём примере и промолчать на законном
коде. Правило, которое молча не срабатывает, даёт зелёный вывод без
проверки — это ловится здесь, а не в разборе чужого модуля. Расхождение
самопроверки — тоже код возврата 1.

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
# Предел VBA на одну логическую строку («Too many line continuations»).
MAX_CONTINUATIONS = 24

# ---------------------------------------------------------------------------
# Встроенные имена VBA / Excel / Office (не требуют объявления)
# ---------------------------------------------------------------------------
# Только существующие имена: выдуманное имя в списке — это молча выключенная
# проверка E08 для опечатки в коде. Легитимное ложное срабатывание — имя
# добавляется сюда тем же коммитом, что и правка модуля.
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
datediff datepart datevalue day debug dir exp fix format hex int
isempty iserror isnull isnumeric isarray isobject isdate ismissing vartype lcase left
len lbound lcase loc lof ltrim mid minute month monthname now oct right rnd
round sgn second split sqr str strconv tan time timer timeserial
timevalue trim typename ucase ubound val variant weekday year
instr instrrev replace lcase ucase strcomp strconv filter join
thisworkbook workbook workbooks worksheet worksheets range ranges cells shape
shapes sheets columns rows selection activewindow windows application screen
font fill line shadow textframe textframe2 textrange paragraphformat
addshape addtextbox msofadeclear nothing
msgbox inputbox createobject getobject copyfile kill name chdir mkdir rmdir
filedatetime filesize freefile open close print input lock unlock seek eof
beep randomize
err rgb iif chrw ascw curdir formatcurrency formatdatetime
isarray
""".split())

CONSTANTS = set("""
vbnullstring vbcr vblf vbcrlf vbobjecterror vbcritical vbexclamation
vbinformation vbyes vbno vbok vbokcancel vbyesno vbretry vbcancel vbquestion
vbblue vbgreen vbred vbblack vbyellow vbwhite vbmagenta vbcyan vbtabs vbtab
vbformfeed vbnewline vbtrue vbfalse vbboolean vbtextcompare vbbinary vbunicode
vbdirectory vbsystem vbsystemfile vbnormal vbmaximized vbminimized vbhide
vbshow vbfirstrow vbfirstcol vbgeneral vbext_ct_stdmodule
msoshaperectangle msoshapetextbox msotextorientationhorizontal msoalignleft
msoaligncenter msoalignright msoalignjustify msoanchor msoanchormiddle
msoanchorbottom msoanchorbottomright msoanchorbottomcenter msoline msolinedash
msotrue msofalse msoautofit msonormal msoedgeraise3d msofadeclear
xlnone xlup xldown xltoleft xlportrait xlexcelcurrentlocale
xlcalculationautomatic xlcalculationmanual xlcalculationsemiautomatic
""".split())

TYPES = set("""
long integer double single string boolean variant object byte currency
decimal date collection
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


# Продолжение — это «_» в конце строки кода, отделённое пробелом (или
# единственное на строке): `x = 1 + _`. Хвост `foo_` — часть идентификатора,
# а не перенос, поэтому endswith("_") здесь не годится.
CONT_RE = re.compile(r"(?:^|\s)_$")


def is_continued(code):
    """Строка кода (уже без комментария) продолжается на следующей строке."""
    return bool(CONT_RE.search(code.rstrip()))


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
        if is_continued(code) and code.count('"') % 2 == 0:
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


# «Static Sub» и «Public Static Sub» — тоже начало процедуры; тип результата
# бывает массивом («As Long()»), иначе такие процедуры выпадали из разбора.
PROC_START = re.compile(
    r"^\s*((?:(?:public|private|friend|static)\s+)*)(sub|function|property (?:get|let|set))"
    r"\s+([^\s(]+)(\(.*\))?\s*(as\s+\w+\s*(?:\(\s*\))?)?\s*$",
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


def module_level(module):
    """Уровень модуля: объявления, имена Type/Enum и члены Enum.

    Возвращает ({имя: (строка, видимость, имя как записано)}, типы,
    члены_enum). Внутри
    Type/Enum объявления — поля и члены, а не переменные модуля: поле
    доступно только через точку, член Enum — по имени (иначе E08 считал бы
    его необъявленным).
    """
    names, types, members = {}, set(), set()
    inside = None
    for no, text, _span in module:
        low = text.strip().lower()
        m = re.match(r"^(?:public\s+|private\s+)?(type|enum)\s+(%s)" % NAME_RE, low)
        if m:
            inside = m.group(1)
            types.add(m.group(2))
            continue
        if re.match(r"^end (type|enum)\b", low):
            inside = None
            continue
        if inside == "enum":
            m = re.match(r"^(%s)\s*(?:=|$)" % NAME_RE, low)
            if m:
                members.add(m.group(1))
            continue
        if inside:
            continue
        m = re.match(
            r"^(private|public|friend)\s+(?:const\s+|dim\s+|static\s+)?"
            r"(?:withevents\s+)?(%s)" % NAME_RE, text.strip(), re.I)
        if m:
            names[m.group(2).lower()] = (no, m.group(1).lower(), m.group(2))
    return names, types, members


def collect_decls(procs, module):
    """Имена, объявленные на уровне модуля (без полей Type и членов Enum)."""
    return set(module_level(module)[0])


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
        if DECL.match(low):
            body = re.sub(r"^\s*(dim|static|const)\s+(withevents\s+)?(const\s+)?", "", low)
            for part in split_top(body):
                m = re.match(r"(%s)" % NAME_RE, part)
                if m:
                    names.add(m.group(1))
        # For / For Each счётчик обязан быть объявлен, но соберём для точности
    return names


def proc_signature(proc):
    """Параметры процедуры: имена, число обязательных, признак ParamArray.

    Optional/ParamArray снимают требование аргумента, ParamArray — и
    верхнюю границу их числа.
    """
    sig = proc["sig"].strip()
    names, required, variadic = [], 0, False
    if sig.startswith("("):
        inner = sig[1:-1] if sig.endswith(")") else sig[1:]
        for arg in split_top(inner):
            low = arg.lower().strip()
            optional = False
            while True:
                m = re.match(r"^(byval|byref|optional|paramarray)\s+", low)
                if not m:
                    break
                if m.group(1) in ("optional", "paramarray"):
                    optional = True
                if m.group(1) == "paramarray":
                    variadic = True
                low = low[m.end():]
            m = re.match(r"(%s)" % NAME_RE, low)
            if m:
                names.append(m.group(1))
                if not optional:
                    required += 1
    return {"names": names, "required": required, "variadic": variadic}


def match_paren(text, open_idx):
    """Индекс парной «)» для «(» на open_idx (строки уже замаскированы)."""
    depth = 0
    for i in range(open_idx, len(text)):
        if text[i] == "(":
            depth += 1
        elif text[i] == ")":
            depth -= 1
            if depth == 0:
                return i
    return -1


# Начало оператора: начало строки, после двоеточия, после Then/Else.
STMT_HEAD = re.compile(r"(?:^|:|\bthen\b|\belse\b)\s*(?:call\s+)?$", re.I)
# Продолжение выражения, а не аргумент: «Имя = …», «Имя > 0», «Имя And …».
NOT_A_CALL = re.compile(r"^\s*(?:[=<>+\-*/\\^&.,)]|and\b|or\b|xor\b|mod\b"
                        r"|is\b|like\b|to\b|then\b|else\b)", re.I)


def stmt_tail(text, pos):
    """Хвост оператора от pos до двоеточия верхнего уровня.

    «:=» — именованный аргумент, а не разделитель операторов: «Двое a:=1,
    b:=2» — один оператор с двумя аргументами.
    """
    depth = 0
    for i in range(pos, len(text)):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == ":" and depth == 0 and text[i + 1:i + 2] != "=":
            return text[pos:i]
    return text[pos:]


def find_calls(text, signatures):
    """Вызовы процедур модуля в операторе: (имя, [аргументы]).

    Три формы VBA: «Имя a, b» (голый вызов), «Имя(a, b)» (в выражении или
    через Call) и «Имя» без аргументов. Имя функции как переменная
    результата («Имя = …», «Имя > 0») вызовом не считается.
    """
    out = []
    for m in re.finditer(r"(?<![.\w])(%s)" % NAME_RE, text):
        name = m.group(1).lower()
        if name not in signatures:
            continue
        rest = text[m.end():]
        stripped = rest.lstrip()
        if stripped.startswith("("):
            open_idx = m.end() + len(rest) - len(stripped)
            close = match_paren(text, open_idx)
            if close < 0:
                continue                      # разобьёт E04
            # «Имя(x) = 5» целым оператором — присваивание элементу массива,
            # а не вызов. Но «If Имя(x) = key Then» — сравнение, и вызов в нём
            # проверяется: разница только в позиции оператора.
            if (STMT_HEAD.search(text[:m.start()])
                    and text[close + 1:].lstrip().startswith("=")):
                continue
            out.append((name, split_top(text[open_idx + 1:close])))
        elif STMT_HEAD.search(text[:m.start()]) and not NOT_A_CALL.match(rest):
            out.append((name, split_top(stmt_tail(text, m.end()))))
    return out


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

# Строка объявления локальной переменной/константы.
DECL = re.compile(r"^\s*(dim|static|const)\b")


def usage_text(proc):
    """Тело процедуры для поиска использований имени — без объявлений.

    «Dim x As Long: x = 1» — объявление и использование в одной строке:
    после двоеточия начинается отдельный оператор, его считаем.
    """
    out = []
    for _no, text, _span in proc["body"]:
        t = mask_strings(text).lower()
        if DECL.match(t):
            t = t.split(":", 1)[1] if ":" in t else ""
        out.append(t)
    return "\n".join(out)


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
        err("E01", 0, "Windows-1251 round-trip не идентичен (битые символы)")
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

    # E20 — предел VBA: 24 продолжения на логическую строку.
    for no, _t, span in logical:
        if len(span) - 1 > MAX_CONTINUATIONS:
            err("E20", no, "%d продолжений «_» > %d"
                % (len(span) - 1, MAX_CONTINUATIONS))

    # E05
    for i, line in enumerate(raw_lines, start=1):
        code = strip_comment(line).rstrip()
        if is_continued(code):
            nxt = raw_lines[i] if i < len(raw_lines) else ""
            if not strip_comment(nxt).strip():
                err("E05", i, "продолжение «_», а следующая строка пуста/комментарий")

    # E19 — VBA продолжает и комментарий: строка после «' … _» тоже станет
    # комментарием, код в ней молча исчезает.
    for i, line in enumerate(raw_lines, start=1):
        comment = line[len(strip_comment(line)):]
        if comment.startswith("'") and re.search(r"\s_\s*$", comment):
            err("E19", i, "комментарий кончается на « _» — строка %d тоже станет "
                          "комментарием" % (i + 1))

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

    mod_decls, mod_types, enum_members = module_level(module)
    mod_names = set(mod_decls)
    proc_names = {p["name"].lower() for p in procs}
    known = (KEYWORDS | BUILTINS | CONSTANTS | TYPES | mod_names | proc_names
             | mod_types | enum_members)
    signatures = {p["name"].lower(): proc_signature(p) for p in procs}
    proc_label = {p["name"].lower(): p["name"] for p in procs}

    # E03 — Type/End Type и Enum/End Enum: они объявляются на уровне модуля
    # (внутри процедуры VBA их не допускает), поэтому проверяются здесь, а
    # не в стеке процедуры.
    mod_stack = []
    for no, text, _span in module:
        s = text.strip().lower()
        m = re.match(r"^(?:public\s+|private\s+)?(type|enum)\s+%s" % NAME_RE, s)
        if m:
            mod_stack.append((m.group(1), no))
        elif re.match(r"^end (type|enum)\b", s):
            kind = s.split()[1]
            _close(mod_stack, kind, no, "End %s без %s" % (kind, kind),
                   err, {"name": "<уровень модуля>"})
    for kind, no in mod_stack:
        err("E03", no, "%s не закрыт на уровне модуля" % kind.capitalize())

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

        # E21/E22 — вызовы процедур модуля: число аргументов и имена
        # именованных. Сигнатуры известны, поэтому это проверяемо статически
        # — а VBA в CI не запускается, и «Argument not optional» иначе
        # всплывает только при импорте модуля в Excel.
        for no, text, _orig in p["body"]:
            for name, args in find_calls(mask_strings(text), signatures):
                sig = signatures[name]
                positional, named = 0, []
                for arg in args:
                    m = re.match(r"^(%s)\s*:=" % NAME_RE, arg)
                    if m:
                        named.append(m.group(1).lower())
                    else:
                        positional += 1
                for nm in named:
                    if nm not in sig["names"]:
                        err("E22", no, "%s: у %s нет параметра «%s»"
                            % (p["name"], proc_label[name], nm))
                covered = set(sig["names"][:positional]) | set(named)
                missing = [nm for nm in sig["names"][:sig["required"]]
                           if nm not in covered]
                want = ("%d…%d" % (sig["required"], len(sig["names"]))
                        if sig["required"] != len(sig["names"])
                        else str(sig["required"]))
                if sig["variadic"]:
                    want = ">= %d" % sig["required"]
                if missing:
                    err("E21", no, "%s: %s — не хватает аргументов (%s), нужно %s"
                        % (p["name"], proc_label[name], ", ".join(missing), want))
                elif not sig["variadic"] and len(args) > len(sig["names"]):
                    err("E21", no, "%s: %s — аргументов %d, нужно %s"
                        % (p["name"], proc_label[name], len(args), want))

        # E10/E11 — обработчики ошибок и метки. Имя метки бывает кириллицей
        # («On Error GoTo Ошибка»), поэтому NAME_RE, а не [a-z_]\w*.
        labels = labels_in_proc(p)
        open_resume = None
        for no, t, _orig in low_body:
            m = re.match(r"^\s*on error goto (%s)" % NAME_RE, t)
            if m and m.group(1) not in labels:
                err("E10", no, "%s: метка «%s» не определена" % (p["name"], m.group(1)))
            if re.match(r"^\s*on error resume next\s*$", t):
                if open_resume is None:
                    open_resume = no
            elif re.match(r"^\s*on error goto", t) or re.match(r"^\s*on error resume\b", t):
                if open_resume is not None and not re.match(r"^\s*on error resume next", t):
                    open_resume = None
        if open_resume is not None:
            err("E11", open_resume, "%s: On Error Resume Next не восстановлен" % p["name"])

        # W05 — неиспользуемые метки
        used_labels = set()
        for no, t, _orig in low_body:
            for m in re.finditer(r"(?:goto|resume)\s+(%s)" % NAME_RE, t):
                used_labels.add(m.group(1))
        for lb, no in labels.items():
            if lb not in used_labels:
                warn("W05", no, "%s: метка «%s:» без переходов" % (p["name"], lb))

        # E12 — функция без присваивания результата. Присваивание бывает и в
        # однострочном If («If x Then Foo = 1»), и после двоеточия.
        if p["kind"] == "function":
            pat = re.escape(p["name"].lower())
            assign = re.compile(r"(?:^\s*|\bthen\s+|\belse\s+|:\s*)(?:set\s+)?%s\s*="
                                % pat)
            if not any(assign.search(t) for _, t, _ in low_body):
                err("E12", p["start"], "%s: имя функции ни разу не присваивается" % p["name"])

        # E14 — дубликаты Case. Список меток свой на каждый Select Case:
        # вложенный Select не должен сбрасывать список внешнего блока.
        case_stack = []
        for no, t, _orig in low_body:
            s = t.strip()
            if s.startswith("select case"):
                case_stack.append([])
            elif s.startswith("end select"):
                if case_stack:
                    for k, v in collections.Counter(case_stack.pop()).items():
                        if v > 1:
                            err("E14", no, "%s: дубль Case %s (×%d)" % (p["name"], k, v))
            elif s.startswith("case ") and case_stack:
                if not s.startswith("case else"):
                    case_stack[-1].extend(x.strip() for x in s[5:].split(","))

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

    # E17 — разорванная конкатенация/выражение: строка начинается с бинарного
    # оператора, а предыдущая не продолжена «_». Чаще всего рвётся именно
    # перед And/Or/Mod, а не перед «+».
    for i, line in enumerate(raw_lines, start=1):
        s = line.strip()
        if not s or s.startswith("'"):
            continue
        if re.match(r"^(&|\+|=|<>|<=|>=|<|>|\*|/|\\|\^|and\b|or\b|xor\b|mod\b)\s*\S",
                    s, re.I):
            prev = strip_comment(raw_lines[i - 2]).rstrip() if i > 1 else ""
            if not is_continued(prev):
                err("E17", i, "строка начинается с «%s», но предыдущая не "
                              "продолжена знаком _" % s.split()[0])

    # E13 — дубликаты номеров ошибок. Комментарий и строковый литерал — не
    # код: упоминание «vbObjectError + 28» в пояснении дублем не считается.
    nums = collections.OrderedDict()
    for i, line in enumerate(raw_lines, start=1):
        m = re.search(r"vbobjecterror\s*\+\s*(\d+)", mask_strings(strip_comment(line)),
                      re.I)
        if m:
            nums.setdefault(int(m.group(1)), []).append(i)
    for code, where in nums.items():
        if len(where) > 1:
            err("E13", where[1], "код ошибки %d reused на строках %s" % (code, where))

    # E16 — строка кода на нулевом отступе внутри процедуры. Границы берутся
    # из разбора модуля: сигнатура в несколько строк («Private Function X( _»)
    # по физической строке не опознаётся, и раньше тело такой процедуры не
    # проверялось вовсе.
    for p in procs:
        for i in range(p["span"][-1] + 1, p.get("end", p["start"])):
            line = raw_lines[i - 1]
            s = line.strip()
            if (not s or line[0] in " \t" or s.startswith("'")
                    or re.match(r"^%s:$" % NAME_RE, s)):
                continue
            err("E16", i, "код внутри процедуры на нулевом отступе: %s" % s[:60])

    # E23 — тип в «As <Тип>» обязан быть известен: встроенный, объектный или
    # объявленный в этом модуле (Type/Enum). E08 такие имена не видит — он
    # вырезает «As …» целиком, — а опечатка в типе это ошибка компиляции VBA
    # («User-defined type not defined»). Класс из другого модуля проекта
    # добавляется в TYPES тем же коммитом (см. докстринг).
    known_types = TYPES | BUILTINS | mod_types
    scans = [("<уровень модуля>", [(no, t) for no, t, _span in module])]
    for p in procs:
        scans.append((p["name"],
                      [(p["start"], p["sig"])]
                      + [(no, t) for no, t, _span in p["body"]]))
    for label, lines in scans:
        for no, text in lines:
            for m in re.finditer(r"\bas\s+(?:new\s+)?(%s)" % NAME_RE,
                                 mask_strings(text), re.I):
                if m.group(1).lower() not in known_types:
                    err("E23", no, "%s: тип «%s» не известен" % (label, m.group(1)))

    # W06 — неиспользуемые объявления уровня модуля. Только Private:
    # Public/Friend может использовать другой модуль проекта, отсюда этого
    # не видно. Использование в инициализаторе другой константы считается.
    bodies_text = "\n".join(mask_strings(t).lower()
                            for p in procs for _, t, _span in p["body"])
    mod_lines = [(no, mask_strings(t).lower()) for no, t, _span in module]
    for nm, (no, visibility, label) in sorted(mod_decls.items()):
        if visibility != "private" or nm in TYPES:
            continue
        pat = r"(?<![\.\w])%s(?![\w])" % re.escape(nm)
        uses = len(re.findall(pat, bodies_text))
        uses += sum(len(re.findall(pat, t)) for mno, t in mod_lines if mno != no)
        if uses == 0:
            warn("W06", no,
                 "«%s» объявлен на уровне модуля и не используется" % label)

    # W04 — неиспользуемые параметры/локали. Строка объявления — не
    # использование: иначе всякий Dim сам себе оправдание и правило молчит.
    for p in procs:
        flat = usage_text(p)
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


# ---------------------------------------------------------------------------
# Самопроверка правил: правила, которые молча не срабатывают, — худший вид
# поломки линтера (вывод зелёный, проверки нет). Каждый кейс — маленький
# модуль и правило, которое обязано сработать (want) или обязано молчать
# (skip). Гоняется тем же check_file, что и Module1.bas, до обхода файлов.
# ---------------------------------------------------------------------------

SELFTEST = [
    ("E16: многострочная сигнатура, тело на нулевом отступе",
     "Private Function Много( _\n"
     "    ByVal a As Long) As Long\n"
     "Много = a\n"
     "End Function\n", {"E16"}, set()),
    ("E16: отступ на месте — молчит",
     "Private Function Много( _\n"
     "    ByVal a As Long) As Long\n"
     "    Много = a\n"
     "End Function\n", set(), {"E16"}),
    ("W04: локальная переменная объявлена и не используется",
     "Private Sub Дело()\n"
     "    Dim лишняя As Long\n"
     "    Dim нужная As Long\n"
     "    нужная = 1\n"
     "End Sub\n", {"W04"}, set()),
    ("W04: Dim — не использование, но «Dim x: x = 1» — использование",
     "Private Sub Дело()\n"
     "    Dim x As Long: x = 1\n"
     "End Sub\n", set(), {"W04"}),
    ("E17: целочисленное деление «\\» в начале строки",
     "Private Sub Дело()\n"
     "    Dim x As Long\n"
     "    x = 10\n"
     "    \\ 2\n"
     "End Sub\n", {"E17"}, set()),
    ("E17: оборванное условие на And",
     "Private Sub Дело(ByVal x As Long)\n"
     "    If x = 1\n"
     "    And x = 2 Then\n"
     "        MsgBox \"x\"\n"
     "    End If\n"
     "End Sub\n", {"E17"}, set()),
    ("E17: продолжение «_» перед And — законно",
     "Private Sub Дело(ByVal x As Long)\n"
     "    If x = 1 _\n"
     "    And x = 2 Then\n"
     "        MsgBox \"x\"\n"
     "    End If\n"
     "End Sub\n", set(), {"E17"}),
    ("E10: кириллическая метка не найдена",
     "Private Sub Дело()\n"
     "    On Error GoTo Ошибка\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"E10"}, set()),
    ("W05: кириллическая метка с переходом — не лишняя",
     "Private Sub Дело()\n"
     "    On Error GoTo Ошибка\n"
     "    Exit Sub\n"
     "Ошибка:\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E10", "W05"}),
    ("E14: дубль Case во внешнем Select при вложенном",
     "Private Sub Дело(ByVal x As Long)\n"
     "    Select Case x\n"
     "        Case 1\n"
     "            Select Case x\n"
     "                Case 2\n"
     "                    MsgBox \"a\"\n"
     "            End Select\n"
     "        Case 1\n"
     "            MsgBox \"b\"\n"
     "    End Select\n"
     "End Sub\n", {"E14"}, set()),
    ("E12: присваивание в однострочном If",
     "Private Function Флаг(ByVal x As Long) As Boolean\n"
     "    If x > 0 Then Флаг = True\n"
     "End Function\n", set(), {"E12"}),
    ("E13: номер ошибки в комментарии — не дубль",
     "Private Sub Дело()\n"
     "    ' vbObjectError + 28 занят ниже\n"
     "    Err.Raise vbObjectError + 28, , \"x\"\n"
     "End Sub\n", set(), {"E13"}),
    ("E13: настоящий дубль номера ошибки",
     "Private Sub Дело(ByVal x As Long)\n"
     "    If x = 1 Then\n"
     "        Err.Raise vbObjectError + 28, , \"x\"\n"
     "    End If\n"
     "    Err.Raise vbObjectError + 28, , \"y\"\n"
     "End Sub\n", {"E13"}, set()),
    ("E19: комментарий кончается на « _»",
     "Private Sub Дело()\n"
     "    ' пояснение с переносом _\n"
     "    MsgBox \"код, а не комментарий\"\n"
     "End Sub\n", {"E19"}, set()),
    ("E20: 25 продолжений в одной логической строке",
     "Private Sub Дело()\n"
     "    Dim x As Long\n"
     "    x = 1 + _\n" + "        1 + _\n" * 24 + "        1\n"
     "End Sub\n", {"E20"}, set()),
    ("E20: ровно 24 продолжения — предел не превышен",
     "Private Sub Дело()\n"
     "    Dim x As Long\n"
     "    x = 1 + _\n" + "        1 + _\n" * 23 + "        1\n"
     "End Sub\n", set(), {"E20"}),
    ("E08: выдуманное «встроенное» имя больше не прикрывает опечатку",
     "Private Sub Дело()\n"
     "    Sleep 1\n"
     "End Sub\n", {"E08"}, set()),
    ("E21: голый вызов без обязательного аргумента",
     "Private Sub Дело()\n"
     "    Двое 1\n"
     "End Sub\n\n"
     "Private Sub Двое(ByVal a As Long, ByVal b As Long)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"E21"}, set()),
    ("E21: вызов в скобках с лишним аргументом",
     "Private Sub Дело()\n"
     "    MsgBox Один(1, 2)\n"
     "End Sub\n\n"
     "Private Function Один(ByVal a As Long) As Long\n"
     "    Один = a\n"
     "End Function\n", {"E21"}, set()),
    ("E21: вызов внутри сравнения — тоже вызов",
     "Private Sub Дело()\n"
     "    If Двое(1) = 2 Then\n"
     "        MsgBox \"x\"\n"
     "    End If\n"
     "End Sub\n\n"
     "Private Function Двое(ByVal a As Long, ByVal b As Long) As Long\n"
     "    Двое = a\n"
     "End Function\n", {"E21"}, set()),
    ("E21: Optional можно опустить и можно передать",
     "Private Sub Дело()\n"
     "    Трое 1, 2\n"
     "    Трое 1, 2, 3\n"
     "End Sub\n\n"
     "Private Sub Трое(ByVal a As Long, ByVal b As Long, _\n"
     "                 Optional ByVal c As Long = 0)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E21"}),
    ("E21: ParamArray снимает верхнюю границу",
     "Private Sub Дело()\n"
     "    Много\n"
     "    Много 1, 2, 3\n"
     "End Sub\n\n"
     "Private Sub Много(ParamArray vals() As Variant)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E21"}),
    ("E21: имя функции как результат — не вызов",
     "Private Function Флаг(ByVal x As Long) As Long\n"
     "    Флаг = x\n"
     "    If Флаг < 0 Then Флаг = 0\n"
     "    Флаг = Флаг + 1\n"
     "End Function\n", set(), {"E21", "E12"}),
    ("E21: вызов с переносом строки считается целиком",
     "Private Sub Дело()\n"
     "    Двое 1, _\n"
     "        2\n"
     "End Sub\n\n"
     "Private Sub Двое(ByVal a As Long, ByVal b As Long)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E21"}),
    ("E22: именованного аргумента нет в сигнатуре",
     "Private Sub Дело()\n"
     "    Двое a:=1, b:=2, c:=3\n"
     "End Sub\n\n"
     "Private Sub Двое(ByVal a As Long, ByVal b As Long)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"E22"}, set()),
    ("E22: верные именованные аргументы в любом порядке",
     "Private Sub Дело()\n"
     "    Двое b:=2, a:=1\n"
     "End Sub\n\n"
     "Private Sub Двое(ByVal a As Long, ByVal b As Long)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E21", "E22"}),
    ("E03: Type без End Type на уровне модуля",
     "Private Type Точка\n"
     "    x As Long\n\n"
     "Private Sub Дело()\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"E03"}, set()),
    ("E23: опечатка в типе («As Longn»)",
     "Private Sub Дело()\n"
     "    Dim x As Longn\n"
     "    x = 1\n"
     "End Sub\n", {"E23"}, set()),
    ("E23: опечатка в типе после New",
     "Private Sub Дело()\n"
     "    Dim c As New Collectionn\n"
     "    c.Add 1\n"
     "End Sub\n", {"E23"}, set()),
    ("E23: опечатка в типе параметра",
     "Private Sub Дело(ByVal ws As Workheet)\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"E23"}, set()),
    ("E23: известные типы, массив и Type/Enum модуля — молчит",
     "Private Type Точка\n"
     "    x As Long\n"
     "End Type\n\n"
     "Private Enum Цвет\n"
     "    Красный = 1\n"
     "End Enum\n\n"
     "Private Sub Дело(ByVal ws As Worksheet)\n"
     "    Dim a() As Long\n"
     "    Dim p As Точка\n"
     "    Dim c As Цвет\n"
     "    Dim o As New Collection\n"
     "    c = Красный\n"
     "    p.x = 1\n"
     "    o.Add p.x\n"
     "    ReDim a(1 To 2)\n"
     "    a(1) = c\n"
     "    MsgBox ws.Name\n"
     "End Sub\n", set(), {"E23", "E08"}),
    ("W06: Private-константа уровня модуля не используется",
     "Private Const ЛИШНЯЯ As Long = 5\n\n"
     "Private Sub Дело()\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", {"W06"}, set()),
    ("W06: используется в инициализаторе другой константы — молчит",
     "Private Const БАЗА As Long = 5\n"
     "Private Const ПРОИЗВ As Long = БАЗА * 2\n\n"
     "Private Sub Дело()\n"
     "    MsgBox ПРОИЗВ\n"
     "End Sub\n", set(), {"W06"}),
    ("W06: Public виден другим модулям проекта — не проверяется",
     "Public gВнешняя As Long\n\n"
     "Private Sub Дело()\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"W06"}),
    ("E03: закрытые Type и Enum на уровне модуля",
     "Private Type Точка\n"
     "    x As Long\n"
     "End Type\n\n"
     "Private Enum Цвет\n"
     "    Красный = 1\n"
     "    Синий = 2\n"
     "End Enum\n\n"
     "Private Sub Дело()\n"
     "    MsgBox \"x\"\n"
     "End Sub\n", set(), {"E03"}),
]


def selftest():
    """Прогнать SELFTEST. Возвращает список расхождений (пустой — всё ок)."""
    import tempfile

    bad = []
    with tempfile.TemporaryDirectory() as tmp:
        for i, (title, body, want, skip) in enumerate(SELFTEST):
            path = os.path.join(tmp, "case%d.bas" % i)
            with open(path, "wb") as fh:
                fh.write(("Option Explicit\n\n" + body).replace("\n", "\r\n")
                         .encode(ENCODING))
            errors, warnings, _logical = check_file(path)
            got = {rule for rule, _no, _msg in errors + warnings}
            missing = sorted(want - got)
            extra = sorted(skip & got)
            if missing or extra:
                bad.append("%s: не сработало %s; сработало зря %s"
                           % (title, ", ".join(missing) or "—",
                              ", ".join(extra) or "—"))
    # Продолжение — только « _», а не хвост идентификатора; сигнатуры
    # «Static Sub» и «As Long()» обязаны попадать в разбор процедур.
    if not is_continued("x = 1 + _") or is_continued("Dim foo_"):
        bad.append("is_continued: « _» и хвост «foo_» различаются неверно")
    for sig in ("Static Sub Дело()", "Public Static Sub Дело()",
                "Private Function Массив() As Long()"):
        if not PROC_START.match(sig):
            bad.append("PROC_START не узнаёт сигнатуру: %s" % sig)
    return bad


def default_targets(root):
    """Все *.bas/*.cls репозитория. В .git не заходим (os.walk в него и не
    спускается), а .github обходим — подстрока «.git» ловила и его."""
    out = []
    for dp, dirs, files in os.walk(root):
        dirs[:] = sorted(d for d in dirs if d != ".git")
        out.extend(os.path.join(dp, f) for f in files
                   if f.lower().endswith((".bas", ".cls")))
    return sorted(out)


def main(argv):
    targets = argv[1:]
    if not targets:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        targets = default_targets(root)
    bad = selftest()
    print("самопроверка правил: %d кейсов, расхождений %d"
          % (len(SELFTEST), len(bad)))
    for msg in bad:
        print("  FAIL " + msg)
    if bad:
        return 1

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
