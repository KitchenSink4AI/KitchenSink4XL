"""core/limits.py: the STATIC REFUSE-CLASS checks (the cheap layer of the
"Excel will refuse this" gate).

THE ARCHITECTURAL GAP THIS CLOSES. verify-after-write (core.verify) is an
honest check on three things: the produced package re-opens as valid OOXML,
no part was lost, and the cells read back as intended. It has no notion of
"Excel will refuse this," and it never did: openpyxl will happily load a
package that Excel throws out. Ten single tool calls once returned
``ok: true, saved: true, verified: true`` and produced a file answering
"Open method of Workbooks class failed."

The gate is closed in three layers, and this module is layer (a):

  (a) STATIC REFUSAL CHECKS, here. Every refuse-class found so far gets a
      cheap, Excel-ground-truthed check BEFORE the write, so the common routes
      refuse loudly with an actionable message and never cost a COM call.
  (b) HONEST REPORTING. ops/validation.py no longer hard-codes
      ``structure.opens_clean: true``; the field says "not checked" unless a
      real Excel check ran.
  (c) THE AUTHORITATIVE CHECK, OPT-IN. ``com_validate_opens_clean`` gets every
      one of these right and is wired as a post-write verification
      (``verify_com``, or KS4XL_VERIFY_COM=1 to make it the default).

Layer (a) can never be complete on its own -- that is exactly why layer (c)
exists -- so nothing here claims a file is safe. Each threshold below is a
REFUSAL BOUNDARY MEASURED AGAINST A REAL EXCEL: a package was built at the
boundary, handed to Excel through ``Workbooks.Open``, and the exact value
where the open flips from success to
"Open method of Workbooks class failed" is recorded next to the constant.

No new refusal vocabulary: these raise ExcelWouldRefuse (a plain XlMcpError,
so the envelope's closed BAD_PARAMS code covers it) or FormulaRejected, both
of which already exist.
"""

from __future__ import annotations

import re

from .errors import ExcelWouldRefuse

# ---------------------------------------------------------------- thresholds

#: Comment text. MEASURED: 32,766 and 32,767 characters open; 32,768 does not.
#: (A 60,000-character comment is the same refusal.)
MAX_COMMENT_CHARS = 32_767

#: A data-validation formula1/formula2 string. MEASURED by bisection: 4,097
#: characters open, 4,098 does not. Excel's DOCUMENTED authoring limit is 255
#: and the tool has always warned above it; this is the separate, harder line
#: where the FILE stops opening at all, which is what a
#: 10,000-item inline list crossed.
MAX_DV_FORMULA_CHARS = 4_097
DV_FORMULA_SOFT_CHARS = 255

#: A header or footer, as ONE assembled string including its &L / &C / &R
#: section codes. MEASURED: a single left section of 253 characters opens
#: ("&L" + 253 = 255) and 254 does not; three sections of 80 open (246 total)
#: and three of 84 do not (258). The limit is on the whole string, not the
#: section, which is why a "255 per section" check would still ship a file
#: Excel refuses.
MAX_HEADER_FOOTER_CHARS = 255

#: Picture display extents. MEASURED: width/height 0 and 10 open; -5 does not
#: (a negative EMU extent is not representable in the drawing XML).
MIN_IMAGE_EXTENT = 0

#: Cell text. Excel's documented ceiling, and check_cell_text below is what
#: consults it. It used to be a bare constant that nothing on the set_cells
#: path read, so a 40,000-character cell paid a full write-and-verify cycle
#: and then refused through the WRONG LAYER: "1 written cell did not read
#: back as intended", a generic verify failure, instead of "Excel stores at
#: most 32,767 characters".
MAX_CELL_CHARS = 32_767

#: Formula text, without its leading '='. A 10,000-character formula once
#: saved "verified" and then would not open.
#: MEASURED by bisection on files opened in a hidden Excel: 8,192
#: characters open, 8,193 do not. Excel's documented limit, exactly.
#: Excel counts the formula as the file stores it, with its prefixes
#: (_xlfn. on newer functions, _xlpm. on LET and LAMBDA names): MEASURED,
#: _xlfn.LET(_xlpm.x,A1,_xlpm.x+...) opens at 8,192 stored characters
#: (2,054 as written) and not at 8,193.
MAX_FORMULA_CHARS = 8_192

#: Nested function calls inside one formula. Excel documents 64 as the
#: limit for typing a formula; the FILE limit, which is what decides whether
#: a saved workbook opens, is one higher. MEASURED: 65 nested ABS() open, 66
#: do not.
MAX_FORMULA_NESTING = 65

#: Arguments to one function call. MEASURED: SUM with 255 arguments opens,
#: 256 does not. Excel's documented limit, exactly.
MAX_FUNCTION_ARGS = 255

#: Excel stores every formula compiled, and the compiled form has its own
#: ceiling that a formula well under 8,192 characters can cross: 2,049 cell
#: references joined by "+" (6,146 characters) do not open. MEASURED: the
#: compiled size is 7 bytes per cell reference, 13 per range, 3 per small
#: whole number, 9 per other number, 1 per operator and per closing
#: parenthesis, and 16,383 bytes open while 16,384 do not (2,048 references
#: = 16,383 bytes; the same plus one pair of parentheses = 16,384).
MAX_FORMULA_COMPILED_BYTES = 16_383

#: A font's name. Excel will not open a workbook whose styles name a font
#: longer than this (MEASURED: a 31-character name opens; 32, 33, 40 and 64
#: characters refuse the whole file).
MAX_FONT_NAME_CHARS = 31


def check_font_name(name, *, what: str = "font.name") -> None:
    """Excel's font-name limit, before anything is written."""
    if name is None:
        return
    if not isinstance(name, str):
        raise ExcelWouldRefuse(f"{what} must be a font name such as 'Calibri'")
    if len(name) > MAX_FONT_NAME_CHARS:
        raise ExcelWouldRefuse(
            f"{what} {name!r} is {len(name)} characters long; Excel allows "
            f"{MAX_FONT_NAME_CHARS}, and it will not open a workbook whose "
            "styles name a longer font.")


# ------------------------------------------------------------ illegal text

#: Characters openpyxl itself refuses to serialize (and Excel cannot store):
#: the C0 controls except tab, newline and carriage return. Without this check
#: openpyxl's IllegalCharacterError escaped the refusal envelope entirely
#: and echoed the raw control bytes back to the caller.
_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def check_text_storable(value, *, what: str = "this value"):
    """Refuse text Excel cannot store, BEFORE openpyxl raises its own
    un-enveloped error. Two populations:

      1. C0 control characters. openpyxl raises IllegalCharacterError, which
         escaped the envelope as a bare tool error with a Rich traceback on
         stderr and the raw control bytes echoed back.
      2. LONE SURROGATES. A request carrying "\\ud800" is valid JSON and the
         server accepted it, but the value cannot be encoded to UTF-8 on the
         way out, so the response was never written and the caller blocked
         until its own timeout with the server still healthy. A PAIRED
         surrogate (a real emoji) is fine and is left alone.

    Returns the value unchanged when it is storable."""
    if not isinstance(value, str):
        return value
    m = _ILLEGAL_CHARS.search(value)
    if m is not None:
        point = value.index(m.group(0))
        raise ExcelWouldRefuse(
            f"{what} contains the control character U+{ord(m.group(0)):04X} "
            f"at position {point}, which Excel cannot store in a worksheet. "
            "Strip the control characters (tab, newline and carriage return "
            "are fine) and retry.")
    for i, ch in enumerate(value):
        if 0xD800 <= ord(ch) <= 0xDFFF:
            raise ExcelWouldRefuse(
                f"{what} contains an unpaired surrogate code point "
                f"U+{ord(ch):04X} at position {i}. It is not encodable as "
                "UTF-8, so neither the workbook nor this server's own reply "
                "can carry it. Send the character as a proper pair, or drop "
                "it.")
    return value


# --------------------------------------------------------------- the checks


def check_cell_text(value, *, what: str = "this cell"):
    """The storable-text check plus Excel's 32,767-character cell ceiling.
    Non-strings pass straight through (a number has no length to check).

    openpyxl writes an oversize string happily and the produced package is
    valid OOXML, so verify-after-write is what used to catch it, one layer too
    late and with the wrong words. Refusing here costs nothing and says what
    the limit is."""
    check_text_storable(value, what=what)
    if isinstance(value, str) and len(value) > MAX_CELL_CHARS:
        raise ExcelWouldRefuse(
            f"{what} is {len(value):,} characters; Excel stores at most "
            f"{MAX_CELL_CHARS:,} in a cell and truncates or refuses beyond "
            "that. Shorten the text, or split it across cells.")
    return value


def check_comment_text(text: str) -> str:
    check_text_storable(text, what="the comment text")
    if len(text) > MAX_COMMENT_CHARS:
        raise ExcelWouldRefuse(
            f"the comment is {len(text):,} characters; Excel stores at most "
            f"{MAX_COMMENT_CHARS:,} and REFUSES TO OPEN a workbook that "
            "carries a longer one (measured against Excel: 32,767 opens, "
            "32,768 does not). Shorten the note, or put the long text in a "
            "cell.")
    return text


def check_dv_formula(formula, *, field: str = "formula1") -> str | None:
    if formula is None:
        return None
    text = str(formula)
    check_text_storable(text, what=f"the data-validation {field}")
    if len(text) > MAX_DV_FORMULA_CHARS:
        raise ExcelWouldRefuse(
            f"the data-validation {field} is {len(text):,} characters; Excel "
            f"REFUSES TO OPEN a workbook whose {field} is longer than "
            f"{MAX_DV_FORMULA_CHARS:,} (measured). An inline list this long "
            "belongs in a range: put the items on a sheet and point the "
            "validation at that range instead.")
    return text


def check_header_footer(sections: dict, *, which: str = "header") -> None:
    """Refuse a header/footer whose ASSEMBLED string crosses Excel's limit.
    ``sections`` maps 'left'/'center'/'right' to the ENCODED text (the & codes
    already expanded), which is what actually lands in the file."""
    codes = {"left": "&L", "center": "&C", "right": "&R"}
    total = 0
    for name, text in sections.items():
        if not text:
            continue
        check_text_storable(text, what=f"the {which} {name} section")
        total += len(codes.get(name, "")) + len(str(text))
    if total > MAX_HEADER_FOOTER_CHARS:
        raise ExcelWouldRefuse(
            f"the assembled {which} is {total:,} characters including its "
            f"&L/&C/&R section codes; Excel REFUSES TO OPEN a workbook whose "
            f"{which} string is longer than {MAX_HEADER_FOOTER_CHARS} "
            "(measured: 255 opens, 256 does not). Note the limit is on the "
            "WHOLE string, so three sections share it. Shorten the text or "
            "drop a section.")


#: Characters legal in the AUTHORITY (host) component of a URI: RFC 3986
#: reg-name plus the userinfo and port punctuation and the IPv6 brackets.
_AUTHORITY_OK = re.compile(r"^[A-Za-z0-9\-._~%!$&'()*+,;=:@\[\]]*$")

#: Schemes this server refuses to plant in a workbook. Excel OPENS all of
#: them (measured), so this is not a corruption guard: it is the same posture
#: the server already takes on formula injection, where =cmd|'/c calc'!A1 is
#: deliberately stored as TEXT. An agent-planted hyperlink target is that
#: threat model with a click instead of a recalculation.
UNSAFE_SCHEMES = ("javascript:", "vbscript:", "data:", "ms-msdt:",
                  "ms-search:", "ms-officecmd:", "search-ms:")

#: Schemes that reach the local machine or the network. Legitimate in a real
#: workbook (a link to a file share is ordinary), so these WARN rather than
#: refuse; the caller sees what was written.
SENSITIVE_PREFIXES = ("file:", "\\\\")


def check_hyperlink_target(target: str) -> list[str]:
    """Validate a hyperlink target and return advisory warnings.

    Refuses two classes:

      1. URI-INVALID targets, which make the workbook unopenable. MEASURED:
         a quote, a tab, or a space in the AUTHORITY (``http://x"y``,
         ``http://x y``) gives "Open method of Workbooks class failed", while
         the identical characters in the PATH (``http://example.com/x"y``)
         open fine. ``http://x"/><evil a="`` is one such target: not an XML
         injection (openpyxl escaped it correctly), just a target Excel
         cannot parse.
      2. UNSAFE SCHEMES, per the server's standing injection posture.
    """
    if not isinstance(target, str):
        raise ExcelWouldRefuse("the hyperlink target must be a string")
    check_text_storable(target, what="the hyperlink target")
    low = target.strip().lower()
    for scheme in UNSAFE_SCHEMES:
        if low.startswith(scheme):
            raise ExcelWouldRefuse(
                f"refusing to write a {scheme} hyperlink target. Excel opens "
                "the workbook and the link fires on a click, which is the "
                "same threat model this server already blocks on the formula "
                "side (=cmd|'/c calc'!A1 is stored as TEXT, never armed). "
                "Use an http(s) target, or an in-workbook 'Sheet!A1' "
                "reference.")
    m = re.match(r"^([A-Za-z][A-Za-z0-9+.\-]*):\/\/([^/?#]*)", target)
    if m is not None:
        authority = m.group(2)
        if not _AUTHORITY_OK.match(authority):
            bad = sorted({c for c in authority if not _AUTHORITY_OK.match(c)})
            shown = ", ".join(
                repr(c) if c.isprintable() else f"U+{ord(c):04X}" for c in bad)
            raise ExcelWouldRefuse(
                f"the hyperlink target {target!r} has {shown} in its host "
                f"({authority!r}), which is not a valid URI authority. Excel "
                "REFUSES TO OPEN a workbook carrying such a target (measured; "
                "the same characters later in the PATH are fine). "
                "Percent-encode them, or fix the host.")
    warnings: list[str] = []
    for prefix in SENSITIVE_PREFIXES:
        if low.startswith(prefix.lower()):
            warnings.append(
                f"the hyperlink target {target!r} points at the local machine "
                "or a network share; a UNC target leaks credentials to "
                "whoever owns that share when a reader clicks it")
            break
    return warnings


def check_image_extents(width, height) -> None:
    for label, value in (("width", width), ("height", height)):
        if value is None:
            continue
        try:
            n = int(value)
        except (TypeError, ValueError):
            raise ExcelWouldRefuse(
                f"image {label} must be a whole number of pixels, got "
                f"{value!r}") from None
        if n < MIN_IMAGE_EXTENT:
            raise ExcelWouldRefuse(
                f"image {label} is {n}; a negative display extent cannot be "
                "represented in the drawing XML and Excel REFUSES TO OPEN the "
                "workbook (measured: 0 and 10 open, -5 does not). Pass a "
                "non-negative size, or omit it to keep the image's natural "
                "size.")


def _formula_shape(body: str) -> tuple[int, int]:
    """(deepest nesting of function calls, most arguments in one call) of a
    formula body, ignoring everything inside string literals. A parenthesis
    that only groups an expression is not a function level."""
    frames: list[list] = []   # [is_function, argument count]
    deepest = 0
    most = 0
    in_str = False
    braces = 0
    i = 0
    n = len(body)
    while i < n:
        ch = body[i]
        if in_str:
            if ch == '"':
                if i + 1 < n and body[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
        elif ch == '"':
            in_str = True
        elif ch == "{":
            braces += 1
        elif ch == "}":
            braces = max(0, braces - 1)
        elif ch == "(":
            is_fn = i > 0 and (body[i - 1].isalnum() or body[i - 1] in "_.")
            frames.append([is_fn, 1])
            depth = sum(1 for f in frames if f[0])
            deepest = max(deepest, depth)
        elif ch == ")":
            if frames:
                f = frames.pop()
                if f[0]:
                    most = max(most, f[1])
        elif ch == "," and frames and not braces:
            frames[-1][1] += 1
        i += 1
    for f in frames:
        if f[0]:
            most = max(most, f[1])
    return deepest, most


#: Measured in a hidden Excel 16 (private desktop), on files this server wrote
#: and Excel then opened, so the test is "Excel opens the workbook". A call to
#: a function Excel does not have built in compiles to a name token plus a
#: call token, and the name token's size depends on the callee. A built-in
#: function's own token is 3 (4 for a variable-argument one).
#:
#:   a defined name, a user-defined or add-in function, any name Excel does
#:   not know (a name with sheet scope costs the same as one with workbook
#:   scope) ...................................... 5 + 4 = 9 bytes
#:   a name qualified by a sheet or workbook (S!q) ... 7 + 4 = 11 bytes
#:   a LET variable or a LAMBDA parameter, which the file stores with an
#:   _xlpm. prefix ................................ 10 bytes
#:
#: MEASURED: 1,260 of =q(1)+... open and 1,261 do not (13 bytes a term, the
#: sheet-scoped name the same); 1,638 of =f()+... open and 1,639 do not (10);
#: 1,260 of =f(1)+... (13); 1,092 of =S!q(1)+... open and 1,093 do not (15);
#: with 100 calls to a LET variable and n references to A1 the file opens up
#: to n = 1,867, and with 100 calls to a LAMBDA parameter up to n = 1,869.
#: Counted as 3, a formula of 1,300 name calls saved and Excel would not open
#: the workbook. Long names look dearer (910 calls of lam_s, 744 of S!lam_s,
#: 682 of Sheet2!q) only because the 8,192-character limit meets their 9 to
#: 12-character terms first. Excel's own entry limit for the LET and LAMBDA
#: shapes is higher than the limit on the file, so only file figures count.
NAME_CALL_BYTES = 9
QUALIFIED_CALL_BYTES = 11
BOUND_CALL_BYTES = 10

#: LET and LAMBDA, MEASURED to the byte in a hidden Excel 16: k copies of a
#: shape plus a filler of n references to A1, the last j of them written
#: A1% (a percent sign is one byte), so Excel's capacity is bisected to one
#: byte; each shape at k = 25, and at k = 12 where a part paid once per
#: formula had to be told from a part paid per copy. The plain filler opens
#: at 16,383 bytes exactly (MAX_FORMULA_COMPILED_BYTES). 41 shape and k
#: results, every one fitted by these figures:
#:
#:   the LET call ................................................ 5
#:   the LAMBDA call (whatever it declares) ...................... 22
#:   a name a LET or LAMBDA declares ............................. 0
#:   each use of a declared name ................................. 6
#:   a LAMBDA called where it is written, LAMBDA(..)(args) ....... 4
#:     plus its arguments; the commas between them cost nothing
#:   a missing argument, as in f(1,) or IF(A1,,1) ................ 1
#:   a flag for LET: once for the formula when a LET is outside
#:     every LAMBDA, and once more for each LAMBDA whose own body
#:     holds a LET ............................................... 4
#:   a name a LAMBDA uses but does not declare (a LET name or an
#:     outer LAMBDA's parameter): once per such name per LAMBDA,
#:     also for a LAMBDA that only passes it on to a nested one .. 5
#:
#: The LET flag is not the volatile one (ONCE_PER_FORMULA_BYTES): a formula
#: with both carries both (LET(x,A1,x)+NOW(): 8 bytes once). A second and a
#: third use of a captured name each measured one byte less than a plain use;
#: they are counted as plain uses, so such a formula is counted a byte or two
#: high, never low. Where a size could not be measured, the count errs toward
#: refusing: a formula that would make Excel refuse the whole workbook is the
#: worse outcome. Counted as LET 5 with 3 once and LAMBDA as a 3-byte call,
#: the shapes LET(f,LAMBDA(a,a+1),f(A1)), LAMBDA(a,a+1)(A1) and
#: LAMBDA(a,b,a+b)(A1,A1) passed this check 41, 50 and 6 references past
#: what Excel opens: the save answered ok, and Excel then refused to open the
#: workbook and dropped the formula on repair.
LET_CALL_BYTES = 5
LAMBDA_CALL_BYTES = 22
LET_DECLARED_NAME_BYTES = 0
BOUND_NAME_USE_BYTES = 6
LAMBDA_INVOKE_BYTES = 4
MISSING_ARG_BYTES = 1
LET_FLAG_BYTES = 4
CAPTURED_NAME_BYTES = 5

#: An array constant, MEASURED: 15 bytes in the formula whatever
#: it holds ({1}, {1,2,3}, {1,2;3,4}, {"a","b"} and ten values all cost
#: the same; the values are stored outside the formula). Counted as one
#: byte, a formula of 100 SUM({1,2,3}) and 1,973 references saved and
#: Excel would not open the workbook; it opens 1,798.
ARRAY_CONSTANT_BYTES = 15

#: The prefix the file puts on a LET variable or LAMBDA parameter, at its
#: declaration and at every use (core.calc.normalize_formula writes it).
_BOUND_PREFIX = "_xlpm."

#: Prefixes Excel writes on names it does know: _xlfn. and _xlws. on its
#: newer functions, _xlpm. and _xlop. on LET and LAMBDA parameters. None of
#: these is a plain name call, so _calls_a_name says no. A call through
#: _xlpm. is priced by _call_bytes (BOUND_CALL_BYTES); the others are not
#: measured and keep the built-in floor of 3.
_KNOWN_NAME_PREFIXES = ("_xlfn.", "_xlws.", "_xlpm.", "_xlop.")

#: A LET or LAMBDA is the only thing that declares a name inside a formula.
_BINDS_NAMES = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(", re.IGNORECASE)


def _as_stored(text: str) -> str:
    """The formula as the file holds it. A call to a LET variable or LAMBDA
    parameter is told from a call to a defined name by the _xlpm. prefix the
    file gives the declared name, so a typed formula is priced as the text
    that will be saved. A formula that cannot be normalized is priced as
    typed (it is refused where it is written)."""
    if not _BINDS_NAMES.search(text):
        return text
    from .calc import normalize_formula

    try:
        return normalize_formula(text)[0]
    except Exception:  # noqa: BLE001
        return text


def _calls_a_name(func_open: str) -> bool:
    """True when a function token (``q(``, ``SUM(``, ``_xludf.f(``) calls
    something Excel does not have built in, which compiles as a name plus
    a call (NAME_CALL_BYTES)."""
    from .calc import _BUILTIN_FUNCS

    name = func_open[:-1] if func_open.endswith("(") else func_open
    if "!" in name:
        return True  # a name qualified by a sheet or a workbook
    low = name.lower()
    if low.startswith("_xludf."):
        return True
    if low.startswith(_KNOWN_NAME_PREFIXES):
        return False
    return name.upper() not in _BUILTIN_FUNCS


#: Built-in functions Excel compiles as a VARIABLE-argument call, a 4-byte
#: token where a fixed-argument call takes 3. The rule is any
#: function whose arguments can vary in number, by an optional or a
#: repeating argument. MEASURED in a hidden Excel 16 on a private desktop,
#: every one listed: 100 calls of it plus n references to A1, with Excel's
#: largest n bisected on files holding the text this server stores, pins
#: the call's own bytes to one integer (XOR: 1,898 references open and
#: 1,899 do not, so 4; ABS: 1,910 and 1,911, so 3). 210 functions were
#: measured and none broke the rule. A function not listed is priced 3, so
#: the count stays a floor. IF and CHOOSE are variable too and are priced
#: with their jump tokens in compiled_size_floor, so they are not listed.
VAR_ARG_FUNCS: frozenset[str] = frozenset({
    "ADDRESS", "AGGREGATE", "AND", "AVEDEV", "AVERAGE", "AVERAGEA",
    "AVERAGEIF", "AVERAGEIFS", "BASE", "CEILING.MATH", "CELL", "CHOOSECOLS",
    "CHOOSEROWS", "COLUMN", "CONCAT", "CONCATENATE", "COUNT", "COUNTA",
    "COUNTIFS", "DB", "DDB", "DEC2BIN", "DEVSQ", "DOLLAR", "DROP", "FILTER",
    "FIND", "FIXED", "FLOOR.MATH", "FV", "GCD", "GEOMEAN", "GETPIVOTDATA",
    "HARMEAN", "HLOOKUP", "HSTACK", "HYPERLINK", "IFS", "INDEX", "INDIRECT",
    "IPMT", "IRR", "KURT", "LCM", "LEFT", "LINEST", "LOG", "LOOKUP", "MATCH",
    "MAX", "MAXA", "MAXIFS", "MEDIAN", "MIN", "MINA", "MINIFS", "MODE",
    "MODE.MULT", "MODE.SNGL", "NETWORKDAYS", "NPER", "NPV", "NUMBERVALUE",
    "OFFSET", "OR", "PERCENTRANK", "PMT", "PPMT", "PRODUCT", "PV",
    "RANDARRAY", "RANK", "RANK.AVG", "RANK.EQ", "RATE", "RIGHT", "ROMAN",
    "ROW", "SEARCH", "SEQUENCE", "SHEET", "SHEETS", "SKEW", "SORT", "SORTBY",
    "STDEV", "STDEV.P", "STDEV.S", "STDEVA", "STDEVP", "STDEVPA",
    "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ",
    "SWITCH", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT",
    "TREND", "TRUNC", "UNIQUE", "VAR", "VAR.P", "VAR.S", "VARA", "VARP",
    "VARPA", "VLOOKUP", "VSTACK", "WEEKDAY", "WEEKNUM", "WORKDAY",
    "XLOOKUP", "XMATCH", "XOR", "YEARFRAC",
    # MEASURED the same way: each 1 byte more a call than a
    # fixed-argument function (IMSUM x 100: 1,860 references open, and the
    # count at 3 accepted 1,873).
    "EXPAND", "IMPRODUCT", "IMSUM", "TOCOL", "TOROW", "WRAPCOLS", "WRAPROWS",
    # MEASURED to the byte with a LAMBDA argument (the LET and LAMBDA
    # figures above): MAP, BYROW and BYCOL take 4. REDUCE, SCAN, MAKEARRAY
    # and ISOMITTED measured 3, although REDUCE and SCAN have an optional
    # argument, so they are not listed.
    "BYCOL", "BYROW", "MAP",
})

#: IFERROR and IFNA compile like a two-argument IF: an attribute token after
#: the value and a jump after the fallback, 4 bytes each, on a 3-byte call.
#: MEASURED: 11 bytes a call for both (IFERROR(A1,0) x 100
#: plus 1,773 references open, 1,774 do not; IFNA the same).
IF_STYLE_FALLBACK_BYTES = 8
_IF_STYLE_FALLBACK = frozenset({"IFERROR", "IFNA"})

#: A formula that calls any of these carries one 4-byte attribute token at
#: its start, however many calls it makes (the volatile ones, and SHEET and
#: SHEETS). MEASURED: each, with 100 calls, opens at exactly 4 bytes fewer
#: references' worth than its calls alone account for (NOW() x 100: 1,997
#: references open, 1,998 do not).
ONCE_PER_FORMULA_BYTES = 4
_ONCE_PER_FORMULA = frozenset({
    "CELL", "INDIRECT", "NOW", "OFFSET", "RAND", "RANDBETWEEN", "SHEET",
    "SHEETS", "TODAY",
})


def _builtin_name(func_open: str) -> str:
    """The function's own name, upper case, without the _xlfn. / _xlws.
    prefixes the file puts on newer functions (STDEV.S keeps its dot)."""
    name = func_open[:-1] if func_open.endswith("(") else func_open
    low = name.lower()
    while low.startswith(("_xlfn.", "_xlws.")):
        name, low = name[6:], low[6:]
    return name.upper()


def _call_bytes(func_open: str) -> int:
    """What the opening of a function call costs once compiled: 3 for a
    fixed-argument built-in (or a name Excel writes with a known prefix), 4
    for a variable-argument one (VAR_ARG_FUNCS), and for a call to a name,
    NAME_CALL_BYTES, QUALIFIED_CALL_BYTES or BOUND_CALL_BYTES."""
    name = func_open[:-1] if func_open.endswith("(") else func_open
    if "!" in name:
        return QUALIFIED_CALL_BYTES
    if name.lower().startswith(_BOUND_PREFIX):
        return BOUND_CALL_BYTES
    if _calls_a_name(func_open):
        return NAME_CALL_BYTES
    builtin = _builtin_name(func_open)
    if builtin == "LET":
        return LET_CALL_BYTES
    if builtin == "LAMBDA":
        return LAMBDA_CALL_BYTES
    return 4 if builtin in VAR_ARG_FUNCS else 3


_CELL_REF = re.compile(r"[A-Za-z]{1,3}[0-9]+")
_AREA_REF = re.compile(
    r"(?:[A-Za-z]{1,3}[0-9]+|[A-Za-z]{1,3}|[0-9]+):(?:[A-Za-z]{1,3}[0-9]+|[A-Za-z]{1,3}|[0-9]+)")

#: A union of references, as in SUM((A1,B2)), compiles inside one memory
#: token around the whole parenthesized reference: 7 bytes when it
#: holds only cell and range references on the formula's own sheet, 3 when
#: it holds something Excel evaluates (a function call, a reference to
#: another sheet). One token per group, however many unions it holds; a
#: group inside a group shares it. MEASURED in a hidden Excel 16 on a
#: private desktop (100 groups plus n references to A1, n bisected):
#: SUM((A1,A1)) costs 7 more than SUM's 4 and its references and unions,
#: and so do 3 or 5 references, two ranges, ((A1,A1),A1), (A1,A1)+1,
#: INDEX's first argument and (A1,A1) on its own; two groups cost 14;
#: SUM((Sheet!A1,Sheet!A1)) and SUM((A1,INDEX(A1:B2,1,1))) cost 3. A
#: parenthesis with no union costs nothing more (SUM((A1)), ABS((A1))).
#: =SUM((A1,...,A1)) opens with 2,046 references and not with 2,047, which
#: the floor accepted before. Anything else inside a group (a number, a
#: name, an operator, a space) is not measured and adds nothing, so the
#: count stays a floor, and so does a group inside a function call inside
#: another group.
UNION_GROUP_BYTES = 7
UNION_GROUP_EVALUATED_BYTES = 3


def compiled_size_floor(formula: str) -> int:
    """The size of a formula as Excel stores it (bytes), counted from token
    sizes measured in Excel.

    Each token this count recognises is counted at its measured size. The
    LET and LAMBDA parts were measured to the byte; where a part could only
    be bounded (a repeated use of a name a LAMBDA captures), it is counted
    high, so the check refuses rather than writes a formula Excel might not
    open. A token this count does not recognise counts one byte, the least
    any token takes. 0 when the formula cannot be tokenized."""
    from openpyxl.formula.tokenizer import Token, Tokenizer

    text = _as_stored(formula if formula.startswith("=") else "=" + formula)
    try:
        items = Tokenizer(text).items
    except Exception:  # noqa: BLE001
        return 0
    bound = _bindings(items)
    size = (LET_FLAG_BYTES * len(bound["let_bodies"])
            + CAPTURED_NAME_BYTES * len(bound["captured"])
            + MISSING_ARG_BYTES * bound["missing"])
    # Measured in a hidden Excel: three token kinds the first
    # floor left out made formulas well under 8,192 characters pass while
    # Excel would not open the file.
    #  - IF compiles with an attribute token after its condition (4 bytes)
    #    and a jump after each value (4 bytes each), and is a
    #    variable-argument call (4 bytes, not 3).
    #  - CHOOSE compiles with a jump table (4 bytes plus 2 per option plus
    #    2) and a jump after each option (4 bytes each), also variable.
    #  - Whitespace is kept as its own 4-byte token, except a space between
    #    two references: the intersection operator, which with the wrapper
    #    Excel puts around it measures 8 bytes (468 intersections of two
    #    areas open, 469 do not; 712 of two cells open, 713 do not).
    # Measured: at a floor of exactly 16,383 the IF, CHOOSE and whitespace
    # forms open and one more term does not, so those counts are exact.
    # An array constant's values are stored outside the formula and do not
    # count against the limit (a SUM over 1,900 constants opens); the
    # constant itself takes ARRAY_CONSTANT_BYTES whatever it holds.
    frames: list[list] = []  # [function name, argument count]
    array_depth = 0
    prev_sig = None
    once = False  # a call that adds its one token per formula
    # Open function calls and LAMBDAs called where written (None),
    # and grouping parentheses
    # ([holds a union, the bytes its union token adds at least]).
    nest: list = []
    for idx, t in enumerate(items):
        top = nest[-1] if nest else None
        if idx in bound["invoke_open"]:
            if top is not None:
                top[1] = min(top[1], UNION_GROUP_EVALUATED_BYTES)
            nest.append(None)
        elif idx in bound["invoke_close"]:
            if nest and top is None:
                nest.pop()
        elif t.type == Token.PAREN and t.subtype == Token.OPEN:
            nest.append([False, UNION_GROUP_BYTES])
        elif t.type == Token.PAREN and t.subtype == Token.CLOSE:
            if top is not None:
                union, add = nest.pop()
                outer = nest[-1] if nest else None
                if outer is not None:  # a group inside a group: one token
                    outer[0] = outer[0] or union
                    outer[1] = min(outer[1], add)
                elif union and not any(f is not None for f in nest):
                    size += add
        elif t.type == Token.FUNC and t.subtype == Token.OPEN:
            if top is not None:
                top[1] = min(top[1], UNION_GROUP_EVALUATED_BYTES)
            nest.append(None)
        elif t.type == Token.FUNC and t.subtype == Token.CLOSE:
            if nest and top is None:
                nest.pop()
        elif top is not None:  # a token at a group's own level
            if t.type == Token.OP_IN and t.value == ",":
                top[0] = True
            elif t.type == Token.OPERAND and t.subtype == Token.RANGE:
                ref = t.value.rsplit("!", 1)[-1].replace("$", "")
                if not (_CELL_REF.fullmatch(ref) or _AREA_REF.fullmatch(ref)):
                    top[1] = 0
                elif "!" in t.value:
                    top[1] = min(top[1], UNION_GROUP_EVALUATED_BYTES)
            else:
                top[1] = 0
        if t.type == Token.WSPACE:
            if array_depth:
                continue
            at = next((j for j in range(idx + 1, len(items))
                       if items[j].type != Token.WSPACE), None)
            nxt = items[at] if at is not None else None
            # MEASURED: LET( x , A1 , x ) costs the spaces of three tokens,
            # not six, and LAMBDA( a , a )( 1 ) the spaces of four. Excel
            # keeps a space with the token that follows it, so spaces before
            # an argument separator join those after it, and spaces before a
            # name a LET or LAMBDA declares (which is no token) join the next
            # ones. Counted as nothing.
            if nxt is not None and (
                    (nxt.type == Token.SEP and nxt.subtype == Token.ARG)
                    or at in bound["declared"]):
                continue
            left = prev_sig is not None and (
                prev_sig.type == Token.OPERAND
                or (prev_sig.subtype == Token.CLOSE
                    and prev_sig.type in (Token.FUNC, Token.PAREN)))
            right = nxt is not None and (
                nxt.type == Token.OPERAND
                or (nxt.subtype == Token.OPEN
                    and nxt.type in (Token.FUNC, Token.PAREN)))
            size += 8 if (left and right) else 4
            continue
        if t.type == Token.ARRAY:
            if t.subtype == Token.OPEN:
                if array_depth == 0:
                    size += ARRAY_CONSTANT_BYTES
                array_depth += 1
            else:
                array_depth = max(0, array_depth - 1)
            prev_sig = t
            continue
        if array_depth:
            prev_sig = t
            continue
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            frames.append([t.value[:-1].upper().split(".")[-1], 1])
        elif t.type == Token.SEP and t.subtype == Token.ARG and frames:
            frames[-1][1] += 1
        elif t.type == Token.FUNC and t.subtype == Token.CLOSE and frames:
            name, args = frames.pop()
            if name == "IF":
                size += 1 + 4 + 4 * max(0, args - 1)
            elif name == "CHOOSE" and args >= 2:
                options = args - 1
                size += 1 + (4 + 2 * (options + 1)) + 4 * options
            elif name in _IF_STYLE_FALLBACK:
                size += IF_STYLE_FALLBACK_BYTES
        prev_sig = t
        if t.type == Token.OPERAND:
            if t.subtype == Token.RANGE:
                ref = t.value.rsplit("!", 1)[-1].replace("$", "")
                sheet = 2 if "!" in t.value else 0
                if _CELL_REF.fullmatch(ref):
                    size += 7 + sheet
                elif _AREA_REF.fullmatch(ref):
                    size += 13 + sheet
                elif "[" in ref:
                    # A structured reference (T[a], T[@a],
                    # [@a]) compiles to a 14-byte list token. MEASURED:
                    # 1,092 of =T[a]+... open, 1,093 do not. A two-column
                    # reference measures larger; 14 keeps this a floor.
                    size += 14
                elif ref.lower().startswith(_BOUND_PREFIX):
                    # A name a LET or LAMBDA declares costs nothing in the
                    # formula; every other LET or LAMBDA name is a use.
                    size += (LET_DECLARED_NAME_BYTES
                             if idx in bound["declared"]
                             else BOUND_NAME_USE_BYTES)
                elif ref:
                    # A defined name (or a table's bare name)
                    # compiles to a 5-byte name token. MEASURED: 2,730 of
                    # =q+q+... open, 2,731 do not. Counted as one byte, a
                    # 4,000-reference formula saved and Excel would not
                    # open the workbook.
                    size += 5
                else:
                    size += 1
            elif t.subtype == Token.NUMBER:
                try:
                    n = float(t.value)
                    size += 3 if n.is_integer() and 0 <= n <= 65535 else 9
                except ValueError:
                    size += 1
            elif t.subtype == Token.TEXT:
                inner = t.value[1:-1].replace('""', '"')
                size += 3 + 2 * len(inner)
            elif t.subtype in (Token.LOGICAL, Token.ERROR):
                size += 2
            else:
                size += 1
        elif t.type == Token.FUNC and t.subtype == Token.OPEN:
            size += _call_bytes(t.value)
            if not once and not _calls_a_name(t.value) \
                    and _builtin_name(t.value) in _ONCE_PER_FORMULA:
                # Once per formula, also when the call is inside a LAMBDA;
                # LET's own flag is counted apart (LET_FLAG_BYTES).
                once = True
                size += ONCE_PER_FORMULA_BYTES
        elif idx in bound["invoke_open"]:
            size += LAMBDA_INVOKE_BYTES
        elif idx in bound["invoke_sep"] or idx in bound["invoke_close"]:
            pass  # part of LAMBDA_INVOKE_BYTES
        elif t.type in (Token.OP_IN, Token.OP_PRE, Token.OP_POST):
            size += 1
        elif t.type == Token.PAREN and t.subtype == Token.CLOSE:
            size += 1
    return size


def _bindings(items) -> dict:
    """What a formula's LET and LAMBDA calls declare, call where written and
    capture, and how many arguments it leaves out, worked out before the
    count (the figures above LAMBDA_CALL_BYTES).

    "declared": item indices of names a LET or LAMBDA declares (a whole
    argument in an odd place of a LET, other than the last; any whole
    argument of a LAMBDA other than the last, its calculation).
    "invoke_open" / "invoke_close" / "invoke_sep": the parentheses and
    commas of a LAMBDA called where it is written, LAMBDA(a,a+1)(A1), or of
    a call on what such a call or a bound name returns, f(1)(2).
    "captured": (LAMBDA, name) pairs, one for each name a LAMBDA uses
    (itself, or through a LAMBDA nested in it) that a LET or LAMBDA outside
    it declares. "let_bodies": "top", and each LAMBDA, whose own body holds
    a LET. "missing": arguments left out, f(1,) or IF(A1,,1)."""
    from openpyxl.formula.tokenizer import Token

    out: dict = {"declared": set(), "invoke_open": set(), "invoke_close": set(),
                 "invoke_sep": set(), "captured": set(), "let_bodies": set(),
                 "missing": 0}
    stack: list[dict] = []        # open calls and parentheses, innermost last
    callable_ends: set[int] = set()
    sig = [i for i, t in enumerate(items) if t.type != Token.WSPACE]

    def is_sep(t) -> bool:
        return t is not None and t.type == Token.SEP and t.subtype == Token.ARG

    def is_comma(t) -> bool:
        return t is not None and t.type == Token.OP_IN and t.value == ","

    def closes(t, kind) -> bool:
        return t is not None and t.type == kind and t.subtype == Token.CLOSE

    def body() -> object:
        for f in reversed(stack):
            if f["kind"] == "call" and f["name"] == "LAMBDA":
                return f["at"]
        return "top"

    def use(name: str) -> None:
        crossed = []
        for f in reversed(stack):
            if f["kind"] != "call":
                continue
            if name in f["decl"]:
                out["captured"].update((lam, name) for lam in crossed)
                return
            if f["name"] == "LAMBDA":
                crossed.append(f["at"])

    for pos, i in enumerate(sig):
        t = items[i]
        prv = items[sig[pos - 1]] if pos else None
        nxt = items[sig[pos + 1]] if pos + 1 < len(sig) else None
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            bound = t.value.lower().startswith(_BOUND_PREFIX)
            if bound:
                use(t.value[len(_BOUND_PREFIX):-1].lower())
            name = None if bound else _builtin_name(t.value)
            if name == "LET":
                out["let_bodies"].add(body())
            stack.append({"kind": "call", "name": name, "at": i, "arg": 1, "decl": set()})
            if is_sep(nxt):
                out["missing"] += 1
        elif is_sep(t):
            if stack and stack[-1]["kind"] == "call":
                stack[-1]["arg"] += 1
            if is_sep(nxt) or closes(nxt, Token.FUNC):
                out["missing"] += 1
        elif closes(t, Token.FUNC):
            if stack and stack[-1]["kind"] == "call":
                # A LAMBDA, or a call through a bound name, may return a
                # LAMBDA that the next parenthesis calls.
                if stack.pop()["name"] in (None, "LAMBDA"):
                    callable_ends.add(i)
        elif t.type == Token.PAREN and t.subtype == Token.OPEN:
            if i - 1 in callable_ends:
                out["invoke_open"].add(i)
                stack.append({"kind": "invoke", "at": i})
                if is_comma(nxt):
                    out["missing"] += 1
            else:
                stack.append({"kind": "paren", "at": i})
        elif closes(t, Token.PAREN):
            if stack and stack[-1]["kind"] in ("invoke", "paren"):
                if stack.pop()["kind"] == "invoke":
                    out["invoke_close"].add(i)
                    callable_ends.add(i)
        elif is_comma(t) and stack and stack[-1]["kind"] == "invoke":
            out["invoke_sep"].add(i)
            if is_comma(nxt) or closes(nxt, Token.PAREN):
                out["missing"] += 1
        elif (t.type == Token.OPERAND and t.subtype == Token.RANGE
              and t.value.lower().startswith(_BOUND_PREFIX)):
            name = t.value[len(_BOUND_PREFIX):].lower()
            top = stack[-1] if stack else None
            whole = is_sep(nxt) and (is_sep(prv) or (
                prv is not None and prv.type == Token.FUNC and prv.subtype == Token.OPEN))
            if (whole and top is not None and top["kind"] == "call"
                    and (top["name"] == "LAMBDA"
                         or (top["name"] == "LET" and top["arg"] % 2 == 1))):
                out["declared"].add(i)
                top["decl"].add(name)
            else:
                use(name)
    return out


#: The prefixes Excel writes into a stored formula.
_STORAGE_PREFIXES = ("_xlfn.", "_xlpm.", "_xlop.", "_xlws.")

#: The prefixes at most quadruple a formula (x+x+... becomes
#: _xlpm.x+_xlpm.x+...), so one this short cannot pass the limit stored.
_SURELY_SHORT = MAX_FORMULA_CHARS // 5


def _stored_form(body: str) -> str:
    """The formula body as the file will hold it. One that cannot be
    normalized is counted as given (its write is refused where it is
    normalized)."""
    if "(" not in body or len(body) <= _SURELY_SHORT:
        return body
    from .calc import normalize_formula

    try:
        return normalize_formula(body)[0]
    except Exception:  # noqa: BLE001
        return body


def _written_form(body: str) -> str:
    """The formula body as a user writes it: without Excel's prefixes."""
    if not any(p in body for p in _STORAGE_PREFIXES):
        return body
    from .calc import denormalize_formula

    return denormalize_formula(body)


def _too_long_text(what: str, written: int, stored: str) -> str:
    """The refusal for a formula past 8,192 stored characters, in the
    length the user wrote. When the prefixes made the difference, it also
    says the most that fits as written for a formula built the same way,
    and the stored count only as the reason."""
    limit = MAX_FORMULA_CHARS
    remedy = ("Split the calculation across helper cells, or move repeated "
              "parts into a defined name.")
    names = "_xlpm." in stored or "_xlop." in stored
    adds = []
    if "_xlfn." in stored:
        adds.append("_xlfn. to each newer function such as LET or LAMBDA" if names
                    else "_xlfn. to each newer function such as XLOOKUP or FILTER")
    if names:
        adds.append("_xlpm. to each LET or LAMBDA name")
    if len(stored) == written or not adds:
        return (f"{what} is {len(stored):,} characters long; Excel's limit is "
                f"{limit:,}, and it will not open a workbook holding a longer "
                f"one. {remedy}")
    fits = written * limit // len(stored)
    return (f"{what} is {written:,} characters as written, and a formula built "
            f"like this one can be at most about {fits:,} characters as "
            f"written: Excel will not open a workbook holding a longer one. "
            f"Excel's limit of {limit:,} characters counts the formula as Excel "
            f"stores it, {len(stored):,} characters here: Excel adds "
            f"{' and '.join(adds)}. {remedy}")


def check_formula_length(formula, *, what: str = "this formula") -> None:
    """Excel's 8,192-character limit, counted as Excel counts it: on the
    formula as the file stores it. The refusal speaks in the length the
    user wrote (_too_long_text)."""
    if not isinstance(formula, str):
        return
    body = formula[1:] if formula.startswith("=") else formula
    stored = _stored_form(body)
    if len(stored) > MAX_FORMULA_CHARS:
        raise ExcelWouldRefuse(_too_long_text(what, len(_written_form(body)), stored))


def check_formula(formula, *, what: str = "this formula") -> None:
    """Excel's hard formula limits, checked before anything reaches the disk:
    a file holding a formula past any of them is one Excel will not open
    (a 10,000-character formula once saved as "verified")."""
    if not isinstance(formula, str):
        return
    body = formula[1:] if formula.startswith("=") else formula
    check_formula_length(body, what=what)
    if len(body) > 2_000:
        compiled = compiled_size_floor(body)
        if compiled > MAX_FORMULA_COMPILED_BYTES:
            raise ExcelWouldRefuse(
                f"{what} is too large for Excel to store: compiled, it needs "
                f"at least {compiled:,} bytes and Excel's limit is "
                f"{MAX_FORMULA_COMPILED_BYTES:,} (about 2,000 cell "
                "references, about 540 IF calls, or about 1,260 calls to a "
                "LAMBDA name or user-defined function, in one formula), so "
                "Excel would not open the workbook. Use a range "
                "(SUM(A1:A5000)) instead of listing cells one by one, or "
                "split it across helper cells.")
    if "(" not in body:
        return
    deepest, most = _formula_shape(body)
    if deepest > MAX_FORMULA_NESTING:
        raise ExcelWouldRefuse(
            f"{what} nests {deepest} levels of functions; Excel allows "
            f"{MAX_FORMULA_NESTING}. Break the calculation into helper "
            "cells.")
    if most > MAX_FUNCTION_ARGS:
        raise ExcelWouldRefuse(
            f"{what} passes {most} arguments to one function; Excel allows "
            f"{MAX_FUNCTION_ARGS}. Pass a range instead of listing cells one "
            "by one.")


def check_sheet_title(title: str) -> str:
    """Excel's sheet-name grammar, for completeness of the refuse-class
    table: 1-31 characters, none of : \\ / ? * [ ], and not bracketed by
    apostrophes."""
    if not isinstance(title, str) or not title:
        raise ExcelWouldRefuse("a sheet name must be a non-empty string")
    check_text_storable(title, what="the sheet name")
    if len(title) > 31:
        raise ExcelWouldRefuse(
            f"the sheet name is {len(title)} characters; Excel's limit is 31")
    bad = sorted(set(title) & set(":\\/?*[]"))
    if bad:
        raise ExcelWouldRefuse(
            f"the sheet name contains {', '.join(repr(c) for c in bad)}; "
            "Excel bans : \\ / ? * [ ] in sheet names")
    if title.startswith("'") or title.endswith("'"):
        raise ExcelWouldRefuse(
            "a sheet name may not start or end with an apostrophe")
    if title.casefold() == "history":
        # Excel reserves it (the change-history sheet of a shared
        # workbook) and refuses it in its own rename box.
        raise ExcelWouldRefuse("Excel reserves the sheet name 'History'; choose another")
    return title


def describe() -> dict:
    """The refuse-class table, for the docs and the safety gate."""
    return {
        "comment_text_chars": MAX_COMMENT_CHARS,
        "data_validation_formula_chars": MAX_DV_FORMULA_CHARS,
        "header_footer_string_chars": MAX_HEADER_FOOTER_CHARS,
        "image_extent_minimum": MIN_IMAGE_EXTENT,
        "cell_text_chars": MAX_CELL_CHARS,
        "formula_chars": MAX_FORMULA_CHARS,
        "formula_compiled_bytes": MAX_FORMULA_COMPILED_BYTES,
        "formula_nesting": MAX_FORMULA_NESTING,
        "function_arguments": MAX_FUNCTION_ARGS,
        "hyperlink": "valid URI authority; unsafe schemes refused",
        "formula": "no whitespace before '('; no bare LET/LAMBDA names",
        "measured_against": "Excel",
    }


__all__ = [
    "MAX_COMMENT_CHARS", "MAX_DV_FORMULA_CHARS", "DV_FORMULA_SOFT_CHARS",
    "MAX_HEADER_FOOTER_CHARS", "MIN_IMAGE_EXTENT", "MAX_CELL_CHARS",
    "MAX_FORMULA_CHARS", "MAX_FORMULA_NESTING", "MAX_FUNCTION_ARGS",
    "MAX_FORMULA_COMPILED_BYTES", "compiled_size_floor", "check_formula_length",
    "check_formula",
    "UNSAFE_SCHEMES", "SENSITIVE_PREFIXES",
    "check_text_storable", "check_cell_text", "check_comment_text",
    "check_dv_formula",
    "check_header_footer", "check_hyperlink_target", "check_image_extents",
    "check_sheet_title", "describe",
]
