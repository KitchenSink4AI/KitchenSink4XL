"""Excel number formats applied to a value: the text a cell DISPLAYS (a ZIP
code formatted 00000 was read and exported as 2134; there was no way to get
what Excel shows).

Covers the formats real workbooks use: General, text (@), fixed and
thousands-separated decimals with literal prefixes and suffixes ("$", "€",
"kg"), zero-padded codes (00000), percentages, scientific, scaling commas
(#,##0, shows thousands), fractions (# ?/?, # ??/??, # ?/8), the four
sections (positive;negative;zero;text, parentheses for negatives), colour
and condition brackets (ignored: a display-text export has no colour),
[$€-x] currency tags, dates and times (yyyy, yy, mmmm, mmm, mm, m, dddd,
ddd, dd, d, hh, h, mm after h or before s, ss, fractions of a second,
AM/PM) and elapsed time ([h]:mm:ss, [mm]:ss, [ss]). Rounding is Excel's:
the value to 15 significant digits, then halves away from zero (0.125 under
0.00 is 0.13). A format outside that set falls back to General, never to an
exception. The expected texts in the export tests were read from Excel's
own Range.Text for the same values and formats.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from decimal import ROUND_HALF_UP, Decimal, localcontext

_EPOCH = _dt.datetime(1899, 12, 30)
_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
           "September", "October", "November", "December"]
_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _sections(fmt: str) -> list[str]:
    out, cur, quote, i = [], [], False, 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            quote = not quote
        elif ch == "\\" and i + 1 < len(fmt):
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        elif ch == ";" and not quote:
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


def _strip_brackets(sec: str) -> str:
    """[Red], [Color10], [>=100], [$-409] change no digit of a display-text
    export; [$€-407] keeps its currency symbol; [h], [mm], [ss] are elapsed
    time and stay."""
    def keep(m: re.Match) -> str:
        body = m.group(1)
        if body.startswith("$"):
            return '"' + body[1:].split("-", 1)[0] + '"'
        if re.fullmatch(r"[hH]+|[mM]+|[sS]+", body):
            return m.group(0)
        return ""
    return re.sub(r"\[([^\]]*)\]", keep, sec)


def _literal(text: str) -> str:
    """Quoted text and backslash escapes as the characters they show; the _
    and * spacing codes dropped."""
    out, i = [], 0
    while i < len(text):
        ch = text[i]
        if ch == '"':
            j = text.find('"', i + 1)
            j = len(text) if j < 0 else j
            out.append(text[i + 1:j])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < len(text):
            out.append(text[i + 1])
            i += 2
            continue
        if ch == "_" and i + 1 < len(text):
            # _x leaves the width of x empty: Excel's text holds one space
            # there (0.00_) shows 1234.5678 as "1234.57 ", measured)
            out.append(" ")
            i += 2
            continue
        if ch == "*" and i + 1 < len(text):
            # *x repeats x to fill the column: how many depends on the
            # column's width in pixels, so the fill is left out
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _unquoted(sec: str) -> str:
    return re.sub(r'"[^"]*"|\\.', "", sec)


def general(value) -> str:
    """Excel's General display of a number (up to 15 significant digits)."""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value or value in (float("inf"), float("-inf")):
            return str(value)
        if value == int(value) and abs(value) < 1e15:
            return str(int(value))
        text = f"{value:.15g}"
        if "e" in text:
            mant, exp = text.split("e")
            return f"{mant}E{'+' if int(exp) >= 0 else '-'}{abs(int(exp)):02d}"
        return text
    return "" if value is None else str(value)


#: The smallest normal double; Excel holds nothing smaller than this but 0.
_TINY = 2.2250738585072014e-308


def _js_precision(v: float, p: int) -> str:
    """JavaScript's Number.prototype.toPrecision(p)."""
    mant, exp = f"{v:.{p - 1}e}".split("e")
    e = int(exp)
    if e < -6 or e >= p:
        return f"{mant}e{'+' if e >= 0 else '-'}{abs(e)}"
    digits = mant.replace("-", "").replace(".", "")
    sign = "-" if v < 0 else ""
    if e >= 0:
        whole, frac = digits[:e + 1], digits[e + 1:]
        return sign + whole + ("." + frac if frac else "")
    return sign + "0." + "0" * (-e - 1) + digits


def _strip_decimal(o: str) -> str:
    return o if "." not in o else re.sub(r"(?:\.0*|(\.\d*[1-9])0+)$", r"\1", o)


def _normalize_exp(o: str) -> str:
    if "E" not in o:
        return o
    o = re.sub(r"(?:\.0*|(\.\d*[1-9])0+)E", r"\1E", o)
    return re.sub(r"(E[+-])(\d)$", r"\g<1>0\2", o)


def general_display(value) -> str:
    """The text General shows for a number in a worksheet cell: at most 11
    characters (12 with a minus), decimal while that fits, otherwise
    scientific with up to 6 significant digits (measured: 0.123456789012
    shows 0.123456789, 123456789.123 shows 123456789.1, 9007199254740993
    shows 9.0072E+15, -0.000001234 shows -0.000001234, a denormal shows 0).
    general() keeps 15 digits for exports, which are data, not display."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return general(value)
    v = float(value)
    if v != v or v in (float("inf"), float("-inf")):
        return general(value)
    if abs(v) < _TINY:
        return "0"
    v = float(_dec15(v))  # Excel's 15 significant digits first
    width = 12 if v < 0 else 11
    mag = math.floor(math.log10(abs(v)))
    if -4 <= mag <= -1:
        o = _js_precision(v, 10 + mag)
    elif abs(mag) <= 9:
        o = _strip_decimal(f"{v:.12f}")
        if len(o) > width:
            o = _js_precision(v, 10)
            if len(o) > width:
                o = f"{v:.5e}"
    elif mag == 10:
        o = f"{v:.10f}"[:12]
    else:
        o = _strip_decimal(f"{v:.11f}")
        if len(o) > width or o in ("0", "-0"):
            o = _js_precision(v, 6)
    return _strip_decimal(_normalize_exp(o.upper()))


def _is_date_section(sec: str) -> bool:
    body = _unquoted(sec)
    if re.search(r"[yYdDhHsS]", body):
        return True
    # months, eras and era years (m, g, e, b) when no digit place says number
    return bool(re.search(r"[mMgGeEbB]", body)) and not re.search(r"[0#?]", body)


def _dec15(v: float) -> Decimal:
    """v as Excel holds it for display: rounded to 15 significant digits
    (2.675 is 2.675 here, not the binary 2.67499999...)."""
    if v == 0 or v != v:
        return Decimal(0)
    return Decimal(f"{v:.15g}")


def _round(d: Decimal, places: int) -> Decimal:
    """d to places decimals, halves away from zero, as Excel displays."""
    with localcontext() as ctx:
        ctx.prec = 1000
        return d.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def _as_datetime(value):
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime.combine(value, _dt.time())
    if isinstance(value, _dt.time):
        return _dt.datetime.combine(_dt.date(1899, 12, 30), value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return _EPOCH + _dt.timedelta(days=float(value))
        except OverflowError:
            return None
    return None


#: Day 0 of the 1904 date system.
_EPOCH_1904 = _dt.datetime(1904, 1, 1)
#: The first serial past 31 December 9999, Excel's last date, per system.
_PAST_LAST = {False: 2958466, True: 2958466 - 1462}


def _serial(value, date1904: bool = False) -> tuple[float, int] | None:
    """(the Excel serial in days behind value, the days a calendar date
    built from that serial must move to be value's own date). Excel counts
    a 29 February 1900, so a real date before 1 March 1900 is one serial
    day further from the epoch than the calendar says (the reader shifts it
    the same way). The 1904 date system counts from 1 January 1904 and has
    no such day."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value), 0
    if isinstance(value, _dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second
                + value.microsecond / 1e6) / 86400, 0
    d = _as_datetime(value)
    if d is None:
        return None
    delta = d - (_EPOCH_1904 if date1904 else _EPOCH)
    serial = delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400
    if not date1904 and 1 < serial < 61:
        return serial - 1, 1
    return serial, 0


_DATE_TOKENS = re.compile(r'(\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|'
                          r'hh|h|ss|s|am/pm|a/p|g+|e+|b+|\.0+|"[^"]*"|\\.|.)', re.IGNORECASE)


class Unmatched(Exception):
    """A format whose displayed text depends on a locale or calendar this
    module cannot render (Excel's own text was not measured for it)."""


#: Tokens whose text depends on the format's locale.
_NAMED = ("dddd", "ddd", "mmmm", "mmm", "mmmmm", "am/pm", "a/p")
_LOCALE_TAG = re.compile(r"\[\$([^\]]*?)-([^\]]+)\]")


def _locale(raw_sec: str):
    """The language a section's [$...-LCID] tag names: an int language id,
    None when it names none, or the tag text when it is a locale name this
    module does not know."""
    m = _LOCALE_TAG.search(_unquoted_positions(raw_sec))
    if m is None:
        return None
    code = raw_sec[m.start(2):m.end(2)].strip()
    if re.fullmatch(r"[0-9A-Fa-f]+", code):
        return int(code, 16)
    from .numfmt_names import TAGS
    low = code.lower()
    if low in TAGS:
        return TAGS[low]
    if low in ("x-sysdate", "x-systime"):
        return 0xF800
    return code


def _names(lcid):
    """(days, short days, months, short months, (AM, PM) or None) for a
    language, or None when Excel's names for it were not measured."""
    from .numfmt_names import AMPM, NAMES
    english = (tuple(_DAYS), tuple(d[:3] for d in _DAYS), tuple(_MONTHS),
               tuple(m[:3] for m in _MONTHS), ("AM", "PM"))
    if lcid is None:
        return english
    if isinstance(lcid, str):
        return None
    if lcid in (0xF800, 0xF400):
        # the system long date and time: Excel follows the computer's
        # regional settings; these are this server's English ones
        return english
    if lcid >> 16:
        return None  # another calendar or digit set
    if lcid in NAMES:
        return NAMES[lcid] + (AMPM.get(lcid),)
    if lcid == 0 or lcid & 0x3FF == 0x09:
        return english  # every English locale names days and months alike
    return None


def _era(d: _dt.datetime):
    from .numfmt_names import ERAS
    found = ERAS[0]
    for era in ERAS:
        if (d.year, d.month, d.day) >= era[0]:
            found = era
    return found, d.year - found[0][0] + 1


def _unit(kind: str) -> str:
    """The date unit a lower-cased token names (h, m, s, d, y; elapsed [h]
    is h), or "" for a literal."""
    if kind in ("am/pm", "a/p"):
        return ""
    k = kind[1:] if kind.startswith("[") else kind
    return k[0] if k and k[0] in "hsdmy" else ""


def _clock_ticks(serial: float, scale: int) -> tuple[int, int]:
    """(the serial in 1/scale seconds as Excel's clock shows it, the whole
    seconds an elapsed total shows).

    Half of the last unit shown is added to the day count, then the hours,
    minutes and seconds are peeled off one after another, each step a
    floating-point product, and cut, never rounded; the fraction of a
    second is read to 15 significant digits before it is cut, and so are
    the seconds of an elapsed total. MEASURED (a private hidden Excel,
    13,262 made-up cells, most of them exactly on half of the last unit
    shown): this gives Excel's text for every cell under h:mm:ss, mm:ss,
    date and time, mm:ss.0, [ss].0, h:mm:ss.00 and ss.000, and all but three
    under [mm]:ss and [h]:mm:ss, where rounding the stored number itself,
    however precisely, gets about one in ten of them wrong. A lone [ss]
    count stays the irregular case (_date)."""
    if serial < 0:
        ticks = math.floor(serial * (86400 * scale) + 0.5)
        return ticks, int(_round(_dec15(serial * 86400), 0))
    x = serial + 0.5 / (86400 * scale)
    days = math.floor(x)
    r = (x - days) * 24
    h = math.floor(r)
    r = (r - h) * 60
    m = math.floor(r)
    r = (r - m) * 60
    sec = math.floor(r)
    frac = math.floor(float(f"{(r - sec) * scale:.15g}"))
    whole = (days * 86400 + h * 3600 + m * 60) + math.floor(float(f"{r:.15g}"))
    return (days * 86400 + h * 3600 + m * 60 + sec) * scale + frac, whole


def _date(value, sec: str, lcid=None, strict: bool = False,
          date1904: bool = False) -> str | None:
    found = _serial(value, date1904)
    if found is None:
        return None
    serial, shift = found
    if serial >= _PAST_LAST[date1904]:
        # past 31 December 9999: Excel fills the cell with #
        return "#" * 8
    if date1904 and serial < 0:
        # the 1904 system shows a negative date or time as its distance
        # before day 0, with a minus (-1 shows -1904-01-02, -0.25 under
        # [h]:mm shows -6:00; measured)
        inner = _date(-serial, sec, lcid, strict, date1904)
        return None if inner is None else "-" + inner
    tokens = _DATE_TOKENS.findall(sec)
    kinds = [t.lower() for t in tokens]
    units = [_unit(k) for k in kinds]
    ampm = any(k in ("am/pm", "a/p") for k in kinds)
    names = _names(lcid)
    japanese = lcid == 0x411
    if strict:
        why = None
        if names is None and any(k in _NAMED for k in kinds):
            why = "day or month names"
        elif ampm and (names is None or names[4] is None):
            why = "AM/PM text"
        elif any(k[0] in "ge" for k in kinds if k) and not japanese:
            why = "era or era-year text"
        elif any(k[0] == "b" for k in kinds if k):
            why = "Buddhist-calendar year"
        if why is not None:
            raise Unmatched(why)
    if names is None:
        names = _names(None)  # not strict: English, as before
    if names[4] is None:
        names = names[:4] + (("AM", "PM"),)
    # Excel rounds the time to what it shows: the second, or the tenths,
    # hundredths or thousandths a ".0" after the seconds asks for.
    frac_of = {}
    for i, k in enumerate(kinds):
        if re.fullmatch(r"\.0+", k):
            prev = next((u for u in reversed(units[:i]) if u), "")
            if prev == "s":
                frac_of[i] = min(3, len(k) - 1)
    places = max(frac_of.values(), default=0)
    scale = 10 ** places
    dated = [k for k, u in zip(kinds, units) if u]
    if serial < 0 and not (len(dated) == 1 and dated[0].startswith("[")):
        # Excel cannot show a negative date or time (1900 date system) and
        # fills the cell with #: eight at the standard column width. A lone
        # elapsed count ([ss]) is a number and keeps its minus.
        return "#" * 8
    # Excel's clock and its elapsed totals ([h], [mm], [ss]) can differ at
    # an exact half: 59.5 seconds shows 01:59 under [mm]:ss (measured).
    ticks, whole_s = _clock_ticks(serial, scale)
    clock_s = ticks // scale
    if places:
        whole_s = clock_s
    elif len(dated) == 1 and dated[0].startswith("[") and serial >= 0:
        # A lone elapsed count ([ss], [mm], [h]) rounds the day's seconds,
        # read to 15 significant digits, half up: the closest of the rules
        # measured (59.5 shows 60 and 60.5 shows 61, as in Excel), though
        # about one in sixteen exact halves of a second still shows one
        # second off.
        whole_days = math.floor(serial)
        whole_s = whole_days * 86400 + math.floor(
            float(f"{(serial - whole_days) * 86400 + 0.5:.15g}"))
    try:
        d = (_EPOCH_1904 if date1904 else _EPOCH) + \
            _dt.timedelta(days=shift, seconds=clock_s)
    except OverflowError:
        return "#" * 8
    # The day name follows Excel's serial, not the calendar: a real date
    # before 1 March 1900 sits one serial day further out (shift), and Excel
    # names serial 1, 1 January 1900, a Sunday (measured).
    year, month, day = d.year, d.month, d.day
    weekday = (d - _dt.timedelta(days=shift)).weekday()
    days = clock_s // 86400
    if not date1904 and not shift and days <= 60:
        # Excel's calendar starts at 0 January 1900 and counts a 29 February
        # 1900: serial 1 is 1 January 1900, 0.5 is noon on "January 0"
        year, month, day = (1900, 1, days) if days <= 31 else (1900, 2, days - 31)
    out = []
    for i, t in enumerate(tokens):
        low = kinds[i]
        prev = next((u for u in reversed(units[:i]) if u), "")
        nxt = next((u for u in units[i + 1:] if u), "")
        if low.startswith("[") and len(low) > 2:
            n = len(low) - 2
            total = abs(whole_s) // (3600 if low[1] == "h" else 60 if low[1] == "m" else 1)
            out.append(("-" if whole_s < 0 else "") + f"{total:0{n}d}")
        elif i in frac_of:
            digits = frac_of[i]
            out.append("." + f"{(ticks % scale) // 10 ** (places - digits):0{digits}d}")
        elif low == "yyyy":
            out.append(f"{year:04d}")
        elif low == "yy":
            out.append(f"{year % 100:02d}")
        elif low in ("mm", "m") and (prev == "h" or nxt == "s"):
            out.append(f"{d.minute:02d}" if low == "mm" else str(d.minute))
        elif low == "mmmmm":
            out.append(names[2][month - 1][0])
        elif low == "mmmm":
            out.append(names[2][month - 1])
        elif low == "mmm":
            out.append(names[3][month - 1])
        elif low == "mm":
            out.append(f"{month:02d}")
        elif low == "m":
            out.append(str(month))
        elif low == "dddd":
            out.append(names[0][weekday])
        elif low == "ddd":
            out.append(names[1][weekday])
        elif japanese and low[0] in "ge":
            era, year_of_era = _era(_dt.datetime(year, month, max(day, 1)))
            if low[0] == "g":
                out.append(era[1] if len(low) >= 3 else era[2] if len(low) == 2
                           else era[3])
            else:
                out.append(f"{year_of_era:02d}" if len(low) >= 2 else str(year_of_era))
        elif low == "dd":
            out.append(f"{day:02d}")
        elif low == "d":
            out.append(str(day))
        elif low in ("hh", "h"):
            h = (d.hour % 12 or 12) if ampm else d.hour
            out.append(f"{h:02d}" if low == "hh" else str(h))
        elif low in ("ss", "s"):
            out.append(f"{d.second:02d}" if low == "ss" else str(d.second))
        elif low == "am/pm":
            out.append(names[4][0] if d.hour < 12 else names[4][1])
        elif low == "a/p":
            out.append("A" if d.hour < 12 else "P")
        else:
            out.append(_literal(t))
    return "".join(out)


_PATTERN = re.compile(r"[0#?][0#?,]*(?:\.[0#?]*)?(?:[Ee][+-][0#]+)?|\.[0#?]+")
#: a fraction: an optional whole-number part and its gap, then numerator /
#: denominator (placeholders, or a fixed denominator such as 8 or 16)
_FRACTION = re.compile(r"(?:([0#?][0#?,]*)(\s+))?([0#?]+)/([0#?]+|[1-9][0-9]*)")


def _digits(pattern: str, value: int, *, show_zero: bool) -> str:
    """A whole number under integer placeholders: 0 pads with zeros, ?
    with spaces, # with nothing, and a comma adds thousands separators.
    show_zero: a zero still shows when the pattern has no 0."""
    if value == 0 and not show_zero:
        text = ""
    else:
        text = f"{value:,}" if "," in pattern else str(value)
    text = text.rjust(pattern.count("0"), "0")
    return text.rjust(sum(1 for ch in pattern if ch in "0?"), " ")


def _convergent(x: float, max_den: int) -> tuple[int, int]:
    """Excel's fraction for x >= 0 under up to max_den in the denominator:
    the last continued-fraction convergent that fits, not the nearest
    fraction (0.1 under ?/? is 0/1, not 1/9; pi under ??/?? is 22/7)."""
    b = x
    p2, p1, q2, q1 = 0, 1, 1, 0
    p = q = 0
    while q1 < max_den:
        a = math.floor(b)
        p, q = a * p1 + p2, a * q1 + q2
        if b - a < 5e-8:
            break
        b = 1 / (b - a)
        p2, p1, q2, q1 = p1, p, q1, q
    if q > max_den:
        p, q = (p2, q2) if q1 > max_den else (p1, q1)
    return p, q


def _fraction(value: float, sec: str, fm: re.Match, *, signed: bool) -> str:
    """value under a fraction section (# ?/?, ?/?, # ?/8), as Excel shows
    it: a whole number keeps the fraction's room blank under ?, shows 0/1
    under 0 and nothing under #; a mixed fraction keeps the minus of a
    negative value even when it shows as 0."""
    int_pat, gap, num_pat, den_pat = fm.groups()
    prefix, suffix = sec[:fm.start()], sec[fm.end():]
    x = float(_dec15(abs(value)))
    if den_pat[0] not in "0#?":
        q = int(den_pat)
        p = int(_round(_dec15(x) * q, 0))
        den_text = den_pat
    else:
        p, q = _convergent(x, 10 ** len(den_pat) - 1)
        den_text = str(q).ljust(len(den_pat), " ") if "?" in den_pat else str(q)
    if int_pat is None:
        text = _digits(num_pat, p, show_zero=True) + "/" + den_text
        minus = signed and value < 0 and p != 0
    else:
        whole, p = divmod(p, q)
        if p == 0 and "0" not in num_pat:
            text = _digits(int_pat, whole, show_zero=True)
            if "?" in num_pat + den_pat:
                text += " " * (len(gap) + len(num_pat) + 1 + len(den_pat))
        else:
            body = _digits(int_pat, whole, show_zero=False)
            lead = gap if body.strip() else (" " * len(gap) if "?" in num_pat else "")
            text = body + lead + _digits(num_pat, p, show_zero=True) + "/" + den_text
        minus = signed and value < 0
    return ("-" if minus else "") + _literal(prefix) + text + _literal(suffix)


def _interleaved(sec: str, plain: str) -> str | None:
    """The section's digits laid out around the literals between them, as
    Excel lays them out (000-00-0000 shows 123 as 000-00-0123, (000)
    000-0000 shows a phone number, "x"0"y"0 shows 12 as x1y2, 0.0 0 shows
    123.456 as 123.4 6): None when no literal stands between two digit
    placeholders (the contiguous pattern below handles those).

    Whole-number places fill from the right, one digit each; digits beyond
    the places all go to the first one. Decimal places fill from the left.
    A 0 place with no digit left shows 0, a ? shows a space, a # nothing."""
    holders = [i for i, ch in enumerate(plain) if ch in "0#?"]
    if len(holders) < 2:
        return None
    first, last = holders[0], holders[-1]
    point = next((i for i in range(first, last) if plain[i] == "."), None)
    if not any(ch not in "0#?,." for ch in plain[first:last + 1]):
        return None
    if re.search(r"[Ee][+-]", plain[first:last + 1]):
        return None
    # the section's characters between the first and last place, as runs:
    # ("ph", char) for a digit place, ("pt", ".") for the decimal point,
    # ("lit", text) for anything else (quoted text, escapes, spacing codes)
    runs: list[tuple[str, str]] = []
    i = first
    while i <= last:
        ch = plain[i]
        if ch in "0#?":
            runs.append(("ph", ch))
            i += 1
        elif i == point:
            runs.append(("pt", "."))
            i += 1
        elif ch == ",":
            i += 1  # a separator between places: no character of its own
        else:
            j = i
            while j <= last and plain[j] not in "0#?" and j != point and \
                    plain[j] != ",":
                j += 1
            # a quoted run is blanked in plain: take the section's own text
            runs.append(("lit", _literal(sec[i:j])))
            i = j
    return sec[:first], runs, sec[last + 1:]


def _lay_out(value: float, layout, *, signed: bool, percent: int) -> str:
    """Render value through an _interleaved layout."""
    prefix, runs, suffix = layout
    int_places = []
    seen_point = False
    for kind, ch in runs:
        if kind == "pt":
            seen_point = True
        elif kind == "ph" and not seen_point:
            int_places.append(ch)
    dec_places = [ch for k, ch in runs if k == "ph"][len(int_places):]
    # each comma right after the last place scales by a thousand
    scale = len(suffix) - len(suffix.lstrip(","))
    suffix = suffix.lstrip(",")
    d = _dec15(abs(value)) * (100 ** percent)
    if scale:
        d = d.scaleb(-3 * scale)
    body = f"{_round(d, len(dec_places)):f}"
    whole, _, frac = body.partition(".")
    if whole == "0":
        whole = ""
    digits = list(whole)
    # whole-number places, right to left
    filled: dict[int, str] = {}
    ph_idx = [n for n, (k, _c) in enumerate(runs) if k == "ph"][:len(int_places)]
    for n in reversed(ph_idx):
        ch = runs[n][1]
        if digits:
            filled[n] = digits.pop()
        else:
            filled[n] = "0" if ch == "0" else (" " if ch == "?" else "")
    if digits and ph_idx:
        filled[ph_idx[0]] = "".join(digits) + filled[ph_idx[0]]
    # decimal places, left to right; trailing zeros as the place says
    dec_idx = [n for n, (k, _c) in enumerate(runs) if k == "ph"][len(int_places):]
    frac = frac.ljust(len(dec_idx), "0")
    trailing = True
    for n, dg in reversed(list(zip(dec_idx, frac))):
        ch = runs[n][1]
        if trailing and dg == "0" and ch != "0":
            filled[n] = " " if ch == "?" else ""
        else:
            trailing = False
            filled[n] = dg
    out = []
    for n, (kind, text) in enumerate(runs):
        out.append(filled.get(n, "") if kind == "ph" else text)
    text = "".join(out)
    minus = signed and value < 0 and any(v.strip("0 ") for v in filled.values())
    return ("-" if minus else "") + _literal(prefix) + text + _literal(suffix)


def _number(value: float, sec: str, *, signed: bool) -> str:
    """value under one numeric section. signed: put a minus on a negative
    value (the section is the only one; a negative section shows its own
    text, parentheses included, around the absolute value)."""
    plain = _unquoted_positions(sec)
    fm = _FRACTION.search(plain)
    if fm:
        return _fraction(value, sec, fm, signed=signed)
    m = _PATTERN.search(plain)
    if not m:
        # a section of text alone still shows the minus of a negative value
        # it was chosen for (measured: [<=5]"low" shows -7 as -low)
        return ("-" if signed and value < 0 else "") + _literal(sec)
    layout = _interleaved(sec, plain)
    if layout is not None:
        return _lay_out(value, layout, signed=signed,
                        percent=_unquoted(sec).count("%"))
    end = m.end()
    # each comma straight after the digits scales the value by a thousand
    # (#,##0, shows 1234567 as 1,235; 0.0,, shows it as 1.2)
    scale = 0
    while end < len(plain) and plain[end] == ",":
        scale += 1
        end += 1
    prefix, pattern, suffix = sec[:m.start()], sec[m.start():m.end()], sec[end:]
    # Excel puts the minus in front of everything, the prefix included
    # (-$5.00); a negative section shows its own text around the absolute
    # value instead.
    minus = signed and value < 0
    d = _dec15(abs(value))
    # each percent sign scales by a hundred (0.0%% shows 0.125 as 1250.0%%)
    d *= 100 ** _unquoted(sec).count("%")
    sci = re.search(r"[Ee]([+-])", pattern)
    if sci:
        int_pat, _, dec_pat = pattern[:sci.start()].partition(".")
        decimals = sum(1 for ch in dec_pat if ch in "0#?")
        n_int = sum(1 for ch in int_pat if ch in "0#?")
        # ##0.0E+0 is engineering notation: the exponent steps by the
        # number of whole-number places
        step = n_int if n_int > 1 and "#" in int_pat else 1
        if d:
            exp = math.floor(d.adjusted() / step) * step
            mant = _round(d.scaleb(-exp), decimals)
            if mant >= Decimal(10) ** step:
                exp += step
                mant = _round(d.scaleb(-exp), decimals)
            mant_text = f"{mant:f}"
        else:
            exp = 0
            mant_text = "0" * step + ("." + "0" * decimals if decimals else "")
        width = max(1, sum(1 for ch in pattern[sci.end():] if ch == "0"))
        # E+ always shows the exponent's sign, E- only a minus
        sign = "-" if exp < 0 else ("+" if sci.group(1) == "+" else "")
        text = f"{mant_text}{pattern[sci.start()]}{sign}{abs(exp):0{width}d}"
    else:
        int_part, dot, frac_part = pattern.partition(".")
        scale += len(int_part) - len(int_part.rstrip(","))
        int_part = int_part.rstrip(",")
        if scale:
            d = d.scaleb(-3 * scale)
        decimals = sum(1 for ch in frac_part if ch in "0#?")
        min_int = sum(1 for ch in int_part if ch == "0")
        thousands = "," in int_part
        body = f"{_round(d, decimals):f}"
        whole, _, frac = body.partition(".")
        # trailing zero decimals: shown under 0, a space under ?, dropped
        # under # (0.0? shows 2.5 as "2.5 ", 0.# shows 12 as "12.")
        holders = [ch for ch in frac_part if ch in "0#?"]
        kept, trailing = [], True
        for ch, dg in reversed(list(zip(holders, frac))):
            if trailing and dg == "0" and ch != "0":
                kept.append(" " if ch == "?" else "")
            else:
                trailing = False
                kept.append(dg)
        frac = "".join(reversed(kept))
        if whole == "0" and min_int == 0:
            whole = ""
        whole = whole.rjust(min_int, "0")
        if thousands and whole:
            whole = f"{int(whole):,}".rjust(min_int, "0")
        whole = whole.rjust(sum(1 for ch in int_part if ch in "0?"), " ")
        # A dot in the pattern always shows ("0.#" shows 12 as "12.").
        text = whole + ("." + frac if dot else "")
    if minus and text.strip("0.,%E+- "):
        return "-" + _literal(prefix) + text + _literal(suffix)
    return _literal(prefix) + text + _literal(suffix)


def _unquoted_positions(sec: str) -> str:
    """sec with quoted text and escapes blanked (same length), so a digit
    placeholder inside a literal ("No. 0") is not taken for the pattern."""
    def blank(m: re.Match) -> str:
        return " " * len(m.group(0))
    return re.sub(r'"[^"]*"|\\.', blank, sec)


#: A section's condition: [>=1000], [<0], [<>5] ...
_COND = re.compile(r"\[\s*(<=|>=|<>|<|>|=)\s*([+-]?(?:\d+\.?\d*|\.\d+)"
                   r"(?:[eE][+-]?\d+)?)\s*\]")


def _condition(raw: str) -> tuple[str, float] | None:
    m = _COND.search(_unquoted_positions(raw))
    return (m.group(1), float(m.group(2))) if m else None


def _holds(cond: tuple[str, float], v: float) -> bool:
    op, t = cond
    return {"<": v < t, "<=": v <= t, ">": v > t, ">=": v >= t,
            "=": v == t, "<>": v != t}[op]


def _only_negatives(cond: tuple[str, float]) -> bool:
    """A condition only negative numbers meet ([<0], [<-100], [=-3]): its
    section shows a negative value without a minus, as the plain negative
    section does. [<=0] admits zero, and Excel shows -3 under [<=0]0.0;0
    as -3.0 (measured)."""
    op, t = cond
    return (op == "<" and t <= 0) or (op in ("<=", "=") and t < 0)


def _pick(raw_secs: list[str], v: float) -> tuple[int, bool] | None:
    """(the section that formats number v, whether it shows v's minus), as
    Excel picks it; (-1, True) for General, past a lone section of
    conditions; None when a format of conditions has no section for v.

    Without conditions: positive;negative;zero, the negative section showing
    the value without its minus. With conditions (measured in Excel): the
    first section whose condition holds, else the next section without one,
    and a minus on a negative value unless the condition admits no positive
    number ([>=1000]0.0,"K";0 shows 512 as 512 and -1500 as -1500;
    [<0]"neg";[>0]"pos";"zero" shows -1 as neg; [<=5]"low";[>5]"high" shows
    -7 as -low)."""
    n = len(raw_secs)
    conds = [_condition(s) for s in raw_secs[:3]]
    if conds[0] is not None or (n > 1 and conds[1] is not None):
        for i in range(min(n, 3)):
            cond = conds[i]
            if cond is None:
                return i, True
            if _holds(cond, v):
                return i, not _only_negatives(cond)
        if n == 1:
            # Excel reads a lone section of conditions as "[>100]0;General"
            # (its own NumberFormat says so): the rest show as General
            return -1, True
        return None
    if v < 0 and n > 1:
        return 1, False
    if v == 0 and n > 2:
        return 2, False
    return 0, True


#: openpyxl's text for Excel's built-in format 44 (the Currency style) has
#: lost its section separators; this is Excel's own (en-US) text for it.
_OPENPYXL_44 = '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)'
_EXCEL_44 = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
#: openpyxl's text for Excel's built-in format 47 has lost its colon; Excel
#: shows that format as mm:ss.0 (59.5 seconds is 00:59.5).
_OPENPYXL_47 = "mmss.0"
_EXCEL_47 = "mm:ss.0"


def display(value, fmt: str | None, *, date1904: bool = False) -> str:
    """The text Excel shows for value under number format fmt (English day
    and month names where the format's locale was not measured). date1904:
    the workbook counts dates from 1904."""
    return _display(value, fmt, strict=False, date1904=date1904)


def display_checked(value, fmt: str | None, *, date1904: bool = False
                    ) -> tuple[str | None, str | None]:
    """(the text Excel shows, None), or (None, why) when that text depends
    on a locale or calendar whose names this module has not measured: the
    caller then says so instead of showing a guess."""
    try:
        return _display(value, fmt, strict=True, date1904=date1904), None
    except Unmatched as exc:
        tag = _LOCALE_TAG.search(fmt or "")
        code = tag.group(0) if tag else "its locale code"
        return None, f"{exc} for {code}"


class _NoText(Exception):
    """The format could not be applied to the value (shown() says nothing
    then, where display() shows General)."""


#: A fraction of a second a format shows ("ss.00" shows hundredths).
_SECOND_DIGITS = re.compile(r"[sS]\]?(\.0+)")


def needs_stored_digits(value, fmt: str | None, *, date1904: bool = False) -> bool:
    """Whether the text Excel shows for this date, time or duration depends
    on digits of the stored number that the value no longer holds. A reader
    hands such a cell over held to the millisecond; when that lands on
    exactly half of the last unit the format shows (59.5 seconds under
    [ss], where the file may hold 59.49999999999999) and the text differs on
    either side of that half, only the stored number says which side Excel
    shows. The same holds for 28 February 1900: a reader hands both serial
    59 and serial 60 over as that date, and Excel shows serial 60 as its
    own 29 February 1900, a Wednesday (measured)."""
    if not date1904 and isinstance(value, _dt.datetime) and \
            value.date() == _dt.date(1900, 2, 28):
        return True
    if isinstance(value, _dt.timedelta):
        us = abs(value) // _dt.timedelta(microseconds=1)
    elif isinstance(value, (_dt.datetime, _dt.time)):
        us = ((value.hour * 60 + value.minute) * 60 + value.second) * 1_000_000 \
            + value.microsecond
    else:
        return False
    m = _SECOND_DIGITS.search(_unquoted(fmt or ""))
    places = min(3, len(m.group(1)) - 1) if m else 0
    if places >= 3:
        # thousandths shown from a value held to the millisecond: the digits
        # past it decide every cell, so the stored number always does
        return True
    step = 10 ** (6 - places)
    # Excel's clock adds half of the last unit shown and then cuts
    # (_clock_ticks), so the side of an exact half depends on the last
    # digits of the stored number, whichever text the value would give
    return abs(us % step - step // 2) <= 1000


def shown(value, fmt: str | None, *, date1904: bool = False
          ) -> tuple[str | None, str | None]:
    """The text Excel shows, for a read that gives it beside the value:
    (text, None); (None, why) when it depends on a locale or calendar this
    module has not measured (display_checked's why); (None, None) when it
    cannot be given exactly: the format could not be applied (display()
    shows General there), or the value has lost digits the text depends on
    (needs_stored_digits; a reader with the stored number passes that
    instead)."""
    if needs_stored_digits(value, fmt, date1904=date1904):
        return None, None
    try:
        return _display(value, fmt, strict=True, date1904=date1904, fallback=False), None
    except Unmatched as exc:
        tag = _LOCALE_TAG.search(fmt or "")
        code = tag.group(0) if tag else "its locale code"
        return None, f"{exc} for {code}"
    except Exception:  # noqa: BLE001
        return None, None


def _display(value, fmt: str | None, *, strict: bool,
             date1904: bool = False, fallback: bool = True) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    fmt = (fmt or "General").strip() or "General"
    if isinstance(value, _dt.timedelta):
        # the reader returns an elapsed-time cell ([h]:mm:ss) as a duration
        value = value.total_seconds() / 86400
    if fmt == _OPENPYXL_44:
        fmt = _EXCEL_44
    elif fmt == _OPENPYXL_47:
        fmt = _EXCEL_47
    try:
        if fmt.lower() == "general":
            if isinstance(value, (_dt.datetime, _dt.date)):
                d = _as_datetime(value)
                return d.strftime("%Y-%m-%d") if d.time() == _dt.time() \
                    else d.strftime("%Y-%m-%d %H:%M:%S")
            if isinstance(value, _dt.time):
                return value.strftime("%H:%M:%S")
            return general_display(value)
        raw_secs = _sections(fmt)
        secs = [_strip_brackets(s) for s in raw_secs]
        if isinstance(value, str):
            text_sec = secs[3] if len(secs) > 3 else (secs[0] if "@" in _unquoted(secs[0])
                                                      else None)
            if text_sec is None:
                return value
            parts = text_sec.split("@")
            return value.join(_literal(p) for p in parts)
        if isinstance(value, (_dt.datetime, _dt.date, _dt.time)) or _is_date_section(secs[0]):
            out = _date(value, secs[0], _locale(raw_secs[0]), strict,
                        date1904)
            if out is None and not fallback:
                raise _NoText(fmt)
            return out if out is not None else general(value)
        num = float(value)
        picked = _pick(raw_secs, num)
        if picked is None:
            # a format of conditions with no section for the value: Excel
            # fills the cell with # (eight at the standard column width)
            return "#" * 8
        idx, signed = picked
        if idx < 0:
            return general_display(value)
        sec = secs[idx]
        if not sec:
            return ""  # an empty section hides the values it is for
        if idx == 0 and _unquoted(sec).strip() == "@":
            return general(value)
        return _number(num, sec, signed=signed)
    except Unmatched:
        raise
    except Exception:  # noqa: BLE001 (a display never fails a read)
        if not fallback:
            raise
        return general(value)


__all__ = ["display", "display_checked", "general", "general_display",
           "needs_stored_digits", "shown", "Unmatched"]
