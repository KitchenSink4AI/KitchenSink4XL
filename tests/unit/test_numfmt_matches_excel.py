"""Number display matches Excel's own text, cell for cell.

tests/fixtures/numfmt/excel_range_text.json holds a made-up grid of values
under number formats with the text Excel showed for each cell (Range.Text,
read in a private hidden Excel, columns wide enough that nothing showed as
#). The groups:

  * rounding: Excel rounds the value to 15 significant digits, then halves
    away from zero (0.125 under 0.00 is 0.13; 2.675 is 2.68), and each % sign
    scales by a hundred;
  * elapsed time: [h]:mm:ss, [mm]:ss, [ss] count the whole duration
    (1.25 days is 30:00:00), and .0 after the seconds shows their fraction;
  * fractions: # ?/?, ?/?, # ?/8 and the rest, with Excel's own choice of
    fraction (the last continued-fraction convergent that fits, so 0.1 under
    ?/? is 0/1, not the nearer 1/9) and its blanks and signs;
  * scaling commas: each comma after the digits divides by a thousand
    (#,##0, shows 1234567 as 1,235).
"""

from __future__ import annotations

import json
from pathlib import Path

import openpyxl
import pytest

import family_harness as fh
from xlsx_mcp.core import numfmt

ORACLE = json.loads((Path(__file__).resolve().parents[1] / "fixtures" / "numfmt"
                     / "excel_range_text.json").read_text("utf-8"))


@pytest.mark.parametrize("group", sorted(ORACLE["groups"]))
def test_display_matches_excels_range_text(group):
    cells = ORACLE["groups"][group]
    wrong = [(v, f, want, numfmt.display(v, f)) for v, f, want in cells
             if numfmt.display(v, f) != want]
    assert not wrong, (f"{len(wrong)} of {len(cells)} cells differ from Excel "
                       f"(value, format, Excel, product): {wrong[:6]}")


def _call(tool, op, **args):
    payload, err, _ = fh.start().call(tool, {"op": op, "args": args})
    assert not err, payload
    return payload


def test_a_read_shows_what_excel_shows(tmp_path):
    # the reader hands an elapsed-time cell over as a duration, which used to
    # come out as "1 day, 6:00:00"
    p = tmp_path / "shown.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    for row, (value, fmt) in enumerate([(1.25, "[h]:mm:ss"), (0.75, "# ?/?"),
                                        (1234567, "#,##0,"), (0.125, "0.00"),
                                        (100.123456, "[h]:mm:ss.00")], start=1):
        ws.cell(row, 1, value).number_format = fmt
    wb.save(p)
    wb.close()
    out = _call("xl_read", "read_range", path=str(p), location="A1:A5",
                values="displayed")
    assert out["values"] == [["30:00:00"], [" 3/4"], ["1,235"], ["0.13"], ["2402:57:46.60"]]


def test_date_values_keep_their_own_calendar_date():
    # the reader hands dated cells over as dates and times; a date before
    # 1 March 1900 sits one serial day further out in Excel (its 29 February
    # 1900), and the display must still show the value's own date
    import datetime as dt
    cases = [(dt.datetime(1900, 1, 1), "d-mmm-yy", "1-Jan-00"),
             (dt.datetime(1900, 2, 28, 12, 0), "yyyy-mm-dd hh:mm", "1900-02-28 12:00"),
             (dt.datetime(1987, 6, 5, 18, 0), "h:mm AM/PM", "6:00 PM"),
             (dt.time(12, 30, 0), "hh:mm:ss", "12:30:00"),
             (dt.date(1987, 6, 5), "yyyy-mm-dd", "1987-06-05"),
             (dt.timedelta(days=1, hours=6), "[h]:mm:ss", "30:00:00")]
    for value, fmt, want in cases:
        assert numfmt.display(value, fmt) == want, (value, fmt)
