"""Displayed text follows Excel's sections and digit places.

Reads that return displayed text (values="displayed") applied a format of
conditions in the wrong section (512 under [>=1000000]0.0,,"M";[>=1000]
0.0,"K";0 showed 0.0M; Excel shows 512), read the digits of a pattern with
literals between them as one run (123 under 000-00-0000 showed 123-00-0000;
Excel shows 000-00-0123), and showed an empty section's value. Every
expected text below is Excel's own Range.Text for the same value and
format, measured in Excel.
"""

from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.core import numfmt
from xlsx_mcp.ops import cells

K = '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0'
SIGNS = '[<0]"neg";[>0]"pos";"zero"'
BIG = '[>100]"big"0;[<-100]"small"0;0.0'
LOWHIGH = '[<=5]"low";[>5]"high"'

#: (value, format, Excel's text)
EXCEL = [
    (512, K, "512"), (1500, K, "1.5K"), (2500000, K, "2.5M"), (-5, K, "-5"),
    (0, K, "0"), (-1500, K, "-1500"),
    (-1, SIGNS, "neg"), (1, SIGNS, "pos"), (0, SIGNS, "zero"),
    (150, BIG, "big150"), (-150, BIG, "small150"), (5, BIG, "5.0"),
    (-5, BIG, "-5.0"),
    (0.5, "[<1]0.00%;0", "50.00%"), (3, "[<1]0.00%;0", "3"),
    (7, LOWHIGH, "high"), (3, LOWHIGH, "low"), (-7, LOWHIGH, "-low"),
    (123, "000-00-0000", "000-00-0123"),
    (123456789, "000-00-0000", "123-45-6789"),
    (5, "00-00", "00-05"), (1234, "#-#", "123-4"),
    (5551234567, "(000) 000-0000", "(555) 123-4567"),
    (12, '"x"0"y"0', "x1y2"), (12345, "0 00 00", "1 23 45"),
    (123.456, "0.0 0", "123.4 6"), (98765, "00000-0", "09876-5"),
    (42, "0;-0;;@", "42"), (0, "0;-0;;@", ""),
    (-1234.5, "#,##0.00;[Red]-#,##0.00", "-1,234.50"),
    (1234567, "#,##0,", "1,235"), (1.5, '0" units"', "2 units"),
    # one section of conditions: Excel reads [>100]0 as [>100]0;General
    (5, "[>100]0", "5"),
    # [<=0] admits zero, so a negative value keeps its minus
    (-3, "[<=0]0.0;0", "-3.0"), (-3, '[=-3]"minus three";0', "minus three"),
    (150, '[>100]"big";[<-100]"small"', "big"),
]


@pytest.mark.parametrize("value,fmt,excel", EXCEL,
                         ids=[f"{v}|{f}" for v, f, _ in EXCEL])
def test_display_matches_excel(value, fmt, excel):
    assert numfmt.display(value, fmt) == excel


def test_a_read_returns_the_same_text(tmp_path):
    p = tmp_path / "fmt.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for i, (v, f, _x) in enumerate(EXCEL, start=1):
        c = ws.cell(i, 1, v)
        c.number_format = f
    wb.save(p)
    got = cells.read_range(str(p), f"A1:A{len(EXCEL)}", sheet="Data",
                           values="displayed")["values"]
    assert [row[0] for row in got] == [x for _v, _f, x in EXCEL]
