"""A defined name is refused exactly when Excel refuses it (measured, not guessed).

Every name below was put through Excel's own Names.Add (Excel 16, a private
desktop, synthetic names only), and the ones that matter for a saved
file were also opened. The rule the tool uses is Excel's: a name that BEGINS with
R, C or RC and a number that is a real row (1 to 1,048,576) or column (1 to
16,384) is refused, unless a period follows the number; a number that is not a
real row or column (R0_x, C16385_x, R999999999999_x, RC0_x) is an ordinary name;
the digits of other scripts count as digits (R then the Arabic-Indic two is
refused), but fullwidth, mathematical, circled and subscript digits do not.
"""
from __future__ import annotations

import unicodedata

import pytest

from xlsx_mcp.core.errors import XlMcpError
from xlsx_mcp.ops import names as _names

#: Names.Add accepted these.
EXCEL_ACCEPTS = [
    'R_x', 'R0_x', 'R1048577_x', 'R00_x', 'R999999999999_x', 'R1048577', 'R0', 'R0x',
    'r0_x', 'R1.x', 'R1.1', 'R0.1', 'R0D0', '_R1_x', '\\R1_x', 'C_x', 'C0_x',
    'C16385_x', 'C0', 'c0_x', 'C0x', 'C1.x', 'C16385.x', 'RC_x', 'RC0_x', 'RC16385_x',
    'Rc0x', 'RCx', 'RC0', 'R0C1_x', 'R0C0_x', 'R1048577C1_x', 'R0C0', 'R0C1',
    'R1048577C16385_x', 'R0C', 'RC0x', 'A_x', 'A1_x', 'A0', 'A0_x', 'XFD1_x', 'XFE1',
    'XFE1_x', 'ZZZ1', 'ZZZ1_x', 'AAAA1', 'A1048576_x', 'A1048577', 'A1048577_x',
    'A1048576x', 'TAX2023x', 'Q21_Name', 'Q1_Sales', 'R\u0660_x', 'RC\u0660_x',
    'Q\u0662_x', 'x\u0662_y', 'A\u0662_x', 'R\uff11_x', 'C\uff11_x', 'R\uff10_x',
    'RC\uff11', 'R\u2082_x', 'RC1.x', 'RC.x', 'R.x', 'R1.', 'R1..x', 'C1.', 'RC1.',
    'R2.C', 'R1.C', 'R1048577.x', 'C0.x', 'R0.x', 'RC0.x', 'R0C1.x', 'R0C1x', 'RC00_x',
    'R00C1_x', 'rc_x', 'R\U0001d7d0_x', 'C\U0001d7d0_x', 'R\u2461_x', 'C\u2461_x',
    'R\u07c2_x', 'R\u0662.x', 'C\u0662.x',
    'R\u0661\u0660\u0664\u0668\u0665\u0667\u0667_x',
    'C\u0661\u0666\u0663\u0668\u0665_x', 'xyz\u0661', 'XFE\u0661', 'A\u0660', 'A\xb2',
    'A\u2082', 'A\uff11', 'Ab\uff11', '\uff21\uff11', 'ab12cd', 'R1.C1_x', 'R1.C1x',
    'R1.A1_x', 'R1.5', 'R1.Rc', 'R1.A', 'R1.x.A1', 'R1.XFE1', 'R1.A0', 'R1.R', 'R1.CC',
    'R1.C0', 'R1.RC', 'R0.A1', 'C0.A1', 'Sheet1.A1', 'a.A1', 'R_x.A1', 'R.A1', 'C.A1',
    'RC.A1', 'Rate.A1', 'Cx.A1',
]

#: Names.Add refused these.
EXCEL_REFUSES = [
    'R1_x', 'R9_x', 'R1048576_x', 'R01_x', 'R001_x', 'R1048576', 'Rx1', 'R1x', 'r1_x',
    'R2D2', 'C1_x', 'C16384_x', 'C01_x', 'C16384', 'C16385', 'c1_x', 'C1x', 'RC1_x',
    'RC16384_x', 'rc1_x', 'RC', 'RC1', 'R1C_x', 'R1C0_x', 'R1C1_x', 'R1C16384_x',
    'R1C16385_x', 'R1048576C1_x', 'R1C1', 'R1C0', 'R2C', 'RC2', 'XFD1', 'TAX2023',
    'R\u0662_x', 'C\u0662_Name', 'R\u0662C\u0662_x', 'R1\u0662_x', 'RC\u0662_x',
    'R\u0662', 'C\u0662', 'R1C\u0662_x', 'R\u0662C1_x', 'x\u0662', 'A\u0662',
    'r\u0662_x', 'c\u0662_x', 'R\u0968_x', 'C\u0968_x', 'R\xb2_x', 'C\xb2_x',
    'R\u0662\u0663_x', 'R2\u0663_x', 'R\u06612_x', 'R1C1.x', 'R1C.x', 'R1.C1', 'R1C1.',
    'R1C1_x.y', 'C1.R1', 'R1C0.x', 'R1C16385.x', 'R1C00_x', 'rC1_x', 'Rc1_x', 'R1c_x',
    'R\u09e8_x', 'C\u09e8_x', 'R\u0e52_x', 'C\u0e52_x', 'R\xb3_x', 'R\xb9_x',
    'C\xb3_x', 'R\u06f2_x', 'C\u06f2_x', 'RC\u09e8_x', 'RC\xb2_x', 'RC\u06f2_x',
    'R\xb2.x', 'R\u0660\u0662_x', 'R\u0662\u0660_x', 'R1\u0660_x', 'R\u06601_x',
    'R\u0661\u0660\u0664\u0668\u0665\u0667\u0666_x',
    'C\u0661\u0666\u0663\u0668\u0664_x', 'x\u0661', 'XFD\u0661', 'AB\u0662\u0663',
    'a1', 'Q\u0662', 'Q\u0662\u0663', 'R', 'C', 'r', 'c', 'rc', 'Rc', 'rC', 'R1.R1',
    'RC1.A1', 'C1.A1', 'R1.AB12', 'R1.A1.x', 'R1.C1.x', 'R1.D5', 'R1.XFD1', 'R1.Q1',
    'R2.D2', 'R10.A1', 'C1.C1', 'RC1.C1', 'R1.C16385', 'R1.R1C1', 'A1.A1', 'Q1.Q1',
    'R\xb2.A', 'R\xb2_', 'C\xb2.x', 'RC\xb2.x', 'R\u0662.A1', 'R\xb9.x', 'R\xb3.x',
]

#: Every character with a digit value 1 to 9 for which "R" + character + "_x"
#: is refused by Excel (code points in hex; the same set refuses after "C").
EXCEL_DIGITS_AFTER_R = {int(h, 16) for h in (
    "31 32 33 34 35 36 37 38 39 B2 B3 B9 661 662 663 664 665 666 667 668 669 6F1 "
    "6F2 6F3 6F4 6F5 6F6 6F7 6F8 6F9 967 968 969 96A 96B 96C 96D 96E 96F 9E7 9E8 "
    "9E9 9EA 9EB 9EC 9ED 9EE 9EF A67 A68 A69 A6A A6B A6C A6D A6E A6F AE7 AE8 AE9 "
    "AEA AEB AEC AED AEE AEF B67 B68 B69 B6A B6B B6C B6D B6E B6F BE7 BE8 BE9 BEA "
    "BEB BEC BED BEE BEF C67 C68 C69 C6A C6B C6C C6D C6E C6F CE7 CE8 CE9 CEA CEB "
    "CEC CED CEE CEF D67 D68 D69 D6A D6B D6C D6D D6E D6F DE7 DE8 DE9 DEA DEB DEC "
    "DED DEE DEF E51 E52 E53 E54 E55 E56 E57 E58 E59 ED1 ED2 ED3 ED4 ED5 ED6 ED7 "
    "ED8 ED9 F21 F22 F23 F24 F25 F26 F27 F28 F29 1041 1042 1043 1044 1045 1046 1047 "
    "1048 1049 17E1 17E2 17E3 17E4 17E5 17E6 17E7 17E8 17E9 1811 1812 1813 1814 "
    "1815 1816 1817 1818 1819 19DA 1A81 1A82 1A83 1A84 1A85 1A86 1A87 1A88 1A89 "
    "1A91 1A92 1A93 1A94 1A95 1A96 1A97 1A98 1A99 A9D1 A9D2 A9D3 A9D4 A9D5 A9D6 "
    "A9D7 A9D8 A9D9 A9F1 A9F2 A9F3 A9F4 A9F5 A9F6 A9F7 A9F8 A9F9 ABF1 ABF2 ABF3 "
    "ABF4 ABF5 ABF6 ABF7 ABF8 ABF9 "
).split()}


@pytest.mark.parametrize("name", EXCEL_ACCEPTS)
def test_a_name_excel_accepts_is_accepted(name):
    assert _names._validate_name(name) == name


@pytest.mark.parametrize("name", EXCEL_REFUSES)
def test_a_name_excel_refuses_is_refused(name):
    with pytest.raises(XlMcpError):
        _names._validate_name(name)


def test_the_digits_excel_reads_after_r_and_c_are_the_digits_the_tool_reads():
    wrong = []
    for cp in range(0x20, 0x110000):
        if 0xD800 <= cp <= 0xDFFF:
            continue
        ch = chr(cp)
        value = unicodedata.digit(ch, None)
        if value is None or value < 1:
            continue
        expect_refused = cp in EXCEL_DIGITS_AFTER_R
        for lead in ("R", "C"):
            try:
                _names._validate_name(lead + ch + "_x")
                refused = False
            except XlMcpError:
                refused = True
            if refused != expect_refused:
                wrong.append((lead, hex(cp), refused))
    assert wrong == []
