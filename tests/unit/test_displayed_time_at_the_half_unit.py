"""Displayed time at exactly half of the last unit shown follows Excel.

Excel's text for a time that falls exactly on half of the last unit its
format shows (59.5 seconds under h:mm:ss, 1.25 seconds under mm:ss.0) does
not follow the stored number in one direction: the clock adds half of that
unit to the day count and then peels the hours, minutes and seconds off in
floating point, cutting each. Every expected text below was read from a
private hidden Excel (Range.Text, column 60 wide) on a made-up workbook
written exactly as these tests write it. A lone [ss] count is the one form
that still differs at some exact halves, so it is not pinned here beyond the
measured cases that hold.

File tier only, no Excel. Every workbook is made up and built by code.
"""
from __future__ import annotations

import openpyxl
import pytest

from xlsx_mcp.ops import cells


def _book(tmp_path, value, fmt):
    p = tmp_path / "made.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Shown"
    ws.column_dimensions["A"].width = 60
    ws["A1"] = value
    ws["A1"].number_format = fmt
    wb.save(p)
    return str(p)


# (value, format, Excel's Range.Text)
CASES = [
    (3600.5 / 86400, "[mm]:ss", "60:01"),
    (3600.5 / 86400, "[h]:mm:ss", "1:00:01"),
    (3600.5 / 86400, "mm:ss", "00:01"),
    (3600.5 / 86400, "h:mm:ss", "1:00:01"),
    (1.25 / 86400, "mm:ss.0", "00:01.3"),
    (60, "dddd", "Wednesday"),
    (60, "ddd", "Wed"),
    (59, "dddd", "Tuesday"),
    (89.5 / 86400 + 45000, "yyyy-mm-dd hh:mm:ss", "2023-03-15 00:01:29"),
    (3599.5 / 86400 + 45000, "yyyy-mm-dd hh:mm:ss", "2023-03-15 00:59:59"),
    (90061.5 / 86400 + 45000, "yyyy-mm-dd hh:mm:ss", "2023-03-16 01:01:01"),
    (60, "yyyy-mm-dd", "1900-02-29"),
    (59.5 / 86400, "h:mm:ss", "0:00:59"),
    (59.5 / 86400, "[ss]", "60"),
    (59.5 / 86400, "[mm]:ss", "01:59"),
    (45018.02299189815, '[h]:mm:ss', '1080432:33:06'),
    (3.840190972222222, '[h]:mm:ss', '92:09:52'),
    (22.32392939814815, '[h]:mm:ss', '535:46:27'),
    (45011.68424189815, '[h]:mm:ss', '1080280:25:18'),
    (45000.00008680556, '[h]:mm:ss', '1080000:00:07'),
    (45013.50950810185, '[mm]:ss', '64819453:41'),
    (11.79282986111111, '[mm]:ss', '16981:40'),
    (45005.16100115741, '[mm]:ss', '64807431:50'),
    (45017.66450810185, '[mm]:ss', '64825436:53'),
    (0.0005497685185185186, '[mm]:ss', '00:47'),
    (45001.225202546295, '[ss]', '3888105857'),
    (45010.79282986111, '[ss]', '3888932500'),
    (16.643605324074073, '[ss]', '1438007'),
    (45007.74692708333, '[ss]', '3888669334'),
    (45019.117193287035, '[ss]', '3889651725'),
    (0.0006649305555555555, '[ss].0', '57.4'),
    (0.0006626157407407407, '[ss].0', '57.2'),
    (0.0007112268518518519, '[ss].0', '61.4'),
    (0.0006163194444444444, '[ss].0', '53.2'),
    (2.1996695601851854, '[ss].0', '190051.4'),
    (0.9465219907407407, 'h:mm', '22:42'),
    (45000.04166087963, 'h:mm', '0:59'),
    (0.9465219907407407, 'h:mm AM/PM', '10:42 PM'),
    (45000.04166087963, 'h:mm AM/PM', '12:59 AM'),
    (45000.00075810185, 'h:mm:ss', '0:01:05'),
    (2.1973668981481485, 'h:mm:ss', '4:44:12'),
    (11.79282986111111, 'h:mm:ss', '19:01:40'),
    (45001.58972800926, 'h:mm:ss', '14:09:12'),
    (15.82177662037037, 'h:mm:ss', '19:43:21'),
    (15.795509317129628, 'h:mm:ss.00', '19:05:32.00'),
    (2.199664409722222, 'h:mm:ss.00', '4:47:31.00'),
    (0.9465162615740741, 'h:mm:ss.00', '22:42:59.00'),
    (0.00043987268518518523, 'h:mm:ss.00', '0:00:38.00'),
    (0.001030150462962963, 'h:mm:ss.00', '0:01:29.00'),
    (20.85583912037037, 'mm:ss', '32:24'),
    (45000.00075810185, 'mm:ss', '01:05'),
    (45005.16100115741, 'mm:ss', '51:50'),
    (21.32392939814815, 'mm:ss', '46:27'),
    (16.27941550925926, 'mm:ss', '42:21'),
    (0.0004056712962962963, 'mm:ss.0', '00:35.0'),
    (2.191671875, 'mm:ss.0', '36:00.4'),
    (8.159722222222222e-05, 'mm:ss.0', '00:07.0'),
    (7.82892650462963, 'mm:ss.0', '53:39.2'),
    (0.00034780092592592594, 'mm:ss.0', '00:30.1'),
    (0.0004282465277777778, 'ss.000', '37.001'),
    (0.0004976909722222223, 'ss.000', '43.001'),
    (0.0006250057870370371, 'ss.000', '54.001'),
    (5.787037037037037e-09, 'ss.000', '00.001'),
    (0.00046296875, 'ss.000', '40.001'),
    (45000.0005150463, 'yyyy-mm-dd hh:mm:ss', '2023-03-15 00:00:44'),
    (45012.01999421296, 'yyyy-mm-dd hh:mm:ss', '2023-03-27 00:28:47'),
    (15.28765625, 'yyyy-mm-dd hh:mm:ss', '1900-01-15 06:54:13'),
    (0.00036458333333333335, 'yyyy-mm-dd hh:mm:ss', '1900-01-00 00:00:32'),
    (10.87904513888889, 'yyyy-mm-dd hh:mm:ss', '1900-01-10 21:05:49'),
]


@pytest.mark.parametrize("value,fmt,excel", CASES, ids=[f"{c[1]}-{c[0]!r}" for c in CASES])
def test_displayed_text_at_a_half_unit_is_excels(tmp_path, value, fmt, excel):
    p = _book(tmp_path, value, fmt)
    shown = cells.read_range(p, "A1", sheet="Shown", values="displayed")["values"][0][0]
    default = cells.read_range(p, "A1", sheet="Shown")
    told = default.get("shows", {}).get("A1", str(default["values"][0][0]))
    assert (shown, told) == (excel, excel)
