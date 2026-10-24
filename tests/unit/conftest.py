"""Shared helpers for the unit suite.

The read surfaces send labels SPARSELY (grouped by label, listing only the
cells that are not ordinary) and get_cells sends its rows as arrays named by
a `fields` list. Both encodings are lossless against the dense forms they
replaced; these helpers read them back so a test can go on asserting about
"the label of B4" without restating the encoding every time.
"""

from __future__ import annotations

import pytest


@pytest.fixture
def rewrite_route(monkeypatch):
    """Force every save in the test through the full openpyxl rewrite.

    Cell edits now keep the rest of the workbook byte-for-byte
    (core/surgical.py), so they no longer meet the hazard gate at all. The
    gate still decides every structural edit, and the tests that pin its
    refusals, warnings and verify backstops were written with cell edits;
    this fixture keeps them testing the rewrite route they were written
    for. tests/unit/test_preserve_routes.py covers what a cell edit now does
    on the same workbooks."""
    monkeypatch.setenv("KS4XL_SURGICAL_SAVE", "off")


def label_at(payload: dict, addr: str) -> str:
    """The label a read payload carries for one address."""
    for label, addrs in (payload.get("labels") or {}).items():
        if addr in addrs:
            return label
    return payload.get("labels_default", "value")


def dense_labels(payload: dict, rows: int, cols: int,
                 min_row: int = 1, min_col: int = 1) -> list[list[str]]:
    """Rebuild the dense label matrix a read_range payload used to send."""
    from openpyxl.utils import get_column_letter
    return [
        [label_at(payload, f"{get_column_letter(min_col + c)}{min_row + r}")
         for c in range(cols)]
        for r in range(rows)
    ]


def cell_rows(payload: dict) -> list[dict]:
    """A get_cells payload back as the row dicts it used to send:
    {sheet, cell, value, label} per requested cell."""
    fields = payload["fields"]
    sheet = payload.get("sheet")
    out = []
    for row in payload["cells"]:
        item = dict(zip(fields, row))
        addr = item["cell"]
        item["label"] = label_at(payload, addr)
        if sheet is not None:
            item["sheet"] = sheet
        elif "!" in addr:
            qualified, bare = addr.rsplit("!", 1)
            item["sheet"] = qualified.strip("'").replace("''", "'")
            item["cell"] = bare
        out.append(item)
    return out
