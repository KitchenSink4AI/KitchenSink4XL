"""core/readcompat.py: read what Excel reads.

Some workbooks hold markup Excel opens clean that openpyxl's reader cannot
load at all, so every read and edit of them failed (UNSUPPORTED_CONTENT, or
"could not read worksheets"). Each case below was measured in Excel:

  - Workbook or sheet protection whose attribute names differ from the
    schema's in letter case (workbookalgorithmName, workbookhashValue ...).
    Excel reads them as the protection they name: the structure is
    protected and the published password unprotects it. openpyxl raised
    TypeError on the unexpected keyword. The attributes are now matched
    without regard to case, and an attribute neither reader knows is left
    out of the model (Excel ignores it too).
  - A TRUE/FALSE cell stored as <v>true</v> or <v>false</v> (Excel reads
    them as TRUE and FALSE; openpyxl's int() raised ValueError and the whole
    sheet failed to load).
  - An error value stored in any letter case (<v>#n/a</v>): Excel shows
    #N/A, and so does the read now.
  - A shared formula whose first cell is missing, or comes after a cell that
    uses it: openpyxl read the formula "=", Excel drops the formula and
    keeps the stored value, and the read now returns that value
    (core.sharedformulas names the cells).
  - A number under a date format that is past Excel's last date (31
    December 9999): openpyxl read it as the error #VALUE!, though the cell
    holds the number (Excel fills the cell with #); it reads as the number.

install() patches openpyxl's reader once per process (idempotent).
"""

from __future__ import annotations

_ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!",
           "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!",
           "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!", "#PYTHON!",
           "#EXTERNAL!")
_ERROR_BY_UPPER = {e.upper(): e for e in _ERRORS}

_INSTALLED = False


def _case_tolerant(cls) -> None:
    """from_tree for a protection class that matches attribute names without
    regard to case and leaves out ones the class does not know."""
    original = cls.from_tree.__func__
    known = {a.lower(): a for a in getattr(cls, "__attrs__", ())}

    def from_tree(klass, node):
        attrib = node.attrib
        fixed = {}
        changed = False
        for key, value in attrib.items():
            if key.startswith("{") or key in klass.__attrs__:
                fixed[key] = value
                continue
            canon = known.get(key.lower())
            changed = True
            if canon is not None and canon not in attrib:
                fixed[canon] = value
        if not changed:
            return original(klass, node)
        import copy
        clone = node.makeelement(node.tag, fixed)
        for child in node:
            clone.append(copy.deepcopy(child))
        return original(klass, clone)

    cls.from_tree = classmethod(from_tree)


def install() -> None:
    global _INSTALLED
    if _INSTALLED:
        return
    _INSTALLED = True
    from openpyxl.workbook.protection import WorkbookProtection
    from openpyxl.worksheet.protection import SheetProtection
    from openpyxl.worksheet import _reader as _wsreader
    from openpyxl.xml.constants import SHEET_MAIN_NS

    for cls in (WorkbookProtection, SheetProtection):
        _case_tolerant(cls)

    value_tag = "{%s}v" % SHEET_MAIN_NS
    parse_cell = _wsreader.WorkSheetParser.parse_cell

    def tolerant_parse_cell(self, element):
        ctype = element.get("t")
        if ctype == "b":
            v = element.find(value_tag)
            if v is not None and v.text is not None:
                word = v.text.strip().lower()
                if word in ("true", "false"):
                    v.text = "1" if word == "true" else "0"
        cell = parse_cell(self, element)
        if cell.get("data_type") == "f" and cell.get("value") == "=" \
                and not self.data_only:
            # A shared formula whose first cell is missing (or comes after
            # this cell): no formula can be rebuilt. Excel drops the formula
            # and keeps the stored value, and so does this read (the cells
            # are named by core.sharedformulas, which the read tools flag).
            self.data_only = True
            try:
                cell = parse_cell(self, element)
            finally:
                self.data_only = False
        if ctype in (None, "n") and cell.get("data_type") == "e" \
                and cell.get("value") == "#VALUE!":
            # a number in a date format past 31 December 9999 (or before
            # the calendar): openpyxl calls it an error, but the cell holds
            # the number, and Excel's value is that number
            v = element.find(value_tag)
            try:
                cell["value"] = float(v.text)
                cell["data_type"] = "n"
            except (AttributeError, TypeError, ValueError):
                pass
        if ctype == "e":
            value = cell.get("value")
            if isinstance(value, str):
                cell["value"] = _ERROR_BY_UPPER.get(value.strip().upper(),
                                                    value)
        return cell

    _wsreader.WorkSheetParser.parse_cell = tolerant_parse_cell


__all__ = ["install"]
