"""ops/properties.py: core document properties + the calc-settings surface.

set_workbook_properties covers the File > Info core properties (title,
author, subject, keywords, category, comments) and the workbook calc
settings the file tier owns: calculation mode and
the fullCalcOnLoad flag. Called with no settings it is a read-only report
of both. The manual-mode caveat is stated wherever it applies: under
calc_mode='manual' Excel will NOT refresh cached results on open, so every
non-Excel read of formula cells returns whatever cache last existed.

Mutation routes through WorkbookPackage (hazard gate, backup, atomic
verified save). openpyxl round-trips coreProps and calcPr natively.
"""

from __future__ import annotations

from typing import Any

from ..core import authorship as _authorship
from ..core.errors import XlMcpError
from ..core.package import WorkbookPackage
from . import gridio

from ..core import vocab as _vocab

#: snake_case values; autoNoTable still accepted.
CALC_MODES = tuple(_vocab.CALC_MODES)
_CALC_ALIASES = {"automatic": "auto"}

MANUAL_MODE_CAVEAT = (
    "calc_mode='manual' means Excel will not recalculate on open; cached "
    "results go stale after every formula or value edit until someone "
    "recalculates explicitly, and every non-Excel reader sees the stale "
    "cache")

_PROP_FIELDS = ("title", "author", "subject", "keywords", "category",
                "comments")


def _read_report(path: str) -> dict:
    wb = gridio.open_wb(path)
    try:
        props = wb.properties
        calc = wb.calculation
        mode = getattr(calc, "calcMode", None) or "auto"
        out: dict[str, Any] = {
            "properties": {
                "title": props.title,
                "author": props.creator,
                "subject": props.subject,
                "keywords": props.keywords,
                "category": props.category,
                "comments": props.description,
                "last_modified_by": props.lastModifiedBy,
            },
            "calc": {
                "calc_mode": _vocab.from_file(mode, _vocab.CALC_MODES),
                "full_calc_on_load": bool(
                    getattr(calc, "fullCalcOnLoad", False)),
                "iterative_calc": bool(getattr(calc, "iterate", False)),
                "max_iterations": getattr(calc, "iterateCount", None) or 100,
                "max_change": getattr(calc, "iterateDelta", None) or 0.001,
            },
            "changed": False,
        }
        if mode == "manual":
            out["calc"]["caveat"] = MANUAL_MODE_CAVEAT
        return out
    finally:
        wb.close()


def set_workbook_properties(path: str, title: str | None = None,
                            author: str | None = None,
                            subject: str | None = None,
                            keywords: str | None = None,
                            category: str | None = None,
                            comments: str | None = None,
                            calc_mode: str | None = None,
                            full_calc_on_load: bool | None = None,
                            iterative_calc: bool | None = None,
                            max_iterations: int | None = None,
                            max_change: float | None = None,
                            allow_loss: bool = False,
                            backup: bool = True,
                            verify_com: bool | None = None) -> dict:
    """Set core document properties and/or calc settings; only the given
    parameters change. With nothing given, returns the current values
    read-only. One backup + one verified save when mutating."""
    given = {"title": title, "author": author, "subject": subject,
             "keywords": keywords, "category": category,
             "comments": comments}
    settings = {k: v for k, v in given.items() if v is not None}
    calc_given = [v for v in (calc_mode, full_calc_on_load, iterative_calc,
                              max_iterations, max_change) if v is not None]
    if not settings and not calc_given:
        return _read_report(path)

    mode = None
    if calc_mode is not None:
        mode = _vocab.to_file(_CALC_ALIASES.get(str(calc_mode).lower(), calc_mode),
                              _vocab.CALC_MODES, "calc_mode")

    pkg = WorkbookPackage.open(path)
    wb = pkg.workbook
    props = wb.properties
    detail: dict[str, Any] = {}
    notes: list[str] = []  # what an email address sent as the author earns
    for field, value in settings.items():
        if not isinstance(value, str):
            raise XlMcpError(f"{field} must be a string")
        note = None
        if field == "author" and "@" in value:
            # An email address is never written as the author: the name that
            # would sign a note stands in (core.authorship).
            value, source, note = _authorship.signer_with_note(
                value, _authorship.PROPERTY)
        target = {"author": "creator", "comments": "description"}.get(
            field, field)
        setattr(props, target, value)
        detail[field] = value
        if note:
            detail["author_source"] = source
            notes.append(note)
    warnings: list[str] = []
    if mode is not None:
        wb.calculation.calcMode = mode
        detail["calc_mode"] = _vocab.from_file(mode, _vocab.CALC_MODES)
        if mode == "manual":
            warnings.append(MANUAL_MODE_CAVEAT)
    if full_calc_on_load is not None:
        wb.calculation.fullCalcOnLoad = bool(full_calc_on_load)
        detail["full_calc_on_load"] = bool(full_calc_on_load)
    if iterative_calc is not None:
        wb.calculation.iterate = bool(iterative_calc)
        detail["iterative_calc"] = bool(iterative_calc)
        if iterative_calc:
            warnings.append(
                "iterative calculation is on: circular references now "
                "converge instead of erroring, bounded by max_iterations "
                "and max_change; unintended circular references become "
                "silent wrong numbers, so use audit_formulas to check. A "
                "circular formula written here has no seed value, and Excel "
                "iterates from the cell's current value, so run recalculate "
                "(xl_app) after writing one or Excel shows #VALUE! until "
                "the cell is re-entered by hand")
    if max_iterations is not None:
        n = int(max_iterations)
        if not 1 <= n <= 32767:
            raise XlMcpError("max_iterations must be 1 to 32767")
        wb.calculation.iterateCount = n
        detail["max_iterations"] = n
    if max_change is not None:
        d = float(max_change)
        if d <= 0:
            raise XlMcpError("max_change must be a positive number")
        wb.calculation.iterateDelta = d
        detail["max_change"] = d
    pkg._changed["properties"] = detail
    result = pkg.save(allow_loss=allow_loss, backup=backup,
                      verify_com=verify_com)
    result["warnings"] = [*notes, *result.get("warnings", []), *warnings]
    return result


__all__ = ["set_workbook_properties", "CALC_MODES", "MANUAL_MODE_CAVEAT"]
