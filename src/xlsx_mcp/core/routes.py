"""core/routes.py: which save route each op takes, in one place.

The save path decides the route by what an edit changed (core/package.py
_route_gate, core/surgical.py plan_edit): an edit that changes only cells is
spliced into the original package and keeps everything else byte-for-byte;
anything else takes the full rewrite, whose hazard gate refuses a loss
unless allow_loss is passed. diagnose_workbook reports that decision ahead
of time from this table, and tests/unit/test_preserve_routes.py runs every
"cell" op below against a workbook full of things the rewrite would lose and
requires the byte-preserving route, so this table and the save path cannot
drift apart.

Classes:
  cell     the op changes cells only (values, formulas, number formats, cell
           formatting, comments): byte-preserving route.
  rewrite  the op changes sheet or workbook structure: full rewrite, hazard
           gate, carry-forward of what the model does not hold.
  file     the op writes a whole new file (create, copy, export, backups) and
           never edits an existing package in place.
  excel    the op runs inside a private hidden Excel (xl_app), which saves the
           workbook itself.
"""

from __future__ import annotations

CELL = "cell"
REWRITE = "rewrite"
FILE = "file"
EXCEL = "excel"

#: Every mutating op, by save route. A "cell" op with a condition in
#: CELL_CONDITIONS takes the rewrite when the condition says so.
OP_EDIT_CLASS: dict[str, str] = {
    # xl_edit
    "set_cell": CELL, "set_cells": CELL, "write_range": CELL,
    "clear_range": CELL, "copy_range": CELL, "move_range": CELL,
    "apply_edits": CELL, "replace_cells": CELL, "set_formula": CELL,
    "import_data": CELL,
    # xl_format
    "format_cells": CELL, "apply_style": CELL, "copy_format": CELL,
    "manage_comment": CELL, "manage_hyperlink": CELL,
    "set_dimensions": REWRITE, "set_merge": REWRITE,
    "manage_conditional_format": REWRITE,
    "manage_image": REWRITE, "manage_chart": REWRITE, "set_view": REWRITE,
    # xl_structure
    "sort_range": CELL,
    "manage_worksheet": REWRITE, "modify_grid_structure": REWRITE,
    "create_table": REWRITE, "manage_table": REWRITE, "manage_name": REWRITE,
    "set_filter": REWRITE, "clear_filter": REWRITE,
    "manage_data_validation": REWRITE,
    # xl_workbook
    "set_workbook_properties": REWRITE, "set_protection": REWRITE,
    "set_page_layout": REWRITE, "set_header_footer": REWRITE,
    "create_workbook": FILE, "copy_workbook": FILE, "manage_backups": FILE,
    "export_range": FILE, "export_file": FILE,
    # xl_app
    "recalculate": EXCEL, "com_manage_pivot": EXCEL, "com_export_pdf": EXCEL,
    "com_render_sheet": EXCEL, "com_convert_format": EXCEL,
    "com_save_with_password": EXCEL, "com_autofit": EXCEL,
    "com_goal_seek": EXCEL, "com_set_sparkline": EXCEL,
    "com_validate_opens_clean": EXCEL,
}

#: When a "cell" op takes the rewrite anyway. Plain words; diagnose shows them.
CELL_CONDITIONS: dict[str, str] = {
    "apply_style": "a style the workbook does not have yet is added first, "
                   "which needs the full rewrite",
    "manage_comment": "a note on a cell that also has a threaded comment, or "
                      "on a sheet whose legacy drawing holds form controls, "
                      "needs the full rewrite",
    "move_range": "moving a range that holds merged cells needs the full "
                  "rewrite",
    "manage_hyperlink": "changing a link that spans a range of cells needs "
                        "the full rewrite",
}


def ops_in(cls: str) -> list[str]:
    return sorted(op for op, c in OP_EDIT_CLASS.items() if c == cls)


def gate_verdict(path: str) -> dict:
    """What the full rewrite's hazard gate does with this workbook, by
    RUNNING that gate (core/package.py _hazard_gate) on it without saving:
    the structural-edit answer diagnose gives is the gate's own, so the two
    cannot disagree (they once did)."""
    from .errors import HazardRefused
    from .package import WorkbookPackage

    pkg = WorkbookPackage.open(path)
    try:
        pkg.workbook  # the gate's second detector reads the model's load
        try:
            warnings = pkg._hazard_gate(False)
        except HazardRefused as exc:
            detail = dict(getattr(exc, "detail", None) or {})
            lost = detail.get("labels") or detail.get("dropped_extensions") or []
            return {"route": "refuse", "message": str(exc), "would_lose": list(lost)}
        return {"route": "openpyxl", "warnings": list(warnings),
                "carried": sorted(pkg._carried_keys())}
    finally:
        pkg.close()


def predict(report, *, verdict: dict | None = None,
            splice_problem: str | None = None) -> dict:
    """The route each class of op takes on this workbook, with the reason:
    cell edits from whether the worksheets can be spliced
    (surgical.splice_check), structural edits from the hazard gate's own
    verdict (gate_verdict)."""
    from . import hazard as _hazard

    if report.error is not None:
        reason = f"cannot scan package: {report.error}"
        return {"cell_edits": {"route": _hazard.ROUTE_REFUSE, "reason": reason},
                "structural_edits": {"route": _hazard.ROUTE_REFUSE,
                                     "reason": reason}}
    verdict = verdict or {"route": "openpyxl", "warnings": [], "carried": []}
    kept = [h.label for h in report.hazards if h.key in set(verdict.get("carried", ()))]
    if verdict["route"] == "refuse":
        struct = {"route": _hazard.ROUTE_REFUSE,
                  "reason": "the full rewrite would lose "
                            + ", ".join(verdict.get("would_lose") or ["content it cannot carry"])
                            + "; these edits refuse unless allow_loss is passed, "
                              "or can be made in Excel itself",
                  "would_lose": verdict.get("would_lose") or []}
    elif verdict.get("warnings"):
        struct = {"route": _hazard.ROUTE_OPENPYXL,
                  "reason": "the full rewrite proceeds and says what may change: "
                            + "; ".join(verdict["warnings"]),
                  "warnings": verdict["warnings"]}
    else:
        struct = {"route": _hazard.ROUTE_OPENPYXL,
                  "reason": "the full rewrite keeps everything this file holds"
                            + (" (" + ", ".join(kept) + " are put back from the "
                               "original)" if kept else "")}
    if kept:
        struct["carried_over"] = kept
    struct["ops"] = ops_in(REWRITE)
    if splice_problem is None:
        cell = {"route": _hazard.ROUTE_RAW_OOXML,
                "reason": "cell edits write only the cells that change; every "
                          "other part of the file is kept byte-for-byte, so "
                          "nothing is at risk and formula results the edit "
                          "does not affect keep their cached values",
                "ops": ops_in(CELL)}
        conditions = {k: v for k, v in CELL_CONDITIONS.items() if v}
        if conditions:
            cell["except_when"] = conditions
    else:
        cell = dict(struct)
        cell["ops"] = ops_in(CELL)
        cell["reason"] = (f"cell edits cannot be spliced into this file "
                          f"({splice_problem}), so they take the full "
                          f"rewrite: " + struct["reason"])
    return {"cell_edits": cell, "structural_edits": struct}


__all__ = ["OP_EDIT_CLASS", "CELL_CONDITIONS", "CELL", "REWRITE", "FILE",
           "EXCEL", "ops_in", "predict", "gate_verdict"]
