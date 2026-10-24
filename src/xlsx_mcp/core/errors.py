"""Exception types for xlsx-mcp (KitchenSink4XL).

Ported from KitchenSink4Word core/errors.py and extended for the grid
domain. Every tool call maps these to actionable messages; envelope.py turns
them into the closed-code refusal shape.

The base class is XlMcpError. The Word taxonomy is carried over (renamed
Workbook/sheet where the Word wording was document-specific) plus the
grid-specific additions: HazardRefused, FormulaRejected, CalcUnavailable.
"""

from __future__ import annotations


class XlMcpError(Exception):
    """Base class; message text is user-facing."""


# ------------------------------------------------------- file / package

class WorkbookNotFound(XlMcpError):
    pass


class WorkbookLocked(XlMcpError):
    """File is open in Excel (or another process holds a lock).

    The empirically confirmed case: a book held
    open in Excel raises PermissionError on the openpyxl write path and
    Excel drops a ~$name.xlsx owner lockfile. Maps to WORKBOOK_LOCKED. Excel
    has no live route: every edit, xl_app's included, needs the file closed;
    reads, exports and copies use the copy last saved to disk."""


class WorkbookCorrupt(XlMcpError):
    """File is not a valid .xlsx/.xlsm package (BadZipFile and kin)."""


class UnsafePackage(WorkbookCorrupt):
    """A package this server will not read: a part declares a document type
    (<!DOCTYPE ...>), its parts would inflate far past what they store (a
    decompression bomb), or it names a part twice; or a picture part whose
    name cannot be a plain file name, which extract refuses. Refused before
    any part is parsed or inflated (core/xmlsafe.py, core/zipsafe.py);
    UNSUPPORTED_CONTENT, as every package the server cannot read."""


class WorkbookProtected(XlMcpError):
    """File is encrypted / password-protected (read needs decryption)."""


# ------------------------------------------------------------- location

class TargetNotFound(XlMcpError):
    """A cell, range, name, table, sheet, or object the tool was told to
    act on does not exist."""


class AmbiguousTarget(XlMcpError):
    """More than one match for a search or name selector; the caller must
    disambiguate. Carries every match on .matches for the refusal shape."""


class RangeOutOfBounds(XlMcpError):
    """A range is inverted or exceeds the grid limits (1,048,576 rows x
    16,384 columns) or the 32,767-char cell ceiling. Maps to
    RANGE_OUT_OF_BOUNDS; the message names the valid bounds."""


class CeilingExceeded(RangeOutOfBounds):
    """The call covers more cells than a write, scan or read ceiling allows.
    Still RANGE_OUT_OF_BOUNDS, with its own hint: the
    range is valid, only too big for one call."""


class StaleAnchor(XlMcpError):
    """A get_grid_view anchor no longer resolves; the sheet changed since
    the view was taken. Maps to STALE_ANCHOR (the view/batch layer)."""


class UnsupportedStructure(XlMcpError):
    """A grid topology or XML shape we refuse to guess about (conservative
    mode). Maps to UNSUPPORTED_CONTENT."""


# ------------------------------------------------------ safety / verify

class ValidationFailed(XlMcpError):
    """Verify-after-write caught a problem; the original file was NOT
    modified (the mutation wrote to a temp path and was not promoted).
    Maps to VALIDATION_FAILED."""


class ExcelWouldRefuse(XlMcpError):
    """A STATIC refuse-class check caught a value Excel itself will not
    accept: a comment past 32,767 characters, a data-validation formula past
    the length where the file stops opening, a header/footer string past 255,
    an invalid hyperlink authority, a negative image extent, a control
    character or lone surrogate in cell text.

    These all used to sail through: verify-after-write proves the package is
    valid OOXML and reads back as intended, which a workbook Excel refuses to
    open can be. Every threshold is measured against Excel (core.limits).
    Maps to the existing BAD_PARAMS code -- a refuse-class value is a bad
    parameter, and the closed vocabulary does not grow for this."""


class HazardRefused(XlMcpError):
    """The round-trip hazard scan found unpreservable parts and no safe
    route, and no allow_loss override was passed. Carries the named parts
    and the available routes on .detail. Maps to HAZARD_REFUSED. This is the
    loud refusal that never repeats the incumbent's silent loss."""


class FormulaRejected(XlMcpError):
    """The formula-injection lint or the unsafe-function policy blocked a
    write, or a recalculation or goal seek would reach the network
    (core.netreach); the message names the function or data connection and
    the override. Maps to FORMULA_REJECTED."""


class CalcUnavailable(XlMcpError):
    """A recalculation was requested and Excel is not available, or the
    removed engine='formulas' was asked for. Maps to CALC_UNAVAILABLE."""


# --------------------------------------------------------- COM / live

class ExcelNotRunning(XlMcpError):
    """No attachable Excel instance (com_ tools need one)."""


class WorkbookNotOpenInExcel(XlMcpError):
    """A live/COM tool targeted a workbook that is not open in the running
    Excel."""


class ExcelBusy(XlMcpError):
    """Excel rejected the call (modal dialog, or a running command)."""


class ExcelBlocked(XlMcpError):
    """Excel is not answering at all (long synchronous operation)."""


class SurfaceRefusal(XlMcpError):
    """A refusal the family surface raises itself: it
    carries a declared code from the closed vocabulary, optional structured
    details and an optional hint. NOT_AVAILABLE (launch settings), NOT_FOUND
    (no such op), BAD_PARAMS (shape, strict args, wrong tool), TIMEOUT (the
    call's deadline), LIMIT_EXCEEDED (the output cap) and INTERNAL_ERROR."""

    def __init__(self, code: str, message: str, *, details: dict | None = None,
                 hint: str | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}
        self.hint = hint


class ExcelDisconnected(XlMcpError):
    """Excel or the workbook closed mid-call; the edit may be partially
    applied."""
