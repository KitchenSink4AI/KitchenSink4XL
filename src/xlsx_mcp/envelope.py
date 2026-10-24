"""The response envelope: every tool answer leaves the server through here.

The family surface, adapted from the Word family-surface envelope (itself
adapted from the PDF server's envelope.py) while keeping Excel's own refusal
shape and closed vocabulary:

1. EMITTED ONCE, COMPACT. A success or a refusal is serialized one time, as
   compact JSON, and sent as the single text content block. No
   structuredContent copy rides beside it and no tool declares an
   outputSchema. (The released server already sent successes once; its
   refusals still carried a structuredContent copy through RefusalResult.)
2. THE SHAPE IS EXCEL'S. Refusals stay {ok: false, error: {code, message,
   hint, ...}} with the closed code list below (the grid additions
   WORKBOOK_LOCKED, HAZARD_REFUSED, FORMULA_REJECTED, CALC_UNAVAILABLE kept);
   `matches` (locate ambiguity), `detail` and `details` ride along when
   present. A control character in a message is never echoed raw.
3. INTERNAL_ERROR IS THE FALLBACK. An exception outside CODE_MAP is ours,
   never the caller's: it becomes INTERNAL_ERROR with its library text
   scrubbed of local folders and clipped, instead of escaping as a raw
   FastMCP tool error (the released server fell back to BAD_PARAMS, which
   blamed the caller for a server fault).

The pack signpost (pack_hint, DisabledToolSignpost), RefusalResult and the
fastmcp InputValidationEnvelope are gone: packs no longer exist and the
dispatcher validates arguments itself (strict), so fastmcp never sees
a tool's arguments. A refusal whose raise site declares `hint_tools` now
carries `details.see_ops`: each named op with the tool that holds it.
"""

from __future__ import annotations

import json as _json
import os
import re as _re
import zipfile
from typing import Any

from openpyxl.utils.exceptions import (
    IllegalCharacterError as _IllegalCharacterError,
)
from openpyxl.utils.exceptions import (
    InvalidFileException as _InvalidFileException,
)

from .core import errors as _err
from .core.safesave import MutationLockTimeout
from .core.xmlsafe import XML_ERRORS as _xml_errors
from .core.xmlsafe import ParseError as _XmlParseError
from .core.sandbox import SandboxViolation

# Order matters: specific classes before their XlMcpError base. Unchanged
# from the released server (engine behaviour).
CODE_MAP: tuple[tuple[type[BaseException], str], ...] = (
    (_err.AmbiguousTarget, "AMBIGUOUS_LOCATION"),
    (_err.StaleAnchor, "STALE_ANCHOR"),
    (_err.RangeOutOfBounds, "RANGE_OUT_OF_BOUNDS"),
    (_err.TargetNotFound, "NOT_FOUND"),
    (_err.WorkbookNotFound, "NOT_FOUND"),
    (MutationLockTimeout, "WORKBOOK_LOCKED"),
    (_err.WorkbookLocked, "WORKBOOK_LOCKED"),
    (_err.WorkbookCorrupt, "UNSUPPORTED_CONTENT"),
    (_err.WorkbookProtected, "UNSUPPORTED_CONTENT"),
    (_err.UnsupportedStructure, "UNSUPPORTED_CONTENT"),
    (_err.ValidationFailed, "VALIDATION_FAILED"),
    # Grid-specific additions.
    (_err.HazardRefused, "HAZARD_REFUSED"),
    (_err.FormulaRejected, "FORMULA_REJECTED"),
    (_err.CalcUnavailable, "CALC_UNAVAILABLE"),
    (_err.ExcelNotRunning, "APP_NOT_RUNNING"),
    (_err.WorkbookNotOpenInExcel, "APP_NOT_RUNNING"),
    (_err.ExcelBusy, "APP_BUSY"),
    (_err.ExcelBlocked, "APP_BLOCKED"),
    (_err.ExcelDisconnected, "CONFLICT"),
    (SandboxViolation, "BAD_PARAMS"),
    (_err.ExcelWouldRefuse, "BAD_PARAMS"),
    (_err.XlMcpError, "BAD_PARAMS"),
    (FileExistsError, "CONFLICT"),
    (FileNotFoundError, "NOT_FOUND"),
    # openpyxl's own refusal for text Excel cannot store (control characters
    # in a cell value); the ops layer refuses these first with a better
    # message (core.limits.check_text_storable). Backstop.
    (_IllegalCharacterError, "BAD_PARAMS"),
    # A file that is not a zip reaches the READ paths as a raw BadZipFile.
    (zipfile.BadZipFile, "UNSUPPORTED_CONTENT"),
    # openpyxl's refusal for a format it does not read (legacy .xls that
    # slipped past the OLE sniff).
    (_InvalidFileException, "UNSUPPORTED_CONTENT"),
    # PermissionError is the file-held-open signal on Windows.
    (PermissionError, "WORKBOOK_LOCKED"),
    (IsADirectoryError, "BAD_PARAMS"),
    (ValueError, "BAD_PARAMS"),
    (TypeError, "BAD_PARAMS"),
    # Parser and recursion errors from hostile XML input refuse in-envelope.
    (_XmlParseError, "BAD_PARAMS"),
    (RecursionError, "UNSUPPORTED_CONTENT"),
    (OverflowError, "BAD_PARAMS"),
    # AttributeError, IndexError and a KeyError that is not
    # a lookup of something the caller named are the server's own faults,
    # never the caller's: they fall through to INTERNAL_ERROR ("the
    # arguments are not the cause") instead of BAD_PARAMS. _lookup_code
    # sorts the KeyErrors (openpyxl's "Worksheet X does not exist" is the
    # caller naming a sheet that is not there: NOT_FOUND).
    # Last-resort OS backstop (a device path, a broken junction, a
    # cross-volume replace).
    (OSError, "BAD_PARAMS"),
)


def _catchable() -> tuple[type[BaseException], ...]:
    """CATCHABLE with lxml's error folded in (core/xmlsafe.py names the XML
    exceptions, so nothing here imports a parser)."""
    types = [t for t, _ in CODE_MAP]
    types.extend(t for t in _xml_errors if t not in types)
    return tuple(types)


CATCHABLE = _catchable()

#: The closed code vocabulary. NOT_AVAILABLE, TIMEOUT, INTERNAL_ERROR and
#: LIMIT_EXCEEDED are the family surface's additions; PROTECTED and
#: GUARD_TRIPPED are the Word editor's, carried over.
CLOSED_CODES = frozenset({
    "AMBIGUOUS_LOCATION", "NOT_FOUND", "WORKBOOK_LOCKED", "APP_NOT_RUNNING",
    "APP_BUSY", "APP_BLOCKED", "VALIDATION_FAILED", "STALE_ANCHOR",
    "RANGE_OUT_OF_BOUNDS", "UNSUPPORTED_CONTENT", "CONFLICT", "BAD_PARAMS",
    "HAZARD_REFUSED", "FORMULA_REJECTED", "CALC_UNAVAILABLE",
    "NOT_AVAILABLE", "TIMEOUT", "INTERNAL_ERROR", "LIMIT_EXCEEDED",
    # The Word editor's codes: PROTECTED (sheet or workbook
    # protection forbids the change) and GUARD_TRIPPED (a safety check
    # the caller passed did not hold).
    "PROTECTED", "GUARD_TRIPPED",
})

#: The shipped hints, carried over. Four mechanical corrections (names
#: only what exists), no new English: NOT_FOUND named `get_workbook_info`,
#: which never existed (the op is get_workbook_metadata); HAZARD_REFUSED and
#: CALC_UNAVAILABLE named "the com pack", which is now the xl_app tool; and
#: WORKBOOK_LOCKED ended with a clause about `live='auto'`, a parameter no
#: Excel op has, which was deleted rather than reworded. A hint also names
#: the tool in front of each op it mentions (xl_read), since
#: on the family surface an op is called through its tool.
_SHIPPED_HINTS: dict[str, str] = {
    "AMBIGUOUS_LOCATION": (
        "several targets matched; use the unambiguous address from the "
        "candidates in the message"
    ),
    "NOT_FOUND": (
        "re-run xl_read get_workbook_metadata or get_grid_view to see current "
        "sheets, the true used range, and anchors"
    ),
    "STALE_ANCHOR": (
        "the sheet changed since the view was taken; re-run xl_read "
        "get_grid_view and resend with fresh anchors"
    ),
    "RANGE_OUT_OF_BOUNDS": (
        "the range is inverted or exceeds the grid; the message names the "
        "valid bounds"
    ),
    "WORKBOOK_LOCKED": (
        "close the file in Excel (or wait out the other process) and "
        "retry"
    ),
    "VALIDATION_FAILED": (
        "the original file was NOT modified; the message says what the "
        "produced package failed"
    ),
    "HAZARD_REFUSED": (
        "the workbook holds parts openpyxl would drop on save; the message "
        "names them, states exactly what each loses, and gives the two real "
        "routes (pass allow_loss:true to accept the loss after a backup, or "
        "leave the workbook to Excel). xl_app cannot perform a cell, "
        "format, or structural write, so it is not a route here"
    ),
    "FORMULA_REJECTED": (
        "the write hit the formula-injection or unsafe-function policy, or a "
        "recalculation or goal seek would reach the network, or its file could "
        "not be read well enough to rule that out; the message names the "
        "function, data connection or unread part and the override"
    ),
    "CALC_UNAVAILABLE": (
        "a recalculation needs Windows with Microsoft Excel (xl_app); reads still return cached values, labelled cached"
    ),
    "APP_NOT_RUNNING": "this operation needs Excel installed and reachable",
    "APP_BUSY": (
        "Excel is showing a dialog or running a command; clear it and retry"
    ),
    "APP_BLOCKED": "Excel is not answering; wait or restart it",
}

from .surface.copy import NEW_HINTS as _NEW_HINTS  # noqa: E402

#: One recovery per code, total over CLOSED_CODES (a test holds it).
#: surface/copy.py wins: a hint there for any code,
#: shipped or new, replaces the shipped line.
HINTS: dict[str, str] = {**_SHIPPED_HINTS, **_NEW_HINTS}

#: The one plain sentence every corrupt-workbook refusal carries. Its tool
#: names are the family-surface calls.
RECOVERY_HINT = (
    "If this workbook was working before, an automatic backup may hold a "
    "good copy: xl_workbook manage_backups(action='list', path=...) shows "
    "the prev and anchor slots, and manage_backups(action='restore', "
    "path=..., source='prev') puts one back. xl_read "
    "get_workflows(task='recover-workbook') walks through it."
)

#: The hint an engine's own BAD_PARAMS refusal carries: its message
#: already names what to change.
ENGINE_REFUSAL_HINT = "The message says what to change; nothing was written."

#: How much library text may ride inside an INTERNAL_ERROR.
DETAIL_CHARS = 200


def _lookup_code(exc: KeyError) -> str:
    """NOT_FOUND for a lookup of something the caller named, UNSUPPORTED_CONTENT
    for a package part every reader needs, INTERNAL_ERROR otherwise."""
    text = str(exc.args[0]) if exc.args else ""
    if "does not exist" in text:
        return "NOT_FOUND"
    if text.strip("'\"").lower().endswith((".xml", ".rels", ".bin")):
        return "UNSUPPORTED_CONTENT"
    return "INTERNAL_ERROR"


def classify(exc: BaseException) -> str:
    if isinstance(exc, KeyError) and not isinstance(exc, _err.XlMcpError):
        return _lookup_code(exc)
    for etype, code in CODE_MAP:
        if isinstance(exc, etype):
            return code
    return "INTERNAL_ERROR"


def _declared_code(exc: BaseException) -> str | None:
    """A raise-site code override, but only when it is one of OUR closed
    codes (xml.etree's ParseError carries an integer expat `.code`)."""
    code = getattr(exc, "code", None)
    return code if isinstance(code, str) and code in CLOSED_CODES else None


#: Control characters must never be echoed back verbatim.
_CTRL = _re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")


def _safe_message(text: str) -> str:
    return _CTRL.sub(lambda m: f"<U+{ord(m.group(0)):04X}>", text)


_WIN_PATH = _re.compile(r"[A-Za-z]:[\\/][^\s'\"<>|]*")
_POSIX_HOME = _re.compile(r"/(?:home|Users)/[^\s'\"<>|]*")


def scrub(text: Any, limit: int = DETAIL_CHARS) -> str:
    """Bound and de-identify library text: local folder paths go, the text
    is clipped."""
    s = _safe_message(str(text))
    s = _WIN_PATH.sub("<path>", s)
    s = _POSIX_HOME.sub("<path>", s)
    for var in ("USERPROFILE", "HOME", "LOCALAPPDATA", "APPDATA", "TEMP", "TMP"):
        value = os.environ.get(var)
        if value and len(value) > 3:
            s = s.replace(value, "<path>")
    if len(s) > limit:
        s = s[:limit].rstrip() + "..."
    return s


def _see_ops(names) -> list[dict[str, str]]:
    from .surface.groups import OP_GROUP
    return [{"op": n, "tool": OP_GROUP[n]} for n in names if n in OP_GROUP]


def refusal(exc: BaseException) -> dict:
    """Build the {ok: false, error: {code, message, hint, ...}} payload."""
    declared = _declared_code(exc)
    code = declared or classify(exc)
    if code == "INTERNAL_ERROR" and declared is None:
        message = (f"something inside the server failed while running this call "
                   f"({type(exc).__name__}: {scrub(str(exc) or type(exc).__name__)}). "
                   "The arguments are not the cause.")
    else:
        raw = getattr(exc, "message", None) if declared else None
        message = _safe_message(raw or str(exc))
        if isinstance(exc, KeyError) and code == "NOT_FOUND" and exc.args:
            message = _safe_message(str(exc.args[0]))
        elif isinstance(exc, LookupError) and len(message) < 40:
            part = message.strip("'\"")
            if part.lower().endswith((".xml", ".rels", ".bin")):
                # A missing OOXML part: say what the package is missing.
                code = "UNSUPPORTED_CONTENT"
                message = (
                    f"the workbook package is missing {part}, which every reader "
                    "needs; the file is not a usable .xlsx/.xlsm")
            else:
                message = (
                    f"internal lookup failed on {message}: a nested parameter "
                    "probably has the wrong shape (list where a dict belongs, or "
                    "vice versa)"
                )
    hint = getattr(exc, "hint", None)
    if (not isinstance(hint, str) or not hint) and isinstance(
            exc, (_err.WorkbookNotFound, FileNotFoundError)):
        # Re-reading a workbook that is not there
        # cannot help; say how to make one.
        from .surface.copy import MISSING_FILE_HINT
        hint = MISSING_FILE_HINT
    if (not isinstance(hint, str) or not hint) and isinstance(exc, _err.CeilingExceeded):
        # The range was valid, only too big for one call.
        from .surface.copy import CEILING_HINT
        hint = CEILING_HINT
    if not isinstance(hint, str) or not hint:
        hint = HINTS.get(code, "")
        if code == "BAD_PARAMS" and declared is None:
            # The Word editor's rule: an engine's own refusal of a
            # value says exactly what is wrong; the generic "did not fit this
            # op" line beside it adds nothing. The engine's hint is the
            # message itself, so the generic one is left out. A schema
            # mismatch (declared by the dispatcher) keeps the xl_help hint.
            hint = ENGINE_REFUSAL_HINT
    if isinstance(exc, (_err.WorkbookCorrupt, zipfile.BadZipFile)):
        hint = f"{hint} {RECOVERY_HINT}".strip()
    error: dict[str, Any] = {"code": code, "message": message, "hint": hint}
    matches = getattr(exc, "matches", None)
    if matches:
        error["matches"] = matches
    detail = getattr(exc, "detail", None)
    if detail:
        error["detail"] = detail
    details = getattr(exc, "details", None)
    details = dict(details) if isinstance(details, dict) else {}
    see = _see_ops(getattr(exc, "hint_tools", None) or [])
    if see:
        details["see_ops"] = see
    if details:
        error["details"] = details
    return {"ok": False, "error": error}


def dumps(payload: Any) -> str:
    """The one serializer. Compact, UTF-8 as-is, surrogates replaced."""
    return _json.dumps(_sanitize(payload), ensure_ascii=False, separators=(",", ":"),
                       default=str)


def _sanitize(value: Any) -> Any:
    if isinstance(value, str):
        try:
            value.encode("utf-8")
            return value
        except UnicodeEncodeError:
            return value.encode("utf-8", "replace").decode("utf-8")
    if isinstance(value, dict):
        return {_sanitize(k): _sanitize(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_sanitize(v) for v in value]
    return value


def tool_result(payload: dict[str, Any], *, is_error: bool = False):
    """Put a payload on the wire ONCE: one compact text block, no structured
    copy, isError set for refusals."""
    from fastmcp.tools.tool import ToolResult
    from mcp.types import TextContent

    return ToolResult(content=[TextContent(type="text", text=dumps(payload))],
                      is_error=is_error)


def unknown_tool_middleware(plan):
    """Envelope a call to a tool this server does not list (no refusal class
    may escape the closed-code shape). A family tool that
    the launch settings removed refuses NOT_AVAILABLE naming who decided
    it; any other name refuses NOT_FOUND with the tools that do exist.
    Nothing runs either way. This is the only refusal decided outside the
    dispatcher, because fastmcp answers an unlisted name before any tool."""
    from fastmcp.exceptions import NotFoundError
    from fastmcp.server.middleware import Middleware

    from .core.errors import SurfaceRefusal
    from .surface import copy
    from .surface.groups import TOOL_NAMES

    class UnknownToolEnvelope(Middleware):
        async def on_call_tool(self, context, call_next):
            try:
                return await call_next(context)
            except NotFoundError:
                name = str(getattr(context.message, "name", ""))
                if name in TOOL_NAMES:
                    exc = SurfaceRefusal(
                        "NOT_AVAILABLE",
                        f"{scrub(name, 64)} is turned off by this server's launch settings",
                        details={"tool": name, "decided_by": copy.SETTINGS_OWNER})
                else:
                    exc = SurfaceRefusal(
                        "NOT_FOUND", f"no tool named {scrub(name, 64)!r} on this server.",
                        details={"tools": list(plan.tools)},
                        hint=copy.UNKNOWN_TOOL_HINT)
                return tool_result(refusal(exc), is_error=True)

    return UnknownToolEnvelope()
