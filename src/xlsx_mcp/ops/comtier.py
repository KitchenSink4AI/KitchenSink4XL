"""ops/comtier.py: the COM tier, driving real Excel for what files cannot do.

Every operation here runs through com.session.ComExecutor: process-wide
serialization (one COM worker thread), DisplayAlerts/EnableEvents suppressed
and restored around every operation, bounded timeouts with clean structured
refusals, save/export retry with backoff, and PID-journaled private workers
only (NEVER the user's Excel session; a file open in the user's Excel refuses
with WORKBOOK_LOCKED semantics).

Design rules:
- CALC: every recalc issues an explicit Application.CalculateFull; open-time
  calculation does NOT fire under manual calculation mode (measured).
- PIVOTS: REAL pivot tables only, created and refreshed through Excel. Never
  a fake static table.
- VBA: preserve/inspect only. No macro authoring, and macro EXECUTION
  (com_run_macro) is not offered.

Mutating operations hold the workbook's write lock from before Excel opens
it until its saved copy is in place, stage the safesave backup slots
before Excel touches the file, and check Excel's copy (structural verify,
integrity, encryption) BEFORE it replaces the workbook, so a copy that fails
never replaces it; after the promote the file is confirmed to be that
copy, byte for byte. Excel's own save replaces the whole package (part names
and order change legitimately), so the file-tier part-inventory diff does not
apply here; the structural check plus Excel-as-writer is the honesty basis,
and com_validate_opens_clean is available as the deep check.
"""

from __future__ import annotations

import os
import re
import threading
import time
from pathlib import Path
from typing import Any, Callable

from ..com import session as _session
from ..core import calc as _calc
from ..core import deadline as _deadline
from ..core import hazard as _hazard
from ..core import safesave as _safesave
from ..core import verify as _verify
from ..core import xmlsafe as _xmlsafe
from ..core.errors import (
    ExcelDisconnected,
    TargetNotFound,
    ValidationFailed,
    XlMcpError,
)
from ..core.sandbox import check_path

# ------------------------------------------------------------ Excel constants

XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_TYPE_PDF = 0
XL_SCREEN = 1     # CopyPicture Appearance
XL_BITMAP = 2     # CopyPicture Format

AGG_FUNCS: dict[str, int] = {
    "sum": -4157, "count": -4112, "average": -4106, "max": -4136,
    "min": -4139, "product": -4149, "count_numbers": -4113,
    "stdev": -4155, "var": -4164,
}

FILE_FORMATS: dict[str, int] = {
    "xlsx": 51, "xlsm": 52, "xlsb": 50, "xls": 56,
    "csv": 62,      # xlCSVUTF8 (Excel 2016+; UTF-8 with BOM)
    "ods": 60,
}

SPARK_TYPES: dict[str, int] = {"line": 1, "column": 2, "win_loss": 3}

_PDF_SCOPES = ("workbook", "sheet", "range")


# ---------------------------------------------------------------- helpers


def _norm_path(path: str, label: str) -> str:
    p = check_path(path, label)
    # Waits out another save's swap; names a workbook a stopped save left
    # aside.
    from ..core import swapwait as _swapwait
    _swapwait.require_workbook(p, f"no such workbook: {p}")
    return p


def _out_path(path: str, label: str, *, overwrite: bool,
              source: str | None = None,
              workbook_target: bool = False) -> tuple[str, dict]:
    """Output-target guard for the COM export/render/convert family: the
    same non-mutating-writer covenant as the file-tier exports
    (core.outguard): never the source workbook, never inside the backup
    store, an existing target refuses without overwrite:true, and
    overwrite first keeps a timestamped .bak of what it replaces."""
    from ..core.outguard import guard_out_file
    p, info = guard_out_file(path, source=source, overwrite=overwrite,
                             what=label, workbook_target=workbook_target)
    parent = os.path.dirname(os.path.abspath(p))
    if parent and not os.path.isdir(parent):
        raise TargetNotFound(f"output directory does not exist: {parent}")
    return p, info


def _ws(wb, sheet: str | None):
    """Worksheet by name (or the first sheet), with an honest not-found."""
    if sheet is None:
        return wb.Worksheets(1)
    try:
        return wb.Worksheets(sheet)
    except Exception:
        names = [str(wb.Worksheets(i + 1).Name)
                 for i in range(int(wb.Worksheets.Count))]
        raise TargetNotFound(
            f"no sheet named {sheet!r}; sheets: {names}") from None


def _op_of(label: str) -> str:
    """The operation name out of an executor label like
    ``recalculate(book.xlsx)``, for use in a refusal message."""
    return str(label).split("(", 1)[0].strip() or "the operation"


def _wrap_com_error(exc: Exception, doing: str) -> XlMcpError:
    """Turn any COM failure into one of this server's errors, with Excel's
    own words in it and never a pywin32 tuple repr.

    Two outcomes. A COM failure that means EXCEL DIED (the
    taskkilled-mid-write case, HRESULT 0x800706BE) is not a bad parameter
    and must not be coded as one: an unattended caller reading BAD_PARAMS
    will not retry, it will go
    "fix" a call that was correct. Those raise ExcelDisconnected, which the
    envelope maps to CONFLICT. Everything else stays the plain refusal."""
    msg = _session.excel_error_message(exc)
    if _session.excel_process_died(exc) or _excel_ended_under(exc):
        return ExcelDisconnected(
            f"the Excel worker process stopped answering while {doing}: "
            f"{msg}. The operation did not complete. Retry: the COM layer "
            "re-arms with a fresh Excel on the next call.")
    return XlMcpError(f"Excel refused while {doing}: {msg}")


def _excel_ended_under(exc: Exception) -> bool:
    """Whether the Excel the COM job in flight uses has ended, whatever the
    exception says (no job in flight on this thread: no). A failed member lookup (AttributeError) on an Excel
    that is going gets a second to finish going; Excel's own refusals come
    from a live Excel and are answered at once."""
    from ..com import instances as _instances
    pid = _instances.current_pid()
    if not pid:
        return False
    return _session.excel_gone(pid, wait=1.0 if isinstance(exc, AttributeError) else 0.0)


def _guarded(body: Callable[[Any, Any], dict], doing: str
             ) -> Callable[[Any, Any], dict]:
    """Run an operation body with the same COM-error translation the open and
    save steps get. Before this, only open_workbook and wb.Save were wrapped,
    so a com_error raised by any COM call inside a body (CalculateFull, a
    Range assignment, a Worksheets lookup) propagated untranslated and
    envelope.refusal() rendered it with str() as the raw pywin32 tuple
    ``(-2147023170, 'The remote procedure call failed.', None, None)``."""
    def run(app, wb) -> dict:
        try:
            return body(app, wb)
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, doing) from exc
    return run


def _refuse_encrypted_without_password(path: str,
                                       password: str | None) -> None:
    """An encrypted (CFB) file opened with NO password HANGS Excel on a modal
    prompt (probe-proven: no supplied argument suppresses it), so it is
    refused by signature BEFORE any COM call.

    A WRONG password is a different case and this message used to describe it
    wrongly. It said a wrong password surfaces as the operation timeout,
    which taught callers to budget 60 seconds for a mistyped password. The
    positional-password delivery in session.open_workbook fixed that: a wrong
    password now comes back in about a tenth of a second as Excel's own
    "The password you supplied is not correct"."""
    if password is None and _session.is_encrypted_package(path):
        from ..core.errors import WorkbookProtected
        raise WorkbookProtected(
            f"{Path(path).name} is password-protected (encrypted package); "
            "this operation needs the password. "
            "com_validate_opens_clean and com_save_with_password accept "
            "one; a WRONG password comes back at once with Excel's own "
            "wrong-password message, so there is nothing to wait out.")


def _refuse_legacy_biff(path: str) -> None:
    """Every route that saves a workbook back to its own path refuses a
    legacy .xls before Excel starts.

    Excel would happily edit and re-save the BIFF file, but the post-save
    structural verify requires an OOXML package, so the mutation would be
    applied and then rolled back. Refuse up front with the actual remedy
    instead. It is also the password-to-modify guard for this format:
    that check reads OOXML packages
    only, and com_save_with_password, which has its own open, used to send
    a .xls with a password to modify straight to a read-write open, where
    Excel showed its Password dialog. Read-only COM ops (com_convert_format
    among them) take .xls fine."""
    if _hazard.ole_container_kind(path) == _hazard.OLE_KIND_BIFF:
        from ..core.errors import UnsupportedStructure
        exc = UnsupportedStructure(
            f"{Path(path).name} is a legacy Excel 97-2003 (BIFF) workbook; "
            "COM operations that save a workbook on this server write and "
            "verify OOXML packages only, so Excel was not started. Convert "
            "it first: com_convert_format(path=..., output='...xlsx'), then "
            "run the operation on the converted file.")
        exc.hint_tools = ("com_convert_format",)
        raise exc


def _refuse_filter_table_damage(path: str) -> None:
    """A workbook carrying the sheet filter over a table that
    earlier versions wrote is one Excel will not open, so a writing route
    would only start Excel to be refused. It is refused here, before Excel
    starts, with the repair: any file-tier edit, clear_filter among them,
    repairs it (core.tablefilter)."""
    from ..core import tablefilter as _tablefilter
    damage = _tablefilter.scan(path)
    if not damage.filters:
        return
    from ..core.errors import UnsupportedStructure
    tables = sorted({t.name for f in damage.filters for t in f.tables})
    exc = UnsupportedStructure(
        f"Excel will not open {Path(path).name}: sheet "
        f"{damage.filters[0].title!r} has a sheet filter over table "
        f"{tables[0]}, which an earlier version of this server wrote. Excel "
        "was not started and nothing was changed. Repair it first with "
        f"xl_structure clear_filter (location {{\"table\": \"{tables[0]}\"}}), "
        "then retry.")
    exc.hint_tools = ("clear_filter",)
    raise exc


def _package_checks(path: str, password: str | None, *,
                    allow_network: bool = False, op: str = "open it") -> None:
    """What a read-only COM route refuses from the file alone, before Excel
    starts and before anything about this session (its clipboard) is asked:
    a package with a DTD, a decompression bomb or a part named twice, an
    encrypted file with no password, and, unless allow_network, a workbook
    whose package or text shows it can reach the network (a file this
    server cannot read is checked by Excel itself, at the shared open)."""
    _xmlsafe.check_file(path)
    _refuse_encrypted_without_password(path, password)
    _refuse_network_before_excel(path, allow_network, op)


def _refuse_network_before_excel(path: str, allow_network: bool, op: str) -> None:
    """The shared open's network check (com.session.network_gate), made
    early where the file alone can answer, so the refusal comes before
    Excel starts."""
    if allow_network:
        return
    from ..core import netreach as _netreach
    _how, reach = _netreach.file_reach(path)
    if reach:
        raise _netreach.network_refusal(path, reach, op)


def _run_readonly(label: str, path: str,
                  body: Callable[[Any, Any], dict],
                  timeout: float | None = None,
                  password: str | None = None,
                  allow_network: bool = False) -> dict:
    """Open the workbook read-only in the pooled worker, run body(app, wb),
    close without saving. Serialized, alert-suppressed, timeout-bounded."""
    _package_checks(path, password, allow_network=allow_network, op=_op_of(label))
    body = _guarded(body, f"running {_op_of(label)} in the workbook")

    def job(manager) -> dict:
        w = manager.acquire()
        wb = None
        prior = _session._assert_hygiene(w.app)
        try:
            try:
                wb = _session.open_workbook(w.app, path, read_only=True,
                                            password=password,
                                            allow_network=allow_network,
                                            network_op=_op_of(label))
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(exc, "opening the workbook") from exc
            out = body(w.app, wb)
            out.setdefault("instance_pid", w.pid)
            return out
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            _session._restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    return _session.run_com(label, job, timeout=timeout)


def _new_save_state(copy: str | None = None) -> dict:
    """What a COM save's job and its calling thread share: whether Excel's
    save began, where Excel writes its copy (chosen by the calling thread,
    so it can remove the copy of a save that was cut off), Excel's copy
    once written, and whether the calling thread has already answered
    without it (_keep_copy)."""
    return {"save_started": False, "copy": copy, "staged": None,
            "settled": False, "lock": threading.Lock()}


def _keep_copy(state: dict, staged: str) -> None:
    """Hand Excel's copy to the calling thread, which puts it in place.

    When that thread has already answered (the call timed out while Excel
    was saving), nobody will: the copy is removed here and the job stops,
    so an Excel that answers late leaves nothing in the temp folder and
    never replaces the workbook. The error goes nowhere; the call's
    answer was given."""
    with state["lock"]:
        if not state["settled"]:
            state["staged"] = staged
            return
    _session.discard_saved(staged)
    raise XlMcpError("the call had already answered when Excel finished its "
                     "save; the copy was dropped")


def _refuse_if_changed(path: str, stamp) -> None:
    """The workbook Excel has just opened must be the file the save stamped
    when it began: a change that landed after the stamp was taken,
    from a program that does not use the write lock, is not in the backup
    staged for this save. Refused before Excel edits anything."""
    if _safesave.file_stamp(path) != stamp:
        raise _safesave.changed_on_disk_error(path)


def _run_mutation(label: str, path: str,
                  body: Callable[[Any, Any], dict],
                  *, backup: bool = True, timeout: float | None = None,
                  save: bool = True, password: str | None = None,
                  allow_network: bool = False) -> dict:
    """The mutating pattern: guard (not open elsewhere), stage the backup,
    open in the pooled worker, run body(app, wb), save a copy with retry,
    close, check the copy (before it replaces the workbook, so a bad save
    never does, backup or not), put it in place, confirm the file is that
    copy.

    From before Excel opens the workbook until the checks after the
    copy is in place, the CALLING thread holds the workbook's write lock
    (core.safesave.write_lock), the one every file-tier save takes. An edit
    by this or another server that used to land between Excel's open and
    the promote, and was then overwritten by Excel's copy while both calls
    answered ok, now waits for the save: it is then refused in the file
    tier's own words (another process is modifying the workbook, or the
    workbook changed after the edit read it), or, made after it, applies on
    top. The COM job runs on the COM worker thread and never takes the
    lock: that thread refuses to (safesave.forbid_write_lock_on_this_
    thread), because a hold there would wait for its own caller. Only this
    thread puts the copy in place, so a COM timeout releases the lock and a
    hung Excel can never promote late. A program that does not use the
    lock is caught by the stamp: the workbook's size and modification time
    when the save began must still be the file's once Excel has opened it
    and again just before the promote."""
    _safesave.refuse_backup_slot_target(path, "edit")  # never a backup in place
    _xmlsafe.check_file(path)  # a part with a DTD: refused before Excel starts
    _refuse_encrypted_without_password(path, password)
    # A password to modify makes Excel ask in a visible dialog on a
    # read-write open; refused before Excel starts.
    _hazard.refuse_if_write_reserved(path)
    _hazard.refuse_xlm_macro_sheet(path)  # ForceDisable cannot disable XLM
    _refuse_legacy_biff(path)
    _refuse_filter_table_damage(path)  # before Excel starts
    _refuse_network_before_excel(path, allow_network, _op_of(label))
    # Where Excel will write its copy, beside the workbook;
    # refused here, before the lock or anything else changes the folder,
    # when Excel could not save there.
    copy = _session.copy_path_for(path) if save else None
    body = _guarded(body, f"running {_op_of(label)} in the workbook")
    with _safesave.write_lock(path):
        return _mutate_locked(label, path, body, backup=backup,
                              timeout=timeout, save=save, password=password,
                              copy=copy, allow_network=allow_network)


def _mutate_locked(label: str, path: str, body: Callable[[Any, Any], dict],
                   *, backup: bool, timeout: float | None, save: bool,
                   password: str | None, copy: str | None = None,
                   allow_network: bool = False) -> dict:
    """_run_mutation's work, under the workbook's write lock."""
    # Taken first, so the backup, the scans and Excel all see this file.
    stamp = _safesave.file_stamp(path)
    warnings = _session.guard_target_closed(path)
    pre_report = None
    try:
        pre_report = _hazard.scan_path(path)
    except Exception:
        pre_report = None
    # Two-phase rotation (the same covenant as the file tier):
    # stage the pre-mutation content now, commit onto the slots only after
    # Excel's save came back. A COM op that fails before or during save
    # leaves the slots exactly as they were instead of burning prev.
    backup_slot = None
    ticket = None
    if backup:
        ticket = _safesave.prepare_rotation(path)
    state = _new_save_state(copy)
    # The broken links the file already had, so the check
    # after Excel's save fails only on ones the save introduced.
    integrity_baseline = None
    if save and password is None:
        try:
            integrity_baseline = _verify.package_problems(path)
        except Exception:  # noqa: BLE001
            integrity_baseline = None

    def job(manager) -> dict:
        w = manager.acquire()
        wb = None
        staged = None
        prior = _session._assert_hygiene(w.app)
        try:
            try:
                wb = _session.open_workbook(w.app, path, password=password,
                                            allow_network=allow_network,
                                            network_op=_op_of(label))
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(exc, "opening the workbook") from exc
            _refuse_if_changed(path, stamp)  # Excel has the stamped file
            changed = body(w.app, wb)
            if save:
                # The point of no return, checked before Excel saves.
                _deadline.commit_or_raise("the save")
                state["save_started"] = True
                try:
                    # Excel writes a copy; the file itself is
                    # replaced only after Excel has closed it (below).
                    staged = _session.save_workbook(
                        wb, label=f"save {Path(path).name}", path=path,
                        copy=copy)
                except XlMcpError:
                    raise
                except Exception as exc:  # noqa: BLE001
                    raise _wrap_com_error(exc, "saving the workbook") from exc
                _keep_copy(state, staged)
            try:
                wb.Close(False)
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(exc, "closing the workbook") from exc
            wb = None
            return {"changed": changed, "instance_pid": w.pid, "staged": staged}
        except BaseException:
            _session.discard_saved(staged)
            raise
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            _session._restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    try:
        out = _session.run_com(label, job, timeout=timeout)
    except BaseException as exc:
        # Excel only ever wrote a copy, so the workbook is as it was.
        _settle_failed_save(path, ticket, state, exc)
        raise
    digest = None
    if out.get("staged"):
        # The copy is checked where Excel wrote it, before anything
        # replaces the workbook (an encrypted workbook's copy must still be
        # encrypted; the structural zip check cannot read that container).
        check = _session.check_saved_copy(
            path, encrypted=password is not None,
            integrity_baseline=integrity_baseline)
        try:
            digest = _session.put_in_place(out["staged"], path, stamp=stamp,
                                           check=check, warnings=warnings,
                                           ticket=ticket)
        except BaseException:
            if ticket is not None:
                ticket.abort()  # the workbook is as it was; so are the slots
            raise
    elif ticket is not None:
        ticket.abort()  # nothing was saved (save=false): the slots stay
    if ticket is not None and ticket.rotated["prev"]:
        backup_slot = "prev"

    verified = False
    if save:
        _confirm_in_place(path, digest, backup_slot)
        # An encrypted workbook's content cannot be read here; Excel as the
        # writer plus a COM reopen (com_validate_opens_clean with the
        # password) is its check.
        verified = password is None

    result: dict[str, Any] = {
        "ok": True,
        "file": path,
        "changed": out.get("changed", {}),
        "saved": bool(save),
        "backup": backup_slot,
        "verified": verified,
        "engine": "com",
        "instance_pid": out.get("instance_pid"),
        "warnings": warnings,
    }
    _carry_network_note(result, out)
    if pre_report is not None and pre_report.hazards:
        result["preserved"] = (
            "Excel saved the workbook itself, so fragile parts "
            f"({', '.join(pre_report.labels())}) are preserved at full "
            "fidelity")
    return result


def _confirm_in_place(path: str, digest: str | None,
                      backup_slot: str | None) -> None:
    """After the promote, the workbook must be, byte for byte, the copy that
    passed the checks.

    The checks ran on the copy before it went in place
    (com.session.check_saved_copy). They re-read every part (the structural
    verify took 12.5 s on a 5.8 MB workbook, measured), so they are not run
    again on the workbook; this compares hashes, which takes milliseconds.
    A mismatch means another program wrote the file at that moment, and
    nothing is put over what is there now."""
    if digest is None or _session.file_sha256(path) == digest:
        return
    raise ValidationFailed(
        f"{Path(path).name} does not match the copy that was checked when it "
        "was put in place (another program may have written to it at that "
        "moment), so this save is not confirmed. "
        + ("The workbook as it was before this call is in the backup: "
           "xl_workbook manage_backups (action='restore')." if backup_slot else
           "No backup was taken (backup=false)."))


def _note(exc: BaseException, note: str) -> BaseException:
    """Append a sentence to an error's message (its args, and the declared
    message a SurfaceRefusal carries), keeping its type and code."""
    text = f"{str(exc).rstrip('. ')}. {note}"
    exc.args = (text,)
    if isinstance(getattr(exc, "message", None), str):
        exc.message = text
    return exc


def _settle_failed_save(path: str, ticket, state: dict,
                        exc: BaseException) -> None:
    """Excel's part of the call failed or was cut off. Excel writes the edit
    to a copy and never to the workbook's own file (session.save_workbook),
    so the file is exactly as it was: the copy is dropped and the slots
    stay as they were (the two-phase rotation's covenant). When Excel saved
    over the file itself, a save cut off after the commit point left the
    file in an unknown state, and this put the backup back over it.

    The call answers now; a job still inside Excel (a timeout) drops its
    own copy when it comes back (_keep_copy)."""
    with state["lock"]:
        state["settled"] = True
        staged, state["staged"] = state.get("staged"), None
    _session.discard_saved(staged)
    # A copy Excel was still writing when the call was cut off (Excel
    # is ended by then) is removed too, by its name beside the workbook.
    _session.discard_saved(state.get("copy"))
    if ticket is not None:
        ticket.abort()
    if state.get("save_started"):
        _note(exc, f"{Path(path).name} was not changed: Excel writes the edit to "
                   "a copy first, and the copy was not put in place")


# ============================================================== recalculate


#: The refusal for engine='formulas', removed in 2.0.0 (CALC_UNAVAILABLE).
FORMULAS_ENGINE_REMOVED = (
    "The formulas engine was removed in 2.0.0. Recalculation needs Windows with "
    "Microsoft Excel: use engine 'com' or 'auto'. Reads still return each formula's "
    "cached value, labelled cached.")

#: The refusal for 'auto' or 'com' where Excel cannot be reached (CALC_UNAVAILABLE);
#: details.excel_unavailable says why.
EXCEL_NEEDED_FOR_RECALC = (
    "Recalculation needs Windows with Microsoft Excel, and Excel is not available "
    "here. Reads still return each formula's cached value, labelled cached.")


def recalculate(path: str, engine: str = "auto",
                timeout_seconds: float | None = None,
                backup: bool = True, allow_network: bool = False) -> dict:
    """Fidelity recalculation in Excel. engine='com' opens the workbook in a
    private hidden Excel, issues an explicit CalculateFull (open-time calc
    does NOT fire under manual mode; spike-proven), saves, and reports how
    many formula cells gained cached values. engine='auto' is the same route.

    engine='formulas' (the pure-Python `formulas` library) was removed in
    2.0.0. It is refused first, before the file is read, before COM is
    probed and before anything imports that library. Without Excel, 'auto'
    and 'com' refuse; nothing falls back to the library.

    A workbook whose formulas can reach the network as Excel calculates
    them (WEBSERVICE, IMAGE, RTD and the rest in core.netreach) or that has
    external data connections is refused before Excel starts, unless
    allow_network is true: no network request the caller did not ask for."""
    from ..core.errors import CalcUnavailable

    # The engine is decided before anything reads the file or probes COM:
    # the removed engine is refused here, and nothing below imports the
    # `formulas` library or hands it a path.
    engine = str(engine).strip().lower()
    if engine == "formulas":
        raise CalcUnavailable(FORMULAS_ENGINE_REMOVED)
    if engine not in ("auto", "com"):
        raise XlMcpError(f"engine must be 'auto' or 'com', got {engine!r}")
    p = _norm_path(path, "recalculate")
    _xmlsafe.check_file(p)  # a bomb, a repeated part, a DTD: refused before Excel reads it
    _refuse_encrypted_without_password(p, None)
    ok, why = _session.com_available()
    _hazard.refuse_if_write_reserved(p)  # before Excel starts
    if not ok:
        exc = CalcUnavailable(EXCEL_NEEDED_FOR_RECALC)
        exc.details = {"excel_unavailable": why}
        raise exc

    reach = _network_gate(p, allow_network, "recalculate")
    absent_before = _count_absent_cached(p)
    seeded: list[int] = [0]
    circular: list[str] = []

    def body(app, wb) -> dict:
        app.CalculateFull()
        iterating = False
        try:
            iterating = bool(app.Iteration)
            if iterating:
                seeded[0] = _reseed_iterative_errors(app, wb)
        except Exception:  # noqa: BLE001
            pass
        if not iterating:
            # Excel knows where a circular reference is
            # (Worksheet.CircularReference, the first cell of a loop on each
            # sheet); say so instead of reporting a clean recalculation.
            for ws in wb.Worksheets:
                try:
                    loop = ws.CircularReference
                    if loop is not None:
                        circular.append(f"{ws.Name}!{str(loop.Address).replace('$', '')}")
                except Exception:  # noqa: BLE001
                    continue
        return {"calculated": "CalculateFull"}

    result = _run_mutation(f"recalculate({Path(p).name})", p, body,
                           backup=backup, timeout=timeout_seconds, allow_network=allow_network)
    absent_after = _count_absent_cached(p)
    result["recalculated"] = True
    result["changed"] = {
        "calculation": "explicit CalculateFull in Excel, cache saved",
        "formula_cells_without_cached_value_before": absent_before,
        "formula_cells_without_cached_value_after": absent_after,
    }
    result["freshness"] = (
        "cached values were computed by Excel just now; reads return them, "
        "labelled cached")
    if circular:
        result["circular_references"] = circular
        result["warnings"] = list(result.get("warnings", [])) + [
            f"Excel found a circular reference (a formula that depends on its own "
            f"result) at {', '.join(circular)}; the cells in the loop hold whatever "
            "Excel last computed, often 0. Fix the formulas, or turn on iterative "
            "calculation with xl_workbook set_workbook_properties if the loop is meant"]
    _note_network(result, reach)
    if seeded[0]:
        result["changed"]["iterative_cells_reseeded"] = seeded[0]
        result["warnings"] = list(result.get("warnings", [])) + [
            f"{seeded[0]} formula cell(s) under iterative calculation had no "
            "seed value and were loading as errors; their formulas were "
            "re-entered so Excel could iterate them"]
    return result


def _network_gate(path: str, allow_network: bool, op: str) -> dict | None:
    """Before Excel starts: what in the workbook can reach the network as
    Excel calculates it (core.netreach). Refused unless allow_network; with
    it, returned so the result can say what Excel was allowed to do."""
    from ..core import netreach as _netreach
    reach = _netreach.network_reach(path)
    if reach and not allow_network:
        raise _netreach.network_refusal(path, reach, op)
    return reach


def _carry_network_note(result: dict, out: dict) -> dict:
    """Copy the shared open's network_allowed and network_note (made in
    com.session, where every COM open passes) from a job's output into the
    result the op builds from it."""
    if out.get("network_allowed"):
        result["network_allowed"] = True
        result["network_note"] = out.get("network_note") or _session.NETWORK_NOTE
    return result


def _note_network(result: dict, reach: dict | None) -> dict:
    """With allow_network and something that can reach the network, list
    what Excel was allowed to fetch from, beside the shared open's
    network_allowed and network_note."""
    if reach:
        result["network_allowed"] = True
        result["network_note"] = _session.NETWORK_NOTE
        result["network_reach"] = reach
        result["warnings"] = list(result.get("warnings", [])) + [
            "allow_network=true: Excel was allowed to fetch from the network for "
            "the formulas or data connections listed in network_reach"]
    return result


def _reseed_iterative_errors(app, wb) -> int:
    """Re-enter error-valued formulas so iterative calculation can converge.

    Under iterative calc Excel seeds every pass from the cell's CURRENT
    value. A formula that arrived from a file-tier write has NO cached value,
    so Excel loads it as an error and each iteration propagates that error:
    the numbers-safety gate caught =B1+1 stuck at #VALUE! forever in a
    workbook where Excel's own authoring of the identical formula converges.
    Re-entering the formula through Excel restores a numeric seed. Cells whose
    error is genuine (#REF!, #DIV/0!) are unharmed: re-entering the same
    formula recomputes the same error. Legacy CSE array cells are skipped,
    since assigning .Formula would break the array.

    The re-entry goes through .Formula2, NOT .Formula. HasArray is False for a
    DYNAMIC array, so the CSE skip above never covered them, and assigning
    legacy .Formula re-enters in implicit-intersection mode: an Excel probe
    took an error-valued
    =FILTER(A1:A3,A1:A3>99), assigned .Formula = .Formula, and Excel then
    reported the formula as =@FILTER(A1:A3,A1:A3>99) and stored it as a plain
    <f> with the t="array" ref= spill attributes GONE. The formula had been
    silently converted. .Formula2 round-trips it unchanged. Pre-2019 Excel has
    no .Formula2 (and no dynamic arrays either), so the fallback below is the
    old path where it is still the correct one."""
    xl_cell_type_formulas, xl_errors = -4123, 16
    total = 0
    for ws in wb.Worksheets:
        try:
            errors = ws.Cells.SpecialCells(xl_cell_type_formulas, xl_errors)
        except Exception:  # noqa: BLE001
            continue  # SpecialCells raises when the sheet has no error cells
        try:
            for cell in errors:
                try:
                    if cell.HasArray:
                        continue
                    try:
                        cell.Formula2 = cell.Formula2
                    except Exception:  # noqa: BLE001 - pre-2019 Excel
                        cell.Formula = cell.Formula
                    total += 1
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            continue
    if total:
        app.CalculateFull()
    return total


def _count_absent_cached(path: str) -> int:
    """Formula cells with NO cached value (read as blank outside Excel)."""
    keep_vba = _xmlsafe.keeps_vba(path)
    fwb = _xmlsafe.load_workbook(path, keep_vba=keep_vba, data_only=False)
    cwb = _xmlsafe.load_workbook(path, keep_vba=keep_vba, data_only=True)
    from . import gridio as _gridio
    _gridio.fill_empty_text_results(cwb)  # "" is a cached value
    try:
        n = 0
        for ws in fwb.worksheets:
            cws = cwb[ws.title]
            for (r, c), cell in getattr(ws, "_cells", {}).items():
                # cell TYPE, not a leading '=': neutralized injection text is
                # data, and counting it as an uncalculated formula made the
                # recalculate report claim work it never had to do.
                if (_calc.is_formula_cell(cell)
                        and cws.cell(r, c).value is None):
                    n += 1
        return n
    finally:
        fwb.close()
        cwb.close()


# ============================================================== real pivots


def com_manage_pivot(path: str, action: str, name: str | None = None,
                     source_sheet: str | None = None,
                     source_range: str | None = None,
                     dest_sheet: str | None = None,
                     dest_cell: str = "A3",
                     rows: list[str] | None = None,
                     columns: list[str] | None = None,
                     filters: list[str] | None = None,
                     values: list[dict] | None = None,
                     timeout_seconds: float | None = None,
                     backup: bool = True, allow_network: bool = False) -> dict:
    """Create, refresh, delete, or list REAL pivot tables through Excel.

    allow_network: a workbook whose formulas or data connections can reach
    the network is refused before Excel opens it (Excel calculates as it
    opens), unless this is true."""
    p = _norm_path(path, "com_manage_pivot")
    action = str(action).strip().lower()
    if action not in ("create", "refresh", "delete", "list"):
        raise XlMcpError(
            f"action must be create, refresh, delete, or list; got {action!r}")

    if action == "list":
        def body(app, wb) -> dict:
            return {"pivot_tables": _list_pivots(wb)}
        return _run_readonly(f"pivot_list({Path(p).name})", p, body,
                             timeout=timeout_seconds, allow_network=allow_network)

    if action == "create":
        if not source_range:
            raise XlMcpError("create needs source_range (an A1 range with a "
                             "header row, e.g. 'A1:D200')")
        if not values:
            raise XlMcpError(
                "create needs values: a list like "
                "[{'field': 'Amount', 'func': 'sum'}]")
        for spec in values:
            # query_range's aggregate says func, and the
            # help example here said function, which was ignored without a
            # word (the default, sum, hid it). Both are read; anything else
            # in an entry refuses.
            if not isinstance(spec, dict):
                raise XlMcpError("every values entry is an object like "
                                 "{'field': 'Amount', 'func': 'sum'}")
            unknown = sorted(set(spec) - {"field", "func", "function", "caption"})
            if unknown:
                raise XlMcpError(
                    f"values entries take field, func and caption; {unknown} "
                    "is not one of them")
            if "func" not in spec and "function" in spec:
                spec["func"] = spec.pop("function")
            func = str(spec.get("func", "sum")).lower()
            if func not in AGG_FUNCS:
                raise XlMcpError(
                    f"unknown aggregation {func!r}; one of "
                    f"{sorted(AGG_FUNCS)}")
            if not spec.get("field"):
                raise XlMcpError("every values entry needs a 'field'")

        def body(app, wb) -> dict:
            try:
                src_ws = _ws(wb, source_sheet)
                src = src_ws.Range(source_range)
                if dest_sheet is not None:
                    try:
                        dst_ws = wb.Worksheets(dest_sheet)
                    except Exception:
                        dst_ws = wb.Worksheets.Add()
                        dst_ws.Name = dest_sheet
                else:
                    dst_ws = wb.Worksheets.Add()
                pt_name = name or f"KS4XLPivot{int(time.time()) % 100000}"
                cache = wb.PivotCaches().Create(
                    XL_DATABASE, f"'{src_ws.Name}'!{src.Address}")
                pt = cache.CreatePivotTable(
                    dst_ws.Range(dest_cell), pt_name)
                for f in rows or []:
                    pt.PivotFields(f).Orientation = XL_ROW_FIELD
                for f in columns or []:
                    pt.PivotFields(f).Orientation = XL_COLUMN_FIELD
                for f in filters or []:
                    pt.PivotFields(f).Orientation = XL_PAGE_FIELD
                added = []
                for spec in values:
                    func = str(spec.get("func", "sum")).lower()
                    caption = spec.get("caption") or \
                        f"{func.capitalize()} of {spec['field']}"
                    pt.AddDataField(pt.PivotFields(spec["field"]),
                                    caption, AGG_FUNCS[func])
                    added.append(caption)
                return {"pivot_created": pt_name,
                        "sheet": str(dst_ws.Name),
                        **_pivot_location(pt),
                        "rows": rows or [], "columns": columns or [],
                        "filters": filters or [], "values": added}
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(
                    exc, "creating the pivot table (check that source_range "
                    "includes a header row and every field name matches a "
                    "header exactly)") from exc

        return _run_mutation(f"pivot_create({Path(p).name})", p, body,
                             backup=backup, timeout=timeout_seconds, allow_network=allow_network)

    if action == "refresh":
        def body(app, wb) -> dict:
            found = []
            for i in range(1, int(wb.Worksheets.Count) + 1):
                ws = wb.Worksheets(i)
                for j in range(1, int(ws.PivotTables().Count) + 1):
                    pt = ws.PivotTables(j)
                    if name is None or str(pt.Name) == name:
                        try:
                            pt.RefreshTable()
                        except Exception as exc:  # noqa: BLE001
                            raise _wrap_com_error(
                                exc, f"refreshing pivot {pt.Name!r}") from exc
                        found.append(str(pt.Name))
            if name is not None and not found:
                raise TargetNotFound(
                    f"no pivot table named {name!r}; use action='list'")
            if not found:
                raise TargetNotFound("this workbook has no pivot tables")
            return {"refreshed": found}
        return _run_mutation(f"pivot_refresh({Path(p).name})", p, body,
                             backup=backup, timeout=timeout_seconds, allow_network=allow_network)

    # delete
    if not name:
        raise XlMcpError("delete needs the pivot table name; "
                         "use action='list' to see them")

    def body(app, wb) -> dict:
        for i in range(1, int(wb.Worksheets.Count) + 1):
            ws = wb.Worksheets(i)
            for j in range(1, int(ws.PivotTables().Count) + 1):
                pt = ws.PivotTables(j)
                if str(pt.Name) == name:
                    addr = str(pt.TableRange2.Address).replace("$", "")
                    sheet_name = str(ws.Name)
                    pt.TableRange2.Clear()
                    return {"pivot_deleted": name, "sheet": sheet_name,
                            "cleared_range": addr}
        raise TargetNotFound(
            f"no pivot table named {name!r}; use action='list'")

    return _run_mutation(f"pivot_delete({Path(p).name})", p, body,
                         backup=backup, timeout=timeout_seconds, allow_network=allow_network)


def _pivot_location(pt) -> dict:
    """The pivot's own range as the file stores it (and
    get_pivot reads it): TableRange1, plain A1. When report filters sit above
    it, location_with_filters is the whole area (TableRange2)."""
    body = str(pt.TableRange1.Address).replace("$", "")
    whole = str(pt.TableRange2.Address).replace("$", "")
    out = {"location": body}
    if whole != body:
        out["location_with_filters"] = whole
    return out


def _list_pivots(wb) -> list[dict]:
    out: list[dict] = []
    for i in range(1, int(wb.Worksheets.Count) + 1):
        ws = wb.Worksheets(i)
        for j in range(1, int(ws.PivotTables().Count) + 1):
            pt = ws.PivotTables(j)
            entry = {"name": str(pt.Name), "sheet": str(ws.Name),
                     **_pivot_location(pt)}
            try:
                entry["source"] = str(pt.PivotCache().SourceData)
            except Exception:
                pass
            try:
                entry["refreshed"] = str(pt.RefreshDate)
            except Exception:
                pass
            out.append(entry)
    return out


# ============================================================ export / render


def com_export_pdf(path: str, output: str, scope: str = "workbook",
                   sheet: str | None = None, range_a1: str | None = None,
                   overwrite: bool = False,
                   timeout_seconds: float | None = None,
                   allow_network: bool = False) -> dict:
    """Export the workbook, one sheet, or a range to PDF via Excel.

    allow_network: a workbook whose formulas or data connections can reach
    the network is refused before Excel opens it (Excel calculates as it
    opens), unless this is true."""
    p = _norm_path(path, "com_export_pdf")
    # The output path is settled HERE, whole, before
    # anything is handed to Excel. ExportAsFixedFormat appends .pdf to a
    # name that lacks it, so output='report' used to pass the guard as
    # 'report' while Excel wrote report.pdf over the user's own file with no
    # .bak. The .pdf is added first and the guard judges the exact file
    # Excel will write; Excel is handed the absolute form of that same path
    # (body below), which leaves it no default to apply.
    resolved = str(output)
    if not resolved.lower().endswith(".pdf"):
        from ..core.outguard import refuse_given_name
        refuse_given_name(resolved, source=p, what="com_export_pdf output")
        resolved += ".pdf"
    out, out_info = _out_path(resolved, "com_export_pdf output",
                              overwrite=overwrite, source=p)
    scope = str(scope).strip().lower()
    if scope not in _PDF_SCOPES:
        raise XlMcpError(f"scope must be one of {_PDF_SCOPES}, got {scope!r}")
    if scope == "range" and not range_a1:
        raise XlMcpError("scope='range' needs range_a1 (e.g. 'A1:F40')")

    def body(app, wb) -> dict:
        try:
            if scope == "workbook":
                target = wb
            elif scope == "sheet":
                target = _ws(wb, sheet)
            else:
                target = _ws(wb, sheet).Range(range_a1)
            # The point of no return, checked before Excel writes.
            _deadline.commit_or_raise("the PDF export")
            _session.com_retry(
                lambda: target.ExportAsFixedFormat(XL_TYPE_PDF,
                                                   os.path.abspath(out)),
                label="ExportAsFixedFormat")
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, "exporting to PDF") from exc
        size = os.path.getsize(out) if os.path.exists(out) else 0
        return {"exported": out, "scope": scope, "bytes": size}

    result = _run_readonly(f"export_pdf({Path(p).name})", p, body,
                           timeout=timeout_seconds, allow_network=allow_network)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ValidationFailed(
            "Excel reported success but the PDF is missing or empty")
    result.update(out_info)
    result["ok"] = True
    result["file"] = p
    return result


def _clipboard_available() -> bool:
    """True when this Windows session lets any process open the clipboard.

    Excel offers exactly one route from a range to a bitmap,
    Range.CopyPicture, and CopyPicture goes through the Windows clipboard.
    A service account, a locked workstation, or a session with no window
    station refuses the clipboard to every process, so the render cannot
    succeed regardless of Excel or this server. The probe runs entirely
    outside Excel, so its answer is about the environment. The COM-tier
    gate has probed this since b7150e3; the product did not, and handed
    back Excel's raw "CopyPicture method of Range class failed" instead.
    """
    try:
        import win32clipboard
    except Exception:  # noqa: BLE001
        return False
    try:
        win32clipboard.OpenClipboard()
    except Exception:  # noqa: BLE001
        return False
    try:
        win32clipboard.CloseClipboard()
    except Exception:  # noqa: BLE001
        pass
    return True


#: A PNG palette this small holds nothing but a background, a border and
#: their antialiasing. Measured on this machine: Excel exports an empty
#: chart canvas with a six-entry palette at every canvas size, while the
#: thinnest real render seen (four columns of EMPTY cells, gridlines only)
#: came back with thirty entries and a single cell holding the number 1
#: came back with eighty. Eight is the ceiling because it sits above the
#: artefact and an order of magnitude below the thinnest real picture.
_BLANK_PALETTE_CEILING = 8

_PNG_MAGIC = b"\x89PNG\r\n\x1a\n"


def _png_palette_size(path: str) -> int | None:
    """How many distinct colours an exported PNG can possibly hold.

    A palette PNG cannot show more colours than its PLTE chunk declares, so
    the chunk length is an exact upper bound on the picture's content and
    reading it costs a few header bytes rather than a decode. Excel exports
    the blank artefact as a palette PNG and anything with real content in it
    as either a larger palette or truecolour.

    Returns the entry count for a palette PNG, and None for any other colour
    type or for a file that is not a readable PNG. None means "no cheap
    bound", never "blank": a truecolour export is rich by construction.
    """
    try:
        with open(path, "rb") as fh:
            if fh.read(8) != _PNG_MAGIC:
                return None
            colour_type = None
            while True:
                head = fh.read(8)
                if len(head) < 8:
                    return None
                length = int.from_bytes(head[:4], "big")
                kind = head[4:8]
                if kind == b"IHDR":
                    data = fh.read(length)
                    if len(data) < 10:
                        return None
                    colour_type = data[9]
                    if colour_type != 3:
                        return None
                    fh.read(4)
                    continue
                if kind == b"PLTE":
                    return length // 3
                if kind == b"IDAT" or kind == b"IEND":
                    return None
                fh.seek(length + 4, os.SEEK_CUR)
    except OSError:
        return None


def _render_is_blank(path: str) -> bool:
    """True when the exported PNG holds no more than a background and a
    border. See _BLANK_PALETTE_CEILING for where the number comes from."""
    bound = _png_palette_size(path)
    return bound is not None and bound <= _BLANK_PALETTE_CEILING


_BLANK_RENDER_REFUSAL = (
    "Excel exported the picture but it came back blank: the canvas holds "
    "a background and a border and nothing else. The range was copied and "
    "the export ran, so this is not a bad range or a missing file. Use "
    "com_export_pdf on the same workbook (scope='sheet' or scope='range'), "
    "which draws the cells through a different Excel route and works here.")


#: The hint on the no-clipboard refusal. The session is
#: the cause; the arguments and the workbook are fine.
CLIPBOARD_HINT = (
    "Nothing is wrong with the arguments or the workbook; this session has no "
    "clipboard. xl_app com_export_pdf draws the same cells without one.")


def _frame_with_objects(ws):
    """The default frame is the used range grown to take
    in every chart, picture and shape on the sheet (one rectangle, from the
    top-left cell of the first to the bottom-right cell of the last), so a
    chart beside the data is in the picture rather than cut off."""
    used = ws.UsedRange
    r0, c0 = int(used.Row), int(used.Column)
    r1, c1 = r0 + int(used.Rows.Count) - 1, c0 + int(used.Columns.Count) - 1
    try:
        for i in range(1, int(ws.Shapes.Count) + 1):
            shape = ws.Shapes(i)
            tl, br = shape.TopLeftCell, shape.BottomRightCell
            r0, c0 = min(r0, int(tl.Row)), min(c0, int(tl.Column))
            r1, c1 = max(r1, int(br.Row)), max(c1, int(br.Column))
    except Exception:  # noqa: BLE001 (a shape with no cell anchor: keep the used range)
        pass
    return ws.Range(ws.Cells(r0, c0), ws.Cells(r1, c1))


def com_render_sheet(path: str, output: str, sheet: str | None = None,
                     range_a1: str | None = None, overwrite: bool = False,
                     timeout_seconds: float | None = None,
                     allow_network: bool = False) -> dict:
    """Render a sheet's used range (or a given range) to a PNG image for
    visual verification, via CopyPicture into a temporary chart canvas.

    The canvas is ACTIVATED before the paste and the paste is checked, and
    the exported file is checked again on the way out. Both checks exist
    because the failure this call had was silent in every direction: the
    clipboard held the bitmap, the paste reported nothing, the export
    reported success, and the caller got a white rectangle under ok:true.
    """
    p = _norm_path(path, "com_render_sheet")
    # The CALLER's own arguments are judged before the environment is. A
    # .bmp target is wrong on every machine, and answering it with a
    # clipboard refusal on the machines that have no clipboard makes the
    # same call report two different problems depending on where it ran.
    # The suffix check is pure: _out_path can write a .bak, so it stays
    # after the probe and no file is touched for a call that refuses.
    if not str(output).lower().endswith(".png"):
        raise XlMcpError("output must be a .png path")
    # The file is judged before the session: a workbook this server refuses
    # is refused as such whether or not this session has a clipboard.
    _package_checks(p, None, allow_network=allow_network, op="render_sheet")
    if not _clipboard_available():
        # Named cause, and no workaround offered for the render itself,
        # because there is none: CopyPicture is Excel's only range-to-bitmap
        # route. An environment code with its own hint,
        # not BAD_PARAMS, whose hint blamed the arguments.
        from ..core.errors import SurfaceRefusal
        raise SurfaceRefusal(
            "APP_NOT_RUNNING",
            "rendering this range needs the Windows clipboard, and this "
            "session does not have one (service accounts and locked or "
            "disconnected sessions often do not). Run it from an ordinary "
            "logged-in desktop session, or export the range as values "
            "instead",
            hint=CLIPBOARD_HINT)
    out, out_info = _out_path(output, "com_render_sheet output",
                              overwrite=overwrite, source=p)

    def body(app, wb) -> dict:
        try:
            ws = _ws(wb, sheet)
            rng = ws.Range(range_a1) if range_a1 else _frame_with_objects(ws)
            rng.CopyPicture(XL_SCREEN, XL_BITMAP)
            width, height = float(rng.Width), float(rng.Height)
            if width <= 0 or height <= 0:
                raise XlMcpError("the target range has no visible size")
            co = ws.ChartObjects().Add(0, 0, width, height)
            try:
                # ACTIVATE BEFORE PASTE. Chart.Paste() into a chart canvas
                # that is not the active chart does nothing at all, and it
                # says nothing about it: no error, no return value, an empty
                # canvas. Chart.Export then succeeds on that empty canvas,
                # which is how this call handed back ok:true over a white
                # rectangle. The hidden instance is not the cause; the
                # clipboard holds the bitmap either way.
                co.Activate()
                co.Chart.Paste()
                pasted = int(co.Chart.Shapes.Count)
                if pasted < 1:
                    raise ValidationFailed(_BLANK_RENDER_REFUSAL)
                # The point of no return, checked before Excel writes.
                _deadline.commit_or_raise("the render")
                _session.com_retry(
                    lambda: co.Chart.Export(os.path.abspath(out), "PNG"),
                    label="Chart.Export")
            finally:
                co.Delete()
            return {"rendered": out, "range": str(rng.Address).replace("$", ""),
                    "sheet": str(ws.Name), "pasted_shapes": pasted}
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, "rendering the range to an "
                                  "image") from exc

    result = _run_readonly(f"render_sheet({Path(p).name})", p, body,
                           timeout=timeout_seconds, allow_network=allow_network)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ValidationFailed(
            "Excel reported success but the PNG is missing or empty")
    # The backstop. The paste is checked in the worker, but a picture can
    # also land and still export as nothing, and a file on disk that a
    # caller was told to trust is the worst place for that to be true. The
    # blank file is removed, because a refusal that leaves the artefact
    # behind is the same lie in a different place.
    if _render_is_blank(out):
        try:
            os.remove(out)
        except OSError:
            pass
        raise ValidationFailed(_BLANK_RENDER_REFUSAL)
    result.update(out_info)
    result["ok"] = True
    result["file"] = p
    return result


def com_convert_format(path: str, output: str, format: str | None = None,
                       overwrite: bool = False,
                       timeout_seconds: float | None = None,
                       allow_network: bool = False) -> dict:
    """Convert between workbook formats via Excel's own SaveAs.

    allow_network: a workbook whose formulas or data connections can reach
    the network is refused before Excel opens it (Excel calculates as it
    opens), unless this is true."""
    p = _norm_path(path, "com_convert_format")
    # Excel's SaveAs adds the format's extension to
    # a name that has none, so the guard below would judge a different file
    # from the one Excel writes. The extension is added here first.
    if format and not Path(str(output)).suffix:
        output = f"{output}.{str(format).strip().lower()}"
    out, out_info = _out_path(output, "com_convert_format output",
                              overwrite=overwrite, source=p,
                              workbook_target=True)
    fmt = (format or Path(out).suffix.lstrip(".")).strip().lower()
    if fmt not in FILE_FORMATS:
        raise XlMcpError(
            f"format must be one of {sorted(FILE_FORMATS)}, got {fmt!r}")
    # canonical_key, not abspath: this is the ONE caller where outguard's
    # rule 1 is load-bearing (workbook_target=True skips the workbook-
    # extension rule that shadows it everywhere else), so its backstop has
    # to be at least as strong as the rule it backs up. abspath normalizes
    # neither case nor symlinks, so "BOOK.XLSX" or a junctioned spelling of
    # the source used to walk straight past it and Excel's own SaveAs
    # overwrote the workbook being converted.
    from ..core.safesave import canonical_key as _canonical_key
    if _canonical_key(out) == _canonical_key(p):
        raise XlMcpError("output must differ from the source path; "
                         "conversion never overwrites its own source")

    def body(app, wb) -> dict:
        try:
            # The point of no return, checked before the replaced
            # output is removed and Excel writes the new one.
            _deadline.commit_or_raise("the conversion")
            if overwrite and os.path.exists(out):
                os.remove(out)
            _session.com_retry(
                lambda: wb.SaveAs(os.path.abspath(out),
                                  FILE_FORMATS[fmt]),
                label="SaveAs")
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, f"converting to {fmt}") from exc
        note = None
        if fmt == "csv":
            note = "CSV carries the FIRST worksheet's cells only (values, " \
                   "no formulas or formatting); that is the format, not a bug"
        return {"converted": out, "format": fmt,
                **({"note": note} if note else {})}

    def job_wrapper(app, wb):
        return body(app, wb)

    result = _run_readonly(f"convert({Path(p).name} -> {fmt})", p,
                           job_wrapper, timeout=timeout_seconds, allow_network=allow_network)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ValidationFailed(
            "Excel reported success but the converted file is missing or "
            "empty")
    result.update(out_info)
    result["ok"] = True
    result["file"] = p
    return result


# =========================================================== encrypt / misc


def com_save_with_password(path: str, password: str,
                           current_password: str | None = None,
                           timeout_seconds: float | None = None,
                           backup: bool = True,
                           allow_network: bool = False) -> dict:
    """Password-protect (or re-key / decrypt) a workbook using Excel's real
    encryption. password='' removes the protection (current_password then
    required). Holds the workbook's write lock from before Excel opens it
    until the encrypted copy is in place and checked, as _run_mutation
    does."""
    p = _norm_path(path, "com_save_with_password")
    _safesave.refuse_backup_slot_target(p, "encrypt")  # never a backup in place
    if not isinstance(password, str):
        raise XlMcpError("password must be a string ('' removes protection)")
    if password == "" and not current_password:
        raise XlMcpError(
            "removing a password needs current_password to open the file")
    _xmlsafe.check_file(p)  # a part with a DTD: refused before Excel starts
    _refuse_encrypted_without_password(
        p, current_password if current_password else None)
    _hazard.refuse_if_write_reserved(p)  # before Excel starts
    _refuse_legacy_biff(p)  # this route has its own open
    _refuse_filter_table_damage(p)  # before Excel starts
    _refuse_network_before_excel(p, allow_network, "save_with_password")
    copy = _session.copy_path_for(p)  # beside it, or refused now
    with _safesave.write_lock(p):
        return _encrypt_locked(p, password, current_password, copy=copy,
                               timeout_seconds=timeout_seconds, backup=backup,
                               allow_network=allow_network)


def _encrypt_locked(p: str, password: str, current_password: str | None, *,
                    timeout_seconds: float | None, backup: bool,
                    copy: str | None = None,
                    allow_network: bool = False) -> dict:
    """com_save_with_password's work, under the workbook's write lock."""
    if copy is None:
        copy = _session.copy_path_for(p)
    stamp = _safesave.file_stamp(p)  # before the backup and Excel's open
    warnings = _session.guard_target_closed(p)
    backup_slot = None
    ticket = None
    if backup:
        ticket = _safesave.prepare_rotation(p)
    state = _new_save_state(copy)

    def job(manager) -> dict:
        w = manager.acquire()
        wb = None
        staged = None
        prior = _session._assert_hygiene(w.app)
        try:
            try:
                wb = _session.open_workbook(w.app, p,
                                            password=current_password,
                                            allow_network=allow_network,
                                            network_op="save with a password")
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(
                    exc, "opening the workbook (a wrong current_password "
                    "surfaces here)") from exc
            _refuse_if_changed(p, stamp)  # Excel has the stamped file
            # The point of no return, checked before Excel saves.
            _deadline.commit_or_raise("the encrypted save")
            state["save_started"] = True
            try:
                wb.Password = password
                # Excel writes an encrypted copy; the file
                # itself is replaced only after Excel has closed it.
                staged = _session.save_workbook(wb, label="encrypted save", path=p,
                                                copy=copy)
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(exc, "applying the password") from exc
            _keep_copy(state, staged)
            wb.Close(False)
            wb = None
            return {"instance_pid": w.pid, "staged": staged}
        except BaseException:
            _session.discard_saved(staged)
            raise
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            _session._restore_hygiene(w.app, *prior)
            manager.release_to_pool(w)

    try:
        out = _session.run_com(f"encrypt({Path(p).name})", job,
                               timeout=timeout_seconds)
    except BaseException as exc:
        # Excel only ever wrote a copy, so the workbook is as it was.
        _settle_failed_save(p, ticket, state, exc)
        raise
    encrypted = password != ""
    digest = None
    if out.get("staged"):
        # The copy is checked before it replaces the
        # workbook. An encrypted save's copy must be Excel's encryption (a
        # plain package means the password did not take); a decrypted one
        # must pass the structural verify.
        check = _session.check_saved_copy(p, encrypted=encrypted)
        try:
            digest = _session.put_in_place(out["staged"], p, stamp=stamp,
                                           check=check, warnings=warnings,
                                           ticket=ticket)
        except BaseException:
            if ticket is not None:
                ticket.abort()  # the workbook is as it was; so are the slots
            raise
    elif ticket is not None:
        ticket.abort()
    if ticket is not None and ticket.rotated["prev"]:
        backup_slot = "prev"
    _confirm_in_place(p, digest, backup_slot)
    result = {
        "ok": True, "file": p, "saved": True, "backup": backup_slot,
        "engine": "com", "instance_pid": out.get("instance_pid"),
        "changed": {"encryption": "applied" if encrypted else "removed"},
        "verified": not encrypted,
        "warnings": warnings + ([
            "the file is now encrypted; every file-based tool on this "
            "server will refuse it until the password is removed. Keep the "
            "password safe; it is NOT recoverable"] if encrypted else []),
    }
    return _carry_network_note(result, out)


def com_autofit(path: str, sheet: str | None = None,
                columns: str | None = None, rows: str | None = None,
                timeout_seconds: float | None = None,
                backup: bool = True, allow_network: bool = False) -> dict:
    """True column/row autofit using Excel's real text metrics.

    allow_network: a workbook whose formulas or data connections can reach
    the network is refused before Excel opens it (Excel calculates as it
    opens), unless this is true."""
    p = _norm_path(path, "com_autofit")

    def body(app, wb) -> dict:
        try:
            ws = _ws(wb, sheet)
            done: dict[str, str] = {"sheet": str(ws.Name)}
            if columns is None and rows is None:
                ws.UsedRange.EntireColumn.AutoFit()
                done["columns"] = "all used columns"
            if columns is not None:
                ws.Columns(columns).AutoFit()
                done["columns"] = columns
            if rows is not None:
                ws.Rows(rows).AutoFit()
                done["rows"] = rows
            return done
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, "autofitting (columns like 'A:D', "
                                  "rows like '1:20')") from exc

    return _run_mutation(f"autofit({Path(p).name})", p, body,
                         backup=backup, timeout=timeout_seconds, allow_network=allow_network)


def com_goal_seek(path: str, target_cell: str, target_value: float,
                  changing_cell: str, sheet: str | None = None,
                  save: bool = True, timeout_seconds: float | None = None,
                  backup: bool = True, allow_network: bool = False) -> dict:
    """Excel's Goal Seek: adjust changing_cell until target_cell (a formula)
    reaches target_value; convergence-checked. Goal Seek recalculates, so a
    workbook that can reach the network is refused as recalculate refuses
    it, unless allow_network is true."""
    p = _norm_path(path, "com_goal_seek")
    try:
        goal = float(target_value)
    except (TypeError, ValueError):
        raise XlMcpError("target_value must be a number") from None
    reach = _network_gate(p, allow_network, "goal_seek")

    def body(app, wb) -> dict:
        try:
            ws = _ws(wb, sheet)
            tgt = ws.Range(target_cell)
            chg = ws.Range(changing_cell)
            if not str(tgt.Formula or "").startswith("="):
                raise XlMcpError(
                    f"target_cell {target_cell} holds no formula; Goal Seek "
                    "adjusts an input until a FORMULA reaches the goal")
            converged = bool(tgt.GoalSeek(goal, chg))
            result_value = tgt.Value
            input_value = chg.Value
            if not converged:
                raise XlMcpError(
                    f"Goal Seek did not converge: {target_cell} reached "
                    f"{result_value!r} against the goal {goal!r}. The file "
                    "was not saved.")
            return {"converged": True, "target_cell": target_cell,
                    "target_value": goal, "achieved": result_value,
                    "changing_cell": changing_cell,
                    "solution": input_value}
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(exc, "running Goal Seek") from exc

    if save:
        return _note_network(_run_mutation(f"goal_seek({Path(p).name})", p, body,
                                           backup=backup, timeout=timeout_seconds,
                                           allow_network=allow_network),
                             reach)
    result = _note_network(_run_readonly(f"goal_seek({Path(p).name})", p, body,
                                         timeout=timeout_seconds,
                                         allow_network=allow_network), reach)
    result["ok"] = True
    result["file"] = p
    result["saved"] = False
    result["note"] = "save:false; the solution is reported but not persisted"
    return result


def _qualify_sparkline_source(source: str, sheet_name: str) -> str:
    """Bind an UNQUALIFIED sparkline source to the group's OWN sheet.

    Excel's ``SparklineGroups.Add`` resolves a bare "B2:F10" against the ACTIVE
    sheet, not the sheet the group lives on. Under the pooled hidden worker
    (whose active sheet is whatever the file stored) an unqualified source
    silently binds the sparkline to the WRONG sheet's data -- the field
    report's "list reports Pivot2!" was that mis-binding, not a display glitch
    (the stored xm:f was literally Pivot2!). Qualifying with the group's sheet
    binds it where the caller meant and makes ``.SourceData`` read back the
    right sheet. A caller-supplied "Sheet!ref" (a deliberate cross-sheet
    sparkline) already carries its sheet and is left untouched.
    """
    if "!" in source:
        return source
    nm = sheet_name
    if re.search(r"[^A-Za-z0-9_]", nm):
        nm = "'" + nm.replace("'", "''") + "'"
    return f"{nm}!{source}"


def com_set_sparkline(path: str, action: str = "create",
                      location: str | None = None,
                      source: str | None = None,
                      type: str = "line", sheet: str | None = None,
                      timeout_seconds: float | None = None,
                      backup: bool = True, allow_network: bool = False) -> dict:
    """Create, clear, or list sparkline groups through Excel (the file tier
    defers sparklines here: they live in x14 extLst XML openpyxl cannot
    round-trip)."""
    p = _norm_path(path, "com_set_sparkline")
    action = str(action).strip().lower()
    if action not in ("create", "clear", "list"):
        raise XlMcpError(
            f"action must be create, clear, or list; got {action!r}")

    if action == "list":
        def body(app, wb) -> dict:
            groups: list[dict] = []
            # sheet scopes the inventory. Create and clear have always
            # honored it; list ignored it and returned the whole workbook,
            # so a caller that named a sheet got an answer about sheets it
            # had not asked about and no signal that it had been ignored.
            if sheet is not None:
                sheets = [_ws(wb, sheet)]
            else:
                sheets = [wb.Worksheets(i)
                          for i in range(1, int(wb.Worksheets.Count) + 1)]
            for ws in sheets:
                try:
                    # ws.Cells, not UsedRange: sparkline cells hold no
                    # values, so they can sit OUTSIDE the used range.
                    scope = ws.Cells
                    cnt = int(scope.SparklineGroups.Count)
                except Exception:
                    cnt = 0
                for j in range(1, cnt + 1):
                    g = scope.SparklineGroups.Item(j)
                    groups.append({
                        "sheet": str(ws.Name),
                        "location": str(g.Location.Address),
                        "source": str(g.SourceData),
                        "type": {v: k for k, v in SPARK_TYPES.items()}.get(
                            int(g.Type), str(g.Type)),
                    })
            return {"sparkline_groups": groups}
        return _run_readonly(f"sparkline_list({Path(p).name})", p, body,
                             timeout=timeout_seconds, allow_network=allow_network)

    if not location:
        raise XlMcpError("location is required (the cell/range that shows "
                         "the sparklines, e.g. 'G2:G10')")

    if action == "clear":
        def body(app, wb) -> dict:
            try:
                ws = _ws(wb, sheet)
                ws.Range(location).SparklineGroups.Clear()
                return {"sparklines_cleared": location,
                        "sheet": str(ws.Name)}
            except XlMcpError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise _wrap_com_error(exc, "clearing sparklines") from exc
        return _run_mutation(f"sparkline_clear({Path(p).name})", p, body,
                             backup=backup, timeout=timeout_seconds, allow_network=allow_network)

    # create
    kind = str(type).strip().lower()
    if kind not in SPARK_TYPES:
        raise XlMcpError(
            f"type must be one of {sorted(SPARK_TYPES)}, got {type!r}")
    if not source:
        raise XlMcpError("create needs source: the data range the "
                         "sparklines chart, e.g. 'A2:F10'")

    def body(app, wb) -> dict:
        try:
            ws = _ws(wb, sheet)
            bound = _qualify_sparkline_source(source, str(ws.Name))
            group = ws.Range(location).SparklineGroups.Add(SPARK_TYPES[kind], bound)
            return {"sparklines_created": location, "source": bound,
                    "type": kind, "sheet": str(ws.Name),
                    "count": int(group.Count) if hasattr(group, "Count")
                    else None}
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise _wrap_com_error(
                exc, "creating sparklines (location and source must be "
                "compatible shapes: one sparkline cell per source row or "
                "column)") from exc

    return _run_mutation(f"sparkline_create({Path(p).name})", p, body,
                         backup=backup, timeout=timeout_seconds, allow_network=allow_network)


# ======================================================= validate / status


def com_validate_opens_clean(path: str, password: str | None = None,
                             timeout_seconds: float | None = None,
                             allow_network: bool = False) -> dict:
    """Open the file in a private hidden Excel and report whether Excel
    accepts it without a repair prompt (the authoritative corruption smoke
    test). Read-only; never saves."""
    p = check_path(path, "com_validate_opens_clean")
    _xmlsafe.check_file(p)  # a part with a DTD: refused before Excel starts
    _refuse_network_before_excel(p, allow_network, "validate that it opens clean")
    if password is not None:
        def body(app, wb) -> dict:
            return {"opens_clean": True, "workbook": str(wb.Name),
                    "worksheets": int(wb.Worksheets.Count)}
        try:
            result = _run_readonly(f"opens_clean({Path(p).name})", p, body,
                                   timeout=timeout_seconds,
                                   password=password,
                                   allow_network=allow_network)
        except XlMcpError:
            raise
        except Exception as exc:  # noqa: BLE001
            return {"ok": True, "file": p, "opens_clean": False,
                    "excel_says": _session.excel_error_message(exc)}
        result["ok"] = True
        result["file"] = p
        return result
    result = _session.opens_clean(p, timeout=timeout_seconds,
                                  allow_network=allow_network)
    result["ok"] = True
    result["file"] = p
    return result


def com_status() -> dict:
    """Honest COM layer status: availability, worker/pool state, journaled
    PIDs, contention and timing. Never spawns Excel."""
    status = _session.get_executor().status()
    status["ok"] = True
    status["policy"] = (
        "private pooled hidden workers only; this server never attaches to "
        "or edits your open Excel session, and refuses files Excel holds "
        "open")
    return status


__all__ = [
    "recalculate", "com_manage_pivot", "com_export_pdf", "com_render_sheet",
    "com_convert_format", "com_save_with_password", "com_autofit",
    "com_goal_seek", "com_set_sparkline", "com_validate_opens_clean",
    "com_status", "AGG_FUNCS", "FILE_FORMATS", "SPARK_TYPES",
]
