"""One call, end to end: arguments in, one compact envelope out (the same
design as the Word family-surface dispatcher, itself taken from the PDF
server's surface/dispatch.py; decode_json_text is the Word editor's, verbatim).

    parse {op, args}            BAD_PARAMS for any other shape
    resolve op                  NOT_AVAILABLE (launch settings / read-only),
                                BAD_PARAMS naming the right tool, NOT_FOUND
    decode JSON text            only where the schema admits an array or an
                                object and never a string
    validate args               the op's generated model, extra="forbid"
    run, wall-clock bounded     ONE Deadline per call; the op runs
                                in a worker thread with the deadline active,
                                and the safe-save routine commits against it
                                before touching the disk
    success envelope            {ok: true, file?, ...result} (ok and file
                                kept because the Workflow product checks file)
    output cap                  core/budget.enforce on EVERY result
    refusal                     Excel's shape {ok: false, error: {code, message,
                                hint, ...}}; anything unexpected is
                                INTERNAL_ERROR, scrubbed

Excel refuses bad arguments with its own code, BAD_PARAMS (the family's
BAD_ARGS); the product's closed vocabulary is kept (envelope.py
"adapted": keep the product's existing refusal shape).

NUMBERS SAFETY. decode_json_text never touches a parameter whose
schema admits a string, so a cell value (set_cell value, write_range rows,
set_cells writes) that arrives as text stays text: "[1,2]" and "1e3" land as
the strings they were sent as, exactly as before the family surface.

Nothing below this module ever sees an MCP type, and nothing above it ever
sees an exception.
"""

from __future__ import annotations

import asyncio
import contextlib
import contextvars
import functools
import itertools
import json
import threading
from typing import Any

from pydantic import ValidationError

from .. import envelope, registry, runtime
from ..core import budget, deadline, protection_gate
from ..core import package as _package
from ..ops import backups as _backups
from ..core.deadline import Deadline
from ..core.errors import SurfaceRefusal
from ..registry import OpSpec
from . import copy, normalize
from .groups import HELP_TOOL
from .plan import SurfacePlan

#: Extra seconds the backstop waits past the bound before answering.
BACKSTOP_GRACE_S = 5.0

#: Keys the success envelope reads the call's file from (the order of the
#: released Excel _SuccessEnvelope._FILE_KEYS).
FILE_KEYS = ("file_path", "target_path", "path", "workbook", "src", "dst")

_TOOL_KEYS = frozenset({"op", "args"})


def _refuse(code: str, message: str, *, details: dict | None = None,
            hint: str | None = None) -> SurfaceRefusal:
    return SurfaceRefusal(code, message, details=details, hint=hint)


# ------------------------------------------------------------------ parsing


def parse(arguments: Any) -> tuple[str, dict[str, Any]]:
    if not isinstance(arguments, dict):
        raise _refuse("BAD_PARAMS", "call a tool with an object: {\"op\": ..., \"args\": {...}}")
    extra = sorted(set(arguments) - _TOOL_KEYS)
    if extra:
        raise _refuse("BAD_PARAMS", f"unknown top-level key(s) {', '.join(extra)}; op arguments "
                      "go inside args", details={"unknown": extra})
    op = arguments.get("op")
    if not isinstance(op, str) or not op:
        raise _refuse("BAD_PARAMS", "op is required and must be an op name")
    args = arguments.get("args")
    if args is None:
        args = {}
    if not isinstance(args, dict):
        raise _refuse("BAD_PARAMS", "args must be an object")
    return op, dict(args)


def resolve(plan: SurfacePlan, tool: str, op: str) -> OpSpec:
    """The op's spec, or the refusal that says exactly why not."""
    specs = registry.public_ops()
    if op in plan.forbidden_ops or plan.op_tool.get(op) in plan.forbidden_tools:
        raise _refuse("NOT_AVAILABLE", f"{op} is turned off by this server's launch settings",
                      details={"op": op, "decided_by": copy.SETTINGS_OWNER})
    if op in plan.hidden_by_read_only:
        raise _refuse("NOT_AVAILABLE", f"{op} changes files and this server runs read-only",
                      details={"op": op, "decided_by": copy.SETTINGS_OWNER})
    if op not in specs or op not in plan.op_tool:
        # (an op this edition leaves off its tool list is not there at all)
        raise unknown_op(plan, op, tool)
    owner = plan.op_tool.get(op)
    if owner != tool and not may_reroute(plan, tool, specs[op]):
        raise _refuse("BAD_PARAMS", f"{op} belongs to the {owner} tool, not {tool}",
                      details={"op": op, "tool": owner})
    return specs[op]


def may_reroute(plan: SurfacePlan, tool: str, spec: OpSpec) -> bool:
    """An op sent to the wrong tool runs anyway (the same rule as the Word
    editor's) ONLY when the tool it was sent to promises no
    less than the op does: an op that writes never runs through a read-only
    tool, and an op that edits files in place never runs through a tool that
    does not say so. The call is echoed in normalized_from.tool."""
    sent = plan.tools.get(tool)
    if sent is None or tool == HELP_TOOL:
        return False
    if spec.writes and sent.read_only:
        return False
    if spec.destructive and not sent.destructive:
        return False
    return True


def unknown_op(plan: SurfacePlan, op: str, tool: str | None) -> SurfaceRefusal:
    """NOT_FOUND for an op name that routes nowhere: nothing ran; at most
    three suggestions, each with its tool and one-line purpose, and never a
    destructive op for a name without that verb."""
    from .plan import op_text
    specs = registry.public_ops()
    visible = [n for n in specs if plan.is_visible(n)]
    near = [{"op": name, "tool": plan.op_tool[name], "purpose": op_text(specs[name])}
            for name in normalize.suggestions(op, visible)]
    details: dict[str, Any] = {"op": op}
    if tool is not None:
        details["tool"] = tool
    if near:
        details["did_you_mean"] = near
    return _refuse("NOT_FOUND", f"there is no op named {op}; nothing was run",
                   details=details, hint=copy.UNKNOWN_OP_HINT)


def shape(plan: SurfacePlan, arguments: Any) -> tuple[str, dict[str, Any], dict[str, Any]]:
    """(canonical op, args, normalized_from) from what was sent.
    The op name routes by its EXACT folded spelling; an op's own argument
    names sent beside op move into args. Everything else is parse()'s."""
    sent: dict[str, Any] = {}
    if isinstance(arguments, dict) and isinstance(arguments.get("op"), str) and arguments["op"]:
        specs = registry.public_ops()
        canonical = normalize.route_op(arguments["op"], specs)
        if canonical is not None and canonical != arguments["op"]:
            sent["op"] = arguments["op"]
            arguments = {**arguments, "op": canonical}
        if canonical is not None:
            fields = specs[canonical].args.model_fields
            arguments, moved = normalize.flatten(arguments, fields)
            if moved:
                sent["flat_args"] = moved
    op, args = parse(arguments)
    return op, args, sent


# ------------------------------------------------------------ JSON text
# Small models sometimes send a nested array or object as JSON
# TEXT inside args ("edits": "[{...}]"). When that parameter's schema admits
# an array or an object and does NOT admit a string, the text is decoded
# before validation, and the result records it. A parameter that admits a
# string is never decoded; text that is not valid JSON of an admitted kind
# is left alone and refused by validation as before.


@functools.lru_cache(maxsize=512)
def _schema(model: type) -> dict[str, Any]:
    return model.model_json_schema()


def _admits(schema: Any, root: dict[str, Any], seen: frozenset[str] = frozenset()) -> set[str]:
    """The JSON kinds a (sub)schema admits: object, array, string, other, any.
    Follows anyOf, oneOf, allOf and local $ref."""
    if not isinstance(schema, dict) or not schema:
        return {"any"}
    kinds: set[str] = set()
    ref = schema.get("$ref")
    if isinstance(ref, str):
        if not ref.startswith("#/") or ref in seen:
            return {"any"}
        node: Any = root
        for part in ref[2:].split("/"):
            node = node.get(part) if isinstance(node, dict) else None
        kinds |= _admits(node, root, seen | {ref})
    for key in ("anyOf", "oneOf"):
        for sub in schema.get(key) or []:
            kinds |= _admits(sub, root, seen)
    if schema.get("allOf"):
        parts = [_admits(sub, root, seen) for sub in schema["allOf"]]
        inter = set.intersection(*parts) if parts else set()
        if all("any" in p for p in parts):
            inter = {"any"}
        kinds |= inter
    t = schema.get("type")
    if t is not None:
        for one in (t if isinstance(t, list) else [t]):
            kinds.add(one if one in ("object", "array", "string") else "other")
    elif "properties" in schema or "additionalProperties" in schema:
        kinds.add("object")
    elif "items" in schema or "prefixItems" in schema:
        kinds.add("array")
    for key in ("const", "enum"):
        if key in schema:
            vals = schema[key] if key == "enum" else [schema[key]]
            for v in vals:
                kinds.add("string" if isinstance(v, str) else "array" if isinstance(v, list)
                          else "object" if isinstance(v, dict) else "other")
    if not kinds and not ({"$ref", "anyOf", "oneOf", "allOf"} & set(schema)):
        return {"any"}
    return kinds


def decode_json_text(spec: OpSpec, args: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """(args with JSON-text values decoded where the schema calls for it, the
    names decoded). Never touches a parameter that admits a string."""
    root = _schema(spec.args)
    props = root.get("properties") or {}
    decoded: list[str] = []
    out = dict(args)
    for name, value in args.items():
        if not isinstance(value, str) or name not in props:
            continue
        text = value.strip()
        if not text or text[0] not in "[{":
            continue
        kinds = _admits(props[name], root)
        if "any" in kinds or "string" in kinds or not ({"array", "object"} & kinds):
            continue
        try:
            parsed = json.loads(text)
        except (ValueError, RecursionError):
            continue
        if (isinstance(parsed, list) and "array" in kinds) or (isinstance(parsed, dict) and "object" in kinds):
            out[name] = parsed
            decoded.append(name)
    return out, decoded


def validate(spec: OpSpec, args: dict[str, Any]) -> dict[str, Any]:
    """Strict validation against the op's generated model. Returns the keyword
    arguments for the engine function: the validated (coerced) values of the
    arguments the caller gave; the function's own defaults cover the rest."""
    try:
        model = spec.args.model_validate(args)
    except ValidationError as exc:
        # An argument a weaker model got wrong names the form
        # that works, in the message a model reads first.
        message = copy.bad_params_message(
            spec.name, exc.errors(include_url=False, include_input=False))
        raise _refuse("BAD_PARAMS", message
                      or f"arguments for {spec.name} did not match its schema",
                      details={"op": spec.name,
                               "problems": registry.validation_problems(exc)}) from None
    return {name: getattr(model, name) for name in model.model_fields_set}


# ------------------------------------------------------------------ running


def execute(spec: OpSpec, kwargs: dict[str, Any], d: Deadline) -> Any:
    """Run one validated call (blocking; the caller bounds it) with the
    call's deadline active, so safe save commits against it."""
    with deadline.active(d):
        if spec.writes:
            _put_back_before_write(kwargs)
        return spec.call(kwargs)


#: Worker threads are named with this prefix. It starts with "asyncio" on
#: purpose: tests that wait for a call's late work to end look for threads
#: of that name (they were asyncio's default-executor threads before each
#: call got a thread of its own).
WORKER_PREFIX = "asyncio-ks4xl-call"

_worker_seq = itertools.count(1)

#: The worker threads still running, each with its call's deadline.
_INFLIGHT: dict[threading.Thread, Deadline] = {}
_INFLIGHT_LOCK = threading.Lock()


def drain_inflight(bound: float = 60.0) -> dict:
    """At server exit, let the calls still running finish (the same design
    as the Word editor; server.main()
    calls this after the transport closes, and nothing a caller sends can
    reach it).

    Worker threads are daemons, so the interpreter no longer waits for
    them the way it waited for asyncio's executor threads. A COM call that
    timed out may still be inside the private Excel with the workbook open.
    So, before the process exits:
      * every call that has NOT committed has its deadline cancelled, so it
        stops at its commit point instead of saving, and closes up;
      * every in-flight worker is waited for, `bound` seconds in total, so
        its Excel job ends and the workbook is released;
      * if any is still running past the bound, ONE stderr line names the
        hidden Excel PIDs this server's pool started, and the exit goes on.
    This step ends no process. (What happens to the pool after it is the
    COM executor's own shutdown.)
    Returns what it did, for the log and the tests."""
    import sys as _sys
    import time as _time

    end = _time.monotonic() + bound
    with _INFLIGHT_LOCK:
        items = [(t, d) for t, d in _INFLIGHT.items() if t.is_alive()]
    if not items:
        return {"cancelled": 0, "waited": 0, "unfinished": 0}
    cancelled = sum(1 for _t, d in items if not d.committed and d.cancel())
    for t, _d in items:
        t.join(max(0.0, end - _time.monotonic()))
    unfinished = sum(1 for t, _d in items if t.is_alive())
    if unfinished:
        from ..com import session as _com_session

        pids = _com_session.owned_excel_pids()
        _sys.stderr.write(
            f"[kitchensink4xl] {unfinished} call(s) still running after {bound:g}s at exit; "
            f"hidden Excel started by this server: {pids or 'none'}\n")
    return {"cancelled": cancelled, "waited": len(items), "unfinished": unfinished}


def _start_worker(fn, d: Deadline | None = None) -> "asyncio.Future":
    """Run fn on a NEW daemon thread, under a copy of the caller's context,
    and return a future for its result (the same design as the Word
    editor).

    asyncio.to_thread used the loop's default executor: min(32, cpu + 4)
    threads. A call that timed out keeps its thread (a thread cannot be
    killed), so enough stuck calls filled the executor and from then on
    every call, however trivial, queued behind them and answered TIMEOUT.
    One thread per call means a stuck call holds only its own thread: it is
    never a slot another call waits for. The thread is a daemon, so a stuck
    one cannot hold up the server's exit either."""
    loop = asyncio.get_running_loop()
    fut: asyncio.Future = loop.create_future()
    ctx = contextvars.copy_context()

    def settle(ok: bool, value) -> None:
        if fut.cancelled():
            return
        if ok:
            fut.set_result(value)
        else:
            fut.set_exception(value)

    def worker() -> None:
        if d is not None:
            with _INFLIGHT_LOCK:
                _INFLIGHT[threading.current_thread()] = d
        try:
            value = ctx.run(fn)
        except BaseException as exc:  # noqa: BLE001 (handed to the awaiting call)
            outcome = (False, exc)
        else:
            outcome = (True, value)
        finally:
            if d is not None:
                with _INFLIGHT_LOCK:
                    _INFLIGHT.pop(threading.current_thread(), None)
        try:
            loop.call_soon_threadsafe(settle, *outcome)
        except RuntimeError:
            pass  # the loop is gone: nobody is waiting for this answer

    # A call answered TIMEOUT never awaits its future; reading the late
    # outcome here keeps asyncio from logging "exception was never retrieved".
    fut.add_done_callback(lambda f: f.cancelled() or f.exception())
    threading.Thread(target=worker, daemon=True,
                     name=f"{WORKER_PREFIX}-{next(_worker_seq)}").start()
    return fut


async def run_bounded(fn, d: Deadline):
    """Run blocking work in its own thread with an outer wall-clock backstop.

    The backstop answers TIMEOUT only after cancelling the call's deadline,
    and cancelling fails once a save has committed. So a TIMEOUT answer means
    nothing was saved, and a call that did save always reports it: the late
    thread can never publish after TIMEOUT. A committed call is waited for
    up to d.post_commit_wait() more seconds; past that (Excel hung inside
    its own save) the answer is TIMEOUT with saved: null, never "nothing was
    saved" (core/deadline.py)."""
    task = _start_worker(fn, d)
    try:
        while True:
            left = d.remaining()
            wait = None if left == float("inf") else left + BACKSTOP_GRACE_S
            done, _pending = await asyncio.wait({task}, timeout=wait)
            if done:
                return task.result()
            if d.remaining() > 0:
                continue
            if d.cancel():
                raise d.timeout_error()
            # The save committed before the backstop fired: its answer is the
            # true one, and it is normally moments away (a rename and a
            # return). The wait is still bounded.
            done, _pending = await asyncio.wait({task}, timeout=d.post_commit_wait())
            if done:
                return task.result()
            raise d.timeout_error()
    except asyncio.CancelledError:
        d.cancel()
        task.cancel()  # the late outcome, if any, is dropped (settle checks)
        raise


def success(spec: OpSpec, kwargs: dict[str, Any], result: Any) -> dict[str, Any]:
    """The success shape the released server put on the wire: object
    results gain ok (and file, when the call named one); any other result
    rides under `result` with the same siblings."""
    head: dict[str, Any] = {"ok": True}
    for key in FILE_KEYS:
        value = kwargs.get(key)
        if isinstance(value, str):
            head["file"] = value
            break
    if isinstance(result, dict):
        if "ok" in result:
            return dict(result)
        return {**head, **result}
    return {**head, "result": result}


#: Ops that make something from a workbook without editing it (a copy, an
#: export, a render). They run on the copy last saved to disk.
_DISK_COPY_OPS = {"copy_workbook": "src", "com_export_pdf": "path",
                  "com_render_sheet": "path", "com_convert_format": "path",
                  "export_range": "path", "export_file": "path"}


def _note_open_in_excel(op: str, kwargs: dict[str, Any], payload: dict[str, Any]) -> None:
    """While Excel has the workbook open, an export or a
    copy says that it used the copy last saved to disk."""
    key = _DISK_COPY_OPS.get(op)
    path = kwargs.get(key) if key else None
    if not isinstance(path, str):
        return
    import os
    owner = os.path.join(os.path.dirname(os.path.abspath(path)), "~$" + os.path.basename(path))
    if os.path.exists(owner):
        warnings = payload.get("warnings")
        if not isinstance(warnings, list):
            payload["warnings"] = warnings = [] if warnings is None else [warnings]
        warnings.append(
            f"{os.path.basename(path)} is open in Excel; this used the copy last saved to "
            "disk, so changes not yet saved in Excel are not in it")


#: The arguments of a writing call that name a workbook it may write.
_WRITE_TARGETS = ("path", "dst", "dest", "output")


def _put_back_before_write(kwargs: dict[str, Any]) -> None:
    """Before a writing call, put back a workbook it names that
    a save which did not finish left aside, so the call starts from the
    workbook in its place and never writes over it or strands it. Reads
    never come here (they never move a file)."""
    from ..core import safesave as _safesave_mod
    for key in _WRITE_TARGETS:
        target = kwargs.get(key)
        if isinstance(target, str) and target:
            _safesave_mod.put_back_before_write(target)


def _refuse_two_sheets(kwargs: dict[str, Any]) -> None:
    """A location that names its own sheet beside a
    different sheet argument used to run on the location's sheet without a
    word. Excel compares sheet names without regard to case."""
    loc, sheet = kwargs.get("location"), kwargs.get("sheet")
    inner = loc.get("sheet") if isinstance(loc, dict) else None
    if isinstance(inner, str) and isinstance(sheet, str) and inner.casefold() != sheet.casefold():
        raise _refuse("BAD_PARAMS", f"location names sheet {inner!r} and sheet names {sheet!r}; "
                      "send one of them (nothing was run)",
                      details={"location.sheet": inner, "sheet": sheet})


def _requested_timeout(kwargs: dict[str, Any]) -> float | None:
    value = kwargs.get("timeout_seconds")
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


async def dispatch(plan: SurfacePlan, tool: str, arguments: Any) -> tuple[dict[str, Any], bool]:
    """The whole pipeline. Returns (payload, is_error); never raises."""
    cap = plan.settings.max_output_chars
    decoded: list[str] = []
    normalized_from: dict[str, Any] = {}
    try:
        op, args, normalized_from = shape(plan, arguments)
        if tool == HELP_TOOL:
            raise _refuse("INTERNAL_ERROR", "xl_help is not a group tool")
        spec = resolve(plan, tool, op)
        if plan.op_tool.get(spec.name) != tool:
            normalized_from["tool"] = tool
        args, decoded = decode_json_text(spec, args)
        args, sent = normalize.normalize_args(spec.name, args, spec.args.model_fields)
        if sent:
            normalized_from["args"] = sent
        kwargs = validate(spec, args)
        _refuse_two_sheets(kwargs)
        if kwargs.get("verify_com") is True:
            why = runtime.excel_forbidden()
            if why:
                raise _refuse("NOT_AVAILABLE", f"verify_com needs Excel, and {why}; the call "
                              "was not run. Leave verify_com out to save with the file checks.",
                              details={"op": spec.name, "decided_by": copy.SETTINGS_OWNER})
        d = Deadline(plan.settings.timeout_for(spec.timeout_s, _requested_timeout(kwargs)))
        # A writing call's saves honour sheet and workbook
        # protection (core.protection_gate, checked inside the save).
        # A destructive call keeps a labelled copy of the
        # workbook as it was (ops.backups.AUTO_SNAPSHOT), named in the result.
        snap_op = _backups.auto_snapshot_op(spec.name, kwargs) if spec.writes else None
        with (protection_gate.dispatched_call(spec.name) if spec.writes
              else contextlib.nullcontext()), _package.auto_snapshot(snap_op) as snap:
            result = await run_bounded(lambda: execute(spec, kwargs, d), d)
        payload = success(spec, kwargs, result)
        if snap is not None and snap["label"] and payload.get("ok") is not False:
            payload["snapshot"] = snap["label"]
        if payload.get("ok") is False:
            # An engine function that returned a refusal-shaped dict.
            return budget.cap_refusal(payload, cap), True
        if decoded:
            payload["json_text_decoded"] = decoded
        if normalized_from:
            payload["normalized_from"] = normalized_from
        _note_open_in_excel(spec.name, kwargs, payload)
        supports_offset = "offset" in spec.args.model_fields
        offset = kwargs.get("offset") if isinstance(kwargs.get("offset"), int) else 0
        payload = budget.enforce(payload, cap, supports_offset=supports_offset,
                                 requested_offset=offset)
        return budget.lead_with_totals(payload), False
    except asyncio.CancelledError:
        raise
    except BaseException as exc:  # noqa: BLE001 (every failure becomes an envelope)
        if isinstance(exc, (KeyboardInterrupt, SystemExit)):
            raise
        payload = envelope.refusal(exc)
        if decoded:
            # Recorded on refusals too: these arguments arrived as JSON text.
            payload["error"].setdefault("details", {}).setdefault("json_text_decoded", decoded)
        if normalized_from:
            payload["error"].setdefault("details", {}).setdefault("normalized_from",
                                                                  normalized_from)
        return budget.cap_refusal(payload, cap), True
