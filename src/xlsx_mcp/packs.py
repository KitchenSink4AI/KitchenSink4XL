"""Tiered loading: the pack registry and the enable/disable machinery.

Ported near-verbatim from KitchenSink4Word packs.py (proven in production
through two siblings) and adapted to the grid domain: env vars are
KS4XL_MODE / KS4XL_PACK_POLICY, and PACK_SUMMARIES carries the
consolidation-phase pack map (lite + three packs, re-cut 2026-09-04 against
measured bills under the cost-aware ruling). server.py registers every tool
up front and non-lite tools start disabled.

Which tools a session sees and may call comes from that session's own
record (punch-list #933). main() applies apply_startup_mode() to the process
default record (_ENABLED); each session copies it into a record of its
own when it initializes (packstate.py), and the server's own
middleware (packgate.py) filters tools/list and admits tools/call from that
record, sending tools/list_changed to that session only. No fastmcp
visibility state is used: fastmcp expired a session's visibility rules a
day after they were set, which is how an enabled pack used to vanish from a
long conversation while this module still said it was on.

Member lists are wired by server.py's @_tool decorator (the single source of
membership truth). The 2026-09-04 re-cut merged the four small planning packs
into two: format+objects+tables-names became design (the report-design usage
cluster) and data folded into io's inspector cluster; rationale per pack in
the DESIGN as-built notes.

Env contract:
- KS4XL_MODE: startup surface for clients without reliable list_changed.
  "lite" (default), "full", or a comma-separated pack list ("design,com").
- KS4XL_PACK_POLICY: "auto" (default; the CLIENT's permission prompt gates
  enable_tools, which is deliberately a plain tool call) or "locked"
  (enable_tools/disable_tools refuse; the surface is fixed at startup).

Saved choices (owner ruling 2026-09-26: "If someone switches it on, they
don't expect it to revert. I know I wouldn't."). Every enable_tools and
disable_tools call is saved (packstore.py) under the launch settings in
force, and main() starts the process default from the launch settings
plus those saved choices, so a choice stays until a person or the AI
changes it again. A launch-time lock (KS4XL_PACK_POLICY=locked) wins:
saved choices are not applied while it holds. The store directory can be
moved with KS4XL_PACK_STORE_DIR.
"""

from __future__ import annotations

import json
import os
import sys
from typing import Any

from . import packstate as _packstate
from . import packstore as _packstore
from .core.errors import TargetNotFound, XlMcpError

# Packs in menu order (cost-aware ruling 2026-09-02, re-cut 2026-09-04
# against measured bills: every sub-1.5k pack merged into a thematic
# neighbor; com stays separate at any size because it is environment-gated).
# "lite" is the always-on core, not a pack. "everything" is a convenience
# alias for all packs.
PACK_SUMMARIES: dict[str, str] = {
    "design": (
        "workbook design and rich features: named cell styles, format "
        "painter, style-bloat audit, conditional formatting, data "
        "validation, images, charts (create/list/delete), advanced table "
        "lifecycle (columns, totals, resize, banding), and named ranges "
        "(define, scope, LAMBDA, cleanup)"
    ),
    "io": (
        "page layout and print, headers/footers, advisory protection, "
        "legacy comments, multi-sheet export, and the read-side "
        "inspectors: external links, VBA, existing pivots, data "
        "connections"
    ),
    "com": (
        "drives a private hidden Excel instance (Windows + Excel "
        "required): real pivot tables, fidelity recalculation, goal seek, "
        "PDF export, sheet render to image, format conversion, real "
        "encryption, sparklines, true autofit, opens-clean validation, "
        "and honest status; never touches your open Excel session"
    ),
}
EVERYTHING = "everything"

#: Returned as `note` by EVERY successful enable(), including one that
#: enabled nothing new.
#:
#: The old sentence ("re-fetch the tool list if your client does not refresh
#: automatically") assumed the client could re-fetch. Several cannot: Claude
#: Code workers and Codex CLI fix their tool list when the session or worker
#: starts, so a pack enabled afterwards never becomes callable there and the
#: agent retries enable_tools forever against a surface that already says the
#: pack is on. That is why the note also rides the no-op re-enable: the second
#: call is exactly where a stuck agent lands, and it is the call that used to
#: come back with no note at all. The route out is KS4XL_MODE at launch, which
#: is why the note names it instead of promising a refresh.
#:
#: The note is PREFIX + BODY, because only the body is true unconditionally.
#:
#: A no-op re-enable sends NO notification: enable_tools in server.py sends
#: tools/list_changed only when this call switched a tool on, so no
#: ToolListChangedNotification reaches the session. Telling the caller
#: "tools/list_changed was sent" on that call was a plain falsehood, and the
#: worst possible one here: an agent debugging a tool it cannot see would go
#: looking for a notification that was never emitted.
LIST_CHANGED_PREFIX = "tools/list_changed was sent."

#: The honest prefix for the call where nothing changed state.
NO_LIST_CHANGE_PREFIX = (
    "These packs were already on, so no list change was sent."
)

#: The advice, identical either way. The launch-env route LEADS because it is
#: the one proven to work in every client, workers included. The orchestrator
#: route comes second and is labelled for the client it works in: in Codex CLI
#: a pack the parent enables never reaches a worker at all, so offering it
#: first would send most readers down a path that cannot help them.
CLIENT_NOTE_BODY = (
    "If the new tools are not in your tool list, this client fixed its list "
    "when the session or worker started: do not retry here. What works in "
    "every client: ask the user to add the packs to KS4XL_MODE (comma list) "
    "in this server's launch settings, restart the app or session, then "
    "start a new worker if needed. Claude Code only: the orchestrator can "
    "instead call enable_tools in the main session and then start a new "
    "worker. If enable_tools refuses a pack, an administrator locked the "
    "tool set: do not retry."
)


def client_note(list_changed: bool) -> str:
    """The enable() note. `list_changed` is whether a notification really
    went out, not whether the caller asked for one."""
    prefix = LIST_CHANGED_PREFIX if list_changed else NO_LIST_CHANGE_PREFIX
    return f"{prefix} {CLIENT_NOTE_BODY}"

#: The same fact, stated once where a client reads it before it calls
#: anything: the server instructions (one handshake) and the get_workflows
#: index. Single-sourced here so the two copies cannot drift.
WORKER_SURFACE_NOTE = (
    "Workers and subagents only see the tools that were on when they "
    "started: start the server with KS4XL_MODE set to a comma list of "
    "packs, or, in Claude Code, enable packs in the main session before "
    "starting workers."
)

# pack -> {tool_name: fastmcp Tool}; "lite" holds the always-on core.
_REGISTRY: dict[str, dict[str, object]] = {"lite": {}}

# tool_name -> enabled? The PROCESS DEFAULT record: the startup surface
# apply_startup_mode() sets (launch settings plus saved choices), which a
# session copies when it initializes. A session's own changes live in its
# own record (packstate.py), and are also applied here and saved, so a
# session that starts later starts from them; callers outside any MCP
# session (the test suite, the measurement scripts) read and write this
# one directly.
_ENABLED: dict[str, bool] = {}

#: The packs the launch settings turned on at startup, set by
#: apply_startup_mode(). Saved choices are kept under this set.
_STARTUP_PACKS: list[str] = []

#: Were saved choices applied at startup? False while the launch settings
#: lock the tool set.
_SAVED_APPLIED = False

#: Moves the saved-choices directory (the test suite points it at a
#: temporary one).
ENV_PACK_STORE = "KS4XL_PACK_STORE_DIR"
_STATE_DIR_NAME = "xlsx-mcp"


def default_record() -> dict[str, bool]:
    """The process default record: what a session starts from."""
    return _ENABLED


def _store_path():
    return _packstore.store_path(ENV_PACK_STORE, _STATE_DIR_NAME)


def _launch_key() -> str:
    return _packstore.launch_key(_STARTUP_PACKS)


def _remember(choices: dict[str, bool]) -> bool:
    """Make `choices` ({pack: on?}) the process default for the packs they
    name, so every session that starts later starts from them, and save
    them. Returns whether they were saved. Called under packstate.LOCK."""
    for pack, on in choices.items():
        for name in _REGISTRY.get(pack, {}):
            _ENABLED[name] = on
    return _packstore.save(
        _store_path(), _launch_key(), choices, _STARTUP_PACKS)


def saved_packs_report() -> dict:
    """The saved choices for this launch settings, and whether they were
    applied at startup (not while the tool set is locked)."""
    choices = _packstore.load(_store_path(), _launch_key(), PACK_SUMMARIES)
    return {
        "applied": _SAVED_APPLIED,
        "packs": {p: ("on" if on else "off") for p, on in choices.items()},
    }


def _apply_saved_choices(startup_packs: list[str]) -> None:
    """Start the process default from the saved choices for this launch,
    unless the launch settings lock the tool set."""
    global _STARTUP_PACKS, _SAVED_APPLIED
    _STARTUP_PACKS = list(startup_packs)
    _SAVED_APPLIED = False
    choices = _packstore.load(_store_path(), _launch_key(), PACK_SUMMARIES)
    if _policy_locked():
        if choices:
            sys.stderr.write(
                "[kitchensink4xl] saved pack choices not applied: the tool set "
                "is locked at startup\n")
        return
    for pack, on in choices.items():
        for name in _REGISTRY.get(pack, {}):
            _ENABLED[name] = on
    _SAVED_APPLIED = True
    if choices:
        listed = ", ".join(
            f"{p} {'on' if on else 'off'}" for p, on in sorted(choices.items()))
        sys.stderr.write(
            f"[kitchensink4xl] saved pack choices applied: {listed}\n")

#: "The session being served", the default for the session parameters
#: below. Passing None means "no session": the process default record.
_CURRENT: Any = object()


def _record(session: Any = _CURRENT) -> dict[str, bool]:
    """The record in force for a session: its own, or the process default
    when it has none (or there is no session)."""
    if session is _CURRENT:
        session = _packstate.current_session()
    own = _packstate.session_record(session)
    return _ENABLED if own is None else own


class PackOff(TargetNotFound):
    """A call to a tool whose pack is off in the calling session.

    Refused as NOT_FOUND through the normal refusal envelope, by the session
    pack gate and, when that is skipped, inside the tool's own boundary
    wrapper, so both give the same payload (punch-list #933, in the #51
    v2.1 erratum S1-A shape). The words are the disabled-tool signpost's,
    unchanged."""

    code = "NOT_FOUND"

    def __init__(self, tool_name: str, pack: str):
        super().__init__(
            f"tool {tool_name!r} exists but is currently disabled: it "
            f"belongs to the {pack!r} pack."
        )
        self.tool_name = tool_name
        self.pack = pack
        self.hint = (
            f"call enable_tools(packs=['{pack}']) to turn it on, then "
            "retry this call"
        )


def register(tool_name: str, pack: str | None, tool: object) -> None:
    """Called by server.py once per tool at import time. pack=None means
    the lite core (enabled at startup); anything else starts disabled."""
    key = pack or "lite"
    if key != "lite" and key not in PACK_SUMMARIES:
        raise ValueError(f"unknown pack {key!r} for tool {tool_name}")
    _REGISTRY.setdefault(key, {})[tool_name] = tool
    _ENABLED[tool_name] = key == "lite"


def pack_names() -> list[str]:
    return list(PACK_SUMMARIES)


def pack_tools(pack: str) -> list[str]:
    return sorted(_REGISTRY.get(pack, {}))


def pack_of(tool_name: str) -> str | None:
    for pack, tools in _REGISTRY.items():
        if tool_name in tools:
            return pack
    return None


def is_tool_enabled(tool_name: str, session: Any = _CURRENT) -> bool:
    return _record(session).get(tool_name, False)


def tool_names() -> dict[str, list[str]]:
    return {pack: sorted(tools) for pack, tools in _REGISTRY.items()}


def approx_tokens(tool: object) -> int:
    """Rough per-tool client cost at ~4 chars per token. Honest enough for the
    informed-approval report; not a billing meter.

    Measured off the tool as tools/list actually serializes it, not off a
    hand-picked pair of fields. It used to sum description + inputSchema only,
    which silently dropped outputSchema, annotations, _meta, name and title:
    the published lite/full figures understated the real wire cost by ~16%
    (fat audit 2026-09-08). A server whose pitch is that it tells you what it
    costs does not get to publish the flattering subset, so the estimator
    measures what it publishes.

    Compact separators, because the transport uses them: the default
    json.dumps spacing is not on the wire and counting it overshot the real
    69-tool surface by ~1,200 tokens. Verified against a live stdio
    tools/list, which is the only figure that can settle it.
    """
    try:
        payload = tool.to_mcp_tool().model_dump(  # type: ignore[attr-defined]
            exclude_none=True, by_alias=True, mode="json")
        return round(len(json.dumps(payload, ensure_ascii=False,
                                    separators=(",", ":"))) / 4)
    except Exception:  # noqa: BLE001
        # Anything that is not a live fastmcp Tool (a stub in a test, a future
        # fastmcp that renames the method) falls back to the old estimate
        # rather than breaking the surface report.
        desc = getattr(tool, "description", "") or ""
        try:
            schema = json.dumps(getattr(tool, "parameters", {}) or {})
        except (TypeError, ValueError):
            schema = ""
        return round((len(desc) + len(schema)) / 4)


def pack_cost(pack: str) -> int:
    return sum(approx_tokens(t) for t in _REGISTRY.get(pack, {}).values())


def surface_report(session: Any = _CURRENT) -> dict:
    """The session's active surface: enabled tool count and approx token
    bill."""
    record = _record(session)
    active = 0
    tokens = 0
    per_pack: dict[str, str] = {}
    for pack, tools in _REGISTRY.items():
        enabled = [t for n, t in tools.items() if record.get(n, False)]
        active += len(enabled)
        tokens += sum(approx_tokens(t) for t in enabled)
        per_pack[pack] = f"{len(enabled)}/{len(tools)} enabled"
    return {
        "active_tools": active,
        "approx_active_tokens": tokens,
        "packs": per_pack,
    }


_POLICIES = ("auto", "locked")


def pack_policy() -> str:
    """The validated KS4XL_PACK_POLICY value. A typo used to FAIL OPEN:
    KS4XL_PACK_POLICY=lockedd served with packs unlockable, while its
    sibling KS4XL_MODE=fulll refused to serve -- the one env var whose
    whole purpose is a security pin was the one that shrugged off a
    misspelling (fresh-eyes round, M-4). Unknown values now refuse loudly,
    at startup (apply_startup_mode) and at every policy consultation."""
    raw = os.environ.get("KS4XL_PACK_POLICY", "auto").strip().lower()
    if not raw:
        return "auto"
    if raw not in _POLICIES:
        raise XlMcpError(
            f"KS4XL_PACK_POLICY={raw!r} is not a recognized policy; use "
            f"one of {_POLICIES}. Refusing rather than letting a typo "
            "silently drop the host's lock.")
    return raw


def _policy_locked() -> bool:
    return pack_policy() == "locked"


def _validate(packs: list[str]) -> list[str]:
    if isinstance(packs, str):
        packs = [packs]
    if not isinstance(packs, list) or not packs:
        raise XlMcpError(
            f"packs must be a non-empty list from {pack_names()} "
            f"(or ['{EVERYTHING}'])"
        )
    out: list[str] = []
    for p in packs:
        name = str(p).strip().lower()
        if name == EVERYTHING:
            return list(PACK_SUMMARIES)
        if name == "lite":
            raise XlMcpError(
                "the lite core is always on; it cannot be enabled or "
                "disabled as a pack"
            )
        if name not in PACK_SUMMARIES:
            raise XlMcpError(
                f"unknown pack {p!r}; valid packs: {pack_names()} "
                f"(or '{EVERYTHING}' for all of them)"
            )
        if name not in out:
            out.append(name)
    return out


def enable(packs: list[str]) -> dict:
    """Idempotent enable. Reports what changed, the approx token cost added,
    and the resulting total surface."""
    if _policy_locked():
        err = XlMcpError(
            "KS4XL_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host. Ask the operator to change KS4XL_MODE or unlock "
            "the policy."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    session = _packstate.current_session()
    with _packstate.LOCK:
        record = _record(session)
        if session is not None:
            record = dict(record)  # copy-on-write: never edit a stored one
        enabled_now: list[str] = []
        already: list[str] = []
        tokens_added = 0
        flipped: set[str] = set()
        for pack in wanted:
            newly = False
            for name, tool in _REGISTRY.get(pack, {}).items():
                if not record.get(name, False):
                    record[name] = True
                    flipped.add(name)
                    tokens_added += approx_tokens(tool)
                    newly = True
            (enabled_now if newly else already).append(pack)
        if session is not None and flipped:
            _packstate.keep_session_record(session, record)
        saved = _remember({pack: True for pack in wanted})
        return {
            "enabled": enabled_now,
            "saved": saved,
            "already_enabled": already,
            "approx_tokens_added": tokens_added,
            **surface_report(session),
            # enable_tools sends tools/list_changed exactly when something
            # flipped, so the prefix tracks the real emission rather than
            # the caller's intent.
            "note": client_note(bool(flipped)),
        }


def disable(packs: list[str]) -> dict:
    """Idempotent disable; the lite core always stays on."""
    if _policy_locked():
        err = XlMcpError(
            "KS4XL_PACK_POLICY=locked: the tool surface is fixed at startup "
            "by the host."
        )
        err.code = "CONFLICT"
        raise err
    wanted = _validate(packs)
    session = _packstate.current_session()
    with _packstate.LOCK:
        record = _record(session)
        if session is not None:
            record = dict(record)  # copy-on-write: never edit a stored one
        disabled_now: list[str] = []
        already: list[str] = []
        tokens_removed = 0
        flipped: set[str] = set()
        for pack in wanted:
            newly = False
            for name, tool in _REGISTRY.get(pack, {}).items():
                if record.get(name, False):
                    record[name] = False
                    flipped.add(name)
                    tokens_removed += approx_tokens(tool)
                    newly = True
            (disabled_now if newly else already).append(pack)
        if session is not None and flipped:
            _packstate.keep_session_record(session, record)
        saved = _remember({pack: False for pack in wanted})
        return {
            "disabled": disabled_now,
            "saved": saved,
            "already_disabled": already,
            "approx_tokens_removed": tokens_removed,
            **surface_report(session),
        }


def apply_startup_mode() -> str:
    """Apply KS4XL_MODE at server start, before any client connects, to the
    process default record every session starts from. Returns the mode
    applied, for logging."""
    pack_policy()  # a misspelled policy pin refuses to serve, like a mode typo
    mode = os.environ.get("KS4XL_MODE", "lite").strip().lower()
    if not mode or mode == "lite":
        _apply_saved_choices([])
        return "lite"
    # "lite" and "full"/"everything" are mode tokens, tolerated inside
    # comma lists alike: lite is always on anyway, full means every pack.
    # Refusing lite or full bricked a sibling at startup; typos still fail
    # LOUDLY via _validate below.
    tokens = [p.strip() for p in mode.split(",") if p.strip()]
    wants_full = any(t in ("full", EVERYTHING) for t in tokens)
    named = [t for t in tokens if t not in ("lite", "full", EVERYTHING)]
    if named:
        _validate(named)  # raises on typos so a bad env fails LOUDLY
    packs = list(PACK_SUMMARIES) if wants_full else named
    valid = _validate(packs) if packs else []
    for pack in valid:
        for name in _REGISTRY.get(pack, {}):
            _ENABLED[name] = True
    _apply_saved_choices(valid)
    return mode if valid else "lite"


def menu() -> dict:
    """The full pack menu with per-pack tool lists and approx token costs."""
    return {
        pack: {
            "summary": PACK_SUMMARIES[pack],
            "tools": pack_tools(pack),
            "approx_tokens": pack_cost(pack),
        }
        for pack in PACK_SUMMARIES
    }
