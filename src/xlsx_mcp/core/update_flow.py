"""The update notice, and the update a user chooses from it.

FAMILY MODULE: byte-identical in KitchenSink4Word, KitchenSink4XL and
KitchenSink4PPT. Every product fact (name, package, variable names, repository,
the tool that holds the ops) comes from core/update_check.py, so a port copies
this file unchanged. The per-product hook is three calls in surface/tools.py
(before_call, end_call, after_call) plus the two ops in server.py.

What it does, and the rule each part keeps:

- THE NOTICE. After a read-only tool call (or help) has produced its result,
  after_call() asks whether a notice is due: the cache shows a newer stable
  release, notices are not off, and none was shown or attempted in the last
  seven days (one window across versions). The attempt is CLAIMED before
  anything is sent: an in-process timestamp under a lock, then last_attempt
  in update-notices.json under a lock file, so two parallel calls or two
  server processes give one notice, and an unwritable state folder gives at
  most one per process.
  - A client that declared FORM elicitation, on a protocol version in
    DIALOG_PROTOCOLS, outside a background task, gets a dialog: "{PRODUCT} {NEW} is
    available. You have {CURRENT}. Update now?" with "Update now" / "Not now".
    The wait is bounded (60 s); a late answer is cancelled and never breaks
    the session. An accept, or a decline or cancel that takes at least
    1.5 s, is the user's answer: anything but "Update now" means ask again
    in seven days. A decline or cancel faster than that (a client answering
    by itself), a JSON-RPC error such as -32601, any exception or the
    timeout means the dialog is unavailable: the `update` key goes on that
    same result and this process uses the key from then on.
  - Any other client gets one `update` key on that successful result:
    {"notice": ..., "relay": ...}, never when the result already has one and
    never past the output cap (then nothing is recorded and it waits).
  Never after a call that wrote: a dialog there could turn a completed write
  into a reported failure.
- THE UPDATE. Only after "Update now" in the dialog, or an explicit update_now
  call (which first asks its own Yes/No where the dialog exists). It acts only
  on the cached, validated newer stable version and runs in ONE background
  child; no tool call waits on an install. Routes, each failing closed to the
  manual command:
  - uvx (sys.prefix in uv's archive cache, uv named by the absolute UV
    variable): [UV, "tool", "run", "--refresh-package", PACKAGE,
    "--refresh-package", CORE_PACKAGE, PACKAGE, "--version"], no shell, stdin closed, no window, a hard timeout; success
    only when the printed version is at least NEW.
  - uv tool install and pip: nothing runs; the exact command comes back.
  - Claude Desktop bundle (<P>_INSTALL_ROUTE=mcpb from the bundle manifest):
    the release's {package}-{NEW}.mcpb and its .sha256 from GitHub (fixed URL,
    GitHub hosts only), hash and in-memory bundle checks, a fixed file name in
    the state folder, then os.startfile so Desktop shows its own install
    window, which is the user's consent.
- SETTINGS that survive updates: update-notices.json beside the update cache
  (honours <P>_UPDATE_CACHE_DIR), the update_notices op (on, off, status) from
  any client, <P>_UPDATE_NOTICES=on|off for administrators, and
  <P>_UPDATE_CHECK=off over everything.
- Business: nothing here runs (both ops refuse, and the surface leaves them out).

Nothing from PyPI or from the AI reaches a command line, a URL or a message
except a version string that packaging validated.
"""

from __future__ import annotations

import asyncio
import atexit
import contextlib
import contextvars
import dataclasses
import hashlib
import io
import json
import logging
import os
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Annotated, Any, Literal

from packaging.version import InvalidVersion, Version
from pydantic import Field

from . import update_check as uc

PRODUCT = uc.PRODUCT
PACKAGE = uc.PACKAGE

#: The two ops this module provides, in the tool named by uc.UPDATE_TOOL.
OPS = ("update_now", "update_notices")

NOTICE_INTERVAL = timedelta(days=7)
#: A last_attempt this far in the future is a clock that moved back: ignored.
FUTURE_SKEW = timedelta(days=1)
DIALOG_TIMEOUT_S = 60.0
INSTALL_TIMEOUT_S = 120.0
UPDATE_NOW_WAIT_S = 30.0
DOWNLOAD_TIMEOUT_S = 30.0
MAX_BUNDLE_BYTES = 1_000_000
LOCK_STALE_S = 30.0

NOTICES_NAME = "update-notices.json"
LOCK_NAME = "update-notices.lock"
BUNDLE_SAVE_NAME = f"{PACKAGE}-update.mcpb"

#: Protocol versions on which a form dialog is sent; any other era degrades
#: to the `update` key. Later eras cannot be negotiated on the pinned stack.
DIALOG_PROTOCOLS = frozenset({"2025-06-18", "2025-11-25"})
#: Hosts a bundle download may be redirected to.
GITHUB_HOSTS = frozenset({"github.com", "objects.githubusercontent.com",
                          "release-assets.githubusercontent.com"})
#: The mcp_config.env entries a bundle may carry, each with its only value:
#: the install route, plus any user_config substitutions the product's
#: bundle declares (uc.BUNDLE_USER_ENV; none by default).
BUNDLE_ENV = {uc.INSTALL_ROUTE_ENV: "mcpb", **getattr(uc, "BUNDLE_USER_ENV", {})}

UPDATE_NOW = "Update now"
NOT_NOW = "Not now"

# ------------------------------------------------------------------ the copy

DIALOG_MESSAGE = "{product} {new} is available. You have {current}. Update now?"
OFF_NOTE = (f"You can turn update notices off at any time: tell your AI to turn off "
            f"{PRODUCT} update notices.")
FALLBACK_NOTICE = "{product} {new} is available (you have {current})."
RELAY_OFFER = ("Tell the user once, in one sentence, and offer to update: if they agree, "
               "call {tool} op update_now.")
RELAY_ONCE = "Tell the user this once, in one sentence, then continue."
MANUAL_HINT = ("run uvx --refresh {package} --version once, or pip install -U {package}, "
               "then restart your AI app. Claude Desktop users install the new bundle from "
               "the GitHub release.")
PIN_LINE = "If your app's settings name a version such as {package}==X, change it to {package}."

UPDATING = "{product} {new} is being installed in the background; the next answer says how it went."
STILL_UPDATING = ("{product} {new} is still being installed in the background; the next "
                  "answer says how it went.")
UVX_DONE = "{product} {new} is installed. To start using it, {reload}. " + PIN_LINE
UVX_FAILED = ("The update to {product} {new} did not finish: {reason}. Nothing changed. "
              "To update by hand, " + MANUAL_HINT)
PIP_COMMAND = "To update {product} to {new}, run pip install -U {package}, then restart your AI app."
UV_TOOL_COMMAND = ("To update {product} to {new}, close your AI app, run uv tool upgrade "
                   "{package}, then start the app again.")
BUNDLE_OPENED = ("The {product} {new} bundle is downloaded and checked, and Claude Desktop is "
                 "opening its install window. Choose Install there to finish.")
BUNDLE_FAILED = ("The {product} {new} bundle could not be prepared: {reason}. Nothing was "
                 "installed. To update by hand, download {package}-{new}.mcpb from "
                 "https://github.com/{repo}/releases/tag/v{new} and open it.")
NOT_NOW_TEXT = "Nothing was installed: the user chose Not now."

#: clientInfo.name selects one fixed sentence; the name is never echoed.
RELOAD_LINES: tuple[tuple[tuple[str, ...], str], ...] = (
    (("claude-code",), "run /mcp, select the {product} server, choose Reconnect"),
    (("cursor",), "turn the server off and on in Settings > MCP"),
    (("visual studio code", "vscode", "vs code"), "run MCP: Restart Server"),
    (("codex", "chatgpt", "openai", "copilot"), "restart the app"),
)
RELOAD_DEFAULT = "restart your AI app"


class UpdateRefused(Exception):
    """An op refusal; server.py turns it into the product's own refusal."""

    def __init__(self, code: str, message: str, hint: str | None = None) -> None:
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint


@dataclasses.dataclass
class NoticeChoice:
    """The dialog's one required field."""

    choice: Annotated[Literal["Update now", "Not now"], Field(description=OFF_NOTE)]


_log = logging.getLogger(uc.LOGGER_NAME)
_SAID: set[str] = set()


def _say_once(key: str, line: str) -> None:
    """One stderr line per process and key, through logging (the server's
    non-blocking stderr writer)."""
    if key in _SAID:
        return
    _SAID.add(key)
    with contextlib.suppress(Exception):
        _log.warning(line)


# ----------------------------------------------------------------- edition


def community() -> bool:
    """True only on the public channel. Business (and anything that cannot
    be read) gets none of this module."""
    try:
        if uc.resolve_backend().name != "pypi":
            return False
        return getattr(uc.license_policy.active(), "edition_id", None) != "business"
    except Exception:
        return False


# ------------------------------------------------------------------- state


def notices_path() -> Path:
    """update-notices.json beside the update cache: the server's own state
    folder, the same path for every version and install route."""
    return uc.cache_path().parent / NOTICES_NAME


def read_state() -> dict:
    """The stored state; anything unreadable means "nothing recorded"."""
    try:
        data = json.loads(notices_path().read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


@contextlib.contextmanager
def _state_lock(wait_s: float = 1.0):
    """A lock file beside the state, so two server processes claim one
    notice. Yields True when held, False when another process holds it,
    None when the folder cannot hold a lock at all (read-only): the
    in-process guard then still allows at most one attempt per process."""
    path = uc.cache_path().parent / LOCK_NAME
    fd = None
    held: bool | None = False
    deadline = time.monotonic() + wait_s
    while True:
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            held = True
            break
        except FileExistsError:
            try:
                if time.time() - path.stat().st_mtime > LOCK_STALE_S:
                    path.unlink()
                    continue
            except OSError:
                pass
            if time.monotonic() >= deadline:
                held = False
                break
            time.sleep(0.05)
        except OSError:
            held = None
            break
    try:
        yield held
    finally:
        if fd is not None:
            with contextlib.suppress(OSError):
                os.close(fd)
            with contextlib.suppress(OSError):
                path.unlink()


def _record(**fields) -> bool:
    """Re-read the state and change only the given fields, then write it
    atomically, so two processes never lose each other's choice."""
    with _state_lock():
        state = read_state()
        state.update(fields)
        return uc.atomic_write_json(notices_path(), state)


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _stamp(value: Any, now: datetime) -> datetime | None:
    """A stored UTC datetime, or None when it is absent, unreadable or more
    than a day in the future (a clock that moved back)."""
    if not isinstance(value, str):
        return None
    try:
        stamp = datetime.fromisoformat(value)
    except ValueError:
        return None
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    return None if stamp - now > FUTURE_SKEW else stamp


def notices_env() -> str | None:
    value = (os.environ.get(uc.NOTICES_ENV) or "").strip().lower()
    return value if value in ("on", "off") else None


def notices_off(state: dict | None = None) -> tuple[bool, str]:
    """(off, decided_by). <P>_UPDATE_NOTICES wins; =on also clears a stored
    off. Otherwise the stored choice decides."""
    env = notices_env()
    if env == "off":
        return True, uc.NOTICES_ENV
    state = read_state() if state is None else state
    if env == "on":
        if state.get("notices_off") is True:
            _record(notices_off=False)
        return False, uc.NOTICES_ENV
    return state.get("notices_off") is True, "setting"


_CLAIM_LOCK = threading.Lock()
_CLAIMED_AT: datetime | None = None


def _recently(now: datetime, state: dict) -> bool:
    last = _stamp(state.get("last_attempt"), now)
    return last is not None and now - last < NOTICE_INTERVAL


def claim(now: datetime | None = None) -> bool:
    """Take this week's notice, before anything is sent. False when this
    process or another one already took it."""
    global _CLAIMED_AT
    now = now or _utcnow()
    with _CLAIM_LOCK:
        if _CLAIMED_AT is not None and now - _CLAIMED_AT < NOTICE_INTERVAL:
            return False
        with _state_lock() as held:
            if held is False:
                return False
            state = read_state()
            if _recently(now, state):
                return False
            _CLAIMED_AT = now
            state["last_attempt"] = now.isoformat()
            uc.atomic_write_json(notices_path(), state)
        return True


def due() -> tuple[str, str] | None:
    """(current, new) when a notice is due now, else None."""
    if not community():
        return None
    pair = uc.versions()
    if pair is None:
        return None
    state = read_state()
    if notices_off(state)[0]:
        return None
    if _recently(_utcnow(), state) or (
            _CLAIMED_AT is not None and _utcnow() - _CLAIMED_AT < NOTICE_INTERVAL):
        return None
    installed = state.get("installed_version")
    with contextlib.suppress(InvalidVersion, TypeError):
        if installed and Version(str(installed)) >= Version(pair[1]):
            return None  # installed already, waiting for a restart
    job = _JOB
    if job is not None and job.new == pair[1]:
        return None
    return pair


# ------------------------------------------------------------- the dialog


def _session_facts(ctx) -> tuple[bool, str]:
    """(form dialog allowed, client name) for one request context."""
    try:
        from mcp.shared.version import LATEST_PROTOCOL_VERSION, SUPPORTED_PROTOCOL_VERSIONS

        params = ctx.session.client_params
        name = str(getattr(getattr(params, "clientInfo", None), "name", "") or "")
        if ctx.is_background_task:
            return False, name
        caps = getattr(params.capabilities, "elicitation", None)
        if caps is None:
            return False, name
        if getattr(caps, "form", None) is None and getattr(caps, "url", None) is not None:
            return False, name  # URL-only: no form
        asked = str(params.protocolVersion)
        negotiated = asked if asked in SUPPORTED_PROTOCOL_VERSIONS else LATEST_PROTOCOL_VERSION
        # Fail closed: the client's own version must be one of the two too,
        # so a client asking for an unknown era gets the `update` key.
        return asked == str(negotiated) and asked in DIALOG_PROTOCOLS, name
    except Exception:
        return False, ""


def _context():
    try:
        from fastmcp.server.dependencies import get_context

        return get_context()
    except Exception:
        return None


def reload_line(client_name: str) -> str:
    name = (client_name or "").lower()
    for needles, line in RELOAD_LINES:
        if any(n in name for n in needles):
            return line.format(product=PRODUCT)
    return RELOAD_DEFAULT


#: A decline or cancel faster than this is a client answering by itself
#: (the Desktop Code tab and local Cowork decline in milliseconds), not the
#: user: the dialog is then treated as unavailable.
HUMAN_ANSWER_S = 1.5

#: time.monotonic, as a seam for the tests' fake clock.
_monotonic = time.monotonic

#: Set when a dialog proved unusable in this process: (reason). The notice
#: then goes straight to the `update` key for the rest of the process.
_DIALOG_UNUSABLE: str | None = None


async def ask(ctx, current: str, new: str) -> tuple[str | None, str]:
    """Show the Yes/No dialog. Returns (choice, outcome):

    - ("Update now" | "Not now", "answered"): an accept with a valid choice;
    - (None, "answered"): a decline or cancel at least HUMAN_ANSWER_S after
      the request, which is the user's answer (ask again in a week);
    - (None, "unavailable:<reason>"): a decline or cancel faster than that
      (reason "auto-decline"), a JSON-RPC error such as -32601 or any other
      exception ("error"), or no answer within DIALOG_TIMEOUT_S ("timeout").

    Cancellation of the tool call propagates. On timeout the request is
    cancelled toward the client, so a late answer is dropped and the session
    goes on."""
    import anyio

    message = DIALOG_MESSAGE.format(product=PRODUCT, new=new, current=current)
    session = ctx.session
    request_id = getattr(session, "_request_id", None)  # the id elicit is about to use
    started = _monotonic()
    try:
        with anyio.move_on_after(DIALOG_TIMEOUT_S) as scope:
            result = await ctx.elicit(message, response_type=NoticeChoice)
        if scope.cancelled_caught:
            await _cancel_request(session, request_id)
            return None, "unavailable:timeout"
        action = getattr(result, "action", None)
        if action in ("decline", "cancel"):
            if _monotonic() - started < HUMAN_ANSWER_S:
                return None, "unavailable:auto-decline"
            return None, "answered"
        data = getattr(result, "data", None)
        if action == "accept" and isinstance(data, NoticeChoice)                 and data.choice in (UPDATE_NOW, NOT_NOW):
            return data.choice, "answered"
        return None, "unavailable:error"
    except Exception:
        return None, "unavailable:error"


def _dialog_failed(outcome: str) -> None:
    global _DIALOG_UNUSABLE
    if outcome.startswith("unavailable:"):
        _DIALOG_UNUSABLE = outcome.split(":", 1)[1]


async def _cancel_request(session, request_id) -> None:
    if not isinstance(request_id, int):
        return
    with contextlib.suppress(Exception):
        from mcp import types

        await session.send_notification(types.ServerNotification(types.CancelledNotification(
            params=types.CancelledNotificationParams(requestId=request_id,
                                                     reason="no answer in time"))))


# --------------------------------------------------------- the surface hooks

#: The consent update_now runs under, set by before_call for every surface
#: call: "yes" or "no" (the dialog's answer), "none" (no dialog possible:
#: the explicit call is the request), "unasked" (a surface call that was not
#: recognised as update_now: refused). Unset means an in-process call.
_CONSENT: contextvars.ContextVar[str | None] = contextvars.ContextVar(
    f"{PACKAGE}_update_consent", default=None)
#: The (current, new) pair before_call decided on; update_now acts on it.
_ASKED: contextvars.ContextVar[tuple[str, str] | None] = contextvars.ContextVar(
    f"{PACKAGE}_update_asked", default=None)


async def before_call(op: str | None):
    """Run before the op. op is the canonical op name the dispatcher will
    route (None when the arguments do not parse). Returns a token for
    end_call."""
    consent = "unasked"
    pair = None
    try:
        if op == "update_now" and community():
            consent = "none"
            import anyio
            # Refresh first (in a worker thread: up to one socket timeout), so the
            # dialog asks about the version update_now will act on even when the
            # cache did not know it yet. update_now's own refresh then finds it fresh.
            pair = await anyio.to_thread.run_sync(_newer_for_update_now, True)
            if pair is not None and install_route() in ("uvx", "mcpb"):
                ctx = _context()
                allowed, _ = _session_facts(ctx) if ctx is not None else (False, "")
                if allowed and _DIALOG_UNUSABLE in (None, "timeout"):
                    choice, outcome = await ask(ctx, *pair)
                    _dialog_failed(outcome)
                    if choice == UPDATE_NOW:
                        consent = "yes"
                    elif outcome in ("unavailable:auto-decline", "unavailable:error"):
                        consent = "none"   # no dialog here: the call is the request
                    else:
                        consent = "no"     # Not now, a real decline, or no answer
    except asyncio.CancelledError:
        raise
    except Exception:
        consent = "no"
        _say_once("hook", f"{PRODUCT}: the update notice step failed and was skipped.")
    return _CONSENT.set(consent), _ASKED.set(pair)


def end_call(token) -> None:
    with contextlib.suppress(Exception):
        _CONSENT.reset(token[0])
        _ASKED.reset(token[1])


def _fits(payload: dict, value: Any, cap: int) -> bool:
    try:
        size = len(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
        size += len(json.dumps(value, ensure_ascii=False, separators=(",", ":"))) + 12
        return size <= cap
    except Exception:
        return False


async def after_call(payload: Any, is_error: bool, *, read_only: bool, cap: int,
                     can_update: bool) -> Any:
    """Run after the op, before the result is returned. read_only: the tool
    is a read-only tool or help. can_update: update_now is on this server's
    tool list. Returns the payload, with an `update` key at most once a
    week. Never raises (cancellation propagates)."""
    try:
        if not community():
            return payload
        uc.maybe_refresh_in_background()
        pair = uc.versions()
        if pair is not None:
            _say_once(f"available-{pair[1]}",
                      f"{PRODUCT} {pair[1]} is available (installed {pair[0]}).")
        if (is_error or not read_only or not isinstance(payload, dict)
                or payload.get("ok") is False or "update" in payload):
            return payload
        told = _untold_outcome()
        if told is not None:
            if _fits(payload, told, cap):
                _mark_told()
                return {**payload, "update": told}
            return payload
        pair = due()
        if pair is None:
            return payload
        current, new = pair
        ctx = _context()
        allowed, client = _session_facts(ctx) if ctx is not None else (False, "")
        if can_update:
            fallback = {"notice": FALLBACK_NOTICE.format(product=PRODUCT, new=new,
                                                         current=current),
                        "relay": RELAY_OFFER.format(tool=uc.UPDATE_TOOL)}
        else:
            fallback = {"notice": FALLBACK_NOTICE.format(product=PRODUCT, new=new,
                                                         current=current)
                        + " To update, " + MANUAL_HINT.format(package=PACKAGE),
                        "relay": RELAY_ONCE}
        # The fallback must fit before anything is recorded, because an
        # unusable dialog falls back to it on this same result.
        if not _fits(payload, fallback, cap):
            return payload
        if (allowed and can_update and _DIALOG_UNUSABLE is None
                and install_route() in ("uvx", "mcpb")):
            if not claim():
                return payload
            choice, outcome = await ask(ctx, current, new)
            _record(last_dialog=outcome)
            _dialog_failed(outcome)
            if outcome != "answered":
                return {**payload, "update": fallback}
            if choice != UPDATE_NOW:
                return payload
            job = start(new, client)
            value = job.value()
            if _fits(payload, value, cap):
                if job.done.is_set():
                    _mark_told()
                return {**payload, "update": value}
            return payload
        if not claim():
            return payload
        return {**payload, "update": fallback}
    except asyncio.CancelledError:
        raise
    except Exception:
        _say_once("hook", f"{PRODUCT}: the update notice step failed and was skipped.")
        return payload


# ----------------------------------------------------------- install routes


def uv_binary() -> str | None:
    """The uv that launched this server: the UV variable uv sets for its
    children, only when it is an absolute path to an existing file named
    uv or uv.exe. Never a PATH lookup."""
    raw = os.environ.get("UV") or ""
    if not raw:
        return None
    path = Path(raw)
    if not path.is_absolute() or path.name.lower() not in ("uv", "uv.exe"):
        return None
    return str(path) if path.is_file() else None


def install_route() -> str:
    """mcpb, uvx, uv-tool or pip. Anything unsure is pip: the manual command."""
    if (os.environ.get(uc.INSTALL_ROUTE_ENV) or "").strip().lower() == "mcpb":
        return "mcpb"
    prefix = Path(sys.prefix)
    if (prefix / "uv-receipt.toml").exists():
        return "uv-tool"
    if uv_binary() and "archive-v0" in [part.lower() for part in prefix.parts]:
        return "uvx"
    return "pip"


def _scrub(text: str, limit: int = 200) -> str:
    """One line of a child's error output, with the home folder hidden."""
    line = (text or "").strip().splitlines()[-1] if (text or "").strip() else ""
    home = os.path.expanduser("~")
    if home and home != "~":
        line = line.replace(home, "~")
    return line[:limit]


def printed_version(out: str) -> Version | None:
    """The version `{PACKAGE} --version` printed: the last token of the last
    non-empty line, when packaging parses it."""
    lines = [ln.strip() for ln in (out or "").splitlines() if ln.strip()]
    if not lines:
        return None
    try:
        return Version(lines[-1].split()[-1])
    except (InvalidVersion, IndexError):
        return None


class _Job:
    """The one background update of this process."""

    def __init__(self, new: str, route: str, client: str) -> None:
        self.new, self.route, self.client = new, route, client
        self.done = threading.Event()
        self.status = "updating"
        self.message = UPDATING.format(product=PRODUCT, new=new)
        self.told = False
        self.proc: subprocess.Popen | None = None

    def finish(self, status: str, message: str) -> None:
        self.status, self.message = status, message
        if status == "updated":
            _record(installed_version=self.new)
        self.done.set()

    def value(self) -> dict:
        status = "updating in the background" if not self.done.is_set() else self.status
        return {"status": status, "notice": self.message, "relay": RELAY_ONCE}


_JOB: _Job | None = None
_JOB_LOCK = threading.Lock()


def start(new: str, client: str = "") -> _Job:
    """Start (or join) the single background update for version new."""
    global _JOB
    with _JOB_LOCK:
        job = _JOB
        if job is not None and (not job.done.is_set() or
                                (job.new == new and job.status in ("updated", "opened"))):
            return job
        route = install_route()
        job = _Job(new, route, client)
        _JOB = job
    if route == "pip":
        job.finish("command", PIP_COMMAND.format(product=PRODUCT, new=new, package=PACKAGE))
    elif route == "uv-tool":
        job.finish("command", UV_TOOL_COMMAND.format(product=PRODUCT, new=new, package=PACKAGE))
    else:
        target = _run_uvx if route == "uvx" else _run_bundle
        threading.Thread(target=_guarded, args=(target, job), daemon=True,
                         name=f"{PACKAGE}-update").start()
    return job


def _guarded(target, job: _Job) -> None:
    try:
        target(job)
    except Exception as exc:  # noqa: BLE001 (every failure is an outcome)
        if not job.done.is_set():
            if job.route == "uvx":
                job.finish("failed", UVX_FAILED.format(
                    product=PRODUCT, new=job.new, package=PACKAGE,
                    reason=f"{type(exc).__name__}"))
            else:
                job.finish("failed", BUNDLE_FAILED.format(
                    product=PRODUCT, new=job.new, package=PACKAGE, repo=uc.GITHUB_REPO,
                    reason=type(exc).__name__))
    finally:
        job.done.set()


def _uvx_failed(job: _Job, reason: str) -> None:
    job.finish("failed", UVX_FAILED.format(product=PRODUCT, new=job.new, package=PACKAGE,
                                           reason=reason))


def _spawn(argv: list[str], **kwargs) -> subprocess.Popen:
    """The one place a child process starts (tests replace it)."""
    return subprocess.Popen(argv, stdin=kwargs.pop("stdin"), stdout=kwargs.pop("stdout"),
                            stderr=kwargs.pop("stderr"), **kwargs)


def _run_uvx(job: _Job) -> None:
    uv = uv_binary()
    if uv is None:
        return _uvx_failed(job, "uv could not be found")
    argv = [uv, "tool", "run", "--refresh-package", PACKAGE,
            "--refresh-package", uc.CORE_PACKAGE, PACKAGE, "--version"]
    proc = _spawn(argv, shell=False, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=dict(os.environ),
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    job.proc = proc
    try:
        out, err = proc.communicate(timeout=INSTALL_TIMEOUT_S)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.communicate()
        return _uvx_failed(job, f"it took longer than {int(INSTALL_TIMEOUT_S)} seconds")
    finally:
        job.proc = None
    text = out.decode("utf-8", "replace") if isinstance(out, bytes) else str(out or "")
    if proc.returncode != 0:
        detail = _scrub(err.decode("utf-8", "replace") if isinstance(err, bytes) else str(err))
        return _uvx_failed(job, f"uv exited with code {proc.returncode}"
                           + (f" ({detail})" if detail else ""))
    got = printed_version(text)
    if got is None or got < Version(job.new):
        return _uvx_failed(job, f"uv started version {got}" if got else
                           "uv did not report the new version")
    job.finish("updated", UVX_DONE.format(product=PRODUCT, new=job.new, package=PACKAGE,
                                          reload=reload_line(job.client)))


class _GitHubOnly(urllib.request.HTTPRedirectHandler):
    """Redirects over https to the GitHub release hosts only."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        parts = urllib.parse.urlsplit(newurl)
        if parts.scheme != "https" or (parts.hostname or "").lower() not in GITHUB_HOSTS:
            raise urllib.error.HTTPError(newurl, code, "redirect outside GitHub refused",
                                         headers, fp)
        return super().redirect_request(req, fp, code, msg, headers, newurl)



def bundle_url(new: str) -> str:
    """Built from constants and the validated version only. No override."""
    return (f"https://github.com/{uc.GITHUB_REPO}/releases/download/"
            f"v{new}/{PACKAGE}-{new}.mcpb")


def _download(url: str, limit: int) -> bytes:
    opener = urllib.request.build_opener(_GitHubOnly)
    request = urllib.request.Request(url, headers={"User-Agent": f"{PACKAGE}/{uc.current_version()}"})
    with opener.open(request, timeout=DOWNLOAD_TIMEOUT_S) as resp:
        final = urllib.parse.urlsplit(resp.geturl())
        if final.scheme != "https" or (final.hostname or "").lower() not in GITHUB_HOSTS:
            raise ValueError("download left GitHub")
        data = resp.read(limit + 1)
    if len(data) > limit:
        raise ValueError("download too large")
    return data


class BundleRefused(ValueError):
    """The bundle failed a check; the message says which, in plain words."""


def expected_sha256(text: str, name: str) -> str:
    """The hash in a .sha256 file: '<64 hex>' or '<64 hex>  <name>'."""
    fields = (text or "").strip().split()
    if not fields or len(fields) > 2:
        raise BundleRefused("the published checksum file is not in the expected form")
    digest = fields[0].lower()
    if len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
        raise BundleRefused("the published checksum file is not in the expected form")
    if len(fields) == 2 and fields[1].lstrip("*") != name:
        raise BundleRefused("the published checksum names another file")
    return digest


def check_bundle(data: bytes, new: str) -> dict:
    """The in-memory bundle check (also run by scripts/bundle_release.py on
    every bundle a release attaches). Returns the manifest or raises
    BundleRefused. Nothing is extracted to disk."""
    if len(data) >= MAX_BUNDLE_BYTES:
        raise BundleRefused("the bundle is larger than 1 MB")
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
        infos = archive.infolist()
    except Exception:
        raise BundleRefused("the bundle is not a readable archive") from None
    if sorted(i.filename for i in infos) != ["icon.png", "manifest.json"]:
        raise BundleRefused("the bundle holds files other than manifest.json and icon.png")
    if sum(i.file_size for i in infos) >= MAX_BUNDLE_BYTES:
        raise BundleRefused("the bundle unpacks to more than 1 MB")
    try:
        manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
    except Exception:
        raise BundleRefused("the bundle manifest is not readable") from None
    if not isinstance(manifest, dict) or manifest.get("name") != PACKAGE:
        raise BundleRefused("the bundle is for another package")
    if manifest.get("version") != new:
        raise BundleRefused("the bundle is for another version")
    server = manifest.get("server")
    config = server.get("mcp_config") if isinstance(server, dict) else None
    if not isinstance(config, dict) or config.get("command") != "uvx" \
            or config.get("args") != [f"{PACKAGE}{getattr(uc, 'BUNDLE_EXTRAS', '')}=={new}"]:
        raise BundleRefused("the bundle starts something other than the published package")
    env = config.get("env", {})
    if not isinstance(env, dict) or any(BUNDLE_ENV.get(k) != v for k, v in env.items()):
        raise BundleRefused("the bundle sets settings it should not")
    return manifest


def _open_bundle(path: Path) -> None:
    """Hand the file to the OS default handler (ShellExecute, no shell
    string): Claude Desktop shows its own install window."""
    os.startfile(str(path))  # type: ignore[attr-defined]


def _run_bundle(job: _Job) -> None:
    def failed(reason: str) -> None:
        job.finish("failed", BUNDLE_FAILED.format(product=PRODUCT, new=job.new, package=PACKAGE,
                                                  repo=uc.GITHUB_REPO, reason=reason))

    if sys.platform != "win32":
        return failed("opening it automatically works on Windows only")
    url = bundle_url(job.new)
    name = url.rsplit("/", 1)[-1]
    try:
        digest = expected_sha256(_download(url + ".sha256", 4096).decode("ascii", "replace"),
                                 name)
        data = _download(url, MAX_BUNDLE_BYTES)
    except BundleRefused as exc:
        return failed(str(exc))
    except Exception:
        return failed("it could not be downloaded from the GitHub release")
    if hashlib.sha256(data).hexdigest() != digest:
        return failed("its checksum does not match the published one")
    try:
        check_bundle(data, job.new)
    except BundleRefused as exc:
        return failed(str(exc))
    target = uc.cache_path().parent / BUNDLE_SAVE_NAME
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(target.name + f".{os.getpid()}.tmp")
        tmp.write_bytes(data)
        os.replace(tmp, target)
        _open_bundle(target)
    except Exception:
        return failed("it could not be saved and opened")
    job.finish("opened", BUNDLE_OPENED.format(product=PRODUCT, new=job.new))


def shutdown() -> None:
    """Kill a running update child (server exit; also registered atexit)."""
    job = _JOB
    proc = job.proc if job is not None else None
    if proc is not None:
        with contextlib.suppress(Exception):
            proc.kill()


atexit.register(shutdown)


def _untold_outcome() -> dict | None:
    job = _JOB
    if job is None or job.told or not job.done.is_set():
        return None
    return job.value()


def _mark_told() -> None:
    job = _JOB
    if job is not None and job.done.is_set():
        job.told = True


# ---------------------------------------------------------------------- ops


def _require_community(op: str) -> None:
    if not community():
        raise UpdateRefused("NOT_AVAILABLE", f"{op} is not part of this edition; your "
                            "administrator installs updates. Nothing was changed.")


def _newer_for_update_now(refresh: bool) -> tuple[str, str] | None:
    if uc.disabled():
        return None
    if refresh:
        if uc.refresh_running():
            thread = uc._REFRESH_THREAD
            if thread is not None:
                thread.join(UPDATE_NOW_WAIT_S)
        elif uc.is_due(uc.read_cache()):
            uc.run_check(timeout=uc.BACKGROUND_TIMEOUT_SECONDS)
    return uc.versions()


def update_now() -> dict:
    """The update_now op (blocking; it runs in the call's worker thread)."""
    _require_community("update_now")
    if uc.disabled():
        raise UpdateRefused("NOT_FOUND", f"the update check is off ({uc.OFF_ENV}=off), so "
                            "no newer version is known. Nothing was installed.")
    consent = _CONSENT.get()
    if consent == "unasked":
        raise UpdateRefused("BAD_PARAMS", "update_now was not recognised; call it by its exact "
                            "name with no arguments. Nothing was installed.")
    # A surface call acts on the version before_call asked about (or found no
    # dialog possible for); only an in-process call refreshes here.
    pair = _newer_for_update_now(refresh=True) if consent is None else _ASKED.get()
    if pair is None:
        raise UpdateRefused("NOT_FOUND", f"no newer stable version of {PRODUCT} is known; this "
                            "build is the newest the last check found. Nothing was installed.")
    current, new = pair
    if consent == "no":
        return {"status": "not_now", "installed_version": current, "new_version": new,
                "message": NOT_NOW_TEXT}
    ctx = _context()
    client = _session_facts(ctx)[1] if ctx is not None else ""
    job = start(new, client)
    job.done.wait(UPDATE_NOW_WAIT_S)
    if job.done.is_set():
        job.told = True
        status, message = job.status, job.message
    else:
        status, message = "updating", STILL_UPDATING.format(product=PRODUCT, new=new)
    return {"status": status, "route": job.route, "installed_version": current,
            "new_version": new, "message": message}


NOTICES_OFF_TEXT = (f"{PRODUCT} update notices are off. The weekly check still runs; say "
                    f"turn on {PRODUCT} update notices to have them back.")
NOTICES_ON_TEXT = f"{PRODUCT} update notices are on."
ENV_WINS_TEXT = "The {env}={value} setting in this server's configuration decides instead."
NOT_SAVED_TEXT = "the setting could not be saved in the server's state folder; nothing changed"


def update_notices(action: str) -> dict:
    """The update_notices op: on, off or status."""
    _require_community("update_notices")
    if action not in ("on", "off", "status"):
        raise UpdateRefused("BAD_PARAMS", "action takes on, off or status")
    if action in ("on", "off"):
        if not _record(notices_off=(action == "off")):
            raise UpdateRefused("CONFLICT", NOT_SAVED_TEXT)
    state = read_state()
    off, decided_by = notices_off(state)
    if uc.disabled():
        off, decided_by = True, uc.OFF_ENV
    out: dict = {"notices": "off" if off else "on", "decided_by": decided_by,
                 "check": "off" if uc.disabled() else "on",
                 "current_version": uc.current_version()}
    pair = uc.versions()
    out["newer_version"] = pair[1] if pair else None
    last = _stamp(state.get("last_attempt"), _utcnow())
    out["last_notice"] = last.isoformat() if last else None
    out["next_notice_after"] = (last + NOTICE_INTERVAL).isoformat() if last and not off else None
    job = _JOB
    if job is not None:
        out["update"] = {"new_version": job.new, **job.value()}
    if action != "status":
        message = NOTICES_OFF_TEXT if action == "off" else NOTICES_ON_TEXT
        env = notices_env()
        if env is not None and env != action:
            message += " " + ENV_WINS_TEXT.format(env=uc.NOTICES_ENV, value=env)
        out["message"] = message
    return out
