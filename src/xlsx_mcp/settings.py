"""Launch settings, read ONCE when the server starts and never changed after.

This is how IT and power users narrow the product now that packs are gone
(the same design as the PDF server's settings.py). Environment
variables in the launch configuration; no tool, argument or elicitation
changes any of them, and a value this module cannot parse stops the server
at start with one readable line rather than being guessed at.

    KS4XL_READ_ONLY          1/true/yes/on or 0/false/no/off (default off).
                            On: only read-only tools are listed.
    KS4XL_FORBID             comma list of tool or op names removed from the
                            tool list and from help; calls to them refuse
                            NOT_AVAILABLE. An unknown name refuses to start.
    KS4XL_MAX_OUTPUT_CHARS   per-call output ceiling in characters
                            (core/budget.py; default 80,000).
    KS4XL_CALL_TIMEOUT_S     wall-clock bound per call, in seconds (default 300).
    KS4XL_MAX_CALL_TIMEOUT_S ceiling for ops that declare a longer bound of
                            their own (registry.LONG_OPS); default 900.
    KS4XL_AUTHOR             who a note is signed by when a call names no
                            author (core/authorship.py): user (default),
                            the user's own name; assistant,
                            "AI Assistant via KS4XL".

KS4XL_ALLOWED_ROOTS keeps its meaning and its reader (core/sandbox.py).

RETIRED: KS4XL_MODE and KS4XL_PACK_POLICY belonged to the 1.x
pack system. An upgraded server meets them in configurations written by
earlier versions, so the old DEFAULTS are accepted and ignored (MODE lite,
full or everything; PACK_POLICY auto; blank is the same as unset) and
named in one stderr line at start and in get_server_info. A value that
limited the tools in 1.x (PACK_POLICY=locked, a MODE list of tool groups,
anything not understood) stops the server at start: starting the full
surface anyway would fail OPEN, a lock that silently stopped locking. All
retired settings are classified together, so one refusal stops the start
even when another is a harmless old default. Refusal text: copy slot
LEGACY_SETTING_REFUSAL; notice text: copy slot LEGACY_SETTING_NOTICE.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Mapping

ENV_READ_ONLY = "KS4XL_READ_ONLY"
ENV_FORBID = "KS4XL_FORBID"
ENV_MAX_OUTPUT = "KS4XL_MAX_OUTPUT_CHARS"
ENV_CALL_TIMEOUT = "KS4XL_CALL_TIMEOUT_S"
ENV_MAX_CALL_TIMEOUT = "KS4XL_MAX_CALL_TIMEOUT_S"
ENV_AUTHOR = "KS4XL_AUTHOR"

LEGACY_SETTINGS = ("KS4XL_MODE", "KS4XL_PACK_POLICY")

#: The values of each retired setting that 1.x servers wrote as a default and
#: that limited nothing by themselves; matched trimmed and case-insensitive.
#: Anything else (a lock, a list of tool groups, an unknown word) is refused.
_LEGACY_ACCEPTED = {
    "KS4XL_MODE": {"lite", "full", "everything"},
    "KS4XL_PACK_POLICY": {"auto"},
}

DEFAULT_CALL_TIMEOUT_S = 300.0
MIN_CALL_TIMEOUT_S = 1.0
MAX_CALL_TIMEOUT_S = 3600.0
DEFAULT_MAX_OP_TIMEOUT_S = 900.0

_TRUE = {"1", "true", "yes", "on"}
_FALSE = {"", "0", "false", "no", "off"}


class SettingsError(ValueError):
    """A launch setting could not be parsed. The server does not start."""


@dataclass(frozen=True)
class Settings:
    read_only: bool = False
    forbid: tuple[str, ...] = ()
    max_output_chars: int = 80_000
    call_timeout_s: float = DEFAULT_CALL_TIMEOUT_S
    max_call_timeout_s: float = DEFAULT_MAX_OP_TIMEOUT_S
    author_mode: str = "user"
    legacy_set: tuple[str, ...] = ()

    def timeout_for(self, op_timeout_s: float | None, requested_s: float | None = None) -> float:
        """The wall-clock bound of one call: the server default, raised to the
        op's own longer bound (capped at max_call_timeout_s), and never below
        a COM timeout the caller passed (timeout_seconds) plus a margin for the
        private hidden Excel to start."""
        bound = self.call_timeout_s
        if op_timeout_s is not None:
            bound = max(bound, min(float(op_timeout_s), self.max_call_timeout_s))
        if requested_s is not None and requested_s > 0:
            bound = max(bound, min(float(requested_s) + 60.0, self.max_call_timeout_s))
        return bound

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "Settings":
        env = os.environ if env is None else env
        raw_ro = (env.get(ENV_READ_ONLY) or "").strip().lower()
        if raw_ro in _TRUE:
            read_only = True
        elif raw_ro in _FALSE:
            read_only = False
        else:
            raise SettingsError(
                f"{ENV_READ_ONLY}={env.get(ENV_READ_ONLY)!r} is not understood; use 1 or 0")
        forbid = tuple(sorted({n.strip() for n in (env.get(ENV_FORBID) or "").split(",")
                               if n.strip()}))
        from .core.budget import parse_max_chars
        try:
            max_chars = parse_max_chars(env.get(ENV_MAX_OUTPUT))
        except ValueError as exc:
            raise SettingsError(str(exc)) from None
        timeout = _seconds(env, ENV_CALL_TIMEOUT, DEFAULT_CALL_TIMEOUT_S)
        max_timeout = _seconds(env, ENV_MAX_CALL_TIMEOUT, DEFAULT_MAX_OP_TIMEOUT_S)
        author_mode = (env.get(ENV_AUTHOR) or "").strip().lower() or "user"
        if author_mode not in ("user", "assistant"):
            raise SettingsError(
                f"{ENV_AUTHOR}={env.get(ENV_AUTHOR)!r} is not understood; use user (the "
                "default, your own name) or assistant")
        legacy = _classify_legacy(env)
        return cls(read_only=read_only, forbid=forbid, max_output_chars=max_chars,
                   call_timeout_s=timeout, max_call_timeout_s=max_timeout,
                   author_mode=author_mode, legacy_set=legacy)


def _classify_legacy(env: Mapping[str, str]) -> tuple[str, ...]:
    """The retired settings that are set to an old default, in order. A value
    that limited the tools in 1.x raises; every name is checked before any is
    accepted. Excel never had ALL_TOOLS or LOCK_TOOLS, so they are not read."""
    ignored: list[str] = []
    for name in LEGACY_SETTINGS:
        raw = (env.get(name) or "").strip().lower()
        if not raw:
            continue
        if raw not in _LEGACY_ACCEPTED[name]:
            from .surface import copy as _copy

            raise SettingsError(_copy.LEGACY_SETTING_REFUSAL.replace("{name}", name))
        ignored.append(name)
    return tuple(ignored)


def _seconds(env: Mapping[str, str], name: str, default: float) -> float:
    raw = (env.get(name) or "").strip()
    if not raw:
        return default
    try:
        value = float(raw)
    except ValueError:
        raise SettingsError(f"{name}={raw!r} is not a number of seconds") from None
    if not (MIN_CALL_TIMEOUT_S <= value <= MAX_CALL_TIMEOUT_S):
        raise SettingsError(f"{name}={raw} is outside {MIN_CALL_TIMEOUT_S:g} to "
                            f"{MAX_CALL_TIMEOUT_S:g} seconds")
    return value
