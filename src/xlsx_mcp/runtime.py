"""The frozen state of a running server: its launch settings and tool plan.

Set once by server.create_server(); read by get_server_info and xl_help.
In-process callers that never started a server (the test suite's
srv.<name>(...) convention, scripts) get a default plan built on first use,
so get_server_info reports the same surface a default launch would.
"""

from __future__ import annotations

import contextvars
from contextlib import contextmanager
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from .surface.plan import SurfacePlan

_plan: "SurfacePlan | None" = None


def set_plan(plan: "SurfacePlan") -> None:
    global _plan
    _plan = plan


#: The plan of the server running the current call. A process may build
#: more than one server (the test suite does); each tool carries its own
#: plan and sets it here for the call, so get_server_info reports the
#: server that was actually called.
_CALL_PLAN: contextvars.ContextVar["SurfacePlan | None"] = contextvars.ContextVar(
    "ks4xl_call_plan", default=None)


@contextmanager
def calling(p: "SurfacePlan") -> Iterator["SurfacePlan"]:
    token = _CALL_PLAN.set(p)
    try:
        yield p
    finally:
        _CALL_PLAN.reset(token)


def plan() -> "SurfacePlan":
    global _plan
    current = _CALL_PLAN.get()
    if current is not None:
        return current
    if _plan is None:
        from .settings import Settings
        from .surface import plan as surface_plan
        return surface_plan.build(Settings())
    return _plan


def excel_forbidden() -> str | None:
    """Why this server may not start Excel, or None. KS4XL_FORBID naming
    xl_app (or com_validate_opens_clean)
    is an administrator's "no Excel here": it closes every route that would
    start the private hidden Excel, verify_com and KS4XL_VERIFY_COM
    included, not only the xl_app tool."""
    if _CALL_PLAN.get() is None and _plan is None:
        return None
    p = plan()
    if "xl_app" in p.forbidden_tools or "com_validate_opens_clean" in p.forbidden_ops:
        named = "xl_app" if "xl_app" in p.forbidden_tools else "com_validate_opens_clean"
        return (f"this server's launch settings turn Excel off (KS4XL_FORBID names {named}), "
                "so nothing here starts Excel")
    return None


def started() -> bool:
    return _plan is not None


def reset() -> None:
    """Test support: forget the started plan."""
    global _plan
    _plan = None
