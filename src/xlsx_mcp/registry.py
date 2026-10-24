"""The op registry: every Excel operation, declared once (family surface).

The same design as the Word family-surface registry, which took it from the
PDF server's registry.

Before the family surface, each module-level function in server.py was
registered as its own MCP tool. Now each one is an OP: server.py's `_op`
decorator records an OpSpec here and leaves the function itself untouched,
so in-process callers (the test suite's `srv.<name>(...)` convention) keep
the raw function, its returns and its typed exceptions. The MCP surface
(xlsx_mcp.surface) builds the fixed tool list from this registry and
dispatches every call through it.

What an OpSpec carries, and where each fact comes from:

- name     the function's name, which is also the op name on the wire;
- group    the tool the op lives on (surface/groups.py);
- summary  the first sentence of the function's shipped docstring;
- args     a pydantic model generated from the function signature with
           extra="forbid", so an unknown key refuses on every op;
- writes / read_only   from core/readonly.py (the product's own
           classification; an unclassified op fails registration there);
- destructive          from core/tool_annotations.py;
- runs_live            the op drives the private hidden Excel (com_ ops and
           recalculate). Excel has no live editing of a user's open workbook.

Excel differs from Word here in one way: there is no derived op (Word's
split of get_document_view has no Excel counterpart).
"""

from __future__ import annotations

import inspect
import re
import typing
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any, Callable, Mapping

from pydantic import BaseModel, ConfigDict, ValidationError, create_model


class StrictArgs(BaseModel):
    """Base of every generated args model: unknown keys refuse."""

    model_config = ConfigDict(extra="forbid", arbitrary_types_allowed=True)


@dataclass(frozen=True)
class OpSpec:
    name: str
    group: str | None
    summary: str
    args: type[StrictArgs]
    handler: Callable[..., Any]
    writes: bool
    read_only: bool
    destructive: bool
    runs_live: bool
    timeout_s: float | None = None
    fixed_args: Mapping[str, Any] = field(default_factory=lambda: MappingProxyType({}))
    target: str = ""

    @property
    def internal(self) -> bool:
        return self.name.startswith("_")

    def call(self, kwargs: dict[str, Any]) -> Any:
        """Run the engine function with validated arguments plus any fixed ones."""
        return self.handler(**{**kwargs, **dict(self.fixed_args)})


REGISTRY: dict[str, OpSpec] = {}

#: Parameters that are never part of an op's arguments.
_SKIPPED_PARAMS = frozenset({"ctx", "context"})

#: Ops whose own bound is longer than the server default (seconds). The
#: dispatcher caps it at KS4XL_MAX_CALL_TIMEOUT_S. These start the private
#: hidden Excel and already carry COM timeouts of their own
#: (timeout_seconds), which the dispatcher also honours.
LONG_OPS: dict[str, float] = {
    "recalculate": 900.0, "com_manage_pivot": 900.0, "com_export_pdf": 900.0,
    "com_render_sheet": 900.0, "com_convert_format": 900.0,
    "com_save_with_password": 900.0, "com_autofit": 900.0,
    "com_goal_seek": 900.0, "com_set_sparkline": 900.0,
    "com_validate_opens_clean": 900.0,
}


def validation_problems(exc: ValidationError) -> list[dict[str, str]]:
    """The ONE way a schema mismatch is described: fields and problems,
    never the value given."""
    return [{"field": ".".join(str(p) for p in e["loc"]) or "(root)", "problem": e["msg"]}
            for e in exc.errors(include_url=False, include_input=False)]


_SENTENCE = re.compile(r"\.\s")


def first_sentence(fn: Callable) -> str:
    """The first sentence of the function's docstring, whitespace collapsed
    (cut at the first '. ')."""
    doc = " ".join((inspect.getdoc(fn) or "").split())
    m = _SENTENCE.search(doc + " ")
    return doc if m is None else doc[: m.start() + 1]


def _camel(name: str) -> str:
    return "".join(part.capitalize() for part in name.split("_")) + "Args"


def args_model(fn: Callable, name: str, drop: frozenset[str] = frozenset()) -> type[StrictArgs]:
    """A strict pydantic model from the function signature (defaults kept,
    Annotated metadata kept, `ctx` dropped)."""
    hints = typing.get_type_hints(fn, include_extras=True)
    fields: dict[str, Any] = {}
    for p in inspect.signature(fn).parameters.values():
        if p.name in drop or p.name in _SKIPPED_PARAMS:
            continue
        if p.kind in (p.VAR_POSITIONAL, p.VAR_KEYWORD):
            raise TypeError(f"op {name!r}: *args/**kwargs cannot become a strict schema")
        ann = hints.get(p.name, Any)
        default = ... if p.default is p.empty else p.default
        fields[p.name] = (ann, default)
    return create_model(_camel(name), __base__=StrictArgs, **fields)


def register(fn: Callable, *, group_of: Callable[[str], str | None]) -> OpSpec:
    """Record one engine function as an op."""
    from .core import readonly, tool_annotations

    name = fn.__name__
    if name in REGISTRY:
        raise ValueError(f"op {name!r} registered twice")
    read_only = readonly.read_only_hint(name)  # raises on an unclassified op
    runs_live = name.startswith("com_") or name == "recalculate"
    spec = OpSpec(
        name=name, group=group_of(name), summary=first_sentence(fn),
        args=args_model(fn, name), handler=fn, writes=not read_only,
        read_only=read_only,
        destructive=(not read_only) and bool(tool_annotations.destructive_hint(name)),
        runs_live=runs_live, timeout_s=LONG_OPS.get(name), target=name)
    REGISTRY[name] = spec
    return spec


def public_ops() -> dict[str, OpSpec]:
    """Every op a caller may reach (internal ops excluded)."""
    return {n: s for n, s in REGISTRY.items() if not s.internal}


def unregister(name: str) -> None:
    """Test support: drop a test-only op."""
    REGISTRY.pop(name, None)
