"""The fixed tool list, computed once at start from REGISTRY and the launch
settings (the same design as the PDF server's surface/plan.py).

Rules:
- one tool per group plus xl_help; every group tool takes {op, args};
- `op` (an enum of that tool's visible ops) and `args` each carry a one-line
  schema description;
- each tool has a title; readOnlyHint only when every op is read-only by
  core/readonly.py, destructiveHint when any op is DESTRUCTIVE by
  core/tool_annotations.py; idempotentHint = read-only; openWorldHint only
  where an op reaches a remote service (OPEN_WORLD_OPS);
- per-op lines are never truncated, one per line; a description holds at
  most 2,048 characters, so the lines that do not fit move, whole and in
  order, into the `op` parameter's description; every input schema stays at
  most 5,000 bytes;
- read-only ops live only in read-only tools and writing ops only in writing
  tools (a violation stops the build of the plan, and the test suite);
- internal ops (leading underscore) are never listed;
- the list never changes during a session: no packs, no enable_tools, no
  list_changed. KS4XL_READ_ONLY and KS4XL_FORBID are applied here, once.

Excel differences from PDF, each deliberate (all shared with the Word
editor, which this module follows):
- op order within a tool is the starting-map order
  (surface/groups.py), not alphabetical;
- each op line carries the op's signature, `name(required; opt optional)`,
  in the format generated from the args model (not English), with
  `backup` and `verify_com` left to xl_help and `path` implied, exactly as
  Word leaves `live`/`backup` and `file_path`.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from .. import registry
from ..registry import OpSpec
from ..settings import ENV_FORBID, Settings, SettingsError
from . import copy
from .groups import GROUPS, HELP_TOOL, READ_ONLY_GROUPS, TOOL_NAMES

ALL_TOOLS = TOOL_NAMES

#: Arguments left out of an op line's signature (xl_help shows them).
QUIET_ARGS = frozenset({"backup", "verify_com"})

#: Head additions for particular tools.
HEAD_ADDITIONS = {
    "xl_app": lambda: copy.XL_APP_HIDDEN_EXCEL,
}


#: Ops that reach a remote service (update_now asks PyPI and GitHub), so the
#: tool holding them says openWorldHint true. Every other tool says false.
OPEN_WORLD_OPS = frozenset({"update_now"})


class PlanError(RuntimeError):
    """The registry breaks a surface rule. A build defect, never a customer state."""


@dataclass
class ToolPlan:
    name: str
    ops: list[str]
    read_only: bool
    destructive: bool
    title: str
    description: str
    input_schema: dict[str, Any]
    annotations: dict[str, Any] = field(default_factory=dict)


@dataclass
class SurfacePlan:
    settings: Settings
    tools: dict[str, ToolPlan]
    op_tool: dict[str, str]           # every public op -> its tool (listed or not)
    forbidden_ops: frozenset[str]
    forbidden_tools: frozenset[str]
    hidden_by_read_only: frozenset[str]

    def visible_ops(self) -> dict[str, list[str]]:
        return {t.name: list(t.ops) for t in self.tools.values() if t.name != HELP_TOOL}

    def is_visible(self, op: str) -> bool:
        tool = self.op_tool.get(op)
        return tool in self.tools and op in self.tools[tool].ops

    def op_count(self) -> int:
        return sum(len(ops) for ops in self.visible_ops().values())


MAX_DESCRIPTION_CHARS = 2048
MAX_SCHEMA_BYTES = 5000


def schema_bytes(schema: dict[str, Any]) -> int:
    import json
    return len(json.dumps(schema, separators=(",", ":"), ensure_ascii=False).encode("utf-8"))


def signature(spec: OpSpec) -> str:
    """`name(required; opt optional)` from the args model."""
    fields = spec.args.model_fields
    req = [n for n, f in fields.items() if f.is_required() and n not in QUIET_ARGS
           and n != "path"]
    opt = [n for n, f in fields.items() if not f.is_required() and n not in QUIET_ARGS]
    parts = ", ".join(req)
    if opt:
        parts += ("; " if parts else "") + "opt " + " ".join(opt)
    return f"{spec.name}({parts})"


def op_text(spec: OpSpec) -> str:
    """The op's one-line meaning: an OP_LINES override, else its docstring
    first sentence, plus any fact suffix."""
    text = copy.OP_LINES.get(spec.name) or spec.summary
    suffix = copy.OP_LINE_SUFFIX.get(spec.name)
    return f"{text} {suffix}" if suffix else text


def _op_line(spec: OpSpec) -> str:
    return f"- {signature(spec)}: {op_text(spec)}"


def _head(group: str) -> str:
    head = copy.TOOLS[group]["description"]
    extra = HEAD_ADDITIONS.get(group)
    return f"{head} {extra()}" if extra else head


def _layout(head: str, lines: list[str]) -> tuple[str, str | None]:
    """(tool description, op parameter description or None). The head and
    as many op lines as fit stay in the description, which then says the list
    continues; the rest go, whole and in order, in the op parameter's
    description. Line wording is never changed or cut."""
    full = "\n".join([head, copy.OPS_HEADING, *lines])
    if len(full) <= MAX_DESCRIPTION_CHARS:
        return full, None
    kept: list[str] = []
    for ln in lines:
        trial = "\n".join([head, copy.OPS_HEADING, *kept, ln, copy.OPS_CONTINUED])
        if len(trial) > MAX_DESCRIPTION_CHARS:
            break
        kept.append(ln)
    rest = lines[len(kept):]
    desc = "\n".join([head, copy.OPS_HEADING, *kept, copy.OPS_CONTINUED])
    return desc, "\n".join([copy.OP_PARAM_CONTINUED, *rest])


def _group_schema(ops: list[str], op_description: str | None = None) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "op": {"type": "string", "enum": ops, "description": op_description or copy.OP_PARAM},
            "args": {"type": "object", "description": copy.ARGS_PARAM,
                     "additionalProperties": True},
        },
        "required": ["op"],
        "additionalProperties": False,
    }


def _help_schema() -> dict[str, Any]:
    # No enum: the group tools already enumerate every op.
    return {
        "type": "object",
        "properties": {
            "op": {"type": "string", "description": copy.HELP_OP_PARAM},
            "topic": {"type": "string", "enum": ["location"], "description": copy.HELP_TOPIC_PARAM},
        },
        "additionalProperties": False,
    }


def check_registry(ops: dict[str, OpSpec]) -> None:
    """Static surface rules over the whole catalog."""
    for name, spec in ops.items():
        if spec.group is None:
            raise PlanError(f"op {name} has no tool group (surface/groups.py)")
        if spec.group not in GROUPS:
            raise PlanError(f"op {name} is in unknown group {spec.group}")
        if spec.group in READ_ONLY_GROUPS and not spec.read_only:
            raise PlanError(f"op {name} can change something but is in the read-only tool "
                            f"{spec.group}")
        if spec.group not in READ_ONLY_GROUPS and spec.read_only:
            raise PlanError(f"op {name} is read-only but sits in {spec.group}; read-only ops "
                            "belong in xl_read so every tool's annotations are honest")
    for group, members in GROUPS.items():
        for op in members:
            if op not in ops:
                raise PlanError(f"surface/groups.py names {op} in {group}, which is not a "
                                "registered op")
    for group in ALL_TOOLS:
        if group not in copy.TOOLS:
            raise PlanError(f"no copy slot for tool {group}")


def build(settings: Settings, ops: dict[str, OpSpec] | None = None) -> SurfacePlan:
    ops = registry.public_ops() if ops is None else {n: s for n, s in ops.items()
                                                     if not s.internal}
    check_registry(ops)
    from ..core import update_flow
    if not update_flow.community():
        # Business: the update ops are not on its tool list (its tools stay
        # byte-identical to the release before them).
        ops = {n: s for n, s in ops.items() if n not in update_flow.OPS}
    known = set(ALL_TOOLS) | set(ops)
    unknown = sorted(set(settings.forbid) - known)
    if unknown:
        raise SettingsError(f"{ENV_FORBID} names {', '.join(unknown)}, which is not a tool or op "
                            "of this server")
    forbidden_tools = frozenset(n for n in settings.forbid if n in ALL_TOOLS)
    forbidden_ops = frozenset(n for n in settings.forbid if n in ops)
    hidden_ro: set[str] = set()

    op_tool = {name: spec.group for name, spec in ops.items()}
    tools: dict[str, ToolPlan] = {}
    for group, order in GROUPS.items():
        read_only = group in READ_ONLY_GROUPS
        members = [n for n in order if n in ops]
        members += sorted(n for n, s in ops.items() if s.group == group and n not in order)
        if settings.read_only and not read_only:
            hidden_ro.update(members)
            continue
        if group in forbidden_tools:
            continue
        visible = [n for n in members if n not in forbidden_ops]
        if not visible:
            continue
        destructive = any(ops[n].destructive for n in visible)
        slot = copy.TOOLS[group]
        description, op_description = _layout(_head(group), [_op_line(ops[n]) for n in visible])
        annotations = {"title": slot["title"], "readOnlyHint": read_only,
                       "destructiveHint": destructive, "idempotentHint": read_only,
                       "openWorldHint": any(n in OPEN_WORLD_OPS for n in visible)}
        tools[group] = ToolPlan(group, visible, read_only, destructive, slot["title"],
                                description, _group_schema(visible, op_description), annotations)

    if HELP_TOOL not in forbidden_tools and tools:
        listed = [op for t in tools.values() for op in t.ops]
        slot = copy.TOOLS[HELP_TOOL]
        tools[HELP_TOOL] = ToolPlan(
            HELP_TOOL, listed, True, False, slot["title"], slot["description"],
            _help_schema(),
            {"title": slot["title"], "readOnlyHint": True, "destructiveHint": False,
             "idempotentHint": True, "openWorldHint": False})
    return SurfacePlan(settings, tools, op_tool, forbidden_ops, forbidden_tools,
                       frozenset(hidden_ro))
