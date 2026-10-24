"""The MCP tools: one FastMCP Tool per entry of the surface plan (the same
design as the PDF server's surface/tools.py).

This is the ONLY module that constructs MCP tools (an AST test pins
it). Each tool's inputSchema is exactly the plan's {op, args} schema (no
generated wrapper schema, no outputSchema) and its run() hands the raw
arguments to the dispatcher, so FastMCP's function-signature validation is
not involved at all.
"""

from __future__ import annotations

from typing import Any

from fastmcp.tools.tool import Tool, ToolResult
from mcp.types import ToolAnnotations
from pydantic import PrivateAttr

from .. import envelope, runtime
from ..core import budget, update_flow
from ..core.errors import SurfaceRefusal
from .dispatch import dispatch, shape
from .groups import HELP_TOOL
from .help import help_payload
from .plan import ToolPlan

_HELP_KEYS = frozenset({"op", "topic"})


async def run_help(plan, arguments: Any) -> tuple[dict[str, Any], bool]:
    cap = plan.settings.max_output_chars
    try:
        if not isinstance(arguments, dict) or set(arguments) - _HELP_KEYS:
            raise SurfaceRefusal("BAD_PARAMS", "xl_help takes two optional arguments, op and "
                                 "topic")
        result = help_payload(plan, arguments.get("op"), arguments.get("topic"))
        result = budget.enforce({"ok": True, **result}, cap)
        return result, False
    except Exception as exc:  # noqa: BLE001
        return budget.cap_refusal(envelope.refusal(exc), cap), True


class XlTool(Tool):
    """A fixed surface tool, bound to the plan of the server that built it."""

    _plan: Any = PrivateAttr(default=None)

    async def run(self, arguments: dict[str, Any]) -> ToolResult:
        plan = self._plan or runtime.plan()
        with runtime.calling(plan):
            if self.name == HELP_TOOL:
                payload, is_error = await run_help(plan, arguments)
            else:
                # The update hook (core/update_flow.py), step 1: update_now's
                # own Yes/No dialog, before the op runs. It sees the op name
                # exactly as the dispatcher routes it.
                token = await update_flow.before_call(_routed_op(plan, arguments))
                try:
                    payload, is_error = await dispatch(plan, self.name, arguments)
                finally:
                    update_flow.end_call(token)
            # Step 2: the weekly background check and, after a read-only call
            # only, the update notice (a dialog, or one `update` key).
            tp = plan.tools.get(self.name)
            payload = await update_flow.after_call(
                payload, is_error, read_only=self.name == HELP_TOOL or bool(tp and tp.read_only),
                cap=plan.settings.max_output_chars, can_update=plan.is_visible("update_now"))
        return envelope.tool_result(payload, is_error=is_error)


def _routed_op(plan, arguments: Any) -> str | None:
    """The canonical op a call routes to, or None when it does not parse."""
    try:
        return shape(plan, arguments)[0]
    except Exception:  # noqa: BLE001 (dispatch reports the refusal itself)
        return None


def make_tool(tp: ToolPlan, plan=None) -> XlTool:
    ann = tp.annotations
    tool = XlTool(
        name=tp.name,
        title=tp.title,
        description=tp.description,
        parameters=tp.input_schema,
        output_schema=None,
        annotations=ToolAnnotations(title=ann["title"], readOnlyHint=ann["readOnlyHint"],
                                    destructiveHint=ann["destructiveHint"],
                                    idempotentHint=ann["idempotentHint"],
                                    openWorldHint=ann["openWorldHint"]),
    )
    tool._plan = plan
    return tool
