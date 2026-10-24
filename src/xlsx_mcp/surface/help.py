"""xl_help: generated from REGISTRY, never hand-written (the same design as
the Word family-surface help, itself taken from the PDF server's
surface/help.py).

With an op: the op's full argument schema (its generated model's JSON
schema), the one-line summary, the op's full shipped docstring (its detailed
meaning, the product's own words), the tool that runs it, whether it writes,
whether it edits in place, and an example call that validates against the
model (a test checks every op's example). With no op: every listed tool and
its ops. topic="location": the location-object explainer.

Examples are synthesized from the schema. Excel's path arguments are plain
strings, so a path is recognised by its name: path and friends get a file of
the kind the op works on (realistic example kinds).
"""

from __future__ import annotations

import copy as _copy
import inspect
from typing import Any

from .. import registry
from ..core.errors import SurfaceRefusal
from ..registry import OpSpec
from . import copy
from .groups import HELP_TOOL
from .plan import SurfacePlan, op_text


def args_schema(spec: OpSpec) -> dict[str, Any]:
    schema = _copy.deepcopy(spec.args.model_json_schema())
    schema.pop("examples", None)
    schema.pop("title", None)
    schema["additionalProperties"] = False
    return schema


#: A path argument gets an example of the file kind it takes.
_NAME_EXAMPLES: dict[str, Any] = {
    "path": "workbook.xlsx", "src": "workbook.xlsx", "dst": "workbook-copy.xlsx",
    "output": "workbook.pdf", "out_file": "range.csv", "out_dir": "output-folder",
    "dest_dir": "output-folder", "directory": "workbooks-folder",
    "image_file": "figure.png", "source_file": "rows.csv",
}

#: Ops whose `output` is not a PDF: the example names what they write.
_OP_OUTPUT_EXAMPLES = {"com_export_pdf": "workbook.pdf", "com_render_sheet": "sheet.png",
                       "com_convert_format": "workbook.xlsm"}


def _placeholder(prop: dict[str, Any], defs: dict[str, Any], name: str, op: str) -> Any:
    """A value that validates against `prop`, built recursively."""
    if "$ref" in prop:
        ref = defs.get(prop["$ref"].rsplit("/", 1)[-1], {})
        return _placeholder(ref, defs, name, op)
    if "enum" in prop:
        return prop["enum"][0]
    if "const" in prop:
        return prop["const"]
    for key in ("anyOf", "oneOf", "allOf"):
        if key in prop:
            options = [o for o in prop[key] if o.get("type") != "null"] or prop[key]
            return _placeholder(options[0], defs, name, op)
    kind = prop.get("type")
    if name in _NAME_EXAMPLES and kind in ("string", "array", None):
        value = _NAME_EXAMPLES[name]
        if name == "output" and op in _OP_OUTPUT_EXAMPLES:
            value = _OP_OUTPUT_EXAMPLES[op]
        if kind == "array" and isinstance(value, str):
            value = [value]
        return value
    if kind == "string":
        return "text"
    if kind == "integer":
        return max(0, int(prop.get("minimum", 0)))
    if kind == "number":
        return float(prop.get("minimum", 1))
    if kind == "boolean":
        return False
    if kind == "array":
        if not prop.get("minItems"):
            return []
        item = _placeholder(prop.get("items", {}), defs, name, op)
        return [item] * int(prop["minItems"])
    if kind == "object" or "properties" in prop:
        props = prop.get("properties", {})
        return {n: _placeholder(props[n], defs, n, op) for n in prop.get("required", [])}
    return None


def example(spec: OpSpec) -> dict[str, Any]:
    from .examples import EXAMPLES

    if spec.name in EXAMPLES:
        return dict(EXAMPLES[spec.name])
    extra = spec.args.model_config.get("json_schema_extra")
    if isinstance(extra, dict) and extra.get("examples"):
        return dict(extra["examples"][0])
    schema = spec.args.model_json_schema()
    defs = schema.get("$defs", {})
    props = schema.get("properties", {})
    return {n: _placeholder(props[n], defs, n, spec.name) for n in schema.get("required", [])}


def _detail(spec: OpSpec) -> str:
    return " ".join((inspect.getdoc(spec.handler) or "").split())


def help_payload(plan: SurfacePlan, op: str | None, topic: str | None = None) -> dict[str, Any]:
    if topic is not None:
        if topic != "location":
            raise SurfaceRefusal("BAD_PARAMS", "topic takes one value, 'location'",
                                 details={"topics": ["location"]})
        return {"topic": "location", "location": copy.LOCATION_HELP}
    if op is None or op == "":
        return {"usage": copy.HELP_USAGE, "tools": plan.visible_ops()}
    if not isinstance(op, str):
        raise SurfaceRefusal("BAD_PARAMS", "op must be an op name")
    specs = registry.public_ops()
    from . import normalize
    canonical = normalize.route_op(op, specs)  # the same folded routing
    if canonical is not None:
        op = canonical
    if op in plan.forbidden_ops or op in plan.hidden_by_read_only \
            or plan.op_tool.get(op) in plan.forbidden_tools:
        raise SurfaceRefusal("NOT_AVAILABLE", f"{op} is turned off by this server's launch "
                             "settings", details={"op": op, "decided_by": copy.SETTINGS_OWNER})
    if op not in specs or not plan.is_visible(op):
        from .dispatch import unknown_op  # the same suggestions as a call
        raise unknown_op(plan, op, None)
    spec = specs[op]
    tool = plan.op_tool[op]
    out = {"op": op, "tool": tool, "summary": op_text(spec),
           "writes": spec.writes, "edits_in_place": spec.destructive,
           "args_schema": args_schema(spec), "example": {"op": op, "args": example(spec)},
           "detail": _detail(spec)}
    if spec.fixed_args:
        out["fixed_args"] = dict(spec.fixed_args)
    if "location" in spec.args.model_fields:
        out["location"] = copy.LOCATION_HELP
    return out


__all__ = ["help_payload", "args_schema", "example", "HELP_TOOL"]
