"""Lenient on routing and shape, strict on meaning (the Word editor's rules,
carried to Excel).

Every call is normalized BEFORE validation, and every change is echoed in the
result as `normalized_from` ({what: what was sent}) so the model learns the
canonical form. Nothing here changes what an op does once it runs: each rule
maps a spelling or a shape onto the ONE canonical value it unambiguously
names, and anything that does not map exactly is left for strict validation
(or the op) to refuse as before.

-   op names: case-insensitive, separators ignored ("SetCell", "setCell",
    "set-cell" route to set_cell). The released 1.2.5 tool names ARE the op
    names (migration/v1_to_v2.json), so there is no rename table. EXACT on
    the folded key only; a test proves no two ops share one. Anything
    fuzzier is a suggestion, never a route.
-   flat arguments: an op's own argument names sent beside `op` instead of
    inside `args` are moved into `args`; `file_path` and `file` (Word's and
    PowerPoint's spelling) are accepted for Excel's `path`. Unknown keys are
    still refused.
-   did you mean: an unknown op refuses NOT_FOUND (nothing ran) with at most
    three suggestions, each with its tool and one-line purpose; a destructive
    op is never suggested for a name that does not itself carry that verb.
-   near-miss values and shapes: values of the arguments each op checks are
    matched without regard to case or separators, plus a short list of exact
    synonyms; {"search": "text"} for {"search": {"text": "text"}} and
    {"region": "B7"} for {"region": {"near": "B7"}}; common spellings of
    formatting keys (colour, fontSize, wrap).
-   values are snake_case everywhere (greater_than, last_7_days,
    text_length, auto_no_table); Excel's camelCase spellings still work and
    are echoed. "anchor" means only a get_grid_view token: manage_chart's
    cell argument is at_cell (anchor still accepted and echoed).
"""

from __future__ import annotations

import copy as _copy
import re
from typing import Any

from ..core import vocab as _vocab

#: Renamed arguments, old name -> new name, per op.
RENAMED_ARGS: dict[str, dict[str, str]] = {
    "manage_chart": {"anchor": "at_cell"},
}

#: Values whose canonical spelling became snake_case, per (op, argument
#: path); the old spellings fold onto the new ones, so they keep working.
VALUE_TABLES: dict[tuple[str, str], dict[str, str]] = {
    ("manage_conditional_format", "params.operator"): {
        **_vocab.OPERATORS, "contains_text": "containsText", "not_contains": "notContains",
        "begins_with": "beginsWith", "ends_with": "endsWith"},
    ("manage_conditional_format", "params.period"): _vocab.TIME_PERIODS,
    ("manage_data_validation", "dv_type"): _vocab.DV_TYPES,
    ("manage_data_validation", "operator"): _vocab.OPERATORS,
    ("set_workbook_properties", "calc_mode"): _vocab.CALC_MODES,
}

#: Synonyms for the file argument: Word and PowerPoint call it file_path.
FILE_SYNONYMS = ("file_path", "file")

#: Enum vocabularies of string arguments checked inside the ops:
#: (op, argument) -> the canonical values.
VOCAB: dict[tuple[str, str], tuple[str, ...]] = {
    ("read_range", "values"): ("cached", "formula", "both", "displayed"),
    ("get_cells", "values"): ("cached", "formula", "both", "displayed"),
    ("get_table", "values"): ("cached", "formula", "both", "displayed"),
    ("get_grid_view", "values"): ("cached", "formula", "both"),
    ("query_range", "values"): ("cached", "formula", "both", "displayed"),
    ("query_range", "match"): ("all", "any"),
    ("export_range", "values"): ("cached", "formula", "both", "displayed"),
    ("export_file", "values"): ("cached", "formula", "both", "displayed"),
    ("export_range", "fmt"): ("csv", "tsv", "json"),
    ("export_file", "fmt"): ("csv", "tsv", "json"),
    ("import_data", "fmt"): ("auto", "csv", "tsv", "json"),
    ("clear_range", "what"): ("contents", "formats", "all"),
    ("copy_range", "what"): ("all", "values", "formulas", "formats"),
    ("set_merge", "action"): ("merge", "unmerge", "list"),
    ("find_cells", "look_in"): ("values", "formulas", "both"),
    ("find_cells", "match"): ("exact", "contains", "regex"),
    ("replace_cells", "look_in"): ("values", "formulas", "both"),
    ("replace_cells", "match"): ("exact", "contains", "regex"),
    ("manage_worksheet", "action"): ("add", "delete", "rename", "copy", "reorder", "hide",
                                     "unhide"),
    ("manage_worksheet", "state"): ("hidden", "very_hidden"),
    ("manage_table", "action"): ("add_row", "delete_row", "add_column", "delete_column",
                                 "rename", "resize", "toggle_totals", "to_range", "set_style"),
    ("manage_name", "action"): ("add", "delete", "rename", "update", "list"),
    ("manage_conditional_format", "action"): ("add", "list", "delete"),
    ("manage_conditional_format", "cf_type"): (
        "cell_is", "color_scale", "data_bar", "icon_set", "formula", "top_bottom", "unique",
        "duplicate", "above_average", "time_period", "blanks", "errors"),
    ("manage_data_validation", "action"): ("add", "list", "delete"),
    ("manage_comment", "action"): ("add", "edit", "delete", "list"),
    ("manage_hyperlink", "action"): ("add", "remove", "list"),
    ("modify_grid_structure", "action"): ("insert_rows", "delete_rows", "insert_cols",
                                          "delete_cols"),
    ("manage_image", "action"): ("insert", "list", "delete", "extract"),
    ("manage_chart", "action"): ("create", "list", "delete"),
    ("manage_chart", "chart_type"): ("bar", "bar_horizontal", "line", "pie", "doughnut",
                                     "area", "scatter"),
    ("set_protection", "action"): ("sheet", "workbook", "unlock", "remove", "status"),
    ("set_protection", "scope"): ("sheet", "workbook", "all"),
    ("set_page_layout", "orientation"): ("portrait", "landscape"),
    ("set_header_footer", "apply_to"): ("odd", "even", "first"),
    ("manage_backups", "action"): ("list", "restore", "purge", "snapshot"),
    ("manage_backups", "source"): ("prev", "anchor"),
    ("manage_backups", "scope"): ("orphans", "slots"),
    # 'formulas' was removed in 2.0.0; it stays here so any spelling of it
    # reaches recalculate's own refusal, which says so and names the next step.
    ("recalculate", "engine"): ("auto", "com", "formulas"),
    ("com_manage_pivot", "action"): ("create", "refresh", "delete", "list"),
    ("com_export_pdf", "scope"): ("workbook", "sheet", "range"),
    ("com_convert_format", "format"): ("xlsx", "xlsm", "xlsb", "xls", "csv", "ods"),
    ("com_set_sparkline", "action"): ("create", "clear", "list"),
    ("com_set_sparkline", "type"): ("line", "column", "win_loss"),
    **{key: tuple(table) for key, table in VALUE_TABLES.items() if "." not in key[1]},
}

#: Exact synonyms (after case/separator folding) that name a canonical value.
SYNONYMS: dict[tuple[str, str], dict[str, str]] = {
    ("modify_grid_structure", "action"): {
        "insertcolumns": "insert_cols", "insertcolumn": "insert_cols",
        "insertcol": "insert_cols", "deletecolumns": "delete_cols",
        "deletecolumn": "delete_cols", "deletecol": "delete_cols",
        "insertrow": "insert_rows", "deleterow": "delete_rows"},
    ("manage_worksheet", "action"): {"show": "unhide", "move": "reorder"},
    ("manage_worksheet", "state"): {"veryhidden": "very_hidden"},
    ("set_merge", "action"): {"split": "unmerge"},
    ("manage_hyperlink", "action"): {"delete": "remove"},
    ("manage_chart", "chart_type"): {"column": "bar", "barhorizontal": "bar_horizontal"},
    ("set_workbook_properties", "calc_mode"): {"automatic": "auto"},
    ("com_set_sparkline", "type"): {"winloss": "win_loss"},
}

#: Keys of a formatting object (format_cells; apply_style define), matched
#: without regard to case or separators, plus exact synonyms.
FORMAT_KEYS = {
    "font": ("name", "size", "bold", "italic", "underline", "strike", "color"),
    "fill": ("color", "pattern", "fg", "bg"),
    "border": ("style", "color", "sides", "left", "right", "top", "bottom"),
    "alignment": ("horizontal", "vertical", "wrap_text", "text_rotation", "indent"),
}
FORMAT_SYNONYMS = {
    "font": {"colour": "color", "fontcolor": "color", "fontsize": "size", "sizept": "size",
             "fontname": "name", "fontfamily": "name", "family": "name",
             "strikethrough": "strike"},
    "fill": {"colour": "color", "fgcolor": "fg", "foreground": "fg", "bgcolor": "bg",
             "background": "bg", "filltype": "pattern", "patterntype": "pattern"},
    "border": {"colour": "color", "linestyle": "style"},
    "alignment": {"wrap": "wrap_text", "wraptext": "wrap_text", "rotation": "text_rotation",
                  "halign": "horizontal", "valign": "vertical"},
}

#: Location selectors whose value may be sent as a bare string.
_TEXT_SELECTORS = {"search": "text", "region": "near"}

#: Words that make an op destructive for the suggestion rule.
DESTRUCTIVE_VERBS = ("delete", "clear", "purge", "remove", "overwrite", "replace")


def fold(value: str) -> str:
    """Case and separators folded: 'Very_Hidden' -> 'veryhidden'."""
    return re.sub(r"[\s_\-]+", "", value).lower()


# ------------------------------------------------------------------ op names

def op_keys(ops) -> dict[str, str]:
    """folded key -> canonical op, for every op. Raises when two names would
    share a key (a test runs this over the registry)."""
    keys: dict[str, str] = {}
    for name in ops:
        k = fold(name)
        if k in keys and keys[k] != name:
            raise ValueError(f"op names {keys[k]} and {name} fold to the same key {k}")
        keys[k] = name
    return keys


def route_op(op: str, ops) -> str | None:
    """The canonical op for `op`, by EXACT folded key; None when there is none."""
    if op in ops:
        return op
    return op_keys(ops).get(fold(op))


def _distance(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def suggestions(op: str, ops, limit: int = 3) -> list[str]:
    """Up to `limit` op names closest to `op`, never a destructive op
    for a name without that verb."""
    q = fold(op)
    scored = []
    for name in sorted(set(ops)):
        f = fold(name)
        if any(v in f for v in DESTRUCTIVE_VERBS) and not any(
                v in f and v in q for v in DESTRUCTIVE_VERBS):
            continue
        d = _distance(q, f)
        if q and (q in f or f in q):
            d = min(d, abs(len(f) - len(q)) // 2)
        scored.append((d, name))
    scored.sort()
    cut = max(3, len(q) // 2)
    return [n for d, n in scored if d <= cut][:limit]


# ------------------------------------------------------------- args and shapes

def flatten(arguments: dict[str, Any], fields) -> tuple[dict[str, Any], list[str]]:
    """Canonical argument names (and the file synonyms) sent beside `op`
    move into args. Returns (arguments, moved names). Anything else stays
    where it is, to be refused."""
    extra = [k for k in arguments if k not in ("op", "args")]
    movable = [k for k in extra if k in fields or k in FILE_SYNONYMS]
    if not movable:
        return arguments, []
    args = dict(arguments.get("args") or {}) if isinstance(arguments.get("args"), dict) else None
    if args is None and arguments.get("args") is not None:
        return arguments, []
    args = args or {}
    out = {k: v for k, v in arguments.items() if k not in movable}
    for k in movable:
        if k in args:
            return arguments, []  # the same name twice: leave it to be refused
        args[k] = arguments[k]
    out["args"] = args
    return out, movable


def _match(value: str, vocab, synonyms: dict[str, str] | None = None) -> str | None:
    hit = {fold(v): v for v in vocab}.get(fold(value))
    if hit is None and synonyms:
        hit = synonyms.get(fold(value))
    return hit


def _canon_value(value: str, table: dict[str, str]) -> str | None:
    """The snake_case value an accepted spelling names."""
    return _vocab.canonical(value, table)


def normalize_args(op: str, args: dict[str, Any], fields) -> tuple[dict[str, Any], dict]:
    """Value and shape rules for one op. Returns (args, {what: sent})."""
    sent: dict[str, Any] = {}
    out = dict(args)
    # The file synonyms.
    if "path" in fields and "path" not in out:
        for syn in FILE_SYNONYMS:
            if syn in out and syn not in fields:
                sent[syn] = out[syn]
                out["path"] = out.pop(syn)
                break
    # Renamed arguments.
    for old, new in RENAMED_ARGS.get(op, {}).items():
        if old in out and new not in out and new in fields:
            sent[old] = out[old]
            out[new] = out.pop(old)
    # Values of checked arguments: case and separators, plus synonyms,
    # and the snake_case values.
    for name, value in list(out.items()):
        if not isinstance(value, str):
            continue
        table = VALUE_TABLES.get((op, name))
        if table is not None:
            canon = _canon_value(value, table) or (
                SYNONYMS.get((op, name), {}).get(fold(value)))
            if canon and canon != value:
                sent.setdefault(name, value)
                out[name] = canon
            continue
        vocab = VOCAB.get((op, name))
        if not vocab or value in vocab:
            continue
        hit = _match(value, vocab, SYNONYMS.get((op, name)))
        if hit:
            sent.setdefault(name, value)
            out[name] = hit
    # Nested values (conditional-format params).
    params = out.get("params")
    if isinstance(params, dict):
        new_params = dict(params)
        for (vop, path), table in VALUE_TABLES.items():
            if vop != op or not path.startswith("params."):
                continue
            key = path.split(".", 1)[1]
            value = params.get(key)
            if isinstance(value, str):
                canon = _canon_value(value, table)
                if canon and canon != value:
                    new_params[key] = canon
                    sent.setdefault(f"params.{key}", value)
        out["params"] = new_params
    # Formatting objects: key spelling (colour -> color, fontSize -> size).
    for obj in ("font", "fill", "border", "alignment"):
        if isinstance(out.get(obj), dict):
            fixed = _fix_format_keys(obj, out[obj])
            if fixed != out[obj]:
                sent.setdefault(obj, _copy.deepcopy(out[obj]))
                out[obj] = fixed
    if isinstance(out.get("define"), dict):
        define = dict(out["define"])
        changed = False
        for obj in ("font", "fill", "border", "alignment"):
            if isinstance(define.get(obj), dict):
                fixed = _fix_format_keys(obj, define[obj])
                if fixed != define[obj]:
                    define[obj] = fixed
                    changed = True
        if changed:
            sent.setdefault("define", _copy.deepcopy(out["define"]))
            out["define"] = define
    # Location objects: {"search": "Revenue"}, {"region": "B7"}.
    before = _copy.deepcopy(out)
    out = _fix_locations(out)
    if out != before:
        for k in out:
            if out[k] != before.get(k):
                sent.setdefault(k, before.get(k))
    return out, sent


def _fix_format_keys(obj: str, fmt: dict) -> dict:
    keys = FORMAT_KEYS[obj]
    known = {fold(k): k for k in keys}
    out = {}
    for k, v in fmt.items():
        if k in keys:
            out[k] = v
            continue
        canon = known.get(fold(k)) or FORMAT_SYNONYMS[obj].get(fold(k))
        target = canon if canon and canon not in fmt else k
        if obj == "border" and target in ("left", "right", "top", "bottom") \
                and isinstance(v, dict):
            v = _fix_format_keys("border", v)
        out[target] = v
    return out


def _fix_locations(value: Any) -> Any:
    if isinstance(value, list):
        return [_fix_locations(v) for v in value]
    if not isinstance(value, dict):
        return value
    out = {k: _fix_locations(v) for k, v in value.items()}
    for sel, inner in _TEXT_SELECTORS.items():
        if isinstance(out.get(sel), str):
            out[sel] = {inner: out[sel]}
    return out


__all__ = ["fold", "route_op", "op_keys", "suggestions", "flatten", "normalize_args",
           "RENAMED_ARGS", "VALUE_TABLES", "VOCAB", "SYNONYMS", "FILE_SYNONYMS"]
