"""Output budget: no answer leaves the server larger than it is allowed to be
(the same design as the PDF and PowerPoint servers).

`enforce()` runs in the dispatcher on EVERY op's result, so an op cannot opt
out. The contract, in both directions:

1. NEVER EXCEEDED. The serialized answer, measured with the same compact
   serializer the wire uses (envelope.dumps), stays at or under the ceiling.
   If nothing in an answer can be cut honestly, the call refuses with
   LIMIT_EXCEEDED instead of exceeding.
2. NOTHING SILENTLY DROPPED. A cut answer carries a `page` block naming the
   field that was cut, the true total, how many came back, how many were
   omitted and, when the op takes an `offset`, the offset that continues it.
3. SMALL ANSWERS UNCHANGED. An answer that fits comes back byte-identical.

The ceiling is KS4XL_MAX_OUTPUT_CHARS, default 80,000 characters (about
20,000 tokens at four characters per token), which keeps every answer under
Claude Code's own 25,000-token MCP output cap, so the client never cuts an
answer silently. PDF's 52,000 default was not taken: an Excel read of a
wide range (read_range, get_grid_view, query_range, export_range with no
out_file) is routinely larger, and the Word editor chose 80,000 for the same
reason; one family ceiling keeps the siblings' answers comparable. An
unparseable value refuses at start.

Excel has no untrusted-content envelope, so the PDF tag-aware
cut is dropped: a string is cut at a character boundary and says so.
"""

from __future__ import annotations

import copy
from typing import Any

DEFAULT_MAX_CHARS = 80_000
MIN_MAX_CHARS = 2_000
ENV_MAX_CHARS = "KS4XL_MAX_OUTPUT_CHARS"

#: Characters reserved for the page block so a full answer can still explain
#: what it left out.
PAGE_BLOCK_RESERVE = 640

TRUNCATION_MARK = "\n[... {n} more characters not returned; output limit reached; continue at character {k}]"


def _limit_error(message: str, details: dict | None = None):
    from .errors import SurfaceRefusal
    return SurfaceRefusal("LIMIT_EXCEEDED", message, details=details)


def _bad(message: str):
    from .errors import SurfaceRefusal
    return SurfaceRefusal("BAD_PARAMS", message)


def truncate_text(text: str, allowance: int) -> tuple[str, int]:
    """Cut one string to fit a serialized allowance and say so in the string.
    Returns (cut_text, characters_omitted)."""
    if _size(text) <= allowance:
        return text, 0

    def build(keep: int) -> str:
        return text[:keep] + TRUNCATION_MARK.format(n=len(text) - keep, k=keep)

    mark_len = len(TRUNCATION_MARK.format(n=len(text), k=len(text))) + 40
    keep = max(0, allowance - mark_len - 2)
    out = build(keep)
    while keep > 0 and _size(out) > allowance:
        keep = int(keep * 0.9)
        out = build(keep)
    if _size(out) > allowance:
        keep = 0
        out = build(0)
    return out, len(text) - keep


def parse_max_chars(raw: str | None) -> int:
    """Validate the launch value. Raises ValueError with a readable reason."""
    if raw is None or not str(raw).strip():
        return DEFAULT_MAX_CHARS
    text = str(raw).strip()
    try:
        value = int(text)
    except ValueError:
        raise ValueError(
            f"{ENV_MAX_CHARS}={raw!r} is not a whole number of characters (at least "
            f"{MIN_MAX_CHARS}); unset it for the {DEFAULT_MAX_CHARS} default") from None
    if value < MIN_MAX_CHARS:
        raise ValueError(f"{ENV_MAX_CHARS}={value} is below the {MIN_MAX_CHARS}-character floor")
    return value


def _size(obj: Any) -> int:
    from ..envelope import dumps
    return len(dumps(obj))


def normalize_window(limit, offset, *, what: str = "items") -> tuple[int | None, int]:
    """Validate caller paging. Refuses loudly rather than clamping."""
    if offset is None:
        offset = 0
    if not isinstance(offset, int) or isinstance(offset, bool) or offset < 0:
        raise _bad(f"offset must be a whole number of {what} at or above 0, got {offset!r}")
    if limit is not None and (not isinstance(limit, int) or isinstance(limit, bool) or limit < 1):
        raise _bad(f"limit must be a whole number of {what} at or above 1, got {limit!r}")
    return limit, offset


def shrink_record(item: Any, allowance: int) -> tuple[Any, int]:
    """Cut a record's longest text field so the record fits."""
    if not isinstance(item, dict):
        if isinstance(item, str):
            return truncate_text(item, allowance)
        return item, 0
    longest, size = None, 0
    for key, value in item.items():
        if isinstance(value, str) and len(value) > size:
            longest, size = key, len(value)
    if longest is None:
        return item, 0
    fixed = _size({k: v for k, v in item.items() if k != longest})
    text, dropped = truncate_text(item[longest], allowance - fixed - len(longest) - 6)
    if not dropped:
        return item, 0
    new = dict(item)
    new[longest] = text
    return new, dropped


def page_items(items: list, *, limit=None, offset=0, budget: int, overhead: int = 0,
               unit: str = "items", shrink=shrink_record) -> tuple[list, dict | None]:
    """Cut `items` to what fits and say exactly what was left out. Returns
    (kept, page); page is None when the caller asked for no window and it all
    fit. At least one item always comes back (shrunk if it must be)."""
    limit, offset = normalize_window(limit, offset, what=unit)
    total = len(items)
    window = items[offset:] if offset else items
    if limit is not None:
        window = window[:limit]
    spent = overhead
    kept: list = []
    stopped = False
    shrunk = 0
    for item in window:
        cost = _size(item) + 1
        if spent + cost > budget:
            stopped = True
            if not kept and shrink is not None:
                item, shrunk = shrink(item, budget - spent - 1)
                if spent + _size(item) + 1 <= budget:
                    kept.append(item)
            break
        kept.append(item)
        spent += cost
    returned = len(kept)
    omitted = max(0, total - offset - returned)
    if not (offset or limit is not None) and not stopped and returned == total:
        return kept, None
    page = {"unit": unit, "total": total, "offset": offset, "returned": returned,
            "omitted": omitted, "truncated": omitted > 0 or bool(shrunk),
            "reason": "budget" if stopped else ("limit" if limit is not None else "offset"),
            "next_offset": (offset + returned) if omitted > 0 else None}
    if shrunk:
        page["chars_cut_from_last"] = shrunk
    return kept, page


# ----------------------------------------------------- totals lead the answer

#: The fields that say how much an answer holds and how to get the rest. A
#: client may show only the first ~2,000 characters of an answer, so they go
#: first (after ok and file), never after the long list they describe.
LEAD_KEYS = ("page", "truncated", "total", "total_rules",
             "extension_only_rules", "count", "row_count", "matched",
             "returned", "omitted", "offset", "next_offset", "scanned",
             "sheet_count", "rows", "cols", "dimensions", "shown")
#: A lead field longer than this stays where it is (a "rows" that is the
#: data itself, not its count).
_LEAD_MAX = 600


def lead_with_totals(result: dict[str, Any]) -> dict[str, Any]:
    """The same answer, with its totals and paging fields first. Key order
    is all that changes."""
    if not isinstance(result, dict):
        return result
    head: dict[str, Any] = {k: result[k] for k in ("ok", "file") if k in result}
    for key in LEAD_KEYS:
        if key in result and key not in head:
            value = result[key]
            if key == "page" or isinstance(value, (bool, int, float)) \
                    or value is None or _size(value) <= _LEAD_MAX:
                head[key] = value
    if len(head) <= 2 and all(k in ("ok", "file") for k in head):
        return result
    return {**head, **{k: v for k, v in result.items() if k not in head}}


# ------------------------------------------------------ the wrapper backstop


def _largest_cuttable(obj: dict, prefix: tuple = (), exclude: frozenset = frozenset()
                      ) -> tuple[tuple, int] | None:
    """Path to the largest list or string in a dict tree (dicts only are
    descended), with its serialized size. The page block is never a target."""
    best: tuple[tuple, int] | None = None
    for key, value in obj.items():
        if not prefix and key in ("page", "ok", "file"):
            continue
        path = prefix + (key,)
        if path in exclude:
            continue
        if isinstance(value, (list, str)):
            size = _size(value)
            if best is None or size > best[1]:
                best = (path, size)
        elif isinstance(value, dict):
            inner = _largest_cuttable(value, path, exclude)
            if inner is not None and (best is None or inner[1] > best[1]):
                best = inner
    return best


def _get(obj: dict, path: tuple):
    for key in path:
        obj = obj[key]
    return obj


def _set(obj: dict, path: tuple, value) -> None:
    for key in path[:-1]:
        obj = obj[key]
    obj[path[-1]] = value


def enforce(result: dict[str, Any], cap: int, *, supports_offset: bool = False,
            requested_offset: int = 0) -> dict[str, Any]:
    """Make `result` (the final success envelope) serialize within `cap`,
    cutting honestly. Cuts the largest list or string first, repeatedly,
    recording each cut in the result's `page` block (a list when more than one
    field was cut). Refuses with LIMIT_EXCEEDED when nothing cuttable is left.
    """
    if _size(result) <= cap:
        return result
    out = copy.deepcopy(result)
    op_page = out.get("page") if isinstance(out.get("page"), dict) else None
    base = requested_offset
    if op_page and isinstance(op_page.get("offset"), int):
        base = op_page["offset"]
    cuts: list[dict] = []
    tried: set[tuple] = set()
    while _size(out) > cap:
        found = _largest_cuttable(out, exclude=frozenset(tried))
        if found is None:
            raise _limit_error(
                "the answer is larger than this server's output limit and has no part that can "
                "be cut without misleading; narrow the request",
                details={"limit_chars": cap})
        path, _size_now = found
        tried.add(path)
        value = _get(out, path)
        _set(out, path, [] if isinstance(value, list) else "")
        out["page"] = {"cut": True}  # reserve room for the real block below
        overhead = _size(out) + PAGE_BLOCK_RESERVE
        field = ".".join(str(p) for p in path)
        if isinstance(value, list):
            is_op_list = bool(op_page) and len(path) == 1
            kept, pg = page_items(value, budget=cap, overhead=overhead, unit=field)
            total = op_page["total"] if is_op_list and isinstance(op_page.get("total"), int) \
                else base + len(value) if supports_offset else len(value)
            start = base if supports_offset or is_op_list else 0
            returned = len(kept)
            omitted = max(0, total - start - returned)
            cut = {"field": field, "total": total, "offset": start, "returned": returned,
                   "omitted": omitted, "reason": "output_limit",
                   "next_offset": (start + returned) if omitted and supports_offset else None,
                   "how_to_continue": "offset" if supports_offset else "narrow_request"}
            if pg and pg.get("chars_cut_from_last"):
                cut["chars_cut_from_last"] = pg["chars_cut_from_last"]
            _set(out, path, kept)
        else:
            text, dropped = truncate_text(value, max(0, cap - overhead))
            cut = {"field": field, "chars_total": len(value), "chars_returned": len(value) - dropped,
                   "chars_omitted": dropped, "continue_at_char": len(value) - dropped,
                   "reason": "output_limit", "how_to_continue": "narrow_request"}
            _set(out, path, text)
        cuts.append(cut)
        page: dict[str, Any] = dict(op_page) if op_page else {}
        if len(cuts) == 1:
            page.update(cuts[0])
        else:
            page["cuts"] = cuts
        page["truncated"] = True
        out["page"] = page
    return out


def cap_refusal(payload: dict[str, Any], cap: int) -> dict[str, Any]:
    """A refusal never exceeds the cap either: its details go first."""
    if _size(payload) <= cap:
        return payload
    out = copy.deepcopy(payload)
    err = out.get("error", {})
    for key in ("details", "detail", "matches"):
        if key in err:
            err[key] = {"omitted": f"{key} exceeded the output limit"}
    if _size(out) > cap:
        err["message"] = err.get("message", "")[: max(100, cap // 4)]
    return out
