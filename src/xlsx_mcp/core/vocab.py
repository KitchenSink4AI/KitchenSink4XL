"""Value vocabularies that Excel spells in camelCase (the rule is the Word
editor's: "values are snake_case everywhere; the old camelCase spellings are
still accepted").

Each table maps the snake_case value a caller sends (and a list reports) to
the spelling the file format stores. Matching ignores case and separators, so
greaterThan, greater_than, GREATER-THAN and "greater than" are one value; a
value that does not fold onto an entry is refused with the canonical list
(strict on meaning).
"""

from __future__ import annotations

import re

from .errors import XlMcpError


def fold(value: str) -> str:
    return re.sub(r"[\s_\-]+", "", value).lower()


#: Comparison operators (conditional-format cell_is rules, data validation).
OPERATORS = {
    "between": "between", "not_between": "notBetween", "equal": "equal",
    "not_equal": "notEqual", "greater_than": "greaterThan", "less_than": "lessThan",
    "greater_than_or_equal": "greaterThanOrEqual", "less_than_or_equal": "lessThanOrEqual",
}

#: Data-validation types.
DV_TYPES = {"list": "list", "whole": "whole", "decimal": "decimal", "date": "date",
            "time": "time", "text_length": "textLength", "custom": "custom"}

#: Conditional-format time periods.
TIME_PERIODS = {
    "today": "today", "yesterday": "yesterday", "tomorrow": "tomorrow",
    "last_7_days": "last7Days", "this_week": "thisWeek", "last_week": "lastWeek",
    "next_week": "nextWeek", "this_month": "thisMonth", "last_month": "lastMonth",
    "next_month": "nextMonth",
}

#: Workbook calculation modes.
CALC_MODES = {"auto": "auto", "auto_no_table": "autoNoTable", "manual": "manual"}


def to_file(value, table: dict[str, str], what: str) -> str:
    """The file-format spelling of `value`, refusing what is not in table."""
    if isinstance(value, str):
        hit = {fold(k): k for k in table}.get(fold(value))
        if hit is not None:
            return table[hit]
    raise XlMcpError(f"{what} must be one of {', '.join(table)}, got {value!r}")


def canonical(value, table: dict[str, str]) -> str | None:
    """The snake_case value for a file spelling or any accepted spelling."""
    if not isinstance(value, str):
        return None
    for snake, xml in table.items():
        if fold(value) in (fold(snake), fold(xml)):
            return snake
    return None


def from_file(value, table: dict[str, str]):
    """How a list reports a stored value: snake_case when known, else as stored."""
    return canonical(value, table) or value


__all__ = ["fold", "OPERATORS", "DV_TYPES", "TIME_PERIODS", "CALC_MODES", "to_file",
           "canonical", "from_file"]
