"""Guard: every tool ships a title, and every tool that can change
something declares whether the change may be destructive.

The Anthropic Connectors Directory requires `title` and the applicable
hints on every tool, and a client shows the title to a human instead of
the raw tool name, so a missing or wrong one is a product defect rather
than a documentation nit. Three things are checked here.

FIRST, coverage and shape on the wire: every registered tool carries a
non-empty title, titles are unique within this server, none exceeds 40
characters, and none carries product or marketing language.

SECOND, the classification itself. The non-destructive allowlist below is
hand-audited and lives in the TEST on purpose. `core/tool_annotations.py`
is where the server reads its classification from; if both sides shared
one list the test would only prove the file equals itself. A new tool
consequently fails twice until somebody classifies it deliberately, once
at registration, where `annotations()` raises on an unclassified mutating
name, and once here.

THIRD, the published tables. `docs/TOOL_TITLES.md` and
`docs/TOOL_ANNOTATIONS.md` are what a reviewer reads and what a maintainer
disputes a row in. They are generated from the module, so they are checked
against it here; a table that drifts from the wire is worse than no table.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from xlsx_mcp import registry, server  # noqa: F401  (the import registers ops)
from xlsx_mcp.core import tool_annotations as ann
from xlsx_mcp.surface import copy, groups

DOCS = Path(__file__).resolve().parents[2] / "docs"

#: Words a tool title must not carry.
MARKETING = {
    "best", "powerful", "easy", "easiest", "fast", "fastest", "pro",
    "premium", "ultimate", "smart", "magic", "magical", "free", "awesome",
    "amazing", "seamless", "effortless", "revolutionary", "unlimited",
    "kitchensink4ai", "kitchensink4word", "kitchensink4xl",
    "kitchensink4ppt", "kitchensink4web",
}

# FAMILY SURFACE: the wire now carries seven tools whose titles are copy
# (surface/copy.py TOOLS); each op keeps its classification here, and the
# tool's annotations are computed from its ops (test_family_surface.py
# test_tool_list_shape_and_annotations). The per-op title table stays the
# module's record of each operation's name and is checked against the docs.


def _ops() -> dict:
    """op name -> OpSpec for every engine function."""
    return dict(registry.public_ops())


def _wire_titles() -> dict:
    return {name: copy.TOOLS[name]["title"] for name in groups.TOOL_NAMES}


def test_every_surface_tool_carries_a_non_empty_title():
    missing = [n for n, t in _wire_titles().items() if not t.strip()]
    assert not missing, missing


def test_titles_are_unique_within_this_server():
    titles = list(_wire_titles().values())
    assert len(titles) == len(set(titles))


def test_no_filled_title_is_longer_than_forty_characters():
    """An unfilled title is a COPY_SLOT placeholder, which the owner-test
    ship gate refuses to build; a filled one must fit."""
    long = {n: t for n, t in _wire_titles().items()
            if copy.SLOT not in t and len(t) > 40}
    assert not long, f"titles over 40 characters: {long}"


def test_no_title_carries_marketing_language():
    offenders = {}
    for name, title in _wire_titles().items():
        if copy.SLOT in title:
            continue
        hit = set(re.findall(r"[a-z0-9]+", title.lower())) & MARKETING
        if hit:
            offenders[name] = sorted(hit)
    assert not offenders, offenders


def test_the_title_table_covers_every_operation_exactly():
    ops = set(_ops())
    assert ops - set(ann.TITLES) == set(), sorted(ops - set(ann.TITLES))
    assert set(ann.TITLES) - ops == set(), sorted(set(ann.TITLES) - ops)


def test_an_untitled_name_raises_rather_than_getting_a_generated_title():
    with pytest.raises(RuntimeError):
        ann.title("a_tool_that_does_not_exist")


def _doc_rows(filename: str) -> list:
    text = (DOCS / filename).read_text(encoding="utf-8")
    rows = []
    for line in text.splitlines():
        if not line.startswith("| `"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append([c.strip("`") for c in cells])
    return rows


def test_the_published_title_table_matches_the_module():
    rows = _doc_rows("TOOL_TITLES.md")
    published = {name: title for name, title, *_ in rows}
    assert published == ann.TITLES, (
        "docs/TOOL_TITLES.md has drifted from tool_annotations.py. "
        "Regenerate it rather than editing either side by hand.")


# --------------------------------------------- destructiveHint and friends

#: HAND-AUDITED. Every name here was classified by reading the tool's
#: implementation, against the rule stated in docs/TOOL_ANNOTATIONS.md: a
#: tool is non-destructive only when every path either ADDS without
#: replacing, or changes no user data at all.
NON_DESTRUCTIVE_TOOLS = {
    "com_validate_opens_clean", "copy_workbook", "create_table",
    "create_workbook",
    # the update ops change the server program or its notice setting,
    # never a workbook (core/update_flow.py)
    "update_notices", "update_now",
}


def test_every_mutating_op_carries_a_destructive_classification():
    missing = [n for n, spec in sorted(_ops().items())
               if not spec.read_only and ann.destructive_hint(n) is None]
    assert not missing, missing


def test_the_destructive_classification_covers_the_surface():
    mutating = {n for n, spec in _ops().items() if not spec.read_only}
    classified = ann.DESTRUCTIVE | ann.NON_DESTRUCTIVE
    assert mutating - classified == set(), sorted(mutating - classified)
    assert classified - mutating == set(), sorted(classified - mutating)
    assert not (ann.DESTRUCTIVE & ann.NON_DESTRUCTIVE)


def test_the_non_destructive_set_matches_the_hand_audited_allowlist():
    assert ann.NON_DESTRUCTIVE == NON_DESTRUCTIVE_TOOLS, (
        "the non-destructive set changed. Added: "
        f"{sorted(ann.NON_DESTRUCTIVE - NON_DESTRUCTIVE_TOOLS)}; removed: "
        f"{sorted(NON_DESTRUCTIVE_TOOLS - ann.NON_DESTRUCTIVE)}.")


def test_the_op_destructive_flag_is_the_module_classification():
    for n, spec in _ops().items():
        if spec.read_only:
            assert spec.destructive is False, n
        else:
            assert spec.destructive is bool(ann.destructive_hint(n)), n


def test_an_unclassified_mutating_tool_raises_at_registration():
    with pytest.raises(RuntimeError):
        ann.annotations("a_tool_that_does_not_exist", read_only=False)


def test_the_published_annotation_tables_match_the_module():
    text = (DOCS / "TOOL_ANNOTATIONS.md").read_text(encoding="utf-8")
    destructive, non_destructive = set(), set()
    bucket = None
    for line in text.splitlines():
        if line.startswith("## Tools that may perform destructive"):
            bucket = destructive
        elif line.startswith("## Tools that may not"):
            bucket = non_destructive
        elif line.startswith("## "):
            bucket = None
        elif bucket is not None and line.startswith("| `"):
            bucket.add(line.strip("|").split("|")[0].strip().strip("`"))
    assert destructive == ann.DESTRUCTIVE, (
        "docs/TOOL_ANNOTATIONS.md destructive table has drifted: added "
        f"{sorted(destructive - ann.DESTRUCTIVE)}, removed "
        f"{sorted(ann.DESTRUCTIVE - destructive)}")
    assert non_destructive == ann.NON_DESTRUCTIVE, (
        "docs/TOOL_ANNOTATIONS.md non-destructive table has drifted: added "
        f"{sorted(non_destructive - ann.NON_DESTRUCTIVE)}, removed "
        f"{sorted(ann.NON_DESTRUCTIVE - non_destructive)}")


def test_every_destructive_row_states_a_reason():
    text = (DOCS / "TOOL_ANNOTATIONS.md").read_text(encoding="utf-8")
    thin = []
    for line in text.splitlines():
        if not line.startswith("| `"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if len(cells) == 2 and len(cells[1]) < 20:
            thin.append(cells[0])
    assert not thin, (
        f"a reviewer cannot dispute a row with no reason on it: {thin}")
