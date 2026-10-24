"""Guard: every tool declares whether it can change anything, and no
tool that can change something claims it cannot.

Claude Desktop groups tools by readOnlyHint and offers the read-only
group a single Always Allow. That makes a wrong hint a permission bug
rather than a documentation nit, so two things are checked here.

The allowlist below is hand-audited and lives in the TEST on purpose.
core/readonly.py is where the server reads its classification from; if
both sides shared one list the test would only prove the file equals
itself. A new tool consequently fails twice until someone classifies it
deliberately: once at import, where read_only_hint raises on an unknown
name, and once here.

The second guard catches a WRONG classification rather than a missing
one. It walks the call graph out of each read-only tool and fails if it
reaches a save, a file write, a temporary file, or the DispatchEx that
starts a hidden Excel. The walk treats a few infrastructure modules as
leaves, named and justified below, because they create lock directories
and read process tables without ever touching a workbook.
"""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

from xlsx_mcp import registry, server  # noqa: F401  (the import registers ops)
from xlsx_mcp.core import readonly

PACKAGE = Path(readonly.__file__).parent.parent


#: HAND-AUDITED. Every name here was classified by reading the tool's
#: implementation, not its name. The bar: the tool cannot change
#: anything at all, temporary files and Office processes included.
READ_ONLY_TOOLS = {
    "audit_formulas", "audit_styles", "com_status", "diagnose_workbook",
    "find_cells", "get_cells", "get_connections", "get_external_links",
    "get_grid_view", "get_pivot", "get_server_info", "get_table",
    "get_workbook_metadata", "get_workflows", "inspect_vba",
    "query_range", "read_range", "validate",
}

#: Tools that MUST fail the mutation walk. Without these the walk could
#: rot into a check that passes because it finds nothing anywhere, and a
#: silent guard is worse than no guard.
KNOWN_MUTATORS = {
    "com_export_pdf", "export_range", "manage_backups",
    "manage_comment", "recalculate", "set_cell", "write_range",
}

#: What counts as changing something, in source terms.
MUTATION_MARKERS = [
    r"\b_edit\(",
    r"\.save\(",
    r"write_lock",
    r"safesave",
    r"shutil\.(copy|copyfile|copy2|move|rmtree)",
    r"os\.(remove|unlink|replace|rename|makedirs|mkdir)",
    r"\.write_text\(",
    r"\.write_bytes\(",
    r"\.unlink\(",
    r"\.mkdir\(",
    r"""open\([^)]*['\"][wax]b?\+?['\"]""",
    r"\.SaveAs",
    r"DispatchEx",
    r"tempfile\.(mkdtemp|mkstemp|NamedTemporaryFile|TemporaryDirectory)",
]

#: Modules the walk stops at. Each is infrastructure a read can
#: legitimately reach: xproc.py and serial.py are the cross-process and in-process locks,
#: dialogs.py reads the OS window layer, update_check.py reads the cached
#: update line, and update_flow.py holds the update notice and the update the
#: user chooses: it runs from the surface hook and from update_now, never from
#: a read op's own path (its generic names, start and finish, would otherwise
#: become call-graph edges to unrelated functions of the same name).
LEAF_MODULES = {"xproc.py", "serial.py", "dialogs.py",
                "update_check.py", "update_flow.py"}
LEAF_FUNCTIONS = {"_screen_repair_async"}

_MARKER_RE = re.compile("|".join(MUTATION_MARKERS))


def _index() -> dict:
    """name -> [(relative path, ast node, file source)], module-level and
    class-level definitions only. Nested closures are deliberately not
    indexed: their text already sits inside the parent's source segment,
    so they are scanned, but their generic names (body, run) must not
    become call-graph edges to unrelated functions elsewhere."""
    out: dict = {}
    for path in sorted(PACKAGE.rglob("*.py")):
        if "__pycache__" in str(path):
            continue
        source = path.read_text(encoding="utf-8")
        try:
            tree = ast.parse(source)
        except SyntaxError:  # pragma: no cover
            continue
        nodes = list(tree.body)
        for cls in [n for n in tree.body if isinstance(n, ast.ClassDef)]:
            nodes.extend(cls.body)
        for node in nodes:
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            if path.name in LEAF_MODULES or node.name in LEAF_FUNCTIONS:
                continue
            out.setdefault(node.name, []).append(
                (path.relative_to(PACKAGE).as_posix(), node, source))
    return out


def _callees(node) -> set:
    names = set()
    for child in ast.walk(node):
        if isinstance(child, ast.Call):
            func = child.func
            if isinstance(func, ast.Name):
                names.add(func.id)
            elif isinstance(func, ast.Attribute):
                names.add(func.attr)
    return names


def _reaches_mutation(index, name, seen=None, depth=0, trail=""):
    if seen is None:
        seen = set()
    if name in seen or depth > 5:
        return None
    seen.add(name)
    for relpath, node, source in index.get(name, []):
        segment = ast.get_source_segment(source, node) or ""
        hit = _MARKER_RE.search(segment)
        if hit:
            line = segment[:hit.start()].count(chr(10)) + node.lineno
            return (f"{trail}{name} ({relpath}:{line}) reaches "
                    f"{hit.group(0)}")
        for callee in sorted(_callees(node)):
            if callee in index and callee not in seen:
                found = _reaches_mutation(
                    index, callee, seen, depth + 1, f"{trail}{name} > ")
                if found:
                    return found
    return None


def _registered() -> dict:
    """op name -> OpSpec, every op of the family surface. Family surface:
    the read-only classification now decides which TOOL an op may sit on
    (xl_read, the one tool carrying readOnlyHint=true, holds only read-only
    ops; test_family_surface.py pins the tool annotations)."""
    return registry.public_ops()


def test_every_registered_tool_carries_an_explicit_hint():
    missing = []
    for name, spec in sorted(_registered().items()):
        if not isinstance(spec.read_only, bool):
            missing.append(name)
    assert not missing, (
        f"these ops reach the surface with no read-only classification: {missing}")


def test_the_classification_covers_the_whole_surface():
    registered = set(_registered())
    classified = readonly.READ_ONLY | readonly.MUTATING
    assert registered - classified == set(), (
        "unclassified tools (add them to core/readonly.py): "
        f"{sorted(registered - classified)}")
    assert classified - registered == set(), (
        "core/readonly.py classifies tools that no longer exist: "
        f"{sorted(classified - registered)}")
    assert not (readonly.READ_ONLY & readonly.MUTATING)


def test_read_only_set_matches_the_hand_audited_allowlist():
    on_the_wire = {
        name for name, spec in _registered().items() if spec.read_only is True
    }
    assert on_the_wire == READ_ONLY_TOOLS, (
        "the read-only group changed. Added: "
        f"{sorted(on_the_wire - READ_ONLY_TOOLS)}; removed: "
        f"{sorted(READ_ONLY_TOOLS - on_the_wire)}. Classify the tool by "
        "reading its implementation, then update BOTH this allowlist and "
        "core/readonly.py.")


def test_no_read_only_tool_reaches_a_mutation_path():
    index = _index()
    offenders = []
    for name in sorted(READ_ONLY_TOOLS):
        found = _reaches_mutation(index, name)
        if found:
            offenders.append(found)
    assert not offenders, (
        "tools declared read-only that can change something:" + chr(10)
        + chr(10).join(offenders))


@pytest.mark.parametrize("name", sorted(KNOWN_MUTATORS))
def test_the_mutation_walk_still_has_teeth(name):
    assert _reaches_mutation(_index(), name), (
        f"{name} mutates, so the walk should find it; a walk that finds "
        "nothing proves nothing")


def test_an_unclassified_tool_raises_rather_than_defaulting():
    with pytest.raises(RuntimeError, match="not classified"):
        readonly.read_only_hint("a_tool_nobody_classified")


def test_the_hint_matches_the_module_for_every_tool():
    from xlsx_mcp.surface.groups import OP_GROUP, READ_ONLY_GROUPS
    for name, spec in _registered().items():
        assert spec.read_only is readonly.read_only_hint(name)
        # ...and the op sits on a read-only tool exactly when it is read-only.
        assert (OP_GROUP[name] in READ_ONLY_GROUPS) is spec.read_only, name
