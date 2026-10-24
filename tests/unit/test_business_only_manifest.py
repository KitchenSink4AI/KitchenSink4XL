"""The Community-only test path is exact, not a silent skip list.

tests/conftest.py names the Business-only test modules (not collected when
KS4_ADAPTER_EDITION=community, the public checkout) and the individual
Business-only tests (deselected there). This module runs in BOTH editions:
the module list must equal the set of test modules that import Business code
at top level, every listed test must exist, the Community path must really be
Community-only, and the private path must really carry Business.
"""
from __future__ import annotations

import ast
import importlib.util
import os
from pathlib import Path

TESTS = Path(__file__).resolve().parents[1]
BUSINESS_PACKAGES = ("ks4_license_business", "ks4word_business_launch",
                     "ks4xl_business_launch", "ks4ppt_business_launch")


def _conftest_tuple(name):
    tree = ast.parse((TESTS / "conftest.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == name for t in node.targets):
            return tuple(ast.literal_eval(node.value))
    raise AssertionError(f"{name} is not defined in tests/conftest.py")


def _top_level_business_importers():
    found = set()
    for path in TESTS.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        stack = list(tree.body)
        while stack:
            node = stack.pop()
            if isinstance(node, (ast.If, ast.Try)):
                stack.extend(node.body + node.orelse + getattr(node, "finalbody", [])
                             + [h for hd in getattr(node, "handlers", []) for h in hd.body])
                continue
            names = []
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names = [node.module]
            if any(n.split(".")[0] in BUSINESS_PACKAGES for n in names):
                found.add(path.relative_to(TESTS).as_posix())
    return found


def test_business_only_modules_are_exactly_the_top_level_importers():
    assert set(_conftest_tuple("BUSINESS_ONLY_MODULES")) == _top_level_business_importers()


def test_every_listed_business_only_test_exists():
    for nodeid in _conftest_tuple("BUSINESS_ONLY_TESTS"):
        rel, _, rest = nodeid.split("[", 1)[0].partition("::")
        path = TESTS.parent / rel
        assert path.is_file(), nodeid
        names = rest.split("::")
        body = ast.parse(path.read_text(encoding="utf-8")).body
        for name in names:
            match = [n for n in body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef,
                                                       ast.ClassDef)) and n.name == name]
            assert match, nodeid
            body = getattr(match[0], "body", [])
        assert not any(nodeid.startswith("tests/" + m) for m in _conftest_tuple("BUSINESS_ONLY_MODULES"))


#: Source markers of Business or private-only material inside a test function.
MARKERS = ("ks4_license_business", "kitchensink4-license-business", "BUSINESS_SRC",
           "BUSINESS_POLICY_DIST", "COMMERCIAL_LICENSE_NOTICE", "audit-community-business",
           "_business_launch", "-business\"", "business/", "LEDGER")
#: Path fragments of material the public tree never carries (hard-denied by the
#: snapshot filter), matched in module-level assignments.
PATH_MARKERS = ('"internal"', "internal/", "DEPENDENCY_LEDGER", '"research"')
#: Tests that mention those markers and still run honestly on the Community path,
#: because they assert the Business material is ABSENT or refused, or only use the
#: names as synthetic fixture text.
COMMUNITY_SAFE = {
    "test_a_clean_artifact_set_passes": "synthetic artifact names",
    "test_the_gate_reads_inside_a_nested_bundle": "synthetic artifact names",
    "test_a_forked_adapter_fails_the_adapter_gate": "edition-aware: forks the resolved adapter",
    "test_the_selected_edition_is_what_is_importable": "asserts Business is absent here",
    "test_a_core_without_a_policy_refuses_before_any_tool_registers": "checks the setup error names the launchers",
    "test_core_never_imports_a_default_policy": "asserts the core imports no policy",
    "test_community_package_cannot_reach_a_document_api": "asserts Business is not reached",
    "test_community_package_cannot_reach_a_deck_api": "asserts Business is not reached",
    "test_this_repository_carries_no_copy_of_a_shared_adapter": "asserts absence",
    "test_community_path_never_hands_out_the_business_adapter": "asserts the refusal",
    "test_this_core_composes_the_powerpoint_order": "its [business] case is listed",
}


def test_every_test_that_touches_business_material_is_listed_or_safe():
    listed = {n.split("[", 1)[0] for n in _conftest_tuple("BUSINESS_ONLY_TESTS")}
    modules = set(_conftest_tuple("BUSINESS_ONLY_MODULES"))
    unlisted = []
    for path in sorted(TESTS.rglob("test_*.py")):
        rel = path.relative_to(TESTS).as_posix()
        if rel in modules or path.resolve() == Path(__file__).resolve():
            continue
        src = path.read_text(encoding="utf-8")
        body = ast.parse(src).body
        # module-level names bound to private-only paths count as markers too
        # (e.g. PARITY = ROOT / "docs" / "internal" / ...).
        local = []
        for node in body:
            if isinstance(node, ast.Assign):
                value = ast.get_source_segment(src, node.value) or ""
                if any(m in value for m in MARKERS + PATH_MARKERS):
                    local += [t.id for t in node.targets if isinstance(t, ast.Name)]
        markers = MARKERS + tuple(local)
        stack = [(n, "") for n in body]
        while stack:
            node, prefix = stack.pop()
            if isinstance(node, ast.ClassDef):
                stack.extend((n, prefix + node.name + "::") for n in node.body)
                continue
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) or \
                    not node.name.startswith("test"):
                continue
            seg = "".join(ast.get_source_segment(src, x) or "" for x in [node, *node.decorator_list])
            if not any(m in seg for m in markers):
                continue
            nodeid = f"tests/{rel}::{prefix}{node.name}"
            if nodeid not in listed and node.name not in COMMUNITY_SAFE:
                unlisted.append(nodeid)
    assert unlisted == [], ("tests that read Business or private-only material must be in "
                            "tests/conftest.py BUSINESS_ONLY_TESTS or reviewed as safe here")


def test_the_selected_edition_is_what_is_importable():
    edition = (os.environ.get("KS4_ADAPTER_EDITION") or "").strip().lower() or "all"
    assert importlib.util.find_spec("ks4_license_community") is not None
    business = importlib.util.find_spec("ks4_license_business")
    if edition == "community":
        assert business is None, "the Community-only path must not see the Business adapter"
    else:
        assert business is not None, "the private path must keep the Business adapter"
