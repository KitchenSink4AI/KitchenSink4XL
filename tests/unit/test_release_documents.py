"""Release-document guards: what the published pages and workflow must not say.

These hold the items of the release checklist that a text scan can decide:
no instruction pipes a download into a shell, the publish workflow pins every
action by full commit SHA, and a package claims to be typed only if it ships
a py.typed marker.
"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PUBLISHED_TEXT = ("README.md", "docs/index.html", "docs/llms.txt", "docs/QUICKSTART.md")


def test_no_published_instruction_pipes_a_download_into_a_shell():
    pattern = re.compile(r"(curl|wget)\b[^\n]*\|\s*(sudo\s+)?(ba|z)?sh\b")
    for name in PUBLISHED_TEXT:
        path = ROOT / name
        if path.exists():
            text = path.read_text(encoding="utf-8")
            assert not pattern.search(text), f"{name} pipes a download into a shell"


def test_every_action_in_the_publish_workflow_is_pinned_by_full_sha():
    text = (ROOT / ".github" / "workflows" / "publish.yml").read_text(encoding="utf-8")
    refs = re.findall(r"^\s*-?\s*uses:\s*(\S+)", text, flags=re.MULTILINE)
    assert refs, "no actions found in publish.yml"
    for ref in refs:
        assert re.fullmatch(r"[\w.-]+/[\w.-]+@[0-9a-f]{40}", ref), f"{ref} is not pinned by SHA"


def test_typed_classifier_needs_a_py_typed_marker():
    for pyproject in [ROOT / "pyproject.toml", *(ROOT / "packages").glob("*/pyproject.toml")]:
        data = tomllib.loads(pyproject.read_text(encoding="utf-8"))
        if "Typing :: Typed" in data["project"].get("classifiers", []):
            markers = list((pyproject.parent / "src").rglob("py.typed"))
            assert markers, f"{pyproject} claims Typing :: Typed but ships no py.typed"
