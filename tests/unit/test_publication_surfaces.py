"""Publication-surface checks: scan the shipped text for leftovers.

These checks read the files a user reads (README, llms.txt, quickstart and
the bundle manifest) and fail on:

- ``COPY PENDING``, ``FACTS TO CONVEY`` and ``PLACEHOLDER``, markers that
  mean a string was never filled in;
- the em dash, which is not used in public text;
- a ``user_config`` entry typed anything but ``boolean``. The install screen
  is checkboxes; a free-text box is easy to mistype and break an install.
  ``EXEMPT_TYPED_FIELDS`` names the one field allowed to be a different type;
- an absolute home-directory path, which would only work on one machine.

The ``src/`` trees are deliberately not in scope: the em dashes in refusal
strings are intentional.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

#: Text surfaces scanned. A surface a user reads belongs here.
ENFORCED = (
    "bundle/manifest.json",
    "README.md",
    "docs/llms.txt",
    "docs/QUICKSTART.md",
)

#: (needle, what it means) for the plain-text patterns. Written as literals
#: rather than a regex so the failure message can quote the needle back.
FORBIDDEN_TEXT = (
    ("COPY PENDING", "a copy marker aimed at a future writer"),
    ("FACTS TO CONVEY", "build instructions, not product copy"),
    ("PLACEHOLDER", "an unfilled string"),
    ("\u2014", "an em dash, banned on every public surface"),
)

#: The absolute-path pattern, in both the spellings a JSON file can hold it.
HOME_PATHS = ("C:\\Users\\", "C:\\\\Users\\\\")

#: Typed install-screen fields that are allowed, with the type each may be.
#: The folder picker is a directory chooser rather than a free-text box.
EXEMPT_TYPED_FIELDS: dict[str, str] = {"allowed_root": "directory"}


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def _sources() -> dict[str, str]:
    return {rel: _read(rel) for rel in ENFORCED}


@pytest.mark.parametrize("needle,why", FORBIDDEN_TEXT)
def test_no_working_state_reaches_a_shipping_surface(needle, why):
    for where, text in _sources().items():
        assert needle not in text, (
            f"{where} carries {needle!r}: {why}. Fix the surface.")


def test_no_absolute_home_path_reaches_a_shipping_surface():
    """A hardcoded user directory means the artifact runs on exactly one
    machine."""
    for where, text in _sources().items():
        for pattern in HOME_PATHS:
            assert pattern not in text, (
                f"{where} hardcodes {pattern!r}. The bundle has to launch on "
                f"any machine.")


def test_every_install_screen_field_is_a_checkbox():
    """No typed field on the install screen except the ones named in
    EXEMPT_TYPED_FIELDS."""
    for where in ENFORCED:
        if not where.endswith(".json"):
            continue
        config = json.loads(_read(where)).get("user_config", {})
        typed = sorted(
            key for key, entry in config.items()
            if entry.get("type") != "boolean"
            and entry.get("type") != EXEMPT_TYPED_FIELDS.get(key))
        assert not typed, (
            f"{where} ships typed install-screen fields {typed}. The install "
            f"screen is checkboxes; a value a user has to spell correctly "
            f"belongs in a launch file, not on this screen.")


def test_every_enforced_surface_is_still_a_real_file():
    """A row pointing at a renamed file would quietly turn the check off."""
    for rel in ENFORCED:
        assert (ROOT / rel).is_file(), rel


#: Surfaces cut from a measured source. Every figure below is published on
#: all of them, so restamping one and forgetting another turns this red.
FIGURE_SURFACES = (
    "README.md",
    "docs/llms.txt",
)

#: (figure, what it measures, the script that produces it). Change one only
#: together with a re-run of the named script.
PUBLISHED_FIGURES = (
    ('129', 'workbook operations', 'scripts/count_operations.py'),
    ('69', 'tools including the two pack toggles', 'scripts/measure_surface.py'),
)

#: (string, what it used to mean). Absent from every FIGURE_SURFACES file. A
#: hit means a surface went backwards.
SUPERSEDED_FIGURES = (
    ('Release: v1.0.0', 'an older release line'),
)


def test_every_published_figure_reads_the_same_on_every_surface():
    """A figure published on several surfaces has to be the same figure on
    all of them. Checking presence rather than equality keeps the check
    honest about locale number formatting, which differs by design."""
    missing = []
    for figure, what, source in PUBLISHED_FIGURES:
        for rel in FIGURE_SURFACES:
            if figure not in _read(rel):
                missing.append(f"{rel} does not carry {figure!r} ({what}, "
                               f"measured by {source})")
    assert not missing, (
        "\n".join(missing) + "\nEither a surface was missed by a restamp or "
        "the figure moved and this table was not updated. Re-run the script "
        "that produces it; do not guess the number.")


def test_no_superseded_figure_survives_on_a_published_surface():
    found = []
    for stale, what in SUPERSEDED_FIGURES:
        for rel in FIGURE_SURFACES:
            if stale in _read(rel):
                found.append(f"{rel} still carries {stale!r}, {what}")
    assert not found, (
        "\n".join(found) + "\nRestamp from the measuring script rather than "
        "deleting the row.")


def test_every_figure_surface_is_still_a_real_file():
    for rel in FIGURE_SURFACES:
        assert (ROOT / rel).is_file(), rel
