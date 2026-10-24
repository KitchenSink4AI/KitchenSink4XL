"""THE PRE-SHIP GREP GATE: a check that reads the artifact.

Every publishing defect an earlier audit found was catchable by one
grep, and not one of them was caught, because what got checked was the
report about the file rather than the file. An order was given, an agent
reported compliance, and nobody opened the manifest. This is the check that
opens it.

Five patterns, all of them things that mean "somebody was still working
here" and none of them things a stranger should ever read:

- ``COPY PENDING`` and ``FACTS TO CONVEY``, the build's own instructions to
  a future writer. Six of these shipped on a sibling's install screen.
- ``PLACEHOLDER``, the marker this product was printing at users out of its
  update-check strings until the real copy replaced them.
- The em dash, banned across every public surface in every language.
- A ``user_config`` entry typed anything but ``boolean``. The install screen
  is checkboxes; a free-text box is a typo waiting to break somebody's
  install, and the manifest schema has no enum type that would make one
  safe. ``EXEMPT_TYPED_FIELDS`` names the one field that is an
  exception, and naming it here is what keeps the exception from spreading.
- An absolute ``C:\\Users\\`` path, which is a build that only runs on the
  machine that built it.

SCOPE. Everything in ``ENFORCED`` is clean of all five as of the commit
that adds this file, so the gate is green the day it lands and every hit
after that is a regression. ``docs/index.html`` joined that list later,
after four landing pages shipped figures their own READMEs
contradicted; the published-figure guard at the bottom of this file
covers the same class from the other direction. The ``src/`` trees are
deliberately NOT in scope: the em dashes in their refusal strings are an
author decision, not a gate decision. This is the KS4Web gate (``tests/unit/
test_ship_gate.py`` there), ported so the
family is checked by one rule rather than by one product's discipline.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

#: Text surfaces scanned. Every one is clean at the commit that adds this
#: gate. A surface a stranger reads belongs here; add it rather than
#: trusting that nobody edits it.
ENFORCED = (
    "bundle/manifest.json",
    "README.md",
    "docs/llms.txt",
    "docs/QUICKSTART.md",
    "docs/index.html",
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

#: The typed install-screen fields that are exceptions, with the
#: type each is allowed to be. Empty for every product but KS4XL, whose
#: folder picker is exempt by name because a directory chooser is a
#: picker rather than a free-text box.
EXEMPT_TYPED_FIELDS: dict[str, str] = {"allowed_root": "directory"}


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def _sources() -> dict[str, str]:
    return {rel: _read(rel) for rel in ENFORCED}


@pytest.mark.parametrize("needle,why", FORBIDDEN_TEXT)
def test_no_working_state_reaches_a_shipping_surface(needle, why):
    for where, text in _sources().items():
        assert needle not in text, (
            f"{where} carries {needle!r}: {why}. This is the pattern "
            f"found on a sibling's release candidate; fix "
            f"the surface, do not loosen the gate.")


def test_no_absolute_home_path_reaches_a_shipping_surface():
    """A hardcoded user directory means the artifact runs on exactly one
    machine. A sibling's first release candidate launched
    ``C:\\Users\\...\\venv\\Scripts\\...exe`` and could not have installed
    for anybody else."""
    for where, text in _sources().items():
        for pattern in HOME_PATHS:
            assert pattern not in text, (
                f"{where} hardcodes {pattern!r}. The bundle has to launch on "
                f"a machine nobody here owns.")


def test_every_install_screen_field_is_a_checkbox():
    """No typed field on the install screen except the ones exempted by
    name. The rule is stated as a scan rather than as a parity assertion,
    so it holds for the manifest a human installs from."""
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
            f"belongs in a launch file, not on this screen. An exception is "
            f"added to EXEMPT_TYPED_FIELDS on purpose, never because the "
            f"gate was inconvenient.")


def test_every_enforced_surface_is_still_a_real_file():
    """A row pointing at a renamed file would quietly turn into no gate at
    all, which is how a total scope stops being total."""
    for rel in ENFORCED:
        assert (ROOT / rel).is_file(), rel


#: Surfaces cut from a measured source. docs/index.html is a redirect stub
#: to the community site and carries no figures. Every figure below is
#: published on both of these, so a wave that restamps one and forgets another turns this
#: red. The stale landing pages lived in exactly that gap: the
#: README said one number and docs/index.html said another, and nothing
#: checked either against the other.
FIGURE_SURFACES = (
    "README.md",
    "docs/llms.txt",
)

#: (figure, what it measures, the script that produces it[, the surfaces that
#: carry it when not all of them]). Values are the measurement at the commit
#: that adds this guard. Change one only together with a re-run of the named
#: script.
#:
#: Every figure is held to README and docs/llms.txt.
PUBLISHED_FIGURES = (
    ('131', 'workbook operations', 'scripts/count_operations.py'),
    ('69 operations in 7 tools', 'the family surface', 'scripts/measure_surface.py'),
    ('~5.5k', 'tokens for the whole tool list', 'scripts/measure_surface.py'),
    ('5,340', 'tests', 'pytest --collect-only', ('README.md', 'docs/llms.txt')),
)

#: (string, what it used to mean). Absent from every FIGURE_SURFACES file at
#: this commit. A hit means a surface went backwards.
#: '1,294' (the 1.2.2 test count) joins this list once docs/llms.txt
#: is rewritten; it still states it as v1.2.2's figure.
SUPERSEDED_FIGURES = (
    ('1,284', 'the pre-1.1.0 test count'),
    ('1,291', 'the test count before the gate gained its figure guard'),
    ('2,309', 'an earlier 2.0 test count'),
    ('2,369', 'an earlier 2.0 test count'),
    ('2,455', 'an earlier 2.0 test count'),
    ('2,529', 'an earlier 2.0 test count'),
    ('2,542', 'an earlier 2.0 test count'),
    ('2,627', 'an earlier 2.0 test count'),
    ('2,631', 'an earlier 2.0 test count'),
    ('2,725', 'an earlier 2.0 test count'),
    ('2,763', 'an earlier 2.0 test count'),
    ('3,003', 'an earlier 2.0 test count'),
    ('3,053', 'an earlier 2.0 test count'),
    ('3,159', 'an earlier 2.0 test count'),
    ('3,352', 'an earlier 2.0 test count'),
    ('3,395', 'an earlier 2.0 test count'),
    ('3,421', 'an earlier 2.0 test count'),
    ('3,443', 'an earlier 2.0 test count'),
    ('3,461', 'an earlier 2.0 test count'),
    ('3,525', 'an earlier 2.0 test count'),
    ('3,671', 'an earlier 2.0 test count'),
    ('3,735', 'an earlier 2.0 test count'),
    ('4,756', 'an earlier 2.0 test count'),
    ('4,730', 'an earlier 2.0 test count'),
    ('4,759', 'an earlier 2.0 test count'),
    ('4,779', 'an earlier 2.0 test count'),
    ('4,704', 'an earlier 2.0 test count'),
    ('4,672', 'an earlier 2.0 test count'),
    ('4,626', 'an earlier 2.0 test count'),
    ('4,593', 'an earlier 2.0 test count'),
    ('4,556', 'an earlier 2.0 test count'),
    ('4,422', 'an earlier 2.0 test count'),
    ('3,736', 'an earlier 2.0 test count'),
    ('4,395', 'an earlier 2.0 test count'),
    ('219 operations across 108 tools', "the sibling's pre-2.1.0 surface"),
    ('138 tools', "the sibling's pre-1.2.0 surface"),
    ('Release: v1.0.0', 'the previous release line'),
    ('5,248', 'the test count before the update-notice tests'),
    ('5,310', 'the test count before the consent-after-refresh test'),
    ('5,311', 'the test count before the dialog-step version test'),
    ('5,312', 'the test count before the old-launch-settings tests'),
    ('5,332', 'the test count before the Community app-wheel checks'),
    ('5,433', 'the tool-list figure before update_now and update_notices'),
    ('5.433', 'the same figure with a german or spanish separator'),
    ('5 433', 'the same figure with a french separator'),
    ('67 operations', 'the 2.0.0 operation count before update_now and update_notices'),
    ('129 actions', 'the 2.0.0 action count before update_now and update_notices'),
    ('129 counting', 'the same count in the README and llms.txt wording'),
)


def test_every_published_figure_reads_the_same_on_every_surface():
    """A figure published on three surfaces has to be the same figure on all
    three. Checking presence rather than equality keeps the guard honest
    about locale number formatting, which differs by design.

    Looped rather than parametrized on purpose: this repo's collected test
    count is one of the figures below, so a parametrized table would move
    the number it is guarding every time somebody added a row to it."""
    missing = []
    for figure, what, source, *only in PUBLISHED_FIGURES:
        for rel in (only[0] if only else FIGURE_SURFACES):
            if figure not in _read(rel):
                missing.append(f"{rel} does not carry {figure!r} ({what}, "
                               f"measured by {source})")
    assert not missing, (
        "\n".join(missing) + "\nEither a surface was missed by a restamp or "
        "the figure moved and this table was not updated. Re-run the script "
        "that produces it; do not guess the number.")


@pytest.mark.timeout(600)
def test_the_published_test_count_is_the_collected_count():
    """The figure guard above checks that the test count
    is present, not that it is right, and the README said 1,294 while 1,368
    were collected. The count is now measured: pytest --collect-only over
    tests/unit, in a fresh process with this one's environment.

    The published figure describes the PUBLIC repository, which carries the
    Community edition only, so the count is always taken on the Community path
    (KS4_ADAPTER_EDITION=community), here and in private staging alike."""
    figure = next(row[0] for row in PUBLISHED_FIGURES if row[1] == "tests")
    out = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "pytest", "--collect-only", "-q",
         "-p", "no:cacheprovider", "tests/unit"],
        cwd=ROOT, env=dict(os.environ, KS4_ADAPTER_EDITION="community"),
        stdin=subprocess.DEVNULL,
        capture_output=True, text=True, timeout=540,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    m = re.search(r"(\d+) tests? collected", out.stdout)
    assert m, out.stdout[-2000:] + out.stderr[-2000:]
    collected = int(m.group(1))
    assert f"{collected:,}" == figure, (
        f"{collected:,} tests are collected but the published figure is {figure}; "
        "restamp the figure on README.md and docs/llms.txt and in "
        "PUBLISHED_FIGURES (the number only), and move the old one to "
        "SUPERSEDED_FIGURES")


def test_no_superseded_figure_survives_on_a_published_surface():
    """The regression guard for the defect this file was widened over."""
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
