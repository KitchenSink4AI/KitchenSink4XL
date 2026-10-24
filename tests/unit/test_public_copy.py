"""Public-copy guards: the published numbers must be the scripts' numbers.

The Word repo shipped stale marketing figures twice in launch week (an
llms.txt claiming the wrong version, a site claiming a retired tool count),
which is why this guard exists here from the first release instead of after
the first mistake. It makes the failure mode mechanical: if the surface
changes and the copy does not, the suite goes red.

Guarded files: README.md, docs/llms.txt and docs/index.html. The page is a
redirect stub to the community site, so it carries no figures: the figures
are pinned in README.md and docs/llms.txt (_figure_files).

The operations snapshot guard is the one to understand. count_operations.py
derives the headline operations figure from committed source, and
scripts/operations_snapshot.json records the per-tool breakdown that produced
the published number. Adding a dispatch value to any multiplexer moves the
figure, and that is a publishing decision, not a silent one. When this test
fails legitimately:

    .venv/Scripts/python.exe -X utf8 scripts/count_operations.py \\
        --json scripts/operations_snapshot.json

then update every published figure, then update this test's expectations. All
three in the same commit.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
SNAPSHOT = ROOT / "scripts" / "operations_snapshot.json"

#: The figures the public copy currently claims. Change these only together
#: with the copy itself.
PUBLISHED_OPERATIONS = 131
#: The family surface (2.0): seven fixed tools holding 69 operations. Every
#: check against README, llms.txt or docs/index.html figures is a RELEASE GATE
#: (marker public_copy; runs with KS4XL_RELEASE_GATE=1, which the release ship
#: gate sets) and passes as of the 2.0 copy pass.
SURFACE_TOOLS = 7
SURFACE_OPS = 69
BETA_LABEL = re.compile(r"\bbeta\b", re.IGNORECASE)


def _public_files() -> list[Path]:
    candidates = [
        ROOT / "README.md",
        ROOT / "docs" / "llms.txt",
        ROOT / "docs" / "index.html",
    ]
    return [p for p in candidates if p.exists()]


def _figure_files() -> list[Path]:
    """The files that carry the published figures (docs/index.html is a
    redirect stub and carries none)."""
    return [ROOT / "README.md", ROOT / "docs" / "llms.txt"]


def _run(script: str, *args: str) -> str:
    out = subprocess.run(
        [sys.executable, "-X", "utf8", str(ROOT / "scripts" / script), *args],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
    )
    assert out.returncode == 0, f"{script} failed: {out.stderr}"
    return out.stdout


def _measured() -> tuple[dict, int]:
    """(the fixed surface as measured over real stdio, operations), straight
    from the scripts."""
    surf = json.loads(_run("measure_surface.py", "--json").strip().splitlines()[-1])
    ops = _run("count_operations.py")
    n_ops = int(re.search(r"TOTAL DISTINCT OPERATIONS:\s+(\d+)", ops).group(1))
    return surf, n_ops


public_copy = pytest.mark.public_copy


def test_operations_snapshot_matches_source():
    """Every tool's operation count is the one the snapshot recorded."""
    stored = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    out = _run("count_operations.py", "--check", str(SNAPSHOT))
    assert "snapshot MATCHES" in out, (
        "count_operations.py disagrees with scripts/operations_snapshot.json. "
        "Re-run it with --json to re-record, then update the published "
        "figures and PUBLISHED_OPERATIONS in this file.\n" + out
    )
    assert stored["total_operations"] == PUBLISHED_OPERATIONS


@public_copy
def test_published_numbers_match_scripts():
    """The headline figures in the public copy are the scripts' figures."""
    surf, n_ops = _measured()
    assert (surf["tools"], surf["ops"], n_ops) == (
        SURFACE_TOOLS, SURFACE_OPS, PUBLISHED_OPERATIONS
    ), (
        f"scripts now report tools={surf['tools']} ops={surf['ops']} "
        f"operations={n_ops}; update the public copy AND this test together."
    )
    for path in _figure_files():
        text = path.read_text(encoding="utf-8")
        assert str(n_ops) in text, f"{path.name} is missing the operations count"
        assert str(surf["ops"]) in text, f"{path.name} is missing the op count"


@public_copy
def test_token_headlines_match_measurement():
    """The fixed surface's token figure in every public file is measured."""
    surf, _ = _measured()
    n = surf["approx_tokens"]
    k = f"{n / 1000:.1f}k"
    comma = f"{n:,}"
    variants = [k, comma] + [comma.replace(",", sep)
                             for sep in (".", " ", "\u00a0", "\u202f")]
    for path in _figure_files():
        text = path.read_text(encoding="utf-8")
        assert any(v in text for v in variants), (
            f"{path.name}: measured surface figure {n} appears in no accepted "
            f"format {variants}")


@public_copy
def test_readme_surface_table_matches_measurement():
    """Every tool row in the README surface table carries its measured op
    count and tokens (the table the pack table became)."""
    surf, _ = _measured()
    lines = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()
    for row in surf["rows"]:
        cands = [ln for ln in lines if ln.startswith(f"| `{row['tool']}`")]
        assert cands, f"README surface table has no row for {row['tool']!r}"
        k = f"{row['approx_tokens'] / 1000:.1f}k"
        assert f" {row['ops']} " in cands[0] and k in cands[0], (row, cands[0])


def test_calc_labels_in_llms_txt_are_the_wire_values():
    """The agent-facing doc uses the labels the code actually emits.

    The lay-facing copy and the wire used to disagree: the website said
    "calculated" where core/calc.py emitted `computed`, and an agent branching
    on the marketing word got nothing. The pre-release rename settled it on
    `calculated` in both places. This guard keeps llms.txt reading the
    constants rather than a copy of them, so the next divergence goes red.
    """
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp.core import calc

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    for label in (calc.LABEL_VALUE, calc.LABEL_CACHED, calc.LABEL_CALCULATED,
                  calc.LABEL_ABSENT, calc.LABEL_FORMULA):
        assert f"`{label}`" in text, (
            f"llms.txt does not document the {label!r} calc label"
        )


@public_copy
def test_closed_error_codes_are_all_documented():
    """Every code an agent can receive appears in llms.txt Section 5."""
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp import envelope

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    missing = sorted(c for c in envelope.CLOSED_CODES if f"`{c}`" not in text)
    assert not missing, f"llms.txt does not document error codes: {missing}"


@public_copy
def test_tool_membership_matches_llms_txt():
    """Every tool and every op is named in llms.txt."""
    sys.path.insert(0, str(ROOT / "src"))
    from xlsx_mcp.surface.groups import GROUPS, TOOL_NAMES

    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    listed = set(re.findall(r"`([a-z_0-9]+)`", text))
    missing = [n for n in TOOL_NAMES if n not in listed]
    missing += [op for ops in GROUPS.values() for op in ops if op not in listed]
    assert not missing, f"llms.txt is missing {missing}"


@pytest.mark.parametrize("path", _public_files(), ids=lambda p: p.name)
def test_no_em_dashes_in_public_copy(path: Path):
    assert "—" not in path.read_text(encoding="utf-8"), (
        f"{path.name} contains an em dash"
    )


def test_documented_module_entry_point_resolves():
    """The README's `python -m` line must name a module that can be run.

    The install gauntlet found the guessable names all wrong: the
    distribution is kitchensink4xl, the package is xlsx_mcp, and
    `import kitchensink4xl` fails. README now points at
    `python -m xlsx_mcp.server`, so that module has to stay runnable, and it
    has to reach the same entry point the console scripts do.

    1.1 adds the shorter `python -m xlsx_mcp` route to the same two
    documents, which is why both routes are asserted here rather than only
    the `.server` one: the package `__main__` shipped in 1.0.0's tree but
    not in its release, so until now the docs deliberately named only the
    route the published wheel could honor.
    """
    import importlib.util

    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    llms = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    for route in ("python -m xlsx_mcp.server", "python -m xlsx_mcp`"):
        assert route in readme, (
            f"README no longer documents the module entry point {route!r}")
    assert "python -m xlsx_mcp`" in llms and "python -m xlsx_mcp.server" in llms, (
        "docs/llms.txt no longer documents both module entry points")
    assert importlib.util.find_spec("xlsx_mcp.server") is not None
    assert importlib.util.find_spec("xlsx_mcp.__main__") is not None
    from xlsx_mcp import __main__ as pkg_main
    from xlsx_mcp.server import main as server_main

    assert pkg_main.main is server_main


def test_mcp_name_marker_survives():
    first = (ROOT / "README.md").read_text(encoding="utf-8").splitlines()[0]
    assert first == (
        "<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4xl -->"
    ), "README line 1 mcp-name marker was lost"


#: The disclaimer as the canonical license block words it. It
#: says one more thing than the wording it replaced ("or sponsored by"),
#: so the guard moved with it rather than being loosened to a substring
#: that both spellings happen to share.
DISCLAIMER = (
    "Not affiliated with, endorsed by, or sponsored by Microsoft "
    "Corporation."
)


def test_non_affiliation_disclaimer_present():
    for path in _figure_files():
        assert DISCLAIMER in path.read_text(encoding="utf-8"), (
            f"{path.name} is missing the non-affiliation disclaimer")


def test_version_is_consistent_across_manifests():
    """pyproject, server.json, the bundle manifest and the package agree."""
    import tomllib

    pyproject = tomllib.loads(
        (ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    version = pyproject["project"]["version"]

    server_json = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))
    assert server_json["version"] == version
    assert server_json["packages"][0]["version"] == version

    manifest = json.loads(
        (ROOT / "bundle" / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["version"] == version
    assert manifest["server"]["mcp_config"]["args"] == [
        f"kitchensink4xl[com]=={version}"
    ], "the bundle's uvx pin does not match the package version"

    sys.path.insert(0, str(ROOT / "src"))
    import xlsx_mcp
    if xlsx_mcp.__version__ == "0.0.0":
        pytest.skip(
            "package version not stamped yet; stamping "
            "src/xlsx_mcp/__init__.py at ship turns this guard on"
        )
    assert xlsx_mcp.__version__ == version, (
        f"src/xlsx_mcp/__init__.py says {xlsx_mcp.__version__}, "
        f"pyproject says {version}"
    )


# The two version stamps the manifest guard above does not read. Both were
# bumped by hand at every release through 1.2.4 with nothing to catch a miss:
# the llms.txt Release line an agent reads first, and the download links,
# where a miss sends a Claude Desktop user to an older bundle. (The showroom
# that carried the download buttons is now a redirect stub; the links that
# remain are in README.md and docs/llms.txt.)

#: Every release-asset link in the public copy, whatever version it names.
RELEASE_DOWNLOAD_LINK = re.compile(
    r"https://github\.com/KitchenSink4AI/KitchenSink4XL/releases/download/"
    r"[^\"\s<>)]*")


def _pyproject_version() -> str:
    import tomllib

    return tomllib.loads(
        (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    )["project"]["version"]


@public_copy
def test_llms_txt_release_line_matches_the_package_version():
    version = _pyproject_version()
    text = (ROOT / "docs" / "llms.txt").read_text(encoding="utf-8")
    release_lines = [line for line in text.splitlines()
                     if line.startswith("Release:")]
    assert release_lines == [f"Release: v{version}"], (
        f"docs/llms.txt Release line(s) {release_lines!r} do not match "
        f"pyproject version {version}")


@public_copy
def test_release_download_links_match_the_package_version():
    version = _pyproject_version()
    expected = ("https://github.com/KitchenSink4AI/KitchenSink4XL/releases/"
                f"download/v{version}/kitchensink4xl-{version}.mcpb")
    stale = []
    for path in _figure_files():
        links = RELEASE_DOWNLOAD_LINK.findall(path.read_text(encoding="utf-8"))
        stale += [(path.name, link) for link in links if link != expected]
    assert not stale, (
        f"release download links do not match pyproject version "
        f"{version}: {stale}")


@pytest.mark.parametrize(
    ("text", "has_beta_label"),
    [
        ("Beta", True),
        ("beta-1", True),
        ("Eingabetaste", False),
    ],
)
def test_beta_label_detection(text: str, has_beta_label: bool):
    assert bool(BETA_LABEL.search(text)) is has_beta_label


def test_no_beta_language_in_public_copy():
    """Excel ships as a full 1.0. No beta labels anywhere."""
    for path in _public_files():
        text = path.read_text(encoding="utf-8")
        assert BETA_LABEL.search(text) is None, (
            f"{path.name} still carries a beta label"
        )
