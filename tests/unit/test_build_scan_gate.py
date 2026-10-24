"""The deterministic build gate: no Community guidance in Business artifacts.

Runtime suppression is not the requirement. The requirement is that the
strings are physically absent from every shipped Business file, and that the
notices that MUST ship are still there, so the gate has to fail in both
directions. It also has to fail when someone injects the guidance into a
build input, which is the case this file spends most of its length on.
"""
import importlib.util
import json
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path

import pytest

# The core version these anchors track (the launcher pyprojects carry the
# same string); read, not hard-coded, so the family version bump keeps the
# drift check meaningful rather than stale.
import tomllib as _tomllib
from pathlib import Path as _Path
_CORE_VERSION = _tomllib.loads((_Path(__file__).resolve().parents[2] / "pyproject.toml").read_text("utf-8"))["project"]["version"]
_VERSION_ANCHOR = f'version = "{_CORE_VERSION}"'


ROOT = Path(__file__).resolve().parents[2]

spec = importlib.util.spec_from_file_location(
    "split_build", ROOT / "scripts/build_editions.py")
build_script = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_script)

#: The build reads the approved notice text out of the ONE canonical adapter
#: source at its pinned commit, not out of a copy in this repository.
def test_the_gate_reads_its_needles_from_the_pinned_canonical_source():
    import adapter_source
    source = build_script.adapter_source.notices_source()
    assert source == adapter_source.notices_source()
    assert source.is_file()
    assert not str(source).startswith(str(ROOT))


def _wheel(path: Path, members: dict) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        for name, text in members.items():
            archive.writestr(name, text)
    return path


NOTICES = "\n".join(build_script.REQUIRED_NOTICES)


def test_the_gate_knows_what_it_is_looking_for():
    needles = build_script.community_needles()
    assert len(needles) >= 5
    from ks4_license_community import notices
    assert any(n in notices.WARNING_13 for n in needles)
    # Short fragments are excluded on purpose; they collide with ordinary prose.
    assert all(len(n) >= 40 for n in needles)


def test_a_gate_with_no_needles_refuses_to_pass_anything(tmp_path):
    clean = _wheel(tmp_path / "clean.whl", {"a/LICENSE": NOTICES})
    with pytest.raises(build_script.BuildScanError, match="would pass anything"):
        build_script.scan_business_artifacts([clean], needles=())


def test_a_clean_artifact_set_passes(tmp_path):
    clean = _wheel(tmp_path / "clean.whl", {
        "ks4_license_business/policy.py": "PRODUCT_ID = 'kitchensink4xl-business'\n",
        "dist-info/licenses/LICENSE": NOTICES,
    })
    report = build_script.scan_business_artifacts([clean])
    assert report["members_scanned"] == 2
    assert report["needles"] >= 5


@pytest.mark.parametrize("member", [
    "xlsx_mcp/server.py", "dist-info/METADATA", "INSTALL.md",
])
def test_injecting_the_guidance_into_any_member_fails_the_build(tmp_path, member):
    from ks4_license_community import notices
    poisoned = _wheel(tmp_path / "poisoned.whl", {
        member: "prose prose\n" + notices.WARNING_13 + "\nmore prose",
        "dist-info/licenses/LICENSE": NOTICES,
    })
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.scan_business_artifacts([poisoned])


def test_the_gate_reads_inside_a_nested_bundle(tmp_path):
    from ks4_license_community import notices
    bundle = _wheel(tmp_path / "bundle.mcpb", {
        "manifest.json": json.dumps({"name": "kitchensink4xl-business"}),
        "INSTALL.md": notices.COMMUNITY_GUIDANCE,
        "LICENSE": NOTICES,
    })
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.scan_business_artifacts([bundle])


def test_stripping_a_required_notice_also_fails_the_build(tmp_path):
    """Zero Community guidance never means deleting a copyright notice."""
    stripped = _wheel(tmp_path / "stripped.whl", {"a/LICENSE": "nothing here"})
    with pytest.raises(build_script.BuildScanError, match="Required notices"):
        build_script.scan_business_artifacts([stripped])


@pytest.mark.business_sources
def test_the_allowlist_lets_the_commercial_notice_keep_its_agpl_sentences():
    """The commercial notice says the Community distribution remains
    available under AGPL-3.0. That sentence belongs in a Business artifact
    and a word-level scan would have deleted it."""
    notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
    assert "AGPL-3.0" in notice
    for needle in build_script.community_needles():
        assert needle not in notice


def test_the_business_install_doc_drops_only_the_community_section():
    text = build_script.business_install_doc()
    assert "## Community" not in text
    assert "## Business" in text
    assert "## Install and connect" in text
    for needle in build_script.community_needles():
        assert needle not in text


# ------------------------------------------------------ anchored rewrites


@pytest.mark.parametrize("anchor", [
    '    "kitchensink4xl-core",',
    '    "kitchensink4-license-business",',
    _VERSION_ANCHOR,
], ids=["core_dependency", "business_adapter_dependency", "version"])
@pytest.mark.parametrize("drift", ["missing", "duplicate"])
@pytest.mark.business_sources
def test_business_build_refuses_metadata_drift(tmp_path, anchor, drift):
    """The fail-closed property the whole split rests on: an anchor that is
    absent or ambiguous stops the build instead of silently skipping a
    rewrite."""
    source = ROOT / "packages/kitchensink4xl-business/pyproject.toml"
    _assert_anchor_fails_closed(tmp_path, source, anchor, drift)


@pytest.mark.parametrize("anchor", [_VERSION_ANCHOR, 'license = "AGPL-3.0-only"'],
                         ids=["version", "license"])
@pytest.mark.parametrize("drift", ["missing", "duplicate"])
def test_core_relabelling_refuses_metadata_drift(tmp_path, anchor, drift):
    """The commercial core is produced by rewriting exactly these two lines.
    If either anchor drifts, the build stops rather than shipping a core
    wheel labelled with the wrong grant."""
    _assert_anchor_fails_closed(tmp_path, ROOT / "pyproject.toml", anchor, drift)


def _assert_anchor_fails_closed(tmp_path, source, anchor, drift):
    text = source.read_text("utf-8")
    assert text.count(anchor) == 1, f"anchor drifted out of {source}"
    # Duplicate anchors in comments keep TOML valid but expose ambiguity.
    drifted = (text + "\n# " + anchor + "\n" if drift == "duplicate"
               else text.replace(anchor, "", 1))
    staged = tmp_path / "pyproject.toml"
    staged.write_text(drifted, encoding="utf-8")
    with pytest.raises(ValueError, match="exactly one build metadata anchor"):
        build_script.staged_config(staged, [(anchor, anchor + "  # rewritten")])


# --------------------------------------------------- the license gate


def _licensed_wheel(path, expression):
    members = {"pkg/__init__.py": "x = 1\n"}
    if expression is not None:
        members["pkg-1.0.dist-info/METADATA"] = (
            "Metadata-Version: 2.4\nName: pkg\nVersion: 1.0\n"
            f"License-Expression: {expression}\n\nbody\n")
    else:
        members["pkg-1.0.dist-info/METADATA"] = (
            "Metadata-Version: 2.4\nName: pkg\nVersion: 1.0\n\nbody\n")
    return _wheel(path, members)


def test_an_agpl_wheel_in_a_business_distribution_fails_the_build(tmp_path):
    """The dual-grant defect: a core wheel saying AGPL inside a commercial
    bundle puts two conflicting grants in one customer's hands."""
    wheel = _licensed_wheel(tmp_path / "core.whl", "AGPL-3.0-only")
    with pytest.raises(build_script.BuildScanError, match="Conflicting license"):
        build_script.scan_license_metadata([wheel], "business")


def test_a_commercial_wheel_in_a_community_distribution_fails_the_build(tmp_path):
    """Same defect wearing the other hat."""
    wheel = _licensed_wheel(tmp_path / "core.whl", "LicenseRef-KitchenSink-Commercial")
    with pytest.raises(build_script.BuildScanError, match="Conflicting license"):
        build_script.scan_license_metadata([wheel], "community")


@pytest.mark.parametrize("edition,expression", [
    ("business", "LicenseRef-KitchenSink-Commercial"),
    ("community", "AGPL-3.0-only"),
])
def test_a_correctly_labelled_distribution_passes(tmp_path, edition, expression):
    wheel = _licensed_wheel(tmp_path / "core.whl", expression)
    assert build_script.scan_license_metadata([wheel], edition)


def test_a_wheel_with_no_license_expression_fails_the_build(tmp_path):
    wheel = _licensed_wheel(tmp_path / "core.whl", None)
    with pytest.raises(build_script.BuildScanError, match="no license expression"):
        build_script.scan_license_metadata([wheel], "business")


def test_a_license_gate_with_no_wheels_refuses_to_pass_anything(tmp_path):
    with pytest.raises(build_script.BuildScanError, match="would pass anything"):
        build_script.scan_license_metadata([], "business")


@pytest.mark.timeout(600)
@pytest.mark.business_sources
def test_a_real_business_build_ships_a_commercially_labelled_core(tmp_path):
    """End to end: the Business core wheel is a distinct artifact identity
    with its own grant, built from the same source revision."""
    artifacts = build_script.build("business", tmp_path / "b", date(2026, 9, 13),
                                   None, "primary")
    declared = build_script.scan_license_metadata(artifacts, "business")
    assert all("AGPL" not in expression for expression in declared.values())
    core, = [n for n in declared if n.startswith("kitchensink4xl_core-")]
    assert "+commercial" in core
    assert declared[core] == build_script.COMMERCIAL


@pytest.mark.timeout(600)
def test_a_real_community_build_ships_an_agpl_core_at_the_public_version(tmp_path):
    artifacts = build_script.build("community", tmp_path / "c", date(2026, 9, 13),
                                   None, "primary")
    declared = build_script.scan_license_metadata(artifacts, "community")
    assert all("LicenseRef" not in expression for expression in declared.values())
    app, = [a for a in artifacts if a.name.startswith("kitchensink4xl-") and a.suffix == ".whl"]
    assert declared[app.name] == build_script.AGPL
    assert len([a for a in artifacts if a.suffix == ".whl"]) == 2
    assert not any(a.name.startswith("kitchensink4xl_core-") for a in artifacts)
    from email.parser import HeaderParser
    from packaging.requirements import Requirement
    with zipfile.ZipFile(app) as wheel:
        names = wheel.namelist()
        assert "xlsx_mcp/__init__.py" in names
        assert "xlsx_mcp/server.py" in names
        assert "ks4xl_community_launch/__init__.py" in names
        metadata, = [n for n in names if n.endswith(".dist-info/METADATA")]
        info = HeaderParser().parsestr(wheel.read(metadata).decode("utf-8"))
        requirements = [Requirement(v) for v in info.get_all("Requires-Dist")]
        pins = [v for v in requirements if v.name == "kitchensink4-license-community"]
        assert len(pins) == 1 and str(pins[0].specifier) == "==0.2.0"
        assert pins[0].marker is None and pins[0].url is None
        assert not any(v.name == "kitchensink4xl-core" for v in requirements)
        expected = _tomllib.loads((ROOT / "pyproject.toml").read_text("utf-8"))["project"]["dependencies"]
        assert {str(Requirement(v)) for v in expected} <= {str(v) for v in requirements}
        scripts, = [n for n in names if n.endswith(".dist-info/entry_points.txt")]
        commands = wheel.read(scripts).decode("utf-8")
        assert "xl-mcp = ks4xl_community_launch:main" in commands
        assert "kitchensink4xl = ks4xl_community_launch:main" in commands
    bundle, = [a for a in artifacts if a.suffix == ".mcpb"]
    with zipfile.ZipFile(bundle) as archive:
        assert sorted(n for n in archive.namelist() if n.endswith(".whl")) == sorted(
            a.name for a in artifacts if a.suffix == ".whl")


@pytest.mark.timeout(600)
@pytest.mark.business_sources
def test_a_poisoned_core_label_fails_the_real_business_build(tmp_path, monkeypatch):
    """Skip the relabelling and the build must refuse to ship the bundle."""
    real = build_script.build_core

    def unlabelled(edition, output, scratch):
        return real("community", output, scratch)
    monkeypatch.setattr(build_script, "build_core", unlabelled)
    with pytest.raises((build_script.BuildScanError, ValueError)):
        build_script.build("business", tmp_path / "poison", date(2026, 9, 13),
                           None, "primary")


@pytest.mark.business_sources
def test_the_required_notices_survive_the_commercial_relabelling(tmp_path):
    """Relabelling a distribution never means dropping a notice."""
    artifacts = build_script.build("business", tmp_path / "n", date(2026, 9, 13),
                                   None, "primary")
    assert build_script.scan_business_artifacts(artifacts)["required_notices"] == 2


@pytest.mark.parametrize("anchor", [
    '    "kitchensink4-license-community",',
])
def test_community_launcher_anchors_are_present(anchor):
    source = ROOT / "packages/kitchensink4xl-community/pyproject.toml"
    assert source.read_text("utf-8").count(anchor) == 1


def test_staged_metadata_that_stops_being_valid_toml_is_rejected(tmp_path):
    staged = tmp_path / "pyproject.toml"
    staged.write_text('[project]\nname = "x"\n', encoding="utf-8")
    with pytest.raises(Exception):
        build_script.staged_config(staged, [('name = "x"', 'name = "x\n[project]')])


# --------------------------------------------------------- the real build


@pytest.mark.timeout(600)
@pytest.mark.business_sources
def test_a_real_business_build_passes_the_gate_and_ships_no_community_module(tmp_path):
    """End to end: build the Business distribution and prove the gate ran on
    real artifacts rather than on fixtures shaped like them."""
    output = tmp_path / "business"
    artifacts = build_script.build("business", output, date(2026, 9, 13), None, "primary")
    names = [a.name for a in artifacts]
    assert any(n.startswith("kitchensink4xl_core-") for n in names)
    assert any(n.startswith("kitchensink4_license_business-") for n in names)
    assert any(n.startswith("kitchensink4xl_business-") for n in names)
    assert any(n.endswith(".mcpb") for n in names)
    for artifact in artifacts:
        with zipfile.ZipFile(artifact) as archive:
            assert not [n for n in archive.namelist() if "ks4_license_community" in n]
    # And the gate itself is not a no-op on this artifact set.
    assert build_script.scan_business_artifacts(artifacts)["members_scanned"] > 20


@pytest.mark.timeout(600)
@pytest.mark.business_sources
def test_a_poisoned_build_input_fails_the_real_build(tmp_path, monkeypatch):
    """Inject the guidance into a build input and the build must refuse."""
    from ks4_license_community import notices
    poisoned = build_script.business_install_doc() + "\n\n" + notices.WARNING_13
    monkeypatch.setattr(build_script, "business_install_doc", lambda: poisoned)
    with pytest.raises(build_script.BuildScanError, match="Community guidance"):
        build_script.build("business", tmp_path / "poisoned", date(2026, 9, 13),
                           None, "primary")


def test_the_build_script_runs_headless_without_a_console_prompt():
    done = subprocess.run([sys.executable, "-X", "utf8",
                           str(ROOT / "scripts/build_editions.py"), "--help"],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0
    assert "--release-date" in done.stdout


# ------------------------------------------- the published long description
#
# kitchensink4xl is the distribution this repository publishes to PyPI, and
# on this branch it is built from packages/kitchensink4xl-community, which declared no
# readme. A release cut from here would have published the blessed one-line
# description and a project page with nothing on it, while the same release
# cut from the community branch carried the whole README, because there the
# published package is the root project and the root project declares one.
# build_launcher already stages the repository README beside the staged
# pyproject, so the declaration is the whole fix, and the assertion is against
# the built wheel rather than against the toml.


def test_the_community_wheel_carries_the_readme_as_its_long_description(
        tmp_path):
    artifacts = build_script.build("community", tmp_path / "c",
                                   date(2026, 9, 16), None)
    wheels = [a for a in artifacts
              if a.name.startswith("kitchensink4xl-")
              and a.suffix == ".whl"]
    assert len(wheels) == 1, [a.name for a in artifacts]
    with zipfile.ZipFile(wheels[0]) as archive:
        name = next(n for n in archive.namelist()
                    if n.endswith(".dist-info/METADATA"))
        metadata = archive.read(name).decode("utf-8").replace(
            "\r\n", "\n")
    headers, _, body = metadata.partition("\n\n")
    assert "Description-Content-Type: text/markdown" in headers
    readme = (ROOT / "README.md").read_text("utf-8")
    first = readme.splitlines()[0]
    assert first and first in body
    # Not a truncated one: the README's own last non-empty line is in there.
    last = [line for line in readme.splitlines() if line.strip()][-1]
    assert last in body
    assert len(body) > len(readme) * 0.9


# ------------------------------------------------ owner-test mode
# (the same design as the Word editor)


def _gate_module():
    import importlib.util
    from pathlib import Path as _P
    root = _P(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("owner_gate_build", root / "scripts/build_editions.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _owner_wheel(tmp_path, name, text):
    import zipfile
    p = tmp_path / name
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("pkg/copy.py", text)
    return p


@pytest.mark.parametrize("text", [
    'X = "' + "COPY" + '_SLOT: facts"', "FACTS" + " TO CONVEY: y", "COPY PENDING",
])
def test_owner_test_gate_refuses_an_unfilled_slot(tmp_path, text):
    build = _gate_module()
    with pytest.raises(SystemExit):
        build.scan_owner_test([_owner_wheel(tmp_path, "a.whl", text)])


def test_owner_test_gate_refuses_the_builders_home_folder(tmp_path):
    from pathlib import Path as _P
    build = _gate_module()
    with pytest.raises(SystemExit):
        build.scan_owner_test([_owner_wheel(tmp_path, "b.whl", f"x = r'{_P.home()}\\f.xlsx'")])


def test_owner_test_gate_passes_a_clean_artifact(tmp_path):
    build = _gate_module()
    report = build.scan_owner_test([_owner_wheel(tmp_path, "c.whl", "X = 'finished copy'")])
    assert report["members_scanned"] == 1


def test_owner_test_gate_sees_a_runtime_assembled_slot(tmp_path):
    """copy.py writes its markers split, so only the slot helper's calls
    reveal an unfilled slot in the shipped source (the marker scan alone
    passed an unfilled copy.py)."""
    import zipfile
    build = _gate_module()
    head = 'SLOT = "COPY" "_SLOT:"\ndef _slot(facts):\n    return SLOT + facts\n'
    p = tmp_path / "core.whl"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("xlsx_mcp/surface/copy.py", head + 'X = _slot("facts")\n')
    with pytest.raises(SystemExit):
        build.scan_owner_test([p])
    q = tmp_path / "filled.whl"
    with zipfile.ZipFile(q, "w") as z:
        z.writestr("xlsx_mcp/surface/copy.py", head + 'X = "finished copy"\n')
    assert build.scan_owner_test([q])["members_scanned"] == 1
