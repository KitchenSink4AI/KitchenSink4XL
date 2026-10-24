"""Community release checks exercise artifact identity and publishing order."""
from __future__ import annotations

import importlib.util
import zipfile
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location("release_verify", ROOT / "scripts/release_verify.py")
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


def wheel(folder, name, version, requires=()):
    folder.mkdir(parents=True)
    stem = name.replace("-", "_")
    path = folder / f"{stem}-{version}-py3-none-any.whl"
    with zipfile.ZipFile(path, "w") as archive:
        if name in {"kitchensink4xl", "kitchensink4xl-core"}:
            archive.writestr("xlsx_mcp/server.py", "# shared code\n")
        if name == "kitchensink4xl":
            archive.writestr("ks4xl_community_launch/__init__.py", "# launcher\n")
        archive.writestr(f"{stem}-{version}.dist-info/METADATA",
                         f"Name: {name}\nVersion: {version}\n" +
                         "".join(f"Requires-Dist: {r}\n" for r in requires) + "\n")
    return path


def artifacts(tmp_path, requirements=("kitchensink4-license-community==0.2.0",), license_version="0.2.0"):
    dist = tmp_path / "dist"
    wheel(dist / "app", "kitchensink4xl", "2.0.0", requirements)
    wheel(dist / "license", release.LICENSE_DIST, license_version)
    return dist


def test_community_release_has_app_and_license_only(tmp_path):
    dist = artifacts(tmp_path)
    assert set(release.check_build(ROOT, "xl", "v2.0.0", dist)) == {"app", "license"}
    core = wheel(dist / "core", "kitchensink4xl-core", "2.0.0")
    smoke_spec = importlib.util.spec_from_file_location("smoke", ROOT / "scripts/smoke_editions.py")
    smoke = importlib.util.module_from_spec(smoke_spec)
    smoke_spec.loader.exec_module(smoke)
    app, = (dist / "app").glob("*.whl")
    assert smoke.payload_digest(app) == smoke.payload_digest(core)
    public = [["xl_app", ["recalculate", "update_now", "update_notices"],
               {"openWorldHint": True, "readOnlyHint": False}]]
    commercial = [["xl_app", ["recalculate"],
                   {"openWorldHint": False, "readOnlyHint": False}]]
    smoke.assert_shared_surface(public, commercial)
    with pytest.raises(AssertionError):
        smoke.assert_shared_surface(public, [["xl_app", ["recalculate"],
                                              {"openWorldHint": False, "readOnlyHint": True}]])
    import shutil
    community = tmp_path / "community"
    community.mkdir()
    shutil.copy2(app, community / app.name)
    policy, = (dist / "license").glob("*.whl")
    shutil.copy2(policy, community / policy.name)
    assert set(smoke.wheels_for(community)) == {"policy", "launcher"}
    shutil.copy2(core, community / core.name)
    with pytest.raises(AssertionError):
        smoke.wheels_for(community)
    with pytest.raises(release.ReleaseCheckError, match="unexpected entries"):
        release.check_build(ROOT, "xl", "v2.0.0", dist)


@pytest.mark.parametrize("requirements", [
    (), ("kitchensink4-license-community>=0.2.0",),
    ("kitchensink4-license-community==0.2.0; python_version >= '3.12'",),
    ("kitchensink4-license-community==0.2.0", "kitchensink4_license_community==0.2.0"),
    ("kitchensink4-license-community==0.2.0", "kitchensink4xl_core==2.0.0"),
])
def test_community_release_refuses_wrong_dependency_identity(tmp_path, requirements):
    with pytest.raises(release.ReleaseCheckError):
        release.check_build(ROOT, "xl", "v2.0.0", artifacts(tmp_path, requirements))


def test_community_release_refuses_a_different_license_version(tmp_path):
    with pytest.raises(release.ReleaseCheckError, match="version 0.2.0"):
        release.check_build(ROOT, "xl", "v2.0.0", artifacts(tmp_path, license_version="0.3.0"))


def test_plan_checks_license_without_uploading_it(tmp_path, monkeypatch):
    dist = artifacts(tmp_path)
    monkeypatch.setattr(release, "remote_state", lambda index, w: "verified" if w["name"] == release.LICENSE_DIST else "absent")
    assert release.plan(dist, "xl", "require", "https://pypi.org") == {"app": "upload", "license": "skip"}


def test_publish_lane_keeps_release_gate_and_verifies_before_registry():
    jobs = yaml.safe_load((ROOT / ".github/workflows/publish.yml").read_text("utf-8"))["jobs"]
    assert "publish-core" not in jobs
    assert "release-gate" in jobs
    assert jobs["verify-dependencies"]["needs"] == "plan"
    assert jobs["publish-app"]["needs"] == ["plan", "verify-dependencies"]
    assert jobs["registry"]["needs"] == "post-verify"
    build = "\n".join(s.get("run", "") for s in jobs["build"]["steps"])
    assert "pypi-dist/app" in build and "pypi-dist/core" not in build
    pre = "\n".join(s.get("run", "") for s in jobs["verify-dependencies"]["steps"])
    assert "--kinds license --wait" in pre
    assert set(jobs["plan"]["outputs"]) == {"app", "license"}
    publisher, = [s for s in jobs["publish-app"]["steps"] if "gh-action-pypi-publish" in s.get("uses", "")]
    assert publisher["with"]["packages-dir"] == "dist/app/"
