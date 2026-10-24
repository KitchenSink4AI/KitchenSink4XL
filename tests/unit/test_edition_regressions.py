"""Regression cases for defects independently found during the edition audit."""
import os
import time
import pytest
import tomllib as _tomllib
from pathlib import Path as _Path

# Read, not hard-coded: the family version bump keeps the drift check live.
_VERSION_ANCHOR = 'version = "%s"' % _tomllib.loads((_Path(__file__).resolve().parents[2] / "pyproject.toml").read_text("utf-8"))["project"]["version"]
from xlsx_mcp.core import update_check
from xlsx_mcp.core import safesave


@pytest.mark.parametrize("release", [[], [{"yanked": True}], None, "invalid", [None]])
def test_uninstallable_release_cannot_return_via_info_fallback(release):
    assert update_check._latest_stable({
        "info": {"version": "9.0.0"}, "releases": {"9.0.0": release}
    }) is None


def test_error_diagnostic_does_not_echo_private_exception_message():
    result = update_check._reason(OSError("C:\\Private\\company.pem: secret proxy credential"))
    assert result == "OSError"


@pytest.mark.parametrize("locks", [safesave])
@pytest.mark.parametrize("stamp", [0, None, time.time() - 36000],
                         ids=["zero", "none", "ten_hours_ago"])
def test_live_writer_is_never_evicted_just_because_it_is_old(monkeypatch, locks, stamp):
    monkeypatch.setattr(locks, "_pid_alive", lambda pid: True)
    info = {"pid": os.getpid() + 1000, "token": "other-process",
            "pid_created": 123, "time": stamp}
    assert not locks._is_stale(info)


@pytest.mark.parametrize("locks", [safesave])
def test_dead_writer_still_recoverable(monkeypatch, locks):
    monkeypatch.setattr(locks, "_pid_alive", lambda pid: False)
    assert locks._is_stale({"pid": os.getpid() + 1000, "time": time.time()})


def test_entrypoint_disables_framework_banner_update_path(monkeypatch):
    from xlsx_mcp import server
    seen = {}
    real = server.create_server

    def built(settings=None):
        mcp = real(settings)
        monkeypatch.setattr(mcp, "run", lambda **kwargs: seen.update(kwargs))
        return mcp
    monkeypatch.setattr(server, "create_server", built)
    monkeypatch.setattr(server._star_nudge, "announce_once", lambda: None)
    server.main()
    assert seen["show_banner"] is False


@pytest.mark.parametrize("locks", [safesave])
def test_remote_host_lock_is_not_reclaimed_or_reentered(tmp_path, monkeypatch, locks):
    import json
    monkeypatch.setattr(locks, "_local_host", lambda: "LOCAL")
    def forbidden(*args):
        pytest.fail("Remote PID must not be checked locally")
    monkeypatch.setattr(locks, "_pid_alive", forbidden)
    info = {"host": "REMOTE", "pid": os.getpid(), "token": locks._OWNER_TOKEN,
            "time": 0, "pid_created": 0}
    assert not locks._is_stale(info)
    assert not locks._is_ours(info)
    path = tmp_path / "remote.lock"
    payload = json.dumps(info).encode("utf-8")
    path.write_bytes(payload)
    monkeypatch.setattr(locks, "LOCK_WAIT_SECONDS", -1)
    with pytest.raises(safesave.MutationLockTimeout):
        locks._acquire_lockfile(path, "example.xlsx")
    assert path.read_bytes() == payload


@pytest.mark.parametrize("locks", [safesave])
def test_host_case_and_legacy_payloads_keep_local_dead_pid_recovery(monkeypatch, locks):
    monkeypatch.setattr(locks, "_local_host", lambda: "Local")
    monkeypatch.setattr(locks, "_pid_alive", lambda pid: False)
    for host in ["LOCAL", "local", ""]:
        assert locks._is_stale({"host": host, "pid": 123})
    assert locks._is_stale({"pid": 123})


@pytest.mark.business_sources
def test_approved_warning_preserves_exact_section_symbol():
    """The approved v3 wording, byte for byte, with a real section symbol.

    The text moved into the shared Community adapter in the licensing split,
    so it is read from the canonical pinned source rather than from a module
    in this repository. The mojibake regression it guards against did not move
    with it: the docs here still have to carry a real section symbol.
    """
    from pathlib import Path

    import adapter_source
    from ks4_license_community.notices import WARNING_13, COMMUNITY_GUIDANCE
    expected = ("This deployment likely violates AGPL-3.0 \u00a713. Providing network "
                "access to a modified version without offering users its source code is "
                "a license violation, not a gray area. Certain commercial and enterprise "
                "deployments require a paid license. Purchase at "
                "kitchensink4.ai/licenses or contact licensing@kitchensink4.ai.")
    assert WARNING_13 == expected
    assert "AGPL \u00a713 source offer" in COMMUNITY_GUIDANCE
    root = Path(__file__).resolve().parents[2]
    checked = [adapter_source.notices_source()]
    checked += [root / name for name in ["docs/editions.md",
                                         "docs/audit-community-business.md"]]
    for path in checked:
        data = path.read_bytes()
        assert bytes.fromhex("c2a7") + b"13" in data, path
        assert b"?13" not in data, path


def test_the_guidance_keeps_its_carve_outs_and_its_register():
    """Coercive instruction register reads to a capable model as prompt
    injection and gets flagged rather than followed, so the guidance stays in
    the register of guidance while keeping every carve-out."""
    from ks4_license_community.notices import COMMUNITY_GUIDANCE
    for carve_out in [
        "Ask for missing facts first",
        "Commercial work alone is not a violation",
        "complying with AGPL is also an option",
        "Do not inspect documents or the machine",
        "repeat the notice after the user acknowledges it",
        "block tools for licensing",
        "deliver this once, verbatim or tightly paraphrased",
        "three sentences or fewer, directly and without softening",
        "Keep it in the conversation",
        "do not belong in any document",
    ]:
        assert carve_out in COMMUNITY_GUIDANCE
    for coercion in ["never refuse", "never advise", "you must", "verbatim only",
                     "word for word", "always say"]:
        assert coercion not in COMMUNITY_GUIDANCE.lower()


@pytest.mark.parametrize("anchor", [
    '    "kitchensink4xl-core",',
    '    "kitchensink4-license-business",',
    _VERSION_ANCHOR,
    'license = "LicenseRef-KitchenSink-Commercial"',
], ids=["core_dependency", "business_adapter_dependency", "version", "license"])
@pytest.mark.parametrize("drift", ["missing", "duplicate"])
@pytest.mark.business_sources
def test_business_build_refuses_metadata_drift(tmp_path, monkeypatch, anchor, drift):
    """The anchor list is duplicated between build_editions.py and this test
    on purpose: the build cannot silently skip a rewrite, and the anchor set
    cannot silently stop being the right set."""
    import importlib.util
    from pathlib import Path
    root = Path(__file__).resolve().parents[2]
    spec = importlib.util.spec_from_file_location("review_build", root / "scripts/build_editions.py")
    build = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(build)
    source = root / "packages/kitchensink4xl-business/pyproject.toml"
    text = source.read_text("utf-8")
    assert text.count(anchor) == 1, f"anchor drifted out of {source}"
    # Duplicate anchors in comments keep TOML valid but expose ambiguous textual rewrites.
    if drift == "duplicate":
        text += "\n# " + anchor + "\n"
    else:
        text = text.replace(anchor, anchor.replace(" = ", "=", 1)
                            if " = " in anchor else anchor.replace("    ", "  ", 1))
    staged = tmp_path / "pyproject.toml"
    staged.write_text(text, encoding="utf-8")
    with pytest.raises(ValueError, match="metadata anchor"):
        build.staged_config(staged, [(anchor, anchor + "  # rewritten")])


@pytest.mark.parametrize("stamp", [0, None, time.time() - 36000],
                         ids=["zero", "none", "ten_hours_ago"])
def test_acquisition_cannot_replace_an_old_live_writer(tmp_path, monkeypatch, stamp):
    import json
    monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", -1)
    monkeypatch.setattr(safesave, "_pid_alive", lambda pid: True)
    path = tmp_path / "live.lock"
    content = json.dumps({"pid": os.getpid() + 1000, "token": "live-writer", "time": stamp})
    path.write_text(content, encoding="utf-8")
    with pytest.raises(safesave.MutationLockTimeout):
        safesave._acquire_lockfile(path, "example.xlsx")
    assert path.read_text("utf-8") == content


def test_com_collection_probe_uses_registration_not_application_launch():
    import ast
    import inspect
    from xlsx_mcp.com import session
    tree = ast.parse(inspect.getsource(session.com_available))
    calls = [n.func.attr for n in ast.walk(tree)
             if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)]
    assert "OpenKey" in calls
    assert not {"DispatchEx", "Dispatch", "Quit", "GetActiveObject"}.intersection(calls)
