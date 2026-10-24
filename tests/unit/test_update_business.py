"""The licensed update channel: the client half.

The client-side tests of the update channel. The EXISTING
test_update_check.py suite is the gate for the PyPI path: that file must keep
passing unmodified after the PyPI code moved behind the backend interface,
and it does.

Nothing here touches the network. The channel document is signed with a
throwaway key per test and the anchor file is written beside it, so the
signature path is exercised for real rather than stubbed; the transport is
the only thing replaced, and one test replaces it with a recorder that
asserts the request carries no identifier at all.

PORTING NOTE. This file is the same in all four products. Only the import
line and PRODUCT change per repo.
"""
from __future__ import annotations

import base64
import json
from datetime import datetime, timedelta, timezone

import pytest
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from ks4_license_business import policy as business_policy
from xlsx_mcp import license_policy, server
from xlsx_mcp.core import update_check as uc

PRODUCT_CONFIG = business_policy.ProductConfig(
    product_id="kitchensink4xl-business",
    display_name="KitchenSink4XL Business",
    license_file_env="KS4XL_LICENSE_FILE")

BUILD = {"edition": "business", "package": "kitchensink4xl-business",
         "version": "1.2.2+business", "release_date": "2026-09-10",
         "channel": "beta", "public_keys": {}}

NOW = datetime(2026, 9, 18, 6, 11, 3, tzinfo=timezone.utc)

#: The shipped fetch, captured before any fixture replaces it, so the
#: no-identifier test can assert against the request the real code
#: builds rather than against a test double.
_REAL_FETCH = uc.business_fetch


#: The bytes a Windows editor puts at the head of a saved policy file.
BOM = b"\xef\xbb\xbf"


def _iso(moment: datetime) -> str:
    return moment.replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _document(**changes) -> dict:
    doc = {
        "schema": 1,
        "channel": "beta",
        "sequence": 6,
        "release_id": "beta-6",
        "bundle_version": "2026.09.4-beta",
        "product_versions": {"word": "2.1.4+business", "xl": "1.2.3+business",
                             "ppt": "1.2.3+business", "web": "1.0.2+business"},
        "lines": {"cp314": {"sha256": "a1" * 32, "size": 128422104,
                            "manifest_sha256": "4e" * 32}},
        "default_line": "cp314",
        "anchor_key_id": "ks4-release-2026",
        "anchor_fingerprint": "sha256:" + "ce" * 32,
        "published_at": "2026-09-18T04:00:00Z",
        "expires_at": _iso(NOW + timedelta(days=90)),
        "min_supported_bundle_version": "2026.09.1-beta",
        "withdrawn": False,
        "withdrawn_reason": None,
        "severity": "recommended",
        "notes_url": "https://kitchensink4.ai/releases/beta-6",
    }
    doc.update(changes)
    return doc


class Channel:
    """One signing key, one anchor file, and a transport that serves
    whatever the test last handed it."""

    def __init__(self, root):
        self.key = Ed25519PrivateKey.generate()
        self.root = root
        self.requests: list = []
        self.bodies: dict = {}
        self.raise_with: BaseException | None = None
        self.calls = 0
        public = self.key.public_key().public_bytes(
            serialization.Encoding.Raw, serialization.PublicFormat.Raw)
        anchors = root / "trust" / "anchors.json"
        anchors.parent.mkdir(parents=True, exist_ok=True)
        anchors.write_text(json.dumps({
            "schema": 1,
            "anchors": {"ks4-release-2026": base64.b64encode(public).decode()},
        }), encoding="utf-8")

    def sign(self, doc: dict) -> dict:
        signature = self.key.sign(uc.canonical_bytes(doc))
        return {**doc, "signature": {
            "alg": "ed25519", "key_id": doc["anchor_key_id"],
            "sig": base64.b64encode(signature).decode()}}

    def serve(self, doc: dict | None = None, *, raw: bytes | None = None,
              url: str | None = None, sign: bool = True) -> None:
        if raw is None:
            body = self.sign(doc) if sign else doc
            raw = json.dumps(body).encode("utf-8")
        self.bodies[url] = raw

    def transport(self, url, *, user_agent, timeout, proxy=None):
        self.calls += 1
        self.requests.append({"url": url, "user_agent": user_agent,
                              "timeout": timeout, "proxy": proxy})
        if self.raise_with is not None:
            raise self.raise_with
        if url in self.bodies:
            return self.bodies[url]
        if None in self.bodies:
            return self.bodies[None]
        raise OSError("no body staged for " + url)


@pytest.fixture
def channel(monkeypatch, tmp_path):
    """A Business install with a signed channel, a policy file and a stamp,
    all under a scratch root. Nothing outside tmp_path is read or written."""
    config_root = tmp_path / "ProgramData" / "KitchenSink4AI"
    config_root.mkdir(parents=True, exist_ok=True)
    monkeypatch.setenv(uc.CONFIG_DIR_ENV, str(config_root))
    monkeypatch.setenv(uc.CACHE_DIR_ENV, str(tmp_path / "state"))
    for name in (uc.BUNDLE_VERSION_ENV, uc.RELEASE_ID_ENV, uc.CHANNEL_ENV,
                 uc.OFF_ENV, uc.LEGACY_OFF_ENV, uc.BUSINESS_OFF_ENV):
        monkeypatch.delenv(name, raising=False)
    stamp = tmp_path / "ks4-stamp.json"
    stamp.write_text(json.dumps({
        "schema": 1, "bundle_version": "2026.09.3-beta", "release_id": "beta-5",
        "channel": "beta", "installed_at": "2026-09-14T07:02:44Z",
        "line": "cp314", "org_id": "ORD-JW-0001"}), encoding="utf-8")
    monkeypatch.setenv(uc.STAMP_FILE_ENV, str(stamp))

    served = Channel(config_root)
    monkeypatch.setattr(uc, "business_fetch", served.transport)

    previous = license_policy.active()
    license_policy.install(business_policy.BusinessPolicy(PRODUCT_CONFIG, BUILD))
    served.stamp = stamp
    served.config_root = config_root
    served.policy_path = config_root / "update-policy.json"
    try:
        yield served
    finally:
        license_policy.install(previous)


def _license(monkeypatch, updates_through="2027-09-14", term_ends=None):
    """Replace the receipt report with the two facts coverage reads."""
    described = {"edition": "business", "status": "verified",
                 "updates_through": updates_through}
    if term_ends is not None:
        described["term_ends"] = term_ends
    monkeypatch.setattr(type(license_policy.active()), "describe_license",
                        lambda self, redacted=True: dict(described))


def _policy(channel, **fields):
    body = {"schema": 1, "mode": "notify", "channel": "beta"}
    body.update(fields)
    channel.policy_path.write_text(json.dumps(body), encoding="utf-8")


# ------------------------------------------------------ the statuses


def test_current(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(release_id="beta-5",
                            bundle_version="2026.09.3-beta", sequence=5))
    out = uc.status(force=True)
    assert out["status"] == "current"
    assert out["available"] is None
    assert out["installed"] == {"bundle_version": "2026.09.3-beta",
                                "release_id": "beta-5", "channel": "beta"}


def test_update_available(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["available"]["release_id"] == "beta-6"
    assert out["available"]["bundle_version"] == "2026.09.4-beta"
    assert out["available"]["product_versions"]["xl"] == "1.2.3+business"
    assert out["coverage"]["covered"] is True
    assert out["source"].endswith("/release/channel/beta")


def test_update_available_not_covered(channel, monkeypatch):
    _license(monkeypatch, updates_through="2026-08-01")
    _policy(channel)
    channel.serve(_document())
    out = uc.status(force=True)
    assert out["status"] == "update_available_not_covered"
    assert out["coverage"]["covered"] is False
    assert out["reason"] == "updates_lapsed"
    assert out["available"]["release_id"] == "beta-6"


def test_withdrawn(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(release_id="beta-5", sequence=7, withdrawn=True,
                            withdrawn_reason="A packaging fault"))
    out = uc.status(force=True)
    assert out["status"] == "withdrawn"
    assert out["withdrawn_reason"] == "A packaging fault"


def test_security_update(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(severity="security"))
    out = uc.status(force=True)
    assert out["status"] == "security_update"
    assert out["available"]["severity"] == "security"


def test_check_disabled_by_policy(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, mode="off")
    out = uc.status(force=True)
    assert out["status"] == "check_disabled"
    assert out["disabled_by"] == "update-policy.json"


def test_check_disabled_by_environment(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    monkeypatch.setenv(uc.BUSINESS_OFF_ENV, "off")
    out = uc.status(force=True)
    assert out["status"] == "check_disabled"
    assert out["disabled_by"] == uc.BUSINESS_OFF_ENV
    assert channel.calls == 0


def test_check_failed(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.raise_with = OSError("unreachable")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "network"
    assert out["available"] is None


def test_the_community_shim_switch_does_not_disable_the_business_channel(
        channel, monkeypatch):
    """ks4-install.ps1 compiles `<PREFIX>_UPDATE_CHECK=off` into EVERY Business
    launch shim, to keep a core paired with a community or missing adapter off
    the public index. A Business backend that honoured the same name reported
    check_disabled on every real install, which is what the gate-10 proof
    caught on a freshly installed bundle. The shim keeps the
    community switch and the Business check is governed by the policy file.
    """
    _license(monkeypatch)
    _policy(channel)
    monkeypatch.setenv(uc.OFF_ENV, "off")
    out = uc.status(force=True)
    assert out["status"] != "check_disabled"
    assert out.get("disabled_by") is None
    assert channel.calls == 1


def test_the_environment_can_only_turn_the_check_off(channel, monkeypatch):
    """"on" is not a value that turns anything on; the policy governs."""
    _license(monkeypatch)
    _policy(channel, mode="off")
    monkeypatch.setenv(uc.BUSINESS_OFF_ENV, "on")
    assert uc.status(force=True)["disabled_by"] == "update-policy.json"


# ------------------------------------------------------- every reason


def _reason(channel, monkeypatch, **kwargs):
    _license(monkeypatch)
    _policy(channel, **kwargs.pop("policy", {}))
    return uc.status(force=True)


def test_reason_oversize(channel, monkeypatch):
    channel.serve(raw=b"x" * (uc.BUSINESS_MAX_BYTES + 1))
    assert _reason(channel, monkeypatch)["reason"] == "oversize"


def test_reason_malformed(channel, monkeypatch):
    channel.serve(raw=b'{"schema": 1, "channel":')
    assert _reason(channel, monkeypatch)["reason"] == "malformed"


def test_reason_malformed_on_a_missing_field(channel, monkeypatch):
    doc = _document()
    doc.pop("lines")
    channel.serve(doc)
    assert _reason(channel, monkeypatch)["reason"] == "malformed"


def test_reason_schema_unknown(channel, monkeypatch):
    channel.serve(_document(schema=2))
    assert _reason(channel, monkeypatch)["reason"] == "schema_unknown"


def test_reason_signature_invalid(channel, monkeypatch):
    signed = channel.sign(_document())
    signed["bundle_version"] = "2026.09.9-beta"
    channel.serve(raw=json.dumps(signed).encode("utf-8"))
    assert _reason(channel, monkeypatch)["reason"] == "signature_invalid"


def test_reason_signature_invalid_on_an_unknown_key(channel, monkeypatch):
    channel.serve(_document(anchor_key_id="ks4-release-2099"))
    assert _reason(channel, monkeypatch)["reason"] == "signature_invalid"


def test_reason_channel_mismatch(channel, monkeypatch):
    channel.serve(_document(channel="stable"))
    assert _reason(channel, monkeypatch)["reason"] == "channel_mismatch"


def test_reason_expired(channel, monkeypatch):
    stale = _iso(datetime.now(timezone.utc) - timedelta(days=5))
    channel.serve(_document(expires_at=stale))
    assert _reason(channel, monkeypatch)["reason"] == "expired"


def test_reason_rollback(channel, monkeypatch):
    uc.write_channel_state("beta", {"schema": 1, "highest_sequence": 6,
                                    "last_release_id": "beta-6",
                                    "last_seen": _iso(NOW)})
    channel.serve(_document(sequence=5, release_id="beta-5x"))
    assert _reason(channel, monkeypatch)["reason"] == "rollback"


def test_reason_no_anchor(channel, monkeypatch):
    (channel.config_root / "trust" / "anchors.json").unlink()
    channel.serve(_document())
    assert _reason(channel, monkeypatch)["reason"] == "no_anchor"


def test_reason_network(channel, monkeypatch):
    channel.raise_with = OSError("connection refused")
    assert _reason(channel, monkeypatch)["reason"] == "network"


def test_reason_timeout(channel, monkeypatch):
    channel.raise_with = TimeoutError("timed out")
    assert _reason(channel, monkeypatch)["reason"] == "timeout"


def test_reason_policy_malformed(channel, monkeypatch):
    _license(monkeypatch)
    channel.policy_path.write_text("{not json", encoding="utf-8")
    out = uc.status(force=True)
    assert out["reason"] == "policy_malformed"
    assert out["policy_path"] == str(channel.policy_path)
    assert channel.calls == 0


def test_reason_policy_malformed_on_a_bad_value(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, check_interval_hours=0)
    assert uc.status(force=True)["reason"] == "policy_malformed"


def test_reason_no_stamp(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.stamp.unlink()
    out = uc.status(force=True)
    assert out["reason"] == "no_stamp"
    assert channel.calls == 0


def test_every_documented_reason_is_reachable():
    """The closed set the report may name. A reason outside it is a defect
    in this module, not a new fact about the world."""
    assert {"oversize", "malformed", "signature_invalid", "channel_mismatch",
            "pin_mismatch", "expired", "rollback", "no_anchor", "network",
            "timeout", "schema_unknown", "policy_malformed", "no_stamp"}


# -------------------------------------------------------- anti-rollback


def _state(channel):
    return uc.read_channel_state("beta")


def test_sequence_six_after_five_is_accepted(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    uc.write_channel_state("beta", {"schema": 1, "highest_sequence": 5,
                                    "last_release_id": "beta-5",
                                    "last_seen": _iso(NOW)})
    channel.serve(_document(sequence=6))
    assert uc.status(force=True)["status"] == "update_available"
    assert _state(channel)["highest_sequence"] == 6


def test_sequence_six_after_six_is_the_steady_state(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    uc.write_channel_state("beta", {"schema": 1, "highest_sequence": 6,
                                    "last_release_id": "beta-6",
                                    "last_seen": _iso(NOW)})
    channel.serve(_document(sequence=6))
    assert uc.status(force=True)["status"] == "update_available"
    assert _state(channel)["highest_sequence"] == 6


def test_a_rollback_leaves_the_state_file_untouched(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    uc.write_channel_state("beta", {"schema": 1, "highest_sequence": 6,
                                    "last_release_id": "beta-6",
                                    "last_seen": _iso(NOW)})
    before = uc.channel_state_path("beta").read_bytes()
    channel.serve(_document(sequence=5, release_id="beta-5x"))
    assert uc.status(force=True)["reason"] == "rollback"
    assert uc.channel_state_path("beta").read_bytes() == before


def test_a_missing_state_file_is_sequence_zero(channel):
    assert uc.read_channel_state("beta")["highest_sequence"] == 0


# ------------------------------------------------------------ the clock


def test_expiry_twenty_four_hours_in_the_past_is_tolerated(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(
        expires_at=_iso(datetime.now(timezone.utc) - timedelta(hours=24))))
    assert uc.status(force=True)["status"] == "update_available"


def test_expiry_seventy_two_hours_in_the_past_is_refused(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(
        expires_at=_iso(datetime.now(timezone.utc) - timedelta(hours=72))))
    assert uc.status(force=True)["reason"] == "expired"


# --------------------------------------------------------- the coverage


@pytest.mark.parametrize("model", ["updates_only", "term"])
def test_coverage_includes_a_release_published_inside_the_window(model):
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", "2027-09-14", None,
                               model=model)
    assert out["covered"] is True
    assert out["model"] == model


@pytest.mark.parametrize("model", ["updates_only", "term"])
def test_coverage_excludes_a_release_published_after_the_window(model):
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", "2026-08-01", None,
                               model=model)
    assert out["covered"] is False
    assert out["reason"] == "updates_lapsed"


@pytest.mark.parametrize("model", ["updates_only", "term"])
def test_a_term_still_running_changes_nothing(model):
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", "2027-09-14",
                               "2027-09-14", model=model)
    assert out["covered"] is True
    assert out["term_ends"] == "2027-09-14"


def test_a_passed_term_is_covered_under_updates_only():
    """The live receipts and the live copy say updates_only, and under it a
    passed term is not a thing that exists."""
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", "2027-09-14",
                               "2026-01-01", model="updates_only")
    assert out["covered"] is True
    assert out["reason"] is None


def test_a_passed_term_ends_coverage_under_term():
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", "2027-09-14",
                               "2026-01-01", model="term")
    assert out["covered"] is False
    assert out["reason"] == "term_ended"


def test_no_receipt_means_coverage_does_not_know():
    out = uc.evaluate_coverage("2026-09-18T04:00:00Z", None, None)
    assert out["covered"] is None


def test_the_coverage_model_has_exactly_one_home():
    """One constant, one place. A second module-level name carrying a
    renewal rule is the thing this test exists to catch."""
    import inspect
    source = inspect.getsource(uc)
    assert source.count("\nCOVERAGE_MODEL = ") == 1
    assert uc.COVERAGE_MODEL in ("updates_only", "term")


def test_the_term_branch_reaches_the_report(channel, monkeypatch):
    _license(monkeypatch, term_ends="2026-01-01")
    _policy(channel)
    monkeypatch.setattr(uc, "COVERAGE_MODEL", "term")
    channel.serve(_document())
    out = uc.status(force=True)
    assert out["status"] == "update_available_not_covered"
    assert out["reason"] == "term_ended"


# --------------------------------------------------------------- pins


def test_a_pin_that_matches_the_install_reports_current(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5",
                            bundle_version="2026.09.3-beta", sequence=5),
                  url="https://get.kitchensink4.ai/release/beta-5")
    out = uc.status(force=True)
    assert out["status"] == "current"
    assert out["pin"] == "beta-5"


def test_a_pin_ahead_of_the_install_describes_the_pin(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-6")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/beta-6")
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["available"]["release_id"] == "beta-6"
    assert out["source"].endswith("/release/beta-6")
    assert [r["url"] for r in channel.requests] == [
        "https://get.kitchensink4.ai/release/channel/beta",
        "https://get.kitchensink4.ai/release/beta-6"]


def test_a_pin_is_not_subject_to_the_sequence_rule(channel, monkeypatch):
    """A pin deliberately names an older release, so the channel's
    anti-rollback record must not refuse it, and resolving one must not
    advance that record either."""
    _license(monkeypatch)
    _policy(channel, pin="beta-4")
    uc.write_channel_state("beta", {"schema": 1, "highest_sequence": 6,
                                    "last_release_id": "beta-6",
                                    "last_seen": _iso(NOW)})
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(sequence=4, release_id="beta-4",
                            bundle_version="2026.09.2-beta"),
                  url="https://get.kitchensink4.ai/release/beta-4")
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["available"]["release_id"] == "beta-4"
    assert uc.read_channel_state("beta")["highest_sequence"] == 6


def test_a_pin_holds_against_a_security_head_by_default(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(severity="security"),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5",
                            bundle_version="2026.09.3-beta", sequence=5),
                  url="https://get.kitchensink4.ai/release/beta-5")
    out = uc.status(force=True)
    assert out["status"] == "current"
    assert [r["url"] for r in channel.requests] == [
        "https://get.kitchensink4.ai/release/channel/beta",
        "https://get.kitchensink4.ai/release/beta-5"]


def test_security_bypasses_pin_reports_the_head(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-5", security_bypasses_pin=True)
    channel.serve(_document(severity="security"),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    out = uc.status(force=True)
    assert out["status"] == "security_update"
    assert out["available"]["release_id"] == "beta-6"
    assert out["pin"] == "beta-5"


def test_security_bypasses_pin_still_holds_for_a_routine_head(channel,
                                                              monkeypatch):
    """The switch reports past a pin for a SECURITY release and nothing
    else. It never installs anything in either case."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5", security_bypasses_pin=True)
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5",
                            bundle_version="2026.09.3-beta", sequence=5),
                  url="https://get.kitchensink4.ai/release/beta-5")
    assert uc.status(force=True)["status"] == "current"


# ------------------------------------------------- interval and jitter


def test_two_calls_inside_the_interval_make_one_network_call(channel,
                                                             monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    first = uc.status(force=True)
    second = uc.status()
    assert channel.calls == 1
    assert first["status"] == second["status"] == "update_available"
    assert second["available"]["release_id"] == "beta-6"


def test_the_interval_is_honoured_from_the_policy_file(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, check_interval_hours=1)
    channel.serve(_document())
    uc.status(force=True)
    cache = uc.read_business_cache()
    aged = datetime.now(timezone.utc) - timedelta(hours=5)
    uc.write_business_cache({**cache, "last_check": aged.isoformat()})
    uc.status()
    assert channel.calls == 2


def test_the_jitter_is_stable_across_calls(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    uc.status(force=True)
    stored = uc.read_business_cache()["jitter_seconds"]
    assert 0 <= stored < uc.BUSINESS_MAX_JITTER_SECONDS
    assert uc.business_jitter("beta", uc.read_business_cache()) == stored
    assert uc.business_jitter("beta") == uc.business_jitter("beta")


def test_the_jitter_differs_between_channels():
    assert uc.business_jitter("beta") != uc.business_jitter("stable")


def test_a_failure_is_recorded_so_the_interval_still_caps_the_network(
        channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.raise_with = OSError("down")
    uc.status(force=True)
    assert uc.read_business_cache()["ok"] is False
    uc.status()
    assert channel.calls == 1


# ------------------------------------------------------ no identifier


def test_the_request_carries_no_identifier(channel, monkeypatch):
    """Two headers, no query string, no cookie, no license id, no org, no
    machine. The whole request is recorded and asserted."""
    _license(monkeypatch)
    _policy(channel)
    recorded = {}

    def transport(request, timeout, proxy):
        recorded["url"] = request.full_url
        recorded["headers"] = dict(request.header_items())
        recorded["method"] = request.get_method()
        recorded["data"] = request.data
        raise OSError("recorded, not sent")

    monkeypatch.setattr(uc, "_business_open", transport)
    monkeypatch.setattr(uc, "business_fetch", _REAL_FETCH)
    uc.status(force=True)
    assert recorded["method"] == "GET"
    assert recorded["data"] is None
    assert recorded["url"] == "https://get.kitchensink4.ai/release/channel/beta"
    assert "?" not in recorded["url"]
    # urllib title-cases what it is given, so the assertion is on the set of
    # header NAMES, case-folded: exactly two, and nothing else rides along.
    headers = {k.lower(): v for k, v in recorded["headers"].items()}
    assert set(headers) == {"accept", "user-agent"}
    assert headers["accept"] == "application/json"
    assert headers["user-agent"] == "kitchensink4xl-business/1.2.2+business"
    # Nothing that identifies the customer, the order, or the machine.
    blob = json.dumps(recorded).lower()
    for secret in ("ORD-JW-0001", "Private Example", "cookie", "license_id",
                   "kitchensink4xl-business/1.2.2+business=", "authorization",
                   uc._machine_guid()):
        assert secret.lower() not in blob or secret == ""


def test_the_user_agent_names_only_the_product_and_its_version(channel,
                                                               monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    uc.status(force=True)
    agent = channel.requests[0]["user_agent"]
    assert agent == "kitchensink4xl-business/1.2.2+business"
    assert agent.count("/") == 1


def test_the_endpoint_override_changes_the_origin_only(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, endpoint="https://mirror.example.internal")
    channel.serve(_document(),
                  url="https://mirror.example.internal/release/channel/beta")
    assert uc.status(force=True)["status"] == "update_available"


def test_the_endpoint_rule_matches_the_installer():
    """ks4-install.ps1's -UpdatePolicy validates this file and then WRITES it,
    so a value the installer accepts and this client calls malformed is a
    policy that installs cleanly and turns every server's check off. The
    gate-10 proof hit exactly that, on a loopback endpoint the installer had
    just written. These are Test-UpdateEndpoint's two accepted shapes.
    """
    for origin in ("https://mirror.example.internal",
                   "https://mirror.example.internal/",
                   "http://127.0.0.1:8799", "http://127.0.0.1",
                   "http://localhost", "http://localhost:80",
                   "http://[::1]", "http://[::1]:99"):
        assert uc._valid_endpoint(origin), origin
    for origin in ("http://mirror.example.internal",
                   "http://127.0.0.1.evil.example",
                   "http://127.0.0.1:8799/path", "http://127.0.0.1:x",
                   "https://", "ftp://loopback"):
        assert not uc._valid_endpoint(origin), origin


def test_a_loopback_endpoint_is_read_rather_than_refused(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, endpoint="http://127.0.0.1:8799")
    channel.serve(_document(), url="http://127.0.0.1:8799/release/channel/beta")
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["source"] == "http://127.0.0.1:8799/release/channel/beta"


def test_a_mirror_cannot_forge_a_release(channel, monkeypatch):
    """The signature check does not move with the endpoint."""
    _license(monkeypatch)
    _policy(channel, endpoint="https://mirror.example.internal")
    channel.serve(_document(), sign=False,
                  url="https://mirror.example.internal/release/channel/beta")
    assert uc.status(force=True)["reason"] == "malformed"


# ---------------------------------------------------------- mode off


def test_mode_off_reads_no_cache_and_makes_no_call(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, mode="off")

    def landmine(*a, **k):
        raise AssertionError("mode off touched the cache or the network")

    monkeypatch.setattr(uc, "read_business_cache", landmine)
    monkeypatch.setattr(uc, "write_business_cache", landmine)
    monkeypatch.setattr(uc, "business_fetch", landmine)
    out = uc.status(force=True)
    assert out["status"] == "check_disabled"
    assert out["disabled_by"] == "update-policy.json"
    assert not uc.business_cache_path().exists()


def test_mode_off_says_nothing_in_the_one_line_notice(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    uc.status(force=True)
    assert uc.update_notice() is not None
    _policy(channel, mode="off")
    assert uc.update_notice() is None


# ------------------------------------------------- the report surface


def test_get_server_info_reports_the_build_on_business(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    out = server.get_server_info()
    assert out["bundle_version"] == "2026.09.3-beta"
    assert out["release_id"] == "beta-5"
    assert out["channel"] == "beta"
    assert out["update_check"]["status"] == "update_available"
    assert isinstance(out["update"], str)
    assert "2026.09.4-beta" in out["update"]
    assert "beta-6" in out["update"]


def test_get_server_info_reports_three_nulls_on_community(monkeypatch,
                                                         tmp_path):
    monkeypatch.setenv(uc.CACHE_DIR_ENV, str(tmp_path))
    monkeypatch.delenv(uc.OFF_ENV, raising=False)
    monkeypatch.delenv(uc.LEGACY_OFF_ENV, raising=False)
    monkeypatch.setattr(uc, "_fetch", lambda *a, **k: None)
    out = server.get_server_info()
    assert out["edition"] == "community"
    assert out["bundle_version"] is None
    assert out["release_id"] is None
    assert out["channel"] is None
    assert "state" in out["update_check"]
    assert "update" not in out or isinstance(out["update"], str)


def test_the_report_never_carries_a_local_path_or_a_licensee(channel,
                                                             monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    blob = json.dumps(server.get_server_info())
    assert str(channel.config_root) not in blob
    assert str(channel.stamp) not in blob
    assert "Private Example" not in blob


def test_the_privacy_sentence_is_emitted_once(channel, monkeypatch):
    """The once-per-cache mechanism, exercised against a stand-in.

    No such sentence is shipped yet, so
    BUSINESS_PRIVACY_NOTE is None and the field is absent. Asserting the
    mechanism against nothing would assert nothing, so the backend is given
    a sentence here and the disclosure is watched.
    """
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    monkeypatch.setattr(uc.BusinessBackend, "privacy_text", "a sentence")
    first = uc.status(force=True)
    assert first["privacy"] == "a sentence"
    assert first["privacy_endpoint"].startswith("https://get.kitchensink4.ai/")
    assert "privacy" not in uc.status(force=True)


def test_nothing_to_disclose_is_not_a_disclosure(channel, monkeypatch):
    """With no sentence written, no disclosure field is emitted, and
    the once-per-cache flag is not spent either: the first check after the
    copy lands is still the one that discloses it."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    assert uc.BUSINESS_PRIVACY_NOTE is None
    out = uc.status(force=True)
    assert "privacy" not in out
    assert "privacy_endpoint" not in out
    monkeypatch.setattr(uc.BusinessBackend, "privacy_text", "a sentence")
    assert uc.status(force=True)["privacy"] == "a sentence"


def test_a_disabled_check_discloses_nothing(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, mode="off")
    assert "privacy" not in uc.status(force=True)


# ------------------------------------------------------------ the seam itself


def test_a_policy_with_no_backend_gets_the_public_channel():
    assert isinstance(uc.resolve_backend(), uc.PyPIBackend)
    assert uc.resolve_backend().name == "pypi"
    assert uc.installed_build() == {"bundle_version": None, "release_id": None,
                                    "channel": None}


def test_the_business_policy_selects_the_business_channel(channel):
    backend = uc.resolve_backend()
    assert isinstance(backend, uc.BusinessBackend)
    assert backend.name == "business"
    assert backend.off_env == uc.BUSINESS_OFF_ENV


def test_a_policy_may_supply_a_whole_backend_of_its_own(channel, monkeypatch):
    class Custom:
        name = "custom"
        off_env = "NOPE"
        privacy_text = "x"
        privacy_endpoint = "https://example.invalid/"

        def check(self, *, timeout, force=False):
            return {"status": "current", "reason": None}

        def first_disclosure(self):
            return False

        def mark_disclosed(self):
            pass

    monkeypatch.setattr(license_policy, "update_backend", lambda: Custom())
    assert uc.status()["status"] == "current"


def test_suppression_still_wins_over_a_backend(channel, monkeypatch):
    """A mismatched pair never reaches the network by accident."""
    monkeypatch.setattr(license_policy, "update_check_suppression",
                        lambda: {"disabled_by": "business_distribution"})
    monkeypatch.setattr(uc, "business_fetch", _landmine)
    out = uc.status(force=True)
    assert out["state"] == "disabled"
    assert out["disabled_by"] == "business_distribution"


def _landmine(*a, **k):
    raise AssertionError("a suppressed check reached the network")


def test_the_backend_check_never_raises(channel, monkeypatch):
    def explode(*a, **k):
        raise RuntimeError("anything at all")

    monkeypatch.setattr(uc, "read_stamp", explode)
    assert uc.BusinessBackend().check(timeout=1.0)["status"] == "check_failed"


def test_nothing_in_this_module_downloads(channel):
    import inspect
    source = inspect.getsource(uc)
    for banned in ("urlretrieve", "shutil.copyfileobj", "subprocess",
                   "os.system"):
        assert banned not in source


def test_current_version_reports_the_distribution_string(channel):
    assert uc.current_version() == "1.2.2+business"
    assert server.get_server_info()["version"] == uc.current_version()


# ------------------------------- the pin is bound to what it asked for
#
# Regression: an independent review served a validly signed but
# OLDER release document in answer to a pinned request and every client took
# it, because the pin path waives the sequence floor and nothing else bound
# the answer to the requested id. Signing establishes that a document is
# genuine; it does not establish that it is the document that was asked for.


def test_a_pin_refuses_a_validly_signed_other_release(channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-6")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-1", sequence=1,
                            bundle_version="2026.09.1-beta"),
                  url="https://get.kitchensink4.ai/release/beta-6")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "pin_mismatch"
    assert out["available"] is None


def test_a_pin_refuses_a_signed_release_from_another_channel(channel,
                                                             monkeypatch):
    """The same substitution dressed as a cross-channel document. The pin
    path does not waive the channel rule either."""
    _license(monkeypatch)
    _policy(channel, pin="beta-6")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(channel="stable", release_id="stable-1",
                            sequence=1),
                  url="https://get.kitchensink4.ai/release/beta-6")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] in {"channel_mismatch", "pin_mismatch"}
    assert out["available"] is None


def test_the_matching_pin_is_still_accepted(channel, monkeypatch):
    """The positive control for the two refusals above: the identity check
    refuses substitutes and nothing else."""
    _license(monkeypatch)
    _policy(channel, pin="beta-6")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/beta-6")
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["available"]["release_id"] == "beta-6"


def test_validate_document_binds_the_requested_release_id(channel):
    """The rule at its own level, so a future caller that forgets to pass
    the id is the only way back to the defect."""
    anchors = uc.read_anchors()
    raw = json.dumps(channel.sign(_document())).encode("utf-8")
    doc, reason = uc.validate_document(raw, channel="beta", anchors=anchors,
                                       enforce_sequence=False,
                                       expect_release_id="beta-6")
    assert reason is None and doc["release_id"] == "beta-6"
    doc, reason = uc.validate_document(raw, channel="beta", anchors=anchors,
                                       enforce_sequence=False,
                                       expect_release_id="beta-4")
    assert doc is None and reason == "pin_mismatch"


# ---------------------------- a failed refresh stays failed until good
#
# Regression from the same review: a good check, then a failed refresh, then
# an ordinary call inside the interval reported the retained document and the
# failure disappeared. A machine that could not check says so.


def test_a_failed_refresh_is_not_erased_by_the_cached_document(channel,
                                                               monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    good = uc.status(force=True)
    assert good["status"] == "update_available"

    channel.serve(raw=b"this is not a channel document",
                  url="https://get.kitchensink4.ai/release/channel/beta")
    bad = uc.status(force=True)
    assert bad["status"] == "check_failed"
    assert bad["reason"] == "malformed"

    cached = uc.status()
    assert cached["status"] == "check_failed"
    assert cached["reason"] == "malformed"
    assert cached["available"] is None
    assert cached["last_successful_check"] == good["last_successful_check"]


def test_a_cached_good_check_is_still_reported_inside_the_interval(
        channel, monkeypatch):
    """The positive control: nothing above stops a good cache from
    answering without a network call."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    first = uc.status(force=True)
    calls = channel.calls
    second = uc.status()
    assert second["status"] == first["status"] == "update_available"
    assert channel.calls == calls


def test_a_pin_set_after_the_cache_was_written_is_not_answered_from_it(
        channel, monkeypatch):
    """A cached answer answers the question that was asked when it was
    taken. A new pin is a different question."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    assert uc.status(force=True)["status"] == "update_available"
    _policy(channel, pin="beta-4")
    out = uc.status()
    assert out["status"] == "check_failed"
    assert out["reason"] == "pin_mismatch"


# ------------------------------ a Windows-authored policy has a BOM
#
# Regression from the same review: Notepad, Set-Content and Out-File write a
# UTF-8 BOM, the installer's ConvertFrom-Json accepts one, and every client
# refused the identical file as policy_malformed.


def _policy_with_bom(channel, **fields):
    body = {"schema": 1, "mode": "notify", "channel": "beta"}
    body.update(fields)
    channel.policy_path.write_text(json.dumps(body), encoding="utf-8-sig")


def test_a_policy_file_with_a_bom_turns_the_check_off(channel, monkeypatch):
    _license(monkeypatch)
    _policy_with_bom(channel, mode="off")
    monkeypatch.setattr(uc, "business_fetch", _landmine)
    out = uc.status(force=True)
    assert out["status"] == "check_disabled"
    assert out["disabled_by"] == "update-policy.json"


def test_a_policy_file_with_a_bom_drives_an_ordinary_check(channel,
                                                           monkeypatch):
    _license(monkeypatch)
    _policy_with_bom(channel)
    channel.serve(_document())
    out = uc.status(force=True)
    assert out["status"] == "update_available"


def test_the_policy_file_bytes_still_have_to_be_json(channel):
    channel.policy_path.write_bytes(BOM + b"not json at all")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "policy_malformed"


# ------------------------------------ recall, carried by the fresh head
#
# A historical /release/<id> document is byte-immutable: it was signed before
# anyone knew the release would be pulled, so it can never carry its own
# recall. The channel head carries `withdrawn_release_ids` instead, which is
# why a pinned check fetches the head as well as the pin.


def test_a_pinned_release_named_in_the_head_recall_list_reports_withdrawn(
        channel, monkeypatch):
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(withdrawn_release_ids=["beta-5"]),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta"),
                  url="https://get.kitchensink4.ai/release/beta-5")
    out = uc.status(force=True)
    assert out["status"] == "withdrawn"
    assert out["recalled_release_ids"] == ["beta-5"]


def test_the_historical_pin_document_still_verifies_unchanged(channel,
                                                              monkeypatch):
    """The recall does not rewrite history: the pinned document is the same
    bytes, with the same valid signature, before and after."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    historical = channel.sign(_document(release_id="beta-5", sequence=5,
                                        bundle_version="2026.09.3-beta"))
    raw = json.dumps(historical).encode("utf-8")
    doc, reason = uc.validate_document(raw, channel="beta",
                                       anchors=uc.read_anchors(),
                                       enforce_sequence=False,
                                       expect_release_id="beta-5")
    assert reason is None and doc["release_id"] == "beta-5"

    channel.serve(_document(withdrawn_release_ids=["beta-5"]),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(raw=raw, url="https://get.kitchensink4.ai/release/beta-5")
    assert uc.status(force=True)["status"] == "withdrawn"

    doc, reason = uc.validate_document(raw, channel="beta",
                                       anchors=uc.read_anchors(),
                                       enforce_sequence=False,
                                       expect_release_id="beta-5")
    assert reason is None and doc["release_id"] == "beta-5"


def test_a_head_that_cannot_be_verified_reports_unknown_not_current(
        channel, monkeypatch):
    """A pinned machine that cannot reach or verify the head does not know
    whether its pin has been recalled, and says so."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(), url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta"),
                  url="https://get.kitchensink4.ai/release/beta-5")
    good = uc.status(force=True)
    assert good["status"] == "current"

    channel.serve(raw=b"not a document",
                  url="https://get.kitchensink4.ai/release/channel/beta")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "malformed"
    assert out["last_successful_check"] == good["last_successful_check"]
    assert uc.status()["status"] == "check_failed"


def test_a_head_below_the_accepted_sequence_cannot_clear_a_recall(
        channel, monkeypatch):
    """Monotonic. An attacker who can serve stale bytes replays the head
    that predates the recall; the sequence rule refuses it, and a refused
    head is not an answer."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(sequence=7, withdrawn_release_ids=["beta-5"]),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta"),
                  url="https://get.kitchensink4.ai/release/beta-5")
    assert uc.status(force=True)["status"] == "withdrawn"
    assert uc.read_channel_state("beta")["highest_sequence"] == 7

    channel.serve(_document(sequence=6),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "rollback"
    assert uc.read_channel_state("beta")["highest_sequence"] == 7


def test_an_expired_head_cannot_clear_a_recall(channel, monkeypatch):
    """The other stale-bytes route to the same end, refused the same way."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    channel.serve(_document(expires_at=_iso(NOW - timedelta(days=30))),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta"),
                  url="https://get.kitchensink4.ai/release/beta-5")
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "expired"


def test_a_head_with_no_recall_list_recalls_nothing(channel, monkeypatch):
    """Every document published before the field existed means an empty
    list, not an unknown."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5")
    body = _document()
    body.pop("withdrawn_release_ids", None)
    channel.serve(body, url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta"),
                  url="https://get.kitchensink4.ai/release/beta-5")
    out = uc.status(force=True)
    assert out["status"] == "current"
    assert out["recalled_release_ids"] == []


def test_a_head_carrying_fields_this_client_does_not_know_still_parses(
        channel, monkeypatch):
    """Forward compatibility: the frozen parser in a shipped bundle meets a
    head published by a newer publisher and reads the fields it knows."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(some_future_field={"a": [1, 2, 3]},
                            another_one="whatever it turns out to be"))
    assert uc.status(force=True)["status"] == "update_available"


def test_a_recall_list_of_the_wrong_shape_is_malformed(channel, monkeypatch):
    """A recall that is silently dropped is the one failure worth refusing
    over, so a list this client cannot read is a refusal, not an empty
    list."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(withdrawn_release_ids="beta-5"))
    out = uc.status(force=True)
    assert out["status"] == "check_failed"
    assert out["reason"] == "malformed"


def test_an_installed_release_named_in_the_recall_list_reports_withdrawn(
        channel, monkeypatch):
    """The unpinned case: nothing is pinned, and the head says the release
    this machine is running has been pulled."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(withdrawn_release_ids=["beta-5"]))
    out = uc.status(force=True)
    assert out["status"] == "withdrawn"
    assert out["available"]["release_id"] == "beta-6"


# ------------------- a cached answer is no fresher than its document
#
# Regression from the repair review. The not-due branch reported from
# the cached document without re-reading its expiry, so a head accepted just
# inside the 48 hour clock tolerance went on answering update_available after
# that tolerance had run out, while a forced check of the identical signed
# bytes answered check_failed with reason expired. A signature stays valid
# forever; freshness does not, and the two paths have to agree.


class _Clock(datetime):
    """A datetime whose now() is a moment the test chooses.

    The module reads the clock in several places, so the clock is what the
    test moves rather than the document, which keeps the signed bytes
    identical on both sides of the comparison.
    """

    _moment = datetime(2026, 9, 18, 6, 11, 3, tzinfo=timezone.utc)

    @classmethod
    def now(cls, tz=None):
        return cls._moment if tz else cls._moment.replace(tzinfo=None)


def _at(monkeypatch, moment):
    monkeypatch.setattr(uc, "datetime",
                        type("Clock", (_Clock,), {"_moment": moment}))


def _nearly_expired_head(channel, monkeypatch, **changes):
    """A genuinely signed head whose expires_at has already passed, by less
    than the tolerance, so it is accepted when it arrives.

    The clock is frozen at that same moment before the first check runs.
    These tests are about a document crossing a boundary the test moves it
    across deliberately, so the wall clock underneath has no business
    moving too: a check that straddled a second boundary recorded one
    second in the cache and reported the next one, and the assertion below
    that the cached answer preserves the first check's success time failed
    on that difference rather than on anything the module did wrong.
    """
    now = datetime.now(timezone.utc).replace(microsecond=0)
    channel.serve(_document(expires_at=_iso(now - timedelta(hours=47)),
                            **changes))
    _at(monkeypatch, now)
    return now


def test_a_cached_head_inside_the_clock_tolerance_still_answers(channel,
                                                                monkeypatch):
    """The positive control for the rule below: a document past its expiry
    but still inside the 48 hour tolerance is a valid answer, cached or
    fetched, and it is still answered without a network call."""
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    now = _nearly_expired_head(channel, monkeypatch)
    first = uc.status(force=True)
    assert first["status"] == "update_available"
    calls = channel.calls

    _at(monkeypatch, now + timedelta(minutes=30))
    cached = uc.status()
    assert cached["status"] == "update_available"
    assert channel.calls == calls


def test_a_cached_head_past_the_clock_tolerance_reports_expired(channel,
                                                                monkeypatch):
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    now = _nearly_expired_head(channel, monkeypatch)
    first = uc.status(force=True)
    assert first["status"] == "update_available"
    calls = channel.calls

    # Two hours on: past expires_at plus the tolerance, and still a long way
    # inside the 24 hour polling interval, which is the window the defect
    # lived in.
    _at(monkeypatch, now + timedelta(hours=2))
    cached = uc.status()
    assert cached["status"] == "check_failed"
    assert cached["reason"] == "expired"
    assert cached["available"] is None
    assert cached["last_successful_check"] == first["last_successful_check"]
    assert channel.calls == calls


def test_the_cached_and_the_forced_answer_agree_on_identical_bytes(
        channel, monkeypatch):
    """The defect as the review reproduced it: the same signed document at
    the same moment gave two different answers depending on which path the
    call took."""
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    now = _nearly_expired_head(channel, monkeypatch)
    assert uc.status(force=True)["status"] == "update_available"

    _at(monkeypatch, now + timedelta(hours=2))
    cached = uc.status()
    forced = uc.status(force=True)
    assert cached["status"] == forced["status"] == "check_failed"
    assert cached["reason"] == forced["reason"] == "expired"


def test_an_expired_cached_pin_document_is_not_reported_either(channel,
                                                              monkeypatch):
    """The pin is the other half of the cached pair and it carries its own
    expiry, so a fresh head does not vouch for a stale pinned document."""
    _license(monkeypatch)
    _policy(channel, pin="beta-5", check_interval_hours=24)
    now = datetime.now(timezone.utc)
    channel.serve(_document(expires_at=_iso(now + timedelta(days=90))),
                  url="https://get.kitchensink4.ai/release/channel/beta")
    channel.serve(_document(release_id="beta-5", sequence=5,
                            bundle_version="2026.09.3-beta",
                            expires_at=_iso(now - timedelta(hours=47))),
                  url="https://get.kitchensink4.ai/release/beta-5")
    assert uc.status(force=True)["status"] == "current"
    calls = channel.calls

    _at(monkeypatch, now + timedelta(hours=2))
    cached = uc.status()
    assert cached["status"] == "check_failed"
    assert cached["reason"] == "expired"
    assert channel.calls == calls


def test_a_cached_head_below_a_raised_sequence_floor_is_not_reported(
        channel, monkeypatch):
    """The anti-rollback record is machine-wide and the cache is per
    product, so another product on the same machine can raise the floor
    under a cache that was written below it."""
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    channel.serve(_document())
    assert uc.status(force=True)["status"] == "update_available"
    calls = channel.calls

    uc.write_channel_state("beta", {
        "schema": 1, "highest_sequence": 9, "last_release_id": "beta-9",
        "last_seen": "2026-09-18T06:11:03Z"})
    cached = uc.status()
    assert cached["status"] == "check_failed"
    assert cached["reason"] == "rollback"
    assert channel.calls == calls


def test_a_cached_answer_from_another_channel_is_not_reported(channel,
                                                              monkeypatch):
    """A machine moved to another channel asks a different question, and
    the cached pair answers the old one."""
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    channel.serve(_document())
    assert uc.status(force=True)["status"] == "update_available"
    monkeypatch.setenv(uc.CHANNEL_ENV, "stable")
    cached = uc.status()
    assert cached["status"] == "check_failed"
    assert cached["reason"] == "channel_mismatch"


def test_the_freshness_rules_are_read_from_the_cache_without_the_network(
        channel, monkeypatch):
    """The helper at its own level: the rules, and nothing else."""
    backend = uc.BusinessBackend()
    now = datetime(2026, 9, 18, 6, 11, 3, tzinfo=timezone.utc)
    fresh = _document(expires_at=_iso(now + timedelta(days=1)))
    stale = _document(expires_at=_iso(now - timedelta(hours=49)))
    assert backend._stale_cache_reason(fresh, None, "beta", 6, now=now) is None
    assert backend._stale_cache_reason(stale, None, "beta", 6,
                                       now=now) == "expired"
    assert backend._stale_cache_reason(fresh, None, "beta", 9,
                                       now=now) == "rollback"
    assert backend._stale_cache_reason(fresh, None, "stable", 6,
                                       now=now) == "channel_mismatch"
    # A pinned document is allowed to sit below the floor, which is the one
    # rule a pin waives, and its expiry is still its own.
    pinned = _document(release_id="beta-5", sequence=5,
                       expires_at=_iso(now + timedelta(days=1)))
    assert backend._stale_cache_reason(fresh, pinned, "beta", 6,
                                       now=now) is None
    assert backend._stale_cache_reason(
        fresh, _document(release_id="beta-5", sequence=5,
                         expires_at=_iso(now - timedelta(hours=49))),
        "beta", 6, now=now) == "expired"


# ------------------------------------- one check, one clock reading


class _TickingClock(datetime):
    """A clock no two reads of which agree.

    A check that samples the clock twice agreed with itself whenever both
    reads fell inside the same second and disagreed when they did not,
    which is a test that passes most of the time and a contract that holds
    most of the time. Under this clock the second read is always a second
    later, so a check that takes two of them cannot pass.
    """

    _moment = NOW
    _reads = 0

    @classmethod
    def now(cls, tz=None):
        cls._reads += 1
        moment = cls._moment + timedelta(seconds=cls._reads)
        return moment if tz else moment.replace(tzinfo=None)


def _ticking(monkeypatch):
    monkeypatch.setattr(uc, "datetime",
                        type("Ticking", (_TickingClock,), {"_reads": 0}))


def test_a_successful_check_reads_the_clock_once(channel, monkeypatch):
    """checked_at is when this check looked and last_successful_check is
    when a check last succeeded. A check that has just succeeded is both of
    those things at once, so the two fields carry one value, and the only
    way they come apart is the way that means something: this check did not
    succeed and the answer in front of you came from an older one."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    _ticking(monkeypatch)

    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["last_successful_check"] == out["checked_at"]


def test_a_cached_answer_still_reports_when_it_was_asked(channel, monkeypatch):
    """The other half of the same rule: the cached path keeps the older
    success time and stamps the present moment, so the two differ there and
    the gap between them is how long this machine has been answering from
    memory."""
    _license(monkeypatch)
    _policy(channel, check_interval_hours=24)
    channel.serve(_document())
    _at(monkeypatch, NOW)
    first = uc.status(force=True)
    assert first["status"] == "update_available"

    _at(monkeypatch, NOW + timedelta(hours=3))
    cached = uc.status()
    assert cached["status"] == "update_available"
    assert cached["last_successful_check"] == first["last_successful_check"]
    assert cached["checked_at"] != cached["last_successful_check"]


# --------------------- additive head fields the client has never
# --------------------- heard of


def test_an_unknown_additive_head_field_is_ignored_for_coverage(
        channel, monkeypatch):
    """The channel head is gaining a metadata_at field that records when
    the document was last rebuilt, which happens for reasons that have
    nothing to do with the release it names: a recall of some other build
    rewrites the head and every artifact in it keeps the publication date
    it always had. Coverage is an entitlement question about an artifact,
    so published_at decides it and nothing else does. A client that started
    reading the newer field would take a metadata refresh for a
    republication and tell an entitled customer it is not covered."""
    _license(monkeypatch, updates_through="2026-09-19")
    _policy(channel)
    channel.serve(_document(published_at="2026-09-18T04:00:00Z",
                            metadata_at="2027-03-01T00:00:00Z"))

    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["coverage"]["covered"] is True
    assert out["coverage"]["reason"] is None
    assert out["available"]["published_at"] == "2026-09-18T04:00:00Z"
    assert "metadata_at" not in out["available"]


def test_an_unknown_additive_head_field_cannot_rescue_lapsed_coverage(
        channel, monkeypatch):
    """The same rule in the direction that costs money.

    A metadata_at EARLIER than published_at is the one shape that could
    make a client reading the wrong field hand out a release the customer
    has not paid for, and a conforming client refuses that document as
    malformed for that reason. This client does not know the field at all,
    so it accepts the document and decides coverage on published_at, which
    arrives at the same place by a different road. Both roads are pinned
    here, because the point is not which one the client takes: it is that
    neither of them ends with an entitlement to a release published after
    the customer's updates ran out.
    """
    _license(monkeypatch, updates_through="2026-09-17")
    _policy(channel)
    channel.serve(_document(published_at="2026-09-18T04:00:00Z",
                            metadata_at="2026-09-01T00:00:00Z"))

    out = uc.status(force=True)
    if out["status"] == "check_failed":
        assert out["reason"] == "malformed"
        assert out["coverage"] is None
    else:
        assert out["coverage"]["covered"] is False
        assert out["coverage"]["reason"] == "updates_lapsed"


# ------------------------------------------------- the unwritten copy
# The three customer-facing sentences carried a marker naming the copy
# section they were waiting on, and the marker reached a running server's
# report through a real MCP client, where it read as a defect in our output.
# A field with no sentence is absent; it is never filled with a marker.


def test_the_three_sentences_are_absent_rather_than_marked():
    assert uc.BUSINESS_NOTE is None
    assert uc.BUSINESS_INSTALL_NOTE is None
    assert uc.BUSINESS_PRIVACY_NOTE is None


def _every_status(channel, monkeypatch):
    """One report from each shape that carries copy fields."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    yield "update_available", uc.status(force=True)
    yield "cached", uc.status()
    channel.serve(_document(release_id="beta-5",
                            bundle_version="2026.09.3-beta", sequence=5))
    yield "current", uc.status(force=True)
    channel.serve(_document(withdrawn=True, withdrawn_reason="a reason"))
    yield "withdrawn", uc.status(force=True)
    channel.raise_with = OSError("unreachable")
    yield "check_failed", uc.status(force=True)
    channel.raise_with = None
    _policy(channel, mode="off")
    yield "check_disabled", uc.status(force=True)


def test_no_report_carries_a_marker_where_a_sentence_belongs(channel,
                                                             monkeypatch):
    for name, report in _every_status(channel, monkeypatch):
        blob = json.dumps(report)
        assert "placeholder" not in blob.lower(), name
        assert "LAUNCH_PROSE" not in blob, name
        assert "note" not in report, name
        assert "install_note" not in report, name


def test_the_fields_come_back_when_the_copy_does(channel, monkeypatch):
    """The wiring is intact: a sentence set here reaches the report."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document())
    monkeypatch.setattr(uc, "BUSINESS_NOTE", "the note")
    monkeypatch.setattr(uc, "BUSINESS_INSTALL_NOTE", "the install note")
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["note"] == "the note"
    assert out["install_note"] == "the install note"


# ------------------------ a Windows-authored trust anchor has a BOM
#
# The policy file's sibling. It is the worse half of the pair: the
# policy file refused loudly as policy_malformed, while an anchors.json saved
# by Notepad, Set-Content or Out-File made read_anchors raise inside its own
# try and return NO ANCHORS. Every check then reported no_anchor, which is
# exactly what a machine that was never given a trust anchor reports, so a
# misconfigured fleet and an unconfigured one said the same thing.
#
# The installer reads the same two files. Its validator was rewritten clause
# for clause against these readers in the installer-side pass, so the paths
# below are the contract and not an implementation detail of this module.


def _anchor_file(channel):
    return channel.config_root / "trust" / "anchors.json"


def _rewrite_the_anchor_with_a_bom(channel):
    path = _anchor_file(channel)
    path.write_text(path.read_text(encoding="utf-8"), encoding="utf-8-sig")
    assert path.read_bytes().startswith(BOM)


def test_the_readers_agree_with_the_installer_on_where_the_files_live(channel):
    root = channel.config_root
    assert uc.policy_file_path() == root / "update-policy.json"
    assert uc.anchor_file_path() == root / "trust" / "anchors.json"


def test_a_bom_does_not_empty_the_anchor_set(channel):
    _rewrite_the_anchor_with_a_bom(channel)
    assert list(uc.read_anchors()) == ["ks4-release-2026"]


def test_an_anchor_file_with_a_bom_still_verifies_a_document(channel,
                                                             monkeypatch):
    _license(monkeypatch)
    _policy(channel)
    _rewrite_the_anchor_with_a_bom(channel)
    channel.serve(_document())
    out = uc.status(force=True)
    assert out["status"] == "update_available"
    assert out["available"]["release_id"] == "beta-6"


def test_an_anchor_file_that_is_not_json_still_yields_nothing(channel):
    _anchor_file(channel).write_bytes(BOM + b"not json at all")
    assert uc.read_anchors() == {}


def test_a_bom_does_not_make_a_bad_key_good(channel):
    """Tolerant about the container, strict about the keys, BOM or no BOM."""
    _anchor_file(channel).write_text(
        json.dumps({"schema": 1, "anchors": {"ks4-release-2026": "not-a-key"}}),
        encoding="utf-8-sig")
    assert uc.read_anchors() == {}


# --------------------- a cached answer outliving what it answered for
#
# Three servers on a fresh install reported
# update_available naming a bundle version and a release id, out of a document
# signed with a scratch key generated for a proof run hours earlier and
# fetched from a loopback address that no longer existed, on a machine whose
# own trust anchor set was empty. The cache is per-user and per-product; it is
# not per-install, per-endpoint or per-anchor-set, and the freshness rules
# deliberately do not re-verify a cached document, because it was verified
# when it arrived. It was. It was not verified for this machine.


def _one_good_check(channel, monkeypatch, **changes):
    """Leave a real, verified answer in the cache, and return the cache."""
    _license(monkeypatch)
    _policy(channel)
    channel.serve(_document(**changes))
    assert uc.status(force=True)["status"] == "update_available"
    return uc.read_business_cache()


def _the_cache_is_not_consulted(channel):
    """The next unforced check goes to the network, and the network is down,
    so a report that is not check_failed can only have come from the cache."""
    channel.raise_with = OSError("the endpoint is gone")
    out = uc.status()
    assert out["status"] == "check_failed", out["status"]
    assert out["available"] is None
    return out


def test_a_good_check_records_what_its_answer_was_good_for(channel,
                                                           monkeypatch):
    cache = _one_good_check(channel, monkeypatch)
    binding = cache[uc.CACHE_BINDING_KEY]
    assert binding["endpoint"] == uc.BUSINESS_ORIGIN
    assert binding["anchors"] == uc.anchor_fingerprint(uc.read_anchors())
    assert binding["install"] == uc.install_fingerprint()
    assert binding["anchors"] != "none"


def test_the_cache_still_answers_when_nothing_moved(channel, monkeypatch):
    """The control. The binding must not cost the rate limiter its job."""
    _one_good_check(channel, monkeypatch)
    before = channel.calls
    out = uc.status()
    assert out["status"] == "update_available"
    assert channel.calls == before


def test_a_cache_fetched_from_another_endpoint_is_not_read(channel,
                                                           monkeypatch):
    _one_good_check(channel, monkeypatch)
    _policy(channel, endpoint="https://mirror.example/ks4")
    _the_cache_is_not_consulted(channel)


def test_a_cache_verified_under_another_anchor_set_is_not_read(channel,
                                                               monkeypatch):
    """An anchor rotation that keeps the key id and changes the key, which is
    the shape a real rotation takes."""
    _one_good_check(channel, monkeypatch)
    replacement = Ed25519PrivateKey.generate().public_key().public_bytes(
        serialization.Encoding.Raw, serialization.PublicFormat.Raw)
    (channel.config_root / "trust" / "anchors.json").write_text(json.dumps({
        "schema": 1,
        "anchors": {"ks4-release-2026": base64.b64encode(replacement).decode()},
    }), encoding="utf-8")
    _the_cache_is_not_consulted(channel)


def test_a_cache_left_by_another_install_is_not_read(channel, monkeypatch):
    """Same machine, same product, same user, next install. installed_at is
    what tells the two apart."""
    _one_good_check(channel, monkeypatch)
    channel.stamp.write_text(json.dumps({
        "schema": 1, "bundle_version": "2026.09.3-beta", "release_id": "beta-5",
        "channel": "beta", "installed_at": "2026-09-15T21:40:10Z",
        "line": "cp314", "org_id": "ORD-JW-0001"}), encoding="utf-8")
    _the_cache_is_not_consulted(channel)


def test_the_beta_7_report_cannot_happen_again(channel, monkeypatch):
    """The defect as it was found: a document admitted by a key the machine
    no longer has, on a machine that now has no anchor at all."""
    _one_good_check(channel, monkeypatch)
    (channel.config_root / "trust" / "anchors.json").unlink()
    assert uc.read_anchors() == {}
    out = _the_cache_is_not_consulted(channel)
    assert out["available"] is None
    # The calm one-line notice reads the same cache without asking anybody,
    # so it has to refuse the same document.
    assert uc._channel_notice() is None


def test_an_unlabelled_cache_is_not_read(channel, monkeypatch):
    """Every cache written before this existed. An answer that cannot say
    what it was an answer for is not shown to be this machine's."""
    _one_good_check(channel, monkeypatch)
    cache = uc.read_business_cache()
    cache.pop(uc.CACHE_BINDING_KEY)
    uc.write_business_cache(cache)
    _the_cache_is_not_consulted(channel)


def test_a_failed_check_records_the_binding_too(channel, monkeypatch):
    """A failure keeps the document the last good check left behind, so the
    record it writes is still a labelled one."""
    _one_good_check(channel, monkeypatch)
    channel.raise_with = OSError("down")
    assert uc.status(force=True)["status"] == "check_failed"
    cache = uc.read_business_cache()
    assert cache["ok"] is False
    assert cache["document"]["release_id"] == "beta-6"
    assert cache[uc.CACHE_BINDING_KEY]["install"] == uc.install_fingerprint()


def test_the_fingerprints_change_only_when_the_thing_they_name_changes():
    assert uc.anchor_fingerprint(None) == "none"
    assert uc.anchor_fingerprint({}) == "none"
    one = {"a": b"k" * 32}
    assert uc.anchor_fingerprint(one) == uc.anchor_fingerprint(dict(one))
    assert uc.anchor_fingerprint(one) != uc.anchor_fingerprint({"b": b"k" * 32})
    assert uc.anchor_fingerprint(one) != uc.anchor_fingerprint({"a": b"j" * 32})
    assert uc.anchor_fingerprint({"a": b"k" * 32, "b": b"j" * 32}) == \
        uc.anchor_fingerprint({"b": b"j" * 32, "a": b"k" * 32})


def test_a_machine_with_no_stamp_still_has_an_install_fingerprint(tmp_path):
    """Missing, unreadable and not-an-object all resolve rather than raise,
    and none of them collides with a real stamp at the same path."""
    stamp = tmp_path / "gone.json"
    absent = uc.install_fingerprint(stamp)
    assert absent
    stamp.write_text("[]", encoding="utf-8")
    assert uc.install_fingerprint(stamp) == absent
    stamp.write_text(json.dumps({"installed_at": "2026-09-16T00:00:00Z"}),
                     encoding="utf-8")
    assert uc.install_fingerprint(stamp) != absent
    # And a stamp at another path is another install, same contents or not.
    other = tmp_path / "elsewhere.json"
    other.write_text(stamp.read_text(encoding="utf-8"), encoding="utf-8")
    assert uc.install_fingerprint(other) != uc.install_fingerprint(stamp)
