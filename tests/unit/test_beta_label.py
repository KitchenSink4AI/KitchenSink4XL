"""The beta label: one word for the stage of the product, carried as data.

The CHANNEL is build metadata. It is generated into `_edition.json`, validated
by the shared policy adapter against a closed set, and reported by
get_server_info beside the release date. It is deliberately not a version
suffix: the wheel stays at X.Y.Z+business so that pip resolution and the
published identity of the artifact do not move with a label.

The LANGUAGE GATE is a spelling check on a commercial promise. A customer who
reads "preview", "RC", "trial" or "demo" has been told something about
support, price or duration that is not true of this release, so the build
refuses to ship a customer-facing string carrying one. Its allowlist is
explicit and every entry is a document whose wording belongs to somebody other
than the build script.

The ISSUER KEY MAP is the purchase flow. Self-serve receipts are signed on the
fulfillment Worker with the web-tier key and manual or enterprise receipts are
signed with the offline root, so a build carries several public keys and the
receipt's own key id selects between them.

The exhaustive adapter-side tests live in the Word repository, which is
where the shared adapter source lives. What is here is what is this product's
own, plus the one multi-key receipt case that proves this product consumes the
fixed adapter rather than a stale copy of it.
"""
import asyncio
import base64
import importlib.util
import inspect
import json
import zipfile
from datetime import date
from pathlib import Path

import pytest
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey


from ks4_license_business import policy as business_policy  # noqa: E402
from ks4_license_business import receipts  # noqa: E402

from xlsx_mcp import license_policy, server

ROOT = Path(__file__).resolve().parents[2]

spec = importlib.util.spec_from_file_location(
    "beta_label_build", ROOT / "scripts/build_editions.py")
build_script = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build_script)

PRODUCT_ID = "kitchensink4xl-business"
DISPLAY_NAME = "KitchenSink4XL Business"
LICENSE_ENV = "KS4XL_LICENSE_FILE"
LAUNCHER = "ks4xl_business_launch"
PRODUCT_CONFIG = business_policy.ProductConfig(
    product_id=PRODUCT_ID, display_name=DISPLAY_NAME,
    license_file_env=LICENSE_ENV)


def _keypair():
    key = Ed25519PrivateKey.generate()
    raw = key.public_key().public_bytes(serialization.Encoding.Raw,
                                        serialization.PublicFormat.Raw)
    return key, base64.b64encode(raw).decode()


def _metadata(keys=None, **extra):
    data = {"edition": "business", "package": PRODUCT_ID,
            "version": "1.2.2+business", "release_date": "2026-09-10",
            "public_keys": keys if keys is not None else {}}
    data.update(extra)
    return data


def _pem(tmp_path, name):
    key = Ed25519PrivateKey.generate()
    path = tmp_path / name
    path.write_bytes(key.public_key().public_bytes(
        serialization.Encoding.PEM, serialization.PublicFormat.SubjectPublicKeyInfo))
    return key, path


def _server_info(**kwargs):
    """get_server_info, whether this product's surface is sync or async.

    Web's lite surface is a coroutine function and the other three are plain
    calls. The assertion is about what the server reports, not about how it is
    invoked, so the difference is absorbed here rather than forked into two
    test files.
    """
    result = server.get_server_info(**kwargs)
    if inspect.iscoroutine(result):
        return asyncio.run(result)
    return result


def _launcher_edition(output: Path) -> dict:
    # The launcher by its own name: the Business adapter wheel lands in the
    # same folder and also matches "*business-*", and which one a bare glob
    # returns first depends on the filesystem (ubuntu and macos CI picked
    # the adapter, which carries no _edition.json).
    wheel = next(output.glob("kitchensink4xl_business-*.whl"))
    with zipfile.ZipFile(wheel) as archive:
        return json.loads(archive.read(f"{LAUNCHER}/_edition.json"))


# --------------------------------------------------------------- the channel


@pytest.mark.parametrize("channel", ["beta", "stable"])
def test_the_adapter_accepts_the_two_channels(channel):
    data = business_policy.validate_build_metadata(
        _metadata(channel=channel), PRODUCT_CONFIG)
    assert data["channel"] == channel


@pytest.mark.parametrize("channel", ["preview", "rc", "trial", "", None])
def test_a_channel_outside_the_set_is_a_setup_error(channel):
    with pytest.raises(business_policy.BuildMetadataError) as problem:
        business_policy.validate_build_metadata(
            _metadata(channel=channel), PRODUCT_CONFIG)
    assert "beta, stable" in str(problem.value)


def test_a_build_from_before_the_field_existed_is_still_valid():
    assert business_policy.BusinessPolicy(PRODUCT_CONFIG, _metadata()).channel is None


def test_business_server_info_reports_the_channel_and_the_release_date(monkeypatch):
    previous = license_policy.active()
    license_policy.install(business_policy.BusinessPolicy(
        PRODUCT_CONFIG, _metadata(channel="beta")))
    monkeypatch.delenv(LICENSE_ENV, raising=False)
    try:
        info = _server_info()
        assert info["channel"] == "beta"
        assert info["release_date"] == "2026-09-10"
        assert info["version"] == "1.2.2+business"
    finally:
        license_policy.install(previous)


def test_community_server_info_reports_no_channel_and_no_release_date():
    info = _server_info()
    assert info["edition"] == "community"
    assert info["channel"] is None
    assert info["release_date"] is None


@pytest.mark.business_sources
def test_the_generated_metadata_carries_the_channel(tmp_path):
    build_script.build("business", tmp_path / "b", date(2026, 9, 14), None,
                       channel="stable")
    data = _launcher_edition(tmp_path / "b")
    assert data["channel"] == "stable"
    assert data["release_date"] == "2026-09-14"
    assert data["version"].endswith("+business")
    # The channel is never folded into the version: the version is the
    # package's own, plus the edition tag. (A substring check for "rc"
    # read the package's own release-candidate number as a fold-in.)
    import tomllib
    package = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    assert data["version"] == f"{package}+business"


@pytest.mark.business_sources
def test_the_default_channel_is_beta(tmp_path):
    build_script.build("business", tmp_path / "b", date(2026, 9, 14), None)
    assert _launcher_edition(tmp_path / "b")["channel"] == "beta"
    assert build_script.DEFAULT_CHANNEL == "beta"


def test_a_build_refuses_a_channel_outside_the_set(tmp_path):
    with pytest.raises(ValueError, match="channel must be one of"):
        build_script.build("business", tmp_path / "b", date(2026, 9, 14), None,
                           channel="preview")


# ---------------------------------------------------------- the wording leaks


@pytest.mark.business_sources
def test_the_bundle_description_calls_the_build_a_beta(tmp_path):
    build_script.build("business", tmp_path / "b", date(2026, 9, 14), None)
    archive = next((tmp_path / "b").glob("*.mcpb"))
    with zipfile.ZipFile(archive) as bundle:
        manifest = json.loads(bundle.read("manifest.json"))
    assert manifest["description"] == (
        "KitchenSink4XL Business, beta release. Install the supplied wheels "
        "and register the launcher as described in INSTALL.md.")
    assert manifest["long_description"] == manifest["description"]


@pytest.mark.business_sources
def test_the_business_readme_calls_the_build_a_beta(tmp_path):
    scratch = tmp_path / "scratch"
    scratch.mkdir()
    build_script.build_launcher("business", tmp_path / "out", scratch,
                                date(2026, 9, 14), {}, "0.0.0+commercial", "0.2.0")
    readme = next(scratch.glob("*business/README.md")).read_text("utf-8")
    assert readme.startswith(
        "# KitchenSink4XL Business\n\nKitchenSink4AI Business, beta release. "
        "Built and verified like production, licensed for evaluation under the "
        "Pilot Evaluation License.\n\n")
    assert "Private development preview" not in readme


# ------------------------------------------------------- the language gate


@pytest.mark.parametrize("phrase, expected", [
    ("a development preview", ["preview"]),
    ("this is the RC", ["RC"]),
    ("release candidate 3", ["release candidate"]),
    ("a 30 day trial", ["trial"]),
    ("try the demo", ["demo"]),
    ("beta release", []),
    ("the arc of the source", []),
    ("democratic recording", []),
])
def test_the_gate_knows_the_words_it_is_looking_for(phrase, expected):
    assert build_script.find_prerelease_language(phrase) == expected


def _fake_wheel(path, members):
    with zipfile.ZipFile(path, "w") as archive:
        for name, text in members.items():
            archive.writestr(name, text)
    return path


def test_a_gate_that_reads_nothing_refuses_to_pass_anything(tmp_path):
    wheel = _fake_wheel(tmp_path / "a.whl", {"pkg/code.py": "x = 'preview'\n"})
    with pytest.raises(build_script.BuildScanError, match="would pass anything"):
        build_script.scan_prerelease_language([wheel])


@pytest.mark.parametrize("member", ["pkg/_edition.json", "a-1.dist-info/METADATA"])
def test_pre_release_language_in_a_scanned_member_fails_the_build(tmp_path, member):
    wheel = _fake_wheel(tmp_path / "a.whl", {
        "pkg/_edition.json": '{"channel": "beta"}',
        "a-1.dist-info/METADATA": "Name: a\n",
        member: "Name: a\nSummary: a development preview\n",
    })
    with pytest.raises(build_script.BuildScanError, match="something other than a beta"):
        build_script.scan_prerelease_language([wheel])


def test_the_licence_text_is_now_read(tmp_path):
    """LICENSE was out of scope only while it said development preview."""
    wheel = _fake_wheel(tmp_path / "a.whl", {
        "pkg/_edition.json": '{"channel": "beta"}',
        "a-1.dist-info/licenses/LICENSE": "a private development preview\n",
    })
    with pytest.raises(build_script.BuildScanError, match="something other than a beta"):
        build_script.scan_prerelease_language([wheel])


def test_a_member_outside_the_target_set_is_not_read(tmp_path):
    """The gate reads build-authored customer text, not every member."""
    wheel = _fake_wheel(tmp_path / "a.whl", {
        "pkg/_edition.json": '{"channel": "beta"}',
        "pkg/NOTICE.md": "a Business preview is being developed\n",
    })
    report = build_script.scan_prerelease_language([wheel])
    assert report["members_scanned"] == 1
    assert "LICENSE-BETA.txt" in build_script.PRERELEASE_ALLOWLIST


@pytest.mark.business_sources
def test_a_real_business_build_passes_the_language_gate(tmp_path):
    artifacts = build_script.build("business", tmp_path / "b", date(2026, 9, 14), None)
    report = build_script.scan_prerelease_language(artifacts)
    assert report["members_scanned"] >= 4


# --------------------------------------------------------- the issuer keys


def test_one_key_with_no_id_keeps_the_default(tmp_path):
    _, pem = _pem(tmp_path, "one.pem")
    assert list(build_script.issuer_keys(pem, None)) == [build_script.DEFAULT_KEY_ID]


def test_several_keys_are_paired_in_order(tmp_path):
    _, web = _pem(tmp_path, "web.pem")
    _, root = _pem(tmp_path, "root.pem")
    keys = build_script.issuer_keys([web, root],
                                    ["ks4-issuer-web-2026", "ks4-issuer-2026"])
    assert list(keys) == ["ks4-issuer-web-2026", "ks4-issuer-2026"]
    assert len({*keys.values()}) == 2


def test_a_duplicate_key_id_is_refused(tmp_path):
    _, a = _pem(tmp_path, "a.pem")
    _, b = _pem(tmp_path, "b.pem")
    with pytest.raises(ValueError, match="Duplicate issuer key id"):
        build_script.issuer_keys([a, b], ["same", "same"])


def test_a_mismatched_pairing_is_refused(tmp_path):
    _, a = _pem(tmp_path, "a.pem")
    _, b = _pem(tmp_path, "b.pem")
    with pytest.raises(ValueError, match="paired in order"):
        build_script.issuer_keys([a, b], ["only-one"])


@pytest.mark.business_sources
def test_a_build_carries_every_key_it_was_given(tmp_path):
    _, web = _pem(tmp_path, "web.pem")
    _, root = _pem(tmp_path, "root.pem")
    build_script.build("business", tmp_path / "b", date(2026, 9, 14),
                       [web, root], ["ks4-issuer-web-2026", "ks4-issuer-2026"])
    data = _launcher_edition(tmp_path / "b")
    assert sorted(data["public_keys"]) == ["ks4-issuer-2026", "ks4-issuer-web-2026"]


def test_a_receipt_signed_by_either_key_verifies(tmp_path, monkeypatch):
    web, web_encoded = _keypair()
    root, root_encoded = _keypair()
    policy = business_policy.BusinessPolicy(PRODUCT_CONFIG, _metadata(
        {"ks4-issuer-web-2026": web_encoded, "ks4-issuer-2026": root_encoded},
        channel="beta"))
    payload = {"schema": 1, "product": PRODUCT_ID, "license_id": "ORD-1",
               "licensee": "Harbourline Surveying Group Pty Ltd",
               "issued_on": "2026-09-01", "updates_through": "2027-09-01"}
    for key_id, signer in (("ks4-issuer-web-2026", web), ("ks4-issuer-2026", root)):
        path = tmp_path / f"{key_id}.json"
        path.write_text(json.dumps({
            "key_id": key_id, "payload": payload,
            "signature": base64.b64encode(signer.sign(receipts._canonical(payload))).decode()}))
        monkeypatch.setenv(LICENSE_ENV, str(path))
        status = policy.describe_license(redacted=False)
        assert status["status"] == "verified", key_id


def test_a_receipt_naming_an_unknown_key_is_refused_safely(tmp_path, monkeypatch):
    _, web_encoded = _keypair()
    stranger, _ = _keypair()
    policy = business_policy.BusinessPolicy(
        PRODUCT_CONFIG, _metadata({"ks4-issuer-web-2026": web_encoded}, channel="beta"))
    payload = {"schema": 1, "product": PRODUCT_ID, "license_id": "ORD-2",
               "licensee": "Someone Else Entirely", "issued_on": "2026-09-01",
               "updates_through": "2027-09-01"}
    path = tmp_path / "stranger.json"
    path.write_text(json.dumps({
        "key_id": "ks4-issuer-nobody-2029", "payload": payload,
        "signature": base64.b64encode(stranger.sign(receipts._canonical(payload))).decode()}))
    monkeypatch.setenv(LICENSE_ENV, str(path))
    status = policy.describe_license(redacted=False)
    assert status["status"] == "unverified"
    assert status["reason"] == "signature_invalid"
    rendered = json.dumps(status)
    assert "Someone Else Entirely" not in rendered
    assert "ks4-issuer-nobody-2029" not in rendered


# -------------------------------------------- the business distribution notice
# What the Business distribution says about itself. The paragraph is fixed text,
# so it is reproduced here rather than described: the assertion worth having is
# that the shipped bytes are the blessed bytes. The second test is the other
# half. The words this paragraph replaced described an unfinished product with
# nothing decided, and none of them may come back through an edit to any of the
# files that carry it.

BUSINESS_TERMS_PARAGRAPH = """\
This Business distribution is licensed to your organization under the Business
terms: a yearly subscription by user band, not transferable. Support and the
Business update channel run for the paid term; the tools keep working after it.
"""

#: Every file in this repository that carries the paragraph.
NOTICE_FILES = (
    "business/COMMERCIAL_LICENSE_NOTICE.md",
    "packages/kitchensink4xl-business/LICENSE",
)

#: Not the build's language gate, which passes "beta release" deliberately.
#: This is the narrower rule for these files: the distribution notice does not
#: describe the product as unfinished or its commercial terms as undecided.
FORBIDDEN_IN_THE_NOTICE = (
    "beta", "preview", "release candidate", "release-candidate", "pre-release",
    "prerelease", "development", "evaluation", "have not been", "not approved")


@pytest.mark.parametrize("relative", NOTICE_FILES)
@pytest.mark.business_sources
def test_the_notice_carries_the_business_terms_paragraph(relative):
    assert BUSINESS_TERMS_PARAGRAPH in (ROOT / relative).read_text("utf-8")


@pytest.mark.parametrize("relative", NOTICE_FILES)
@pytest.mark.business_sources
def test_the_notice_does_not_call_the_distribution_unfinished(relative):
    text = (ROOT / relative).read_text("utf-8").lower()
    assert [word for word in FORBIDDEN_IN_THE_NOTICE if word in text] == []


@pytest.mark.parametrize("relative", NOTICE_FILES)
@pytest.mark.business_sources
def test_the_notice_passes_the_build_language_gate(relative):
    assert build_script.find_prerelease_language(
        (ROOT / relative).read_text("utf-8")) == []
