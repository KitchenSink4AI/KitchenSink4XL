"""Edition-aware adapter resolution.

The public repository builds and tests the Community edition only and never
receives the Business adapter source, so KS4_ADAPTER_EDITION=community
resolves, imports and digests the Community adapter alone, against
COMMUNITY_DIGEST. The private two-adapter path (unset or "all") is unchanged:
it still requires both packages and ADAPTER_DIGEST, and still refuses a
Community-only directory.
"""
from __future__ import annotations

import shutil

import pytest

import adapter_source as A


@pytest.fixture
def community_only(tmp_path):
    root = tmp_path / "community-only"
    shutil.copytree(A.materialise() / A.COMMUNITY_POLICY_DIST, root / A.COMMUNITY_POLICY_DIST,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    return root


def _set(monkeypatch, edition, source):
    monkeypatch.setenv(A.ENV_EDITION, edition)
    monkeypatch.setenv(A.ENV_MATERIALISED, str(source))


def test_the_community_package_has_the_recorded_community_digest(community_only):
    assert A.tree_digest(community_only, (A.COMMUNITY_POLICY_DIST,)) == A.COMMUNITY_DIGEST


def test_community_path_accepts_a_community_only_tree(monkeypatch, community_only):
    _set(monkeypatch, "community", community_only)
    assert A.materialise().resolve() == community_only.resolve()
    assert A.required_dists() == (A.COMMUNITY_POLICY_DIST,)


def test_community_path_never_hands_out_the_business_adapter(monkeypatch, community_only):
    _set(monkeypatch, "community", community_only)
    assert A.import_paths() == [str(community_only / A.COMMUNITY_POLICY_DIST / 'src')]
    with pytest.raises(A.AdapterSourceError):
        A.package_dir(A.BUSINESS_POLICY_DIST)


def test_community_path_refuses_an_edited_adapter(monkeypatch, community_only):
    victim = next((community_only / A.COMMUNITY_POLICY_DIST).rglob("*.py"))
    victim.write_text(victim.read_text("utf-8") + "\n# edited\n", "utf-8")
    _set(monkeypatch, "community", community_only)
    with pytest.raises(A.AdapterSourceError):
        A.materialise()


def test_private_path_still_requires_business(monkeypatch, community_only):
    # The two-adapter path refuses a Community-only directory.
    _set(monkeypatch, "all", community_only)
    with pytest.raises(A.AdapterSourceError, match="business"):
        A.materialise()


def test_unknown_edition_is_refused(monkeypatch, community_only):
    _set(monkeypatch, "business-lite", community_only)
    with pytest.raises(A.AdapterSourceError):
        A.edition()


@pytest.mark.parametrize("value", ["", "abc1234", "0" * 39, "G" * 40, "A" * 40])
def test_public_pin_refuses_anything_but_a_full_sha(monkeypatch, value):
    monkeypatch.setattr(A, "PUBLIC_ADAPTER_COMMIT", value)
    with pytest.raises(A.AdapterSourceError):
        A.public_pin()


def test_public_pin_names_the_license_repository(monkeypatch):
    monkeypatch.setattr(A, "PUBLIC_ADAPTER_COMMIT", "0123456789abcdef0123456789abcdef01234567")
    pin = A.public_pin()
    assert pin["repository"] == "KitchenSink4AI/kitchensink4-license-community"
    assert pin["path"] == "packages/kitchensink4-license-community"
    assert pin["digest"] == A.COMMUNITY_DIGEST
