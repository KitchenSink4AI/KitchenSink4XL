"""The license adapters are ONE shared source: whatever tree the
suite, the build or CI materialises must be the pinned content. A CI export
supplied through KS4_LICENSE_ADAPTERS is checked against the recorded
digest exactly like the git-materialised cache.
Edition-aware: the private path checks both adapters against ADAPTER_DIGEST,
the Community-only path the Community adapter against COMMUNITY_DIGEST."""

from __future__ import annotations

import shutil

import pytest

from adapter_source import (ENV_MATERIALISED, AdapterSourceError, expected_digest,
                            materialise, required_dists, tree_digest)


def test_the_pinned_tree_has_the_recorded_digest():
    assert tree_digest(materialise()) == expected_digest()


def test_a_faithful_export_is_accepted(tmp_path, monkeypatch):
    for dist in required_dists():
        shutil.copytree(materialise() / dist, tmp_path / dist)
    monkeypatch.setenv(ENV_MATERIALISED, str(tmp_path))
    assert materialise() == tmp_path


def test_red_an_edited_export_is_refused(tmp_path, monkeypatch):
    for dist in required_dists():
        shutil.copytree(materialise() / dist, tmp_path / dist)
    victim = next((tmp_path / required_dists()[-1]).rglob("*.py"))
    victim.write_text(victim.read_text("utf-8") + "\n# edited\n", "utf-8")
    monkeypatch.setenv(ENV_MATERIALISED, str(tmp_path))
    with pytest.raises(AdapterSourceError):
        materialise()
