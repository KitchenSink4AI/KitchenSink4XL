"""Suite-wide setup for the KitchenSink4XL tests.

One job so far: NO NETWORK IN TESTS. get_server_info runs the on-demand
update check, so the suite turns that check off for every test by default.
The update tests switch it back on for themselves and mock the fetch, and
setdefault means a developer who exports the variable still wins.
"""

from __future__ import annotations

import os

import pytest


def pytest_configure(config):
    os.environ.setdefault("KS4XL_UPDATE_CHECK", "off")


@pytest.fixture(autouse=True)
def _pack_store_per_test(tmp_path_factory, monkeypatch):
    """Saved pack choices (packstore.py) go to a fresh directory for every
    test, never to the developer's own state directory, so no test starts
    from a choice another test (or a real session) saved."""
    monkeypatch.setenv(
        "KS4XL_PACK_STORE_DIR", str(tmp_path_factory.mktemp("pack-store")))
