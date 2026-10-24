"""Explicit policy injection for repository developer scripts.

A script that imports the editing core is doing what a launcher does, so it
has to choose an edition on purpose rather than have one default into place.
Measurement and counting scripts choose Community, because Community is the
edition whose published numbers they produce.

The Community adapter comes from the canonical pinned source, not from a copy
in this repository. See adapter_source for why there is no copy.

Import this before importing xlsx_mcp.server.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = ROOT / "packages"

sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source  # noqa: E402


def use_community() -> None:
    for entry in (str(ROOT / "src"),
                  str(PACKAGES / "kitchensink4xl-community" / "src")):
        if entry not in sys.path:
            sys.path.insert(0, entry)
    adapter_source.ensure_importable()
    from ks4_license_community.policy import CommunityPolicy, ProductConfig
    from xlsx_mcp import __version__, license_policy
    if not license_policy.installed():
        license_policy.install(CommunityPolicy(
            ProductConfig(product_id="kitchensink4xl",
                          display_name="KitchenSink4XL"), __version__))
