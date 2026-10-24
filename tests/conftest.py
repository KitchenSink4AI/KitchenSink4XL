"""Suite-wide setup for the KitchenSink4XL tests.

Two jobs now.

NO NETWORK IN TESTS. get_server_info runs the on-demand update check, so the
suite turns that check off for every test by default. The update tests switch
it back on for themselves and mock the fetch, and setdefault means a developer
who exports the variable still wins.

A LICENSE POLICY, INJECTED EXPLICITLY. The core does not choose an edition and
does not run without a policy, so the suite plays the part a launcher plays in
a real installation. Community is what a source checkout exercises by default;
the Business tests build and inject their own policy per test. The path entries
make the launchers importable from a bare checkout, and the shared license
policy adapters come from the canonical pinned source rather than from a copy
in this repository (see scripts/adapter_source.py).
"""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = ROOT / "packages"

#: The Business edition's sources (tests marked business_sources need them).
BUSINESS_SOURCES = (
    ROOT / "business" / "COMMERCIAL_LICENSE_NOTICE.md",
    PACKAGES / "kitchensink4xl-business" / "pyproject.toml",
    PACKAGES / "kitchensink4xl-business" / "LICENSE",
    PACKAGES / "kitchensink4xl-business" / "src" / "ks4xl_business_launch" / "__init__.py",
    ROOT / "scripts" / "sign_business_license.py",
)

_spec = importlib.util.spec_from_file_location(
    "_ks4xl_adapter_source", ROOT / "scripts" / "adapter_source.py")
adapter_source = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(adapter_source)
sys.modules.setdefault("adapter_source", adapter_source)


# EDITIONS. KS4_ADAPTER_EDITION=community is the public checkout: it carries the
# Community edition only (no Business adapter, Business launcher or business/
# tree). The modules and tests listed below exercise Business code (or, for a
# few tests, the commercial notice or an internal audit the public tree does
# not carry; each entry says which). They run wherever those sources exist
# (private staging and every Business build) and are not collected or are
# deselected on the Community path; the session header names them. tests/unit/test_business_only_manifest.py
# keeps the lists exact: every test module that imports Business code at top
# level is listed, and nothing else is.
EDITION = (os.environ.get("KS4_ADAPTER_EDITION") or "").strip().lower() or "all"
if EDITION not in ("all", "community"):
    raise RuntimeError(f"KS4_ADAPTER_EDITION={EDITION!r}: expected all or community")
BUSINESS_ONLY_MODULES = (
    "unit/test_beta_label.py",
    "unit/test_chain_verification.py",
    "unit/test_editions.py",
    "unit/test_product_neutrality.py",
    "unit/test_receipt_compat.py",
    "unit/test_update_business.py",
)
BUSINESS_ONLY_TESTS = {
    "tests/unit/test_license_seam.py::test_setup_error_is_not_a_receipt_error":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_business_package_imports_no_community_module":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_business_package_does_not_depend_on_the_editing_core":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_business_package_imports_cleanly_without_community_installed":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_additional_tool_registration_is_optional_and_neither_policy_uses_it":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_both_shipped_policies_satisfy_the_tightened_validator":
        "Business adapter or launcher code",
    "tests/unit/test_license_seam.py::test_only_the_business_policy_declares_an_update_backend":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_the_allowlist_lets_the_commercial_notice_keep_its_agpl_sentences":
        "the commercial notice business/COMMERCIAL_LICENSE_NOTICE.md (not in the public tree)",
    "tests/unit/test_build_scan_gate.py::test_business_build_refuses_metadata_drift":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_real_business_build_ships_a_commercially_labelled_core":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_poisoned_core_label_fails_the_real_business_build":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_the_required_notices_survive_the_commercial_relabelling":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_real_business_build_passes_the_gate_and_ships_no_community_module":
        "Business adapter or launcher code",
    "tests/unit/test_build_scan_gate.py::test_a_poisoned_build_input_fails_the_real_build":
        "Business adapter or launcher code",
    "tests/unit/test_edition_regressions.py::test_approved_warning_preserves_exact_section_symbol":
        "the internal audit docs/audit-community-business.md (not in the public tree)",
    "tests/unit/test_edition_regressions.py::test_business_build_refuses_metadata_drift":
        "Business adapter or launcher code",
    "tests/unit/test_dependency_ledger.py::test_ledger_exists":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_every_required_dependency_is_in_the_ledger":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_every_optional_dependency_is_in_the_ledger":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_no_copyleft_in_the_required_install":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_required_licenses_are_recognizably_permissive":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
    "tests/unit/test_dependency_ledger.py::test_the_ledger_records_the_font_attribution":
        "DEPENDENCY_LEDGER.md, an internal record the snapshot filter never ships; the ledger checks run in private staging",
}
collect_ignore = list(BUSINESS_ONLY_MODULES) if EDITION == "community" else []


def pytest_report_header(config):
    if EDITION != "community":
        return ["edition: all (Community and Business tests)"]
    return [f"edition: Community only. Not run here (Business code, run in private staging): "
            f"{len(BUSINESS_ONLY_MODULES)} modules {list(BUSINESS_ONLY_MODULES)}, "
            f"{len(BUSINESS_ONLY_TESTS)} tests {list(BUSINESS_ONLY_TESTS)}"]


def _deselect_business_only(config, items):
    if EDITION != "community" or not BUSINESS_ONLY_TESTS:
        return
    keep, drop = [], []
    for item in items:
        # an entry names one parametrized case exactly, or every case of a test
        hit = item.nodeid in BUSINESS_ONLY_TESTS or item.nodeid.split("[", 1)[0] in BUSINESS_ONLY_TESTS
        (drop if hit else keep).append(item)
    if drop:
        config.hook.pytest_deselected(items=drop)
        items[:] = keep

# The launchers live here; the adapters do not, and deliberately so.
for _pkg in (("kitchensink4xl-community",) if EDITION == "community"
             else ("kitchensink4xl-community", "kitchensink4xl-business")):
    _src = str(PACKAGES / _pkg / "src")
    if _src not in sys.path:
        sys.path.insert(0, _src)
adapter_source.ensure_importable()

from ks4_license_community.policy import CommunityPolicy as _CommunityPolicy  # noqa: E402
from ks4_license_community.policy import ProductConfig as _CommunityConfig  # noqa: E402
from xlsx_mcp import __version__ as _core_version  # noqa: E402
from xlsx_mcp import license_policy as _license_policy  # noqa: E402

# Product identity comes from the launcher in a real installation; here the
# suite plays that part, naming the product explicitly rather than relying on
# a default the adapters deliberately no longer have.
_license_policy.install(_CommunityPolicy(
    _CommunityConfig(product_id="kitchensink4xl", display_name="KitchenSink4XL"),
    _core_version))


def pytest_configure(config):
    os.environ.setdefault("KS4XL_UPDATE_CHECK", "off")
    # The save journal of a test run is its own, never the one a
    # real server on this machine keeps in LOCALAPPDATA.
    import tempfile
    os.environ.setdefault(
        "KS4XL_SAVE_JOURNAL_DIR",
        os.path.join(tempfile.gettempdir(), f"ks4xl-test-journal-{os.getpid()}"))


def pytest_collection_modifyitems(config, items):
    _deselect_business_only(config, items)
    # Release gates on public copy (README, llms.txt, site) run only when the
    # release ship gate asks for them; the public copy is written late, close
    # to release. They are not removed.
    if os.environ.get("KS4XL_RELEASE_GATE") != "1":
        import pytest

        skip_copy = pytest.mark.skip(reason="public_copy release gate: set KS4XL_RELEASE_GATE=1")
        for item in items:
            if item.get_closest_marker("public_copy"):
                item.add_marker(skip_copy)
    # The Business edition's sources are not in a tree made from the public
    # release, which ships the Community edition only; the tests that build
    # or read them say so instead of failing there.
    if not all(p.is_file() for p in BUSINESS_SOURCES):
        import pytest

        skip_business = pytest.mark.skip(
            reason="needs the Business edition sources, which the public release does not include")
        for item in items:
            if item.get_closest_marker("business_sources"):
                item.add_marker(skip_business)
