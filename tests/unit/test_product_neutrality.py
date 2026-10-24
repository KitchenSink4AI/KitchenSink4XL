"""The shared adapters name no product.

The defect this file guards against is a packaging one rather than a
behavioural one. Four products cannot each publish a wheel called
kitchensink4-license-business at the same version with different payloads
inside; pip, its cache and any co-installation would then resolve one name to
four different things. So the adapters carry no product identity at all, and
the launchers, whose distributions ARE product-specific, supply it.

Two properties follow and both are tested here: one adapter wheel serves
every editor, and a receipt still cannot wander between products.
"""
import ast
import json
import tomllib
from pathlib import Path

import pytest

import adapter_source

from ks4_license_business import receipts
from ks4_license_business.policy import (BuildMetadataError, BusinessPolicy,
                                         ProductConfig, ProductConfigError)
from ks4_license_community import notices
from ks4_license_community.policy import CommunityPolicy
from ks4_license_community.policy import ProductConfig as CommunityConfig

ROOT = Path(__file__).resolve().parents[2]
PACKAGES = ROOT / "packages"
#: One canonical adapter pair for the suite, read from the pinned source.
#: This repository carries no copy, which is the packaging property under test.
ADAPTERS = {
    "business": adapter_source.package_dir(adapter_source.BUSINESS_POLICY_DIST)
                / "src/ks4_license_business",
    "community": adapter_source.package_dir(adapter_source.COMMUNITY_POLICY_DIST)
                 / "src/ks4_license_community",
}
FIXTURES = Path(__file__).resolve().parents[1] / "fixtures" / "receipts"
MANIFEST = json.loads((FIXTURES / "MANIFEST.json").read_text(encoding="utf-8"))

WORD = ProductConfig(product_id="kitchensink4word-business",
                     display_name="KitchenSink4Word Business",
                     license_file_env="KS4W_LICENSE_FILE")
XL = ProductConfig(product_id="kitchensink4xl-business",
                   display_name="KitchenSink4XL Business",
                   license_file_env="KS4XL_LICENSE_FILE")


def _build(package: str, version: str = "1.2.2+business") -> dict:
    return {"edition": "business", "package": package, "version": version,
            "release_date": "2026-09-10",
            "public_keys": MANIFEST["build"]["public_keys"]}


# ------------------------------------------- no product names in the wheels


#: The four product brands, their PyPI names and their receipt environment
#: prefixes. None may appear in a shared wheel, in its Python OR in its legal
#: notices.
PRODUCT_MARKERS = (
    "KitchenSink4Word", "KitchenSink4XL", "KitchenSink4PPT", "KitchenSink4Web",
    "kitchensink4word", "kitchensink4xl", "kitchensink4ppt", "kitchensink4web",
    "KS4W_", "KS4XL_", "KS4P_", "KS4WEB_",
)


@pytest.mark.parametrize("edition", ["business", "community"])
def test_no_product_name_appears_anywhere_in_a_shared_adapter(edition):
    """The literal check. A product name in these files is the defect."""
    forbidden = list(PRODUCT_MARKERS) + ["word_mcp", "xlsx_mcp"]
    offenders = []
    for path in ADAPTERS[edition].rglob("*.py"):
        if "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8")
        for name in forbidden:
            if name in text:
                offenders.append(f"{path.name}: {name}")
    assert offenders == []


@pytest.mark.parametrize("edition", ["business", "community"])
def test_no_product_name_appears_in_a_shared_adapter_notice_or_license(edition):
    """The notice files, not just the Python.

    Found by the PowerPoint sibling against the canonical adapters: the code
    was neutral while NOTICE.md and both LICENSE files still said
    KitchenSink4Word, so one shared wheel would have shipped Word-named legal
    notices into all four products. This is the assertion that proves the pin
    bump to 5d3842b did what it was for; at 5bb0f20 it fails.
    """
    package = ADAPTERS[edition].parents[1]
    offenders = []
    for path in sorted(package.iterdir()):
        if path.is_dir() or path.name.startswith("."):
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        for marker in PRODUCT_MARKERS:
            if marker in text:
                offenders.append(f"{path.name}: {marker}")
    assert offenders == []


def test_the_per_product_distributions_keep_their_own_notices():
    """Neutralising the shared adapters must not neutralise the product
    distributions, whose notice is supposed to name the product."""
    community = PACKAGES / "kitchensink4xl-community" / "NOTICE.md"
    assert "KitchenSink4XL" in community.read_text(encoding="utf-8")


def test_the_community_distribution_license_is_the_plain_agpl_text():
    """The copyright notice lives in NOTICE.md and the README; the LICENSE
    the distribution ships is the repository's AGPL-3.0 text, byte for byte."""
    community = PACKAGES / "kitchensink4xl-community" / "LICENSE"
    assert community.read_bytes() == (ROOT / "LICENSE").read_bytes()


@pytest.mark.business_sources
def test_the_business_distribution_keeps_its_own_notice():
    business = PACKAGES / "kitchensink4xl-business" / "LICENSE"
    assert "KitchenSink4XL" in business.read_text(encoding="utf-8")


def test_the_shared_notices_still_carry_their_obligations():
    """Neutral naming, same obligations. Relabelling never drops a notice."""
    community = adapter_source.package_dir(adapter_source.COMMUNITY_POLICY_DIST)
    notice = (community / "NOTICE.md").read_text(encoding="utf-8")
    assert "Copyright (c) 2026 Alvut Consulting, LLC" in notice
    assert "KitchenSink4AI is a product line of" in notice
    assert "Not affiliated with, endorsed by, or sponsored by Microsoft" in notice
    assert "Required Notice: Copyright (c) 2026 Alvut Consulting, LLC" in notice
    assert "AGPL-3.0 (open source)" in notice
    agpl = (community / "LICENSE").read_text(encoding="utf-8")
    assert agpl.lstrip().startswith("GNU AFFERO GENERAL PUBLIC LICENSE")
    assert agpl.encode("utf-8") == (ROOT / "LICENSE").read_bytes()
    assert "GNU AFFERO GENERAL PUBLIC LICENSE" in agpl
    commercial = (adapter_source.package_dir(adapter_source.BUSINESS_POLICY_DIST)
                  / "LICENSE").read_text(encoding="utf-8")
    assert "Copyright (c) 2026 Alvut Consulting, LLC" in commercial
    assert "separate commercial agreement" in commercial
    # The clause wraps in the source file; match it the way it is written.
    assert "Third-party components retain their" in commercial
    assert "does not relicense third-party dependencies" in commercial


def test_the_microsoft_trademark_line_is_shared_and_stays_with_the_product():
    """The shared adapter notice carries one trademark line naming every
    Office product it serves; this repository's own notices name this product's
    format as well.
    """
    shared = (adapter_source.package_dir(adapter_source.COMMUNITY_POLICY_DIST)
              / "NOTICE.md").read_text(encoding="utf-8")
    assert "Microsoft, Word, Excel and PowerPoint are trademarks of the Microsoft group of" in shared
    for path in (ROOT / "NOTICE.md",
                 PACKAGES / "kitchensink4xl-community" / "NOTICE.md"):
        text = path.read_text(encoding="utf-8")
        assert "Microsoft, Excel, GitHub and Copilot are trademarks of the Microsoft group" in text


def test_the_adapters_are_consumed_not_copied():
    """This repository builds its adapter wheels from the canonical pinned
    source. A local copy would be a second payload under a shared name, which
    is the packaging defect the shared adapter exists to prevent."""
    for dist in adapter_source.ADAPTER_DISTS:
        assert not (PACKAGES / dist).exists()
    assert all(str(path).startswith(str(adapter_source.materialise()))
               for path in ADAPTERS.values())


@pytest.mark.parametrize("edition", ["business", "community"])
def test_a_shared_adapter_ships_no_generated_metadata(edition):
    """Generated per-product files would make the wheel product-specific
    while leaving its name and version alone, which is the exact thing the
    review forbids."""
    assert list(ADAPTERS[edition].rglob("_edition.json")) == []
    config = tomllib.loads(
        (ADAPTERS[edition].parents[1] / "pyproject.toml").read_text(encoding="utf-8"))
    package_data = config.get("tool", {}).get("setuptools", {}).get("package-data", {})
    assert package_data == {}


@pytest.mark.business_sources
def test_the_business_launcher_owns_every_product_specific_value():
    """All four product-specific values live in the launcher, which is the
    distribution that is allowed to be product-specific."""
    import ks4xl_business_launch as launch
    assert launch.PRODUCT_ID == "kitchensink4xl-business"
    assert launch.DISPLAY_NAME == "KitchenSink4XL Business"
    assert launch.LICENSE_FILE_ENV == "KS4XL_LICENSE_FILE"
    source = ast.parse((PACKAGES / "kitchensink4xl-business/src"
                        / "ks4xl_business_launch/__init__.py")
                       .read_text(encoding="utf-8"))
    assert "_edition.json" in ast.dump(source)


# ------------------------------------------------ two products, one wheel


def test_two_products_coexist_in_one_business_environment():
    """The coexistence case the review asked for: one adapter import, two
    configured products, neither disturbing the other."""
    word = BusinessPolicy(WORD, _build("kitchensink4word-business",
                                       "2.1.3+business"))
    xl = BusinessPolicy(XL, _build("kitchensink4xl-business"))
    assert word.product_id == "kitchensink4word-business"
    assert xl.product_id == "kitchensink4xl-business"
    assert word.distribution_version == "2.1.3+business"
    assert xl.distribution_version == "1.2.2+business"
    assert "KitchenSink4Word" in word.initialization_text()["introduction"]
    assert "KitchenSink4XL" in xl.initialization_text()["introduction"]
    assert type(word).__module__ == type(xl).__module__


def test_two_products_read_different_receipt_environment_variables(tmp_path,
                                                                   monkeypatch):
    receipt = (FIXTURES / "covered.json").read_bytes()
    path = tmp_path / "xl-receipt.json"
    path.write_bytes(receipt)
    monkeypatch.setenv("KS4XL_LICENSE_FILE", str(path))
    monkeypatch.delenv("KS4W_LICENSE_FILE", raising=False)
    word = BusinessPolicy(WORD, _build("kitchensink4word-business",
                                       "2.1.3+business"))
    xl = BusinessPolicy(XL, _build("kitchensink4xl-business"))
    assert xl.describe_license()["status"] == "verified"
    # The Word policy does not see the XL receipt: different variable.
    assert word.describe_license()["reason"] == "not_configured"


def test_a_community_adapter_introduces_whatever_product_it_is_given():
    word = CommunityPolicy(CommunityConfig(product_id="kitchensink4word",
                                           display_name="KitchenSink4Word"), "2.1.3")
    xl = CommunityPolicy(CommunityConfig(product_id="kitchensink4xl",
                                         display_name="KitchenSink4XL"), "1.2.2")
    assert word.initialization_text()["introduction"] == (
        "KitchenSink4Word (kitchensink4word on PyPI), part of the "
        "KitchenSink4AI suite. ")
    assert xl.initialization_text()["introduction"] == (
        "KitchenSink4XL (kitchensink4xl on PyPI), part of the "
        "KitchenSink4AI suite. ")
    # The guidance itself is product-independent and identical.
    assert (word.initialization_text()["instructions"]
            == xl.initialization_text()["instructions"]
            == notices.COMMUNITY_GUIDANCE)


# -------------------------------------------- receipts cannot cross products


def test_an_xl_receipt_is_refused_by_a_word_installation():
    raw = (FIXTURES / "covered.json").read_bytes()   # product: xl-business
    build = _build("kitchensink4word-business", "2.1.3+business")
    with pytest.raises(receipts.ReceiptError) as problem:
        receipts.verify_receipt(raw, build, WORD.product_id)
    assert problem.value.category == "product_mismatch"


def test_a_word_receipt_is_refused_by_an_xl_installation():
    raw = (FIXTURES / "wrong_product.json").read_bytes()  # product: word-business
    with pytest.raises(receipts.ReceiptError) as problem:
        receipts.verify_receipt(raw, MANIFEST["build"], XL.product_id)
    assert problem.value.category == "product_mismatch"
    # ...and the same bytes verify against the installation they belong to.
    assert receipts.verify_receipt(raw, _build("kitchensink4word-business",
                                               "2.1.3+business"),
                                   WORD.product_id)["status"] == "verified"


def test_a_receipt_cannot_nominate_the_product_it_is_checked_against(tmp_path,
                                                                     monkeypatch):
    """Self-selection is the failure mode parameterisation could have
    introduced. The expected product comes from configuration, and a receipt
    claiming a different one is refused rather than believed."""
    raw = (FIXTURES / "wrong_product.json").read_bytes()
    path = tmp_path / "receipt.json"
    path.write_bytes(raw)
    monkeypatch.setenv("KS4XL_LICENSE_FILE", str(path))
    status = BusinessPolicy(XL, _build("kitchensink4xl-business")).describe_license()
    assert (status["status"], status["reason"]) == ("unverified", "product_mismatch")


@pytest.mark.parametrize("expected", [None, "", "   ", 17, b"word"])
def test_verification_refuses_to_run_without_a_configured_product(expected):
    """There is no permissive default to fall into."""
    raw = (FIXTURES / "covered.json").read_bytes()
    with pytest.raises(ValueError, match="expected product"):
        receipts.verify_receipt(raw, MANIFEST["build"], expected)


def test_build_metadata_must_agree_with_the_launcher_configuration():
    with pytest.raises(BuildMetadataError, match="different product"):
        BusinessPolicy(XL, _build("kitchensink4word-business"))


@pytest.mark.parametrize("bad", [
    {"product_id": "", "display_name": "X", "license_file_env": "Y"},
    {"product_id": "x", "display_name": "  ", "license_file_env": "Y"},
    {"product_id": "x", "display_name": "X", "license_file_env": 42},
])
def test_a_malformed_product_config_is_refused_at_construction(bad):
    with pytest.raises(ProductConfigError):
        ProductConfig(**bad)


def test_a_business_policy_refuses_a_bare_dict_for_configuration():
    with pytest.raises(ProductConfigError, match="ProductConfig"):
        BusinessPolicy({"product_id": "kitchensink4xl-business"},
                       _build("kitchensink4xl-business"))
