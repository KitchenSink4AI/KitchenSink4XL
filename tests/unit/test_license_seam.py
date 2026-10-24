"""The licensing seam: injection, setup errors, and adapter isolation.

Two properties are under test here and they are deliberately separate. A core
without a policy is a BROKEN INSTALLATION and must say so before a tool
exists to call. A Business receipt that is missing or unverifiable is a
REPORTED PROPERTY of a working installation and blocks nothing. Confusing
the two is the failure this architecture exists to prevent.
"""
import ast
import os
import subprocess
import sys
from pathlib import Path

import pytest

import adapter_source

from xlsx_mcp import license_policy

ROOT = Path(__file__).resolve().parents[2]
PACKAGES = ROOT / "packages"
#: The adapters are NOT in this repository. One canonical pair serves the
#: whole suite and is read from the pinned source; see scripts/adapter_source.
#: None on the public Community-only path, which carries no Business adapter; the
#: tests that read it are listed in tests/conftest.py BUSINESS_ONLY_TESTS.
BUSINESS_SRC = (adapter_source.package_dir(adapter_source.BUSINESS_POLICY_DIST) / "src"
                if adapter_source.edition() != "community" else None)
COMMUNITY_SRC = adapter_source.package_dir(
    adapter_source.COMMUNITY_POLICY_DIST) / "src"


class _Good:
    policy_schema = license_policy.POLICY_SCHEMA
    edition_id = "community"
    product_id = "kitchensink4xl"
    distribution_version = "0.0.0"

    def describe_license(self, redacted: bool = True):
        return {"edition": "community"}

    def initialization_text(self):
        return {"introduction": "I. ", "instructions": "L."}


def _variant(**members):
    return type("Variant", (_Good,), members)()


# ------------------------------------------------------------ setup errors


def test_a_core_without_a_policy_refuses_before_any_tool_registers(tmp_path):
    """Imported with no policy installed, the server module raises and no
    tool object is ever created. Run in a subprocess because the suite's own
    conftest injects a policy for every other test."""
    script = tmp_path / "probe.py"
    script.write_text(
        "import sys\n"
        f"sys.path[:0] = {[str(COMMUNITY_SRC), str(ROOT / 'src')]!r}\n"
        "from xlsx_mcp import license_policy\n"
        "try:\n"
        "    import xlsx_mcp.server\n"
        "except license_policy.PolicySetupError as problem:\n"
        "    print('SETUP_ERROR', problem)\n"
        "    print('SERVER_IMPORTED', 'xlsx_mcp.server' in sys.modules and "
        "hasattr(sys.modules.get('xlsx_mcp.server'), 'mcp'))\n"
        "else:\n"
        "    print('NO_ERROR')\n", encoding="utf-8")
    done = subprocess.run([sys.executable, "-X", "utf8", str(script)],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0, done.stderr
    assert "SETUP_ERROR" in done.stdout
    assert "SERVER_IMPORTED False" in done.stdout
    # The message has to tell an operator what to do, not just what failed.
    assert "launcher" in done.stdout
    assert "kitchensink4xl-business" in done.stdout


def test_setup_error_is_not_a_receipt_error():
    assert issubclass(license_policy.PolicySetupError, RuntimeError)
    from ks4_license_business.receipts import ReceiptError
    assert not issubclass(ReceiptError, license_policy.PolicySetupError)
    assert not issubclass(license_policy.PolicySetupError, ValueError)


@pytest.mark.parametrize("missing", ["policy_schema", "edition_id", "product_id",
                                     "distribution_version", "describe_license",
                                     "initialization_text"])
def test_an_incomplete_policy_is_a_setup_error(missing):
    with pytest.raises(license_policy.PolicyContractError, match=missing):
        license_policy._validate(_variant(**{missing: None}))


def test_a_mismatched_policy_schema_is_a_setup_error():
    with pytest.raises(license_policy.PolicyContractError, match="schema"):
        license_policy._validate(_variant(policy_schema=999))


def test_initialization_text_must_carry_both_labelled_components():
    broken = _variant(initialization_text=lambda self: {"introduction": "x"})
    with pytest.raises(license_policy.PolicyContractError, match="instructions"):
        license_policy._validate(broken)


# ------------------------------------- mistyped policy members are refused


def test_a_boolean_policy_schema_is_refused():
    """True == 1, so an isinstance(schema, int) check accepted policy_schema
    =True and then compared equal to schema 1. Types are checked exactly."""
    with pytest.raises(license_policy.PolicyContractError, match="non-integer"):
        license_policy._validate(_variant(policy_schema=True))
    assert license_policy.installed() is True  # nothing was swapped in


@pytest.mark.parametrize("member,value", [
    ("edition_id", 42), ("product_id", 17), ("distribution_version", 0),
])
def test_integer_identity_members_are_refused(member, value):
    with pytest.raises(license_policy.PolicyContractError,
                       match="not a non-empty string"):
        license_policy._validate(_variant(**{member: value}))


@pytest.mark.parametrize("text", [
    {"introduction": None, "instructions": []},
    {"introduction": "ok", "instructions": []},
    {"introduction": 3, "instructions": "ok"},
    {"introduction": "", "instructions": "ok"},
    {"introduction": "   ", "instructions": "ok"},
])
def test_non_string_initialization_components_are_refused(text):
    with pytest.raises(license_policy.PolicyContractError,
                       match="not a non-empty string"):
        license_policy._validate(_variant(initialization_text=lambda self: text))


def test_a_policy_with_every_member_mistyped_is_refused_by_install():
    """A policy with every member mistyped is refused at install() rather
    than accepted."""
    class Probe:
        policy_schema = True
        edition_id = 42
        product_id = 17
        distribution_version = 0

        def describe_license(self, redacted: bool = True):
            return {}

        def initialization_text(self):
            return {"introduction": None, "instructions": []}

    previous = license_policy.active()
    with pytest.raises(license_policy.PolicyContractError):
        license_policy.install(Probe())
    assert license_policy.active() is previous  # the good policy still stands


def test_initialization_text_that_is_not_a_dict_is_refused():
    broken = _variant(initialization_text=lambda self: "just a string")
    with pytest.raises(license_policy.PolicyContractError, match="must return a dict"):
        license_policy._validate(broken)


def test_an_unknown_edition_id_is_refused():
    with pytest.raises(license_policy.PolicyContractError, match="unknown edition_id"):
        license_policy._validate(_variant(edition_id="enterprise"))


@pytest.mark.parametrize("hook", ["additional_tool_registration",
                                  "update_check_suppression",
                                  "update_backend"])
def test_a_non_callable_optional_hook_is_refused(hook):
    with pytest.raises(license_policy.PolicyContractError, match=hook):
        license_policy._validate(_variant(**{hook: "not callable"}))


def test_the_two_setup_error_classes_are_distinguishable():
    """Different operator actions: start through a launcher, versus install a
    matched core and adapter."""
    assert issubclass(license_policy.PolicyMissingError, license_policy.PolicySetupError)
    assert issubclass(license_policy.PolicyContractError, license_policy.PolicySetupError)
    assert not issubclass(license_policy.PolicyContractError,
                          license_policy.PolicyMissingError)


def test_core_never_imports_a_default_policy():
    """No module in the core may import either adapter. The core receives a
    prepared object or it fails; it does not find one."""
    offenders = []
    for path in (ROOT / "src").rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), str(path))
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names = [node.module]
            for name in names:
                if name.split(".")[0] in {"ks4_license_community",
                                          "ks4_license_business"}:
                    offenders.append(f"{path.relative_to(ROOT)}: {name}")
    assert offenders == []


def test_no_licensing_prose_survives_in_the_core():
    """The guidance moved; it did not get copied."""
    from ks4_license_community import notices
    for needle in notices.guidance_needles():
        for path in (ROOT / "src").rglob("*.py"):
            assert needle not in path.read_text(encoding="utf-8"), path


# -------------------------------------------------------- adapter isolation


def _imports(root: Path) -> set[str]:
    found = set()
    for path in root.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                found.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                found.add(node.module.split(".")[0])
    return found


def test_business_package_imports_no_community_module():
    assert "ks4_license_community" not in _imports(BUSINESS_SRC)


def test_business_package_does_not_depend_on_the_editing_core():
    """Business must have no transitive dependency on Community, and the
    adapters reach the contract by shape rather than by import, so neither
    one can pull an editing surface in behind it."""
    assert "xlsx_mcp" not in _imports(BUSINESS_SRC)


def test_community_package_cannot_reach_a_document_api():
    reached = _imports(COMMUNITY_SRC)
    assert "xlsx_mcp" not in reached
    assert "ks4_license_business" not in reached
    assert not (reached & {"openpyxl", "lxml", "win32com"})


def test_community_package_declares_no_dependencies():
    import tomllib
    config = tomllib.loads((adapter_source.package_dir(
        adapter_source.COMMUNITY_POLICY_DIST)
        / "pyproject.toml").read_text(encoding="utf-8"))
    assert config["project"]["dependencies"] == []


def test_this_repository_carries_no_copy_of_a_shared_adapter():
    """One canonical adapter pair for the whole suite; a copy here would be
    a second thing wearing the same name."""
    for dist in adapter_source.ADAPTER_DISTS:
        assert not (PACKAGES / dist).exists()
    for module in ("ks4_license_community", "ks4_license_business"):
        assert list(ROOT.rglob(f"**/{module}/policy.py")) == []


def test_the_adapters_come_from_the_pinned_canonical_revision():
    """The build and the suite read the same pinned source, and it is read
    out of a commit rather than out of somebody's working tree."""
    assert len(adapter_source.PINNED_COMMIT) == 40
    described = adapter_source.describe()
    assert described["pinned_commit"] == adapter_source.PINNED_COMMIT
    stamp = Path(described["materialised"]) / ".pinned"
    # A CI export (the public run's Community adapter from the pinned public Word
    # commit) may carry no stamp; the recorded digest proves its provenance either way.
    if stamp.exists() or not os.environ.get(adapter_source.ENV_MATERIALISED):
        assert stamp.read_text(encoding="utf-8").strip() == adapter_source.PINNED_COMMIT
    assert (adapter_source.tree_digest(Path(described["materialised"]))
            == adapter_source.expected_digest())


def test_business_package_imports_cleanly_without_community_installed(tmp_path):
    script = tmp_path / "probe.py"
    script.write_text(
        "import sys\n"
        f"sys.path[:0] = {[str(BUSINESS_SRC)]!r}\n"
        "import ks4_license_business\n"
        "print('COMMUNITY_LOADED', any(m.startswith('ks4_license_community') "
        "for m in sys.modules))\n"
        "print('CORE_LOADED', any(m.startswith('xlsx_mcp') for m in sys.modules))\n",
        encoding="utf-8")
    done = subprocess.run([sys.executable, "-X", "utf8", str(script)],
                          capture_output=True, text=True, stdin=subprocess.DEVNULL)
    assert done.returncode == 0, done.stderr
    assert "COMMUNITY_LOADED False" in done.stdout
    assert "CORE_LOADED False" in done.stdout


# ------------------------------------------------------- text composition


def test_labelled_components_survive_either_assembly_order():
    """Word, Excel and Web put the licensing text last; PowerPoint puts it
    ahead of the product prose. One contract has to serve both without the
    adapter knowing which core it is talking to."""
    from ks4_license_community.policy import CommunityPolicy, ProductConfig
    text = CommunityPolicy(ProductConfig(product_id="kitchensink4xl",
                                         display_name="KitchenSink4XL"),
                           "1.2.2").initialization_text()
    prose = "PRODUCT PROSE."
    trailing_order = text["introduction"] + prose + " " + text["instructions"]
    ppt_order = text["introduction"] + text["instructions"] + " " + prose
    assert trailing_order.index(prose) < trailing_order.index(text["instructions"])
    assert ppt_order.index(text["instructions"]) < ppt_order.index(prose)
    assert text["introduction"] == trailing_order[:len(text["introduction"])]
    assert text["introduction"] == ppt_order[:len(text["introduction"])]


def test_additional_tool_registration_is_optional_and_neither_policy_uses_it():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    community = CommunityPolicy(
        CommunityConfig(product_id="kitchensink4xl",
                        display_name="KitchenSink4XL"), "1.2.2")
    business = BusinessPolicy(
        ProductConfig(product_id="kitchensink4xl-business",
                      display_name="KitchenSink4XL Business",
                      license_file_env="KS4XL_LICENSE_FILE"),
        {"edition": "business", "package": "kitchensink4xl-business",
         "version": "1.2.2+business",
         "release_date": "2026-09-10", "public_keys": {}})
    assert not hasattr(community, "additional_tool_registration")
    assert not hasattr(business, "additional_tool_registration")
    previous = license_policy.active()
    try:
        license_policy.install(community)
        license_policy.register_additional_tools(object())  # declines, no error
    finally:
        license_policy.install(previous)


def test_both_shipped_policies_satisfy_the_tightened_validator():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    license_policy._validate(CommunityPolicy(
        CommunityConfig(product_id="kitchensink4xl",
                        display_name="KitchenSink4XL"), "1.2.2"))
    license_policy._validate(BusinessPolicy(
        ProductConfig(product_id="kitchensink4xl-business",
                      display_name="KitchenSink4XL Business",
                      license_file_env="KS4XL_LICENSE_FILE"),
        {"edition": "business", "package": "kitchensink4xl-business",
         "version": "1.2.2+business", "release_date": "2026-09-10",
         "public_keys": {}}))


def test_only_the_business_policy_declares_an_update_backend():
    from ks4_license_business.policy import BusinessPolicy, ProductConfig
    from ks4_license_community.policy import CommunityPolicy
    from ks4_license_community.policy import ProductConfig as CommunityConfig
    community = CommunityPolicy(
        CommunityConfig(product_id="kitchensink4xl",
                        display_name="KitchenSink4XL"), "1.2.2")
    assert not hasattr(community, "update_check_suppression")
    assert not hasattr(community, "update_backend")
    business = BusinessPolicy(
        ProductConfig(product_id="kitchensink4xl-business",
                      display_name="KitchenSink4XL Business",
                      license_file_env="KS4XL_LICENSE_FILE"),
        {"edition": "business", "package": "kitchensink4xl-business",
         "version": "1.2.2+business",
         "release_date": "2026-09-10", "public_keys": {}})
    # The Business adapter dropped its suppression return when it gained a
    # channel of its own. Declaring both would report the check as off
    # while also knowing how to run it, which is a mismatched pair.
    assert not hasattr(business, "update_check_suppression")
    assert business.update_backend().backend_id == "business"
