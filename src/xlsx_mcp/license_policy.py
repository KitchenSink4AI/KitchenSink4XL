"""The neutral licensing seam. Contains no notice prose, no receipt logic, and
no edition metadata of its own.

The core editing package never chooses an edition and never imports a policy.
An edition launcher constructs the policy object appropriate to the
distribution it belongs to and injects it here BEFORE importing the server.
A core that is started without one is a broken installation, and it says so
before any document is touched.

Two failure classes, deliberately kept apart:

  * setup errors (no policy installed, wrong policy_schema, missing member)
    raise PolicySetupError at import time, so no tool ever registers;
  * a missing or unverifiable Business receipt is a reported property of a
    correctly installed server and never blocks anything.

CONTRACT (duck-typed on purpose: an adapter implements these names and
imports nothing from this package, which is what keeps a notice adapter
structurally unable to reach the document APIs).

Every member below is type-checked at injection, exactly rather than by
isinstance, and a violation raises PolicyContractError before a tool exists.

Required:
    policy_schema          exact int, must equal POLICY_SCHEMA (bool refused)
    edition_id             non-empty str, one of SUPPORTED_EDITIONS
    product_id             non-empty str, the distribution name; also the
                           MCP server name
    distribution_version   non-empty str, the edition distribution's version
    describe_license(redacted: bool = True) -> dict
    initialization_text() -> dict with non-empty string values under the
                           keys "introduction" and "instructions"

Optional (when present, must be callable):
    additional_tool_registration(mcp) -> None
    update_check_suppression() -> dict | None
    update_backend() -> object | None

initialization_text returns LABELLED COMPONENTS rather than one assembled
string. The adapter owns the content and the wording of each component; the
core owns where each lands in its own instructions. Word, Excel and Web put
the licensing component last; PowerPoint puts it ahead of the product prose.
Both orders are expressible without any adapter special-casing.
"""
from __future__ import annotations

POLICY_SCHEMA = 1

#: The editions a core will start under. An adapter announcing anything else
#: is a mismatched or unknown build, not a new edition to accommodate.
SUPPORTED_EDITIONS = frozenset({"community", "business"})

_REQUIRED_STRINGS = ("edition_id", "product_id", "distribution_version")
_REQUIRED_CALLABLES = ("describe_license", "initialization_text")
_OPTIONAL_CALLABLES = ("additional_tool_registration",
                       "update_check_suppression", "update_backend")
_TEXT_COMPONENTS = ("introduction", "instructions")

_active = None


class PolicySetupError(RuntimeError):
    """Installation is incomplete or mismatched. Never a receipt problem."""


class PolicyMissingError(PolicySetupError):
    """Nothing was injected. The launcher did not run, or was bypassed."""


class PolicyContractError(PolicySetupError):
    """Something was injected, and it does not satisfy the contract.

    Kept distinct from PolicyMissingError because the operator actions differ:
    a missing policy means start through a launcher, a contract violation
    means the core and the adapter are not a matched pair. Neither is a
    receipt problem and neither is recoverable at document time.
    """


_NO_POLICY = (
    "No license policy is installed. xlsx_mcp is the shared editing core and "
    "does not run on its own: start the server through an edition launcher "
    "(the kitchensink4xl, xl-mcp or kitchensink4xl-business console command, "
    "or python -m ks4xl_community_launch / python -m ks4xl_business_launch). "
    "If you are running from a source checkout, install one edition adapter "
    "and its launcher into this environment first."
)

_MISMATCH = " Install a core and an adapter built for each other."


def _reject(policy, detail: str) -> None:
    raise PolicyContractError(
        f"License policy {type(policy).__name__} {detail}." + _MISMATCH)


def _nonempty_string(value) -> bool:
    return type(value) is str and bool(value.strip())


def _validate(policy) -> None:
    """Reject a malformed policy loudly, at injection, before anything runs.

    Types are checked exactly rather than by isinstance, because bool is a
    subclass of int and True == 1: an isinstance(schema, int) check accepts
    policy_schema=True and then compares equal to schema 1. Every value here
    is checked for the type it actually has to be.
    """
    schema = getattr(policy, "policy_schema", None)
    if type(schema) is not int:
        _reject(policy, f"declares a non-integer policy_schema {schema!r}")
    if schema != POLICY_SCHEMA:
        _reject(policy, f"declares policy_schema {schema!r}, but this core "
                        f"speaks schema {POLICY_SCHEMA}")
    for name in _REQUIRED_STRINGS:
        value = getattr(policy, name, None)
        if not _nonempty_string(value):
            _reject(policy, f"has {name}={value!r}, which is not a non-empty "
                            "string")
    if policy.edition_id not in SUPPORTED_EDITIONS:
        _reject(policy, f"declares the unknown edition_id "
                        f"{policy.edition_id!r}; this core supports "
                        f"{sorted(SUPPORTED_EDITIONS)}")
    for name in _REQUIRED_CALLABLES:
        if not callable(getattr(policy, name, None)):
            _reject(policy, f"is missing the required method {name!r}")
    for name in _OPTIONAL_CALLABLES:
        hook = getattr(policy, name, None)
        if hook is not None and not callable(hook):
            _reject(policy, f"has a non-callable {name!r}")
    text = policy.initialization_text()
    if not isinstance(text, dict):
        _reject(policy, f"returned {type(text).__name__} from "
                        "initialization_text(), which must return a dict")
    for component in _TEXT_COMPONENTS:
        if component not in text:
            _reject(policy, "returned initialization_text() without the "
                            f"{component!r} component")
        if not _nonempty_string(text[component]):
            _reject(policy, f"returned initialization_text()[{component!r}]="
                            f"{text[component]!r}, which is not a non-empty "
                            "string")


def install(policy) -> None:
    """Inject the prepared policy. Called by a launcher before server import."""
    _validate(policy)
    global _active
    _active = policy


def reset() -> None:
    """Drop the installed policy. For launcher and test harness use only."""
    global _active
    _active = None


def installed() -> bool:
    return _active is not None


def active():
    """The injected policy, or an actionable setup error."""
    if _active is None:
        raise PolicyMissingError(_NO_POLICY)
    return _active


def register_additional_tools(mcp) -> None:
    """Give the policy its optional registration hook. Most policies decline."""
    hook = getattr(active(), "additional_tool_registration", None)
    if hook is not None:
        hook(mcp)


def update_check_suppression() -> dict | None:
    """Extra fields a distribution wants reported when it suppresses the
    public update check, or None when the public channel applies. The core
    does not know which editions suppress it."""
    hook = getattr(active(), "update_check_suppression", None)
    return hook() if hook is not None else None


def update_backend():
    """Which update channel this distribution uses, or None for the public
    one.

    A policy that does not implement this returns None and the core uses
    its PyPI backend, which is why a Community install's update behaviour
    is unchanged by construction: the Community adapter implements nothing.

    The object a policy returns is whatever its own core knows how to use.
    The Business adapter returns a SELECTOR rather than a working backend,
    because that adapter wheel is product-neutral and imports nothing from
    any editing core, so the implementation has to live in the core and the
    adapter's job is to name it. A policy that does return a working
    backend (an object with a callable check) is used as given.
    """
    hook = getattr(active(), "update_backend", None)
    return hook() if hook is not None else None
