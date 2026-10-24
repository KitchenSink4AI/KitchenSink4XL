"""Community bootstrap: build the Community policy, inject it, start the core.

Layer three. The launcher is the only place that knows which edition this
installation is, and it knows because of which distribution it belongs to,
never because of a file it found or an entry point it happened to win. It is
also the only place that knows which PRODUCT this is: the shared Community
policy adapter names no product, so the two values below are what make this
Excel rather than Word.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

DISTRIBUTION = "kitchensink4xl"
PRODUCT_ID = "kitchensink4xl"
DISPLAY_NAME = "KitchenSink4XL"

_NO_ADAPTER = (
    "This Community launcher needs the Community license policy, which is "
    "not installed in this Python environment. Install the supplied "
    "kitchensink4-license-community wheel here, or select the Python "
    "environment where the Community edition was installed."
)
_BOTH_EDITIONS = (
    "This Python environment contains both the Community and the Business "
    "license policy. An edition is an installation choice and the two are "
    "supported only in separate environments; neither launcher guesses. "
    "Install each edition in its own environment."
)


def installed_here(module: str) -> bool:
    """True when a module resolves inside the tree this launcher lives in.

    This is a co-location check on import provenance, NOT environment
    isolation. It answers one question: did the other edition's policy get
    installed into the same directory as this launcher? A deliberately
    constructed PYTHONPATH, an editable install, or a --target directory can
    still put the other adapter somewhere importable, and this check will not
    see it. It is a guard against the ordinary accident of pip-installing
    both editions into one site-packages, and it is not claimed to be more
    than that. Separate environments remain the supported arrangement.
    """
    spec = importlib.util.find_spec(module)
    if spec is None or not spec.origin:
        return False
    here = Path(__file__).resolve().parent.parent
    return Path(spec.origin).resolve().is_relative_to(here)


def _distribution_version() -> str:
    from importlib.metadata import PackageNotFoundError, version
    try:
        return version(DISTRIBUTION)
    except PackageNotFoundError:
        # Source checkout: the facade is not installed as a distribution, so
        # report the core it is running against rather than inventing one.
        from xlsx_mcp import __version__
        return __version__


def product_config():
    from ks4_license_community.policy import ProductConfig
    return ProductConfig(product_id=PRODUCT_ID, display_name=DISPLAY_NAME)


def build_policy():
    if importlib.util.find_spec("ks4_license_community") is None:
        raise SystemExit(_NO_ADAPTER)
    if installed_here("ks4_license_business"):
        raise SystemExit(_BOTH_EDITIONS)
    from ks4_license_community.policy import CommunityPolicy
    return CommunityPolicy(product_config(), _distribution_version())


def main() -> None:
    import sys
    if sys.argv[1:] == ["--version"]:
        # Prints the installed version and exits without starting the
        # server. The update (uv tool run --refresh-package ... --version)
        # reads this line to confirm which version uv now starts.
        print(f"{DISTRIBUTION} {_distribution_version()}")
        return
    from xlsx_mcp import license_policy
    license_policy.install(build_policy())
    from xlsx_mcp.server import main as serve
    serve()
