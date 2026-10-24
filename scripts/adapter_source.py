"""Locate the ONE canonical shared license policy adapter source.

The distributions kitchensink4-license-community and
kitchensink4-license-business are a single pair of packages for the whole
KitchenSink4AI suite, not one pair per editor. Four editors publishing a
wheel called kitchensink4-license-business at one version with four different
payloads inside would make pip, its cache and any co-installation resolve one
name to four different things. So this repository carries NO copy of either
adapter. It consumes them, and every artifact it builds consumes them, from
the canonical source at a pinned revision.

THE CANONICAL SOURCE, TODAY. The adapters were written in the Word editions
checkout and live under its packages/ directory. This module reads them out
of that repository at PINNED_COMMIT using git, so the content is the pinned
revision no matter what the Word working tree currently holds, and no matter
which branch that checkout happens to have in place. Nothing is read from a
working tree.

INTERIM ARRANGEMENT. Reaching into a sibling product's checkout is an
interim arrangement, not the destination. The adapters belong in their own
repository (or their own publish pipeline) that all four editors depend on.
When they move there, only the resolution below changes and nothing that
calls into it does.

EDITIONS. KS4_ADAPTER_EDITION selects what is resolved. Unset (or "all")
is the private two-adapter path every Business build uses: both packages,
checked against ADAPTER_DIGEST, unchanged. "community" is the public path:
the public repositories carry and build the Community edition only and never
receive the Business adapter source, so only kitchensink4-license-community is
resolved, imported and digested, against COMMUNITY_DIGEST. Asking for the
Business adapter on that path is an error, not a fallback.

PUBLIC CI. The public repository reads the Community adapter from the
reviewed public release commit PUBLIC_ADAPTER_COMMIT of the license repository,
KitchenSink4AI/kitchensink4-license-community, the one repository that publishes
kitchensink4-license-community (`python scripts/adapter_source.py
--public-pin`). Until that commit exists the constant is empty and the pin
refuses, so no public run can use an unpinned or moving source.

PUBLISHED WHEEL. A release never uploads the adapter and never rebuilds it
for PyPI. It requires the exact wheel the publisher put on PyPI, pinned by
COMMUNITY_WHEEL_FILENAME and COMMUNITY_WHEEL_SHA256 (`python
scripts/adapter_source.py --fetch-wheel DIR` downloads and checks it). A
rebuild carries the build backend of its day, so its bytes could differ
from the published wheel's; the pin removes that.

RESOLUTION ORDER
    1. KS4_LICENSE_ADAPTERS   an already-materialised directory holding the
                              two adapter package directories. Set this to
                              point at a dedicated adapter repository once one
                              exists, or at a vendored export in CI.
    2. KS4_WORD_CHECKOUT      a Word editions git checkout to read the pinned
                              commit out of.
    3. the sibling checkout   ../../Word/editions relative to THIS repository's
                              main checkout (resolved through git, so running
                              from a linked worktree still finds it).

Materialisation goes to an OS scratch directory keyed by the pinned commit,
never inside this checkout.
"""
from __future__ import annotations

import hashlib
import os
import re
import subprocess
import tempfile
import zipfile
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

#: The reviewed adapter revision. Bumping this is a deliberate act with a
#: re-review attached, which is the entire point of pinning it. What a pin
#: has to hold:
#:
#: - The adapters are product-neutral. Their Python, NOTICE.md and LICENSE
#:   name no single editor, because one pair of packages ships into all of
#:   them.
#: - Their LICENSE and NOTICE files are part of what is pinned. If one
#:   editor materialised a stale LICENSE while another built the current one
#:   from its own tree, one distribution would carry two payloads, and the
#:   bundle build refuses that.
#: - A change to the hooks an adapter declares moves its distribution
#:   version, so a build can refuse an old core paired with a new adapter
#:   (the old core would call a hook that is gone and fall back to the
#:   public PyPI channel on a Business install).
#: - The pinned commit is reachable from a clean clone of the canonical
#:   checkout. A commit that only a local history holds would make the pin
#:   unreachable everywhere else.
#:
#: At this pin the Business adapter validates the build's release channel
#: against a closed set of two (beta, stable), and selects the issuer public
#: key by the id the receipt names, refusing an unknown id with the existing
#: signature_invalid category instead of letting a KeyError fall through to
#: "unreadable". It declares update_backend(), a selector that names the
#: licensed release channel for the core to run. The Community adapter
#: implements neither hook.
#: b2b404c: the Community adapter is the package of its own repository,
#: kitchensink4-license-community, which differs from the previous pin only in
#: comment wording and its PyPI author line. The Business adapter differs from
#: the previous pin only in comment wording.
#: 9c7dc45: the Community adapter is kitchensink4-license-community 0.2.0 (the plain AGPL
#: LICENSE text, trademark lines in NOTICE.md, tidier comments); the Business adapter is
#: unchanged.
PINNED_COMMIT = "9c7dc45404e294a63b685f4186db52769a7ed2e3"

#: Content digest of the two adapter package trees at PINNED_COMMIT, with
#: PowerPoint's tree_digest algorithm (sorted relative paths and file
#: SHA-256s). Recorded rather than computed, so a materialised tree that is
#: not the pinned content (a CI export, KS4_LICENSE_ADAPTERS, a stale cache)
#: fails loudly instead of becoming this product's private copy. The value is
#: IDENTICAL to PowerPoint's recorded ADAPTER_DIGEST: both products pin the
#: same adapter content.
ADAPTER_DIGEST = "70829e50d60509fc71daf0c49a3cdb0d4039e656ad6b8d4f1d46a009ad4fdac7"

#: Content digest of the Community adapter package ALONE at PINNED_COMMIT
#: (same algorithm, Community directory only). The Community-only path is
#: checked against this value and never against ADAPTER_DIGEST.
COMMUNITY_DIGEST = "ce7c493e264397b127fe07a0529469ea662f0b7f470af1f418afcd4d1ea2e1c5"

COMMUNITY_POLICY_DIST = "kitchensink4-license-community"
BUSINESS_POLICY_DIST = "kitchensink4-license-business"
ADAPTER_DISTS = (COMMUNITY_POLICY_DIST, BUSINESS_POLICY_DIST)

ENV_MATERIALISED = "KS4_LICENSE_ADAPTERS"
ENV_CHECKOUT = "KS4_WORD_CHECKOUT"
#: "community" selects the Community-only path; unset or "all" keeps both.
ENV_EDITION = "KS4_ADAPTER_EDITION"
EDITIONS = ("all", "community")

#: Where public CI reads the Community adapter: the publisher's reviewed public
#: release commit carrying packages/kitchensink4-license-community.
#: Full 40-character SHA. Empty until that release exists; public_pin() refuses.
PUBLIC_ADAPTER_REPOSITORY = "KitchenSink4AI/kitchensink4-license-community"
PUBLIC_ADAPTER_COMMIT = "c33f1a4cda744bd4f95234993e37f03c27705278"

#: The exact kitchensink4-license-community wheel on PyPI that releases require,
#: as the publisher's post-verify job prints it. Empty until that release exists;
#: pinned_wheel() refuses.
COMMUNITY_WHEEL_FILENAME = "kitchensink4_license_community-0.2.0-py3-none-any.whl"
COMMUNITY_WHEEL_SHA256 = "4ad5bf6ae2a677a2b51b9c62b7f1a44464940350f8e8c9ce04635a91ea73ff81"
PYPI_JSON_URL = "https://pypi.org/pypi/{name}/{version}/json"
PYPI_FILES_HOST = "files.pythonhosted.org"

#: Where the adapters sit inside the canonical checkout.
_PREFIX = "packages"

_NOT_FOUND = (
    "The canonical license policy adapters could not be located. This "
    "repository deliberately carries no copy of them (see adapter_source's "
    "module docstring). Point {env} at a directory containing "
    "{dists}, or point {checkout} at a Word editions checkout that contains "
    "commit {commit}."
).format(env=ENV_MATERIALISED, dists=" and ".join(ADAPTER_DISTS),
         checkout=ENV_CHECKOUT, commit=PINNED_COMMIT[:7])


class AdapterSourceError(RuntimeError):
    """The pinned canonical adapter source is unavailable or not at the pin."""


def edition() -> str:
    """'all' (private, both adapters) or 'community' (public, Community only)."""
    value = (os.environ.get(ENV_EDITION) or "").strip().lower() or "all"
    if value not in EDITIONS:
        raise AdapterSourceError(
            f"{ENV_EDITION}={value!r} is not one of {', '.join(EDITIONS)}")
    return value


def required_dists() -> tuple[str, ...]:
    """The adapter packages the selected edition resolves."""
    if edition() == "community":
        return (COMMUNITY_POLICY_DIST,)
    return ADAPTER_DISTS


def expected_digest() -> str:
    return COMMUNITY_DIGEST if edition() == "community" else ADAPTER_DIGEST


def public_pin() -> dict:
    """Where public CI checks the Community adapter out from. Refuses unless pinned."""
    if not re.fullmatch(r"[0-9a-f]{40}", PUBLIC_ADAPTER_COMMIT):
        raise AdapterSourceError(
            "PUBLIC_ADAPTER_COMMIT is not a full 40-character commit SHA. Public "
            "builds read the Community adapter from the reviewed public "
            f"{PUBLIC_ADAPTER_REPOSITORY} release commit; set it (and "
            "COMMUNITY_DIGEST if the adapter content changed) in a reviewed edit.")
    return {"repository": PUBLIC_ADAPTER_REPOSITORY, "commit": PUBLIC_ADAPTER_COMMIT,
            "path": f"{_PREFIX}/{COMMUNITY_POLICY_DIST}", "digest": COMMUNITY_DIGEST}


def pinned_wheel() -> dict:
    """The exact published adapter wheel a release requires. Refuses unless pinned."""
    m = re.fullmatch(r"kitchensink4_license_community-([0-9]+(?:\.[0-9]+)*)-py3-none-any\.whl",
                     COMMUNITY_WHEEL_FILENAME)
    if not m or not re.fullmatch(r"[0-9a-f]{64}", COMMUNITY_WHEEL_SHA256):
        raise AdapterSourceError(
            "COMMUNITY_WHEEL_FILENAME and COMMUNITY_WHEEL_SHA256 are not set. Releases require "
            "the exact kitchensink4-license-community wheel its publisher put on PyPI; pin the "
            "values its post-verify job printed, in a reviewed edit.")
    return {"filename": COMMUNITY_WHEEL_FILENAME, "sha256": COMMUNITY_WHEEL_SHA256,
            "version": m.group(1)}


def _https_get(url: str) -> bytes:
    import urllib.request
    request = urllib.request.Request(url, headers={"User-Agent": "ks4-adapter-source/2"})
    with urllib.request.urlopen(request, timeout=60) as response:  # noqa: S310 (https only)
        return response.read()


def fetch_pinned_wheel(destination: Path, get=_https_get) -> Path:
    """Download the pinned wheel from PyPI into destination; refuse any other bytes."""
    import json
    import urllib.parse
    pin = pinned_wheel()
    meta = json.loads(get(PYPI_JSON_URL.format(name=COMMUNITY_POLICY_DIST, version=pin["version"])))
    entries = [f for f in meta.get("urls", []) if f.get("filename") == pin["filename"]]
    if len(entries) != 1:
        raise AdapterSourceError(
            f"{pin['filename']} is not a file of {COMMUNITY_POLICY_DIST} {pin['version']} on PyPI")
    url = entries[0].get("url", "")
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme != "https" or parsed.hostname != PYPI_FILES_HOST:
        raise AdapterSourceError(f"unexpected download URL {url!r}")
    if (entries[0].get("digests") or {}).get("sha256") != pin["sha256"]:
        raise AdapterSourceError("PyPI lists a different sha256 for the pinned wheel")
    data = get(url)
    if hashlib.sha256(data).hexdigest() != pin["sha256"]:
        raise AdapterSourceError("the downloaded wheel does not match COMMUNITY_WHEEL_SHA256")
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / pin["filename"]
    target.write_bytes(data)
    return target


def _git(args, cwd=None) -> str:
    done = subprocess.run(["git", *args], cwd=str(cwd) if cwd else None,
                          capture_output=True, text=True,
                          stdin=subprocess.DEVNULL)
    if done.returncode != 0:
        raise AdapterSourceError(
            f"git {' '.join(args)} failed: {done.stderr.strip()}")
    return done.stdout.strip()


def _checkout_bases() -> list[Path]:
    """Every directory this repository is checked out into.

    A linked worktree's parents are an OS scratch path, so walking up from
    __file__ finds nothing on its own. The common git directory belongs to the
    main checkout, and `git worktree list` names the rest, including the
    checkout whose neighbouring folders hold the Word checkout that the
    sibling lookup below needs.
    """
    bases: list[Path] = []
    try:
        common = _git(["-C", str(ROOT), "rev-parse", "--path-format=absolute",
                       "--git-common-dir"])
        bases.append(Path(common).parent)
    except AdapterSourceError:
        return bases
    try:
        for line in _git(["-C", str(ROOT), "worktree", "list",
                          "--porcelain"]).splitlines():
            if line.startswith("worktree "):
                bases.append(Path(line.split(" ", 1)[1]))
    except AdapterSourceError:
        pass
    return bases


#: Where a sibling Word checkout sits relative to one of our own checkouts.
#: Both layouts are real: the Word checkout one folder down inside a
#: sibling folder of a product checkout ("*" is any one folder), and the
#: Excel and Word editions folders side by side under one parent.
_SIBLING_LAYOUTS = (
    ("..", "*", "Word", "editions"),
    ("..", "..", "Word", "editions"),
)


def _layout_folders(base: Path, layout: tuple[str, ...]) -> list[Path]:
    """The existing folders a layout names from `base`; "*" stands for every
    folder at that level, in name order."""
    found = [base]
    for part in layout:
        if part != "*":
            found = [p / part for p in found]
            continue
        level = []
        for p in found:
            try:
                level.extend(sorted(c for c in p.iterdir() if c.is_dir()))
            except OSError:
                continue
        found = level
    return [p for p in found if p.is_dir()]


def canonical_checkout() -> Path:
    """The git checkout holding the canonical adapters at the pinned commit."""
    override = (os.environ.get(ENV_CHECKOUT) or "").strip()
    candidates = [Path(override)] if override else []
    seen = set()
    expanded = set()
    for base in _checkout_bases():
        for layout in _SIBLING_LAYOUTS:
            # Checkouts side by side share a parent: each one is listed once.
            start = os.path.abspath(base.joinpath(layout[0]))
            key = (os.path.normcase(start), layout[1:])
            if key in expanded:
                continue
            expanded.add(key)
            for candidate in _layout_folders(Path(start), layout[1:]):
                try:
                    candidate = candidate.resolve()
                except OSError:
                    continue
                if candidate not in seen:
                    seen.add(candidate)
                    candidates.append(candidate)
    for candidate in candidates:
        if not candidate.is_dir():
            continue
        try:
            if _git(["-C", str(candidate), "cat-file", "-t",
                     PINNED_COMMIT]) == "commit":
                return candidate
        except AdapterSourceError:
            continue
    raise AdapterSourceError(_NOT_FOUND)


def _cache_root() -> Path:
    suffix = "-community" if edition() == "community" else ""
    return Path(tempfile.gettempdir()) / f"ks4-license-adapters-{PINNED_COMMIT[:12]}{suffix}"


def tree_digest(source: Path, dists: tuple[str, ...] | None = None) -> str:
    """Content digest over every adapter file, path included and sorted
    (PowerPoint's algorithm, so the two products' recorded digests compare).
    Covers the selected edition's packages unless dists is given."""
    parts = []
    for dist in (dists or required_dists()):
        root = source / dist
        for path in sorted(q for q in root.rglob("*") if q.is_file()):
            if "__pycache__" in path.parts or path.suffix == ".pyc":
                continue
            parts.append((path.relative_to(source).as_posix(),
                          hashlib.sha256(path.read_bytes()).hexdigest()))
    if not parts:
        raise AdapterSourceError(f"No adapter files found under {source}")
    return hashlib.sha256(repr(sorted(parts)).encode()).hexdigest()


def _verified(root: Path) -> Path:
    found = tree_digest(root)
    expected = expected_digest()
    if found != expected:
        raise AdapterSourceError(
            f"The license adapters at {root} are not the pinned content "
            f"({edition()} edition): expected "
            f"digest {expected}, found {found}. Either the pin moved without "
            "ADAPTER_DIGEST moving with it, or the adapters were edited outside the "
            "canonical source.")
    return root


def materialise(refresh: bool = False) -> Path:
    """Return a directory holding both adapter package directories.

    Extraction is from the pinned COMMIT, not from a working tree, so a dirty
    or differently-branched Word checkout cannot change what gets built here.
    The result is cached in OS scratch so a suite run extracts once.
    """
    override = (os.environ.get(ENV_MATERIALISED) or "").strip()
    if override:
        root = Path(override)
        missing = [d for d in required_dists() if not (root / d).is_dir()]
        if missing:
            raise AdapterSourceError(
                f"{ENV_MATERIALISED}={root} does not contain {missing}")
        return _verified(root)

    cache = _cache_root()
    stamp = cache / ".pinned"
    if stamp.is_file() and not refresh:
        return _verified(cache)
    checkout = canonical_checkout()
    cache.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ks4xl-adapter-export-") as scratch:
        bundle = Path(scratch) / "adapters.zip"
        # -c core.autocrlf=false -c core.eol=lf: git archive applies the
        # working-tree end-of-line conversion, so on a checkout with
        # autocrlf on it hands back CRLF copies of LF blobs. The adapters
        # are ONE shared source for four products; a product whose export
        # differs from the committed bytes is a private copy of the shared
        # wheel, which is the defect the pin exists to prevent. The sibling
        # products read the same commit with cat-file semantics and get the
        # blob, so this export has to as well.
        _git(["-c", "core.autocrlf=false", "-c", "core.eol=lf",
              "-C", str(checkout), "archive", "--format=zip",
              "--output", str(bundle), PINNED_COMMIT,
              *[f"{_PREFIX}/{dist}" for dist in required_dists()]])
        with zipfile.ZipFile(bundle) as archive:
            archive.extractall(scratch)
        for dist in required_dists():
            source = Path(scratch) / _PREFIX / dist
            if not source.is_dir():
                raise AdapterSourceError(
                    f"{dist} is absent from {checkout} at {PINNED_COMMIT[:7]}")
            target = cache / dist
            if target.exists():
                import shutil
                shutil.rmtree(target)
            source.replace(target)
    stamp.write_text(PINNED_COMMIT + "\n", encoding="utf-8")
    return _verified(cache)


def package_dir(dist: str) -> Path:
    """The adapter distribution's source directory (holds its pyproject)."""
    if dist not in ADAPTER_DISTS:
        raise AdapterSourceError(f"{dist} is not a license policy adapter")
    if dist not in required_dists():
        raise AdapterSourceError(
            f"{dist} is not available on the {edition()} edition path "
            f"({ENV_EDITION}); the public repositories carry the Community adapter only")
    return materialise() / dist


def import_paths() -> list[str]:
    """sys.path entries that make the selected edition's adapters importable."""
    return [str(package_dir(dist) / "src") for dist in required_dists()]


def ensure_importable() -> list[str]:
    """Put the canonical adapters on sys.path, ahead of anything installed."""
    import sys
    added = []
    for entry in import_paths():
        if entry not in sys.path:
            sys.path.insert(0, entry)
            added.append(entry)
    return added


def notices_source() -> Path:
    """The approved Community notice module, read straight from the pin."""
    return (package_dir(COMMUNITY_POLICY_DIST) / "src"
            / "ks4_license_community" / "notices.py")


def describe() -> dict:
    """What a report or a build log should say about where this came from."""
    try:
        checkout = str(canonical_checkout())
    except AdapterSourceError:
        checkout = ""
    return {"pinned_commit": PINNED_COMMIT, "checkout": checkout,
            "edition": edition(), "materialised": str(materialise())}


if __name__ == "__main__":
    import json
    if sys.argv[1:] == ["--public-pin"]:
        for key, value in public_pin().items():
            print(f"{key}={value}")
    elif len(sys.argv) == 3 and sys.argv[1] == "--fetch-wheel":
        print(fetch_pinned_wheel(Path(sys.argv[2])))
    else:
        print(json.dumps(describe(), indent=2))
