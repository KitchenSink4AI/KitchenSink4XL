"""Build the split edition distributions and MCP bundles, without publishing.

One shared editing core, two license policy adapters, two launcher facades.
The Community app carries its core and launcher and requires its policy.
Business retains separate core, policy and launcher wheels. Nothing
selects an edition at runtime: the wheels in the bundle are the choice.

The adapters are NOT part of this repository. They are the one canonical pair
for the whole suite, consumed from the pinned source that adapter_source
resolves, so the wheel this build emits is byte-for-byte the same adapter
every other editor ships.

The anchored fail-closed rewrite discipline is unchanged. Every textual
rewrite goes through replace_once(), which raises when an anchor is missing
or ambiguous, and the staged metadata is parsed before it is built so a
malformed or duplicated table fails here rather than in a customer's
environment. A Business build additionally passes a deterministic scan gate
that proves no Community guidance string reached any shipped artifact while
the legally required notices survived.
"""
from __future__ import annotations
import argparse
import base64
import importlib.util
import json
import re
import shutil
import subprocess
import sys
import tempfile
import tomllib
import zipfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = ROOT / "packages"

CORE_DIST = "kitchensink4xl-core"
COMMUNITY_POLICY_DIST = adapter_source.COMMUNITY_POLICY_DIST
BUSINESS_POLICY_DIST = adapter_source.BUSINESS_POLICY_DIST

#: License expressions, and the rule about which may appear where. One shared
#: source revision produces two legally labelled core artifacts: the public
#: one at X.Y.Z under AGPL, and the commercial one at X.Y.Z+commercial under
#: the commercial grant. Different artifact identities, so nothing ever ships
#: two wheels with the same name and version, different contents and
#: different grants.
AGPL = "AGPL-3.0-only"
COMMERCIAL = "LicenseRef-KitchenSink-Commercial"
COMMERCIAL_CORE_SUFFIX = "+commercial"
#: A distribution must not contain a wheel granted under the other license.
FORBIDDEN_LICENSE = {"community": COMMERCIAL, "business": AGPL}

#: Notices that MUST survive in a Business artifact. "No Community guidance"
#: never means deleting a required copyright, trademark or dependency notice,
#: so the gate asserts these are present as well as asserting the guidance is
#: absent.
REQUIRED_NOTICES = (
    "Copyright (c) 2026 Alvut Consulting, LLC",
    "Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.",
)

#: Artifact members whose bytes are not text and are never scanned.
_BINARY_SUFFIXES = {".png", ".ico", ".pyc", ".whl", ".mcpb", ".so", ".pyd"}

# ------------------------------------------------------- the release channel

#: Release channels a build may declare, matching the policy adapter's
#: ALLOWED_CHANNELS. The adapter is the enforcing copy; this one exists so a
#: bad --channel is rejected at the command line rather than at a customer's
#: first startup.
CHANNELS = ("beta", "stable")
DEFAULT_CHANNEL = "beta"

#: Words a Business customer must never read in a shipped artifact. The
#: product is a beta and "beta" is the word for it. The other five each claim
#: something different and none of them is true here: "preview" and "RC"
#: promise a finished build behind them on a schedule nobody has published,
#: "trial" and "demo" say the licence expires or the software is crippled, and
#: a customer who reads any of them has been told the wrong thing about what
#: they paid for. The gate is a spelling check on a commercial promise.
_PRERELEASE_WORDS = re.compile(
    r"\b(?:previews?|previewing|release[\s-]+candidates?|trials?|demos?)\b",
    re.IGNORECASE)
#: "RC" only as a standalone uppercase token, so ordinary words that contain
#: the letters are not swept up. Case-sensitive on purpose: the marketing term
#: is capitalised, and lowercase "rc" appears in ordinary technical text.
_PRERELEASE_RC = re.compile(r"(?<![A-Za-z0-9])RC(?![A-Za-z0-9])")

#: Members of a Business artifact set that the language gate does NOT read,
#: each with the reason it is out of scope. This list is the gate's honesty:
#: everything named here still carries pre-release language today, and every
#: entry is a document whose wording belongs to somebody other than the build
#: script. Nothing is allowlisted for being inconvenient.
PRERELEASE_ALLOWLIST = {
    "NOTICE.md":
        "third-party and trademark notices; obligations, not product copy",
    "CHANGELOG.md":
        "release history, which must be free to say what a past release was",
    "LICENSE-BETA.txt":
        "the beta licence terms, shipped verbatim with every bundle and "
        "every receipt envelope; it must be "
        "byte-identical wherever it appears",
}

#: The members the gate DOES read, by suffix or exact name, out of the
#: Business artifact set. Narrow on purpose: these are the surfaces the build
#: script itself composes, plus the generated metadata, and a customer reads
#: every one of them.
def _prerelease_targets(label: str) -> bool:
    """True when this artifact member is customer-facing build-authored text."""
    name = label.rsplit("::", 1)[-1]
    base = name.rsplit("/", 1)[-1]
    if base in PRERELEASE_ALLOWLIST:
        return False
    if base in ("LICENSE", "INSTALL.md"):
        # The commercial licence notice and the filtered install text. Both
        # were allowlisted while they still called this a development
        # preview; both now say beta release, so both are read.
        return True
    if base == "manifest.json":
        return True
    if base == "_edition.json":
        return True
    if base == "METADATA" and name.endswith(".dist-info/METADATA"):
        return True
    return False


def find_prerelease_language(text: str) -> list[str]:
    """Every pre-release word in one string, in the order they appear."""
    return ([match.group(0) for match in _PRERELEASE_WORDS.finditer(text)]
            + [match.group(0) for match in _PRERELEASE_RC.finditer(text)])


def scan_prerelease_language(paths, exempt_text=()) -> dict:
    """Prove no shipped Business artifact calls this build anything but a beta.

    exempt_text holds documents that are embedded verbatim inside a scanned
    member: the commercial notice inside the generated launcher README, and
    the repository README inside every wheel's METADATA. Subtracting the
    embedded document keeps the gate on the build script's own words without
    allowlisting the whole member, so new pre-release language in the
    generated half is still caught.

    Line endings are normalised on both sides before the subtraction. A
    document read through Path.read_text arrives with LF and the same
    document read out of a zip member keeps whatever it was written with, and
    a gate that silently failed to subtract would report the embedded
    document's own words as a leak.
    """
    exempt_text = tuple(e.replace("\r\n", "\n") for e in exempt_text if e)
    leaks: list[str] = []
    scanned = 0
    for path in paths:
        path = Path(path)
        for label, text in _members(path):
            if not _prerelease_targets(label):
                continue
            scanned += 1
            text = text.replace("\r\n", "\n")
            for embedded in exempt_text:
                text = text.replace(embedded, "")
            for word in find_prerelease_language(text):
                leaks.append(f"{label}: {word!r}")
    if not scanned:
        raise BuildScanError(
            "The pre-release language gate read no customer-facing members, "
            "which means it would pass anything.")
    if leaks:
        raise BuildScanError(
            "A shipped Business artifact calls this build something other "
            "than a beta:\n  " + "\n  ".join(leaks))
    return {"members_scanned": scanned,
            "allowlisted": sorted(PRERELEASE_ALLOWLIST)}




class BuildScanError(RuntimeError):
    """A shipped Business artifact failed the deterministic content gate."""


def replace_once(text: str, old: str, new: str) -> str:
    """Fail closed if source metadata drifts or an anchor becomes ambiguous."""
    if text.count(old) != 1:
        raise ValueError(f"Expected exactly one build metadata anchor: {old!r}")
    return text.replace(old, new, 1)


def staged_config(path: Path, rewrites: list[tuple[str, str]]) -> str:
    """Apply anchored rewrites and prove the result is still valid metadata."""
    text = path.read_text("utf-8")
    for old, new in rewrites:
        text = replace_once(text, old, new)
    # Reject duplicate tables or malformed staged metadata before building.
    tomllib.loads(text)
    return text


# ------------------------------------------------------------- the scan gate


def community_needles() -> tuple[str, ...]:
    """Load the forbidden sentences from the Community notice module itself.

    Read out of the canonical pinned source file rather than an installed
    package, so the gate cannot pass because the adapter happened not to be
    installed, and cannot drift when the approved wording is revised.
    """
    source = adapter_source.notices_source()
    spec = importlib.util.spec_from_file_location("_scan_notices", source)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.guidance_needles()


def _members(path: Path):
    """Yield (label, text) for every scannable member of an artifact."""
    if path.suffix in {".whl", ".mcpb", ".zip"}:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if Path(name).suffix in _BINARY_SUFFIXES or name.endswith("/"):
                    continue
                yield f"{path.name}::{name}", archive.read(name).decode(
                    "utf-8", errors="replace")
    elif path.suffix not in _BINARY_SUFFIXES:
        yield path.name, path.read_text("utf-8", errors="replace")


#: Owner-test mode. An unfilled copy slot or a builder's home folder in
#: any shipped member fails the build of an owner-test (or release) artifact.
#: The served strings (instructions, tool text, help, hints) are also scanned
#: for these and for em dashes by scripts/smoke_editions.py, on the INSTALLED
#: wheels, because raw source members legitimately carry em dashes in
#: comments and the word PLACEHOLDER in code.
OWNER_TEST_MARKERS = ("COPY" "_SLOT:", "COPY PENDING", "FACTS TO CONVEY")


def _home_needles() -> tuple[str, ...]:
    home = Path.home()
    forms = {str(home), home.as_posix(), str(home).replace("\\", "\\\\")}
    return tuple(sorted(f for f in forms if len(f) > 3))


def scan_owner_test(paths) -> dict:
    """Refuse any artifact member that carries an unfilled copy slot or the
    building machine's home folder. Returns the scan report; raises
    SystemExit naming every hit."""
    hits, scanned = [], 0
    homes = _home_needles()
    for path in paths:
        for label, text in _members(Path(path)):
            scanned += 1
            for marker in OWNER_TEST_MARKERS:
                if marker in text:
                    hits.append(f"{label}: {marker}")
            # surface/copy.py assembles its markers at run time (the literals
            # are written split so a FILLED file passes), so a static scan of
            # the member never sees them. An unfilled slot is any remaining
            # call of the slot helper outside its own definition.
            if label.replace("\\", "/").endswith("surface/copy.py"):
                calls = [ln for ln in text.splitlines()
                         if "_slot(" in ln and not ln.lstrip().startswith("def _slot(")]
                if calls:
                    hits.append(f"{label}: {len(calls)} unfilled copy slot(s)")
            low = text.lower()
            for home in homes:
                if home.lower() in low:
                    hits.append(f"{label}: home folder path")
    if hits:
        sep = chr(10) + "  "
        more = f"{sep}... and {len(hits) - 40} more" if len(hits) > 40 else ""
        raise SystemExit("owner-test ship gate refused the build:" + sep
                         + sep.join(hits[:40]) + more)
    return {"members_scanned": scanned, "markers": list(OWNER_TEST_MARKERS)}


def scan_business_artifacts(paths, needles=None, required=REQUIRED_NOTICES) -> dict:
    """Prove a Business artifact set carries no Community guidance and still
    carries its required notices. Raises rather than warning."""
    needles = tuple(needles if needles is not None else community_needles())
    if not needles:
        raise BuildScanError("The scan gate found no Community needles to "
                             "look for, which means it would pass anything.")
    leaks: list[str] = []
    found_required = {text: False for text in required}
    scanned = 0
    for path in paths:
        for label, text in _members(Path(path)):
            scanned += 1
            for needle in needles:
                if needle in text:
                    leaks.append(f"{label}: {needle[:60]}...")
            for text_required in found_required:
                if text_required in text:
                    found_required[text_required] = True
    missing = [t for t, seen in found_required.items() if not seen]
    if leaks:
        raise BuildScanError(
            "Community guidance reached a Business artifact:\n  "
            + "\n  ".join(leaks))
    if missing:
        raise BuildScanError(
            "Required notices are missing from the Business artifacts:\n  "
            + "\n  ".join(missing))
    return {"members_scanned": scanned, "needles": len(needles),
            "required_notices": len(required)}


def wheel_license(path: Path) -> tuple[str, str]:
    """The license expression a wheel declares, read from its METADATA."""
    with zipfile.ZipFile(path) as archive:
        names = [n for n in archive.namelist() if n.endswith(".dist-info/METADATA")]
        if len(names) != 1:
            raise BuildScanError(f"{path.name} has {len(names)} METADATA files")
        declared = ""
        for line in archive.read(names[0]).decode("utf-8").splitlines():
            if line.startswith("License-Expression:"):
                declared = line.split(":", 1)[1].strip()
            elif line.startswith("License:") and not declared:
                declared = line.split(":", 1)[1].strip()
            elif not line.strip():
                break
    return path.name, declared


def scan_license_metadata(paths, edition: str) -> dict:
    """No distribution ships a wheel granted under the other edition's license.

    This is the check that keeps a Business bundle from carrying an
    AGPL-labelled core wheel beside a commercial launcher, which would put
    two conflicting grants in one customer's hands and make the shipped
    metadata, not the agreement, the thing a customer reads.
    """
    forbidden = FORBIDDEN_LICENSE[edition]
    declared, wrong = {}, []
    for path in paths:
        path = Path(path)
        if path.suffix != ".whl":
            continue
        name, expression = wheel_license(path)
        declared[name] = expression
        if not expression:
            wrong.append(f"{name}: declares no license expression")
        elif forbidden in expression:
            wrong.append(f"{name}: declares {expression}, forbidden in a "
                         f"{edition} distribution")
    if not declared:
        raise BuildScanError("The license gate found no wheels to check, "
                             "which means it would pass anything.")
    if wrong:
        raise BuildScanError(
            f"Conflicting license metadata in the {edition} distribution:\n  "
            + "\n  ".join(wrong))
    return declared


def business_install_doc() -> str:
    """The Business install text, filtered out of the shared editions doc.

    Mechanical section removal, not a rewrite: the Community section is the
    one that carries guidance prose, so it is the one that does not ship.
    """
    lines = (ROOT / "docs/editions.md").read_text("utf-8").splitlines(keepends=True)
    out, skipping = [], False
    for line in lines:
        if line.startswith("## "):
            skipping = line.strip() == "## Community"
        if not skipping:
            out.append(line)
    return "".join(out)


# ------------------------------------------------------------ wheel builders


def _build_wheel(stage: Path, output: Path) -> Path:
    before = set(output.glob("*.whl"))
    subprocess.run([sys.executable, "-m", "build", "--wheel", "--no-isolation",
                    "--outdir", str(output), str(stage)], check=True,
                   stdin=subprocess.DEVNULL)
    made = set(output.glob("*.whl")) - before
    if len(made) != 1:
        raise ValueError(f"Expected exactly one new wheel, got {sorted(made)}")
    return made.pop()


def build_core(edition: str, output: Path, scratch: Path) -> tuple[Path, str]:
    """The editing core, from ONE source revision, labelled for its edition.

    Community gets X.Y.Z under AGPL, which is what goes to PyPI. Business gets
    X.Y.Z+commercial under the commercial grant, with the commercial notice as
    its LICENSE. Same tested source, two artifact identities, no wheel that
    says AGPL sitting inside a commercial bundle.

    NOTICE.md ships in BOTH. It carries the Alvut copyright and the Microsoft
    trademark disclaimer, and those obligations do not change with the grant.
    Relabelling a distribution never means dropping a notice.
    """
    stage = scratch / "core"
    stage.mkdir()
    shutil.copytree(ROOT / "src", stage / "src",
                    ignore=shutil.ignore_patterns("__pycache__", "*.egg-info"))
    for name in ("README.md", "LICENSE", "NOTICE.md"):
        shutil.copy2(ROOT / name, stage / name)
    version = BASE_VERSION
    rewrites = []
    if edition == "business":
        version = BASE_VERSION + COMMERCIAL_CORE_SUFFIX
        rewrites = [
            (f'version = "{BASE_VERSION}"', f'version = "{version}"'),
            (f'license = "{AGPL}"', f'license = "{COMMERCIAL}"'),
        ]
        notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
        (stage / "LICENSE").write_text(notice, encoding="utf-8")
    (stage / "pyproject.toml").write_text(
        staged_config(ROOT / "pyproject.toml", rewrites), encoding="utf-8")
    return _build_wheel(stage, output), version


def build_policy_adapter(edition: str, output: Path, scratch: Path) -> tuple[Path, str]:
    """A license policy adapter wheel, built from the ONE canonical source.

    This repository carries no copy of either adapter. The source comes from
    the pinned canonical revision, so the wheel built here has the same
    contents as the one every other editor in the suite ships. Nothing
    product-specific is generated into it.
    """
    name = COMMUNITY_POLICY_DIST if edition == "community" else BUSINESS_POLICY_DIST
    source = adapter_source.package_dir(name)
    stage = scratch / name
    shutil.copytree(source, stage, ignore=shutil.ignore_patterns(
        "__pycache__", "*.egg-info", "_edition.json"))
    version = tomllib.loads(
        (source / "pyproject.toml").read_text("utf-8"))["project"]["version"]
    return _build_wheel(stage, output), version


def build_launcher(edition: str, output: Path, scratch: Path, release_date: date,
                   keys: dict, core_version: str,
                   policy_version: str,
                   channel: str = DEFAULT_CHANNEL, roots: dict | None = None,
                   trust_list: dict | None = None,
                   require_chain: bool = False) -> tuple[Path, str]:
    """The edition distribution: the public console names, the pinned wheels,
    and, for Business, the product configuration and generated build metadata.

    The metadata lands HERE rather than in the shared adapter, because a
    generated _edition.json inside the adapter would make that wheel
    product-specific while keeping its name and version.
    """
    name = f"kitchensink4xl-{edition}"
    source = PACKAGES / name
    stage = scratch / name
    shutil.copytree(source, stage, ignore=shutil.ignore_patterns(
        "__pycache__", "*.egg-info", "_edition.json"))
    policy_dist = (COMMUNITY_POLICY_DIST if edition == "community"
                   else BUSINESS_POLICY_DIST)
    rewrites = [
        (f'    "{policy_dist}",', f'    "{policy_dist}=={policy_version}",'),
    ]
    if edition == "business":
        rewrites.append((f'    "{CORE_DIST}",',
                         f'    "{CORE_DIST}=={core_version}",'))
    else:
        if policy_version != "0.2.0":
            raise ValueError("Community requires license notices version 0.2.0")
        shutil.copytree(ROOT / "src", stage / "src", dirs_exist_ok=True,
                        ignore=shutil.ignore_patterns("__pycache__", "*.egg-info"))
        for notice in ("LICENSE", "NOTICE.md"):
            shutil.copy2(ROOT / notice, stage / notice)
    version = BASE_VERSION
    if edition == "business":
        version = f"{BASE_VERSION}+business"
        rewrites.append((f'version = "{BASE_VERSION}"', f'version = "{version}"'))
    (stage / "pyproject.toml").write_text(
        staged_config(source / "pyproject.toml", rewrites), encoding="utf-8")
    shutil.copy2(ROOT / "README.md", stage / "README.md")
    if edition == "business":
        notice = (ROOT / "business/COMMERCIAL_LICENSE_NOTICE.md").read_text("utf-8")
        (stage / "README.md").write_text(
            "# KitchenSink4XL Business\n\nKitchenSink4AI Business, beta "
            "release. Built and verified like production, licensed for "
            "evaluation under the Pilot Evaluation License.\n\n" + notice,
            encoding="utf-8")
        (stage / "src/ks4xl_business_launch/_edition.json").write_text(
            json.dumps({
                "edition": "business",
                # Must match the launcher's own PRODUCT_ID; the adapter
                # refuses the pair when it does not.
                "package": "kitchensink4xl-business",
                "version": version,
                # The stage of the product, carried as data so the server
                # can report it and the adapter can refuse a value that is
                # not one of the two. It is deliberately NOT a PEP 440
                # pre-release suffix on the version: the wheel version stays
                # X.Y.Z+business so that pip resolution, the installer's
                # expectations and the published identity of the artifact do
                # not move with the label.
                "channel": channel,
                "release_date": release_date.isoformat(),
                "public_keys": keys,
                # THE ROOTS ARE THE PIN. A signing key can be rotated by
                # issuing it a certificate under an intermediate; a root
                # cannot, which is why only roots are frozen into an artifact
                # and why they are kept offline. public_keys above stays for
                # the receipts already issued under the flat model and comes
                # out a release after every one of them has been re-signed.
                "root_keys": dict(roots or {}),
                # The baseline signed list, carried so a build knows which
                # keys have been pulled without asking anything over a
                # network. A newer list rides in a receipt or an update.
                "trust_list": trust_list,
                # accept_flat false is the switch that ends the compatibility
                # window. It is set at build time and read at verify time; an
                # installation cannot change it, because an installation that
                # could relax it could accept a receipt this artifact was not
                # built to accept.
                "chain_policy": {"accept_flat": not require_chain},
            }, indent=2), encoding="utf-8")
    return _build_wheel(stage, output), version


BASE_VERSION = tomllib.loads(
    (ROOT / "pyproject.toml").read_text("utf-8"))["project"]["version"]


# ------------------------------------------------------------------- bundles


#: The key id a single unnamed --public-key gets, unchanged from before the
#: flags became repeatable.
DEFAULT_KEY_ID = "primary"


def _public_key_b64(path) -> str:
    """One Ed25519 public key file as raw base64, PEM or .pub.

    The issuer keys are handed over as PEM and the ROOT keys as the base64
    .pub files keygen.py writes, so the loader that reads a root has to accept
    both rather than make an operator convert a key by hand at the one moment
    where a mistake is expensive. The 32 raw bytes are the same either way and
    that is what a build records.
    """
    blob = Path(path).read_bytes()
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey

    if b"-----BEGIN" in blob:
        key = serialization.load_pem_public_key(blob)
        if not isinstance(key, Ed25519PublicKey):
            raise ValueError("An Ed25519 public key is required")
        return base64.b64encode(key.public_bytes(
            serialization.Encoding.Raw,
            serialization.PublicFormat.Raw)).decode("ascii")
    try:
        raw = base64.b64decode(b"".join(blob.split()), validate=True)
    except (ValueError, TypeError) as problem:
        raise ValueError(f"{path} is neither a PEM nor a base64 public key: "
                         f"{problem}") from problem
    if len(raw) != 32:
        raise ValueError(f"{path} holds {len(raw)} bytes, not the 32 bytes of an "
                         "Ed25519 public key")
    return base64.b64encode(raw).decode("ascii")


def root_keys(root_key, root_id) -> dict:
    """Map root key id to base64 public key, from paired repeatable arguments.

    Paired POSITIONALLY like --public-key and --key-id, and for the same
    reason. A build carries the global root AND the pre-trusted spare, so that
    losing the global root is a trust-list update rather than a recall: a
    product that already trusts the spare can be told to move to it by a list
    the spare signs. Both are named on the command line, because a root a
    build trusts and nobody wrote down is a root nobody can audit.
    """
    paths = ([] if root_key is None
             else [root_key] if isinstance(root_key, (str, Path))
             else list(root_key))
    ids = ([] if root_id is None
           else [root_id] if isinstance(root_id, str)
           else list(root_id))
    if not paths:
        return {}
    if len(ids) != len(paths):
        raise ValueError(
            f"{len(paths)} root key(s) and {len(ids)} root id(s). --root-key "
            "and --root-id are paired in order, so each root needs exactly "
            "one id.")
    keys: dict = {}
    for path, name in zip(paths, ids):
        if not isinstance(name, str) or not name.strip() or len(name) > 80:
            raise ValueError(f"Invalid root key id: {name!r}")
        if name in keys:
            raise ValueError(
                f"Duplicate root key id {name!r}. Each id names exactly one "
                "root, and a build that anchors a key the command line did "
                "not name is a build nobody can audit.")
        keys[name] = _public_key_b64(path)
    return keys


def load_trust_list(path, roots: dict) -> dict:
    """Read a signed trust list and prove it against the roots being built in.

    Proved HERE, at build time, rather than only in the adapter, because a
    build that ships a list its own roots do not vouch for is an artifact that
    fails at a customer's first status call. The adapter checks it again at
    verification time; this check is the one that stops the artifact from
    existing.
    """
    adapter_source.ensure_importable()
    from ks4_license_business import chain

    data = json.loads(Path(path).read_text(encoding="utf-8"))
    try:
        chain.verify_trust_list(data, roots=roots)
    except chain.ChainError as problem:
        raise ValueError(f"{path} does not verify against the roots given: "
                         f"{problem}") from problem
    # The SIGNED object is what a build embeds, not the body it wraps: a
    # product has to be able to check that signature for itself.
    return data


def issuer_keys(public_key, key_id) -> dict:
    """Map key id to base64 raw public key, from paired repeatable arguments.

    --public-key and --key-id are paired POSITIONALLY: the nth id names the
    nth key. A build may need several, because the purchase flow has two
    issuers. The web tier signs self-serve receipts on the Worker with one
    key; the offline root signs manual and enterprise receipts with another.
    An installation cannot know which counter a customer bought at, so it
    carries both and the receipt's own key id selects between them.

    One pair is the ordinary case and behaves exactly as it did before this
    function existed, including the default id when no --key-id is given.
    Duplicate ids are refused rather than resolved: a map cannot hold one id
    twice, and silently keeping the last one would mean a build that trusts a
    different key than the command line named.
    """
    paths = ([] if public_key is None
             else [public_key] if isinstance(public_key, (str, Path))
             else list(public_key))
    ids = ([] if key_id is None
           else [key_id] if isinstance(key_id, str)
           else list(key_id))
    if not paths:
        return {}
    if not ids:
        ids = [DEFAULT_KEY_ID]
    if len(ids) != len(paths):
        raise ValueError(
            f"{len(paths)} public key(s) and {len(ids)} key id(s). "
            "--public-key and --key-id are paired in order, so each key needs "
            "exactly one id.")
    keys: dict[str, str] = {}
    for path, name in zip(paths, ids):
        if not isinstance(name, str) or not name.strip() or len(name) > 80:
            raise ValueError(f"Invalid issuer key id: {name!r}")
        if name in keys:
            raise ValueError(
                f"Duplicate issuer key id {name!r}. Each id names exactly one "
                "public key, and a build that trusts a key the command line "
                "did not name is a build nobody can audit.")
        # PEM or the one-line base64 .pub the key ceremony actually writes.
        # Reading only PEM here cost the first attempt at the beta build, and
        # the two encodings carry the same 32 bytes, so refusing one of them
        # bought nothing.
        keys[name] = _public_key_b64(path)
    return keys


def build(edition: str, output: Path, release_date: date, public_key=None,
          key_id=DEFAULT_KEY_ID, channel: str = DEFAULT_CHANNEL,
          root_key=None, root_id=None, trust_list=None,
          require_chain: bool = False, owner_test: bool = False) -> list[Path]:
    """Build one edition into a NEW output folder, then run its gates.

    A REFUSED BUILD LEAVES NOTHING. The gates run on the finished artifacts,
    so they are written first; when any gate refuses (the owner-test scan, the
    Business scan, the language or the license gate) or the build fails part
    way, the output folder this call created is removed before the refusal
    propagates. A refused build used to leave its wheels and .mcpb in
    --output, looking exactly like a passing build's output."""
    if channel not in CHANNELS:
        raise ValueError(f"channel must be one of {CHANNELS}, got {channel!r}")
    business = edition == "business"
    keys = issuer_keys(public_key, key_id)
    roots = root_keys(root_key, root_id)
    if require_chain and not roots:
        raise ValueError(
            "--require-chain without --root-key would build an edition that "
            "refuses flat receipts and cannot check chained ones, which is an "
            "edition that verifies nothing.")
    trust = load_trust_list(trust_list, roots) if trust_list is not None else None
    # Temp build roots are OS scratch, never inside the source checkout.
    output.mkdir(parents=True, exist_ok=False)  # this call owns it from here
    try:
        return _build_into(edition, output, release_date, keys, channel, roots, trust,
                           require_chain, owner_test, business)
    except BaseException:
        remove_refused_output(output)
        raise


def remove_refused_output(output: Path) -> None:
    """Delete an output folder a refused build created. Only ever
    called on the folder build() made itself with exist_ok=False."""
    shutil.rmtree(output, ignore_errors=True)
    if output.exists():  # pragma: no cover (a locked file on Windows)
        leftovers = sorted(p.name for p in output.rglob("*") if p.is_file())
        print(f"refused build: could not remove {output}: {leftovers}", file=sys.stderr)


def _build_into(edition: str, output: Path, release_date: date, keys, channel, roots, trust,
                require_chain: bool, owner_test: bool, business: bool) -> list[Path]:
    with tempfile.TemporaryDirectory(prefix="ks4xl-edition-") as scratch_dir:
        scratch = Path(scratch_dir)
        core, core_version = (build_core(edition, output, scratch)
                              if business else (None, BASE_VERSION))
        policy, policy_version = build_policy_adapter(edition, output, scratch)
        launcher, version = build_launcher(
            edition, output, scratch, release_date, keys, core_version,
            policy_version, channel, roots, trust, require_chain)
    wheels = ([core] if business else []) + [policy, launcher]
    package = "kitchensink4xl-business" if business else "kitchensink4xl"
    module = "ks4xl_business_launch" if business else "ks4xl_community_launch"

    manifest = json.loads((ROOT / "bundle/manifest.json").read_text("utf-8"))
    manifest.update(name=package, display_name=f"KitchenSink4XL {edition.title()}",
                    version=version, license=("LicenseRef-KitchenSink-Commercial"
                    if business else "AGPL-3.0-only"))
    manifest["description"] = (
        f"KitchenSink4XL {edition.title()}, beta release. Install the "
        "supplied wheels and register the launcher as described in INSTALL.md.")
    manifest["long_description"] = manifest["description"]
    manifest["author"] = {"name": "KitchenSink4AI"}
    manifest["server"]["mcp_config"]["command"] = "${user_config.python_path}"
    manifest["server"]["mcp_config"]["args"] = ["-m", module]
    # An edition bundle runs its own wheels, not the public Desktop bundle,
    # so it must not announce the bundle install route (core/update_flow.py).
    manifest["server"]["mcp_config"]["env"].pop("KS4XL_INSTALL_ROUTE", None)
    manifest["user_config"]["python_path"] = {
        "type": "string", "title": f"{edition.title()} Python executable",
        "description": "Absolute path to Python in the dedicated environment containing this edition.",
        "required": True,
    }
    if business:
        manifest["user_config"]["license_file"] = {
            "type": "string", "title": "Business license receipt",
            "description": "Path to the offline signed JSON receipt, if supplied.",
            "required": False, "default": "",
        }
        manifest["server"]["mcp_config"]["env"]["KS4XL_LICENSE_FILE"] = "${user_config.license_file}"
    archive = output / f"{package}-{version}.mcpb"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        bundle.writestr("manifest.json", json.dumps(manifest, indent=2))
        bundle.write(ROOT / "bundle/icon.png", "icon.png")
        bundle.write(ROOT / ("business/COMMERCIAL_LICENSE_NOTICE.md" if business else "LICENSE"),
                     "LICENSE")
        if business:
            # The shared editions doc carries Community guidance prose, so the
            # Business bundle ships the filtered install text instead.
            bundle.writestr("INSTALL.md", business_install_doc())
        else:
            bundle.write(ROOT / "docs/editions.md", "INSTALL.md")
        for wheel in wheels:
            bundle.write(wheel, wheel.name)
    artifacts = wheels + [archive]
    if business:
        report = scan_business_artifacts(artifacts)
        print(f"scan gate: {report['members_scanned']} members, "
              f"{report['needles']} needles, all required notices present")
        language = scan_prerelease_language(artifacts)
        print(f"language gate: {language['members_scanned']} customer-facing "
              f"members carry no pre-release wording; allowlisted verbatim "
              f"documents: {', '.join(language['allowlisted'])}")
    # Both editions are checked. A commercial wheel loose in the public
    # distribution is the same defect wearing the other hat.
    if owner_test:
        report = scan_owner_test(artifacts)
        print(f"owner-test gate: {report['members_scanned']} members carry no copy slot "
              "and no home folder")
    declared = scan_license_metadata(artifacts, edition)
    print(f"license gate ({edition}): "
          + ", ".join(f"{n.split('-')[0]}={e}" for n, e in sorted(declared.items())))
    return artifacts


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--edition", choices=["community", "business", "both"], default="both")
    parser.add_argument("--output", type=Path, required=True,
                        help="New destination: existing edition directories are never overwritten")
    parser.add_argument("--release-date", type=date.fromisoformat, required=True)
    parser.add_argument("--public-key", type=Path, action="append",
                        help="Issuer public PEM only; omitted builds cannot verify "
                             "receipts. Repeatable: pair each one with a --key-id "
                             "in the same order")
    parser.add_argument("--key-id", action="append",
                        help=f"Id for the public key in the same position "
                             f"(default: {DEFAULT_KEY_ID} for a single key)")
    parser.add_argument("--channel", choices=CHANNELS, default=DEFAULT_CHANNEL,
                        help="the release stage recorded in the Business build "
                             "metadata and reported by get_server_info "
                             f"(default: {DEFAULT_CHANNEL})")
    parser.add_argument("--root-key", type=Path, action="append",
                        help="Global root public key, PEM or base64 .pub. "
                             "Repeatable: pair each one with a --root-id in the "
                             "same order. A build trusts these and nothing else "
                             "at the top of a certificate chain")
    parser.add_argument("--root-id", action="append",
                        help="Id for the root key in the same position")
    parser.add_argument("--trust-list", type=Path,
                        help="Signed baseline trust list to embed; it must be "
                             "signed by one of the --root-key roots")
    parser.add_argument("--owner-test", action="store_true",
                        help="Owner-test mode: refuse any artifact carrying an unfilled "
                             "copy slot or this machine's home folder")
    parser.add_argument("--require-chain", action="store_true",
                        help="Refuse receipts that carry no certificate chain. "
                             "Off by default while receipts issued under the "
                             "flat model are still in customers' hands")
    args = parser.parse_args()
    for name in (args.key_id or []):
        if not name.strip() or len(name) > 80:
            parser.error("invalid key id")
    if args.key_id and not args.public_key:
        parser.error("--key-id names a --public-key; supply the key it names")
    for name in (args.root_id or []):
        if not name.strip() or len(name) > 80:
            parser.error("invalid root id")
    if args.root_id and not args.root_key:
        parser.error("--root-id names a --root-key; supply the key it names")
    if args.root_key and not args.root_id:
        parser.error("--root-key needs a --root-id: a root nobody named is a "
                     "root nobody can audit")
    if args.trust_list and not args.root_key:
        parser.error("--trust-list needs the --root-key roots it is signed by")
    try:
        # Pair them here, before a wheel is built, so a mismatched or
        # duplicated pair is a command-line error rather than a half-built
        # output directory. The trust list is proved against the roots in the
        # same pass, for the same reason.
        issuer_keys(args.public_key, args.key_id)
        roots = root_keys(args.root_key, args.root_id)
        if args.require_chain and not roots:
            parser.error("--require-chain needs --root-key")
        if args.trust_list is not None:
            load_trust_list(args.trust_list, roots)
    except ValueError as problem:
        parser.error(str(problem))
    print("license policy adapters: "
          + json.dumps(adapter_source.describe(), indent=2))
    editions = ("community", "business") if args.edition == "both" else (args.edition,)
    built: list[Path] = []
    try:
        for edition in editions:
            target = args.output.resolve() / edition
            artifacts = build(edition, target, args.release_date, args.public_key,
                              args.key_id, args.channel, args.root_key, args.root_id,
                              args.trust_list, args.require_chain,
                              owner_test=args.owner_test)
            built.append(target)
            for artifact in artifacts:
                print(artifact)
    except BaseException:
        # With --edition both, an edition that passed is removed too
        # when a later one is refused; a refused run ships nothing.
        for target in built:
            remove_refused_output(target)
        raise


if __name__ == "__main__":
    main()
