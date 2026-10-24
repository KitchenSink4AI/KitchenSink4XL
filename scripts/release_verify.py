#!/usr/bin/env python3
"""Release identity, dependency and PyPI payload checks, without uploading.

The Excel Community release contains app/ (kitchensink4xl, including the
editing core and launcher) and license/ (kitchensink4-license-community).
The app pins license notices version 0.2.0 exactly and must not require a
separate editing-core distribution. With --product license it checks the single
license wheel.

check-build verifies canonical tag/source/wheel versions and dependency pins.
plan checks PyPI identity and payload equivalence before deciding upload/skip.
verify-remote polls for the same exact published wheel names and payloads.
ZIP timestamps and offsets do not affect unpacked payload equivalence.

Run with --help for each subcommand. Standard library plus packaging.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import time
import tomllib
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from email.parser import HeaderParser
from pathlib import Path

from packaging.requirements import InvalidRequirement, Requirement
from packaging.utils import canonicalize_name, parse_wheel_filename
from packaging.version import InvalidVersion, Version

LICENSE_DIST = "kitchensink4-license-community"
PRODUCTS = {
    "word": ("kitchensink4word-core", "kitchensink4word"),
    "xl": ("kitchensink4xl-core", "kitchensink4xl"),
    "ppt": ("kitchensink4ppt-core", "kitchensink4ppt"),
}
KINDS = ("license", "core", "launcher")
COMMUNITY_KINDS = ("app", "license")
#: The shared adapter's own workflow publishes one distribution: the Community adapter.
LICENSE_LANE = "license"
LICENSE_PYPROJECT = "packages/kitchensink4-license-community/pyproject.toml"
TAG_PREFIXES = ("v", "license-v")


def kinds_for(product: str) -> tuple[str, ...]:
    if product == LICENSE_LANE:
        return ("license",)
    return COMMUNITY_KINDS if product == "xl" else KINDS


VERSION_RE = re.compile(
    r"(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)"
    r"(?:(a|b|rc)(0|[1-9]\d*))?(?:\.post(0|[1-9]\d*))?(?:\.dev(0|[1-9]\d*))?")
DEFAULT_INDEX = "https://pypi.org"
FILES_HOST = "files.pythonhosted.org"


class ReleaseCheckError(Exception):
    pass


def fail(msg: str):
    raise ReleaseCheckError(msg)


# ------------------------------------------------------------------ identity
def canonical_version(text: str, what: str) -> str:
    if not VERSION_RE.fullmatch(text):
        fail(f"{what} {text!r} is not a canonical release version")
    try:
        if str(Version(text)) != text:
            fail(f"{what} {text!r} is not in PEP 440 normal form")
    except InvalidVersion:
        fail(f"{what} {text!r} is not a valid version")
    return text


def version_from_tag(tag: str, prefix: str = "v") -> str:
    if prefix not in TAG_PREFIXES:
        fail(f"tag prefix {prefix!r} is not one of {TAG_PREFIXES}")
    if not tag.startswith(prefix) or tag[len(prefix):].startswith("v"):
        fail(f"release tag {tag!r} must be {prefix!r} followed by the version")
    return canonical_version(tag[len(prefix):], "release tag version")


def source_versions(root: Path, product: str) -> dict[str, object]:
    if product == LICENSE_LANE:
        p = root / LICENSE_PYPROJECT
        if not p.is_file():
            fail(f"required version source {LICENSE_PYPROJECT} is missing")
        return {LICENSE_PYPROJECT: tomllib.loads(p.read_text(encoding="utf-8"))
                .get("project", {}).get("version")}
    core, launcher = PRODUCTS[product]
    found: dict[str, object] = {}
    for rel in ("pyproject.toml", f"packages/{launcher}-community/pyproject.toml"):
        p = root / rel
        if not p.is_file():
            fail(f"required version source {rel} is missing")
        found[rel] = tomllib.loads(p.read_text(encoding="utf-8")).get("project", {}).get("version")
    sj = root / "server.json"
    if not sj.is_file():
        fail("required version source server.json is missing")
    server = json.loads(sj.read_text(encoding="utf-8"))
    found["server.json version"] = server.get("version")
    packages = server.get("packages") or []
    if not packages:
        fail("server.json has no packages[] entry")
    for i, pkg in enumerate(packages):
        found[f"server.json packages[{i}].version"] = pkg.get("version")
        if pkg.get("registryType") == "pypi" and pkg.get("identifier") != launcher:
            fail(f"server.json packages[{i}] names {pkg.get('identifier')!r}, expected {launcher!r}")
    bm = root / "bundle" / "manifest.json"
    if not bm.is_file():
        fail("required version source bundle/manifest.json is missing")
    found["bundle/manifest.json"] = json.loads(bm.read_text(encoding="utf-8")).get("version")
    return found


# ------------------------------------------------------------------ wheels
def wheel_metadata(path: Path):
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if n.endswith(".dist-info/METADATA") and n.count("/") == 1]
        if len(names) != 1:
            fail(f"{path.name}: expected one top-level .dist-info/METADATA, found {names}")
        text = z.read(names[0]).decode("utf-8").replace("\r\n", "\n")
    msg = HeaderParser().parsestr(text)
    return msg.get("Name", ""), msg.get("Version", ""), msg.get_all("Requires-Dist") or []


def one_wheel(dist: Path, kind: str) -> Path:
    d = dist / kind
    files = sorted(p for p in d.iterdir()) if d.is_dir() else []
    if len(files) != 1 or files[0].suffix != ".whl" or not files[0].is_file():
        fail(f"{kind}: expected exactly one .whl in {d}, found {[p.name for p in files]}")
    return files[0]


def local_set(dist: Path, product: str) -> dict[str, dict]:
    """{kind: {path, name, version, filename}} after strict naming checks."""
    kinds = kinds_for(product)
    extra = sorted(p.name for p in dist.iterdir() if p.name not in kinds) if dist.is_dir() else ["<missing>"]
    if extra:
        fail(f"unexpected entries in {dist}: {extra}")
    want = {"license": LICENSE_DIST}
    if product != LICENSE_LANE:
        core, launcher = PRODUCTS[product]
        if product == "xl":
            want["app"] = launcher
        else:
            want.update(core=core, launcher=launcher)
    out = {}
    for kind in kinds:
        path = one_wheel(dist, kind)
        try:
            fn_name, fn_ver, _build, _tags = parse_wheel_filename(path.name)
        except Exception as e:  # noqa: BLE001
            fail(f"{kind}: bad wheel filename {path.name}: {e}")
        name, version, requires = wheel_metadata(path)
        if canonicalize_name(name) != want[kind] or fn_name != want[kind]:
            fail(f"{kind}: wheel is {name!r} ({path.name}), expected {want[kind]}")
        canonical_version(version, f"{kind} wheel version")
        if str(fn_ver) != version:
            fail(f"{kind}: filename version {fn_ver} != METADATA version {version}")
        out[kind] = {"path": path, "name": want[kind], "version": version, "filename": path.name,
                     "requires": requires}
    return out


def exact_pin(requires: list[str], dist: str, version: str) -> None:
    hits = []
    for raw in requires:
        try:
            req = Requirement(raw)
        except InvalidRequirement as e:
            fail(f"launcher Requires-Dist {raw!r} does not parse: {e}")
        if canonicalize_name(req.name) == dist:
            hits.append(req)
    if len(hits) != 1:
        fail(f"launcher must declare {dist} exactly once, found {[str(h) for h in hits]}")
    req = hits[0]
    if req.marker is not None or req.extras or req.url:
        fail(f"launcher requirement {req} must be unconditional, without extras or URL")
    specs = list(req.specifier)
    if len(specs) != 1 or specs[0].operator != "==" or specs[0].version != version:
        fail(f"launcher requirement {req} must be exactly {dist}=={version}")


def check_build(root: Path, product: str, tag: str, dist: Path, prefix: str = "v") -> dict:
    version = version_from_tag(tag, prefix)
    bad = {k: v for k, v in source_versions(root, product).items() if v != version}
    if bad:
        fail(f"source versions differ from the tag version {version!r}: {bad}")
    got = local_set(dist, product)
    if product == LICENSE_LANE:
        if got["license"]["version"] != version:
            fail(f"license wheel is {got['license']['version']}, the release is {version}")
        return {"license": {"filename": got["license"]["filename"], "version": version}}
    for kind in (("app",) if product == "xl" else ("core", "launcher")):
        if got[kind]["version"] != version:
            fail(f"{kind} wheel is {got[kind]['version']}, the release is {version}")
    if product == "xl":
        requires = got["app"]["requires"]
        if any(canonicalize_name(Requirement(raw).name) == PRODUCTS[product][0]
               for raw in requires):
            fail("Community app must not require a separate editing core")
        if got["license"]["version"] != "0.2.0":
            fail("Community app requires license notices version 0.2.0")
        exact_pin(requires, LICENSE_DIST, "0.2.0")
    else:
        exact_pin(got["launcher"]["requires"], PRODUCTS[product][0], version)
        exact_pin(got["launcher"]["requires"], LICENSE_DIST, got["license"]["version"])
    return {k: {"filename": v["filename"], "version": v["version"]} for k, v in got.items()}


def normalized_members(data: bytes) -> list[tuple[str, str]]:
    """Sorted (member path, sha256 of uncompressed bytes). ZIP container metadata ignored."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            fail("wheel has duplicate member names")
        return sorted((n, hashlib.sha256(z.read(n)).hexdigest()) for n in names)


def same_payload(local: bytes, remote: bytes) -> tuple[bool, str]:
    a, b = normalized_members(local), normalized_members(remote)
    if a == b:
        return True, ""
    an, bn = dict(a), dict(b)
    diff = sorted(n for n in set(an) | set(bn) if an.get(n) != bn.get(n))
    return False, f"members differ: {diff[:10]}"


# ------------------------------------------------------------------ PyPI
def http_get(url: str, timeout: float = 30.0) -> tuple[int, bytes]:
    req = urllib.request.Request(url, headers={"User-Agent": "ks4-release-verify/1"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:  # noqa: S310 (https only, checked)
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, b""


FETCH = http_get  # tests replace this


def fetch(url: str, tries: int = 3) -> tuple[int, bytes]:
    last = None
    for attempt in range(tries):
        try:
            status, body = FETCH(url)
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            last = e
        else:
            if status in (200, 404):
                return status, body
            last = f"HTTP {status}"
        time.sleep(min(2 ** attempt, 10) if FETCH is http_get else 0)
    fail(f"{url}: no definite answer after {tries} tries ({last})")


def remote_state(index: str, w: dict) -> str:
    """'absent' or 'verified'; raises on anything else (including a mismatch)."""
    url = f"{index}/pypi/{w['name']}/{w['version']}/json"
    status, body = fetch(url)
    if status == 404:
        return "absent"
    info = json.loads(body)
    files = [f for f in info.get("urls", []) if f.get("filename") == w["filename"]]
    if len(files) != 1:
        listed = [f.get("filename") for f in info.get("urls", [])]
        fail(f"{w['name']} {w['version']} exists on PyPI but {w['filename']} is not its file "
             f"(listed: {listed}); a new version is required")
    entry = files[0]
    durl = entry.get("url", "")
    host = urllib.parse.urlparse(durl).hostname
    allowed = FILES_HOST if index == DEFAULT_INDEX else urllib.parse.urlparse(index).hostname
    if urllib.parse.urlparse(durl).scheme not in ("https", "http") or host != allowed or \
            (index == DEFAULT_INDEX and not durl.startswith("https://")):
        fail(f"{w['filename']}: unexpected download URL {durl}")
    dstatus, data = fetch(durl)
    if dstatus != 200:
        fail(f"{w['filename']}: download returned HTTP {dstatus}")
    want = (entry.get("digests") or {}).get("sha256")
    if not want or hashlib.sha256(data).hexdigest() != want:
        fail(f"{w['filename']}: downloaded bytes do not match PyPI's advertised sha256")
    ok, why = same_payload(Path(w["path"]).read_bytes(), data)
    if not ok:
        fail(f"{w['name']} {w['version']} on PyPI is NOT the wheel this release built ({why}). "
             "PyPI files cannot be replaced: bump the version.")
    return "verified"


def plan(dist: Path, product: str, license_mode: str, index: str) -> dict[str, str]:
    if product == LICENSE_LANE and license_mode != "publish":
        fail("the license workflow is the adapter's publisher: use --license-mode publish")
    got = local_set(dist, product)
    decision = {}
    for kind in kinds_for(product):
        state = remote_state(index, got[kind])
        if kind == "license" and license_mode == "require":
            if state != "verified":
                fail(f"{LICENSE_DIST} {got['license']['version']} is not on PyPI. This repository "
                     "never uploads the shared adapter: its own license workflow publishes it; "
                     "release and verify it there first.")
            decision[kind] = "skip"
        else:
            decision[kind] = "upload" if state == "absent" else "skip"
        print(f"{kind:8} {got[kind]['filename']}: {state} -> {decision[kind]}")
    return decision


def verify_remote(dist: Path, product: str, index: str, wait: int, kinds=None) -> None:
    kinds = kinds or kinds_for(product)
    got = local_set(dist, product)
    deadline = time.monotonic() + wait
    for kind in kinds:
        while True:
            state = remote_state(index, got[kind])
            if state == "verified":
                print(f"{kind:8} {got[kind]['filename']}: on PyPI and identical")
                break
            if time.monotonic() >= deadline:
                fail(f"{got[kind]['filename']} is not visible on PyPI")
            time.sleep(15)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("check-build", "plan", "verify-remote"):
        p = sub.add_parser(name)
        p.add_argument("--product", required=True, choices=sorted([*PRODUCTS, LICENSE_LANE]))
        p.add_argument("--dist", required=True)
        p.add_argument("--index", default=DEFAULT_INDEX)
        if name == "check-build":
            p.add_argument("--tag", required=True)
            p.add_argument("--tag-prefix", default="v", choices=TAG_PREFIXES)
            p.add_argument("--root", default=".")
        if name == "plan":
            p.add_argument("--license-mode", required=True, choices=("publish", "require"))
            p.add_argument("--github-output", default=os.environ.get("GITHUB_OUTPUT"))
        if name == "verify-remote":
            p.add_argument("--wait", type=int, default=0)
            p.add_argument("--kinds", default="",
                           help="comma-separated subset of app,license,core,launcher "
                                "(default: every wheel of the product)")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "check-build":
            res = check_build(Path(a.root), a.product, a.tag, Path(a.dist), a.tag_prefix)
            print("release wheels, exact versions and pins:", json.dumps(res))
        elif a.cmd == "plan":
            d = plan(Path(a.dist), a.product, a.license_mode, a.index.rstrip("/"))
            if a.github_output:
                with open(a.github_output, "a", encoding="utf-8") as f:
                    for k, v in d.items():
                        f.write(f"{k}={v}\n")
        else:
            allowed = kinds_for(a.product)
            kinds = tuple(k for k in a.kinds.split(",") if k) or allowed
            if any(k not in allowed for k in kinds):
                raise ReleaseCheckError(f"--kinds {a.kinds!r}: expected a subset of {allowed}")
            verify_remote(Path(a.dist), a.product, a.index.rstrip("/"), a.wait, kinds)
    except ReleaseCheckError as e:
        print(f"RELEASE CHECK FAILED: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
