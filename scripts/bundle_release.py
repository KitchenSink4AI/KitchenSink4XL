"""Name and hash the Claude Desktop bundle for a GitHub Release.

    mcpb pack bundle bundle/kitchensink4xl.mcpb
    python -X utf8 scripts/bundle_release.py bundle/kitchensink4xl.mcpb

Runs the same in-memory check the server runs before it opens a downloaded
bundle (core/update_flow.check_bundle), then writes, beside the input:

    kitchensink4xl-<version>.mcpb          the same bytes, versioned name
    kitchensink4xl-<version>.mcpb.sha256   "<sha256>  kitchensink4xl-<version>.mcpb"

The release attaches the versioned bundle and its .sha256: those are what
update_now downloads, and the site links the versioned name. Exit 0 on
success.
"""

from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


def release_files(packed: Path) -> tuple[Path, Path]:
    from xlsx_mcp.core import update_flow

    data = packed.read_bytes()
    with zipfile.ZipFile(packed) as archive:
        version = json.loads(archive.read("manifest.json").decode("utf-8")).get("version")
    if not isinstance(version, str):
        raise SystemExit("the bundle manifest has no version")
    update_flow.check_bundle(data, version)
    name = f"{update_flow.PACKAGE}-{version}.mcpb"
    bundle = packed.with_name(name)
    if bundle.resolve() != packed.resolve():
        bundle.write_bytes(data)
    digest = packed.with_name(name + ".sha256")
    digest.write_text(f"{hashlib.sha256(data).hexdigest()}  {name}\n", encoding="ascii",
                      newline="\n")
    return bundle, digest


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    try:
        bundle, digest = release_files(Path(sys.argv[1]))
    except ValueError as exc:
        print(f"refused: {exc}", file=sys.stderr)
        return 1
    print(bundle)
    print(digest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
