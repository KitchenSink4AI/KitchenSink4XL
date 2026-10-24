"""Validate installed edition wheels, MCP wire identity and bundle contents.

Community ships an app wheel containing its core and launcher plus a license
policy wheel. Business retains three separate wheels. Both editions expose
an identical tool surface, Business carries no Community guidance, console
and stdio entry points start and nothing reaches the network.
"""
from __future__ import annotations
import argparse
import asyncio
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import adapter_source  # noqa: E402

CORE_VERSION = __import__("tomllib").loads(
    (ROOT / "pyproject.toml").read_text("utf-8"))["project"]["version"]

#: A distribution must not ship a wheel granted under the other license.
FORBIDDEN_LICENSE = {"community": "LicenseRef-KitchenSink-Commercial",
                     "business": "AGPL-3.0-only"}

CHECK = r"""
import asyncio, json, os, sys
from pathlib import Path
sys.path.insert(0, sys.argv[1])
edition = sys.argv[2]
root = Path(sys.argv[1]).resolve()

# The launcher injects the policy; nothing else in this process chooses one.
if edition == "community":
    import ks4xl_community_launch as launcher
    build_policy = launcher.build_policy
else:
    import ks4xl_business_launch as launcher
    build_policy = launcher.build_policy
from xlsx_mcp import license_policy
license_policy.install(build_policy())

from xlsx_mcp import server
from fastmcp import Client
from openpyxl import Workbook, load_workbook

assert Path(server.__file__).resolve().is_relative_to(root), server.__file__
os.environ["KS4XL_UPDATE_CHECK"] = "off"
# Failure is loud if either installed edition attempts a network update.
import urllib.request
def blocked(*a, **k):
    raise AssertionError("Unexpected runtime network request")
urllib.request.urlopen = blocked
path = root / "smoke.xlsx"
wb = Workbook()
wb.active["A1"] = "original"
wb.save(path)
wb.close()

if edition == "community":
    from ks4_license_community.notices import WARNING_13
else:
    WARNING_13 = None

async def check():
    async with Client(server.create_server()) as client:
        init = client.initialize_result
        result = await client.call_tool("xl_read", {"op": "get_server_info"})
        info = json.loads(result.content[0].text)
        assert info["ok"] is True, info
        tools = await client.list_tools()
        changed = await client.call_tool("xl_edit", {"op": "set_cell", "args": {
            "path": str(path), "location": {"cell": "A1"}, "value": "changed",
            "verify_com": False}})
        assert not changed.is_error
        check = load_workbook(path)
        assert check.active["A1"].value == "changed"
        check.close()
        return {"name": init.serverInfo.name, "version": init.serverInfo.version,
                "core_version": info["core_version"],
                "edition": info["edition"], "license": info["licensing"]["license"],
                "status": info["licensing"]["status"],
                "warning_in_instructions": bool(WARNING_13) and WARNING_13 in init.instructions,
                "tools": [t.name for t in tools],
                "surface": [[t.name, t.inputSchema["properties"].get("op", {}).get("enum"),
                             t.annotations.model_dump(exclude_none=True)] for t in tools],
                # Community reports "state", the Business channel "status"
                "update_state": info["update_check"].get("state") or info["update_check"].get("status"),
                "excel_edit": "passed"}
async def served_strings():
    # Owner-test mode: every string the
    # server SERVES (instructions, tool text and schemas, xl_help for every op
    # and topic, refusal hints) carries no unfilled copy slot, no em dash and
    # no home folder.
    from xlsx_mcp import envelope
    texts = []
    async with Client(server.create_server()) as client:
        texts.append(client.initialize_result.instructions or "")
        for t in await client.list_tools():
            texts.append(json.dumps(t.model_dump(exclude_none=True), ensure_ascii=False))
        listing = json.loads((await client.call_tool("xl_help", {})).content[0].text)["tools"]
        for tool_name, ops in listing.items():
            for op in ops:
                texts.append((await client.call_tool("xl_help", {"op": op})).content[0].text)
        texts.append((await client.call_tool("xl_help", {"topic": "location"})).content[0].text)
    texts += list(envelope.HINTS.values())
    blob = json.loads(json.dumps(texts))
    home = str(Path.home())
    bad = []
    for text in blob:
        for marker in ("COPY" "_SLOT:", "COPY PENDING", "FACTS" " TO CONVEY", "PLACEHOLDER",
                       chr(0x2014)):
            if marker in text:
                bad.append(marker)
        if home.lower() in text.lower() or home.replace(chr(92), "/").lower() in text.lower():
            bad.append("home folder")
    return sorted(set(bad))
result = asyncio.run(check())
if len(sys.argv) > 3 and sys.argv[3] == "owner-test":
    result["served_string_defects"] = asyncio.run(served_strings())
print(json.dumps(result))
"""


def community_needles() -> tuple[str, ...]:
    """Read the approved wording out of the ONE canonical adapter source."""
    import importlib.util
    source = adapter_source.notices_source()
    spec = importlib.util.spec_from_file_location("_smoke_notices", source)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.guidance_needles()


def payload_digest(wheel: Path) -> str:
    """Content hash of a wheel's SOURCE members, ignoring zip timestamps.

    Compare only xlsx_mcp members: the Community app also carries its launcher,
    while Business keeps the core in a separate wheel. The code must remain
    identical despite different distribution metadata and ZIP timestamps.
    """
    with zipfile.ZipFile(wheel) as archive:
        parts = sorted((name, hashlib.sha256(archive.read(name)).hexdigest())
                       for name in archive.namelist()
                       if name.startswith("xlsx_mcp/") and not name.endswith("/"))
    assert parts, f"{wheel.name} has no payload outside dist-info"
    return hashlib.sha256(repr(parts).encode()).hexdigest()


def wheels_for(folder: Path) -> dict:
    found = {}
    for wheel in folder.glob("*.whl"):
        if wheel.name.startswith("kitchensink4xl_core-"):
            found["core"] = wheel
        elif wheel.name.startswith("kitchensink4_license_"):
            found["policy"] = wheel
        else:
            found["launcher"] = wheel
    expected = ({"policy", "launcher"} if folder.name == "community"
                else {"core", "policy", "launcher"})
    assert set(found) == expected, f"{folder}: expected {sorted(expected)}, got {sorted(found)}"
    assert len(list(folder.glob("*.whl"))) == len(expected), "duplicate edition wheel"
    return found



def assert_shared_surface(community: list, business: list) -> None:
    """Only Community offers public updates; all other surface fields agree."""
    assert len(community) == len(business)
    for public, commercial in zip(community, business):
        assert public[0] == commercial[0]
        if public[0] == "xl_app":
            updates = {"update_now", "update_notices"}
            assert updates <= set(public[1])
            assert not updates & set(commercial[1])
            assert public[2]["openWorldHint"] is True
            assert commercial[2]["openWorldHint"] is False
            shared = [public[0], [op for op in public[1] if op not in updates],
                      {**public[2], "openWorldHint": False}]
            assert shared == commercial, (shared, commercial)
        else:
            assert public == commercial, (public, commercial)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("artifacts", type=Path)
    parser.add_argument("--owner-test", action="store_true",
                        help="Also scan every served string for copy slots, em dashes "
                             "and home folders")
    args = parser.parse_args()
    needles = community_needles()
    reports = {}
    for name in ("community", "business"):
        folder = args.artifacts.resolve() / name
        wheels = wheels_for(folder)
        bundle, = folder.glob("*.mcpb")
        licenses = {}
        for role, wheel in wheels.items():
            with zipfile.ZipFile(wheel) as z:
                names = z.namelist()
                assert not any(n.endswith((".pem", ".xlsx", ".env")) or "/tests/" in n
                               for n in names)
                meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
                declared = ""
                for line in z.read(meta).decode("utf-8").splitlines():
                    if line.startswith("License-Expression:"):
                        declared = line.split(":", 1)[1].strip()
                        break
                # No conflicting grant anywhere in a shipped distribution.
                assert declared, f"{wheel.name} declares no license"
                assert FORBIDDEN_LICENSE[name] not in declared, (
                    f"{wheel.name} declares {declared} in a {name} distribution")
                licenses[role] = declared
                if name == "business":
                    for member in names:
                        if Path(member).suffix in {".png", ".pyc"}:
                            continue
                        text = z.read(member).decode("utf-8", errors="replace")
                        for needle in needles:
                            assert needle not in text, f"{wheel.name}::{member}"
        with zipfile.ZipFile(wheels["policy"]) as z:
            names = z.namelist()
            if name == "community":
                source = z.read("ks4_license_community/notices.py")
                assert bytes.fromhex("c2a7") + b"13" in source
                assert b"?13" not in source
            # The shared adapter wheel is product-neutral: no generated
            # metadata, no product name, identical for every editor.
            assert not [n for n in names if n.endswith("_edition.json")]
            meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
            if name == "business":
                assert "Requires-Dist: cryptography>=42" in z.read(meta).decode()
        with zipfile.ZipFile(wheels["launcher"]) as z:
            names = z.namelist()
            meta, = [n for n in names if n.endswith(".dist-info/METADATA")]
            metadata = z.read(meta).decode()
            expected_license = ("LicenseRef-KitchenSink-Commercial" if name == "business"
                                else "AGPL-3.0-only")
            assert "License-Expression: " + expected_license in metadata
            if name == "business":
                assert "Requires-Dist: kitchensink4xl-core==" + CORE_VERSION + "+commercial" in metadata
            else:
                assert "Requires-Dist: kitchensink4xl-core" not in metadata
                assert "Requires-Dist: kitchensink4-license-community==0.2.0" in metadata
                assert "xlsx_mcp/__init__.py" in names and "xlsx_mcp/server.py" in names
            ep, = [n for n in names if n.endswith("/entry_points.txt")]
            entry_points = z.read(ep).decode()
            if name == "business":
                assert "kitchensink4xl-business" in entry_points
                assert "ks4xl_business_launch/_edition.json" in names
            else:
                assert "xl-mcp" in entry_points and "kitchensink4xl" in entry_points
        module = f"ks4xl_{name}_launch"
        with zipfile.ZipFile(bundle) as z:
            guide = z.read("INSTALL.md").decode("utf-8")
            if name == "community":
                assert "§13" in guide
            else:
                for needle in needles:
                    assert needle not in guide
            manifest = json.loads(z.read("manifest.json"))
            assert manifest["display_name"].endswith(name.title())
            assert manifest["server"]["mcp_config"]["args"] == ["-m", module]
            assert manifest["server"]["mcp_config"]["command"] == "${user_config.python_path}"
            for wheel in wheels.values():
                assert z.read(wheel.name) == wheel.read_bytes()
            assert {"LICENSE", "INSTALL.md", "icon.png"} <= set(z.namelist())
        with tempfile.TemporaryDirectory(prefix="ks4xl-installed-smoke-") as scratch:
            subprocess.run([sys.executable, "-m", "pip", "install", "--no-deps",
                            "--no-compile", "--target", scratch,
                            *[str(w) for w in wheels.values()]],
                           check=True, capture_output=True, text=True,
                           stdin=subprocess.DEVNULL)
            child_env = {**os.environ, "KS4XL_UPDATE_CHECK": "off", "PYTHONDONTWRITEBYTECODE": "1"}
            child_env.pop("KS4XL_LICENSE_FILE", None)
            result = subprocess.run([sys.executable, "-I", "-c", CHECK, scratch, name,
                                     "owner-test" if args.owner_test else ""],
                                    check=True, capture_output=True, text=True,
                                    encoding="utf-8", env=child_env,
                                    stdin=subprocess.DEVNULL)
            info = json.loads(result.stdout)
            if args.owner_test:
                assert info["served_string_defects"] == [], (
                    f"{name}: served strings carry {info['served_string_defects']}")
            from fastmcp import Client
            from fastmcp.client.transports import StdioTransport

            # A retired 1.x setting that locked the tools stops the server
            # with one line, rather than being ignored. The built edition
            # must do that.
            legacy = subprocess.run(
                [sys.executable, "-m", module], cwd=scratch,
                env={**os.environ, "PYTHONPATH": scratch,
                     "PYTHONDONTWRITEBYTECODE": "1", "KS4XL_UPDATE_CHECK": "off",
                     "KS4XL_PACK_POLICY": "locked"},
                stdin=subprocess.DEVNULL, capture_output=True, text=True,
                timeout=120, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if legacy.returncode != 2 or "limited the tools in an earlier version" not in legacy.stderr:
                raise SystemExit(f"{name}: KS4XL_PACK_POLICY=locked did not stop the server "
                                 f"(rc {legacy.returncode}): {legacy.stderr[-300:]}")

            async def stdio_check():
                # Over real stdio the list is the fixed surface (names, order,
                # op enums, annotations); serverInfo.version is the build's;
                # listChanged is not advertised; a forbidden op refuses
                # inside the call.
                transport = StdioTransport(
                    command=sys.executable,
                    args=["-m", module],
                    env={"PYTHONPATH": scratch, "PYTHONDONTWRITEBYTECODE": "1",
                         "KS4XL_UPDATE_CHECK": "off", "KS4XL_LICENSE_FILE": "",
                         "KS4XL_FORBID": "audit_styles",
                         "KS4XL_VERIFY_COM": "false"},
                    cwd=scratch, keep_alive=False)
                async with Client(transport) as client:
                    init = client.initialize_result
                    assert init.serverInfo.name == info["name"]
                    assert init.serverInfo.version == info["version"]
                    assert not (init.capabilities.tools and init.capabilities.tools.listChanged)
                    response = await client.call_tool("xl_read", {"op": "get_server_info"})
                    assert json.loads(response.content[0].text)["edition"] == name
                    refused = await client.call_tool_mcp("xl_read", {
                        "op": "audit_styles", "args": {"path": "x.xlsx"}})
                    body = json.loads(refused.content[0].text)
                    assert refused.isError and body["error"]["code"] == "NOT_AVAILABLE", body
                    wire = await client.list_tools()
                    return [[t.name, t.inputSchema["properties"].get("op", {}).get("enum"),
                             t.annotations.model_dump(exclude_none=True)] for t in wire]

            stdio_surface = asyncio.run(asyncio.wait_for(stdio_check(), timeout=60))
            # The forbidden op is the only difference from the in-process list.
            expected = [[n, [o for o in ops if o != "audit_styles"] if ops else ops, ann]
                        for n, ops, ann in info["surface"]]
            assert stdio_surface == expected, (stdio_surface, expected)
            info["stdio_startup_tools"] = len(stdio_surface)
            assert info["edition"] == name
            assert info["warning_in_instructions"] == (name == "community")
            assert ("+business" in info["version"]) == (name == "business")
            info["wheel_sha256"] = {role: hashlib.sha256(w.read_bytes()).hexdigest()
                                    for role, w in wheels.items()}
            core_wheel = wheels["core"] if name == "business" else wheels["launcher"]
            info["core_payload_digest"] = payload_digest(core_wheel)
            info["licenses"] = licenses
            info["core_wheel"] = core_wheel.name
            reports[name] = info
    assert reports["community"]["tools"] == reports["business"]["tools"]
    assert_shared_surface(reports["community"]["surface"], reports["business"]["surface"])
    # One shared core under both editions: same members, same content.
    assert (reports["community"]["core_payload_digest"]
            == reports["business"]["core_payload_digest"])
    for info in reports.values():
        info["tool_count"] = len(info.pop("tools"))
        info["op_count"] = sum(len(ops or []) for _n, ops, _a in info.pop("surface"))
    print(json.dumps(reports, indent=2))


if __name__ == "__main__":
    main()
