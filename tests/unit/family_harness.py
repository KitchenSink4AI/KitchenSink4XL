"""Test harness for the family surface: an in-process server built through
server.create_server (the fastmcp Client over the in-memory transport), and
a real stdio round trip through the Community launcher.

The same design as the Word editor's family-surface harness. Every call returns
(payload, is_error, raw_result). Nothing here reaches a user's Excel: the
COM ops these tests call refuse before starting Excel unless the test says
otherwise, and the COM tier only ever drives its own private hidden pool.
"""

from __future__ import annotations

import asyncio
import importlib.util
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _adapter_paths() -> list[str]:
    spec = importlib.util.spec_from_file_location(
        "_ks4xl_adapter_source_h", ROOT / "scripts" / "adapter_source.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.import_paths()


PKG_SRC = [str(ROOT / "packages" / "kitchensink4xl-community" / "src"), *_adapter_paths()]


class Harness:
    def __init__(self, mcp):
        self.mcp = mcp

    async def _with_client(self, fn):
        from fastmcp import Client
        async with Client(self.mcp) as client:
            return await fn(client)

    def list_tools(self):
        return asyncio.run(self._with_client(lambda c: c.list_tools()))

    def capabilities(self):
        async def caps(c):
            return c.initialize_result.capabilities
        return asyncio.run(self._with_client(caps))

    def call(self, tool: str, arguments):
        async def go(c):
            return await c.call_tool_mcp(tool, arguments)
        res = asyncio.run(self._with_client(go))
        text = res.content[0].text if res.content else ""
        try:
            payload = json.loads(text)
        except ValueError:
            payload = {"_raw": text}
        return payload, bool(res.isError), res

    def calls(self, seq):
        """Several calls in one client session (much faster)."""
        async def go(c):
            out = []
            for tool, arguments in seq:
                res = await c.call_tool_mcp(tool, arguments)
                text = res.content[0].text if res.content else ""
                try:
                    out.append((json.loads(text), bool(res.isError), res))
                except ValueError:
                    out.append(({"_raw": text}, bool(res.isError), res))
            return out
        return asyncio.run(self._with_client(go))


def start(settings=None) -> Harness:
    from xlsx_mcp import server
    return Harness(server.create_server(settings))


# ------------------------------------------------------------------ stdio

LAUNCH = ("import sys; sys.path[:0] = {paths!r}; "
          "from ks4xl_community_launch import main; main()")


def launch_cmd() -> list[str]:
    return [sys.executable, "-X", "utf8", "-c", LAUNCH.format(paths=PKG_SRC)]


def _env(extra: dict | None) -> dict:
    env = {k: v for k, v in os.environ.items()
           if not k.startswith("KS4XL_") or k in ("KS4XL_UPDATE_CHECK",)}
    env["KS4XL_UPDATE_CHECK"] = "off"
    env["PYTHONPATH"] = os.pathsep.join([str(ROOT / "src"), *PKG_SRC])
    env.update(extra or {})
    return env


def stdio_session(calls, env: dict | None = None, raw_first: list[dict] | None = None,
                  timeout: float = 120.0) -> dict:
    """Start the Community launcher over real stdio, initialize, list tools,
    make each call, list tools again. Raw JSON-RPC, no client library, so the
    wire is exactly what a client receives. raw_first: messages sent BEFORE
    initialize."""
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    proc = subprocess.Popen(launch_cmd(), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=_env(env), creationflags=flags)
    rid = 0
    out: dict = {"results": [], "raw_first": []}

    def send(msg):
        proc.stdin.write((json.dumps(msg) + "\n").encode("utf-8"))
        proc.stdin.flush()

    def recv(want_id):
        while True:
            line = proc.stdout.readline()
            if not line:
                raise RuntimeError("server closed stdout: "
                                   + proc.stderr.read().decode("utf-8", "replace")[-2000:])
            msg = json.loads(line)
            if msg.get("id") == want_id:
                return msg

    import threading
    killer = threading.Timer(timeout, proc.kill)
    killer.start()
    try:
        for m in raw_first or []:
            rid += 1
            send({**m, "jsonrpc": "2.0", "id": rid})
            out["raw_first"].append(recv(rid))
        rid += 1
        send({"jsonrpc": "2.0", "id": rid, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {},
            "clientInfo": {"name": "ks4xl-test", "version": "0"}}})
        init = recv(rid)["result"]
        out["capabilities"] = init["capabilities"]
        out["server_version"] = init["serverInfo"]["version"]
        out["instructions"] = init.get("instructions", "")
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        rid += 1
        send({"jsonrpc": "2.0", "id": rid, "method": "tools/list"})
        out["tools_raw"] = recv(rid)["result"]
        out["tools"] = out["tools_raw"]["tools"]
        for tool, arguments in calls:
            rid += 1
            send({"jsonrpc": "2.0", "id": rid, "method": "tools/call",
                  "params": {"name": tool, "arguments": arguments}})
            res = recv(rid)
            r = res.get("result") or {}
            content = r.get("content") or []
            out["results"].append({
                "is_error": bool(r.get("isError")) or "error" in res,
                "blocks": len(content),
                "structured": r.get("structuredContent"),
                "text": content[0]["text"] if content else json.dumps(res.get("error")),
            })
        rid += 1
        send({"jsonrpc": "2.0", "id": rid, "method": "tools/list"})
        out["tools_after_raw"] = recv(rid)["result"]
    finally:
        killer.cancel()
        try:
            proc.stdin.close()
        except OSError:
            pass
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=20)
        out["stderr"] = proc.stderr.read().decode("utf-8", "replace")
        out["returncode"] = proc.returncode
    return out


def stdio_start_only(env: dict | None = None, timeout: float = 60.0):
    """Start the launcher with stdin closed; returns (returncode, stderr)."""
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    res = subprocess.run(launch_cmd(), stdin=subprocess.DEVNULL, capture_output=True,
                         env=_env(env), timeout=timeout, creationflags=flags)
    return res.returncode, res.stderr.decode("utf-8", "replace")
