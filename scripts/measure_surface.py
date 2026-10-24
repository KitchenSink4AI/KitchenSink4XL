"""Measure the real tool surface over real stdio.

Starts the Community launcher as a subprocess, reads tools/list exactly as a
client receives it, and prints per tool: description characters, input
schema bytes (compact JSON), and tokens by Excel's own counter (each tool as
serialized, compact JSON, length / 4: the method the released README
figures used, packs.approx_tokens before the family surface). The total is
the figure the ceiling in tests/unit/test_family_stdio.py pins.

    python -X utf8 scripts/measure_surface.py           # table
    python -X utf8 scripts/measure_surface.py --json    # one JSON line

"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests" / "unit"))


def approx_tokens(tool: dict) -> int:
    return round(len(json.dumps(tool, ensure_ascii=False, separators=(",", ":"))) / 4)


def measure(env: dict | None = None) -> dict:
    import family_harness as fh

    res = fh.stdio_session([], env=env)
    rows = []
    for t in res["tools"]:
        rows.append({
            "tool": t["name"],
            "description_chars": len(t.get("description") or ""),
            "schema_bytes": len(json.dumps(t.get("inputSchema") or {}, ensure_ascii=False,
                                           separators=(",", ":")).encode("utf-8")),
            "approx_tokens": approx_tokens(t),
            "ops": len(((t.get("inputSchema") or {}).get("properties") or {})
                       .get("op", {}).get("enum") or []),
        })
    return {"tools": len(rows), "ops": sum(r["ops"] for r in rows),
            "approx_tokens": sum(r["approx_tokens"] for r in rows),
            "instructions_chars": len(res.get("instructions") or ""),
            "server_version": res.get("server_version"), "rows": rows}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    report = measure()
    if a.json:
        print(json.dumps(report))
        return
    print(f"{'tool':<18} {'ops':>4} {'desc':>6} {'schema':>7} {'~tokens':>8}")
    for r in report["rows"]:
        print(f"{r['tool']:<18} {r['ops']:>4} {r['description_chars']:>6} "
              f"{r['schema_bytes']:>7} {r['approx_tokens']:>8}")
    print(f"TOTAL: {report['tools']} tools, {report['ops']} ops, "
          f"{report['approx_tokens']} tokens (Excel's counter); instructions "
          f"{report['instructions_chars']} chars")


if __name__ == "__main__":
    main()
