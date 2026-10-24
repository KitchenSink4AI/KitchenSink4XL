"""Copy guard for docs/index.html and docs/llms.txt.

docs/index.html is a redirect stub: the product page now lives at
https://community.kitchensink4.ai/xl/. This guard checks that the stub
sends every reader there (canonical link, meta refresh and the script
redirect all name the same address), that the page stays out of search
indexes, and that no em dash reaches a reader in the stub or in llms.txt.

Run:  python -X utf8 scripts/check_site.py
"""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
EM_DASH = chr(0x2014)  # built at run time: the copy guard forbids the literal
TARGET = "https://community.kitchensink4.ai/xl/"

failures: list[str] = []


def fail(msg: str) -> None:
    failures.append(msg)


def main(docs: Path = DOCS) -> int:
    failures.clear()
    html = (docs / "index.html").read_text(encoding="utf-8")

    canonical = re.findall(r'<link rel="canonical" href="([^"]*)"', html)
    if canonical != [TARGET]:
        fail(f"canonical link is {canonical}, expected [{TARGET!r}]")
    refresh = re.findall(
        r'<meta http-equiv="refresh" content="0; url=([^"]*)"', html)
    if refresh != [TARGET]:
        fail(f"meta refresh target is {refresh}, expected [{TARGET!r}]")
    script = re.findall(r'location\.replace\("([^"]*)" \+ location\.hash\)', html)
    if script != [TARGET]:
        fail(f"location.replace target is {script}, expected [{TARGET!r}]")
    if '<meta name="robots" content="noindex">' not in html:
        fail("the stub is not marked noindex")
    if EM_DASH in html:
        fail("em dash in the stub")

    llms = docs / "llms.txt"
    if llms.exists() and EM_DASH in llms.read_text(encoding="utf-8"):
        fail("em dash in llms.txt")
    return report()


def report() -> int:
    if failures:
        print(f"FAIL ({len(failures)}):")
        for f in failures:
            print("  -", f)
        return 1
    print("site copy guard: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
