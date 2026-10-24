"""The install card and the package docstring name only entry points that start
the server (2.0.0rc5).

In 2.0 the shared core, `xlsx_mcp`, carries no license policy: `python -m
xlsx_mcp` (and `python -m xlsx_mcp.server`) stop with a message that names the
launcher. The server starts from the `kitchensink4xl` or `xl-mcp` command, or
from `python -m ks4xl_community_launch`. The product page's install card still
told readers to point a client at the `xl-mcp` executable "in the clone's
virtual environment", and the package's `__main__` docstring said `python -m
xlsx_mcp` starts the server. The product page is now a redirect stub, so the
install wording is pinned in README.md and docs/llms.txt.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
INSTALL_TEXT = (ROOT / "README.md", ROOT / "docs" / "llms.txt")
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

#: Where the public copy says what `python -m xlsx_mcp` does.
_CORE_STOPS = "stop with a message that names the launcher"


def test_the_core_does_not_start_the_server_and_says_which_launcher_does():
    env = dict(os.environ, PYTHONPATH=str(ROOT / "src"), PYTHONUTF8="1")
    out = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "xlsx_mcp"], cwd=ROOT, env=env,
        stdin=subprocess.DEVNULL, capture_output=True, text=True, timeout=120,
        creationflags=NO_WINDOW)
    assert out.returncode != 0
    text = out.stdout + out.stderr
    assert "No license policy is installed" in text
    assert "ks4xl_community_launch" in text


def test_the_install_text_names_the_module_that_starts_the_server():
    for path in INSTALL_TEXT:
        text = path.read_text(encoding="utf-8")
        assert "python -m ks4xl_community_launch" in text, path.name
        assert "xl-mcp" in text, path.name
        # The core is named only to say it does not start the server.
        for m in re.finditer(r"python -m xlsx_mcp", text):
            assert _CORE_STOPS in text[m.start():m.start() + 200], (
                path.name, text[m.start():m.start() + 80])


def test_the_package_docstring_does_not_promise_that_the_core_starts_the_server():
    from xlsx_mcp import __main__ as pkg_main

    doc = pkg_main.__doc__ or ""
    assert "does not start the server" in doc
    assert "starts the same server" not in doc
