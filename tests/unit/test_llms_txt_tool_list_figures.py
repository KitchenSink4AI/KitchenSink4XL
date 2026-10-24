"""docs/llms.txt states the tool list's exact size, overall and per tool, as
scripts/measure_surface.py measures it over real stdio. The figures must be
the measured ones. Runs the server over stdio; no Excel."""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path


def test_llms_txt_figures_are_the_measured_ones(pytestconfig):
    root = Path(pytestconfig.rootpath)
    r = subprocess.run([sys.executable, "-X", "utf8", "scripts/measure_surface.py", "--json"], cwd=root,
                       capture_output=True, text=True, timeout=300, stdin=subprocess.DEVNULL)
    m = json.loads(r.stdout[r.stdout.index("{"):])
    text = (root / "docs" / "llms.txt").read_text(encoding="utf-8")
    total = re.search(r"about ([\d,]+) approximate tokens", text).group(1).replace(",", "")
    per = {t: int(n.replace(",", "")) for t, n in re.findall(r"`(xl_\w+)` \d+ / ([\d,]+)", text)}
    measured = {row["tool"]: row["approx_tokens"] for row in m["rows"]}
    assert (int(total), per) == (m["approx_tokens"], measured)
