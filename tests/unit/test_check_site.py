"""The product page's copy guard runs with the suite (scripts/
check_site.py): docs/index.html is a redirect stub, and the guard checks that
its canonical link, meta refresh and script redirect all name the community
site, that it is noindex, and that no em dash reaches a reader."""

from __future__ import annotations

import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
TARGET = "https://community.kitchensink4.ai/xl/"


def _check_site():
    spec = importlib.util.spec_from_file_location("check_site", ROOT / "scripts" / "check_site.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_the_product_page_passes_its_copy_guard(capsys):
    mod = _check_site()
    assert mod.main() == 0, capsys.readouterr().out


def test_the_guard_catches_a_wrong_redirect_target_and_a_missing_noindex(tmp_path, capsys):
    mod = _check_site()
    html = (ROOT / "docs" / "index.html").read_text(encoding="utf-8")
    assert html.count(TARGET) == 4
    broken = html.replace(f'content="0; url={TARGET}"', 'content="0; url=https://example.invalid/"', 1)
    broken = broken.replace('<meta name="robots" content="noindex">', "", 1)
    (tmp_path / "index.html").write_text(broken, encoding="utf-8")
    assert mod.main(tmp_path) == 1
    out = capsys.readouterr().out
    assert "meta refresh target" in out and "noindex" in out
