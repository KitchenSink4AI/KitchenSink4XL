"""The public text says where a note's signing name comes from in the words the
code supports: the Office user name, else the Windows sign-in name.

It does not say that Excel signs that way, which the server does not check. It
reads the Office user name from the registry and the Windows sign-in name from
the environment, and says so.
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
PUBLIC_FILES = ("README.md", "CHANGELOG.md", "docs/llms.txt")
#: Claims about Excel's own behaviour that this server cannot show for signing.
UNPROVEN = ("as Excel does", "your own Excel uses", "name Excel signs with", "Excel signs")


def _signing_lines(name: str) -> list[str]:
    text = (ROOT / name).read_text("utf-8")
    return [ln for ln in text.splitlines()
            if re.search(r"\bsign(s|ed|ing)?\b", ln, re.I)
            and re.search(r"\bnotes?\b|KS4XL_AUTHOR", ln)]


@pytest.mark.parametrize("name", PUBLIC_FILES)
def test_no_public_text_says_notes_are_signed_as_excel_does(name):
    lines = _signing_lines(name)
    assert lines, f"{name} no longer describes how a note is signed"
    for line in lines:
        for claim in UNPROVEN:
            assert claim not in line, f"{name}: {line.strip()[:160]}"


@pytest.mark.parametrize("name,marker", [("README.md", "`KS4XL_AUTHOR`"),
                                         ("CHANGELOG.md", "Notes you add are signed")])
def test_the_text_names_the_office_user_name_and_the_windows_sign_in_name(name, marker):
    line = next(ln for ln in (ROOT / name).read_text("utf-8").splitlines() if marker in ln)
    assert "Office user name" in line and "Windows sign-in name" in line, line[:200]
