# Dependency license ledger

Every declared dependency, its license, and why it is here. Enforced by
`tests/unit/test_dependency_ledger.py`, which fails the build if
`pyproject.toml` grows a dependency that is not listed here. Ported from
KitchenSink4Web, which carried the only ledger in the family until the
2026-09-15 license audit (finding D-02).

Licenses below were read from the installed package metadata in this repo's
virtual environment (`importlib.metadata`), not from a search result.

## The rule

**No copyleft and no source-available dependency anywhere in the REQUIRED
install.** Anything questionable lives behind an optional extra, so a
dependency's license never becomes the server's problem. The ship license is
`AGPL-3.0-only` and every permissive license below is one-way compatible into
it, which is the direction that matters: the obligations run to whoever
redistributes those packages, and `pip` resolves each one from its own
publisher with its own license files. KitchenSink4XL redistributes none of
them, so no attribution obligation attaches to the wheel.

## Required install

The package name is the FIRST column and the license the SECOND, in every
table in this file. The enforcing test reads them positionally.

| Package | License | Direction | Note |
|---|---|---|---|
| `fastmcp` | Apache-2.0 | permissive, one-way into anything | The MCP server framework. Pinned `>=3.4,<4` because a minor bump moved the visibility API mid-build. |
| `openpyxl` | MIT | permissive | The .xlsx object model the file tier is built on, reached in 26 source files. |
| `lxml` | BSD-3-Clause | permissive | The XML engine underneath openpyxl, and reached directly where the object model is not enough. The BSD-3 no-endorsement clause is the only obligation and it binds redistribution, which does not happen here. |
| `regex` | Apache-2.0 AND CNRI-Python | both permissive | Backs the caller-pattern guard in `core/_regex.py`. Needed rather than convenient: stdlib `re` has no match timeout, the server is single-threaded stdio, and one pathological caller pattern would deny service to the whole session. |
| `packaging` | Apache-2.0 OR BSD-2-Clause | either, permissive | Version comparison in `core/update_check.py`. |

No copyleft. No obligation triggered by the current distribution model.

## Optional extras

| Package | License | Extra | Why it is optional |
|---|---|---|---|
| `pywin32` | PSF | `com` | The COM tier, which drives a private hidden Excel for pivot tables, recalculation, PDF export and encryption. Optional because the file tier is the whole product on a machine with no Excel, and declared under a win32 platform marker inside the extra so it is never installed where it cannot work. |
| `pytest` | MIT | `dev` | Test-time only, never distributed. |
| `pytest-timeout` | MIT | `dev` | Test-time only. A hung COM call is this family's most common CI failure and an unbounded one wedges the runner. |
| `pillow` | MIT-CMU | `dev` | Image handling in `ops/objects.py`, exercised by the test suite. MIT-CMU is the historical-permission-notice variant Pillow ships; permissive, attribution on redistribution only. |
| `psutil` | BSD-3-Clause | `dev` | Process accounting in `com/instances.py`, which is how the COM tier proves it leaves no orphaned Excel behind. |

## Transitive obligations

`et-xmlfile` (MIT) arrives through openpyxl and is the one transitive package
worth naming, because it writes the XML openpyxl streams. It is resolved by
`pip` from its own publisher and is not redistributed here. A full transitive
audit has not been run; it is the right companion to the one the
KitchenSink4Web ledger also defers.

## Fonts, which are the one real attribution obligation

`docs/fonts/` ships Fraunces and IBM Plex Mono as `.woff2` files on the
documentation page. Both are under the SIL Open Font License 1.1 (OFL),
which requires the copyright notice and license to travel with the font
software. `docs/fonts/OFL-Fraunces.txt` and `docs/fonts/OFL-IBMPlexMono.txt`
are those notices. Unlike every package
above, these files ARE redistributed, which is why the obligation is real
here and nowhere else in this ledger.
