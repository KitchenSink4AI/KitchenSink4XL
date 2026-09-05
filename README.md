<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4xl -->

# 🔌 KitchenSink4XL

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4XL/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4XL/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4xl)](https://pypi.org/project/kitchensink4xl/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/LICENSE)

**Read and edit Excel workbooks with your AI assistant, and know whether each formula result is fresh, cached or missing.**

KitchenSink4XL is an MCP server that lets Claude Code, the Codex CLI, Copilot CLI, Claude Desktop and other local MCP clients work with .xlsx and .xlsm files: formulas, tables, formatting, structural edits and data checks. When it reads a formula cell, it tells your AI whether the value was just calculated, remembered from the last save, or never calculated at all, so a stale number does not end up in your report. This is the free Community edition, licensed under the AGPL-3.0.

[Product page](https://kitchensink4.ai/KitchenSink4XL/) · [llms.txt](https://kitchensink4.ai/KitchenSink4XL/llms.txt) · [Changelog](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/CHANGELOG.md)

**Platform:** the file tools run on Windows, macOS and Linux. Recalculation, pivot tables, PDF export and the other Excel-powered features need Windows with Microsoft Excel installed.

## Install

The package needs Python 3.12 or newer. Run the commands below in PowerShell or a terminal, not in an AI chat.

**Claude Code or Codex CLI.** Install [uv](https://docs.astral.sh/uv/getting-started/installation/), run the line for your app, and start a new session:

```sh
claude mcp add xl -s user -- uvx kitchensink4xl
```

```sh
codex mcp add xl -- uvx kitchensink4xl
```

**Claude Desktop.** Install uv and restart Claude Desktop. Download the `.mcpb` file from the [latest release](https://github.com/KitchenSink4AI/KitchenSink4XL/releases/latest), then open Settings > Extensions > Advanced settings > Install extension and choose the file. The first launch downloads the Python package, so it needs a network connection.

**Any other local MCP client.** Use `uvx` as the command and `kitchensink4xl` as its argument, or install the package and run `kitchensink4xl` as the server command:

```sh
pip install kitchensink4xl
```

The first launch through uvx builds the environment and can take 20 to 30 seconds; later launches start in about two. If your client times out on the first run, start it once from a terminal and try again.

## What it can do

- Tell whether a formula's value was calculated, cached or missing before you use it, with the affected cells named.
- Read a range or summarize a large table without pulling every row into your AI's context.
- Write formulas, values and number formats into an existing workbook.
- Insert and delete rows and columns while formulas, tables, named ranges, charts and cross-sheet references follow.
- Sort and filter the way Excel does, and format tables for a client or a colleague.
- Import and export CSV and JSON.
- Check a workbook for structural problems before you share it.
- On Windows with Excel: recalculate for real, build pivot tables, run goal seek, export to PDF and render a sheet to an image.

The server offers 67 workbook tools covering 129 operations, plus two tools that switch capability packs on and off.

## Why the labels matter

A workbook stores the last answer each formula produced, not a guarantee that the answer is current. If something upstream changed and nobody recalculated, the old number still looks like a real one. KitchenSink4XL labels every formula result it returns: freshly calculated, remembered from the last save, or not calculated yet. Values someone typed in carry no label, because there is nothing to recalculate. When it matters, the COM pack asks a private, hidden Excel to recalculate and returns results marked as calculated.

## Structural edits that keep the workbook intact

Inserting a column affects everything to its right: formulas, totals, charts, tables and named ranges on other sheets. KitchenSink4XL moves what should move and is tested against Excel's own behavior for formulas, tables, named ranges, charts and cross-sheet references. When an edit cannot be done safely, it refuses and says why instead of guessing.

## Tool packs and context cost

A session starts with a core set of tools and loads larger capability packs only when a task needs them. All figures come from `scripts/measure_surface.py` and count the full tool entry the client receives.

| Pack | Tools | Approx. tokens | What it covers |
|---|---:|---:|---|
| **lite** (loaded at start) | 40 | ~15.5k | A grid view of the workbook, labeled reads, server-side query and aggregation, formula writing and auditing, row and column edits that carry their references, sort and filter, tables, formatting, import and export, backups, diagnostics, pack controls |
| design | 9 | ~4.3k | Named cell styles, format painter, style cleanup, conditional formatting, data validation, images, charts, table lifecycle, named ranges including LAMBDA definitions |
| io | 9 | ~3.0k | Page layout and printing, headers and footers, protection, legacy comments, multi-sheet export, and inspectors for external links, VBA, pivot tables and data connections |
| com | 11 | ~3.1k | A private hidden Excel, never your open session: recalculation, pivot tables, goal seek, PDF export, sheet rendering, format conversion, encryption, sparklines, autofit, an opens-clean check |
| **Everything** | **69** | **~25.8k** | All packs plus `enable_tools` and `disable_tools` |

Your AI turns a pack on by name:

```
enable_tools(["design"])
enable_tools(["com"])
```

A request outside the lite core is refused with the exact pack and call that unlocks it. To load everything at start, set `KS4XL_MODE=full`, or give a comma-separated pack list. Administrators can fix the selection with `KS4XL_PACK_POLICY=locked`. Clients that load tool schemas on first use pay close to nothing until a tool is called.

## Safety model

- Before every change, the workbook is backed up into two rotating slots in a hidden `.ks4xl-backups/` folder next to it. `manage_backups` lists, restores and prunes them. The slots capture changes made through this server only; an edit made in Excel or another program is not captured.
- Saves are atomic and validated, and each write is checked after saving, with the backup restored if the check fails. With `verify_com=true` (or `KS4XL_VERIFY_COM=1` for every save), the saved file must also open cleanly in a hidden Excel.
- Before a save, a scan names any part of the workbook the writer cannot preserve, and refuses the loss unless the call explicitly allows that specific kind.
- Text that merely looks like a formula is never turned into one. The cell's type decides.
- Ambiguous targets are refused with every candidate listed.
- The COM pack uses its own hidden Excel instance, never the workbook you have open, and leaves no stray Excel processes behind. A workbook open in Excel can be read but not changed.
- `query_range`, `export_range` and the aggregates include rows hidden by an autofilter. Excel's SUBTOTAL skips them, so totals on a filtered sheet can differ from Excel's status bar.
- Exclude `.ks4xl-backups/` from cloud sync tools; the backup files change on every edit and sync clients can hold locks that slow saves down.

## Privacy

The server runs on your computer. KitchenSink4AI never receives your files or usage data. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. The optional update check contacts PyPI and sends nothing beyond the request itself. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support.

## Configuration

| Variable | What it does |
|---|---|
| `KS4XL_MODE` | Tools at startup: `lite` (default), `full`, or a comma-separated pack list |
| `KS4XL_PACK_POLICY` | `auto` (default) or `locked`, which fixes the tool set at startup |
| `KS4XL_ALLOWED_ROOTS` | Restricts every path the server reads or writes to these folders (separated by `;` on Windows, `:` elsewhere) |
| `KS4XL_VERIFY_COM` | `1` makes the hidden-Excel open check the default for every save |
| `KS4XL_VALIDATE_COM` | `1` uses Excel's own verdict for `validate`'s structure check |
| `KS4XL_COM_TIMEOUT` | Maximum time for a COM call |
| `KS4XL_UPDATE_CHECK` | `off` disables the update check completely (`KS4XL_NO_UPDATE_CHECK=1` also works) |

**Sandboxing.** With `KS4XL_ALLOWED_ROOTS` set, the containment check runs on canonical paths, so `..\` traversal, symbolic links and junctions, short names, case tricks and lookalike folder names are all caught. A blocked call returns a typed error naming the path and the allowed roots before any file is opened.

**Update check.** The server checks PyPI for a newer release only when you call `get_server_info`: never at startup or on a timer, at most once every seven days, with a two-second limit. It never downloads or installs anything.

**Claude Desktop extension settings.** **Verify every save with Excel** turns on the hidden-Excel open check (off by default, since it adds a round trip through Excel on every write). **Limit the server to one folder** confines every path the server touches, reads included, to a folder you choose.

<details>
<summary><b>Running from source or with other clients</b></summary>

Point any MCP client at the installed executable:

```json
{"mcpServers": {"xl": {"command": "kitchensink4xl", "env": {"KS4XL_MODE": "lite"}}}}
```

`xl-mcp` is an equivalent entry point, and `python -m xlsx_mcp` starts the same server. `python -m xlsx_mcp.server` also works and is the route for releases before 1.1. The COM pack is an optional extra (`pip install kitchensink4xl[com]`) and does nothing outside Windows.

From a clone on Windows:

```
git clone https://github.com/KitchenSink4AI/KitchenSink4XL
cd KitchenSink4XL
python -m venv .venv
.venv\Scripts\pip install -e ".[com]"
claude mcp add xl -s user -- <absolute-path>\.venv\Scripts\xl-mcp.exe
```

On macOS and Linux, use `python3 -m venv .venv`, `.venv/bin/pip install -e ".[com]"` and the path `<absolute-path>/.venv/bin/xl-mcp`.

</details>

## Known limits

- **`com_render_sheet` needs a signed-in desktop session.** Rendering a sheet to an image goes through Excel's clipboard, which a hidden Excel in a non-interactive session cannot use. Every other COM tool runs unattended.
- **A desktop tool, not a shared service.** The server runs under your account on your computer, and the COM pack drives an Excel in your Windows session. It is not designed to be shared by a team from one installation.
- **Deleting or renaming a sheet does not rewrite references.** That matches Excel: formulas and names that point at a deleted sheet show `#REF!` when Excel opens the file. Audit references first if you are unsure.
- **File edits cannot recalculate.** After a write through the file tools, cells downstream keep their old cached values until Excel recalculates. The labels tell you which values are affected, and `recalculate` in the COM pack fixes them.
- **Re-enabled tools in some clients.** Some MCP clients drop a tool's schema when its pack is turned off and do not reload it when the pack comes back. If a tool reports "no such tool", refresh the client's tool list.

## Testing

The unit suite has about 1,300 tests and runs on every pull request. Structural edits are checked against Excel's own recalculation, formulas across every documented function family are written and reopened without repair prompts, and the fixture workbooks include features that only Excel can author, such as pivot tables, slicers and shapes. A separate test set that drives a real Excel runs on Windows.

## Reporting problems

Please [open an issue](https://github.com/KitchenSink4AI/KitchenSink4XL/issues/new?template=bug_report.yml) when something fails, is refused unexpectedly or returns a wrong number. Rebuild the structure with placeholder data rather than attaching a private workbook, and include the output of `get_server_info`, which is designed to be safe to share. Feature ideas are welcome [here](https://github.com/KitchenSink4AI/KitchenSink4XL/issues/new?template=feature_request.yml). For security issues, follow [SECURITY.md](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/SECURITY.md).

## Business edition

The Business edition pairs these tools with a Windows installer, a signed update channel and support under the Business terms. Update checks tell you when a covered release is available, and nothing installs on its own. See [pricing](https://kitchensink4.ai/pricing/). The Community edition stays free under the AGPL, including business use that meets its terms.

## Contributing

Bug reports, fixes, tests and documentation are welcome. See [CONTRIBUTING.md](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/CONTRIBUTING.md); first-time contributors agree to the [Contributor License Agreement](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/CLA.md).

## License

KitchenSink4XL is dual-licensed.

- **AGPL-3.0.** Free for anyone, including businesses, for any use that complies with the AGPL. Its terms include sharing your source code, with your modifications, when you distribute the software or make it available over a network.
- **Commercial license.** For organizations that want to build KitchenSink4XL into their own products or services without the AGPL's source-sharing obligations. Contact licensing@kitchensink4.ai.

Provided as is, without warranty of any kind, under the terms of the license.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line of Alvut Consulting, LLC. KitchenSink4AI tools are built with the help of AI coding assistants, and every release goes through the full test suite before it ships.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and Excel are trademarks of the Microsoft group of companies.
