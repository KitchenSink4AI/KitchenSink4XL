<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4xl -->

# <img src="https://raw.githubusercontent.com/KitchenSink4AI/KitchenSink4XL/main/bundle/icon.png" alt="" width="40" height="40"> KitchenSink4XL

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4XL/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4XL/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4xl)](https://pypi.org/project/kitchensink4xl/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/LICENSE)

**Read and edit Excel workbooks with your AI assistant, and know whether each formula result is fresh, cached or missing.**

Your AI gets 131 actions in Excel, a backup before every change, and a check after every save.

KitchenSink4XL is an MCP server that lets Claude Code, the Codex CLI, Copilot CLI, Claude Desktop and other local MCP clients work with .xlsx and .xlsm files: formulas, tables, formatting, structural edits and data checks. When it reads a formula cell, it tells your AI whether the value was just calculated, remembered from the last save, or never calculated at all, so a stale number does not end up in your report. This is the free Community edition, licensed under the AGPL-3.0.

[Product page](https://community.kitchensink4.ai/xl/) · [llms.txt](https://community.kitchensink4.ai/xl/llms.txt) · [Changelog](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/CHANGELOG.md)

**Platform:** the file tools run on Windows and Linux (macOS is untested in this release). Recalculation, pivot tables, PDF export and the other Excel-powered features need Windows with Microsoft Excel installed.

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

In total the server offers 69 operations in 7 tools, 131 counting each action an operation can take (for example each kind of worksheet or table change). The whole tool list costs about 5.5k tokens of your AI's working memory at startup. Figures come from `scripts/measure_surface.py` and `scripts/count_operations.py`.

## Why the labels matter

A workbook stores the last answer each formula produced, not a guarantee that the answer is current. If something upstream changed and nobody recalculated, the old number still looks like a real one. KitchenSink4XL labels every formula result it returns: freshly calculated, remembered from the last save, or not calculated yet. Values someone typed in carry no label, because there is nothing to recalculate. When it matters, the `xl_app` tool asks a private, hidden Excel to recalculate and returns results marked as calculated.

## Structural edits that keep the workbook intact

Inserting a column affects everything to its right: formulas, totals, charts, tables and named ranges on other sheets. KitchenSink4XL moves what should move and is tested against Excel's own behavior for formulas, tables, named ranges, charts and cross-sheet references. When an edit cannot be done safely, it refuses and says why instead of guessing.

## The seven tools and context cost

The server lists seven tools, all from the first message, with nothing to switch on. Every tool except `xl_help` takes an `op` (the operation to run) and `args` (its arguments, including the `path` of the workbook). `xl_help` returns the arguments and a worked example for any operation. Token figures count the full `tools/list` entry the client receives and come from `scripts/measure_surface.py`.

| Tool | Operations | Approx. tokens | What it covers |
|---|---:|---:|---|
| `xl_read` | 18 | ~1.0k | The grid view, labeled reads of cells and ranges, server-side query and aggregation, search, tables, diagnostics, validation checks, formula and style audits, and the inspectors for pivots, connections, external links and VBA |
| `xl_edit` | 10 | ~0.8k | Cell values and formulas, range writes, copy and move, the anchored batch editor, find and replace, and data import |
| `xl_format` | 11 | ~0.9k | Number and cell formats, named styles, format copying, column widths and row heights, merges, conditional formatting, comments, hyperlinks, images, charts and sheet view settings |
| `xl_structure` | 9 | ~0.8k | Worksheets, row and column inserts and deletes that carry their references, tables, named ranges including LAMBDA definitions, sorting, filters and data validation |
| `xl_workbook` | 9 | ~0.8k | New workbooks and copies, backups, document properties, protection, page layout, headers and footers, and export |
| `xl_app` | 12 | ~0.9k | A private hidden Excel on Windows, never your own session: real recalculation, pivot tables, PDF export, sheet images, format conversion, encryption, autofit, goal seek, sparklines and an opens-clean check. Also updates the server when you ask (`update_now`) and turns update notices on or off (`update_notices`) |
| `xl_help` | 0 | ~0.2k | The arguments and a worked example for any operation, and how locations are written |
| **All seven** | **69** | **~5.5k** | Loaded once at the start; nothing to switch on later |

Your AI calls a tool by naming the operation and passing its arguments:

```
xl_read(op="get_grid_view", args={"path": "C:/data/budget.xlsx"})
xl_edit(op="set_cell", args={"path": "C:/data/budget.xlsx", "location": "C1", "value": "Total"})
```

For a recipe, ask your AI to call `get_workflows`. The server's own instructions add about 500 tokens to the figures above. Clients that load tool schemas on first use pay close to nothing until a tool is called.

Administrators who want a narrower server can set `KS4XL_READ_ONLY=1`, which lists only the operations that read, or list the tools or operations to remove in `KS4XL_FORBID` (see Configuration). What is removed no longer counts toward the cost.

## Safety model

- Before every change, the workbook is backed up into two rotating slots in a hidden `.ks4xl-backups/` folder next to it. `manage_backups` lists, restores and prunes them. The slots capture changes made through this server only; an edit made in Excel or another program is not captured.
- Saves are atomic and validated, and each write is checked after saving, with the backup restored if the check fails. Other programs may see the workbook's path missing for a few milliseconds during a save, as with Office's own saves; this server's own calls wait that out. With `verify_com=true` (or `KS4XL_VERIFY_COM=1` for every save), the saved file must also open cleanly in a hidden Excel.
- An ordinary cell edit keeps the rest of the workbook byte for byte (printer settings, shared strings, row heights, pictures, external links and other parts), so Excel opens the result without a repair prompt. The full rewrite used by structural edits carries those parts back, along with custom properties, sensitivity labels and the password to modify.
- Before a save, a scan names any part of the workbook the writer cannot preserve, and refuses the loss unless the call explicitly allows that specific kind.
- Writes honor sheet and workbook protection and refuse with `PROTECTED` when the protection forbids the change.
- Destructive calls keep a labeled snapshot of the workbook as it was, restorable by label.
- Text that merely looks like a formula is never turned into one. The cell's type decides.
- Ambiguous targets are refused with every candidate listed.
- `xl_app` uses its own hidden Excel instance, never the workbook you have open, and leaves no stray Excel processes behind. A workbook open in Excel can be read but not changed.
- A workbook with a password to modify is refused by every `xl_app` operation that would open it for editing, with the remedy, so Excel's Password dialog does not land on your desktop. An encrypted workbook is refused before Excel starts unless a password is given.
- `query_range`, `export_range` and the aggregates include rows hidden by an autofilter. Excel's SUBTOTAL skips them, so totals on a filtered sheet can differ from Excel's status bar.
- Exclude `.ks4xl-backups/` from cloud sync tools; the backup files change on every edit and sync clients can hold locks that slow saves down.

## Privacy

The server runs on your computer. KitchenSink4AI never receives your files or usage data. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. The update check described below contacts PyPI at most once a week and sends only the package name and the installed version. Choosing Update now downloads the new version from PyPI, or the new bundle from GitHub for Claude Desktop. The server's startup banner is switched off, so the framework it runs on makes no version check of its own. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support.

## Configuration

Launch settings are environment variables in the server's configuration. They are read once at start, and no tool call, argument or prompt can change them.

| Variable | What it does |
|---|---|
| `KS4XL_READ_ONLY` | `1` lists only the operations that read, so nothing can change a file |
| `KS4XL_FORBID` | A comma-separated list of tool or operation names to remove from the tool list and from `xl_help`; a call to one refuses with `NOT_AVAILABLE`. `xl_app` here closes every route that starts Excel |
| `KS4XL_ALLOWED_ROOTS` | Restricts every path the server reads or writes to these folders (separated by `;` on Windows, `:` elsewhere) |
| `KS4XL_CALL_TIMEOUT_S` | Wall-clock bound for one call, in seconds (default 300) |
| `KS4XL_MAX_OUTPUT_CHARS` | Ceiling on the size of one reply, in characters (default 80,000) |
| `KS4XL_AUTHOR` | Who signs a note when the call names no author, or names an email address, which is never written as an author: `user` (the default) signs with your Office user name, or your Windows sign-in name when Office has none; `assistant` signs "AI Assistant via KS4XL", as does a call that passes `author="assistant"`, whatever the setting |
| `KS4XL_VERIFY_COM` | `1` makes the hidden-Excel open check the default for every save |
| `KS4XL_VALIDATE_COM` | `1` uses Excel's own verdict for `validate`'s structure check |
| `KS4XL_COM_TIMEOUT` | Maximum time for a COM call |
| `KS4XL_UPDATE_CHECK` | `off` disables the update check completely (`KS4XL_NO_UPDATE_CHECK=1` also works) |
| `KS4XL_UPDATE_NOTICES` | `off` turns update notices off; `on` turns them back on after you turned them off |

Settings from earlier versions are handled for you: an old default (`KS4XL_MODE` set to lite, full or everything, or `KS4XL_PACK_POLICY=auto`) is ignored with a one-line notice, and a setting that locked the tools (`KS4XL_PACK_POLICY=locked` or a `KS4XL_MODE` list of tool groups) stops the start with a message naming what to use instead. To limit the tools yourself, use `KS4XL_READ_ONLY` or `KS4XL_FORBID`.

**Sandboxing.** With `KS4XL_ALLOWED_ROOTS` set, the containment check runs on canonical paths, so `..\` traversal, symbolic links and junctions, short names, case tricks and lookalike folder names are all caught. A blocked call returns a typed error naming the path and the allowed roots before any file is opened.

**Update check and update notices.** At most once a week, after you use a tool, the server asks PyPI in the background which version of kitchensink4xl is newest. The request names only the package and the installed version and contains no documents, file names or identifiers. When a newer version exists and the server runs through uvx or the Claude Desktop extension, your AI app shows a short notice with Update now and Not now, where the app supports one. Otherwise your AI mentions the new version once and offers to update. It never downloads or installs a new version unless you choose Update now in its notice or ask your AI to update it. To turn notices off, ask your AI to turn off KitchenSink4XL update notices; `KS4XL_UPDATE_NOTICES=off` turns them off everywhere, and `KS4XL_UPDATE_CHECK=off` turns the check off completely (`KS4XL_NO_UPDATE_CHECK=1` also works). The setting is kept in the server's own state folder, so it survives updates.

**Claude Desktop extension settings.** **Verify every save with Excel** turns on the hidden-Excel open check (off by default, since it adds a round trip through Excel on every write). **Limit the server to one folder** confines every path the server touches, reads included, to a folder you choose.

<details>
<summary><b>Other clients and module entry points</b></summary>

Point any MCP client at the installed command:

```json
{"mcpServers": {"xl": {"command": "kitchensink4xl"}}}
```

`xl-mcp` is an equivalent command, and `python -m ks4xl_community_launch` starts the same server for a client that wants an interpreter and a module. The shared core, `xlsx_mcp`, does not run on its own: `python -m xlsx_mcp` and `python -m xlsx_mcp.server` stop with a message that names the launcher. Driving Excel (`xl_app`) uses the optional extra, `pip install kitchensink4xl[com]`, which does nothing outside Windows.

</details>

## Known limits

- **`com_render_sheet` needs a signed-in desktop session.** Rendering a sheet to an image goes through Excel's clipboard, which a hidden Excel in a non-interactive session cannot use. Every other `xl_app` operation runs unattended.
- **A desktop tool, not a shared service.** The server runs under your account on your computer, and `xl_app` drives an Excel in your Windows session. It is not designed to be shared by a team from one installation.
- **Deleting a sheet breaks references to it.** Formulas and names that point at a deleted sheet show `#REF!`, as they do in Excel. Renaming a sheet rewrites them. Audit references first if you are unsure.
- **File edits cannot recalculate.** After a write through the file tools, cells downstream keep their old cached values until Excel recalculates. The labels tell you which values are affected, and `recalculate` in `xl_app` fixes them.

## Upgrading from 1.x

Version 2.0 regroups the 69 flat tools of 1.2.x into the seven tools above. Every operation keeps its name except `enable_tools` and `disable_tools`, which are gone, but you now call it as an operation of a tool: `xl_edit` with `op="set_cell"` instead of a tool named `set_cell`. The command names, the package name and the install routes are unchanged.

| If you called | It now lives in |
|---|---|
| `get_grid_view`, `read_range`, `get_cells`, `query_range`, `find_cells`, `get_table`, `get_workbook_metadata`, `diagnose_workbook`, `validate`, `audit_formulas`, `audit_styles`, `get_pivot`, `get_connections`, `get_external_links`, `inspect_vba`, `get_workflows`, `get_server_info`, `com_status` | `xl_read` |
| `set_cell`, `set_cells`, `write_range`, `clear_range`, `copy_range`, `move_range`, `apply_edits`, `replace_cells`, `set_formula`, `import_data` | `xl_edit` |
| `format_cells`, `apply_style`, `copy_format`, `set_dimensions`, `set_merge`, `manage_conditional_format`, `manage_comment`, `manage_hyperlink`, `manage_image`, `manage_chart`, `set_view` | `xl_format` |
| `manage_worksheet`, `modify_grid_structure`, `create_table`, `manage_table`, `manage_name`, `sort_range`, `set_filter`, `clear_filter`, `manage_data_validation` | `xl_structure` |
| `create_workbook`, `copy_workbook`, `manage_backups`, `set_workbook_properties`, `set_protection`, `set_page_layout`, `set_header_footer`, `export_range`, `export_file` | `xl_workbook` |
| `recalculate`, `com_manage_pivot`, `com_export_pdf`, `com_render_sheet`, `com_convert_format`, `com_save_with_password`, `com_autofit`, `com_goal_seek`, `com_set_sparkline`, `com_validate_opens_clean` | `xl_app` |
| `enable_tools`, `disable_tools` | Gone. Every operation is available from the start |

A few values changed spelling, and the old spellings still work: conditional-format and validation values are snake_case (`greater_than`, `last_7_days`, `text_length`), and `manage_chart` takes `at_cell` instead of `anchor`. Your old `KS4XL_MODE` and `KS4XL_PACK_POLICY` defaults need no editing; see the note on settings above. [`docs/surface_map.md`](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/docs/surface_map.md) lists every 1.2.5 tool beside its replacement and every changed value, and the `migrate-from-v1` recipe in `get_workflows` gives your AI the same map. The [changelog](https://github.com/KitchenSink4AI/KitchenSink4XL/blob/main/CHANGELOG.md) lists everything else that changed.

## Testing

The suite has 5,340 tests in `tests/unit`, and it runs on every pull request. Structural edits are checked against Excel's own recalculation, formulas across every documented function family are written and reopened without repair prompts, and the fixture workbooks include features that only Excel can author, such as pivot tables, slicers and shapes. A separate set of gates that drives a real Excel runs locally on Windows.

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

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and Excel are trademarks of the Microsoft group of companies. GitHub and Copilot are trademarks of the Microsoft group of companies. Claude is a trademark of Anthropic. Codex and OpenAI are trademarks of OpenAI. All other trademarks are the property of their respective owners.
