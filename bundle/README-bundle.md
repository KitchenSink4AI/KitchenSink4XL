# KitchenSink4XL .mcpb Bundle

This directory builds the MCP Bundle (`.mcpb`) for one-click Claude Desktop
installs and Smithery's Local (MCPB Bundle) listing. It follows the
KitchenSink4Word bundle exactly; only the package name, the version pin, the
icon, and the `user_config` block differ.

**Design: thin uvx launcher.** The archive carries only `manifest.json` and
`icon.png` (a few kB). The manifest's `mcp_config` tells the client to run
`uvx kitchensink4xl[com]==<version>`, so the actual server is fetched from
PyPI at first launch. Nothing from `src/` is bundled, because pywin32's `.pth`
bootstrap and compiled wheels are not portable under `server/lib` bundling.
End users need uv installed (`uvx` ships with uv); this is the standard
arrangement for Python MCP servers.

The `[com]` extra is in the pin on purpose. It pulls pywin32 on Windows and
is a declared no-op everywhere else (`pywin32; sys_platform == 'win32'`), so
the Excel tool (`xl_app`) is actually usable for
the Desktop audience that installs by double-click.

## Files

| File | Role | In the archive? |
|---|---|---|
| `manifest.json` | The bundle manifest (binary/uvx shape, manifest_version 0.3) | yes |
| `icon.png` | 512x512 icon: the KitchenSink4AI app icon (ringed symbol on the dark tile), the same file in every product bundle | yes |
| `smoke_test.py` | Post-build verification | no |
| `.mcpbignore` | Keeps everything but manifest + icon out of the archive | no |
| `kitchensink4xl.mcpb` | The packed bundle | output |

## User configuration

Claude Desktop renders `user_config` as a settings form on the extension and
passes the answers to the server as environment variables:

| Setting | Environment variable | Default |
|---|---|---|
| Verify every save with Excel | `KS4XL_VERIFY_COM` | off |
| Limit the server to one folder | `KS4XL_ALLOWED_ROOTS` | empty (no limit) |

Both are read at call time by the server, so a change takes effect on the
next server start. The verify setting maps to a boolean the server parses as
`1/true/yes/on`; anything else, including the unchecked `false`, leaves deep
verification off. The folder setting is deliberately ONE directory: the
server's own `KS4XL_ALLOWED_ROOTS` accepts an `os.pathsep`-separated list,
but the bundle host's join behavior for multi-value config is not pinned
down, so the form asks for one folder and power users set the full list in
their own client config.

The manifest also sets one fixed variable, `KS4XL_INSTALL_ROUTE=mcpb`. It is
not a setting: it tells the server it runs from this bundle, so `update_now`
offers the new bundle for Claude Desktop to install rather than a `uvx`
refresh.

## Rebuild per release

The pinned version appears in `manifest.json` **twice** and must match the
version just published to PyPI:

- `"version": "X.Y.Z"`
- `"args": ["kitchensink4xl[com]==X.Y.Z"]`

The pin is deliberate: uvx caches resolved environments, and an unpinned spec
can serve a stale cached version. **Every PyPI release therefore needs a
manifest bump and a repack.** After editing the two version strings:

```
mcpb pack bundle bundle/kitchensink4xl.mcpb
.venv\Scripts\python.exe -X utf8 bundle\smoke_test.py
.venv\Scripts\python.exe -X utf8 scripts\bundle_release.py bundle\kitchensink4xl.mcpb
```

`scripts/bundle_release.py` runs the same check the server runs before it opens a
downloaded bundle, then writes `kitchensink4xl-X.Y.Z.mcpb` (the same bytes) and
`kitchensink4xl-X.Y.Z.mcpb.sha256` beside the packed file.

(`mcpb` comes from `npm install -g @anthropic-ai/mcpb`. The smoke test
verifies archive contents, manifest shape, the version pin, and performs a
live stdio initialize handshake through uvx against the published package, so
run it only after the new version is actually on PyPI.)

The icon is the KitchenSink4AI brand app icon, replaced as a file when the brand art changes.

## Claude Desktop install (manual test / direct distribution)

Double-click `kitchensink4xl.mcpb`, or drag it into the Claude Desktop
window, or Settings > Extensions > Advanced settings > Install Extension.
Desktop writes the manifest's `mcp_config` into its config; the first launch
resolves the package from PyPI (needs network and uv on the machine).

## Publish to Smithery (Local / MCPB Bundle listing)

Authentication is yours, not an agent's: log in yourself with the Smithery
CLI (`smithery login`, or `npx @smithery/cli login`) before publishing.

```
smithery mcp publish bundle/kitchensink4xl.mcpb -n <your-org>/kitchensink4xl
```

**Tools-array constraint (a limitation of the Smithery CLI):** the Smithery registry
rejects (HTTP 400) any bundle whose manifest declares a `tools` array,
because the MCPB tools schema is incompatible with the registry's Tool
schema. This bundle's manifest deliberately carries **no** `tools` array, so
it publishes as-is. `smoke_test.py` fails if `tools` sneaks into the packed
manifest.

## Attach to a GitHub release

Attach the `.mcpb` as a release asset (do not commit the binary):

```
gh release upload vX.Y.Z bundle/kitchensink4xl-X.Y.Z.mcpb bundle/kitchensink4xl-X.Y.Z.mcpb.sha256 --repo KitchenSink4AI/KitchenSink4XL
```

Both files go on every release from 2.0.0 on: `update_now` downloads the
versioned bundle and its `.sha256` from the release of the new version, and the
site links the versioned bundle.

Because the manifest pins the PyPI version, each release's asset is specific
to that release; never re-upload an old `.mcpb` to a new tag.

## Platform note

`compatibility.platforms` declares all three platforms, which is where this
bundle differs from the Word one. macOS is untested in this release. The file
tier here needs no Excel installation and runs on Windows and Linux; only the
`xl_app` tool asks for Windows with Excel, and it refuses honestly when it is
not there. The manifest format has
no per-tool capability notes, so that story lives in `long_description`, the
README, and that tool's own refusals.

## Signing

`mcpb sign` is optional and currently skipped; unsigned bundles install with
a warning. Revisit if a directory submission requires a signature.

## Before the first release, verify by hand

1. `uvx kitchensink4xl[com]==<version>` resolves and starts from a clean uv
   cache (the extras-in-spec form is what the manifest ships).
2. The two `user_config` settings arrive as the environment variables above,
   with the checkbox off producing something the server reads as off.
3. Desktop's install path on the machine under test (the Microsoft Store
   build has no `.mcpb` file association; Settings > Extensions > Advanced
   settings > Install extension is the reliable route).
