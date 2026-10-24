# Community and Business editions

Both editions use one Excel editing core. Safety, backup and document integrity
behavior are identical. Community is free under the AGPL-3.0. Business adds a
Windows installer, a signed update channel and support under the Business terms;
current options are at https://kitchensink4.ai/pricing/.

## Community

The public source is AGPL-3.0-only. Its initialization instructions identify the
edition and include this contextual warning:

> This deployment likely violates AGPL-3.0 §13. Providing network access to a modified version without offering users its source code is a license violation, not a gray area. Certain commercial and enterprise deployments require a paid license. Purchase at kitchensink4.ai/licenses or contact licensing@kitchensink4.ai.

The assistant should use that warning only when the user describes that
specific unmet obligation. Missing facts call for a targeted question.
Commercial document work alone is not a violation. AGPL compliance remains
an alternative to a commercial license. The server neither scans documents
for business clues nor refuses tools over licensing. Client assistants may
interpret server instructions differently; this is guidance, not a legal
detector or a guaranteed client behavior.

Licensing email and page are returned by get_server_info. The licensing URL
is the intended destination, not a claim that a checkout is already live.
The existing Community update check retains its operator off switch.
Both entry points suppress FastMCP's banner, whose own version checker could
otherwise contact PyPI at startup independently of the product off switch.

## Business

The private build is kitchensink4xl-business, with a +business build-version
suffix and its own executable name. It has commercial distribution metadata
and a separate commercial notice. The notice is not a substitute for an
executed commercial agreement. Nothing establishes new pricing, capacity
limits, or support promises.

Business never calls the public update endpoint, even if the Community
update environment variable is turned on. Receipt verification is entirely
local. A missing, malformed, wrong-product, tampered or unknown-key receipt
is reported without blocking document tools or declaring an infringement.

A receipt whose updates-through date precedes this build's release date
reports updates_not_covered; an older covered build stays covered regardless
of today's date. This is a perpetual entitlement, not a clock-based runtime
subscription. Download authorization and update distribution are handled outside the
product.

get_server_info() reports edition and verification status without a company
name. get_server_info(include_licensee=true) explicitly adds the signed
company identity. The default remains suitable for public support reports.
Initialization does not disclose a company's identity or read a receipt.

## Build both editions

Use a development Python environment with the project's dependencies plus
build, setuptools>=77 and cryptography>=42. In the source root:

```text
python scripts/build_editions.py --edition both --release-date YYYY-MM-DD --output PATH_TO_NEW_OUTPUT
```

The Community directory contains two wheels: kitchensink4xl carries the editing
core and launcher, and kitchensink4-license-community supplies the license
notices at exact version 0.2.0. Business retains three wheels: its editing core,
license policy adapter and launcher. Each MCPB archive contains its manifest,
icon, license, installation guide and all of that edition's wheels.

The license policy adapters are shared across the whole KitchenSink4AI suite
rather than owned by this repository. The build reads them from a pinned
canonical revision, so the adapter wheel shipped here is the same one every
other editor ships.

Both editions are built from one source revision, and each is labelled for the
grant it ships under. The Community app includes the core at the public
version under AGPL-3.0-only. The separate Business core is the same source at
VERSION+commercial under LicenseRef-KitchenSink-Commercial, carrying the
commercial notice as its LICENSE. Two artifact identities, never one version
label covering two different grants. The build refuses to finish if any wheel
in a distribution declares the other edition's license. Output edition
directories must not already exist, preventing accidental artifact overwrite.
Builds stage in the operating system's temporary directory, never inside the
source tree.
The script does not publish, install into your active environment or modify
client settings. Wheels are built from src and an explicit metadata allowlist,
not the test files or the rest of the repository.

For an issuer-enabled Business build, also supply
--public-key ISSUER_PUBLIC.pem --key-id IDENTIFIER. Only an Ed25519 public key
is embedded. Without one, a build can report missing/unverified receipts
but cannot authenticate an issuer. Never embed an issuer private key.

## Install and connect

1. Extract the MCPB archive, or use its accompanying wheel.
2. Create a dedicated environment: python -m venv PATH_TO_EDITION_ENV.
3. Install all of the edition's wheels using that environment's Python:
   Community: PATH_TO_EDITION_PYTHON -m pip install PATH_TO_APP_WHEEL PATH_TO_POLICY_WHEEL.
   Business: PATH_TO_EDITION_PYTHON -m pip install PATH_TO_CORE_WHEEL PATH_TO_POLICY_WHEEL PATH_TO_LAUNCHER_WHEEL.
   Installation may fetch dependencies; runtime license verification does not.
   For an offline machine, prepare the dependencies on a connected machine
   first and install with --no-index --find-links PATH_TO_WHEELHOUSE.
4. Import the MCPB into your client and set its Python executable field to
   the absolute path of that environment's Python.
5. For Business, optionally set the receipt path field to your signed JSON
   file. This maps to KS4XL_LICENSE_FILE.
6. Call get_server_info and confirm the edition before opening a document.

Community and Business share the xlsx_mcp editing core and differ only in
which license policy adapter and launcher are installed beside it. Install
them in SEPARATE environments; installing both policy adapters into one
environment is not supported, and each launcher refuses to start when it
finds the other edition's policy installed alongside it. Both generated MCPBs launch the specified installed environment and verify
it before starting the server; an incorrectly selected environment
fails with a setup explanation. They do not use uvx against public PyPI. This prevents a Business launcher from silently
starting Community. For source development, the original public manifest
still pins the Community PyPI package and is not the Business installer.

## Validation

Run tests with the worktree src first on PYTHONPATH and the project dependency
versions installed. Use pytest -m "not com" for file-based and mocked COM
coverage. Real Excel COM gates require a separate deliberate run with disposable
workbooks, and the command above does not run them.
