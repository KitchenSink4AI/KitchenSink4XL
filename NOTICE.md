# KitchenSink4XL: License Notice

KitchenSink4XL is Copyright (c) 2026 Alvut Consulting, LLC.

KitchenSink4XL is dual-licensed:

**AGPL-3.0 (open source).** Free for anyone (individuals, academics, and businesses) for any use that complies with the AGPL's terms. Those terms include sharing source, including your modifications, when you distribute the software or make it available over a network.

**Commercial license.** For organizations that want to build KitchenSink4XL into their own products or services without the AGPL's source-sharing obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line of Alvut Consulting, LLC.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and Excel are trademarks of the Microsoft group of companies.

## Third-party software

KitchenSink4XL does not redistribute any of the packages below. They are installed from their own publishers by `pip` or `uv`, under their own licenses. The package name is the first column and the license the second in every table.

### Required dependencies

| Package | License | Purpose |
|---|---|---|
| `fastmcp` | Apache-2.0 | MCP server framework (https://github.com/jlowin/fastmcp) |
| `openpyxl` | MIT | .xlsx object model (https://openpyxl.readthedocs.io) |
| `lxml` | BSD-3-Clause | XML processing (https://lxml.de) |
| `regex` | Apache-2.0 AND CNRI-Python | Regular expressions with match timeouts (https://github.com/mrabarnett/mrab-regex) |
| `packaging` | Apache-2.0 OR BSD-2-Clause | Version comparison (https://github.com/pypa/packaging) |

`et-xmlfile` (MIT) is installed as a dependency of `openpyxl`.

### Optional dependencies

| Package | License | Extra | Purpose |
|---|---|---|---|
| `pywin32` | PSF | `com` | Windows COM automation of a private, hidden Excel instance (https://github.com/mhammond/pywin32) |
| `pytest` | MIT | `dev` | Test runner (https://docs.pytest.org) |
| `pytest-timeout` | MIT | `dev` | Test timeouts (https://github.com/pytest-dev/pytest-timeout) |
| `pillow` | MIT-CMU | `dev` | Image handling in tests (https://python-pillow.org) |
| `psutil` | BSD-3-Clause | `dev` | Process inspection in tests (https://github.com/giampaolo/psutil) |
