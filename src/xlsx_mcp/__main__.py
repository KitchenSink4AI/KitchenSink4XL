"""`python -m xlsx_mcp` does not start the server in 2.0.

`xlsx_mcp` is the shared editing core. It carries no license policy, so
running it stops with a message that names the launcher. The server starts
from the `kitchensink4xl` or `xl-mcp` command, or from
`python -m ks4xl_community_launch`; `python -m xlsx_mcp.server` stops the
same way.
"""

from __future__ import annotations

from xlsx_mcp.server import main

if __name__ == "__main__":
    main()
