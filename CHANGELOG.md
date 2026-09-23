# Changelog

### 1.2.5
- `enable_tools` now says what to do when a client fixed its tool list at startup. Put a comma list of packs in `KS4XL_MODE`, restart the app or session, and then start a new worker. Claude Code can instead enable packs in its main session before it starts workers. A locked policy still caps the surface.

### 1.2.4
- Documentation and listing only. The MCP Registry entry moves to the io.github.KitchenSink4AI namespace and is published from the release workflow. The bundle manifest carries contact details. No change to tools or behaviour.

### 1.2.3
- A workbook mutation that runs longer than ten minutes keeps its write lock. The lock was previously broken on age alone even while the holding process was alive, which let a second writer into the same file and reopened the read-modify-save race the lock exists to close.
- A lock written by another machine is never reclaimed from this one. A PID number on a network share says nothing about a process on a different computer, so a foreign-host lock now makes the waiter wait and then refuse by name.
- `com_render_sheet` renders. It previously returned a success result without producing an image; it now produces the image, or refuses and says why.
- Charts this server creates are drawn with their axes.

### 1.2.2
- The license story is stated correctly and in one place. KitchenSink4XL is dual-licensed: AGPL-3.0 for anyone whose use meets the AGPL's terms, including its source-sharing obligations, and a commercial license for organizations that need to ship it inside a closed product (licensing@kitchensink4.ai). NOTICE.md previously said "free for any use (personal, academic, commercial) under AGPL terms", which contradicted the commercial model.
- Copyright is attributed: Alvut Consulting, LLC, named at the top of LICENSE, in NOTICE.md, in the README, in the package author field, and as the grantee in the CLA.
- The landing page serves its fonts itself instead of loading them from Google's CDN.

### 1.2.1
- The server says what it is. A connected agent now reads "KitchenSink4XL (kitchensink4xl on PyPI), part of the KitchenSink4AI suite" at the head of the instructions it receives, and `get_server_info` reports the product name, the package to install, the documentation homepage, and the three sibling packages.
- Published addresses moved to kitchensink4.ai, and the registry record is republished at this version.
- The tool figure is stated as measured on every surface: 129 workbook operations across 69 tools, 67 workbook tools plus the two pack toggles.

### 1.2.0
- Read responses are dramatically smaller (seven to eleven times on the heavy calls) with nothing lost; the token estimator now matches the wire exactly, and every published figure is re-measured.
- A performance bug that reopened the workbook per cell is fixed; large cell reads are near-instant.
- Update notice: a weekly, disclosed check for newer releases, off with KS4XL_UPDATE_CHECK=off.
- The install screen and info card are rewritten in plain language.
