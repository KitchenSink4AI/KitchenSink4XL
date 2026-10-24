"""Core engine for KS4XL: the cross-cutting infrastructure built once,
centrally, before the feature families.

Ported from the family: errors, sandbox, safesave, _regex. Written for the
grid: package, hazard, verify, locate, refs, calc.
"""

# Cell text is read and written as Excel stores it (_xHHHH_ escapes), for
# every load and save in the process: reads, the model a write edits and the
# verify read-back all hold the characters Excel shows (core.xltext).
from . import xltext as _xltext
from . import readcompat as _readcompat

_xltext.install()
# Markup Excel opens clean and openpyxl's reader refused (protection
# attributes in another letter case, <v>true</v>) loads (core.readcompat).
_readcompat.install()
