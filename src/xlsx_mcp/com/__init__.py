"""COM tier for KS4XL: the environment-gated fidelity escape hatch.

Excel is a TRUE multi-instance COM server (unlike POWERPNT); the PowerPoint
singleton lessons do NOT transfer. The danger moved to PID-precise
accounting: self-PID exclusion from GetActiveObject, a disk PID journal, and
a two-phase minutes-grace zombie sweep by owned PID only.

Nothing in this package invokes Excel at import time.
"""
