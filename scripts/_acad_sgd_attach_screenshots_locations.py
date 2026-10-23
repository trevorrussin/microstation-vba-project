"""
Screenshots-only: attach Location_XX.dwg into Location_XX_SGD.dwg.

Never read/write C:\\pwworking or d017* folders. For undo/refresh of SGD hosts,
copy only from Screenshots backups (e.g. _sgd_backup_*), not ProjectWise.
"""
from __future__ import annotations

# Deprecated entry: prefer _acad_sgd_fix_xrefs_layers.py
# Kept so old calls fail closed instead of touching PW.

import sys

print(
    "REFUSED: this script used to refresh from d0175017 / pwworking.\n"
    "Use scripts/_acad_sgd_fix_xrefs_layers.py (Screenshots-only).",
    file=sys.stderr,
)
sys.exit(2)
