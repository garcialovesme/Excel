#!/usr/bin/env bash
set -euo pipefail
python3 scripts/build_workbook.py
for f in README.md docs/workbook-spec.md scripts/build_workbook.py vba/TableSetup.bas vba/AllocationEngine.bas vba/Reconciliation.bas vba/NotesSync.bas workbook.xlsx; do
  [[ -f "$f" ]] || { echo "missing required file: $f" >&2; exit 1; }
done
echo "Bootstrap complete: workbook scaffold and VBA modules are present."
