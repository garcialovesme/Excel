# Copilot Instructions for Excel Workbook Project

## Project Overview

This is a **multi-module Excel accounting system** built with VBA macros. The workbook automates allocation calculations, reconciliation tracking, and note synchronization using Excel tables as the data layer.

### Core Architecture

**VBA Modules** (in `vba/`):
- `TableSetup.pas` - Initializes 12 core tables (tblAccounts_Current, tblAllocations, tblNotes, etc.)
- `AllocationEngine.pas` - Performs financial allocation calculations
- `Reconciliation.pas` - Identifies and logs reconciliation issues
- `NotesSync.pas` - Manages asynchronous note queuing/syncing (SaveNoteToQueue → SyncPendingNotes → MergeSyncedNotesIntoLocalHistory)

**Execution Order** (critical):
1. TableSetup.BootstrapTables
2. AllocationEngine.RunAllocationBuild
3. Reconciliation.RefreshReconIssues

**Data Pattern**: Configuration + Overrides
- `tblRanges_Default` + `tblRangeOverrides` - Avoid hardcoding values; use lookup patterns
- All user data flows through core tables first
- Integration tables use XLOOKUP for cross-table references

### Key Workflows

**Building the Workbook**:
```bash
./scripts/bootstrap_repo.sh  # Validates all required files exist, runs build_workbook.py
python3 scripts/build_workbook.py  # Generates minimal workbook.xlsx scaffold
```

The scaffold structure requires **manual VBA module injection** (copy-paste into Excel VBA editor) - this is by design for complex macro development.

**Module Organization**:
- Each `.bas` file is Option Explicit (strict type checking)
- All modules are public; use `ModuleName.SubroutineName()` to call across modules
- Queue tables (tblNoteSyncQueue, tblNoteSyncLog) decouple note operations from UI

### Project-Specific Conventions

1. **Naming**: All table names prefixed with `tbl` (tblAccounts_Current, tblStaff, tblConfig)
2. **Queue Pattern**: Async operations use queue/log table pair (e.g., NoteSyncQueue → NoteSyncLog)
3. **VBA Dependencies**: Requires Excel 365 or 2021+ (XLOOKUP function required)
4. **Error Handling**: tblReconIssues is the canonical place for problems (populated by RefreshReconIssues)

### When Adding Features

- Start with table structure in spec (`docs/workbook-spec.md`)
- Add VBA module/subroutine to appropriate `.bas` file
- Follow queue-based pattern if operation is async or needs logging
- Update run order in bootstrap checks if module initialization changes
- Use `MsgBox` for placeholder functions (see existing modules)

### File Structure Reference

- `docs/workbook-spec.md` - Authoritative spec for all 12 tables and their relationships
- `vba/TableSetup.bas` - Examine this first to understand table initialization patterns
- `scripts/bootstrap_repo.sh` - Validates project integrity; add new required files here
- `workbook.xlsx` - Auto-generated; do not edit directly
