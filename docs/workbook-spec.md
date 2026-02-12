# Workbook Spec (MVP)

Minimum version: Excel Microsoft 365 desktop or Excel 2021+ (XLOOKUP required).

Core tables:
- tblAccounts_Current
- tblStaff
- tblRanges_Default
- tblRangeOverrides
- tblDigitMap
- tblAllocations
- tblNotes
- tblNoteTypes
- tblConfig
- tblReconIssues
- tblNoteSyncQueue
- tblNoteSyncLog

Run order:
1) TableSetup.BootstrapTables
2) AllocationEngine.RunAllocationBuild
3) Reconciliation.RefreshReconIssues
