# Relocate unpaid invoice rows

Google Apps Script in `Code.gs`. `moveNotPaidRows()` scans the configured tail of filled invoice project rows, recognizes unpaid status in the configured columns, clears the documented scheduling fields, appends copies at the real project-data end and deletes the original rows. `sortPayouts()` transforms the batch-payment tab; it is not a read-only sort preview.

- Read source configuration and helpers before changing the workflow; the README is only a label. Preserve `KEY_COL`-based last-row detection so unrelated U-Y data below the projects does not expand the scan.
- Preserve bottom-up grouped deletion so earlier row numbers remain valid, the configured scan window and exact dropdown/status matching.
- `../Send to Payout script` calls `moveNotPaidRows()` in the same deployed Invoices Script project and must re-read invoice rows after this function shifts them. Coordinate public entry-point and row-schema changes.
- Integrate menu changes into the existing `onOpen()`; adding another in the same Apps Script project silently replaces it. All globals are shared with the payout and other invoice scripts.
- These entry points clear cells, move/delete sheet rows and rewrite payout values. Do not run them as passive diagnostics.
- No local automated test harness or live-deployment manifest is documented. Validate a future functional change on a suitable copy with representative statuses and trailing unrelated data.

When working in `D:\Dev`, consult `../workspace-context/PROJECTS.md` for workspace routing and `../workspace-context/WORKSPACE_CONTEXT.md` for shared context if present. These workspace files are local context, not a dependency of the project. Keep the cross-project contract current without copying invoice records.
