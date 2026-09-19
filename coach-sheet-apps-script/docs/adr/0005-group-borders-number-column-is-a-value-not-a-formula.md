---
status: accepted
date: 2026-09-19
---

# Draw Group Borders' "#" column is a plain value, not a formula

Draw Group Borders (`GroupBorders.gs`) can optionally populate an existing `#` header column with each row's 1-based position within its group, resetting to 1 at every group boundary: the same numbering `populateNumberColumn` already writes as a spreadsheet formula for Practice Roster and Game Roster Prep (`=IF(A1<>A2,1,B1+1)`, resetting on the group-by columns). This one writes a plain value instead, which breaks the sheet-wide rule that every derived column is a formula (ADR 0003, ADR 0004, `AGENTS.md` "Sheet conventions").

The deviation is deliberate: Draw Group Borders' border itself cannot be a formula (Google Sheets has no formula-driven border), so it is necessarily a static snapshot that goes stale after a later resort until the coach reruns the command. Numbering the same rows with a formula would make the `#` column self-heal after that resort while the border next to it stayed wrong, a visible, confusing disagreement between two cells describing the same group on the same row. A plain value keeps both artifacts stale (and both fixed) together: rerun the command, everything is right again; don't, and both are equally wrong.

## Considered options

- **Formula, matching `populateNumberColumn` and the sheet-wide convention.** Rejected: correct in isolation, but produces the split-staleness problem above, and building an arbitrary-column-letter formula string for a value only ever shown next to a value-based border added complexity for a numbering scheme this feature doesn't otherwise need.
- **Plain value (this decision).** Simpler code (reuses the same in-memory group-boundary computation the border already does) and a single, consistent "rerun when stale" story for the whole feature.

## Consequences

- Unlike every other derived cell in this codebase, a Draw Group Borders `#` cell does not update itself when the sheet is resorted or edited; the coach must rerun the command, same as for the border.
- `populateNumberColumn`'s formula-based numbering for Practice Roster and Game Roster Prep is unchanged; this ADR governs only the `#` column Draw Group Borders itself writes.
