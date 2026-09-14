---
status: accepted
date: 2026-09-14
---

# Availability rows are keyed by a typed PlayerID and derive every other player column from the Roster by formula

Practice Availability and Game Availability hold one row per Player that the family portal reads and writes by PlayerID. Until 3.27 the build wrote Full Name, Grade, and Gender Identification into each seeded row as copied values, with Full Name in column A because the roster prep sheets looked players up against `A:A`. Copied values go stale as soon as the Roster changes (a preferred name edit, a Grade correction from Final Forms, a gender resolution) and are indistinguishable from a coach's edit, so nothing ever reconciled them. The coach hand-patched the live fall 2026 Practice Availability to formulas before this change landed.

We decided the availability sheets follow the Roster's own rule (ADR 0003): the key is a typed value and everything derived is a per-row formula. Column A is `PlayerID`, the only per-player value the script writes. Full Name, Grade, and Gender Identification are per-row lookups of that PlayerID against the 📋 Roster in the Roster's house shape, for example `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`, with the Roster column letters read from the live Roster header at build time. Every build rewrites those three cells on any row that has a PlayerID, so a sheet built by an older version converts itself; rows with no PlayerID keep whatever the coach typed, since that name is what identifies the row to fix. Readers that need a Player's Full Name in an availability sheet (Build Practice Roster, Build Game Roster Prep, Apply Activation Status) find the column by header rather than assuming column A.

## Considered options

- **Keep copied values and re-sync them on each build.** Rejected: a sync still leaves a window of staleness, needs its own diff logic, and cannot tell a stale copy from a deliberate coach edit.
- **Key the availability sheets by Full Name and keep PlayerID as a secondary column** (the 3.27 layout). Rejected: the portal matches on PlayerID, so PlayerID is the real key; putting Full Name first only existed to serve hardcoded `A:A` lookups, which are cheaper to fix than a second key is to maintain.
- **Typed PlayerID plus Roster formulas** (this decision). Same shape as the Roster and Extra Player Info, so one convention covers every PlayerID-keyed sheet.

## Consequences

- A new availability sheet is laid out `PlayerID | Full Name | Grade | Gender Identification | date columns...`; an existing sheet keeps its column order and is still found by header, so a sheet with Full Name in column A keeps working (the formulas reference whichever column holds PlayerID).
- Renaming a Player, or fixing their Grade or gender in a Source, shows up in the availability sheets on the next recalculation, with no build.
- `findAvailabilityColumns` returns `fullNameColumn`; every prep-sheet XLOOKUP into an availability sheet uses it instead of `A:A`.
- Build Practice/Game Availability reports how many existing rows had their derived cells rewritten as formulas.
