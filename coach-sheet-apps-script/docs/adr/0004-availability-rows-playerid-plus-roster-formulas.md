---
status: accepted
date: 2026-09-14
---

# Availability rows and printout rows are keyed by a typed PlayerID and derive every other player column by formula

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

## Amendment (2026-09-14): the printouts carry a hidden PlayerID too

The first version of this decision left Build Practice Roster and Build Game Roster Prep joined to the availability sheets by Full Name, resolved by header, because those sheets had no PlayerID of their own. The coach asked for the printouts to carry a PlayerID as well, and they now do: `PlayerID` is the last base column of the coach printouts (`CONFIG.rosterPrintoutBaseColumns.playerId`, index 6) and the last column of the parent game roster, written as a value by `seedPrintoutPlayerRows` and hidden after the build so the printed page is unchanged. Every other cell on a printout row is a per-row lookup on that PlayerID: Full Name, Team, Gender, and Grade into the Roster, and each availability, activation status, and note column into the availability sheet's PlayerID column. Apply Activation Status joins the prep sheet back to Game Availability by PlayerID the same way. The only Full Name joins left are the fallback for an availability sheet built before 3.28 (no PlayerID column yet, which `availabilityJoin` detects and warns about) and the Sheet Builder's custom sheets, which were out of scope.

Consequences added: two Players with the same Full Name no longer collide on a printout; a preferred-name edit reaches every open printout without a rebuild; Build Practice Roster's update path widens a sheet built before the column existed and clears any availability dropdown left in the column it now uses.
