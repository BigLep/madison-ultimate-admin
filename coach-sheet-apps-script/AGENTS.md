# Agent Instructions for Apps Script Deployment

## Before Deploying

**ALWAYS** increment the `SCRIPT_VERSION` constant in `Code.gs` before running `clasp push`. With every change you deploy, increment the minor version (x in `2.x`) first (e.g. `'2.1'` → `'2.2'`).

Use version format `2.x`; increment x for each release.

## Deployment Process

1. Update `SCRIPT_VERSION` in `Code.gs`
2. Run `clasp push` to deploy changes
3. Test the deployed functionality

## Commits

**All changes** to this project must be committed using [Conventional Commits](https://www.conventionalcommits.org/): use a type and optional scope (e.g. `feat(menu): add X`, `fix(roster): correct Y`, `chore(coach-sheet): Z`), and add a short body when it helps.

## Sheet conventions

- **Only key columns are typed values; everything derived is a formula.** The 📋 Roster writes PlayerID in column A and a formula in every other cell (ADR 0003). Practice Availability and Game Availability follow the same rule (ADR 0004): `PlayerID` in column A is the only per-player value the script writes, and Full Name, Grade, and Gender Identification are per-row Roster lookups. Practice rosters and game roster prep sheets (both audiences) carry a hidden `PlayerID` column (`CONFIG.rosterPrintoutBaseColumns.playerId`, last in the parent view) seeded by `seedPrintoutPlayerRows`, and every other cell on the row is a lookup on it via `fillLookupColumn` (Roster for Full Name/Team/Gender/Grade, the availability sheet for date columns, joined by `availabilityJoin`). All of these use `playerIdLookupFormula` in `SheetBuilderUtils.gs`, the Roster's house shape: `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`. Never copy a value the Roster already holds and never join sheets on Full Name; a copied value goes stale the moment the Roster changes, and a name join collides on duplicates.
- **Find columns by header, never by position.** Column letters in generated formulas come from the live header row (`readRosterTable(...).col`, `getExistingColumns`, `findAvailabilityColumns(...).playerIdColumn`); do not hardcode `A:A` or a letter for a column that a build could move.
- **Seeded rows are never deleted or reordered** by any build; the coach removes rows by hand.

## Notes

- Do NOT modify the version field in `appsscript.json` - use `SCRIPT_VERSION` in `Code.gs` instead
- The version helps track which deployment is currently active
