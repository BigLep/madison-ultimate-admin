# Fall 2026 roster rebuild: decisions log

Companion to [the plan](2026-09-roster-rebuild.md) and [ADR 0001](../adr/0001-roster-keyed-by-signups-playerid.md). Written during the autonomous implementation run on 2026-09-07. Each entry records what was unclear, the options considered, what was chosen, and why. The last section lists what was and was not verified against the live workbook.

## Judgment calls

### D1: Version numbering

AGENTS.md says versions are `2.x`, but Code.gs was already at `3.5` and the README says `3.x`. Chose to continue `3.x` (first push is 3.6) because that is what the live menu shows and what the README documents; the AGENTS.md wording is stale. Not changed in AGENTS.md since that file is outside this plan's scope.

### D2: `.clasp.json` filePushOrder does not list AdditionalInfoAnalysis.gs

The plan says to drop AdditionalInfoAnalysis.gs from `filePushOrder`, but it was never listed there (clasp pushes every file regardless; the list only orders them). Nothing to drop; the two new files ExtraPlayerInfo.gs and AnalyzeSignups.gs are added after Diagnostics.gs so they load after Code.gs.

### D3: Extra Player Info tab already exists with draft headers

The live workbook already has an "Extra Player Info" tab whose header row is `Player Id, Full Name, Returning Player, Team` and no data rows. The plan's header is `PlayerID, Full Name, Team, Returning, Include In Generated Rosters`. Options: (a) fail and ask; (b) always overwrite the header; (c) overwrite only when there are no data rows, otherwise stop with an error naming the mismatch. Chose (c): it fixes the live draft without risk, and once coaches have authored rows a silent header rewrite could reorder meaning under their data.

### D4: Roster tab still has the spring five-row metadata layout

The live 📋 Roster has the spring layout (label in column A, rows 1 to 5 metadata, the real headers offset by one column). Generate Fresh Roster clears all contents, notes, and data validations, unfreezes and refreezes row 1, and rewrites from A1, as the plan says. Conditional formatting is left untouched per the plan. Data validations are cleared too (not mentioned in the plan) because a formula-only sheet has no use for leftover checkbox or dropdown rules, and a checkbox rule under an array formula renders confusingly.

### D5: Guard Final Forms lookups on a blank SPS Student ID

`XLOOKUP(TO_TEXT(""), TO_TEXT('Final Forms'!A:A), ...)` matches the first blank row of the Final Forms tab instead of failing, so a signup with no SPS Student ID would silently read a blank Final Forms row. The plan's pattern only guards on PlayerID. Added an inner `IF(spsStudentId="", <missing>, ...)` guard to every Final Forms lookup, where `<missing>` is `""` for text columns and `FALSE` for the three boolean flags (per the ADR: a missing SPS Student ID makes Final Forms Cleared FALSE). Same result on every normal row, correct result on the unmatched rows.

### D6: Profile Complete reads Roster columns, not Signups directly

The plan's column table marks Profile Complete? as derived (D) from Grade, Date of Birth, and Caretaker 1 Email. Roster Grade prefers Final Forms, so a player with a Final Forms grade but a blank Signups grade counts as complete in the Roster while Analyze Signups section 5 (which reads Signups directly, per the plan) would list them. Followed the plan for both. In practice Signups Grade is a Seeded Field copied from Final Forms at the join, so the two rarely diverge.

### D7: `LET` inside `ARRAYFORMULA`

Several columns need the same lookup twice (Signup Gender, Include In Generated Rosters, Date of Birth, Student Personal Email, Photo Link). Used `LET` to bind the lookup once instead of repeating the XLOOKUP. `LET` is supported in Google Sheets and evaluates elementwise inside `ARRAYFORMULA`; this keeps the formulas readable in the cell.

### D8: Date of Birth accepts either an ISO string or a real date

IMPORTRANGE currently delivers the Signups Date of Birth as an ISO text string (`2013-05-15`), and the plan says `DATEVALUE` it. If the portal ever writes real dates, `DATEVALUE` of a number errors. The formula returns the value as-is when `ISNUMBER`, else `DATEVALUE`, else blank.

### D9: Testing seam

There is no test runner for Apps Script in this repo. The seam chosen for /tdd is the pure part: `ROSTER_COLUMNS` formula builders and the column-letter resolvers take plain objects and return strings, so a Node harness (kept in the session scratchpad, not committed) loads Code.gs with stubbed globals and asserts the exact formulas for a known letter map. Every `.gs` file is syntax-checked with `node --check` before each push as the "typecheck" step.

## Verification status

(Filled in as work packages land.)
