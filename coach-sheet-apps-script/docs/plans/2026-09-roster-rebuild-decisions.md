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

### D10: Team dropdown is seeded once, not re-applied

The plan says the Team dropdown is "seeded Blue, Gold (edit in the sheet after tryouts)". If Sync Extra Player Info re-applied that rule on every run it would undo the coaches' post-tryout edit. The sync now applies the Team rule only when the Team column has no data validation yet; the Returning and Include TRUE/FALSE rules are re-applied every run because they never change.

### D11: Menu item renamed to "Find Emails Not Subscribed to Newsletter"

Not in the plan, but CONTEXT.md says to avoid "mailing list", and the sibling item was already renamed to "Caretakers Not Subscribed to Newsletter". README updated to match.

### D12: WP4 items pulled forward into the WP1 push

The CONFIG rename removed `CONFIG.additionalInfo` and `CONFIG.columns.parent1Email`, which Diagnostics and Build Email List read at runtime. Rather than push a build where Run Diagnostics throws, the Additional Info diagnostics check, the AdditionalInfoAnalysis.gs deletion, and the Build Email List constant rename went out with 3.6 (as their own commit).

### D13: Three TestCleared rows, plus other fixtures

The brief mentions two test signups with last name TestCleared; the live Signups tab has three (GrillTest, Inspect, MobileTest) plus at least one more fixture row (RefreshTest TestFixture). None of them have an SPS Student ID, so the Final Forms formulas were verified against the 62 real players whose SPS Student ID resolves in the export instead of via a scratch cell (see below); no scratch cell was needed or written.

### D14: Practice Roster and Game Roster Prep not driven end-to-end

The claude-in-chrome screenshot capture stopped working part way through the session (CDP capture errors), and those two builders open HTML dialogs whose contents are not exposed to the accessibility tree, so I could not fill them in blind. Static check instead: both builders and `SheetBuilderUtils` look up Full Name, Team, Gender Identification, Grade, and Include In Generated Rosters by header name and read from `ROSTER_FIRST_DATA_ROW`; all five headers exist in the generated Roster with the expected value types (Include is a real boolean, which `isIncludedInGeneratedRosters` accepts). Two-minute check for Steven: 🥏 menu, Build Practice Roster, pick any practice date; the sheet should list the 80 (minus purged test rows) Players with Team blank, Gender Gx/Bx, and Grade filled.

### D15: Version 3.10, not 4.0

Versions went 3.6, 3.7, 3.8, 3.9, 3.10 (one bump per push). "3.10" sorts oddly as text but matches the `3.x` rule in the README.

## Verification status

Everything below was run against the live "2026 Fall Coach Sheets" workbook on 2026-09-07 with the menu driven from Chrome (logged in as madisonultimate@gmail.com) and the tabs read back with `gog sheets get`.

### Verified

- **Run Diagnostics** (3.6 and again on 3.10): all required checks pass, including the new Signups mirror, referenced-headers, Roster header, key formula, and Extra Player Info checks. The only warning is the pre-existing missing 🏃 Attendance tab.
- **Generate Fresh Roster** (3.6): 80 Signups rows produced 80 Roster rows in Last Name then Preferred First Name order. A script compared every one of the 39 columns for every row against the Signups, Final Forms, and Newsletter Subscribers tabs (passthroughs, Full Name, gender collapse and fallback, Grade preference, the three Final Forms flags, Final Forms Cleared?, Date of Birth serials, seattleschools.org blanking, all three newsletter statuses including "not a member" and blank-when-no-email, Media Opt-Out, Profile Complete?, Include fallback, Photo Link): 0 mismatches.
- **Final Forms formulas against real StudentIDs**: 65 players have an SPS Student ID and 62 of them resolve in the Final Forms tab (numeric export IDs matched against text Signups IDs via TO_TEXT). Spot check: Arlen Amsberry (8161416) shows Male, Grade 6 as a number, Parent Signed TRUE, Student Signed FALSE, Physical Cleared TRUE, Final Forms Cleared? FALSE, matching the export row.
- **Sync Extra Player Info** (3.7): rewrote the draft header (D3), added 80 rows, every Full Name formula resolved, no blanks. Write-through test on Inspect TestCleared (Extra Player Info row 72, Roster row 72): Team Blue, Returning TRUE, Include FALSE flowed into the Roster; blanking Include made the Roster Include resolve to TRUE (its Profile Complete?). The test values were then cleared again.
- **Analyze Signups** (3.8): sheet written with the five sections; counts 15, 38, 3, 6, 12 match an independent computation from the raw tabs. It found two real pairs of duplicate signups (Hannah Hitchingham 9654j/j5aq2; Kavani Kenney xddbg/pp6ww, which also share SPS Student ID 8189828) and three SPS Student IDs not in the export (Nova Alleen-Willems, Mateo Andrade, Flip Tillinghast). Worth a look.
- **WP4 grep**: `grep -rn "Additional Info\|additionalInfo\|FIRST_DATA_ROW\b\|parent1Email\|StudentID'" *.gs` returns nothing apart from `ROSTER_FIRST_DATA_ROW` and the Final Forms header literal.
- **Node harness**: 18 formula assertions and 7 Analyze Signups assertions pass against the committed code; every `.gs` file passes `node --check`.

### Not verified

- Build Practice Roster and Build Game Roster Prep end-to-end (D14).
- The Extra Player Info dropdown rules were applied by the script but not read back (gog does not expose data validation). Two-minute check: open Extra Player Info, click any Team cell; the dropdown should offer Blue and Gold, and Returning and Include should offer TRUE and FALSE.
- Show Statistics, Find Emails Not Subscribed to Newsletter, and Caretakers Not Subscribed to Newsletter were rewritten but only syntax-checked; each is one menu click to confirm.
