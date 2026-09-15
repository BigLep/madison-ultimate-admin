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

### D16: Analyze Signups reads Final Forms names from columns D and E

Section 2 (Final Forms students with no signup) needs a name, and the plan's fixed-position list covers only A, P, Q, U, W, AB. The export's First Name and Last Name have sat in D and E every season (the portal plan notes the same layout), so the report reads them by position like the rest. If finalforms-export ever validates headers, D and E should join that list.

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

### D17: Gender Special Attention rule (Steven's request, 2026-09-07)

Steven asked for a column before Gender Identification that is TRUE when Final Forms and chosen gender disagree, or when the requested pronouns do not go with the traditional gender. Made precise as: (a) Final Forms Gender mapped to Gx/Bx and Signup Gender both non-blank and different; or (b) with Gender Identification Bx, the lowercased Pronouns list has anything left after stripping he, him, and separators; or (c) with Gender Identification Gx, anything left after stripping she and her. "Anything left" rather than a fixed list of they/them so custom pronouns flag too. Blank pronouns never flag. Live data at the time: pronoun tokens are he, him, she, her, they, them, semicolon-joined.

### Follow-ups completed after Steven's go-ahead (2026-09-07)

- Deleted photo-mapper/ and its root README mention in its own commit. The gitignored service account key file under photo-mapper/backend/ was left on disk for Steven to remove or revoke.
- Purged the five named test rows from the portal's Signups sheet and their orphaned Extra Player Info rows; Signups, Roster, and Extra Player Info each hold 75 rows and every Extra Player Info Full Name resolves. Steven then confirmed Bob Larson (g94df), A E (qk4mx), and Big Loep (tt3t7) were fixtures too; deleted with their Extra Player Info rows, leaving 72 Players.

### Final verification after the column changes (2026-09-07, v3.16)

Steven ran Generate Fresh Roster on 3.16 after the reorder, the Gender Default Handling column, and the Boolean inversion. The comparison script (every column of every row against Signups, Final Forms, and Newsletter Subscribers, including the inverted Gender Default Handling and Media OK rules) reported 70 Players, header order exactly as requested, and 0 mismatches.

### Not verified

- Build Practice Roster and Build Game Roster Prep end-to-end (D14).
- The Extra Player Info dropdown rules were applied by the script but not read back (gog does not expose data validation). Two-minute check: open Extra Player Info, click any Team cell; the dropdown should offer Blue and Gold, and Returning and Include should offer TRUE and FALSE.
- Show Statistics, Find Emails Not Subscribed to Newsletter, and Caretakers Not Subscribed to Newsletter were rewritten but only syntax-checked; each is one menu click to confirm.

### D18: Array formulas replaced by per-row formulas (Steven's report, 2026-09-08)

Steven reported that filter views do not work on the array-formula Roster at all and that it cannot be sorted. Switched to the spring layout keyed by PlayerID: column A written as values, every other cell a per-row formula with row-relative sibling references, one row per Player, rewritten in full by Generate Fresh Roster. Run Diagnostics now checks the per-row shape and compares Roster PlayerIDs with Signups instead of looking for `=SORT(FILTER(`. Recorded as [ADR 0003](../adr/0003-roster-per-row-formulas.md) (status proposed until Steven confirms sorting and filter views behave). Pushed as 3.21 for live verification; not committed at Steven's request until he has checked it. Steven's first run showed every Final Forms column blank or FALSE and every newsletter status "not a member": a plain per-row formula implicitly intersects whole-column arguments like `TO_TEXT('Final Forms'!$A:$A)` with its own row. 3.22 wrapped each cell formula in `ARRAYFORMULA` as a quick fix. Steven asked why the fall 2025 plain formulas would not do; live tests showed the real cause is a type mismatch (Signups SPS Student ID is text, Final Forms StudentID is a number) and that XLOOKUP is already case-insensitive. 3.23 drops the wrapper: Final Forms lookups use `VALUE($B<row>)` with a text fallback, newsletter lookups drop `LOWER`, and Run Diagnostics checks for the plain `=IF($A2=""` shape.

### D19: Availability rows keyed by PlayerID with formula-derived player columns (Steven's report, 2026-09-14)

Steven found that Build Practice/Game Availability wrote Full Name, Grade, and Gender Identification as copied values next to the PlayerID, and hand-patched the live Practice Availability rows 2-70 to per-row Roster XLOOKUP formulas keyed by column A (PlayerID). The generator had also still put Full Name in column A and PlayerID in column B. Changed `seedAvailabilityRows_` so a new sheet is laid out PlayerID, Full Name, Grade, Gender Identification; PlayerID is the only typed value and the other three cells are `availabilityRosterFormula_` output (`=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))` and so on), with Roster column letters read from the live Roster header. Existing rows that have a PlayerID get those cells rewritten as formulas on every run, so older value-based sheets convert themselves. The prep sheets (Build Practice Roster, Build Game Roster Prep, both audiences) had hardcoded `'Practice Availability'!A:A` / `'Game Availability'!A:A` as their Full Name lookup range, which the PlayerID-first layout would have broken; `findAvailabilityColumns` now returns `fullNameColumn` resolved by header and every one of those formulas uses it. Apply Activation Status from a prep sheet likewise finds Game Availability's Full Name by header instead of reading column A. Verified offline: the formula the source generates matches all 69 hand-patched live rows exactly (Node harness against the gog export). Pushed as 3.28. Recorded as [ADR 0004](../adr/0004-availability-rows-playerid-plus-roster-formulas.md). Not yet run live: Build Practice Availability on 3.28 should report 0 rows added and 0 formulas repaired against the hand-patched sheet; Build Game Availability will lay out the fall sheet once Game Info has rows (the tab still holds spring headers, which the build appends after rather than replaces).

### D20: Printouts carry a hidden PlayerID column and join on it (Steven's request, 2026-09-14)

Asked why the prep sheets still joined to the availability sheets by Full Name after D19, Steven decided they should have a PlayerID column. Build Practice Roster (create and update), the coach game roster prep, and the parent game roster now seed a `PlayerID` column from the Roster (`seedPrintoutPlayerRows`, respecting Include In Generated Rosters), write Full Name, Team, Gender, Grade, and every availability, activation, and note cell as per-row lookups keyed by it (`fillLookupColumn` over `playerIdLookupFormula`, the shared house-shape helper that Availability.gs now uses too), and hide the column after formatting so printouts look the same. `findAvailabilityColumns` returns `playerIdColumn`; `availabilityJoin` falls back to a Full Name join with a warning when an availability sheet predates 3.28. Apply Activation Status joins by PlayerID when both sheets have one. Fixed in passing: the practice roster update path labeled its base columns Grade, Gender, Team while the data was written Team, Gender, Grade; both now come from CONFIG. Verified offline with a Node harness over fake sheets (row seeding with the Include filter, every formula on a practice roster with next-game columns, the Full Name fallback against an old availability sheet, the coach game prep layout and formulas, and the availability seeding through the shared helper): all assertions pass. Pushed as 3.29. Not yet run live: Build Practice Roster for a fall practice date, and Build Game Roster Prep once Game Info and Game Availability have fall rows. ADR 0004 amended.

### D21: Printout PlayerID moves to column A (Steven's choice, 2026-09-15)

Steven noticed Build Practice Roster adds the PlayerID column and hides it, and offered two options: key the short-lived practice roster by name after all, or put PlayerID in column A and hide it. Took column A: it matches the Roster and the availability sheets (PlayerID is column A on every PlayerID-keyed sheet) and keeps duplicate names from colliding. `rosterPrintoutBaseColumns` is now PlayerID 1, # 2, Full Name 3, Team 4, Gender 5, Grade 6; the game prep layout and the parent roster take their fixed indices from it; the # formula and group borders use the # column's index instead of assuming column A; the coach game prep summary tables sit under Full Name instead of B/C/D. Harness re-run green. Pushed as 3.30.

### D22: No Activation Status this season (Steven's report, 2026-09-15)

Steven reported that Build Practice Roster added a next-game Activation Status column, which fall 2026 does not use. `CONFIG.gameRosterPrep.hasActivationStatus` already governed the coach game roster prep; it now also governs Build Game Availability (no `$date Activation Status` column per game), Build Practice Roster's next-game block (`practiceRosterNextGameColumns` decides whether the block is two or three columns), and whether the Apply Activation Status menu item is shown. Set to false for fall 2026. The live Game Availability tab had already been built with `9/26 Activation Status` and `10/3 Activation Status` columns (and still carries a stray spring `3/7 Availability`); builds never delete columns, so those are for Steven to remove by hand. Pushed as 3.31.

### D23: Fall 2026 teams and Team sort order (Steven's request, 2026-09-15)

Steven confirmed the teams: Blue, Gold, Silver, TBD (for Players not yet placed; goes away after assignments), Practice Squad, and asked that the build roster commands sort by that order; on follow-up, only the practice roster and the coach game prep, not the parent roster. Added `CONFIG.teams` in that order. Build Practice Roster and the coach Build Game Roster Prep sort by Team rank first (a temporary numeric key column, since Range.sort only orders by cell values), blank or unlisted Team last; the coach game prep's # column now also resets on a Team change. `CONFIG.gameRosterPrep.hasTeam` is now true so the coach game prep has a Team column to sort on. The Extra Player Info Team dropdown seeds from `CONFIG.teams` (only when no rule exists yet; the live dropdown was already edited by hand and is left alone). Live Extra Player Info at the time: 12 Blue, 9 Gold, 3 Silver, 19 TBD, 4 Practice Squad, 32 blank. Pushed as 3.32.

### Fix: menu missing since 3.31 (Steven's report, 2026-09-15)

Steven refreshed the sheet and saw no Madison Ultimate menu. Cause: the 3.31 change that made Apply Activation Status conditional split the createMenu chain but never assigned it to the `menu` variable it then used, so onOpen threw a ReferenceError. `node --check` cannot see that. Fixed in 3.33; the Node harness now builds the menu against a stub UI with the switch both off and on, so a broken onOpen fails offline.
