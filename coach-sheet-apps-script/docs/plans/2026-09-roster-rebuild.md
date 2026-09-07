# Fall 2026 roster rebuild: implementation plan

Status: draft for review (2026-09-07). Decisions behind this plan: [ADR 0001](../adr/0001-roster-keyed-by-signups-playerid.md). Vocabulary: [CONTEXT.md](../../CONTEXT.md). Target workbook: "2026 Fall Coach Sheets" (`1kMg6OgT8gzZyZChS6MtobLZ6GjmcrieV_VRUTTiuinE`), bound script per `.clasp.json`.

## Outcome

Generate Fresh Roster rewrites the 📋 Roster tab as one header row plus one array formula per column in row 2, all derived from the 2026 Fall Signups tab (an IMPORTRANGE mirror of the portal's Signups sheet), Extra Player Info, Final Forms, and Newsletter Subscribers. Nothing is authored in the Roster. Two new menu items support it: Sync Extra Player Info (adds missing PlayerIDs for coaches to annotate) and Analyze Signups (the report of Players whose Sources disagree or are incomplete). Everything Additional Info is deleted.

## Contract shared by every work package

### Roster column list (order is code order)

Source abbreviations: S = Signups by PlayerID, E = Extra Player Info by PlayerID, F = Final Forms by SPS Student ID, N = Newsletter Subscribers by email, D = derived from other Roster columns.

| # | Header | Type | Source | Rule |
|---|---|---|---|---|
| 1 | PlayerID | String | S | Key: every non-empty PlayerID in Signups, sorted by Last Name then Preferred First Name |
| 2 | SPS Student ID | String | S | passthrough |
| 3 | Preferred First Name | String | S | passthrough |
| 4 | Legal First Name | String | S | passthrough |
| 5 | Last Name | String | S | passthrough |
| 6 | Full Name | String | D | TRIM(Preferred First Name & " " & Last Name) |
| 7 | Grade | Number | F, S fallback | Final Forms Grade when the SPS Student ID resolves, else Signups Grade |
| 8 | Elementary School | String | S | passthrough |
| 9 | Final Forms Gender | Enum | F | Final Forms Gender (Male/Female) |
| 10 | Signup Gender | Enum | S | Signups Gender Identification collapsed: Girl or Gx to "Gx", Boy or Bx to "Bx", else blank |
| 11 | Gender Identification | Enum | D | Signup Gender if set, else Final Forms Gender mapped Female to Gx, Male to Bx, else blank |
| 12 | Pronouns | String | S | passthrough (semicolon-joined) |
| 13 | Team | Enum | E | passthrough |
| 14 | Returning | Boolean | E | passthrough |
| 15 | Include In Generated Rosters | Boolean | D | Extra Player Info value when not blank, else Profile Complete? |
| 16 | Profile Complete? | Boolean | D | Grade, Date of Birth, and Caretaker 1 Email all non-empty |
| 17 | Are All Forms Parent Signed | Boolean | F | Final Forms P equals TRUE |
| 18 | Are All Forms Student Signed | Boolean | F | Final Forms Q equals TRUE |
| 19 | Physical Cleared | Boolean | F | Final Forms Physical Clearance equals "Cleared" |
| 20 | Final Forms Cleared? | Boolean | D | AND of 17, 18, 19 |
| 21 | Date of Birth | Date | S | DATEVALUE of the ISO string, date number format |
| 22 | Student SPS Email | Email | S | passthrough |
| 23 | Student Personal Email | Email | S | passthrough, blanked if the domain is seattleschools.org |
| 24 | Student Newsletter Status | Enum | N | status for Student Personal Email, "not a member" when absent, blank when no email |
| 25 | Caretaker 1 Name | String | S | passthrough |
| 26 | Caretaker 1 Email | Email | S | passthrough |
| 27 | Caretaker 1 Newsletter Status | Enum | N | as 24 for Caretaker 1 Email |
| 28 | Caretaker 2 Name | String | S | passthrough |
| 29 | Caretaker 2 Email | Email | S | passthrough |
| 30 | Caretaker 2 Newsletter Status | Enum | N | as 24 for Caretaker 2 Email |
| 31 | Player Allergies | String | S | Signups "Allergies" |
| 32 | Competing Sports and Activities | String | S | passthrough |
| 33 | Jersey Size | String | S | passthrough |
| 34 | Playing Experience | String | S | passthrough |
| 35 | Player hopes for the season | String | S | Signups "Hopes" |
| 36 | Other Player Info | String | S | Signups "Other Info" |
| 37 | Media Opt-Out | Boolean | S | LOWER(value) = "true" |
| 38 | Photo Drive File ID | String | S | passthrough |
| 39 | Photo Link | Hyperlink | D | HYPERLINK("https://drive.google.com/uc?id=" & file id, "photo") when file id present |

Each header cell gets a Sheets note: `Type: ... / Source: ... / <explanation>`.

### Signups headers the formulas reference (resolved by name at generation time)

PlayerID, SPS Student ID, Preferred First Name, Legal First Name, Last Name, Grade, Elementary School, Gender Identification, Pronouns, Date of Birth, Student SPS Email, Student Personal Email, Caretaker 1 Name, Caretaker 1 Email, Caretaker 2 Name, Caretaker 2 Email, Allergies, Competing Sports and Activities, Jersey Size, Playing Experience, Hopes, Other Info, Media Opt-Out, Photo Drive File ID.

### Final Forms columns (fixed positions, validated by finalforms-export)

StudentID A, Gender U, Grade W, Parent Signed P, Student Signed Q, Physical Clearance AB.

### Extra Player Info tab

Header row: PlayerID, Full Name, Team, Returning, Include In Generated Rosters. Full Name is a per-row formula (XLOOKUP of PlayerID into the Roster). Team has a dropdown seeded Blue, Gold (edit in the sheet after tryouts). Returning and Include In Generated Rosters are both dropdowns with TRUE and FALSE, blank allowed. No checkboxes: a checkbox cannot be blank, and for Include In Generated Rosters blank is what means "use Profile Complete"; Returning uses the same control for consistency.

### Formula patterns

All row 2 formulas are `ARRAYFORMULA` over open ranges guarded by `IF($A2:$A="","",...)` so rows past the last PlayerID stay blank. Column letters below are placeholders resolved at generation time.

- Key (column A): `=SORT(FILTER(S!A2:A, S!A2:A<>""), FILTER(S!Last2:Last, S!A2:A<>""), TRUE, FILTER(S!Pref2:Pref, S!A2:A<>""), TRUE)`
- Signups lookup: `=ARRAYFORMULA(IF($A2:$A="","",IFERROR(XLOOKUP($A2:$A, S!$A:$A, S!$X:$X),"")))`
- Final Forms lookup: `=ARRAYFORMULA(IF($A2:$A="","",IFERROR(XLOOKUP(TO_TEXT($B2:$B), TO_TEXT(F!$A:$A), F!$X:$X),"")))`. Coerce both ID sides with TO_TEXT: the CSV import may store StudentID as a number while Signups stores text.
- Final Forms booleans: `UPPER(TO_TEXT(value))="TRUE"` so the result is a real boolean whether the import stored text or boolean.
- Newsletter status: `IF(email="","",IFERROR(VLOOKUP(LOWER(email), ARRAYFORMULA(LOWER(N!$A$2:$A)) paired with N!$B$2:$B ...),"not a member"))`. Implement with XLOOKUP over LOWER of both sides.
- Derived columns reference sibling Roster columns by resolved letter (for example Final Forms Cleared? is `AND($Q2:$Q, $R2:$R, $S2:$S)` in array form).

Generation ensures the sheet has at least 300 rows, freezes row 1, bolds the header, applies a date format to Date of Birth, and writes header notes. Existing conditional formatting on the Roster is left untouched.

## Work packages

### WP1: Roster generation core (Code.gs)

- Replace `buildRosterSheet`, `clearRosterDataInternal`, `clearRosterData`, `validateRosterMetadata`, `populateStudentIdsFromFinalForms`, `updateRosterStudentIdsInternal`, and the placeholder-replacement block with: a `ROSTER_COLUMNS` definition list (name, type, source, note, and a formula builder that receives resolved column letters), a `resolveSignupsColumns()` helper that maps the Signups header names above to letters and throws naming any missing header, a `resolveRosterColumns()` helper for sibling references, and `generateRoster()` that creates the tab if missing, clears contents and notes, writes header, notes, and row 2 formulas, and applies the formatting above.
- Delete `FIRST_DATA_ROW`. Keep `ROSTER_HEADER_ROW = 1` and set `ROSTER_FIRST_DATA_ROW = 2` as two explicit constants. Fix every remaining `FIRST_DATA_ROW` use in Code.gs (statistics, missing emails, pending caretakers) to the roster constants.
- CONFIG: remove `additionalInfo`; add `signups.sheetName = '2026 Fall Signups'` and `extraPlayerInfo.sheetName = 'Extra Player Info'`; rename column constants to the sheet names (`playerId: 'PlayerID'`, `spsStudentId`, `preferredFirstName`, `caretaker1Name`, `caretaker1Email`, `caretaker2Name`, `caretaker2Email`, `profileComplete`, drop `studentId`, `parent*`); keep `fullName`, `lastName`, `grade`, `team`, `genderIdentification`, `includeInGeneratedRosters`.
- Menu: remove Clear Roster Data and Analyze Additional Info Responses; add Sync Extra Player Info and Analyze Signups; rename "Parents Not Members of Mailing List" to "Caretakers Not Subscribed to Newsletter".
- `refreshAllData` message no longer mentions Additional Info.
- Rewrite `showStatistics` on the new columns (total, Profile Complete, Final Forms Cleared, each signed/cleared flag, newsletter regular counts for caretakers, grade distribution) and drop the undefined `additionalInfoCol` reference. Read the sheet in one `getValues` call instead of per-cell.
- `findMissingEmails`: read Full Name for the report instead of columns 1 and 2. `findPendingParents`: use Caretaker 1/2 Name and Email columns, one name field.

### WP2: Extra Player Info sync (new ExtraPlayerInfo.gs)

- `syncExtraPlayerInfo()`: create the tab with the header above if missing; apply validations (Team dropdown Blue/Gold with blank allowed, Returning and Include dropdowns TRUE/FALSE with blank allowed); read PlayerIDs from Signups, append a row for each not already present with the Full Name formula; never delete or reorder; alert with counts added and total.
- Diagnostics check that the tab and its five headers exist.

### WP3: Analyze Signups (new AnalyzeSignups.gs, delete AdditionalInfoAnalysis.gs)

- `analyzeSignups()` writes or replaces an "Analyze Signups" sheet with a timestamp and five sections, each a small table with PlayerID, Full Name, and the relevant detail: (1) signups with no SPS Student ID; (2) Final Forms students whose StudentID is on no signup row (name, grade, StudentID); (3) signups whose SPS Student ID is missing from the Final Forms tab; (4) suspected duplicates: rows sharing normalized last name and birthdate, or sharing an SPS Student ID; (5) signups not Profile Complete, listing which of Grade, Date of Birth, Caretaker 1 Email are missing. Reads Signups and Final Forms tabs directly, not the Roster, so it works before the Roster is generated.
- Summary alert with the five counts.

### WP4: Consumers, diagnostics, cleanup

- `BuildEmailList.gs`: caretaker email constants; header text "Caretaker" instead of "Parent".
- `Diagnostics.gs`: remove the Additional Info spreadsheet check; add checks that the Signups tab exists and its A1 holds an IMPORTRANGE that resolved (A2 non-error), that every referenced Signups header exists, that the Roster header contains every `ROSTER_COLUMNS` name and every `CONFIG.columns` value, that Roster A2 begins with `=SORT(FILTER(`, and the Extra Player Info check from WP2.
- `FullNameDiff.gs`, `SheetBuilderUtils.gs`: no logic change; confirm they use `ROSTER_HEADER_ROW`/`ROSTER_FIRST_DATA_ROW` only.
- `.clasp.json` filePushOrder: drop AdditionalInfoAnalysis.gs, add ExtraPlayerInfo.gs and AnalyzeSignups.gs.
- `grep -rn "Additional Info\|additionalInfo\|FIRST_DATA_ROW\b\|parent1Email\|StudentID'" *.gs` returns nothing when done (Final Forms header literal excepted).

### WP5: Docs

- README: rewrite Spreadsheet Structure (Required Sheets adds 2026 Fall Signups and Extra Player Info, drops Additional Info; Roster layout section becomes "one header row, formulas in row 2, sort with filter views only"; delete the metadata-rows and Column Source Types sections), Data Sources (add Signups, remove Additional Info), Menu Functions (new items, removed items), Key Concepts (PlayerID as primary key, Full Name downstream key; delete Student ID and Additional Info sections), File Structure, Troubleshooting.
- DESIGN.md: rewrite the data sources, roster structure, join logic, and formula pattern sections to match the ADR; drop the Fall 2025 column spec.
- Portal repo `docs/fall-2026/signup-plan.md`, "Roster population" paragraph: mark superseded by this ADR (formula-keyed Roster, Analyze Signups carries the two reports), keeping the original text struck or quoted for history.
- `steps I took for setting up the fall 2026 sheet.md`: add the manual steps below.

### WP6: Deploy and verify (single live target, sequential)

1. Bump `SCRIPT_VERSION`, `clasp push`.
2. Run Diagnostics; fix anything red.
3. Generate Fresh Roster with the two TestCleared rows still present; check every column against the Signups values, the Final Forms row for a real StudentID pasted into one test row's SPS Student ID cell (then removed), and the newsletter statuses.
4. Sync Extra Player Info, set Team and Include values on a test row, confirm they flow through; confirm blank Include resolves to Profile Complete.
5. Analyze Signups; confirm the five sections populate sensibly.
6. Build Practice Roster and Build Game Roster Prep against the generated Roster to confirm Full Name keyed builders still work.
7. Steven purges the test rows in the portal's Signups sheet and re-points the Roster Pivot source range at row 1.
8. Commit per work package with Conventional Commits.

## Out of scope, noted

- photo-mapper reads `Full Name` and `StudentID` from a roster. The portal now ties each Player Photo to a PlayerID, so the tool has no remaining use case. Follow-up after this plan lands: delete `photo-mapper/` and its README mention in the repo root README, in its own commit.
- Team values beyond Blue and Gold, and flipping `CONFIG.gameRosterPrep.hasTeam` to true, wait for tryouts.
- Re-keying availability and generated sheets to PlayerID (deliberately not done, per ADR).

## Amendments

- 2026-09-07 (after the plan landed, Steven's request): Elementary School moved before Grade; Pronouns moved before Gender Identification; new derived Boolean column **Gender Special Attention** inserted between Pronouns and Gender Identification. Rule: TRUE when Final Forms Gender and Signup Gender both exist and disagree, or when Pronouns include anything outside he/him for a Bx or she/her for a Gx. Column order is now: PlayerID, SPS Student ID, Preferred First Name, Legal First Name, Last Name, Full Name, Elementary School, Grade, Final Forms Gender, Signup Gender, Pronouns, Gender Special Attention, Gender Identification, Team, then the rest of the table unchanged (40 columns).
- 2026-09-07, later (Steven's request): full reorder. Order is now PlayerID, SPS Student ID, Preferred First Name, Legal First Name, Last Name, Full Name, Elementary School, Date of Birth, Player Allergies, Competing Sports and Activities, Jersey Size, Playing Experience, Player hopes for the season, Other Player Info, Grade, Final Forms Gender, Signup Gender, Pronouns, Gender Special Attention, Gender Identification, Team, Returning, Include In Generated Rosters, Profile Complete?, Are All Forms Parent Signed, Are All Forms Student Signed, Physical Cleared, Final Forms Cleared?, Caretaker 1 Name, Caretaker 1 Email, Caretaker 1 Newsletter Status, Caretaker 2 Name, Caretaker 2 Email, Caretaker 2 Newsletter Status, Student SPS Email, Student Personal Email, Student Newsletter Status, Media Opt-Out, Photo Drive File ID, Photo Link. Rules unchanged; `ROSTER_COLUMNS` in Code.gs is the authority on order.
