# Coach Sheet Apps Script: design

The coach-facing Google Sheets workbook for a Madison Ultimate season and the Apps Script menu bound to it. It joins data mastered elsewhere (the family portal's Signups sheet, the district's Final Forms export, Buttondown) with a small set of coach-authored facts, and derives the roster and every printout from that join. Vocabulary is in [CONTEXT.md](./CONTEXT.md); the reasoning behind the roster design is in [ADR 0001](./docs/adr/0001-roster-keyed-by-signups-playerid.md) and [ADR 0003](./docs/adr/0003-roster-per-row-formulas.md); the plan that implemented it is in [docs/plans/2026-09-roster-rebuild.md](./docs/plans/2026-09-roster-rebuild.md).

**Seasons:** Fall 2025 (initial development, Final Forms keyed), Spring 2026, Fall 2026 (Signups keyed, this document).

## System architecture

### The workbook

One Google Sheet per season (fall 2026: "2026 Fall Coach Sheets"), duplicated from the prior season, with this script bound to it (see the README for how to re-point `.clasp.json` and the per-season `CONFIG` values). The tabs the script depends on are listed in the README's Required Sheets table and checked by Run Diagnostics.

### Sources

Every authored fact lives in exactly one of four Sources. Nothing is authored in the Roster.

| Source | Mastered by | Tab | Key | How it arrives |
|---|---|---|---|---|
| Signups | Families, through the portal | `2026 Fall Signups` | PlayerID | IMPORTRANGE of the portal's Signups sheet; refreshes on its own; read-only here |
| Extra Player Info | Coaches | `Extra Player Info` | PlayerID | Typed into the tab; rows added by Sync Extra Player Info |
| Final Forms | The district | `Final Forms` | SPS Student ID | Update Final Forms imports the newest CSV from the Drive folder the finalforms-export automation writes to |
| Newsletter Subscribers | Buttondown | `Newsletter Subscribers` | Email | Update Newsletter Subscribers pages through the Buttondown API |

Signups carries the SPS Student ID once the portal's Final Forms Join succeeds, which is what lets the Roster reach Final Forms by PlayerID alone.

## Roster structure

The 📋 Roster is one header row plus one row per Player, each row self-contained:

- **Row 1** is the header. Each header cell carries a note `Type: ... / Source: ... / <rule>` so the sheet explains itself without metadata rows.
- **Column A** holds the PlayerID as a plain value, one row for every non-empty PlayerID in Signups, written in Last Name then Preferred First Name order.
- **Every other cell** is a plain per-row formula guarded by `IF($A<row>="","",...)`, in the same shape as the fall 2025 sheet. No scalar function is applied to a whole-column range (a plain formula implicitly intersects such a range with its own row and every lookup misses); ranges go straight into XLOOKUP and only the single lookup key is converted. It joins by the PlayerID on its row (Signups, Extra Player Info), by SPS Student ID (Final Forms), or by email (Newsletter Subscribers), or derives from sibling cells on the same row. Sibling references are row-relative (`$F2` on row 2), so in-place sorting moves a whole row and keeps it consistent.

Generate Fresh Roster reads every Signups row, sorts the PlayerIDs, clears the Roster's contents, notes, and data validations, writes the header, notes, PlayerID values, and one formula row per Player, freezes row 1, bolds the header, applies a date format to Date of Birth, and finishes by running Format Spruce Up's formatting worker on the Roster (alternating row banding, a data filter, vertical centering, frozen row 1 and column A) so it never needs a separate manual pass. Existing conditional formatting and filter views are left alone. `ROSTER_HEADER_ROW = 1` and `ROSTER_FIRST_DATA_ROW = 2` are the only layout constants readers use.

Fall 2026 first shipped this as one `ARRAYFORMULA` per column in row 2 with a `SORT(FILTER(...))` key formula in A2 (the first form of ADR 0001). That was abandoned within the day because filter views and sorting do not work on array-formula output; [ADR 0003](./docs/adr/0003-roster-per-row-formulas.md) records the switch.

### Column definitions

`ROSTER_COLUMNS` in `Code.gs` is the single source of truth: an ordered list of `{ name, type, source, note, formula }` where `formula` is a builder that receives resolved column letters and a row number and returns that row's formula (`null` for the PlayerID key column, which holds values). The 43 columns and their rules are tabulated in the plan (with its amendments section). Adding or reordering a column means editing that list and running Generate Fresh Roster.

`CONFIG.columns` holds the header names other files look up (`Full Name`, `Team`, `Gender Identification`, `Grade`, `Include In Generated Rosters`, the Caretaker email columns, and so on); every value there must be a `ROSTER_COLUMNS` name, and Run Diagnostics checks the live header row for all of them.

### Consequences

- The Roster can be sorted and filtered freely (filter views, the basic filter, Data > Sort range); every row is self-contained.
- The Roster is a snapshot of Signups' PlayerIDs: a new or removed Signup reaches it only when Generate Fresh Roster runs again (which also resets the row order). Run Diagnostics compares the Roster's PlayerIDs with Signups and fails when they differ, and checks the data rows still have the per-row formula shape.
- A signup with no SPS Student ID shows blank or FALSE Final Forms columns rather than being hidden; a Final Forms student with no signup does not appear. Analyze Signups surfaces both.
- Downstream sheets (availability, practice roster, game roster prep, email lists) stay keyed by Full Name; only the Roster is keyed by PlayerID.

## Join logic

| Target | Key | Formula shape |
|---|---|---|
| Signups | PlayerID | `IFERROR(XLOOKUP($A2, Signups!$A:$A, Signups!$X:$X), "")` on row 2 (and `$A3` on row 3, and so on) with `X` resolved from the Signups header name at generation time |
| Extra Player Info | PlayerID | Same shape against the Extra Player Info tab; column letters come from `EXTRA_PLAYER_INFO_HEADERS` |
| Final Forms | SPS Student ID | `IF($B2="", <missing>, IFERROR(XLOOKUP(VALUE($B2), 'Final Forms'!$A:$A, ...), IFERROR(XLOOKUP($B2, 'Final Forms'!$A:$A), 'Final Forms'!$X:$X), ""))`. Both sides are coerced with `TO_TEXT` because the CSV import stores StudentID as a number while Signups stores text; the inner guard stops a blank ID from matching a blank export row. `<missing>` is `""` for text and `FALSE` for the signature and clearance flags. Final Forms columns are fixed positions (StudentID A, Parent Signed P, Student Signed Q, Gender U, Grade W, Physical Clearance AB), validated by finalforms-export. |
| Newsletter Subscribers | Email | `IF(email="", "", IFERROR(XLOOKUP(LOWER(email), LOWER(Subscribers!$A$2:$A), Subscribers!$B$2:$B), "not a member"))` |
| Derived | Sibling Roster columns | Referenced by resolved letter, for example Full Name is `TRIM($C2:$C&" "&$E2:$E)` and Final Forms Cleared? is `($Q2:$Q=TRUE)*($R2:$R=TRUE)*($S2:$S=TRUE)=1` |

Rules worth knowing:

- **Grade** prefers Final Forms and falls back to Signups.
- **Signup Gender** collapses the family's Gender Identification to Gx or Bx; **Gender Identification** is Signup Gender when set, else Final Forms Gender mapped Female to Gx and Male to Bx. Generated Rosters print Gender Identification.
- Every Roster Boolean reads TRUE for "all is well" and FALSE for "needs attention". **Gender Default Handling** is FALSE when Final Forms Gender and Signup Gender disagree, or when Pronouns include anything outside he/him for a Bx or she/her for a Gx (the pronoun list is lowercased, the expected pronouns and separators are stripped, and anything left over flags). A prompt for a coach to check in, not a verdict. **Media OK** is FALSE when the family declared a Media Opt-Out.
- **Profile Complete?** is a passthrough of the portal's `Profile Complete` column (Player Info, Caretaker Info, and Photo Upload all done; portal ADR 0006), coerced to a boolean. **Include In Generated Rosters** is the Extra Player Info value when set, else TRUE (coach sheet ADR 0002).
- **Final Forms Cleared?** is TRUE only when all forms are parent signed, all forms are student signed, and the physical is cleared; a missing SPS Student ID makes it FALSE.
- **Student Personal Email** is blanked when its domain is seattleschools.org.
- **Date of Birth** accepts the ISO text Signups stores (`DATEVALUE`) or a real date.
- **Photo Link** is a `HYPERLINK` to the Drive file when a Photo Drive File ID is present.

## Extra Player Info

Header: `PlayerID, Full Name, Team, Returning, Number of Past Seasons, Signup Playing Experience, Tryout Group, Signup Grade, Include In Generated Rosters`. Full Name and Signup Playing Experience are per-row XLOOKUPs (into the Roster and Signups respectively, by PlayerID), not authored; Signup Playing Experience exists only so a coach filling in Number of Past Seasons can see the source text right next to it. Team is a dropdown seeded Blue and Gold (seeded once; coaches edit the list after tryouts and the next sync leaves it alone). Returning and Include In Generated Rosters are TRUE/FALSE dropdowns with blank allowed; no checkboxes, because a checkbox cannot be blank and blank Include is what means "included". Number of Past Seasons is a non-negative number with blank allowed, meaning "not yet reviewed" rather than zero. Tryout Group and Signup Grade are coach-typed during tryouts with no validation yet; Tryout Group also passes through to the Roster.

Sync Extra Player Info appends a row for each Signups PlayerID not already present (in Roster order) and never deletes or reorders. An existing tab with a different header is rewritten only while it has no data rows; otherwise the mismatch is reported as an error.

## Analyze Signups

Reads Signups and Final Forms directly (not the Roster, so it works before the Roster exists) and writes or replaces the "Analyze Signups" sheet with a timestamp and six sections: signups with no SPS Student ID, Final Forms students not yet seeded or joined, signups whose SPS Student ID is not in Final Forms, suspected duplicates (same normalized last name and birthdate, or same SPS Student ID), signups not Profile Complete (as the portal wrote it, marked Seeded Signup or family-created), and Seeded Signups the family has not finished. Name normalization follows the portal's rules (trim, lowercase, strip whitespace and apostrophes, fold accents, keep hyphens). It only reports; the portal's Seed Signups from Final Forms does the joining and seeding.

## Technical implementation

### Files

| File | Purpose |
|---|---|
| `Code.gs` | `CONFIG`, `SIGNUPS_HEADERS`, `ROSTER_COLUMNS`, Generate Fresh Roster, menu, Final Forms import, statistics, newsletter reports |
| `Diagnostics.gs` | Run Diagnostics setup checks |
| `ExtraPlayerInfo.gs` | Sync Extra Player Info |
| `AnalyzeSignups.gs` | Analyze Signups report |
| `NewsletterSubscribers.gs` | Buttondown subscriber import |
| `Availability.gs`, `ManagedConditionalFormatting.gs` | Practice and game availability sheets and their shared formatting rules |
| `BuildPracticeRoster.gs`, `BuildGameRosterPrepSheet.gs`, `SheetBuilder.gs`, `SheetBuilderUtils.gs` | Generated Rosters and the custom sheet builder, keyed by Full Name |
| `BuildEmailList.gs` | Caretaker email lists from Full Names |
| `ApplyActivationStatusFromRoster.gs`, `ConvertToAttendance.gs` | Game-day activation and attendance helpers |
| `CreatePracticeCalendarEvents.gs`, `CreateGameCalendarEvents.gs`, `ExportGameInfoToMarkdown.gs` | Calendar sync and exports |
| `FormatSpruceUp.gs`, `DeleteEmptyRowsColumns.gs`, `OrganizeSheets.gs`, `FullNameDiff.gs` | Utilities |

### Deployment

`clasp push` from `coach-sheet-apps-script/`, after bumping `SCRIPT_VERSION` in `Code.gs` (the menu title shows the version, which is how you confirm the sheet picked up a push). Conventional Commits for every change. See the README for new-season setup and Script Properties.

### Testing

There is no Apps Script test runner in this repo. The pure seam is `resolveSignupsColumns`, `resolveRosterColumns`, `buildRosterFormulas`, and `analyzeSignupsData`: they take plain values and return strings or plain objects, so a Node harness can load `Code.gs` and `AnalyzeSignups.gs` with stubbed globals and assert exact formulas and report rows. After a push, Run Diagnostics, Generate Fresh Roster, Sync Extra Player Info, and Analyze Signups from the menu and read the tabs back (for example with `gog sheets get`) to compare against the Source tabs.

## Key decisions

- **Why Google Sheets and Apps Script?** Coach preference, easy sharing, no infrastructure.
- **Why is the Roster keyed by Signups PlayerID?** Signups is the only Source that has every Player, and PlayerID is permanent. See ADR 0001 for the options rejected (Final Forms as key, an append-only value-writer, per-row formulas).
- **Why formulas rather than written values?** A formula cannot clobber anything and stays right as its Source changes. Only the PlayerID key column is written as values, because a formula-generated key list cannot be sorted or filtered in the sheet UI (ADR 0003); the price is that new or removed Signups need a Generate Fresh Roster run, and Run Diagnostics says when.
- **Why header notes rather than metadata rows?** The column list in code is the source of truth; notes carry the same explanation without shifting the data down five rows.
- **Why is Full Name still the downstream key?** Availability and Generated Rosters are read and printed by humans; re-keying them to PlayerID was deliberately not done.

## Out of scope, noted

- The old `photo-mapper/` tool (removed 2026-09-07) read Full Name and StudentID from a roster; the portal now ties each Player Photo to a PlayerID.
- Team values beyond Blue and Gold, and `CONFIG.gameRosterPrep.hasTeam`, wait for tryouts.
