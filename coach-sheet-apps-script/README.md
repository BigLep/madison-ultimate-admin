# Coach Sheet Apps Script

Google Apps Script tools for managing the Madison Middle School Ultimate Frisbee team roster spreadsheet.

**Season note:** This README and the `main` branch apply to the **Fall 2026** season. Past seasons get their own dated branch.

## Purpose

This script provides a custom menu in Google Sheets ("🥏 Madison Ultimate") that automates:
- Generating the 📋 Roster as a formula-only view of every Player in Signups, joined to Final Forms, Extra Player Info, and Newsletter Subscribers ([ADR 0001](./docs/adr/0001-roster-keyed-by-signups-playerid.md), [ADR 0003](./docs/adr/0003-roster-per-row-formulas.md))
- Importing Final Forms registration data and the Buttondown Newsletter subscriber list
- Keeping the coach-authored Extra Player Info tab in sync with Signups
- Building practice and game rosters with availability
- Generating email lists for Caretaker communication
- Reporting on Signups whose Sources disagree or are incomplete (Analyze Signups)

Vocabulary (Player, PlayerID, Source, Signups, Extra Player Info, Profile Complete, and so on) is defined in [CONTEXT.md](./CONTEXT.md); code, headers, and menu text use those terms.

## Prerequisites

- [Node.js](https://nodejs.org/) (for clasp CLI)
- [clasp](https://github.com/google/clasp) - Google Apps Script CLI
- Access to the team's Google Sheet with Editor permissions

## Installation

1. Install clasp globally:
   ```bash
   npm install -g @google/clasp
   ```

2. Login to clasp:
   ```bash
   clasp login
   ```

3. The `.clasp.json` file already points to the correct Apps Script project.

## Deployment

1. **Increment `SCRIPT_VERSION`** in `Code.gs` **before every push**, not after; this is a standing rule, not optional:
   ```javascript
   const SCRIPT_VERSION = '3.4';  // Increment the x in 3.x for each release
   ```

2. **Push changes**:
   ```bash
   clasp push
   ```

3. **Refresh the Google Sheet** and use the 🥏 Madison Ultimate menu.

4. **Run 🩺 Run Diagnostics** from the menu to confirm nothing broke.

### Deploy to a new season's spreadsheet

To bind this script to a **new** spreadsheet, the usual path is duplicating the prior season's spreadsheet (File → Make a copy in Drive), which also duplicates its bound Apps Script project as-is (same code, same internal project title, but a **different Script ID**). If instead you start from a blank spreadsheet, Extensions → Apps Script there creates a fresh empty project bound to it.

1. Open the new season's spreadsheet in Drive.
2. **Extensions → Apps Script**. This opens the script editor for the project already bound to that sheet (via duplication) or creates a new one (via a blank sheet).
3. In the script editor, open **Project settings** (gear icon) and copy the **Script ID** from the URL (`https://script.google.com/u/0/home/projects/<scriptId>/edit`).
   - **There is no API or CLI shortcut for this step.** `gog`'s Apps Script commands (and the underlying Apps Script/Drive APIs) all require a Script ID you already have; none of them can look up "the script bound to spreadsheet X." The Apps Script editor UI is the only way to discover a newly-duplicated bound script's ID.
   - The duplicated project's internal title stays whatever the source project was called (e.g. still "2026 Spring Coach Sheet Admin" after duplicating for fall); consider renaming it in Project settings for clarity, though this is cosmetic and doesn't affect deployment.
4. In this repo, update `coach-sheet-apps-script/.clasp.json`: set `"scriptId"` to that Script ID (leave `rootDir` and `filePushOrder` as-is).
5. **Update per-season values in `Code.gs`** before pushing:
   - `CONFIG.finalForms.folderId`: must point at **this season's** FinalForms exports Drive folder, and must match the `finalforms-export` GitHub Action's `DRIVE_FOLDER_ID` repo variable (`gh variable list -R BigLep/madison-ultimate-admin`). If these two drift apart, "Update Final Forms" silently imports the wrong season's data with no error.
   - `CONFIG.teams`: this season's Team values in display and sort order (fall 2026: Blue, Gold, Silver, TBD, Practice Squad). Seeds the Extra Player Info Team dropdown once, and orders Build Practice Roster and the coach Build Game Roster Prep; a blank or unlisted Team sorts last. TBD holds Players not yet placed and drops out of the list once assignments are done.
   - `CONFIG.gameRosterPrep.hasTeam`: `true` if this season uses teams and the roster has a Team column (fall 2026: true); the coach game roster prep then includes Team and sorts by it first. `false` to omit it from game roster prep sheets.
   - `CONFIG.gameRosterPrep.hasActivationStatus`: `true` if this season tracks a per-game Activation Status (Active/Inactive/TBD). When true, Build Game Availability adds a `$date Activation Status` column per game, the coach game roster prep includes it (sorted first), Build Practice Roster shows it for the next game, and the **Apply Activation Status** menu item is offered. When false (fall 2026), none of those appear. The build never deletes columns, so flip this before the first Build Game Availability of the season, or delete any `Activation Status` columns it already added by hand.
6. **Increment `SCRIPT_VERSION`**, then from `coach-sheet-apps-script/` run:
   ```bash
   clasp push
   ```
7. Back in the spreadsheet, refresh the page; the 🥏 Madison Ultimate menu should appear.
8. **Confirm required sheets exist.** Duplicating a spreadsheet copies every tab, including data sheets the script itself writes into (e.g. **Final Forms**) that look like leftover imported data and are easy to delete by mistake during season cleanup; don't delete a sheet during setup without first grepping `Code.gs` for its name to check whether the script depends on it.
9. **Set any required Script Properties** (see below); these are per-project, not copied by duplicating the spreadsheet, so they must be re-entered on every new bound script.
10. **Run 🩺 Run Diagnostics** from the menu as the final step. It checks that required sheets exist, configured Drive folders/spreadsheets are reachable, and required Script Properties are set; this is the "confirm everything is configured correctly" step for a new season, and the fast way to catch the kind of drift described above before a coach hits it mid-season.

## Script Properties (secrets)

Some integrations need credentials that must never be committed to this repo. Those are stored per-project in Apps Script's **Script Properties**, not in `Code.gs`:

1. Open the spreadsheet → **Extensions → Apps Script**.
2. **Project Settings** (gear icon) → **Script Properties** → **Add script property**.
3. Enter the property name and value, then read it in code with `PropertiesService.getScriptProperties().getProperty('NAME')`.

Script Properties are scoped to the Apps Script project (not the spreadsheet), so duplicating the spreadsheet for a new season does **not** carry them over; they must be re-added on the newly bound project each season. `🩺 Run Diagnostics` checks that required properties are set (and, for Buttondown, that the key actually works).

### Required properties

| Property | Used by | How to get it |
|----------|---------|----------------|
| `BUTTONDOWN_API_KEY` | 📬 Update Newsletter Subscribers | Buttondown → Settings → Programming → API Keys. Read access to subscribers is enough; do not paste it into `Code.gs` or this repo. |

## Spreadsheet Structure

### Required Sheets

The script expects these sheets to exist (created manually or via the menu):

| Sheet Name | Purpose |
|------------|---------|
| `📋 Roster` | Formula-only view of every Player; rewritten by Generate Fresh Roster |
| `2026 Fall Signups` | Read-only IMPORTRANGE mirror of the portal's Signups sheet (one row per Player, keyed by PlayerID). Its A1 holds the IMPORTRANGE; `CONFIG.signups.sheetName` names the tab, so rename both together each season. |
| `Extra Player Info` | Coach-authored per-player facts (Team, Returning, Include In Generated Rosters), keyed by PlayerID. Created and extended by Sync Extra Player Info. |
| `Final Forms` | Imported CSV data from SPS Final Forms |
| `Newsletter Subscribers` | Imported Buttondown subscriber list (email + status); the mailing-list system of record. Google Groups was retired in spring 2026. |
| `Practice Availability` | Player availability for practices; one row per Player: `PlayerID` in column A (the only typed per-player value), Full Name, Grade, and Gender Identification as Roster lookup formulas, then one column per date |
| `Game Availability` | Player availability for games; same row layout as Practice Availability |

### Roster layout

The 📋 Roster is one header row plus one row per Player: the PlayerID as a plain value in column A and a formula in every other cell that looks up that row's PlayerID. Nothing is authored in it: every value traces to exactly one Source (Signups, Extra Player Info, Final Forms, Newsletter Subscribers) and a wrong value is fixed there. In `Code.gs`, `ROSTER_HEADER_ROW` is 1 and `ROSTER_FIRST_DATA_ROW` is 2; every reader (Build Practice Roster, Game Roster Prep, Full Name Diff, the reports) uses those two constants.

Sort and filter however you like: filter views, the basic filter, or Data > Sort range all work because every row is self-contained (this is why the Roster moved from array formulas to per-row formulas, see [ADR 0003](./docs/adr/0003-roster-per-row-formulas.md)). The trade-off is that column A is values, so a new or removed Signup only reaches the Roster when you run Generate Fresh Roster again; Run Diagnostics compares the Roster's PlayerIDs with Signups and reports when the Roster is stale.

Every Boolean column reads TRUE for "all is well" and FALSE for "something needs a coach's attention": Profile Complete?, the three Final Forms flags and Final Forms Cleared?, Include In Generated Rosters, Gender Default Handling (FALSE when Final Forms and Signup gender disagree or the pronouns are off-pattern for the Gx/Bx), and Media OK (FALSE when the family opted out of media). Filter any Boolean column to FALSE to get a to-do list.

Column definitions live in the `ROSTER_COLUMNS` list in `Code.gs` (name, type, source, note, formula). Adding or reordering a column means editing that list and running Generate Fresh Roster, which rewrites the header row, the header notes (hover a header to see its type, Source, and rule), and every data row. The plan that introduced this layout is in [docs/plans/2026-09-roster-rebuild.md](./docs/plans/2026-09-roster-rebuild.md).

## Data Sources

### Signups (family portal)

- **Source**: the portal's Signups sheet, mastered by the family portal (`../madison-ultimate`)
- **Location**: the `2026 Fall Signups` tab, an IMPORTRANGE of that sheet (read-only here; refreshes on its own)
- **Join Key**: PlayerID. The Roster's column A lists every PlayerID in Signups (written as values, initially sorted by Last Name then Preferred First Name), and every other column joins back by the PlayerID on its own row.
- **Headers referenced by name**: listed in `SIGNUPS_HEADERS` in `Code.gs`; the portal may reorder or add columns freely, and Run Diagnostics reports any referenced header that goes missing.

### Extra Player Info (coach-authored)

- **Source**: coaches, in the `Extra Player Info` tab
- **Columns**: PlayerID, Full Name (formula), Team (dropdown seeded from `CONFIG.teams`; edit the list in the sheet if the teams change), Returning (TRUE/FALSE dropdown), Include In Generated Rosters (TRUE/FALSE dropdown; blank means "included")
- **Sync**: Menu → "Sync Extra Player Info" appends a row for every Signups PlayerID not already present; it never deletes or reorders rows.

### Final Forms (SPS Registration)

- **Source**: CSV exports from SPS Final Forms system
- **Location**: Google Drive folder (configured in `CONFIG.finalForms.folderId`)
- **Import**: Menu → "Update Final Forms" (auto-discovers most recent CSV)
- **Join Key**: SPS Student ID (the Signups row's `SPS Student ID` against Final Forms column A, both coerced to text). Fixed export columns: StudentID A, Parent Signed P, Student Signed Q, Gender U, Grade W, Physical Clearance AB.

### Newsletter Subscribers (Buttondown)

- **Source**: Buttondown Subscribers API
- **Access**: `BUTTONDOWN_API_KEY` script property (see [Script Properties](#script-properties-secrets))
- **Import**: Menu → "Update Newsletter Subscribers" (paginates through all subscribers)
- **Lookup**: Email address → status (`"regular"` = subscribed, `"unactivated"` = pending confirmation, `"unsubscribed"`, or `"not a member"` if not a subscriber at all); matched case-insensitively
- Google Groups was retired as the mailing-list system of record in spring 2026; there is no CSV import anymore.

## Menu Functions

### Diagnostics
- **Run Diagnostics** - Checks that required sheets exist (including Extra Player Info with its header row), that the Signups IMPORTRANGE resolved and carries every header the Roster formulas reference, that the Roster header row has every defined column and its A2 key formula is intact, that the Final Forms Drive folder is reachable, and that the Buttondown key works. Run this after deploying to a new season's spreadsheet, or any time something is misbehaving, before digging further.

### Roster Management
- **Generate Fresh Roster** - Rewrite the 📋 Roster: header row, header notes, PlayerID values in column A, and per-row formulas everywhere else, all keyed by Signups PlayerID. Safe to run any time; nothing authored is lost because nothing is authored there. Run it again whenever Signups gains or loses a Player (Run Diagnostics says when).
- **Sync Extra Player Info** - Create the Extra Player Info tab if missing, apply its dropdowns, and append a row for every Signups PlayerID not already present.
- **Refresh All Data** - Update Final Forms and Newsletter Subscribers imports (Signups refreshes on its own through IMPORTRANGE)

### Data Import
- **Update Final Forms** - Import latest Final Forms CSV
- **Update Newsletter Subscribers** - Import the full Buttondown subscriber list (email + status) into the Newsletter Subscribers sheet. Requires the `BUTTONDOWN_API_KEY` script property (see [Script Properties](#script-properties-secrets)).

### Sheet Builders
- **Build Practice Roster** - Create roster with practice availability columns. Column A is a hidden `PlayerID` (the row's key and the only typed value, in column A like every other PlayerID-keyed sheet), then `#`, Full Name, Team, Gender, Grade; every other cell is a per-row lookup on the PlayerID (Roster for Full Name/Team/Gender/Grade, Practice Availability and Game Availability for the date columns). Unhide column A if you need to see the key.
- **Build Game Roster Prep Sheet** - Create game day roster (coach or parent view); like the practice roster, each row is keyed by a hidden `PlayerID` in column A and every other cell is a lookup on it. If **Game Info** has multiple rows on the **same calendar date**, the prep sheet includes **all** of those games (see [Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)).
- **Build Email List** - Generate email lists (Caretaker 1 and 2 emails) for family communication
- **Build Practice/Game Availability** - Create availability tracking sheets: date columns plus one row per Player. Each run appends a row for every Roster Player whose **Include In Generated Rosters** is TRUE and whose PlayerID is not already in the tab: `PlayerID` as a typed value in column A, then Full Name, Grade, and Gender Identification as Roster lookup formulas keyed by that PlayerID (the same `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))` shape the Roster's own columns use), so they stay live when the Roster changes. Existing rows that have a PlayerID get those three cells rewritten as formulas on every run, which converts sheets built by older versions that copied values. Rows are never deleted or reordered, so **set Include In Generated Rosters FALSE for cut players before the first build** (a row seeded earlier stays until you delete it by hand). An existing row with a Full Name but no PlayerID is filled in when exactly one Roster Player has that name; otherwise the run reports it. The `PlayerID` column is what the player portal matches rows on; the prep sheets find Full Name by header. For **multiple games on the same calendar day**, see **[Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)** below. Normally each game date gets three columns (in order): *$Date* Availability, *$Date* Activation Status (dropdown: Active / Inactive / TBD with green/red/grey backgrounds), and *$Date* Note (free text). If there are multiple **Game Info** rows with the **same date and the same Team**, the script adds a second (or third) set with **`(Game 2)`** / **`(Game 3)`** in the header so they match the player portal; rows for different teams on the same date share one set. Headers and cells use text wrapping.
- **Build Custom Sheet** - Interactive builder for custom column selection

### Analysis Tools
- **Show Statistics** - Totals for Players, Profile Complete, Include In Generated Rosters, each Final Forms flag, Newsletter subscriptions, and the grade distribution
- **Find Emails Not Subscribed to Newsletter** - Identify roster emails that aren't Buttondown Newsletter subscribers
- **Caretakers Not Subscribed to Newsletter** - Find Caretakers whose Buttondown subscriber status isn't "regular" (haven't joined, haven't confirmed, or unsubscribed)
- **Analyze Signups** - Write the "Analyze Signups" sheet: signups with no SPS Student ID, Final Forms students not yet seeded or joined, signups whose SPS Student ID is not in Final Forms, suspected duplicate signups, signups not Profile Complete, and Seeded Signups the family has not finished. It only reports; the portal's Seed Signups from Final Forms (`/admin/final-forms`) does the joining and seeding.
- **Full Name Diff** - Compare names across data sources

### Utilities
- **Format Spruce Up** - Apply consistent formatting
- **Delete Empty Rows & Columns** - Clean up empty space
- **Convert to Actual Attendance** - Convert availability to attendance records
- **Organize Sheets** - Reorder sheet tabs
- **Sync Practice Info to Calendar** - Syncs "🥏🏃 Practice" events from the Practice Info sheet to the team calendar (create/update/delete to match sheet). Uses a **Google Calendar Event ID** column so the script can keep sheet and calendar in sync without guessing.
- **Sync Game Info to Calendar** - Syncs game and warmup events from the Game Info sheet to the team calendar ("🎯 Game vs. X", "🎯 TBD Game", "🥏 Game Warmup"; create/update/delete to match sheet). Uses **Google Calendar Event ID** and **Google Calendar Warmup Event ID** columns. Warmup time comes from the **Warmup Arrival** column.

#### Calendar sync logic (Game and Practice)

The sync scripts store Google Calendar event IDs in the spreadsheet so logic stays tight:

| Row has stored event ID? | Day-of match on calendar (by title or time)? | Action |
|--------------------------|-----------------------------------------------|--------|
| Yes                      | n/a                                           | Ensure calendar event matches the sheet (update if needed). |
| No                       | Yes                                           | Ensure calendar matches sheet and **write the event ID** into the spreadsheet. |
| No                       | No                                            | **Create** a new calendar event and **write the event ID** into the spreadsheet. |

**Spreadsheet columns:** Add (or ensure) these headers so the sync can read/write IDs:

- **📍Game Info:** `Google Calendar Event ID`, `Google Calendar Warmup Event ID`
- **📍Practice Info:** `Google Calendar Event ID`

## Key Concepts

### Multiple events on the same calendar day (“double headers”)

You can schedule **several games on one calendar date** (same `M/D` in **Game Info** more than once). Two cases look alike in the sheet and are told apart by the **Team** column:

- **Several teams, one date** (the normal multi-team Saturday): one Game Info row per team-game, each with its Team. Every team's first game that day is "game 1", so all of them share the single `M/D Availability` / `Activation Status` / `Note` column set; a Player only ever plays their own team's game, so one cell per date is enough.
- **Double-header** (the same team twice on one date, e.g. pool play then finals): two rows with the same Date and the same Team. The second row is "game 2" and gets the **`(Game 2)`** column set.

**Game Info (📍Game Info)**  
- Enter **one row per team-game**, not one row per calendar day. Reuse the same **Date** value for every game that day.
- Fill the **Team** column with the team that row belongs to (the values used in Extra Player Info, e.g. Blue, Gold, Silver). Leave it **blank for an all-team event**; a blank Team is its own group and the portal shows that row to every Player, including those with no team yet.
- Keep rows in **true game order** (earlier game first). The script and the portal assign "game 1" / "game 2" **in row order within each date and Team** (same rule as the portal API).

**Game Availability**  
After you add or change rows in Game Info, run **Build Game Availability**. For each distinct date-and-ordinal, the script ensures columns exist (the activation column only when `CONFIG.gameRosterPrep.hasActivationStatus` is true):

| Occurrence that day (per Team) | Example availability header | Example activation header | Example note header |
|---------------------|----------------------------|---------------------------|---------------------|
| 1st game on that date | `5/9 Availability` | `5/9 Activation Status` | `5/9 Note` |
| 2nd game for the same team | `5/9 Availability (Game 2)` | `5/9 Activation Status (Game 2)` | `5/9 Note (Game 2)` |
| 3rd game for the same team | `5/9 Availability (Game 3)` | … | … |

Do **not** use two identical headers like two columns both named `5/9 Availability`; a same-team second game must use the **`(Game N)`** suffix so the portal can tell them apart.

**Build Game Roster Prep**  
The game picker lists **one option per Game Info row** (date, label, and Team when present). Whichever row you pick, the prep sheet includes **every distinct game ordinal on that calendar day** in order: for each, *Activation Status* (if enabled in `CONFIG.gameRosterPrep`), *Availability*, and *Note*; for example a same-team double-header on `5/16` produces `5/16 Activation Status`, `5/16 Availability`, `5/16 Note`, then `5/16 Activation Status (Game 2)`, `5/16 Availability (Game 2)`, `5/16 Note (Game 2)`. Three teams playing on `5/16` produce the first set only, since they share it.

**Practice roster “next game” columns**  
When a practice roster includes columns for the next game after that practice, **find next game** uses the next Game Info row in order; if that day has two games, you get columns for the **first** game on that date (unless you change Game Info order intentionally).

### Practice & Game Availability: dropdowns and cell colors

After **Build Practice Availability** or **Build Game Availability**, the script applies **data validation** in bulk (one shared rule type for all availability columns, and for games one shared rule for all activation columns, fewer duplicate rules than per-column). **Conditional formatting** fills cells by value: **one rule per distinct availability or activation value**, each rule’s range is the **entire sheet grid** (simple and reliable; only values that exactly match are colored). Rebuilding refreshes those managed rules so they do not stack.

The same **managed** whole-sheet rules are applied to **Build Practice Roster Prep** and **Build Game Roster Prep** sheets when those tabs include practice date columns (`M/D`) and/or game-style `M/D Availability` / `M/D Activation Status` headers (see `ManagedConditionalFormatting.gs`).

**Dropdown “chip” colors** in the validation dropdown list are still a Sheets **UI** feature; Apps Script does not style the list UI. **Cell background** colors come from the automated conditional formatting above.

Roster prep sheets **copy** data validation from Practice / Game Availability. They no longer copy conditional formatting from **📋 Roster** for availability coloring; use the managed rules above so prep tabs match Game Availability cell colors.

### Dynamic Column Positioning

Columns are discovered by header name at runtime, not by position. This means:
- Users can reorder columns freely
- Formulas adapt to current column positions
- New columns can be added anywhere

### PlayerID as the Roster key

The Roster's column A holds every PlayerID in Signups as a plain value, one row per Player; every other cell is a per-row formula that joins by the PlayerID on its row (Signups, Extra Player Info), by SPS Student ID (Final Forms), or by email (Newsletter Subscribers), guarded by `IF($A<row>="","",...)`. They are plain formulas, not `ARRAYFORMULA`, so no scalar function is ever applied to a whole-column range (Sheets would implicitly intersect it with the current row and every lookup would miss); ranges go straight into XLOOKUP and only the single lookup key is converted. Sibling references are row-relative (`$F2` on row 2), so sorting the sheet in place keeps each row consistent. Signups column letters are resolved by header name when the Roster is generated. See [ADR 0001](./docs/adr/0001-roster-keyed-by-signups-playerid.md) for why PlayerID is the key and [ADR 0003](./docs/adr/0003-roster-per-row-formulas.md) for why the rows are per-row formulas rather than one array formula.

### Full Name as the downstream key

Full Name (Preferred First Name followed by Last Name, derived in the Roster) is how humans refer to a Player on every printout and email list, but it is never the join key. The availability sheets are keyed by `PlayerID` in column A, the only per-player value Build Practice/Game Availability types; their Full Name, Grade, and Gender Identification cells are per-row Roster lookup formulas on that PlayerID (ADR 0004). The practice roster and game roster prep sheets carry a hidden `PlayerID` in column A too, and every cell on those rows (Full Name, Team, Gender, Grade, each availability, activation, and note column) is a lookup on it, so two Players with the same name never collide and a preferred-name edit shows up everywhere without a rebuild. Apply Activation Status joins a prep sheet back to Game Availability by PlayerID as well (falling back to Full Name only for a sheet built before the column existed). The player portal never matches on Full Name: it reads and writes availability cells by PlayerID. The Sheet Builder's custom sheets are the one remaining feature that looks up by Full Name.

## File Structure

| File | Purpose |
|------|---------|
| `Code.gs` | Main entry point, menu, core roster functions |
| `Diagnostics.gs` | Run Diagnostics setup checks |
| `ExtraPlayerInfo.gs` | Sync Extra Player Info |
| `AnalyzeSignups.gs` | Analyze Signups report |
| `NewsletterSubscribers.gs` | Update Newsletter Subscribers (Buttondown import) |
| `Availability.gs` | Practice/game availability sheet builders |
| `ManagedConditionalFormatting.gs` | Shared whole-sheet CF for availability/activation values (availability tabs + roster prep) |
| `BuildPracticeRoster.gs` | Practice roster generation |
| `BuildGameRosterPrepSheet.gs` | Game day roster generation |
| `BuildEmailList.gs` | Email list generation |
| `SheetBuilder.gs` | Custom sheet builder |
| `SheetBuilderUtils.gs` | Shared utilities |
| `ConvertToAttendance.gs` | Attendance conversion |
| `FormatSpruceUp.gs` | Formatting utilities |
| `DeleteEmptyRowsColumns.gs` | Cleanup utilities |
| `OrganizeSheets.gs` | Sheet organization |
| `CreatePracticeCalendarEvents.gs` | Sync Practice Info to team calendar (shared sync helper) |
| `CreateGameCalendarEvents.gs` | Sync Game Info to team calendar |
| `FullNameDiff.gs` | Name matching analysis |

## Troubleshooting

### "Column not found" or "missing header" errors
Roster readers look columns up by header name in row 1 of 📋 Roster; run "Generate Fresh Roster" to restore the header row. Generate Fresh Roster itself names any Signups header it cannot find; check the portal's Signups sheet and the IMPORTRANGE in `2026 Fall Signups`!A1.

### Roster is empty or shows #REF!
- Open the `2026 Fall Signups` tab; if A1 shows "You need to connect these sheets", click Allow access
- Run 🩺 Run Diagnostics: it checks the IMPORTRANGE resolved, that the Roster's data rows still have the per-row formula shape, and that the Roster's PlayerIDs match Signups (if not, run "Generate Fresh Roster")

### Final Forms columns blank or FALSE for a Player
The Player's Signups row has no SPS Student ID yet, or that ID is not in the latest export. Run "Analyze Signups" to see which; the portal's Seed Signups from Final Forms does the joining. Run "Update Final Forms" to refresh the export.

### Newsletter status columns show "not a member" everywhere
Run "Update Newsletter Subscribers" to populate the Newsletter Subscribers sheet.

## See Also

- [DESIGN.md](./DESIGN.md) - Detailed architecture and design decisions
- [Initial Requirements.md](./Initial%20Requirements.md) - Original project requirements
