# Coach Sheet Apps Script

Google Apps Script tools for managing the Madison Middle School Ultimate Frisbee team roster spreadsheet.

**Season note:** This README and the `main` branch apply to the **Fall 2026** season. Past seasons get their own dated branch.

## Purpose

This script provides a custom menu in Google Sheets ("🥏 Madison Ultimate") that automates:
- Importing and syncing player data from Final Forms registration
- Tracking mailing list membership status
- Building practice and game rosters with availability
- Generating email lists for parent communication
- Analyzing data quality and missing information

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

1. **Increment `SCRIPT_VERSION`** in `Code.gs` **before every push**, not after — this is a standing rule, not optional:
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
   - **There is no API or CLI shortcut for this step.** `gog`'s Apps Script commands (and the underlying Apps Script/Drive APIs) all require a Script ID you already have — none of them can look up "the script bound to spreadsheet X." The Apps Script editor UI is the only way to discover a newly-duplicated bound script's ID.
   - The duplicated project's internal title stays whatever the source project was called (e.g. still "2026 Spring Coach Sheet Admin" after duplicating for fall); consider renaming it in Project settings for clarity, though this is cosmetic and doesn't affect deployment.
4. In this repo, update `coach-sheet-apps-script/.clasp.json`: set `"scriptId"` to that Script ID (leave `rootDir` and `filePushOrder` as-is).
5. **Update per-season values in `Code.gs`** before pushing:
   - `CONFIG.finalForms.folderId` — must point at **this season's** FinalForms exports Drive folder, and must match the `finalforms-export` GitHub Action's `DRIVE_FOLDER_ID` repo variable (`gh variable list -R BigLep/madison-ultimate-admin`). If these two drift apart, "Update Final Forms" silently imports the wrong season's data with no error.
   - `CONFIG.gameRosterPrep.hasTeam` — `true` if this season uses teams (e.g. A/B squads) and the roster has a Team column; `false` to omit it from game roster prep sheets.
   - `CONFIG.gameRosterPrep.hasActivationStatus` — `true` if Game Availability has an Activation Status column per date (Active/Inactive/TBD) and you want it on the coach game roster prep sheet (sorted first); `false` to omit it.
6. **Increment `SCRIPT_VERSION`**, then from `coach-sheet-apps-script/` run:
   ```bash
   clasp push
   ```
7. Back in the spreadsheet, refresh the page; the 🥏 Madison Ultimate menu should appear.
8. **Confirm required sheets exist.** Duplicating a spreadsheet copies every tab, including data sheets the script itself writes into (e.g. **Final Forms**) that look like leftover imported data and are easy to delete by mistake during season cleanup — don't delete a sheet during setup without first grepping `Code.gs` for its name to check whether the script depends on it.
9. **Set any required Script Properties** (see below) — these are per-project, not copied by duplicating the spreadsheet, so they must be re-entered on every new bound script.
10. **Run 🩺 Run Diagnostics** from the menu as the final step. It checks that required sheets exist, configured Drive folders/spreadsheets are reachable, and required Script Properties are set — this is the "confirm everything is configured correctly" step for a new season, and the fast way to catch the kind of drift described above before a coach hits it mid-season.

## Script Properties (secrets)

Some integrations need credentials that must never be committed to this repo. Those are stored per-project in Apps Script's **Script Properties**, not in `Code.gs`:

1. Open the spreadsheet → **Extensions → Apps Script**.
2. **Project Settings** (gear icon) → **Script Properties** → **Add script property**.
3. Enter the property name and value, then read it in code with `PropertiesService.getScriptProperties().getProperty('NAME')`.

Script Properties are scoped to the Apps Script project (not the spreadsheet), so duplicating the spreadsheet for a new season does **not** carry them over — they must be re-added on the newly bound project each season. `🩺 Run Diagnostics` checks that required properties are set (and, for Buttondown, that the key actually works).

### Required properties

| Property | Used by | How to get it |
|----------|---------|----------------|
| `BUTTONDOWN_API_KEY` | 📬 Update Newsletter Subscribers | Buttondown → Settings → Programming → API Keys. Read access to subscribers is enough; do not paste it into `Code.gs` or this repo. |

## Spreadsheet Structure

### Required Sheets

The script expects these sheets to exist (created manually or via the menu):

| Sheet Name | Purpose |
|------------|---------|
| `📋 Roster` | Main roster with player data and formulas |
| `Final Forms` | Imported CSV data from SPS Final Forms |
| `Additional Info` | Imported questionnaire responses (via IMPORTRANGE) |
| `Newsletter Subscribers` | Imported Buttondown subscriber list (email + status); the mailing-list system of record. Google Groups was retired in spring 2026. |
| `Practice Availability` | Player availability for practices |
| `Game Availability` | Player availability for games |

### Roster layout (configurable)

In `Code.gs`, **ROSTER_HEADER_ROW** (default 1) is the row that contains column headers (Full Name, etc.) in the 📋 Roster sheet. **ROSTER_FIRST_DATA_ROW** is the next row (header + 1) and is used when reading the roster for Build Practice Roster, Game Roster, Full Name Diff, and Additional Info analysis. Change these if your roster uses a different layout (e.g. set ROSTER_HEADER_ROW to 5 and ROSTER_FIRST_DATA_ROW to 6 if you use 5 metadata rows).

### Roster Metadata Rows (1-5)

The roster sheet can use 5 metadata rows before player data:

| Row | Purpose | Example |
|-----|---------|---------|
| 1 | Column headers | "First Name", "Grade", etc. |
| 2 | Data type | "String", "Email", "Boolean" |
| 3 | Data source | "Final Forms", "Manual", "Formula" |
| 4 | Notes | Implementation details |
| 5 | Repeat headers | For pivot table compatibility |

**Row 6+** contains player data.

### Column Source Types

The `source` row (row 3) controls how columns are handled:

| Source | Behavior |
|--------|----------|
| `Final Forms` | Populated by XLOOKUP formulas from Final Forms sheet |
| `Additional Info` | Populated by INDEX/MATCH from questionnaire data |
| `Newsletter Subscribers Buttondown status` | Populated by VLOOKUP from the Newsletter Subscribers sheet |
| `Manual` | User-entered data, preserved during roster regeneration |
| `Formula` | Custom formulas, preserved during roster regeneration |
| (empty) | Preserved during roster regeneration |

## Data Sources

### Final Forms (SPS Registration)

- **Source**: CSV exports from SPS Final Forms system
- **Location**: Google Drive folder (configured in `CONFIG.finalForms.folderId`)
- **Import**: Menu → "Update Final Forms" (auto-discovers most recent CSV)
- **Join Key**: Student ID (column A)

### Additional Info Questionnaire

- **Source**: Google Form responses
- **Location**: Linked spreadsheet (configured in `CONFIG.additionalInfo.spreadsheetId`)
- **Import**: Auto-updates via IMPORTRANGE
- **Join Key**: Full Name (must match roster's "Full Name" column exactly)

### Newsletter Subscribers (Buttondown)

- **Source**: Buttondown Subscribers API
- **Access**: `BUTTONDOWN_API_KEY` script property (see [Script Properties](#script-properties-secrets))
- **Import**: Menu → "Update Newsletter Subscribers" (paginates through all subscribers)
- **Lookup**: Email address → status (`"regular"` = subscribed, `"unactivated"` = pending confirmation, `"unsubscribed"`, or `"not a member"` if not a subscriber at all)
- Google Groups was retired as the mailing-list system of record in spring 2026; there is no CSV import anymore.

## Menu Functions

### Diagnostics
- **Run Diagnostics** - Checks that required sheets exist, configured Drive folders/spreadsheets are reachable, and required Script Properties are set. Run this after deploying to a new season's spreadsheet, or any time something is misbehaving, before digging further.

### Roster Management
- **Generate Fresh Roster** - Rebuild all formulas (preserves Manual/Formula columns)
- **Clear Roster Data** - Clear data rows, keep metadata and Manual/Formula columns
- **Refresh All Data** - Update Final Forms and Newsletter Subscribers imports

### Data Import
- **Update Final Forms** - Import latest Final Forms CSV
- **Update Newsletter Subscribers** - Import the full Buttondown subscriber list (email + status) into the Newsletter Subscribers sheet. Requires the `BUTTONDOWN_API_KEY` script property (see [Script Properties](#script-properties-secrets)).

### Sheet Builders
- **Build Practice Roster** - Create roster with practice availability columns
- **Build Game Roster Prep Sheet** - Create game day roster (coach or parent view). If **Game Info** has multiple rows on the **same calendar date**, the prep sheet includes **all** of those games (see [Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)).
- **Build Email List** - Generate email lists for parent communication
- **Build Practice/Game Availability** - Create availability tracking sheets. For **multiple games on the same calendar day**, see **[Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)** below. Normally each game date gets three columns (in order): *$Date* Availability, *$Date* Activation Status (dropdown: Active / Inactive / TBD with green/red/grey backgrounds), and *$Date* Note (free text). If there are multiple **Game Info** rows with the **same date**, the script adds a second (or third) set with **`(Game 2)`** / **`(Game 3)`** in the header so they match the player portal. Headers and cells use text wrapping.
- **Build Custom Sheet** - Interactive builder for custom column selection

### Analysis Tools
- **Show Statistics** - Display roster completion stats
- **Find Emails Not on Mailing List** - Identify roster emails that aren't Buttondown newsletter subscribers
- **Parents Not Members of Mailing List** - Find parents whose Buttondown subscriber status isn't "regular" (haven't joined, haven't confirmed, or unsubscribed)
- **Analyze Additional Info Responses** - Check questionnaire matching
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

You can schedule **several games on one calendar date** (same `M/D` in **Game Info** more than once). That is fully supported end-to-end: sheet columns, roster prep, and the player portal all treat each row as a separate game, in **sheet row order** for that date.

This section is also what people mean by **double-headers** (e.g. two league games Saturday, or pool play then finals the same day).

**Game Info (📍Game Info)**  
- Enter **one row per game**—not one row per calendar day. Reuse the same **Date** value for every game that day (e.g. two rows both `5/9`).
- Keep rows in **true game order** (earlier game first). The script and the portal assign “game 1” / “game 2” **in row order** for that date (same rule as the portal API).

**Game Availability**  
After you add or change rows in Game Info, run **Build Game Availability**. For each distinct game row, the script ensures columns exist:

| Occurrence that day | Example availability header | Example activation header | Example note header |
|---------------------|----------------------------|---------------------------|---------------------|
| 1st game on that date | `5/9 Availability` | `5/9 Activation Status` | `5/9 Note` |
| 2nd game | `5/9 Availability (Game 2)` | `5/9 Activation Status (Game 2)` | `5/9 Note (Game 2)` |
| 3rd game | `5/9 Availability (Game 3)` | … | … |

Do **not** use two identical headers like two columns both named `5/9 Availability`—the second game must use the **`(Game N)`** suffix so the portal can tell them apart.

**Build Game Roster Prep**  
The game picker lists **one option per Game Info row** (date plus label when present). Whichever row you pick, the prep sheet includes **every game on that calendar day** in order: for each game, *Activation Status* (if enabled in `CONFIG.gameRosterPrep`), *Availability*, and *Note*—for example two games on `5/16` produce `5/16 Activation Status`, `5/16 Availability`, `5/16 Note`, then `5/16 Activation Status (Game 2)`, `5/16 Availability (Game 2)`, `5/16 Note (Game 2)`.

**Practice roster “next game” columns**  
When a practice roster includes columns for the next game after that practice, **find next game** uses the next Game Info row in order; if that day has two games, you get columns for the **first** game on that date (unless you change Game Info order intentionally).

### Practice & Game Availability — dropdowns and cell colors

After **Build Practice Availability** or **Build Game Availability**, the script applies **data validation** in bulk (one shared rule type for all availability columns, and for games one shared rule for all activation columns—fewer duplicate rules than per-column). **Conditional formatting** fills cells by value: **one rule per distinct availability or activation value**, each rule’s range is the **entire sheet grid** (simple and reliable; only values that exactly match are colored). Rebuilding refreshes those managed rules so they do not stack.

The same **managed** whole-sheet rules are applied to **Build Practice Roster Prep** and **Build Game Roster Prep** sheets when those tabs include practice date columns (`M/D`) and/or game-style `M/D Availability` / `M/D Activation Status` headers (see `ManagedConditionalFormatting.gs`).

**Dropdown “chip” colors** in the validation dropdown list are still a Sheets **UI** feature; Apps Script does not style the list UI. **Cell background** colors come from the automated conditional formatting above.

Roster prep sheets **copy** data validation from Practice / Game Availability. They no longer copy conditional formatting from **📋 Roster** for availability coloring; use the managed rules above so prep tabs match Game Availability cell colors.

### Dynamic Column Positioning

Columns are discovered by header name at runtime, not by position. This means:
- Users can reorder columns freely
- Formulas adapt to current column positions
- New columns can be added anywhere

### Student ID as Primary Key

All Final Forms data uses XLOOKUP with Student ID:
```javascript
=IFERROR(XLOOKUP(StudentID,'Final Forms'!A:A,'Final Forms'!D:D),"")
```
This ensures formulas work correctly regardless of row sorting.

### Full Name as Additional Info Join Key

The "Full Name" column is a manually-maintained join key for Additional Info lookups. It must match the name format in the questionnaire responses exactly.

## File Structure

| File | Purpose |
|------|---------|
| `Code.gs` | Main entry point, menu, core roster functions |
| `Diagnostics.gs` | Run Diagnostics setup checks |
| `NewsletterSubscribers.gs` | Update Newsletter Subscribers (Buttondown import) |
| `Availability.gs` | Practice/game availability sheet builders |
| `ManagedConditionalFormatting.gs` | Shared whole-sheet CF for availability/activation values (availability tabs + roster prep) |
| `BuildPracticeRoster.gs` | Practice roster generation |
| `BuildGameRosterPrepSheet.gs` | Game day roster generation |
| `BuildEmailList.gs` | Email list generation |
| `SheetBuilder.gs` | Custom sheet builder |
| `SheetBuilderUtils.gs` | Shared utilities |
| `AdditionalInfoAnalysis.gs` | Questionnaire analysis |
| `ConvertToAttendance.gs` | Attendance conversion |
| `FormatSpruceUp.gs` | Formatting utilities |
| `DeleteEmptyRowsColumns.gs` | Cleanup utilities |
| `OrganizeSheets.gs` | Sheet organization |
| `CreatePracticeCalendarEvents.gs` | Sync Practice Info to team calendar (shared sync helper) |
| `CreateGameCalendarEvents.gs` | Sync Game Info to team calendar |
| `FullNameDiff.gs` | Name matching analysis |

## Troubleshooting

### "Column not found" errors
The script requires specific column headers. Check that the required column exists in row 1 of the roster sheet.

### Formulas showing errors
- Ensure Final Forms and Newsletter Subscribers sheets have data
- Run "Update Final Forms" and "Update Newsletter Subscribers" to refresh imports
- Check that Student ID column has values (populated during "Generate Fresh Roster")

### Additional Info not matching
The "Full Name" column must exactly match names in the questionnaire. Use "Full Name Diff" to identify mismatches.

## See Also

- [DESIGN.md](./DESIGN.md) - Detailed architecture and design decisions
- [Initial Requirements.md](./Initial%20Requirements.md) - Original project requirements
