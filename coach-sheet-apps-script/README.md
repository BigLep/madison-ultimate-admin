# Coach Sheet Apps Script

Google Apps Script tools for managing the Madison Middle School Ultimate Frisbee team roster spreadsheet.

**Season note:** This README and the `main` branch apply to the **Spring 2026** season. For the 2025 fall season, see the `fall-2025` branch.

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

1. **Increment the version** in `Code.gs` (use format `2.x`, increment x):
   ```javascript
   const SCRIPT_VERSION = '2.1';  // Increment x for each release
   ```

2. **Push changes**:
   ```bash
   clasp push
   ```

3. **Refresh the Google Sheet** and use the 🥏 Madison Ultimate menu.

### Deploy to a new season's spreadsheet (e.g. Spring 2026)

To bind this script to a **new** spreadsheet (the script must be created from the sheet):

1. Open the season's spreadsheet (e.g. [Spring 2026](https://docs.google.com/spreadsheets/d/1kV3Y_GST_Y-X9PZFXu9yFkCzGWvhk9f7G24Y8QNuayU/edit)).
2. **Extensions → Apps Script**. This opens the script editor and creates a new Apps Script project bound to that sheet.
3. In the script editor, open **Project settings** (gear icon) and copy the **Script ID**.
4. In this repo, update `coach-sheet-apps-script/.clasp.json`: set `"scriptId"` to that Script ID (leave `rootDir` and `filePushOrder` as-is).
5. From `coach-sheet-apps-script/` run:
   ```bash
   clasp push
   ```
6. Back in the spreadsheet, refresh the page; the 🥏 Madison Ultimate menu should appear.
7. **Set season-specific Game Roster Prep options** in `Code.gs` under `CONFIG.gameRosterPrep`:
   - **`hasTeam`** – `true` if this season uses teams (e.g. A/B squads) and the roster has a Team column; `false` to omit the Team column from coach and parent game roster prep sheets.
   - **`hasActivationStatus`** – `true` if Game Availability has an Activation Status column per date (Active/Inactive/TBD) and you want it on the coach game roster prep sheet (and sorted first); `false` to omit it.

Note: `.clasp.json` currently points at the Spring 2026 script. After you deploy to a new sheet, switch the `scriptId` in `.clasp.json` depending on which season you’re editing.

## Spreadsheet Structure

### Required Sheets

The script expects these sheets to exist (created manually or via the menu):

| Sheet Name | Purpose |
|------------|---------|
| `📋 Roster` | Main roster with player data and formulas |
| `Final Forms` | Imported CSV data from SPS Final Forms |
| `Additional Info` | Imported questionnaire responses (via IMPORTRANGE) |
| `Mailing List` | Imported CSV of Google Groups membership |
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
| `Mailing List` | Populated by VLOOKUP from mailing list |
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

### Mailing List (Google Groups)

- **Source**: CSV export from Google Groups
- **Location**: Google Drive folder (configured in `CONFIG.mailingList.folderId`)
- **Import**: Menu → "Update Mailing List" (auto-discovers most recent CSV)
- **Lookup**: Email address → membership status ("member", "invited", "not a member")

## Menu Functions

### Roster Management
- **Generate Fresh Roster** - Rebuild all formulas (preserves Manual/Formula columns)
- **Clear Roster Data** - Clear data rows, keep metadata and Manual/Formula columns
- **Refresh All Data** - Update Final Forms and Mailing List imports

### Data Import
- **Update Final Forms** - Import latest Final Forms CSV
- **Update Mailing List** - Import latest mailing list CSV

### Sheet Builders
- **Build Practice Roster** - Create roster with practice availability columns
- **Build Game Roster Prep Sheet** - Create game day roster (coach or parent view). If **Game Info** has multiple rows on the **same calendar date**, the prep sheet includes **all** of those games (see [Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)).
- **Build Email List** - Generate email lists for parent communication
- **Build Practice/Game Availability** - Create availability tracking sheets. For **multiple games on the same calendar day**, see **[Multiple events on one calendar day](#multiple-events-on-the-same-calendar-day-double-headers)** below. Normally each game date gets three columns (in order): *$Date* Availability, *$Date* Activation Status (dropdown: Active / Inactive / TBD with green/red/grey backgrounds), and *$Date* Note (free text). If there are multiple **Game Info** rows with the **same date**, the script adds a second (or third) set with **`(Game 2)`** / **`(Game 3)`** in the header so they match the player portal. Headers and cells use text wrapping.
- **Build Custom Sheet** - Interactive builder for custom column selection

### Analysis Tools
- **Show Statistics** - Display roster completion stats
- **Find Emails Not on Mailing List** - Identify missing mailing list signups
- **Parents Not Members of Mailing List** - Find parents who haven't joined
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
- Ensure Final Forms and Mailing List sheets have data
- Run "Update Final Forms" and "Update Mailing List" to refresh imports
- Check that Student ID column has values (populated during "Generate Fresh Roster")

### Additional Info not matching
The "Full Name" column must exactly match names in the questionnaire. Use "Full Name Diff" to identify mismatches.

## See Also

- [DESIGN.md](./DESIGN.md) - Detailed architecture and design decisions
- [Initial Requirements.md](./Initial%20Requirements.md) - Original project requirements
