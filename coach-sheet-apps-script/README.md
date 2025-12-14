# Coach Sheet Apps Script

Google Apps Script tools for managing the Madison Middle School Ultimate Frisbee team roster spreadsheet.

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

1. **Increment the version** in `Code.gs`:
   ```javascript
   const SCRIPT_VERSION = '122';  // Increment this
   ```

2. **Push changes**:
   ```bash
   clasp push
   ```

3. **Refresh the Google Sheet** and use the 🥏 Madison Ultimate menu.

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

### Roster Metadata Rows (1-5)

The roster sheet uses 5 metadata rows before player data:

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
- **Build Game Roster Prep Sheet** - Create game day roster (coach or parent view)
- **Build Email List** - Generate email lists for parent communication
- **Build Practice/Game Availability** - Create availability tracking sheets
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

## Key Concepts

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
