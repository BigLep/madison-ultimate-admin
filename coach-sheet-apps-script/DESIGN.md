# Madison Middle School Ultimate Frisbee Roster System - Design Document

## Project Overview
This is a Google Sheets-based roster management system for the Madison Middle School Ultimate Frisbee team. The system combines data from multiple sources to create a comprehensive team roster with 50-100 players in grades 6-8.

**Seasons:** Fall 2025 (initial development), ready for future seasons

## Key Stakeholder
- **Coach/Admin**: Manages the team roster, experienced in software engineering but not proficient in pandas/spreadsheets
- **Contact**: Uses this system to track player information, parent contacts, and team logistics

## System Architecture

### Primary Google Sheet
- **URL**: https://docs.google.com/spreadsheets/d/1ZZA5TxHu8nmtyNORm3xYtN5rzP3p1jtW178UgRcxLA8
- **Main Tab**: "Roster" - Contains the combined data with metadata rows and formulas
- **Supporting Tabs**: 
  - "Final Forms" - Import of registration data
  - "Additional Info" - Import of questionnaire responses
  - "Mailing List" - Import of Google Groups membership

### Data Sources

#### 1. SPS Final Forms (Primary Source)
- **Location**: Google Drive folder - https://drive.google.com/drive/folders/1SnWCxDIn3FxJCvd1JcWyoeoOMscEsQcW
- **Current File ID**: `1pWUIw2rM0MfNWnaC3Ltsz6Wj8_PGFHrH`
- **Format**: CSV export with one row per player
- **Key Fields**: Student info, parent/guardian contacts (2 sets), forms status, physical clearance
- **Update Frequency**: Manual export, filename includes ISO8601 timestamp
- **Determines**: The authoritative list of registered players

#### 2. Additional Questionnaire (Google Form Responses)
- **Sheet ID**: `1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs`
- **Format**: Google Sheets form responses, auto-updating
- **Key Fields**: Jersey size, pronouns, experience, transportation needs, parent volunteer info
- **Join Key**: Player name (First Last format)

#### 3. Team Mailing List (Google Groups Export)
- **Location**: Google Drive folder - https://drive.google.com/drive/folders/1pAeQMEqiA9QdK9G5yRXsqgbNVzEU7R1E
- **Current File ID**: `1n0h81l31lsGvvSPrZUT5SOuS6jXT4h6E`
- **Format**: CSV with one row per email address
- **Key Fields**: Email address (Column A), Posting permissions (Column F)
- **Purpose**: Track which parents/students are on the team mailing list with posting privileges

## Roster Structure

### Metadata Rows (1-5)
1. **Row 1**: Column Names (the field names)
2. **Row 2**: Data Types (String, Boolean, Email, Date, etc.)
3. **Row 3**: Data Sources (e.g., "FinalForms First Name")
4. **Row 4**: Additional Notes (implementation notes and business rules)
5. **Row 5**: Repeated Column Names (for pivot table selection convenience)

### Data Rows (6+)
- **Row 6 onwards**: Student/player data with formulas that reference the imported data

### Column Specifications (34 Standard Columns)

#### Identity Columns (1-3)
- **First Name** - From Final Forms
- **Preferred Name** - Manual entry field
- **Last Name** - From Final Forms

#### Email Columns (4-6)
- **Student SPS Email** - Only populated if domain is @seattleschools.org
- **Student Personal Email** - Only if NOT @seattleschools.org AND not a parent email
- **Student Personal Email On Mailing List?** - Boolean: TRUE if email is on list with "allowed" posting

#### Forms Status (7-9)
- **Are All Forms Parent Signed** - Boolean from Final Forms
- **Are All Forms Student Signed** - Boolean from Final Forms
- **Physical Cleared** - Boolean: TRUE if status is "Cleared"

#### Demographics (10-12)
- **Gender** - From Final Forms
- **Grade** - Number from Final Forms
- **Date of Birth** - Date from Final Forms

#### Parent 1 Info (13-16)
- **Parent 1 First Name**
- **Parent 1 Last Name**
- **Parent 1 Email**
- **Parent 1 Email On Mailing List?** - Boolean: TRUE if on list with "allowed" posting

#### Parent 2 Info (17-20)
- **Parent 2 First Name**
- **Parent 2 Last Name**
- **Parent 2 Email**
- **Parent 2 Email On Mailing List?** - Boolean: TRUE if on list with "allowed" posting

#### Additional Info Form Data (21-34)
- **Additional Info Questionnaire Filled Out?** - Boolean: TRUE if match found
- **Player Pronouns** - From form
- **Player Gender Identification** - Simplified to "Gx" or "Bx"
- **Player Allergies**
- **Competing Sports and Activities**
- **Jersey Size**
- **Playing Experience**
- **Player hopes for the season**
- **Other Player Info**
- **Are you interested in helping coach?**
- **Have you played or coached Ultimate before?**
- **Have you played or coached other team sports?**
- **Are you interested in helping in other ways?**
- **Anything else you want to share?**

## Technical Implementation

### Google Apps Script Files

The script is organized into multiple `.gs` files:
- **`Code.gs`** - Main entry point, menu creation, core roster functions
- **`Availability.gs`** - Practice and game availability sheet builders
- **`BuildPracticeRoster.gs`** - Practice roster sheet generation
- **`BuildGameRosterPrepSheet.gs`** - Game day roster preparation sheets
- **`BuildEmailList.gs`** - Email list generation for parent communication
- **`SheetBuilder.gs`** - Custom sheet builder utility
- **`SheetBuilderUtils.gs`** - Shared utilities for sheet builders
- **`AdditionalInfoAnalysis.gs`** - Analysis of questionnaire responses
- **`ConvertToAttendance.gs`** - Convert availability to actual attendance
- **`FormatSpruceUp.gs`** - Sheet formatting utilities
- **`DeleteEmptyRowsColumns.gs`** - Cleanup utilities
- **`OrganizeSheets.gs`** - Sheet tab organization
- **`FullNameDiff.gs`** - Name matching analysis

### Menu Functions (🥏 Madison Ultimate)

#### Roster Management
- **`generateRoster()`** - Build/rebuild roster with XLOOKUP formulas
- **`clearRosterData()`** - Clear data rows while preserving metadata and Manual/Formula columns
- **`refreshAllData()`** - Update all CSV data sources

#### Data Import
- **`updateFinalForms()`** - Import latest Final Forms CSV (auto-discovers most recent file)
- **`updateMailingList()`** - Import latest Mailing List CSV (auto-discovers most recent file)

#### Sheet Builders
- **`buildCustomSheet()`** - Interactive custom sheet builder
- **`buildPracticeRoster()`** - Create/update practice roster with availability
- **`buildGameRosterPrepSheet()`** - Create game day roster (coach or parent view)
- **`buildEmailList()`** - Generate email lists for parent communication
- **`buildPracticeAvailability()`** - Create practice availability tracking sheet
- **`buildGameAvailability()`** - Create game availability tracking sheet

#### Analysis & Utilities
- **`showStatistics()`** - Display roster completion statistics
- **`findMissingEmails()`** - Find emails not on mailing list
- **`findPendingParents()`** - Find parents who haven't joined mailing list
- **`analyzeAdditionalInfoResponses()`** - Analyze questionnaire response matching
- **`fullNameDiff()`** - Compare names across data sources

#### Formatting & Cleanup
- **`formatSpruceUp()`** - Apply consistent formatting
- **`deleteEmptyRowsAndColumns()`** - Remove empty rows/columns
- **`convertToActualAttendance()`** - Convert availability to attendance
- **`organizeSheets()`** - Organize sheet tab order

#### Internal Utilities
- **`getColumnLetter(columnNumber)`** - Convert column number to letter(s)
- **`createCustomMenu()`** - Create the menu
- **`onOpen()`** - Auto-create menu on sheet open
- **`validateRosterMetadata()`** - Validate script/sheet metadata sync
- **`findMostRecentCsvFile()`** - Find latest CSV in a Drive folder

### Key Design Features

#### 1. Dynamic Column Positioning
- Columns are found by header name, not position
- Users can reorder columns freely
- Formulas adapt to current column positions
- New columns can be added anywhere

#### 2. Join Logic
- **Primary Key for Final Forms**: Student ID (XLOOKUP-based formulas)
- **Primary Key for Additional Info**: Full Name column (manually maintained join key)
- **Parent Email Exclusion**: Student personal email must not match parent emails

#### 3. Mailing List Status
Uses VLOOKUP to check membership status (returns "member", "invited", "not a member", etc.)

#### 4. Gender Simplification
- "Girl-Matching/Gx/Non-binary" → "Gx"
- "Boy-Matching/Bx/Non-binary" → "Bx"

### Formula Patterns

#### XLOOKUP with Student ID (Primary Pattern for Final Forms)
```javascript
=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!D:D),"")
// A6 = Student ID column (dynamically replaced at runtime)
```

#### Conditional Field (SPS Email)
```javascript
=IFERROR(IF(REGEXMATCH(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),"@seattleschools\\.org"),XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),""),"")
```

#### Lookup from Additional Info (using Full Name)
```javascript
=IFERROR(INDEX('Additional Info'!C:C,MATCH(E6,'Additional Info'!B:B,0)),"")
// E6 = Full Name column (dynamically replaced at runtime)
```

#### Mailing List Status Check
```javascript
=IFERROR(VLOOKUP(email,'Mailing List'!$A$3:$C,3,FALSE),"not a member")
```

**Note:** All column references (A6, E6, etc.) are placeholders that get replaced with actual column letters at runtime based on dynamic column discovery.

## Current File Structure

### Repository Structure
```
madison-ultimate-admin/
├── coach-sheet-apps-script/     # Google Apps Script files
│   ├── Code.gs                  # Main script with core functions
│   ├── Availability.gs          # Practice/game availability builders
│   ├── BuildPracticeRoster.gs   # Practice roster generation
│   ├── BuildGameRosterPrepSheet.gs # Game roster generation
│   ├── BuildEmailList.gs        # Email list generation
│   ├── SheetBuilder.gs          # Custom sheet builder
│   ├── SheetBuilderUtils.gs     # Shared utilities
│   ├── AdditionalInfoAnalysis.gs # Response analysis
│   ├── ConvertToAttendance.gs   # Attendance conversion
│   ├── FormatSpruceUp.gs        # Formatting utilities
│   ├── DeleteEmptyRowsColumns.gs # Cleanup utilities
│   ├── OrganizeSheets.gs        # Sheet organization
│   ├── FullNameDiff.gs          # Name matching
│   ├── appsscript.json          # Apps Script manifest
│   ├── .clasp.json              # Clasp deployment config
│   ├── CLAUDE.md                # Claude Code instructions
│   ├── README.md                # Quick reference
│   ├── DESIGN.md                # This document
│   └── Initial Requirements.md  # Original requirements
└── photo-mapper/                # Photo-to-player mapping tool
    ├── frontend/                # NextJS React app
    └── backend/                 # Flask Python API
```

### Deployment
Uses `clasp` for deployment to Google Apps Script:
```bash
clasp push  # Deploy changes (remember to increment SCRIPT_VERSION in Code.gs)
```

## Known Issues and Limitations

1. **No Fuzzy Matching** - Exact name matches only for Additional Info lookups
2. **Manual Column Updates** - If new columns are added to source data, script needs updating
3. **Performance** - Large rosters (>150 students) may be slow to regenerate

## Completed Enhancements (Fall 2025)

- ✅ Automatic CSV file discovery (finds latest file by timestamp)
- ✅ XLOOKUP-based formulas with Student ID as primary key
- ✅ Practice and game availability tracking
- ✅ Practice and game roster builders with multiple view options
- ✅ Email list generation for parent communication
- ✅ Mailing list status tracking (member/invited/not a member)
- ✅ Additional Info response analysis
- ✅ Sheet organization and formatting utilities

## Future Enhancements

### Medium Priority
1. Implement fuzzy name matching for better Additional Info join accuracy
2. Add data validation rules for manual entry fields
3. Create a configuration sheet for easier maintenance

### Low Priority
1. Create a dashboard sheet with charts
2. Add integration with team communication tools
3. Implement automatic team assignment based on grade/experience

## Testing Checklist

When modifying the system, verify:
- [ ] Roster generates without errors
- [ ] Student ID column populates from Final Forms
- [ ] XLOOKUP formulas pull correct data from Final Forms
- [ ] Additional Info lookups work via Full Name
- [ ] Mailing list status shows member/invited/not a member
- [ ] Column reordering doesn't break formulas (dynamic positioning)
- [ ] Manual and Formula source columns are preserved during clear
- [ ] Statistics function shows accurate counts
- [ ] Practice/Game roster builders work correctly
- [ ] Gender identification shows Gx/Bx correctly

## Development Notes

### Development History
- **Sept 2025**: Initial development - basic roster with ROW()-based formulas
- **Oct-Nov 2025**: XLOOKUP migration, Student ID as primary key, availability tracking
- **Dec 2025**: Game roster builders, email list generation, parent roster views

### Key Decisions
- **Why Google Sheets?** - Client preference, easy sharing, no infrastructure needed
- **Why Apps Script?** - Native integration, no external dependencies
- **Why metadata rows?** - Self-documenting, preserves context for future maintainers
- **Why dynamic columns?** - Flexibility for coach to customize layout
- **Why XLOOKUP with Student ID?** - Sort-safe formulas that don't break when roster is reordered
- **Why Full Name as Additional Info join key?** - Google Form doesn't collect Student ID

## Contact and Context
- **Team**: Madison Middle School Ultimate Frisbee
- **Size**: 50-100 players
- **Grades**: 6-8
- **Season Planning Doc**: https://docs.google.com/document/d/1A2F7ThHtcMm23bxk8-30rMT2svaqT3gMbRWeSR_QXXY

## How to Resume Development

1. **Clone the repository** and navigate to `coach-sheet-apps-script/`
2. **Login to clasp**: `clasp login` (if not already authenticated)
3. **Push to test**: `clasp push` (remember to increment SCRIPT_VERSION in Code.gs)
4. **Test in Google Sheets**: Open the spreadsheet and use the 🥏 Madison Ultimate menu
5. **Review data sources**: CSV imports auto-discover the latest file by timestamp
6. **Use the testing checklist** above to verify functionality

## Critical Implementation Details

### Student ID as Primary Key
All Final Forms data uses XLOOKUP with Student ID:
```javascript
=IFERROR(XLOOKUP(StudentID,'Final Forms'!A:A,'Final Forms'!D:D),"")
```
This ensures formulas work correctly regardless of row ordering.

### Full Name as Additional Info Join Key
Additional Info uses the "Full Name" column (manually maintained) as the join key:
```javascript
=IFERROR(INDEX('Additional Info'!C:C,MATCH(FullName,'Additional Info'!B:B,0)),"")
```

### Dynamic Column Discovery
All column positions are discovered at runtime by searching headers:
```javascript
const columnMap = new Map();
headers.forEach((header, index) => columnMap.set(header, index + 1));
const studentIdCol = columnMap.get('StudentID');
```

### Mailing List Status
Uses VLOOKUP to get membership status from column C:
```javascript
=IFERROR(VLOOKUP(email,'Mailing List'!$A$3:$C,3,FALSE),"not a member")
```
Returns: "member", "invited", or "not a member"

### Service Account (for photo-mapper)
- Email: stevel@cedar-scene-471205-t3.iam.gserviceaccount.com
- Used by photo-mapper tool for Drive/Sheets API access

---

*Last updated: December 2025. Use this to understand the system architecture, make modifications, or hand off development.*