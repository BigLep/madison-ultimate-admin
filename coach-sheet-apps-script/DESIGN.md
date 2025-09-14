# Madison Middle School Ultimate Frisbee Roster System - Design Document

## Project Overview
This is a Google Sheets-based roster management system for the Madison Middle School Ultimate Frisbee team (Fall 2025 season). The system combines data from multiple sources to create a comprehensive team roster with 50-100 players in grades 6-8.

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

### Google Apps Script Functions

#### Core Functions
- **`generateRoster()`** - Main function to build/rebuild the roster with all formulas
- **`clearRosterData()`** - Clears rows 6+ while preserving metadata
- **`clearRosterFormatting()`** - Removes all formatting, conditional formatting, etc.

#### Data Update Functions
- **`updateFinalForms()`** - Refreshes Final Forms CSV import
- **`updateMailingList()`** - Refreshes Mailing List CSV import
- **`refreshAllData()`** - Updates both CSV sources (Additional Info auto-updates via IMPORTRANGE)

#### Utility Functions
- **`showStatistics()`** - Displays roster completion statistics
- **`getColumnLetter(columnNumber)`** - Converts column number to letter(s) for formulas
- **`createCustomMenu()`** - Creates the "🥏 Madison Ultimate" menu
- **`onOpen()`** - Auto-creates menu when sheet opens

### Key Design Features

#### 1. Dynamic Column Positioning
- Columns are found by header name, not position
- Users can reorder columns freely
- Formulas adapt to current column positions
- New columns can be added anywhere

#### 2. Join Logic
- **Primary Key**: Player full name (First + Last)
- **Fuzzy Matching**: Not implemented in sheets (would need Apps Script enhancement)
- **Email Matching**: Used as fallback for Additional Info form
- **Parent Email Exclusion**: Student personal email must not match parent emails

#### 3. Boolean Logic for Mailing List
```javascript
IF(AND(
  COUNTIF('Mailing List'!A:A, email) > 0,  // Email exists in list
  INDEX('Mailing List'!F:F, MATCH(email, 'Mailing List'!A:A, 0)) = "allowed"  // Has posting permission
), TRUE, FALSE)
```

#### 4. Gender Simplification
- "Girl-Matching/Gx/Non-binary" → "Gx"
- "Boy-Matching/Bx/Non-binary" → "Bx"

### Formula Patterns

#### Basic Field Import
```javascript
=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!D:D,ROW()-4),""))
```

#### Conditional Field (SPS Email)
```javascript
=IF(ROW()<6,"",IFERROR(IF(REGEXMATCH(INDEX('Final Forms'!F:F,ROW()-4),"@seattleschools\\.org"),INDEX('Final Forms'!F:F,ROW()-4),""),""))
```

#### Lookup from Additional Info
```javascript
=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!C:C,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))
```

## Current File Structure

### Local Files (in Claude's environment)
- `/home/claude/madison_roster_final.gs` - The main Apps Script file
- `/mnt/user-data/outputs/madison_roster_final.gs` - Copy for download

### Related Documentation
- `Roster_Columns.xlsx` - Original column specification from client
- `2025_Fall_Coach_Sheets.xlsx` - Sample of generated roster for testing

## Known Issues and Limitations

1. **No Fuzzy Matching** - Exact name matches only (could add Levenshtein distance in Apps Script)
2. **Manual Column Updates** - If new columns are added to source data, script needs updating
3. **CSV File IDs** - Must manually update when new exports are created
4. **Performance** - Large rosters (>150 students) may be slow to regenerate

## Future Enhancements

### High Priority
1. Implement fuzzy name matching for better join accuracy
2. Add automatic CSV file discovery (find latest by timestamp)
3. Create a configuration sheet for easier maintenance

### Medium Priority
1. Add data validation rules for manual entry fields
2. Implement automatic archiving of old rosters
3. Add email notification when roster is incomplete

### Low Priority
1. Create a dashboard sheet with charts
2. Add integration with team communication tools
3. Implement automatic team assignment based on grade/experience

## Testing Checklist

When modifying the system, verify:
- [ ] Roster generates without errors
- [ ] All 34 standard columns populate correctly
- [ ] Boolean fields show TRUE/FALSE (not Yes/No)
- [ ] Mailing list status shows TRUE for valid emails
- [ ] Column reordering doesn't break formulas
- [ ] Additional custom columns are preserved
- [ ] Statistics function shows accurate counts
- [ ] Clear functions work as expected
- [ ] Gender identification shows Gx/Bx correctly

## Development Notes

### Recent Changes (Sept 11, 2025)
1. Changed from row 7 to row 6 for first data row
2. Removed "END OF METADATA" marker
3. Added repeated headers in row 5 for pivot tables
4. Converted mailing list columns to TRUE/FALSE booleans
5. Implemented dynamic column positioning
6. Fixed apostrophe escaping in column names

### Key Decisions
- **Why Google Sheets?** - Client preference, easy sharing, no infrastructure needed
- **Why Apps Script?** - Native integration, no external dependencies
- **Why metadata rows?** - Self-documenting, preserves context for future maintainers
- **Why dynamic columns?** - Flexibility for coach to customize layout

## Contact and Context
- **Team**: Madison Middle School Ultimate Frisbee
- **Season**: Fall 2025
- **Size**: 50-100 players
- **Grades**: 6-8
- **Previous Season Doc**: https://docs.google.com/document/d/1A2F7ThHtcMm23bxk8-30rMT2svaqT3gMbRWeSR_QXXY

## How to Resume Development

1. **Load the Apps Script**: Copy `madison_roster_final.gs` to Google Apps Script editor
2. **Test Current State**: Run `generateRoster()` to verify it works
3. **Review Data Sources**: Check if CSV file IDs need updating
4. **Check Column Mappings**: Verify all 34 columns still match source data
5. **Test Your Changes**: Use the testing checklist above

## Critical Implementation Details

### Column Reference Updates
When formulas reference other columns (like combining First + Last name), the script:
1. Finds the current position of each referenced column
2. Converts position to column letter
3. Updates formula with actual positions

Example:
```javascript
const firstNameCol = columnMap.get('First Name') || 1;
const firstNameLetter = getColumnLetter(firstNameCol);
formula = formula.replace(/A6/g, firstNameLetter + '6');
```

### Mailing List Validation
The mailing list check is a two-part validation:
1. Email exists in column A of Mailing List sheet
2. Column F (Posting permissions) = "allowed"

Both conditions must be TRUE for the boolean to be TRUE.

### Service Account
For potential future automation:
- Email: stevel@cedar-scene-471205-t3.iam.gserviceaccount.com
- Could be used for automated imports if configured

---

*This design document captures the complete state of the Madison Ultimate Roster system as of September 11, 2025. Use this to understand the system architecture, make modifications, or hand off development to another engineer.*