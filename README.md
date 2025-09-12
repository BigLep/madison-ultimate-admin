# Madison Ultimate Admin

Tools and scripts for managing the Madison Middle School Ultimate Frisbee team roster and communications.

## Data Integration & Join Keys

### 🔑 Full Name Column as Join Key

The **"Full Name" column** is the critical join key between the Roster and Additional Info sheets:

- **Formula Column**: "Full Name" is a manual formula column (not in script metadata) 
- **Join Key**: This column has been manually aligned to match names in the Additional Info sheet
- **Runtime Discovery**: Script finds this column at runtime by searching headers
- **Primary Usage**: All Additional Info lookups use "Full Name" as the join key, NOT concatenated First+Last names
- **Manual Maintenance**: Any name discrepancies between roster and Additional Info should be resolved by updating the "Full Name" formula

**Important**: When adding formulas that reference Additional Info data, always use the "Full Name" column as the lookup key.

## Dynamic Column Positioning

### 🚫 No Hardcoded Column References

The Google Apps Script is designed to be **completely position-independent**:

- **Runtime Discovery**: All column positions are discovered dynamically by searching headers
- **No Hardcoded References**: Script never assumes columns are in specific positions (A, B, C, etc.)
- **Error on Missing Columns**: Script throws clear errors if required columns are not found
- **Flexible Rearrangement**: Users can reorder columns freely without breaking functionality

### 🔧 Implementation Pattern

**❌ Wrong (Hardcoded):**
```javascript
const emailCol = columnMap.get('Student Personal Email') || 5; // BAD: Assumes column E
formula = `=VLOOKUP(E6,'Mailing List'!A:C,3,FALSE)`;           // BAD: Hardcoded E6
```

**✅ Correct (Dynamic):**
```javascript
const emailCol = columnMap.get('Student Personal Email');       // Dynamic lookup
if (!emailCol) throw new Error('Column not found');             // Clear error
const emailLetter = getColumnLetter(emailCol);                  // Convert to letter
formula = `=VLOOKUP(${emailLetter}6,'Mailing List'!A:C,3,FALSE)`;  // Dynamic reference
```

### 📋 Design Principles

1. **Sheet Layout Freedom**: Users control column order through the Google Sheet interface
2. **Runtime Column Discovery**: Script finds columns by name, not position
3. **Explicit Error Handling**: Missing columns cause clear, actionable error messages
4. **Formula Generation**: All cell references in formulas are generated dynamically
5. **No Position Assumptions**: Script works regardless of how columns are arranged

## Metadata Management

The Google Apps Script respects the **sheet as the source of truth** for metadata configuration. The script validates metadata but never overwrites it.

### 📊 Sheet Structure

The roster sheet uses 5 metadata rows:
- **Row 1**: Column Name (header)
- **Row 2**: Type (String, Email, Boolean, etc.)
- **Row 3**: Data Source (Final Forms, Additional Info, Manual, Formula, etc.)
- **Row 4**: Additional Note (description/instructions)
- **Row 5**: Repeat Column Name (for pivot table compatibility)

### 🔄 Expected Workflow

#### Adding New Columns:
1. **Manually add column** in the Google Sheet with proper metadata in rows 1-5
2. **Update the script** to include the new column definition in `rosterColumns` array
3. **Test the script** - validation will ensure everything matches

#### Modifying Existing Columns:
1. **Edit metadata directly** in the Google Sheet (rows 1-5)
2. **Update script definitions** in `Code.gs` to match your changes
3. **Run the script** - validation will catch any mismatches

#### Script Validation:
- ✅ **Validates** that each column defined in the script exists in the sheet
- ✅ **Compares** metadata between sheet and script definitions
- 🚨 **Throws detailed error** if there are mismatches
- 🔒 **Never overwrites** metadata - sheet always wins

#### Error Resolution:
When validation fails, you have two options:
1. **Update the script** - modify `Code.gs` to match what's in the sheet
2. **Fix the sheet** - adjust metadata rows to match script expectations

### 🛡️ Protection Features

- **Metadata rows (1-5) are never modified** by the script
- **Manual/Formula columns** are preserved during data clearing
- **Sheet controls all column definitions** - script follows sheet design
- **Validation ensures consistency** between sheet and script

## Commit Convention

This project uses **Conventional Commits**. Format: `type(scope): description`

Examples:
- `feat(roster): add menu option to find missing emails`
- `fix(mailing-list): correct formula range to start at row 3`
- `refactor(clear): eliminate code duplication`