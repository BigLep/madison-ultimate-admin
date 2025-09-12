# Madison Ultimate Admin

Tools and scripts for managing the Madison Middle School Ultimate Frisbee team roster and communications.

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