# Roster Bugs Found and Fixes Needed

## Issues Identified

### 1. **Source Format Issue** (CRITICAL)
- **Problem**: Source values use wrong format like `"FinalForms First Name"` instead of `"Final Forms"`
- **Impact**: Column preservation logic fails - columns that should be cleared are preserved
- **Fix**: Update all source values to use consistent format:
  - `"Final Forms"` (not FinalForms)
  - `"Additional Info"` (not AdditionalInfoForm)
  - `"Mailing List"` (not MailingList)

### 2. **Column A Issue**
- **Problem**: Column A appears to be a header column but has source `"Data Source:"`
- **Impact**: This column gets cleared when it shouldn't
- **Fix**: Column A should either:
  - Have an empty source `""` to be preserved
  - Be removed from the columns array if it's just for display

### 3. **Mailing List Header Row**
- **Status**: Already handled correctly
- The formulas use `$A$3:$A` which skips the extra header rows

### 4. **Formula Issues**
- Using `ROW()-4` assumes 5 metadata rows, fragile if structure changes
- Could use MATCH to find header row dynamically

## Recommended Fixes

### Fix 1: Update all source values in madison_roster_final.gs

Change from:
```javascript
source: 'FinalForms First Name'
source: 'AdditionalInfoForm'
source: 'MailingList Email address'
```

To:
```javascript
source: 'Final Forms'
source: 'Additional Info'  
source: 'Mailing List'
```

### Fix 2: Fix Column A
Either remove it from the columns array or give it an empty source:
```javascript
source: ''  // This will preserve it
```

### Fix 3: Update the clearing logic
The current logic is correct but the source values are wrong. Once sources are fixed, the logic will work:
```javascript
if (sourceString === 'Manual' || sourceString === 'Formula' || sourceString === '') {
  // Preserve these columns
}
```

## Summary
The main bug is inconsistent source naming. The script defines sources like `"FinalForms First Name"` but the clearing logic expects `"Final Forms"`, `"Manual"`, `"Formula"`, or empty string. This mismatch causes the wrong columns to be preserved/cleared.