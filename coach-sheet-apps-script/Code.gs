/**
 * Madison Middle School Ultimate Frisbee Roster Builder
 * Creates a roster with metadata rows and proper data mapping
 * 
 * To use:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete ALL existing code and paste this entire script
 * 4. Save (Ctrl+S or Cmd+S)
 * 5. Run 'generateRoster' function
 * 6. Grant permissions when prompted
 */

// Script Version - Increment this number when making changes  
const SCRIPT_VERSION = '80';

// Constants
const FIRST_DATA_ROW = 6; // First row containing actual student data (after metadata rows 1-5)

// Configuration
const CONFIG = {
  finalForms: {
    folderId: '1SnWCxDIn3FxJCvd1JcWyoeoOMscEsQcW', 
    sheetName: 'Final Forms'
  },
  additionalInfo: {
    spreadsheetId: '1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs',
    sheetName: 'Additional Info',
    rangeToImport: 'Form Responses 1!A:Z'
  },
  mailingList: {
    folderId: '1pAeQMEqiA9QdK9G5yRXsqgbNVzEU7R1E',
    sheetName: 'Mailing List'
  },
  roster: {
    sheetName: '📋 Roster'
  },
  attendance: {
    sheetName: '🏃 Attendance'
  },
  practiceInfo: {
    sheetName: '📍Practice Info'
  },
  gameInfo: {
    sheetName: '📍Game Info'
  },
  practiceAvailability: {
    sheetName: 'Practice Availability'
  },
  gameAvailability: {
    sheetName: 'Game Availability'
  }
};


/**
 * Main function to generate the roster
 */
function generateRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('Generating roster...');
  
  // Build the roster sheet with metadata and formulas
  buildRosterSheet(ss);
  
  // Populate Student IDs from Final Forms if available
  let studentIdResult = null;
  try {
    studentIdResult = populateStudentIdsFromFinalForms(ss);
  } catch (error) {
    console.warn('Could not populate Student IDs from Final Forms:', error);
    // Don't fail the entire generation if Student ID population fails
  }
  
  // Create custom menu
  createCustomMenu();
  
  SpreadsheetApp.flush();
  
  console.log('Roster generated successfully!');
  
  // Create success message with Student ID information
  let message = 'The roster has been created with metadata rows and formulas.';
  
  if (studentIdResult && studentIdResult.success) {
    message += `\n\n📊 Student IDs: ${studentIdResult.validCount} populated`;
    if (studentIdResult.emptyIdCount > 0) {
      message += `\n⚠️ Warning: ${studentIdResult.emptyIdCount} students in Final Forms have empty Student IDs and were skipped`;
    }
  } else if (studentIdResult) {
    message += `\n\n⚠️ Warning: Could not populate Student IDs - ${studentIdResult.emptyIdCount} of ${studentIdResult.totalCount} entries have empty Student IDs`;
  } else {
    message += '\n\n⚠️ Warning: Could not access Final Forms to populate Student IDs';
  }
  
  message += '\n\nNext Steps:\n1. Run "Update Mailing List" to populate email status columns\n2. Formulas will automatically pull data as students are added\n\nNote: Columns are matched by header name, so you can safely reorder columns and regenerate.';
  
  SpreadsheetApp.getUi().alert('Roster Generated!', message, SpreadsheetApp.getUi().ButtonSet.OK);
}


/**
 * Clear roster data while preserving metadata rows and Manual/Formula columns
 * Internal function that does the actual clearing
 */
function clearRosterDataInternal(rosterSheet) {
  const lastRow = rosterSheet.getMaxRows();
  const lastCol = rosterSheet.getMaxColumns();
  
  if (lastRow >= FIRST_DATA_ROW) {
    // Get the source row (row 3) to determine which columns to preserve
    const sourceRow = rosterSheet.getRange(3, 1, 1, lastCol).getValues()[0];
    
    // Clear data column by column, preserving Manual, Formula, and blank sources
    for (let col = 1; col <= lastCol; col++) {
      const columnSource = sourceRow[col - 1];
      const sourceString = columnSource ? columnSource.toString().trim() : '';
      
      // Skip columns with source "Manual", "Formula", or blank/empty
      if (sourceString === 'Manual' || sourceString === 'Formula' || sourceString === '') {
        console.log(`Preserving column ${col} (source: "${sourceString}")`);
        continue;
      }
      
      // Clear content for all other columns
      const columnRange = rosterSheet.getRange(FIRST_DATA_ROW, col, lastRow - 5, 1);
      columnRange.clearContent();
    }
  }
}

/**
 * Clear roster data while preserving metadata rows and Manual/Formula columns
 * User-facing function with UI alert
 */
function clearRosterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  clearRosterDataInternal(rosterSheet);
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Data Cleared', 'Roster data has been cleared. Metadata rows 1-5 and Manual/Formula columns preserved.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Validate that script metadata matches what's in the sheet
 * Throws an error if there are mismatches, indicating the script needs updating
 */
function validateRosterMetadata(rosterSheet, rosterColumns) {
  const maxCols = rosterSheet.getMaxColumns();
  const existingHeaders = rosterSheet.getRange(1, 1, 1, maxCols).getValues()[0];
  const existingTypes = rosterSheet.getRange(2, 1, 1, maxCols).getValues()[0];
  const existingSources = rosterSheet.getRange(3, 1, 1, maxCols).getValues()[0];
  const existingNotes = rosterSheet.getRange(4, 1, 1, maxCols).getValues()[0];
  
  const mismatches = [];
  
  // Check each column the script knows about
  rosterColumns.forEach(col => {
    const colIndex = existingHeaders.indexOf(col.name);
    
    if (colIndex === -1) {
      // Column is missing from sheet
      mismatches.push(`❌ MISSING COLUMN: "${col.name}" is defined in the script but not found in the sheet. Please add this column manually.`);
    } else {
      // Column exists in sheet - validate metadata
      const sheetType = existingTypes[colIndex]?.toString().trim() || '';
      const sheetSource = existingSources[colIndex]?.toString().trim() || '';
      const sheetNote = existingNotes[colIndex]?.toString().trim() || '';
      
      // Compare with script definitions (allow empty values to match)
      if (sheetType && sheetType !== col.type) {
        mismatches.push(`Column "${col.name}": Type mismatch. Sheet has "${sheetType}", script expects "${col.type}"`);
      }
      
      if (sheetSource && sheetSource !== col.source) {
        mismatches.push(`Column "${col.name}": Source mismatch. Sheet has "${sheetSource}", script expects "${col.source}"`);
      }
      
      if (sheetNote && col.note && sheetNote !== col.note) {
        mismatches.push(`Column "${col.name}": Note mismatch. Sheet has different note than script expects`);
      }
    }
  });
  
  if (mismatches.length > 0) {
    const errorMessage = `❌ METADATA MISMATCH DETECTED ❌\n\nThe sheet metadata doesn't match the script definitions. Please update the script or fix the sheet metadata:\n\n${mismatches.join('\n\n')}\n\n⚠️ The sheet is now the source of truth. Update the script to match what's in the sheet, or fix the sheet metadata to match the script.`;
    
    throw new Error(errorMessage);
  }
  
  console.log('✅ Metadata validation passed - sheet and script are in sync');
}

/**
 * Build the roster sheet with metadata and correct formulas
 */
function buildRosterSheet(spreadsheet) {
  const rosterSheet = spreadsheet.getSheetByName(CONFIG.roster.sheetName);
  
  // Get existing headers from row 1 to find column positions
  const maxCols = rosterSheet.getMaxColumns();
  const existingHeaders = rosterSheet.getRange(1, 1, 1, maxCols).getValues()[0];
  
  // Create a map of header names to column numbers
  const columnMap = new Map();
  existingHeaders.forEach((header, index) => {
    if (header && header !== '') {
      columnMap.set(header, index + 1);
    }
  });
  
  
  // Clear data rows (FIRST_DATA_ROW+) for defined columns only, preserve Manual/Formula columns
  clearRosterDataInternal(rosterSheet);
  
  // Define all roster columns with metadata
  const rosterColumns = [
    {
      name: 'StudentID',
      type: 'String',
      source: 'FinalForms StudentID',
      note: ''
      // No formula - this column gets populated with actual values during Final Forms import
    },
    {
      name: 'First Name',
      type: 'String',
      source: 'FinalForms First Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!D:D),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Last Name',
      type: 'String',
      source: 'FinalForms Last Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!E:E),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Student SPS Email',
      type: 'Email Address',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is seattleschools.org',
      formula: `=IFERROR(IF(REGEXMATCH(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),"@seattleschools\\.org"),XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),""),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Student Personal Email',
      type: 'Email',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is not seattleschools.org and the email address is not used as a FinalForms parent email',
      formula: `=IFERROR(IF(AND(NOT(REGEXMATCH(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),"@seattleschools\\.org")),XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F)<>XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AO:AO),XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F)<>XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AU:AU)),XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!F:F),""),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Student Personal Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(STUDENT_PERSONAL_EMAIL_COLUMN,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Are All Forms Parent Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Parent Signed',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!P:P),FALSE)` // Using XLOOKUP with Student ID
    },
    {
      name: 'Are All Forms Student Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Student Signed',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!Q:Q),FALSE)` // Using XLOOKUP with Student ID
    },
    {
      name: 'Physical Cleared',
      type: 'Boolean',
      source: 'FinalForms Physical Cleared',
      note: '',
      formula: `=IFERROR(IF(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AB:AB)="Cleared",TRUE,FALSE),FALSE)` // Using XLOOKUP with Student ID
    },
    {
      name: 'Gender',
      type: 'Enum',
      source: 'FinalForms Gender',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!U:U),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Grade',
      type: 'Number',
      source: 'FinalForms Grade',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!W:W),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Date of Birth',
      type: 'Date',
      source: 'FinalForms Date of Birth',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!X:X),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 1 First Name',
      type: 'String',
      source: 'FinalForms Parent 1 First Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AM:AM),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 1 Last Name',
      type: 'String',
      source: 'FinalForms Parent 1 Last Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AN:AN),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 1 Email',
      type: 'Email',
      source: 'FinalForms Parent 1 Email',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AO:AO),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 1 Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(PARENT1_EMAIL_COLUMN,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Parent 2 First Name',
      type: 'String',
      source: 'FinalForms Parent 2 First Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AS:AS),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 2 Last Name',
      type: 'String',
      source: 'FinalForms Parent 2 Last Name',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AT:AT),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 2 Email',
      type: 'Email',
      source: 'FinalForms Parent 2 Email',
      note: '',
      formula: `=IFERROR(XLOOKUP(A6,'Final Forms'!A:A,'Final Forms'!AU:AU),"")` // Using XLOOKUP with Student ID
    },
    {
      name: 'Parent 2 Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(PARENT2_EMAIL_COLUMN,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Player Pronouns (select all that apply)',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!C:C,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Supplied Gender Identification',
      type: 'Enum',
      source: 'AdditionalInfoForm',
      note: 'Set this to values of either "Gx" or "Bx".',
      formula: `=IFERROR(IF(REGEXMATCH(INDEX('Additional Info'!D:D,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"Girl|Gx"),"Gx",IF(REGEXMATCH(INDEX('Additional Info'!D:D,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"Boy|Bx"),"Bx","")),"")`
    },
    {
      name: 'Player Allergies',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!E:E,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Competing Sports and Activities',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!F:F,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Jersey Size',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!G:G,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Playing Experience',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!H:H,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Player hopes for the season',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!I:I,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Other Player Info',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!K:K,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Are you interested in helping coach?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!L:L,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: "Have you played or coached Ultimate before? What's been your experience?",
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!M:M,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: "Have you played or coached other team sports? What's been your experience?",
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!N:N,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Are you interested in helping in other ways?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!O:O,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Anything else you want to share?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!P:P,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    }
  ];
  
  // Validate that script metadata matches sheet metadata
  validateRosterMetadata(rosterSheet, rosterColumns);
  
  // Find Full Name column at runtime (critical join key to Additional Info)
  const fullNameCol = columnMap.get('Full Name');
  if (!fullNameCol) {
    throw new Error('❌ CRITICAL: "Full Name" column not found in roster sheet. This column is required as the join key to Additional Info data.');
  }
  const fullNameLetter = getColumnLetter(fullNameCol);
  
  
  // Process each column definition (validation ensures all exist)
  rosterColumns.forEach((col) => {
    const colNum = columnMap.get(col.name);
    // Note: validation above ensures this column exists
    
    // Now we need to update formulas to reference the correct columns dynamically
    // Create a formula that adapts to the current column positions
    let formula = col.formula;
    
    // Process formula if it exists
    if (!formula) {
      console.log(`Column "${col.name}" has no formula - column will be populated manually or via import`);
    } else {
      // Replace column references with dynamic lookups
      // For formulas that reference other columns (like A6, C6, E6, etc.)
      // we need to find those columns' current positions
      
      if (formula.includes('TRIM(A6&" "&C6)')) {
      // Replace concatenated First+Last names with Full Name column reference
      // This is the critical join key to Additional Info data
      formula = formula.replaceAll('TRIM(A6&" "&C6)', `${fullNameLetter}6`);
    } else if (formula.includes('XLOOKUP(A6,')) {
      // Handle XLOOKUP formulas that reference StudentID column
      const studentIdCol = columnMap.get('StudentID');
      if (!studentIdCol) {
        throw new Error('❌ CRITICAL: "StudentID" column not found. Required for Final Forms XLOOKUP formulas.');
      }
      
      // Convert column number to letter
      const studentIdLetter = getColumnLetter(studentIdCol);
      
      // Replace A6 with actual StudentID column in XLOOKUP formulas
      formula = formula.replace(/XLOOKUP\(A6,/g, `XLOOKUP(${studentIdLetter}6,`);
    } else if (formula.includes('A6') || formula.includes('C6')) {
      // Handle other column references (non-Additional Info lookups, non-XLOOKUP)
      const firstNameCol = columnMap.get('First Name');
      const lastNameCol = columnMap.get('Last Name');
      
      if (!firstNameCol || !lastNameCol) {
        throw new Error(`Cannot find required columns: First Name (${firstNameCol}), Last Name (${lastNameCol})`);
      }
      
      // Convert column number to letter
      const firstNameLetter = getColumnLetter(firstNameCol);
      const lastNameLetter = getColumnLetter(lastNameCol);
      
      // Replace A6 with actual First Name column and C6 with Last Name column
      formula = formula.replace(/A6/g, firstNameLetter + '6');
      formula = formula.replace(/C6/g, lastNameLetter + '6');
    }
    
    // Only replace E6 references for columns that actually use Student Personal Email
    // Don't replace E6 if it was created by our Full Name replacement above
    if (formula.includes('E6') && col.source && col.source.includes('Email')) {
      // Find Student Personal Email column
      const emailCol = columnMap.get('Student Personal Email');
      if (!emailCol) {
        throw new Error('Cannot find required column: Student Personal Email');
      }
      const emailLetter = getColumnLetter(emailCol);
      formula = formula.replace(/E6/g, emailLetter + '6');
    }
    
    if (formula.includes('O6')) {
      // Find Parent 1 Email column
      const parent1EmailCol = columnMap.get('Parent 1 Email');
      if (!parent1EmailCol) {
        throw new Error('Cannot find required column: Parent 1 Email');
      }
      const parent1EmailLetter = getColumnLetter(parent1EmailCol);
      formula = formula.replace(/O6/g, parent1EmailLetter + '6');
    }
    
    if (formula.includes('S6')) {
      // Find Parent 2 Email column
      const parent2EmailCol = columnMap.get('Parent 2 Email');
      if (!parent2EmailCol) {
        throw new Error('Cannot find required column: Parent 2 Email');
      }
      const parent2EmailLetter = getColumnLetter(parent2EmailCol);
      formula = formula.replace(/S6/g, parent2EmailLetter + '6');
    }
    
    // Replace placeholder column references with dynamic lookups
    if (formula.includes('STUDENT_PERSONAL_EMAIL_COLUMN')) {
      const studentEmailCol = columnMap.get('Student Personal Email');
      if (!studentEmailCol) {
        throw new Error('Cannot find required column: Student Personal Email');
      }
      const studentEmailLetter = getColumnLetter(studentEmailCol);
      formula = formula.replace(/STUDENT_PERSONAL_EMAIL_COLUMN/g, `${studentEmailLetter}6`);
    }
    
    if (formula.includes('PARENT1_EMAIL_COLUMN')) {
      const parent1EmailCol = columnMap.get('Parent 1 Email');
      if (!parent1EmailCol) {
        throw new Error('Cannot find required column: Parent 1 Email');
      }
      const parent1EmailLetter = getColumnLetter(parent1EmailCol);
      formula = formula.replace(/PARENT1_EMAIL_COLUMN/g, `${parent1EmailLetter}6`);
    }
    
    if (formula.includes('PARENT2_EMAIL_COLUMN')) {
      const parent2EmailCol = columnMap.get('Parent 2 Email');
      if (!parent2EmailCol) {
        throw new Error('Cannot find required column: Parent 2 Email');
      }
      const parent2EmailLetter = getColumnLetter(parent2EmailCol);
      formula = formula.replace(/PARENT2_EMAIL_COLUMN/g, `${parent2EmailLetter}6`);
    }
    
      // Set formula for FIRST_DATA_ROW
      rosterSheet.getRange(FIRST_DATA_ROW, colNum).setFormula(formula);
      
      // Copy formula down to row 200
      const sourceRange = rosterSheet.getRange(FIRST_DATA_ROW, colNum, 1, 1);
      const targetRange = rosterSheet.getRange(FIRST_DATA_ROW + 1, colNum, 194, 1);
      sourceRange.copyTo(targetRange);
    } // End of formula processing else block
  });
  
  console.log('Roster sheet built with dynamic column mapping');
}

/**
 * Convert column number to letter(s)
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const modulo = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    columnNumber = Math.floor((columnNumber - modulo) / 26);
  }
  return letter;
}

/**
 * Create custom menu for easy access
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(`🥏 Madison Ultimate (v${SCRIPT_VERSION})`)
    .addItem('📝 Generate Fresh Roster', 'generateRoster')
    .addItem('🗑️ Clear Roster Data (Keep Metadata)', 'clearRosterData')
    .addSeparator()
    .addItem('🔄 Refresh All Data', 'refreshAllData')
    .addItem('📊 Update Final Forms', 'updateFinalForms')
    .addItem('📧 Update Mailing List', 'updateMailingList')
    .addSeparator()
    .addItem('🏗️ Build Custom Sheet', 'buildCustomSheet')
    .addItem('🎨 Format Spruce Up', 'formatSpruceUp')
    .addItem('🏃 Build Practice Availability', 'buildPracticeAvailability')
    .addItem('🎮 Build Game Availability', 'buildGameAvailability')
    .addItem('📋 Organize Sheets', 'organizeSheets')
    .addSeparator()
    .addItem('📈 Show Statistics', 'showStatistics')
    .addItem('🔍 Find Emails Not on Mailing List', 'findMissingEmails')
    .addItem('👥 Parents Not Members of Mailing List', 'findPendingParents')
    .addItem('📊 Analyze Additional Info Responses', 'analyzeAdditionalInfoResponses')
    .addItem('🔀 Full Name Diff', 'fullNameDiff')
    .addToUi();
}

/**
 * Refresh all data sources
 */
function refreshAllData() {
  updateFinalForms();
  updateMailingList();
  // Additional Info updates automatically via IMPORTRANGE
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Data Refreshed', 'Final Forms and Mailing List data have been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Helper function to find the most recent CSV file in a folder
 * @param {string} folderId - The Google Drive folder ID
 * @returns {File|null} - The most recent CSV file, or null if none found
 */
function findMostRecentCsvFile(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  let mostRecentFile = null;
  let mostRecentDate = new Date(0); // Start with epoch time
  
  // Find the most recent CSV file
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();
    
    // Only consider CSV files
    if (fileName.endsWith('.csv')) {
      const fileDate = file.getLastUpdated();
      if (fileDate > mostRecentDate) {
        mostRecentDate = fileDate;
        mostRecentFile = file;
      }
    }
  }
  
  return mostRecentFile;
}

/**
 * Helper function to calculate and display differences between old and new data
 * @param {Array} oldData - Previous data array
 * @param {Array} newData - New data array  
 * @param {string} dataType - Type of data (for logging)
 */
function reportDataDifferences(oldData, newData, dataType) {
  const oldCount = oldData ? oldData.length - 1 : 0; // Subtract 1 for header
  const newCount = newData.length - 1; // Subtract 1 for header
  
  const difference = newCount - oldCount;
  const diffSign = difference > 0 ? '+' : '';
  
  console.log(`📊 ${dataType} Import Summary:`);
  console.log(`   Previous: ${oldCount} rows`);
  console.log(`   New: ${newCount} rows`);
  console.log(`   Change: ${diffSign}${difference} rows`);
  
  // Show detailed row changes for debugging
  if (oldData && newData.length > 1) {
    console.log(`🔍 ${dataType} Row Changes:`);
    
    // Create maps using row key (first non-empty column) for comparison
    const oldRowMap = new Map();
    const newRowMap = new Map();
    
    // Build old data map (skip header row)
    if (oldData.length > 1) {
      for (let i = 1; i < oldData.length; i++) {
        const row = oldData[i];
        const key = getRowKey(row, i);
        oldRowMap.set(key, { data: row, index: i });
      }
    }
    
    // Build new data map (skip header row)
    for (let i = 1; i < newData.length; i++) {
      const row = newData[i];
      const key = getRowKey(row, i);
      newRowMap.set(key, { data: row, index: i });
    }
    
    const addedRows = [];
    const removedRows = [];
    const modifiedRows = [];
    
    // Find added and modified rows
    for (const [key, newRow] of newRowMap) {
      if (!oldRowMap.has(key)) {
        // New row
        addedRows.push({ key, data: newRow.data });
      } else {
        // Check if row content changed
        const oldRow = oldRowMap.get(key);
        if (JSON.stringify(oldRow.data) !== JSON.stringify(newRow.data)) {
          modifiedRows.push({
            key,
            oldData: oldRow.data,
            newData: newRow.data
          });
        }
      }
    }
    
    // Find removed rows
    for (const [key, oldRow] of oldRowMap) {
      if (!newRowMap.has(key)) {
        removedRows.push({ key, data: oldRow.data });
      }
    }
    
    // Log added rows
    if (addedRows.length > 0) {
      console.log(`   ➕ Added ${addedRows.length} rows:`);
      addedRows.forEach(({ key, data }) => {
        console.log(`      + ${key}`);
      });
    }
    
    // Log removed rows
    if (removedRows.length > 0) {
      console.log(`   ➖ Removed ${removedRows.length} rows:`);
      removedRows.forEach(({ key, data }) => {
        console.log(`      - ${key}`);
      });
    }
    
    // Log modified rows with before/after
    if (modifiedRows.length > 0) {
      console.log(`   ✏️ Modified ${modifiedRows.length} rows:`);
      modifiedRows.forEach(({ key, oldData, newData }) => {
        console.log(`      📝 ${key}:`);
        
        // Compare each column to show what changed
        const maxColumns = Math.max(oldData.length, newData.length);
        for (let col = 0; col < maxColumns; col++) {
          const oldValue = oldData[col] || '';
          const newValue = newData[col] || '';
          
          if (oldValue !== newValue) {
            console.log(`         Column ${col + 1}: "${oldValue}" → "${newValue}"`);
          }
        }
      });
    }
    
    if (addedRows.length === 0 && removedRows.length === 0 && modifiedRows.length === 0) {
      console.log(`   ✅ No changes detected - data is identical`);
    }
  } else if (!oldData) {
    console.log(`   🆕 Initial import - no previous data to compare`);
  }
  
  return {
    oldCount,
    newCount, 
    difference
  };
}

/**
 * Helper function to get a unique key for a row (used for row matching)
 * @param {Array} row - The row data array
 * @param {number} index - Row index as fallback
 * @returns {string} - Unique key for the row
 */
function getRowKey(row, index) {
  // Try to find a good identifier from the first few columns
  for (let i = 0; i < Math.min(3, row.length); i++) {
    const value = row[i];
    if (value && value.toString().trim()) {
      return value.toString().trim();
    }
  }
  
  // Fallback to row index if no good identifier found
  return `Row_${index}`;
}

/**
 * Populate Student IDs in roster from Final Forms sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function populateStudentIdsFromFinalForms(ss) {
  // Get Final Forms sheet data
  const finalFormsSheet = ss.getSheetByName(CONFIG.finalForms.sheetName);
  if (!finalFormsSheet) {
    console.warn('Final Forms sheet not found, skipping Student ID population');
    return;
  }
  
  const lastRow = finalFormsSheet.getLastRow();
  if (lastRow <= 1) {
    console.warn('No data found in Final Forms sheet, skipping Student ID population');
    return;
  }
  
  // Get all Final Forms data
  const finalFormsData = finalFormsSheet.getRange(1, 1, lastRow, finalFormsSheet.getLastColumn()).getValues();
  
  return updateRosterStudentIdsInternal(ss, finalFormsData);
}

/**
 * Internal helper function to update the roster Student ID column with values from Final Forms
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {Array} finalFormsData - The Final Forms CSV data array
 */
function updateRosterStudentIdsInternal(ss, finalFormsData) {
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  if (!rosterSheet) {
    console.warn('Roster sheet not found, skipping Student ID update');
    return;
  }
  
  // Check if StudentID column exists in roster
  const headerRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const studentIdColIndex = headerRow.findIndex(header => header === 'StudentID');
  
  if (studentIdColIndex === -1) {
    console.warn('StudentID column not found in roster, skipping update');
    return;
  }
  
  const studentIdCol = studentIdColIndex + 1; // Convert to 1-based column number
  
  // Extract Student IDs from Final Forms data (assuming column A contains Student IDs)
  if (finalFormsData.length <= 1) {
    console.warn('No Final Forms data to extract Student IDs from');
    return;
  }
  
  // Extract and filter Student IDs (skip header row, filter out empty IDs)
  const allStudentData = finalFormsData.slice(1); // Skip header row
  const validStudentIds = [];
  let emptyIdCount = 0;
  
  allStudentData.forEach((row, index) => {
    const studentId = row[0]; // Column A is Student ID
    if (studentId && studentId.toString().trim() !== '') {
      validStudentIds.push(studentId.toString().trim());
    } else {
      emptyIdCount++;
      console.warn(`Row ${index + 2} in Final Forms has empty Student ID`);
    }
  });
  
  if (validStudentIds.length === 0) {
    console.warn('No valid Student IDs found in Final Forms data');
    return {
      success: false,
      validCount: 0,
      emptyIdCount: emptyIdCount,
      totalCount: allStudentData.length
    };
  }
  
  // Update the roster Student ID column with valid values only
  const startRow = FIRST_DATA_ROW; // Start from first data row
  
  // Clear existing data in Student ID column (data rows only)
  const maxRows = rosterSheet.getMaxRows();
  if (maxRows >= startRow) {
    rosterSheet.getRange(startRow, studentIdCol, maxRows - startRow + 1, 1).clearContent();
  }
  
  // Write the valid Student IDs
  const range = rosterSheet.getRange(startRow, studentIdCol, validStudentIds.length, 1);
  const values = validStudentIds.map(id => [id]); // Convert to 2D array
  range.setValues(values);
  
  console.log(`✅ Updated ${validStudentIds.length} Student IDs in roster column ${studentIdCol}`);
  if (emptyIdCount > 0) {
    console.warn(`⚠️ Skipped ${emptyIdCount} entries with empty Student IDs`);
  }
  
  return {
    success: true,
    validCount: validStudentIds.length,
    emptyIdCount: emptyIdCount,
    totalCount: allStudentData.length
  };
}

/**
 * Update Final Forms data
 */
function updateFinalForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.finalForms.sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', 'Final Forms sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // Get existing data for comparison
    const lastRow = sheet.getLastRow();
    const oldData = lastRow > 0 ? sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues() : null;
    
    // Get the most recent CSV file from the Final Forms folder
    const mostRecentFile = findMostRecentCsvFile(CONFIG.finalForms.folderId);
    
    if (!mostRecentFile) {
      SpreadsheetApp.getUi().alert('Error', 'No CSV files found in the Final Forms folder.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Using most recent Final Forms file: ${mostRecentFile.getName()} (${mostRecentFile.getLastUpdated()})`);
    
    const csvData = mostRecentFile.getBlob().getDataAsString();
    const csvArray = Utilities.parseCsv(csvData);
    
    // Report differences
    const diff = reportDataDifferences(oldData, csvArray, 'Final Forms');
    
    sheet.clear();
    if (csvArray.length > 0) {
      sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
    }
    
    const fileName = mostRecentFile.getName();
    const studentCount = csvArray.length - 1; // Subtract 1 for header row
    
    console.log(`✅ Updated Final Forms from: ${fileName}`);
    
    SpreadsheetApp.getUi().alert('Final Forms Updated', 
      `Successfully imported ${studentCount} students from:\n${fileName}\n\nChange: ${diff.difference >= 0 ? '+' : ''}${diff.difference} students\n\nNote: Run "Generate Fresh Roster" to populate Student IDs and formulas in the roster.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (e) {
    console.error('Error updating Final Forms:', e);
    SpreadsheetApp.getUi().alert('Error', 'Could not update Final Forms data. Check the folder and file permissions.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Update Mailing List data
 */
function updateMailingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.mailingList.sheetName);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', 'Mailing List sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // Get existing data for comparison
    const lastRow = sheet.getLastRow();
    const oldData = lastRow > 0 ? sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues() : null;
    
    // Get the most recent CSV file from the mailing list folder
    const mostRecentFile = findMostRecentCsvFile(CONFIG.mailingList.folderId);
    
    if (!mostRecentFile) {
      SpreadsheetApp.getUi().alert('Error', 'No CSV files found in the mailing list folder.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Using most recent mailing list file: ${mostRecentFile.getName()} (${mostRecentFile.getLastUpdated()})`);
    
    const csvData = mostRecentFile.getBlob().getDataAsString();
    const csvArray = Utilities.parseCsv(csvData);
    
    // Report differences
    const diff = reportDataDifferences(oldData, csvArray, 'Mailing List');
    
    sheet.clear();
    if (csvArray.length > 0) {
      sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
    }
    
    const fileName = mostRecentFile.getName();
    const emailCount = csvArray.length - 1; // Subtract 1 for header row
    
    console.log(`✅ Updated Mailing List from: ${fileName}`);
    
    SpreadsheetApp.getUi().alert('Mailing List Updated!', 
      `Successfully imported ${emailCount} emails from:\n${fileName}\n\nChange: ${diff.difference >= 0 ? '+' : ''}${diff.difference} emails`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.error('Error updating Mailing List:', e);
    SpreadsheetApp.getUi().alert('Error', `Could not update Mailing List data:\n${e.toString()}\n\nCheck the folder ID in CONFIG.mailingList.folderId`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show statistics about the roster
 */
function showStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get column positions dynamically
  const headers = rosterSheet.getRange(1, 1, 1, rosterSheet.getMaxColumns()).getValues()[0];
  const columnMap = new Map();
  headers.forEach((header, index) => {
    if (header) columnMap.set(header, index + 1);
  });
  
  // Get column positions we need
  const firstNameCol = columnMap.get('First Name');
  if (!firstNameCol) {
    throw new Error('Cannot find required column: First Name');
  }
  const formsParentSignedCol = columnMap.get('Are All Forms Parent Signed');
  const formsStudentSignedCol = columnMap.get('Are All Forms Student Signed');
  const physicalClearedCol = columnMap.get('Physical Cleared');
  const parent1MailingCol = columnMap.get('Parent 1 Email On Mailing List?');
  const parent2MailingCol = columnMap.get('Parent 2 Email On Mailing List?');
  const gradeCol = columnMap.get('Grade');
  
  // Start counting from FIRST_DATA_ROW
  const firstDataRow = FIRST_DATA_ROW;
  const lastRow = rosterSheet.getLastRow();
  
  // Count non-empty student rows
  let totalStudents = 0;
  for (let i = firstDataRow; i <= lastRow; i++) {
    const firstName = rosterSheet.getRange(i, firstNameCol).getValue();
    if (firstName && firstName !== '') {
      totalStudents++;
    }
  }
  
  // Count various statistics
  let statsData = {
    formsParentSigned: 0,
    formsStudentSigned: 0,
    physicalCleared: 0,
    additionalInfo: 0,
    parent1OnList: 0,
    parent2OnList: 0,
    grades: {}
  };
  
  for (let i = firstDataRow; i < firstDataRow + totalStudents; i++) {
    // Check forms signed
    if (formsParentSignedCol && rosterSheet.getRange(i, formsParentSignedCol).getValue() === true) {
      statsData.formsParentSigned++;
    }
    if (formsStudentSignedCol && rosterSheet.getRange(i, formsStudentSignedCol).getValue() === true) {
      statsData.formsStudentSigned++;
    }
    
    // Check physical cleared
    if (physicalClearedCol && rosterSheet.getRange(i, physicalClearedCol).getValue() === true) {
      statsData.physicalCleared++;
    }
    
    // Check additional info
    if (additionalInfoCol && rosterSheet.getRange(i, additionalInfoCol).getValue() === true) {
      statsData.additionalInfo++;
    }
    
    // Check mailing list status
    if (parent1MailingCol) {
      const parent1Status = rosterSheet.getRange(i, parent1MailingCol).getValue();
      if (parent1Status === true) statsData.parent1OnList++;
    }
    if (parent2MailingCol) {
      const parent2Status = rosterSheet.getRange(i, parent2MailingCol).getValue();
      if (parent2Status === true) statsData.parent2OnList++;
    }
    
    // Count grade distribution
    if (gradeCol) {
      const grade = rosterSheet.getRange(i, gradeCol).getValue();
      if (grade) {
        statsData.grades[grade] = (statsData.grades[grade] || 0) + 1;
      }
    }
  }
  
  // Build statistics message
  let message = `📊 Roster Statistics\n\n`;
  message += `Total Students: ${totalStudents}\n\n`;
  
  message += `Forms Status:\n`;
  message += `  Parent Signed: ${statsData.formsParentSigned} (${Math.round(statsData.formsParentSigned/totalStudents*100)}%)\n`;
  message += `  Student Signed: ${statsData.formsStudentSigned} (${Math.round(statsData.formsStudentSigned/totalStudents*100)}%)\n`;
  message += `  Physical Cleared: ${statsData.physicalCleared} (${Math.round(statsData.physicalCleared/totalStudents*100)}%)\n\n`;
  
  message += `Additional Info Completed: ${statsData.additionalInfo} (${Math.round(statsData.additionalInfo/totalStudents*100)}%)\n\n`;
  
  message += `Mailing List Status:\n`;
  message += `  Parent 1 on List (Yes): ${statsData.parent1OnList}\n`;
  message += `  Parent 2 on List (Yes): ${statsData.parent2OnList}\n\n`;
  
  message += `Grade Distribution:\n`;
  for (const [grade, count] of Object.entries(statsData.grades).sort()) {
    message += `  Grade ${grade}: ${count} students\n`;
  }
  
  SpreadsheetApp.getUi().alert('Madison Ultimate Roster Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Find all email addresses in roster that are not on the mailing list
 * Excludes Seattle School email addresses
 */
function findMissingEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  const mailingSheet = ss.getSheetByName(CONFIG.mailingList.sheetName);
  
  if (!rosterSheet || !mailingSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find Roster or Mailing List sheets');
    return;
  }
  
  // Get all email addresses from mailing list (column A, starting from row 3)
  const mailingListData = mailingSheet.getRange('A3:A').getValues();
  const mailingListEmails = new Set(
    mailingListData
      .flat()
      .filter(email => email && email.toString().trim())
      .map(email => email.toString().toLowerCase().trim())
  );
  
  console.log(`Found ${mailingListEmails.size} emails in mailing list`);
  
  // Find email columns in roster (looking for columns with "Email" in header)
  const headers = rosterSheet.getRange('1:1').getValues()[0];
  const emailColumns = [];
  
  headers.forEach((header, index) => {
    if (header && header.toString().toLowerCase().includes('email')) {
      emailColumns.push(index + 1); // 1-based column index
      console.log(`Found email column: ${header} at column ${index + 1}`);
    }
  });
  
  // Collect all unique emails from roster
  const lastRow = rosterSheet.getLastRow();
  const uniqueRosterEmails = new Set();
  const missingEmails = [];
  
  // Process each email column
  emailColumns.forEach(colNum => {
    if (lastRow > 5) { // Skip metadata rows
      const emailData = rosterSheet.getRange(FIRST_DATA_ROW, colNum, lastRow - 5, 1).getValues();
      emailData.forEach((row, rowIndex) => {
        const email = row[0];
        if (email && email.toString().trim()) {
          const emailStr = email.toString().trim();
          const emailLower = emailStr.toLowerCase();
          
          // Skip Seattle School email addresses
          if (emailLower.includes('@seattleschools.org')) {
            console.log(`Skipping Seattle Schools email: ${emailStr}`);
            return;
          }
          
          // Check if this email is not in mailing list and not already added
          if (!mailingListEmails.has(emailLower) && !uniqueRosterEmails.has(emailLower)) {
            uniqueRosterEmails.add(emailLower);
            
            // Get student name from same row
            const firstName = rosterSheet.getRange(rowIndex + FIRST_DATA_ROW, 1).getValue();
            const lastName = rosterSheet.getRange(rowIndex + FIRST_DATA_ROW, 2).getValue();
            const columnName = headers[colNum - 1];
            
            missingEmails.push({
              email: emailStr,
              name: `${firstName} ${lastName}`.trim(),
              source: columnName,
              row: rowIndex + FIRST_DATA_ROW
            });
          }
        }
      });
    }
  });
  
  // Sort missing emails alphabetically
  missingEmails.sort((a, b) => a.email.localeCompare(b.email));
  
  // Display results
  if (missingEmails.length === 0) {
    SpreadsheetApp.getUi().alert(
      'All Emails on Mailing List',
      'All roster email addresses are already on the mailing list (Seattle Schools emails excluded).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    // Create a formatted message
    let message = `Found ${missingEmails.length} email addresses not on the mailing list:\n\n`;
    
    // Group by source column for better readability
    const bySource = {};
    missingEmails.forEach(item => {
      if (!bySource[item.source]) {
        bySource[item.source] = [];
      }
      bySource[item.source].push(item);
    });
    
    // Format the message
    Object.keys(bySource).sort().forEach(source => {
      message += `\n${source}:\n`;
      bySource[source].forEach(item => {
        message += `  • ${item.email} (${item.name})\n`;
      });
    });
    
    message += '\n\nYou can copy these addresses to add them to the mailing list.';
    
    // For easier copying, also create a comma-separated list
    const emailList = missingEmails.map(item => item.email).join(', ');
    message += `\n\nComma-separated list for easy copying:\n${emailList}`;
    
    // Show in a dialog (alert has size limits, so using custom HTML dialog for long lists)
    if (missingEmails.length > 10) {
      showMissingEmailsDialog(missingEmails, emailList);
    } else {
      SpreadsheetApp.getUi().alert(
        'Emails Not on Mailing List',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

/**
 * Show missing emails in a scrollable HTML dialog for long lists
 */
function showMissingEmailsDialog(missingEmails, emailList) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 10px; }
      h3 { color: #1a73e8; }
      .email-group { margin-bottom: 20px; }
      .email-item { margin: 5px 0; padding: 5px; background: #f8f9fa; }
      .copy-section { 
        margin-top: 20px; 
        padding: 10px; 
        background: #e8f0fe; 
        border-radius: 5px;
      }
      textarea { 
        width: 100%; 
        height: 100px; 
        margin-top: 10px;
        font-family: monospace;
      }
      .stats { color: #5f6368; margin-bottom: 15px; }
    </style>
    <div>
      <h3>Emails Not on Mailing List</h3>
      <div class="stats">Found ${missingEmails.length} email addresses (Seattle Schools excluded)</div>
      
      <div class="copy-section">
        <strong>All emails (comma-separated):</strong>
        <textarea readonly onclick="this.select()">${emailList}</textarea>
      </div>
      
      <h4>Detailed List by Source:</h4>
      ${Object.entries(
        missingEmails.reduce((acc, item) => {
          if (!acc[item.source]) acc[item.source] = [];
          acc[item.source].push(item);
          return acc;
        }, {})
      ).map(([source, items]) => `
        <div class="email-group">
          <strong>${source}:</strong>
          ${items.map(item => `
            <div class="email-item">
              ${item.email} - ${item.name} (Row ${item.row})
            </div>
          `).join('')}
        </div>
      `).join('')}
    </div>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Emails Not on Mailing List');
}

/**
 * Find all parents/caretakers who are not members of the mailing list
 * Shows those with any status other than "member" (includes "invited" and "not a member")
 */
function findPendingParents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find Roster sheet');
    return;
  }
  
  // Get column positions dynamically
  const headers = rosterSheet.getRange(1, 1, 1, rosterSheet.getMaxColumns()).getValues()[0];
  const columnMap = new Map();
  headers.forEach((header, index) => {
    if (header) columnMap.set(header, index + 1);
  });
  
  // Get the columns we need
  const firstNameCol = columnMap.get('First Name');
  const lastNameCol = columnMap.get('Last Name');
  const parent1FirstCol = columnMap.get('Parent 1 First Name');
  const parent1LastCol = columnMap.get('Parent 1 Last Name');
  const parent1EmailCol = columnMap.get('Parent 1 Email');
  const parent1StatusCol = columnMap.get('Parent 1 Email On Mailing List?');
  const parent2FirstCol = columnMap.get('Parent 2 First Name');
  const parent2LastCol = columnMap.get('Parent 2 Last Name');
  const parent2EmailCol = columnMap.get('Parent 2 Email');
  const parent2StatusCol = columnMap.get('Parent 2 Email On Mailing List?');
  
  if (!firstNameCol || !lastNameCol) {
    SpreadsheetApp.getUi().alert('Error: Could not find required columns');
    return;
  }
  
  const lastRow = rosterSheet.getLastRow();
  const pendingParents = [];
  const seenEmails = new Set(); // Avoid duplicates
  
  // Process each data row (starting from FIRST_DATA_ROW)
  for (let row = FIRST_DATA_ROW; row <= lastRow; row++) {
    const studentFirst = rosterSheet.getRange(row, firstNameCol).getValue();
    const studentLast = rosterSheet.getRange(row, lastNameCol).getValue();
    
    // Skip empty rows
    if (!studentFirst) continue;
    
    // Check Parent 1
    if (parent1StatusCol && parent1EmailCol && parent1FirstCol && parent1LastCol) {
      const status = rosterSheet.getRange(row, parent1StatusCol).getValue();
      const email = rosterSheet.getRange(row, parent1EmailCol).getValue();
      const firstName = rosterSheet.getRange(row, parent1FirstCol).getValue();
      const lastName = rosterSheet.getRange(row, parent1LastCol).getValue();
      
      if (status && status !== 'member' && email && firstName && !seenEmails.has(email.toLowerCase())) {
        seenEmails.add(email.toLowerCase());
        pendingParents.push({
          firstName: firstName.toString().trim(),
          lastName: lastName.toString().trim(),
          email: email.toString().trim(),
          status: status.toString(),
          student: `${studentFirst} ${studentLast}`.trim(),
          parentType: 'Parent 1'
        });
      }
    }
    
    // Check Parent 2
    if (parent2StatusCol && parent2EmailCol && parent2FirstCol && parent2LastCol) {
      const status = rosterSheet.getRange(row, parent2StatusCol).getValue();
      const email = rosterSheet.getRange(row, parent2EmailCol).getValue();
      const firstName = rosterSheet.getRange(row, parent2FirstCol).getValue();
      const lastName = rosterSheet.getRange(row, parent2LastCol).getValue();
      
      if (status && status !== 'member' && email && firstName && !seenEmails.has(email.toLowerCase())) {
        seenEmails.add(email.toLowerCase());
        pendingParents.push({
          firstName: firstName.toString().trim(),
          lastName: lastName.toString().trim(),
          email: email.toString().trim(),
          status: status.toString(),
          student: `${studentFirst} ${studentLast}`.trim(),
          parentType: 'Parent 2'
        });
      }
    }
  }
  
  // Sort by last name, then first name
  pendingParents.sort((a, b) => {
    const lastNameCompare = a.lastName.localeCompare(b.lastName);
    return lastNameCompare !== 0 ? lastNameCompare : a.firstName.localeCompare(b.firstName);
  });
  
  // Generate table data
  const tableData = pendingParents;
  
  if (pendingParents.length === 0) {
    SpreadsheetApp.getUi().alert(
      'All Parents Are Members',
      'All parents have "member" status on the mailing list.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    showPendingParentsDialog(pendingParents);
  }
}

/**
 * Show pending parents in a modal dialog with HTML table for easy copy/paste
 */
function showPendingParentsDialog(pendingParents) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      h3 { color: #1a73e8; margin-bottom: 5px; }
      .stats { color: #5f6368; margin-bottom: 20px; font-size: 14px; }
      
      .table-container { 
        background: white;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        overflow: hidden;
        margin-bottom: 20px;
      }
      
      table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
      }
      
      th {
        background: #f8f9fa;
        color: #202124;
        font-weight: 600;
        padding: 12px 16px;
        text-align: left;
        border-bottom: 2px solid #e8eaed;
      }
      
      td {
        padding: 10px 16px;
        border-bottom: 1px solid #e8eaed;
        vertical-align: top;
      }
      
      tr:hover {
        background: #f8f9fa;
      }
      
      .status-invited {
        background: #fff3cd;
        color: #856404;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
      }
      
      .status-not-member {
        background: #f8d7da;
        color: #721c24;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
      }
      
      .instructions {
        background: #e8f0fe;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 15px;
        font-size: 13px;
      }
      
      .copy-instructions {
        color: #5f6368;
        font-style: italic;
      }
    </style>
    <div>
      <h3>Parents Not Members of Mailing List</h3>
      <div class="stats">Found ${pendingParents.length} parents who are not members</div>
      
      <div class="instructions">
        <strong>Instructions:</strong> Select the table below and copy (Ctrl+C / Cmd+C) to paste into spreadsheets or emails.
        <br><span class="copy-instructions">The table will copy with proper formatting and can be pasted directly into Excel, Google Sheets, or email.</span>
      </div>
      
      <div class="table-container">
        <table id="parentTable">
          <thead>
            <tr>
              <th>First Name</th>
              <th>Last Name</th>
              <th>Email Address</th>
              <th>Status</th>
              <th>Student</th>
            </tr>
          </thead>
          <tbody>
            ${pendingParents.map(parent => `
              <tr>
                <td>${parent.firstName}</td>
                <td>${parent.lastName}</td>
                <td>${parent.email}</td>
                <td><span class="status-${parent.status.replace(' ', '-')}">${parent.status}</span></td>
                <td>${parent.student}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      
      <div style="margin-top: 15px; font-size: 12px; color: #5f6368;">
        <strong>Tip:</strong> You can select individual rows or the entire table and copy to paste elsewhere.
      </div>
    </div>
    
    <script>
      // Auto-select table when clicked for easy copying
      document.getElementById('parentTable').addEventListener('click', function() {
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(this);
        selection.removeAllRanges();
        selection.addRange(range);
      });
    </script>
  `)
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Parents Not Members of Mailing List');
}

/**
 * Analyze Additional Info responses for matches, suggestions, and potential duplicates
 * Creates a separate "Additional Info Analysis" sheet with the results
 *
 * This function is now implemented in AdditionalInfoAnalysis.gs
 */
// This function is implemented in AdditionalInfoAnalysis.gs

/**
 * Run on spreadsheet open to create menu
 */
function onOpen() {
  createCustomMenu();
}
