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
const SCRIPT_VERSION = '8';

// Constants
const FIRST_DATA_ROW = 6; // First row containing actual student data (after metadata rows 1-5)

// Configuration
const CONFIG = {
  finalForms: {
    fileId: '1p4cX6RmO0abXHdniSMnvxFtPYJFauCVj',
    sheetName: 'Final Forms'
  },
  additionalInfo: {
    spreadsheetId: '1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs',
    sheetName: 'Additional Info',
    rangeToImport: 'Form Responses 1!A:Z'
  },
  mailingList: {
    fileId: '1OZO3lo-WIdOp5piegWxVyR-R9PyO2ZoU',
    sheetName: 'Mailing List'
  },
  roster: {
    sheetName: 'Roster'
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
  
  // Create custom menu
  createCustomMenu();
  
  SpreadsheetApp.flush();
  
  console.log('Roster generated successfully!');
  SpreadsheetApp.getUi().alert('Roster Generated!', 'The roster has been created with metadata rows and all data mappings.\n\nNote: Columns are matched by header name, so you can safely reorder columns and regenerate.', SpreadsheetApp.getUi().ButtonSet.OK);
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
      name: 'First Name',
      type: 'String',
      source: 'FinalForms First Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!D:D,ROW()-4),"")`
    },
    {
      name: 'Last Name',
      type: 'String',
      source: 'FinalForms Last Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!E:E,ROW()-4),"")`
    },
    {
      name: 'Student SPS Email',
      type: 'Email Address',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is seattleschools.org',
      formula: `=IFERROR(IF(REGEXMATCH(INDEX('Final Forms'!F:F,ROW()-4),"@seattleschools\\.org"),INDEX('Final Forms'!F:F,ROW()-4),""),"")`
    },
    {
      name: 'Student Personal Email',
      type: 'Email',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is not seattleschools.org and the email address is not used as a FinalForms parent email',
      formula: `=IFERROR(IF(AND(NOT(REGEXMATCH(INDEX('Final Forms'!F:F,ROW()-4),"@seattleschools\\.org")),INDEX('Final Forms'!F:F,ROW()-4)<>INDEX('Final Forms'!AO:AO,ROW()-4),INDEX('Final Forms'!F:F,ROW()-4)<>INDEX('Final Forms'!AU:AU,ROW()-4)),INDEX('Final Forms'!F:F,ROW()-4),""),"")`
    },
    {
      name: 'Student Personal Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(G6,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Are All Forms Parent Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Parent Signed',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!P:P,ROW()-4),FALSE)`
    },
    {
      name: 'Are All Forms Student Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Student Signed',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!Q:Q,ROW()-4),FALSE)`
    },
    {
      name: 'Physical Cleared',
      type: 'Boolean',
      source: 'FinalForms Physical Cleared',
      note: '',
      formula: `=IFERROR(IF(INDEX('Final Forms'!AB:AB,ROW()-4)="Cleared",TRUE,FALSE),FALSE)`
    },
    {
      name: 'Gender',
      type: 'Enum',
      source: 'FinalForms Gender',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!U:U,ROW()-4),"")`
    },
    {
      name: 'Grade',
      type: 'Number',
      source: 'FinalForms Grade',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!W:W,ROW()-4),"")`
    },
    {
      name: 'Date of Birth',
      type: 'Date',
      source: 'FinalForms Date of Birth',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!X:X,ROW()-4),"")`
    },
    {
      name: 'Parent 1 First Name',
      type: 'String',
      source: 'FinalForms Parent 1 First Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AM:AM,ROW()-4),"")`
    },
    {
      name: 'Parent 1 Last Name',
      type: 'String',
      source: 'FinalForms Parent 1 Last Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AN:AN,ROW()-4),"")`
    },
    {
      name: 'Parent 1 Email',
      type: 'Email',
      source: 'FinalForms Parent 1 Email',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AO:AO,ROW()-4),"")`
    },
    {
      name: 'Parent 1 Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(O6,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Parent 2 First Name',
      type: 'String',
      source: 'FinalForms Parent 2 First Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AS:AS,ROW()-4),"")`
    },
    {
      name: 'Parent 2 Last Name',
      type: 'String',
      source: 'FinalForms Parent 2 Last Name',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AT:AT,ROW()-4),"")`
    },
    {
      name: 'Parent 2 Email',
      type: 'Email',
      source: 'FinalForms Parent 2 Email',
      note: '',
      formula: `=IFERROR(INDEX('Final Forms'!AU:AU,ROW()-4),"")`
    },
    {
      name: 'Parent 2 Email On Mailing List?',
      type: 'Enum',
      source: 'MailingList Email address',
      note: 'Returns "not a member", "invited", "member", etc. based on Group status column',
      formula: `=IFERROR(VLOOKUP(S6,'Mailing List'!$A$3:$C,3,FALSE),"not a member")`
    },
    {
      name: 'Additional Info Questionnaire Filled Out?',
      type: 'Boolean',
      source: '',
      note: 'True if was able to find a match for the player in Final Forms and the Additional Info form. \\nFalse otherwise.',
      formula: `=IF(COUNTIF('Additional Info'!B:B,TRIM(A6&" "&C6))>0,TRUE,FALSE)`
    },
    {
      name: 'Player Pronouns (select all that apply)',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IFERROR(INDEX('Additional Info'!C:C,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"")`
    },
    {
      name: 'Player Gender Identification',
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
  
  // Process each column definition
  rosterColumns.forEach((col) => {
    let colNum = columnMap.get(col.name);
    
    // If column doesn't exist, find the next available column
    if (!colNum) {
      colNum = maxCols + 1;
      // Expand sheet if needed
      if (colNum > rosterSheet.getMaxColumns()) {
        rosterSheet.insertColumnsAfter(rosterSheet.getMaxColumns(), 1);
      }
    }
    
    // Set metadata for this column
    // Row 1: Column Name
    rosterSheet.getRange(1, colNum).setValue(col.name);
    // Row 2: Type
    rosterSheet.getRange(2, colNum).setValue(col.type);
    // Row 3: Data Source
    rosterSheet.getRange(3, colNum).setValue(col.source);
    // Row 4: Additional Note
    rosterSheet.getRange(4, colNum).setValue(col.note);
    // Row 5: Repeat Column Name (for pivot tables)
    rosterSheet.getRange(5, colNum).setValue(col.name);
    
    // Now we need to update formulas to reference the correct columns dynamically
    // Create a formula that adapts to the current column positions
    let formula = col.formula;
    
    // Replace column references with dynamic lookups
    // For formulas that reference other columns (like A6, C6, E6, etc.)
    // we need to find those columns' current positions
    
    if (formula.includes('A6') || formula.includes('C6')) {
      // Find current positions of First Name and Last Name columns
      const firstNameCol = columnMap.get('First Name') || 1;
      const lastNameCol = columnMap.get('Last Name') || 3;
      
      // Convert column number to letter
      const firstNameLetter = getColumnLetter(firstNameCol);
      const lastNameLetter = getColumnLetter(lastNameCol);
      
      // Replace A6 with actual First Name column and C6 with Last Name column
      formula = formula.replace(/A6/g, firstNameLetter + '6');
      formula = formula.replace(/C6/g, lastNameLetter + '6');
    }
    
    if (formula.includes('E6')) {
      // Find Student Personal Email column
      const emailCol = columnMap.get('Student Personal Email') || 5;
      const emailLetter = getColumnLetter(emailCol);
      formula = formula.replace(/E6/g, emailLetter + '6');
    }
    
    if (formula.includes('O6')) {
      // Find Parent 1 Email column
      const parent1EmailCol = columnMap.get('Parent 1 Email') || 15;
      const parent1EmailLetter = getColumnLetter(parent1EmailCol);
      formula = formula.replace(/O6/g, parent1EmailLetter + '6');
    }
    
    if (formula.includes('S6')) {
      // Find Parent 2 Email column
      const parent2EmailCol = columnMap.get('Parent 2 Email') || 19;
      const parent2EmailLetter = getColumnLetter(parent2EmailCol);
      formula = formula.replace(/S6/g, parent2EmailLetter + '6');
    }
    
    // Set formula for FIRST_DATA_ROW
    rosterSheet.getRange(FIRST_DATA_ROW, colNum).setFormula(formula);
    
    // Copy formula down to row 200
    const sourceRange = rosterSheet.getRange(FIRST_DATA_ROW, colNum, 1, 1);
    const targetRange = rosterSheet.getRange(FIRST_DATA_ROW + 1, colNum, 194, 1);
    sourceRange.copyTo(targetRange);
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
    .addItem('📈 Show Statistics', 'showStatistics')
    .addItem('🔍 Find Emails Not on Mailing List', 'findMissingEmails')
    .addItem('👥 Parents Not Members of Mailing List', 'findPendingParents')
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
    const file = DriveApp.getFileById(CONFIG.finalForms.fileId);
    const csvData = file.getBlob().getDataAsString();
    const csvArray = Utilities.parseCsv(csvData);
    
    sheet.clear();
    if (csvArray.length > 0) {
      sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
    }
    
    console.log(`Updated Final Forms: ${csvArray.length - 1} students`);
  } catch (e) {
    console.error('Error updating Final Forms:', e);
    SpreadsheetApp.getUi().alert('Error', 'Could not update Final Forms data. Check the file ID.', SpreadsheetApp.getUi().ButtonSet.OK);
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
    const file = DriveApp.getFileById(CONFIG.mailingList.fileId);
    const csvData = file.getBlob().getDataAsString();
    const csvArray = Utilities.parseCsv(csvData);
    
    sheet.clear();
    if (csvArray.length > 0) {
      sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
    }
    
    console.log(`Updated Mailing List: ${csvArray.length - 1} emails`);
  } catch (e) {
    console.error('Error updating Mailing List:', e);
    SpreadsheetApp.getUi().alert('Error', 'Could not update Mailing List data. Check the file ID.', SpreadsheetApp.getUi().ButtonSet.OK);
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
  const firstNameCol = columnMap.get('First Name') || 1;
  const formsParentSignedCol = columnMap.get('Are All Forms Parent Signed');
  const formsStudentSignedCol = columnMap.get('Are All Forms Student Signed');
  const physicalClearedCol = columnMap.get('Physical Cleared');
  const additionalInfoCol = columnMap.get('Additional Info Questionnaire Filled Out?');
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
 * Run on spreadsheet open to create menu
 */
function onOpen() {
  createCustomMenu();
}
