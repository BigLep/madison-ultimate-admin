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

// Configuration
const CONFIG = {
  finalForms: {
    fileId: '1pWUIw2rM0MfNWnaC3Ltsz6Wj8_PGFHrH',
    sheetName: 'Final Forms'
  },
  additionalInfo: {
    spreadsheetId: '1f_PPULjdg-5q2Gi0cXvWvGz1RbwYmUtADChLqwsHuNs',
    sheetName: 'Additional Info',
    rangeToImport: 'Form Responses 1!A:Z'
  },
  mailingList: {
    fileId: '1n0h81l31lsGvvSPrZUT5SOuS6jXT4h6E',
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
 * Clear all formatting from the roster sheet
 */
function clearRosterFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Clear all formatting
  const fullRange = rosterSheet.getDataRange();
  fullRange.clearFormat();
  
  // Clear conditional formatting rules
  rosterSheet.clearConditionalFormatRules();
  
  // Reset row heights to default
  for (let i = 1; i <= rosterSheet.getMaxRows(); i++) {
    rosterSheet.setRowHeight(i, 21); // Default height
  }
  
  // Reset column widths to default
  for (let i = 1; i <= rosterSheet.getMaxColumns(); i++) {
    rosterSheet.setColumnWidth(i, 100); // Default width
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Formatting Cleared', 'All formatting has been removed from the Roster sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear roster data while preserving metadata rows
 */
function clearRosterData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Clear all data from row 6 onwards, preserving metadata rows 1-5
  const lastRow = rosterSheet.getMaxRows();
  const lastCol = rosterSheet.getMaxColumns();
  
  if (lastRow >= 6) {
    // Clear content and formulas from row 6 to the last row
    const dataRange = rosterSheet.getRange(6, 1, lastRow - 5, lastCol);
    dataRange.clearContent();
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Data Cleared', 'Roster data has been cleared. Metadata rows 1-5 preserved.', SpreadsheetApp.getUi().ButtonSet.OK);
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
  
  // Clear data rows (6+) if they exist
  const lastRow = rosterSheet.getMaxRows();
  if (lastRow >= 6) {
    const dataRange = rosterSheet.getRange(6, 1, lastRow - 5, maxCols);
    dataRange.clearContent();
  }
  
  // Define all roster columns with metadata
  const rosterColumns = [
    {
      name: 'First Name',
      type: 'String',
      source: 'FinalForms First Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!D:D,ROW()-4),""))`
    },
    {
      name: 'Preferred Name',
      type: 'String', 
      source: 'Maybe FinalForms or AdditionalInfoForm',
      note: 'If you get any sense that this person goes by a different name than their official first name, then go ahead and add it here. For example, "Pete" instead of "Peter".',
      formula: `=IF(ROW()<6,"","")` // Manual entry
    },
    {
      name: 'Last Name',
      type: 'String',
      source: 'FinalForms Last Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!E:E,ROW()-4),""))`
    },
    {
      name: 'Student SPS Email',
      type: 'Email Address',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is seattleschools.org',
      formula: `=IF(ROW()<6,"",IFERROR(IF(REGEXMATCH(INDEX('Final Forms'!F:F,ROW()-4),"@seattleschools\\.org"),INDEX('Final Forms'!F:F,ROW()-4),""),""))`
    },
    {
      name: 'Student Personal Email',
      type: 'Email',
      source: 'FinalForms Email',
      note: 'Only set this if the domain is not seattleschools.org and the email address is not used as a FinalForms parent email',
      formula: `=IF(ROW()<6,"",IFERROR(IF(AND(NOT(REGEXMATCH(INDEX('Final Forms'!F:F,ROW()-4),"@seattleschools\\.org")),INDEX('Final Forms'!F:F,ROW()-4)<>INDEX('Final Forms'!AO:AO,ROW()-4),INDEX('Final Forms'!F:F,ROW()-4)<>INDEX('Final Forms'!AU:AU,ROW()-4)),INDEX('Final Forms'!F:F,ROW()-4),""),""))`
    },
    {
      name: 'Student Personal Email On Mailing List?',
      type: 'Boolean',
      source: 'MailingList Email address and Posting permissions',
      note: 'TRUE if email is on mailing list with posting permissions allowed, FALSE otherwise',
      formula: `=IF(ROW()<6,"",IF(E6="",FALSE,IF(AND(COUNTIF('Mailing List'!A:A,E6)>0,INDEX('Mailing List'!F:F,MATCH(E6,'Mailing List'!A:A,0))="allowed"),TRUE,FALSE)))`
    },
    {
      name: 'Are All Forms Parent Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Parent Signed',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!P:P,ROW()-4),FALSE))`
    },
    {
      name: 'Are All Forms Student Signed',
      type: 'Boolean',
      source: 'FinalForms Are All Forms Student Signed',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!Q:Q,ROW()-4),FALSE))`
    },
    {
      name: 'Physical Cleared',
      type: 'Boolean',
      source: 'FinalForms Physical Cleared',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(IF(INDEX('Final Forms'!AB:AB,ROW()-4)="Cleared",TRUE,FALSE),FALSE))`
    },
    {
      name: 'Gender',
      type: 'Enum',
      source: 'FinalForms Gender',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!U:U,ROW()-4),""))`
    },
    {
      name: 'Grade',
      type: 'Number',
      source: 'FinalForms Grade',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!W:W,ROW()-4),""))`
    },
    {
      name: 'Date of Birth',
      type: 'Date',
      source: 'FinalForms Date of Birth',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!X:X,ROW()-4),""))`
    },
    {
      name: 'Parent 1 First Name',
      type: 'String',
      source: 'FinalForms Parent 1 First Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AM:AM,ROW()-4),""))`
    },
    {
      name: 'Parent 1 Last Name',
      type: 'String',
      source: 'FinalForms Parent 1 Last Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AN:AN,ROW()-4),""))`
    },
    {
      name: 'Parent 1 Email',
      type: 'Email',
      source: 'FinalForms Parent 1 Email',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AO:AO,ROW()-4),""))`
    },
    {
      name: 'Parent 1 Email On Mailing List?',
      type: 'Boolean',
      source: 'MailingList Email address and Posting permissions',
      note: 'TRUE if email is on mailing list with posting permissions allowed, FALSE otherwise',
      formula: `=IF(ROW()<6,"",IF(O6="",FALSE,IF(AND(COUNTIF('Mailing List'!A:A,O6)>0,INDEX('Mailing List'!F:F,MATCH(O6,'Mailing List'!A:A,0))="allowed"),TRUE,FALSE)))`
    },
    {
      name: 'Parent 2 First Name',
      type: 'String',
      source: 'FinalForms Parent 2 First Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AS:AS,ROW()-4),""))`
    },
    {
      name: 'Parent 2 Last Name',
      type: 'String',
      source: 'FinalForms Parent 2 Last Name',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AT:AT,ROW()-4),""))`
    },
    {
      name: 'Parent 2 Email',
      type: 'Email',
      source: 'FinalForms Parent 2 Email',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Final Forms'!AU:AU,ROW()-4),""))`
    },
    {
      name: 'Parent 2 Email On Mailing List?',
      type: 'Boolean',
      source: 'MailingList Email address and Posting permissions',
      note: 'TRUE if email is on mailing list with posting permissions allowed, FALSE otherwise',
      formula: `=IF(ROW()<6,"",IF(S6="",FALSE,IF(AND(COUNTIF('Mailing List'!A:A,S6)>0,INDEX('Mailing List'!F:F,MATCH(S6,'Mailing List'!A:A,0))="allowed"),TRUE,FALSE)))`
    },
    {
      name: 'Additional Info Questionnaire Filled Out?',
      type: 'Boolean',
      source: '',
      note: 'True if was able to find a match for the player in Final Forms and the Additional Info form. \\nFalse otherwise.',
      formula: `=IF(ROW()<6,"",IF(COUNTIF('Additional Info'!B:B,TRIM(A6&" "&C6))>0,TRUE,FALSE))`
    },
    {
      name: 'Player Pronouns (select all that apply)',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!C:C,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Player Gender Identification',
      type: 'Enum',
      source: 'AdditionalInfoForm',
      note: 'Set this to values of either "Gx" or "Bx".',
      formula: `=IF(ROW()<6,"",IFERROR(IF(REGEXMATCH(INDEX('Additional Info'!D:D,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"Girl|Gx"),"Gx",IF(REGEXMATCH(INDEX('Additional Info'!D:D,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),"Boy|Bx"),"Bx","")),""))`
    },
    {
      name: 'Player Allergies',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!E:E,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Competing Sports and Activities',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!F:F,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Jersey Size',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!G:G,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Playing Experience',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!H:H,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Player hopes for the season',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!I:I,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Other Player Info',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!K:K,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Are you interested in helping coach?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!L:L,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: "Have you played or coached Ultimate before? What's been your experience?",
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!M:M,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: "Have you played or coached other team sports? What's been your experience?",
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!N:N,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Are you interested in helping in other ways?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!O:O,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
    },
    {
      name: 'Anything else you want to share?',
      type: 'String',
      source: 'AdditionalInfoForm',
      note: '',
      formula: `=IF(ROW()<6,"",IFERROR(INDEX('Additional Info'!P:P,MATCH(TRIM(A6&" "&C6),'Additional Info'!B:B,0)),""))`
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
    
    // Set formula for row 6
    rosterSheet.getRange(6, colNum).setFormula(formula);
    
    // Copy formula down to row 200
    const sourceRange = rosterSheet.getRange(6, colNum, 1, 1);
    const targetRange = rosterSheet.getRange(7, colNum, 194, 1);
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
  ui.createMenu('🥏 Madison Ultimate')
    .addItem('📝 Generate Fresh Roster', 'generateRoster')
    .addItem('🗑️ Clear Roster Data (Keep Metadata)', 'clearRosterData')
    .addItem('🧹 Clear All Formatting', 'clearRosterFormatting')
    .addSeparator()
    .addItem('🔄 Refresh All Data', 'refreshAllData')
    .addItem('📊 Update Final Forms', 'updateFinalForms')
    .addItem('📧 Update Mailing List', 'updateMailingList')
    .addSeparator()
    .addItem('📈 Show Statistics', 'showStatistics')
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
  
  // Start counting from row 6 (first data row)
  const firstDataRow = 6;
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
 * Run on spreadsheet open to create menu
 */
function onOpen() {
  createCustomMenu();
}
