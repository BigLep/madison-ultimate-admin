/**
 * Sheet Builder Utilities
 * Shared functions for building custom sheets with XLOOKUP formulas
 */

/**
 * Copy Full Name column from roster to a new sheet
 * @param {Sheet} targetSheet - The sheet to copy Full Name to
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {number} startRow - The row to start copying data (defaults to 2)
 * @return {Object} Object with fullNameColIndex and rowCount
 */
function copyFullNameColumn(targetSheet, rosterSheet, startRow = 2) {
  return copyFullNameColumnToColumn(targetSheet, rosterSheet, startRow, 1);
}

/**
 * Copy Full Name column from roster to a specific column in new sheet
 * @param {Sheet} targetSheet - The sheet to copy Full Name to
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {number} startRow - The row to start copying data
 * @param {number} targetColumn - The column to copy Full Name to (1-based)
 * @return {Object} Object with fullNameColIndex and rowCount
 */
function copyFullNameColumnToColumn(targetSheet, rosterSheet, startRow, targetColumn) {
  const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const fullNameColIndex = rosterHeaderRow.findIndex(name => name === CONFIG.columns.fullName);
  
  if (fullNameColIndex === -1) {
    throw new Error(`${CONFIG.columns.fullName} column not found in roster sheet`);
  }
  
  const rosterDataRange = rosterSheet.getRange(FIRST_DATA_ROW, fullNameColIndex + 1, rosterSheet.getLastRow() - FIRST_DATA_ROW + 1, 1);
  const fullNameValues = rosterDataRange.getValues();
  
  const nonEmptyFullNames = fullNameValues.filter(row => row[0] && row[0].toString().trim() !== '');
  
  if (nonEmptyFullNames.length === 0) {
    throw new Error('No student data found in roster');
  }
  
  targetSheet.getRange(startRow, targetColumn, nonEmptyFullNames.length, 1).setValues(nonEmptyFullNames);
  
  return {
    fullNameColIndex: fullNameColIndex,
    rowCount: nonEmptyFullNames.length
  };
}

/**
 * Create XLOOKUP formulas for columns from roster sheet
 * @param {Sheet} targetSheet - The sheet to add formulas to
 * @param {Array} columnNames - Array of column names to create formulas for
 * @param {number} startColumn - The column index to start adding formulas (1-based)
 * @param {number} rowCount - Number of rows to copy formulas down
 * @param {string} lookupSheetName - Name of the sheet to lookup from (defaults to CONFIG.roster.sheetName)
 */
function createRosterXlookupFormulas(targetSheet, columnNames, startColumn, rowCount, lookupSheetName = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(lookupSheetName || CONFIG.roster.sheetName);
  const sourceHeaderRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  const fullNameColIndex = sourceHeaderRow.findIndex(name => name === CONFIG.columns.fullName);
  if (fullNameColIndex === -1) {
    throw new Error(`${CONFIG.columns.fullName} column not found in source sheet`);
  }
  const fullNameColumnLetter = getColumnLetter(fullNameColIndex + 1);
  
  columnNames.forEach((columnName, index) => {
    const targetColumn = startColumn + index;
    
    const sourceColumnIndex = sourceHeaderRow.findIndex(name => name === columnName);
    if (sourceColumnIndex === -1) {
      console.warn(`Column "${columnName}" not found in ${lookupSheetName || CONFIG.roster.sheetName}`);
      return;
    }
    
    const sourceColumnLetter = getColumnLetter(sourceColumnIndex + 1);
    const sheetName = lookupSheetName || CONFIG.roster.sheetName;
    
    const formula = `=IFERROR(XLOOKUP(A2,'${sheetName}'!${fullNameColumnLetter}:${fullNameColumnLetter},'${sheetName}'!${sourceColumnLetter}:${sourceColumnLetter}),"")`;
    
    targetSheet.getRange(2, targetColumn).setFormula(formula);
    
    if (rowCount > 1) {
      const sourceRange = targetSheet.getRange(2, targetColumn, 1, 1);
      const targetRange = targetSheet.getRange(3, targetColumn, rowCount - 1, 1);
      sourceRange.copyTo(targetRange);
    }
  });
}

/**
 * Copy column formatting from source sheet to target sheet
 * @param {Sheet} targetSheet - The sheet to apply formatting to
 * @param {Sheet} sourceSheet - The source sheet to copy formatting from
 * @param {Array} headers - Array of column headers in target sheet
 * @param {Array} sourceHeaderRow - Header row from source sheet
 */
function copyColumnFormatting(targetSheet, sourceSheet, headers, sourceHeaderRow) {
  headers.forEach((columnName, newColumnIndex) => {
    const sourceColumnIndex = sourceHeaderRow.findIndex(name => name === columnName);
    
    if (sourceColumnIndex === -1) {
      console.warn(`Column "${columnName}" not found in source sheet for formatting`);
      return;
    }
    
    const newColumn = newColumnIndex + 1;
    const sourceColumn = sourceColumnIndex + 1;
    
    try {
      const sourceColumnWidth = sourceSheet.getColumnWidth(sourceColumn);
      targetSheet.setColumnWidth(newColumn, sourceColumnWidth);
      
      const sourceFormatCell = sourceSheet.getRange(FIRST_DATA_ROW, sourceColumn);
      const newFormatCell = targetSheet.getRange(2, newColumn);
      
      const numberFormat = sourceFormatCell.getNumberFormat();
      if (numberFormat) {
        const newColumnRange = targetSheet.getRange(2, newColumn, targetSheet.getMaxRows() - 1, 1);
        newColumnRange.setNumberFormat(numberFormat);
      }
      
      const textWrapping = sourceFormatCell.getWrap();
      const newColumnRange = targetSheet.getRange(2, newColumn, targetSheet.getMaxRows() - 1, 1);
      newColumnRange.setWrap(textWrapping);
      
      const horizontalAlignment = sourceFormatCell.getHorizontalAlignment();
      newColumnRange.setHorizontalAlignment(horizontalAlignment);
      
      const verticalAlignment = sourceFormatCell.getVerticalAlignment();
      newColumnRange.setVerticalAlignment(verticalAlignment);
      
      const fontFamily = sourceFormatCell.getFontFamily();
      const fontSize = sourceFormatCell.getFontSize();
      newColumnRange.setFontFamily(fontFamily);
      newColumnRange.setFontSize(fontSize);
      
      console.log(`✅ Copied formatting for column "${columnName}" (width: ${sourceColumnWidth}px)`);
      
    } catch (error) {
      console.warn(`Could not copy formatting for column "${columnName}":`, error);
    }
  });
}

/**
 * Copy conditional formatting from source sheet to target sheet
 * @param {Sheet} targetSheet - The sheet to apply conditional formatting to
 * @param {Sheet} sourceSheet - The source sheet to copy formatting from
 * @param {number} totalRows - Total number of rows in target sheet
 * @param {number} totalColumns - Total number of columns in target sheet
 */
function copyConditionalFormatting(targetSheet, sourceSheet, totalRows, totalColumns) {
  try {
    const sourceRules = sourceSheet.getConditionalFormatRules();
    
    if (sourceRules.length === 0) {
      console.log('No conditional formatting rules found in source sheet');
      return;
    }
    
    console.log(`Found ${sourceRules.length} conditional formatting rules in source sheet`);
    
    const entireSheetRange = targetSheet.getRange(1, 1, totalRows, totalColumns);
    const newRules = [];
    
    sourceRules.forEach((rule, ruleIndex) => {
      try {
        const newRule = rule.copy().setRanges([entireSheetRange]);
        newRules.push(newRule);
        console.log(`✅ Applied conditional formatting rule ${ruleIndex + 1} to target sheet`);
      } catch (ruleError) {
        console.warn(`Could not copy conditional formatting rule ${ruleIndex + 1}:`, ruleError);
      }
    });
    
    if (newRules.length > 0) {
      targetSheet.setConditionalFormatRules(newRules);
      console.log(`✅ Applied ${newRules.length} conditional formatting rules to target sheet`);
    }
    
  } catch (error) {
    console.warn('Could not copy conditional formatting:', error);
  }
}

/**
 * Copy data validation from source sheet to target sheet for specific columns
 * @param {Sheet} targetSheet - The sheet to apply data validation to
 * @param {Sheet} sourceSheet - The source sheet to copy validation from
 * @param {Array} columnMappings - Array of {sourceColumn, targetColumn} objects
 * @param {number} rowCount - Number of rows to apply validation to
 */
function copyDataValidation(targetSheet, sourceSheet, columnMappings, rowCount) {
  const sourceHeaderRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  columnMappings.forEach(mapping => {
    const sourceColIndex = sourceHeaderRow.indexOf(mapping.sourceColumn) + 1;
    if (sourceColIndex === 0) {
      console.warn(`Source column "${mapping.sourceColumn}" not found`);
      return;
    }
    
    const sourceCell = sourceSheet.getRange(FIRST_DATA_ROW, sourceColIndex);
    const validation = sourceCell.getDataValidation();
    
    if (!validation) {
      console.log(`No data validation found for column "${mapping.sourceColumn}"`);
      return;
    }
    
    const targetRange = targetSheet.getRange(2, mapping.targetColumn, rowCount, 1);
    targetRange.setDataValidation(validation);
    
    console.log(`✅ Copied data validation from "${mapping.sourceColumn}" to column ${mapping.targetColumn}`);
  });
}

/**
 * Style header row with standard formatting
 * @param {Sheet} sheet - The sheet to style
 * @param {number} columnCount - Number of columns in header
 */
function styleHeaderRow(sheet, columnCount) {
  const headerRange = sheet.getRange(1, 1, 1, columnCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
}