/**
 * Delete Empty Rows and Columns Utility
 * Removes empty rows and columns from the active sheet to clean up the workspace
 */

/**
 * Main function to delete empty rows and columns from the active sheet
 * Called from the menu
 */
function deleteEmptyRowsAndColumns() {
  console.log('🧹 Starting Delete Empty Rows and Columns...');
  
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    console.log(`📋 Processing sheet: "${sheetName}"`);
    
    // Get current sheet dimensions
    const maxRows = sheet.getMaxRows();
    const maxColumns = sheet.getMaxColumns();
    
    console.log(`📐 Current sheet size: ${maxRows} rows x ${maxColumns} columns`);
    
    // Find the actual data boundaries
    const lastRowWithData = findLastRowWithData(sheet);
    const lastColumnWithData = findLastColumnWithData(sheet);
    
    console.log(`📊 Data boundaries: row ${lastRowWithData}, column ${lastColumnWithData}`);
    
    // Calculate what needs to be deleted
    const rowsToDelete = maxRows - lastRowWithData;
    const columnsToDelete = maxColumns - lastColumnWithData;
    
    if (rowsToDelete <= 0 && columnsToDelete <= 0) {
      SpreadsheetApp.getUi().alert(
        'No Empty Rows or Columns',
        `Sheet "${sheetName}" has no empty rows or columns to delete.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Show confirmation dialog
    const message = buildConfirmationMessage(sheetName, rowsToDelete, columnsToDelete, lastRowWithData, lastColumnWithData);
    const response = SpreadsheetApp.getUi().alert(
      'Delete Empty Rows and Columns',
      message,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      console.log('❌ User cancelled deletion');
      return;
    }
    
    // Perform the deletion
    let deletionsPerformed = 0;
    
    // Delete empty rows first (do this before columns to avoid index issues)
    if (rowsToDelete > 0) {
      console.log(`🗑️ Deleting ${rowsToDelete} empty rows (from row ${lastRowWithData + 1} onwards)`);
      sheet.deleteRows(lastRowWithData + 1, rowsToDelete);
      deletionsPerformed++;
      console.log(`✅ Deleted ${rowsToDelete} empty rows`);
    }
    
    // Delete empty columns
    if (columnsToDelete > 0) {
      console.log(`🗑️ Deleting ${columnsToDelete} empty columns (from column ${lastColumnWithData + 1} onwards)`);
      sheet.deleteColumns(lastColumnWithData + 1, columnsToDelete);
      deletionsPerformed++;
      console.log(`✅ Deleted ${columnsToDelete} empty columns`);
    }
    
    // Show success message
    const successMessage = buildSuccessMessage(sheetName, rowsToDelete, columnsToDelete, sheet.getMaxRows(), sheet.getMaxColumns());
    SpreadsheetApp.getUi().alert('Cleanup Complete!', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log(`✅ Cleanup complete for sheet "${sheetName}"`);
    
  } catch (error) {
    console.error('Error deleting empty rows and columns:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to delete empty rows and columns: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Fast discovery pass to find approximate data boundaries
 * Starts at A1 and expands until hitting empty space, then tries 5 more cells
 * @param {Sheet} sheet - The sheet to analyze
 * @return {Object} Object with maxSearchRow and maxSearchColumn
 */
function fastDiscoveryPass(sheet) {
  console.log('🔍 Starting fast discovery pass from A1...');
  
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  
  let maxRowFound = 1;
  let maxColumnFound = 1;
  let emptyRowStreak = 0;
  let emptyColumnStreak = 0;
  const maxEmptyStreak = 5;
  
  // Expand row-wise (going down)
  for (let row = 1; row <= maxRows; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    const hasContent = cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '';
    
    if (hasContent) {
      maxRowFound = row;
      emptyRowStreak = 0;
    } else {
      emptyRowStreak++;
      if (emptyRowStreak > maxEmptyStreak) {
        break; // Stop searching after 5 consecutive empty cells
      }
    }
  }
  
  // Expand column-wise (going right)
  for (let col = 1; col <= maxColumns; col++) {
    const cellValue = sheet.getRange(1, col).getValue();
    const hasContent = cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '';
    
    if (hasContent) {
      maxColumnFound = col;
      emptyColumnStreak = 0;
    } else {
      emptyColumnStreak++;
      if (emptyColumnStreak > maxEmptyStreak) {
        break; // Stop searching after 5 consecutive empty cells
      }
    }
  }
  
  // Add some buffer to the discovered boundaries
  const searchRowLimit = Math.min(maxRows, maxRowFound + 20);
  const searchColumnLimit = Math.min(maxColumns, maxColumnFound + 20);
  
  console.log(`🔍 Fast discovery found data up to approximately row ${maxRowFound}, column ${maxColumnFound}`);
  console.log(`🔍 Will search thoroughly up to row ${searchRowLimit}, column ${searchColumnLimit}`);
  
  return {
    maxSearchRow: searchRowLimit,
    maxSearchColumn: searchColumnLimit
  };
}

/**
 * Find the last row that contains any data (optimized with discovery pass)
 * @param {Sheet} sheet - The sheet to analyze
 * @return {number} The last row number with data (1-based)
 */
function findLastRowWithData(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  
  if (maxRows === 0 || maxColumns === 0) {
    return 1; // Keep at least one row
  }
  
  // Fast discovery pass to limit search range
  const searchLimits = fastDiscoveryPass(sheet);
  const searchUpToRow = searchLimits.maxSearchRow;
  const searchUpToColumn = searchLimits.maxSearchColumn;
  
  console.log(`🔍 Thorough row search from row ${searchUpToRow} down to 1`);
  
  // Start from the discovered boundary and work our way up
  for (let row = searchUpToRow; row >= 1; row--) {
    const rowRange = sheet.getRange(row, 1, 1, searchUpToColumn);
    const rowValues = rowRange.getValues()[0];
    
    // Check if any cell in this row has content
    const hasContent = rowValues.some(cell => {
      if (cell === null || cell === undefined) return false;
      const cellString = cell.toString().trim();
      return cellString !== '';
    });
    
    if (hasContent) {
      console.log(`📍 Last row with data: ${row}`);
      return row;
    }
  }
  
  console.log('📍 No data found, keeping row 1');
  return 1; // Keep at least one row if no data found
}

/**
 * Find the last column that contains any data (optimized with discovery pass)
 * @param {Sheet} sheet - The sheet to analyze
 * @return {number} The last column number with data (1-based)
 */
function findLastColumnWithData(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  
  if (maxRows === 0 || maxColumns === 0) {
    return 1; // Keep at least one column
  }
  
  // Fast discovery pass to limit search range
  const searchLimits = fastDiscoveryPass(sheet);
  const searchUpToRow = searchLimits.maxSearchRow;
  const searchUpToColumn = searchLimits.maxSearchColumn;
  
  console.log(`🔍 Thorough column search from column ${searchUpToColumn} down to 1`);
  
  // Start from the discovered boundary and work our way left
  for (let col = searchUpToColumn; col >= 1; col--) {
    const columnRange = sheet.getRange(1, col, searchUpToRow, 1);
    const columnValues = columnRange.getValues();
    
    // Check if any cell in this column has content
    const hasContent = columnValues.some(row => {
      const cell = row[0];
      if (cell === null || cell === undefined) return false;
      const cellString = cell.toString().trim();
      return cellString !== '';
    });
    
    if (hasContent) {
      console.log(`📍 Last column with data: ${col}`);
      return col;
    }
  }
  
  console.log('📍 No data found, keeping column 1');
  return 1; // Keep at least one column if no data found
}

/**
 * Build the confirmation message for the user
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowsToDelete - Number of rows to delete
 * @param {number} columnsToDelete - Number of columns to delete
 * @param {number} lastRowWithData - Last row with data
 * @param {number} lastColumnWithData - Last column with data
 * @return {string} Confirmation message
 */
function buildConfirmationMessage(sheetName, rowsToDelete, columnsToDelete, lastRowWithData, lastColumnWithData) {
  let message = `Sheet: "${sheetName}"\n\n`;
  message += `Data found up to:\n`;
  message += `• Row ${lastRowWithData}\n`;
  message += `• Column ${lastColumnWithData} (${getColumnLetter(lastColumnWithData)})\n\n`;
  
  message += `This will delete:\n`;
  if (rowsToDelete > 0) {
    message += `• ${rowsToDelete} empty rows\n`;
  }
  if (columnsToDelete > 0) {
    message += `• ${columnsToDelete} empty columns\n`;
  }
  
  message += `\nProceed with cleanup?`;
  
  return message;
}

/**
 * Build the success message after deletion
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowsDeleted - Number of rows deleted
 * @param {number} columnsDeleted - Number of columns deleted
 * @param {number} newMaxRows - New max rows
 * @param {number} newMaxColumns - New max columns
 * @return {string} Success message
 */
function buildSuccessMessage(sheetName, rowsDeleted, columnsDeleted, newMaxRows, newMaxColumns) {
  let message = `Successfully cleaned up sheet "${sheetName}"!\n\n`;
  
  if (rowsDeleted > 0) {
    message += `✅ Deleted ${rowsDeleted} empty rows\n`;
  }
  if (columnsDeleted > 0) {
    message += `✅ Deleted ${columnsDeleted} empty columns\n`;
  }
  
  message += `\nNew sheet size: ${newMaxRows} rows x ${newMaxColumns} columns`;
  
  return message;
}

/**
 * Convert column number to letter (e.g., 1 = A, 27 = AA)
 * @param {number} columnNumber - 1-based column number
 * @return {string} Column letter(s)
 */
function getColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--; // Convert to 0-based
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}