/**
 * Format Spruce Up Module
 * Quickly applies professional formatting to any sheet
 */

/**
 * Main function to apply professional formatting to the active sheet
 * Called from the menu
 */
function formatSpruceUp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  if (!activeSheet) {
    SpreadsheetApp.getUi().alert('Error', 'No active sheet found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const sheetName = activeSheet.getName();
  console.log(`🎨 Starting Format Spruce Up for sheet: "${sheetName}"`);
  
  try {
    applySpruceUpFormatting(activeSheet);
    
    console.log(`✅ Format Spruce Up complete for "${sheetName}"`);
    SpreadsheetApp.getUi().alert(
      'Format Spruce Up Complete!',
      `Successfully applied professional formatting to "${sheetName}":\n\n✓ Alternating row colors\n✓ Data filtering enabled\n✓ Vertical center alignment`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error in Format Spruce Up:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to apply formatting: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Apply all spruce up formatting to the given sheet
 * @param {Sheet} sheet - The sheet to format
 */
function applySpruceUpFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    throw new Error('Sheet appears to be empty - no formatting applied');
  }
  
  console.log(`📊 Sheet dimensions: ${lastRow} rows × ${lastCol} columns`);
  
  // 1. Apply alternating colors for the full sheet
  console.log('🌈 Applying alternating row colors...');
  applyAlternatingColors(sheet, lastRow, lastCol);
  
  // 2. Turn on data filtering (treating first row as header)
  console.log('🔍 Enabling data filtering...');
  enableDataFiltering(sheet, lastRow, lastCol);
  
  // 3. Center cells vertically
  console.log('📐 Setting vertical center alignment...');
  setCenterVerticalAlignment(sheet, lastRow, lastCol);
  
  console.log('✅ All formatting applied successfully');
}

/**
 * Apply alternating row colors to the sheet
 * @param {Sheet} sheet - The sheet to format
 * @param {number} lastRow - Last row with data
 * @param {number} lastCol - Last column with data
 */
function applyAlternatingColors(sheet, lastRow, lastCol) {
  try {
    // Remove any existing banding first
    const existingBanding = sheet.getBandings();
    if (existingBanding.length > 0) {
      console.log('Removing existing alternating colors...');
      existingBanding.forEach(banding => banding.remove());
    }
    
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    
    // Apply alternating row colors with header
    dataRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY, // Use light grey theme
      true, // Show header
      false // Don't show footer
    );
    
    console.log('✅ Alternating colors applied');
    
  } catch (error) {
    console.warn('Could not apply alternating colors:', error);
    throw new Error('Failed to apply alternating colors');
  }
}

/**
 * Enable data filtering on the sheet
 * @param {Sheet} sheet - The sheet to format  
 * @param {number} lastRow - Last row with data
 * @param {number} lastCol - Last column with data
 */
function enableDataFiltering(sheet, lastRow, lastCol) {
  try {
    // Remove any existing filters first
    const existingFilter = sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
      console.log('Removed existing filter');
    }
    
    // Create new filter for the entire data range
    const filterRange = sheet.getRange(1, 1, lastRow, lastCol);
    filterRange.createFilter();
    
    console.log('✅ Data filtering enabled');
    
  } catch (error) {
    console.warn('Could not enable data filtering:', error);
    throw new Error('Failed to enable data filtering');
  }
}

/**
 * Set vertical center alignment for all cells
 * @param {Sheet} sheet - The sheet to format
 * @param {number} lastRow - Last row with data  
 * @param {number} lastCol - Last column with data
 */
function setCenterVerticalAlignment(sheet, lastRow, lastCol) {
  try {
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    dataRange.setVerticalAlignment('middle');
    
    console.log('✅ Vertical center alignment applied');
    
  } catch (error) {
    console.warn('Could not set vertical alignment:', error);
    throw new Error('Failed to set vertical center alignment');
  }
}