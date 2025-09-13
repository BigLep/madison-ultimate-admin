/**
 * Sheet Builder Module
 * Creates custom sheets with selected columns from the main roster
 */

/**
 * Main function to start the custom sheet building process
 * Called from the menu
 */
function buildCustomSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found. Please generate the roster first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get all column names from roster sheet
  const headerRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const columnNames = headerRow.filter(name => name && name.toString().trim() !== '');
  
  if (columnNames.length === 0) {
    SpreadsheetApp.getUi().alert('Error', 'No columns found in roster sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Show column selection dialog
  showColumnSelectionDialog(columnNames);
}

/**
 * Show dialog for selecting columns and entering sheet name
 * Uses HTML dialog with checkboxes for better UX
 */
function showColumnSelectionDialog(columnNames) {
  // Create HTML content for the dialog
  const htmlContent = createColumnSelectionHtml(columnNames);
  
  // Create HTML dialog
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(600)
    .setTitle('Build Custom Sheet');
  
  // Show the dialog
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sheet Builder');
}

/**
 * Create HTML content for column selection dialog
 */
function createColumnSelectionHtml(columnNames) {
  const checkboxes = columnNames
    .filter(name => name !== 'Full Name') // Exclude Full Name as it's always included
    .map(name => `
      <div class="checkbox-item">
        <label>
          <input type="checkbox" name="columns" value="${name}"> ${name}
        </label>
      </div>
    `).join('');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .form-group { margin-bottom: 15px; }
          label { font-weight: bold; display: block; margin-bottom: 5px; }
          input[type="text"] { width: 100%; padding: 8px; font-size: 14px; }
          .columns-container { 
            max-height: 300px; 
            overflow-y: auto; 
            border: 1px solid #ccc; 
            padding: 10px; 
            margin: 10px 0;
          }
          .checkbox-item { 
            margin: 5px 0; 
            padding: 2px 0;
          }
          .checkbox-item label { 
            font-weight: normal; 
            display: flex;
            align-items: center;
          }
          .checkbox-item input[type="checkbox"] { 
            margin-right: 8px; 
            margin-top: 0;
            margin-bottom: 0;
          }
          .buttons { 
            text-align: center; 
            margin-top: 20px; 
            padding-top: 15px;
            border-top: 1px solid #eee;
          }
          .btn { 
            padding: 10px 20px; 
            margin: 0 10px; 
            font-size: 14px;
            cursor: pointer;
          }
          .btn-primary { 
            background-color: #4285f4; 
            color: white; 
            border: none;
          }
          .btn-secondary { 
            background-color: #f8f9fa; 
            color: #333; 
            border: 1px solid #ccc;
          }
          .note { 
            font-size: 12px; 
            color: #666; 
            font-style: italic; 
            margin-top: 5px;
          }
          .auto-included {
            background-color: #e8f5e8;
            padding: 10px;
            margin: 10px 0;
            border-left: 4px solid #4caf50;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="sheetName">Sheet Name:</label>
          <input type="text" id="sheetName" placeholder="Enter custom sheet name">
          <div class="note">Choose a unique name for your new sheet</div>
        </div>
        
        <div class="auto-included">
          ✓ <strong>Full Name</strong> column will be automatically included (used for lookups)
        </div>
        
        <div class="form-group">
          <label>Select Additional Columns:</label>
          <div class="note">Check the boxes for columns you want to include in your custom sheet</div>
          <div class="columns-container">
            ${checkboxes}
          </div>
        </div>
        
        <div class="buttons">
          <button class="btn btn-primary" onclick="createSheet()">Create Sheet</button>
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
        
        <script>
          function createSheet() {
            const sheetName = document.getElementById('sheetName').value.trim();
            
            if (!sheetName) {
              alert('Please enter a sheet name');
              return;
            }
            
            const checkboxes = document.querySelectorAll('input[name="columns"]:checked');
            const selectedColumns = Array.from(checkboxes).map(cb => cb.value);
            
            // Call server-side function
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .processSheetCreation(sheetName, selectedColumns);
          }
          
          function onSuccess(message) {
            alert(message);
            google.script.host.close();
          }
          
          function onFailure(error) {
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Server-side function called from HTML dialog
 * Processes the sheet creation request
 */
function processSheetCreation(sheetName, selectedColumns) {
  // Validate inputs
  if (!sheetName || !sheetName.trim()) {
    throw new Error('Sheet name cannot be empty');
  }
  
  // Check if sheet name already exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(sheetName)) {
    throw new Error(`A sheet named "${sheetName}" already exists. Please choose a different name.`);
  }
  
  // Validate that we have at least one column selected or allow empty selection
  if (!Array.isArray(selectedColumns)) {
    selectedColumns = [];
  }
  
  // Create the sheet
  createCustomSheetWithColumns(sheetName, selectedColumns);
  
  // Return success message
  const message = selectedColumns.length === 0 
    ? `Created "${sheetName}" with only the Full Name column.`
    : `Created "${sheetName}" with Full Name + ${selectedColumns.length} additional columns.`;
    
  return message;
}

/**
 * Create the custom sheet with selected columns
 */
function createCustomSheetWithColumns(sheetName, selectedColumns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  try {
    // Create new sheet
    const newSheet = ss.insertSheet(sheetName);
    
    // Find Full Name column in roster
    const headerRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    const fullNameColIndex = headerRow.findIndex(name => name === 'Full Name');
    
    if (fullNameColIndex === -1) {
      throw new Error('Full Name column not found in roster sheet');
    }
    
    // Set up headers: Full Name + selected columns
    const headers = ['Full Name', ...selectedColumns];
    
    // Set headers with XLOOKUP formulas for column names (for debugging)
    const headerFormulas = headers.map((columnName, index) => {
      if (index === 0) {
        // First column is always Full Name
        return 'Full Name';
      } else {
        // Use XLOOKUP to get the column name from roster (helps spot bugs)
        const rosterColumnIndex = headerRow.findIndex(name => name === columnName);
        if (rosterColumnIndex === -1) {
          return columnName; // Fallback to static name if not found
        }
        const rosterColumnLetter = getColumnLetter(rosterColumnIndex + 1);
        return `=XLOOKUP("${columnName}",'${CONFIG.roster.sheetName}'!1:1,'${CONFIG.roster.sheetName}'!1:1)`;
      }
    });
    
    newSheet.getRange(1, 1, 1, headers.length).setValues([headerFormulas]);
    
    // Style the header row
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // Get Full Name values from roster (data rows only)
    const rosterDataRange = rosterSheet.getRange(FIRST_DATA_ROW, fullNameColIndex + 1, rosterSheet.getLastRow() - FIRST_DATA_ROW + 1, 1);
    const fullNameValues = rosterDataRange.getValues();
    
    // Filter out empty rows
    const nonEmptyFullNames = fullNameValues.filter(row => row[0] && row[0].toString().trim() !== '');
    
    if (nonEmptyFullNames.length === 0) {
      throw new Error('No student data found in roster');
    }
    
    // Copy Full Name values to new sheet
    const newSheetDataStartRow = 2;
    newSheet.getRange(newSheetDataStartRow, 1, nonEmptyFullNames.length, 1).setValues(nonEmptyFullNames);
    
    // Create XLOOKUP formulas for each selected column
    selectedColumns.forEach((columnName, columnIndex) => {
      const targetColumn = columnIndex + 2; // +2 because Full Name is column 1
      
      // Find the column in roster sheet
      const rosterColumnIndex = headerRow.findIndex(name => name === columnName);
      if (rosterColumnIndex === -1) {
        console.warn(`Column "${columnName}" not found in roster sheet`);
        return;
      }
      
      const rosterColumnLetter = getColumnLetter(rosterColumnIndex + 1);
      const fullNameColumnLetter = getColumnLetter(fullNameColIndex + 1);
      
      // Create XLOOKUP formula
      const formula = `=IFERROR(XLOOKUP(A2,'${CONFIG.roster.sheetName}'!${fullNameColumnLetter}:${fullNameColumnLetter},'${CONFIG.roster.sheetName}'!${rosterColumnLetter}:${rosterColumnLetter}),"")`;
      
      // Set formula in first data row
      newSheet.getRange(newSheetDataStartRow, targetColumn).setFormula(formula);
      
      // Copy formula down for all rows with Full Name data
      if (nonEmptyFullNames.length > 1) {
        const sourceRange = newSheet.getRange(newSheetDataStartRow, targetColumn, 1, 1);
        const targetRange = newSheet.getRange(newSheetDataStartRow + 1, targetColumn, nonEmptyFullNames.length - 1, 1);
        sourceRange.copyTo(targetRange);
      }
    });
    
    // Auto-resize columns
    newSheet.autoResizeColumns(1, headers.length);
    
    // Apply alternating colors to the entire data range
    const totalRows = nonEmptyFullNames.length + 1; // +1 for header row
    const dataRange = newSheet.getRange(1, 1, totalRows, headers.length);
    
    // Set up alternating row colors with header row
    dataRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY, // Use light grey theme
      true, // Show header
      true  // Show footer (not used but required parameter)
    );
    
    // Ensure header row styling is preserved
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // Activate the new sheet
    newSheet.activate();
    
    console.log(`✅ Created custom sheet "${sheetName}" with ${selectedColumns.length} columns and ${nonEmptyFullNames.length} students`);
    
    // Don't show alert here as it's handled by the HTML dialog success handler
    
  } catch (error) {
    console.error('Error creating custom sheet:', error);
    throw error; // Let the HTML dialog handle the error display
  }
}