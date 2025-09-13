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
    .setHeight(700) // Increased height to prevent button cutoff
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
            max-height: 250px; 
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
          .progress-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.9);
            z-index: 1000;
          }
          .progress-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            padding: 30px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #4285f4;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .progress-text {
            font-size: 16px;
            font-weight: bold;
            color: #333;
            margin-bottom: 5px;
          }
          .progress-detail {
            font-size: 14px;
            color: #666;
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
        
        <!-- Progress Overlay -->
        <div class="progress-overlay" id="progressOverlay">
          <div class="progress-content">
            <div class="spinner"></div>
            <div class="progress-text" id="progressText">Creating Sheet...</div>
            <div class="progress-detail" id="progressDetail">Please wait while we build your custom sheet</div>
          </div>
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
            
            // Show progress overlay
            showProgress('Creating Sheet...', 'Preparing your custom sheet');
            
            // Call server-side function
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .processSheetCreationWithProgress(sheetName, selectedColumns);
          }
          
          function showProgress(title, detail) {
            document.getElementById('progressText').textContent = title;
            document.getElementById('progressDetail').textContent = detail;
            document.getElementById('progressOverlay').style.display = 'block';
          }
          
          function hideProgress() {
            document.getElementById('progressOverlay').style.display = 'none';
          }
          
          function updateProgress(title, detail) {
            document.getElementById('progressText').textContent = title;
            document.getElementById('progressDetail').textContent = detail;
          }
          
          function onSuccess(message) {
            hideProgress();
            google.script.host.close();
          }
          
          function onFailure(error) {
            hideProgress();
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Server-side function called from HTML dialog with progress updates
 * Processes the sheet creation request
 */
function processSheetCreationWithProgress(sheetName, selectedColumns) {
  // This function provides better console logging for progress tracking
  console.log('🏗️ Starting sheet creation process...');
  console.log(`   Sheet name: "${sheetName}"`);
  console.log(`   Selected columns: ${selectedColumns.length}`);
  
  return processSheetCreation(sheetName, selectedColumns);
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
    console.log('📄 Creating new sheet...');
    const newSheet = ss.insertSheet(sheetName);
    
    // Find Full Name column in roster
    const headerRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    const fullNameColIndex = headerRow.findIndex(name => name === 'Full Name');
    
    if (fullNameColIndex === -1) {
      throw new Error('Full Name column not found in roster sheet');
    }
    
    // Set up headers: Full Name + selected columns
    console.log('📋 Setting up headers...');
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
    console.log('👥 Copying student names...');
    const newSheetDataStartRow = 2;
    newSheet.getRange(newSheetDataStartRow, 1, nonEmptyFullNames.length, 1).setValues(nonEmptyFullNames);
    
    // Create XLOOKUP formulas for each selected column
    console.log('🔗 Creating XLOOKUP formulas...');
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
    
    // Copy column formatting from roster sheet
    console.log('🎨 Copying column formatting...');
    copyColumnFormattingFromRoster(newSheet, rosterSheet, headers, headerRow);
    
    // Apply alternating colors to the entire data range
    console.log('🌈 Applying alternating colors...');
    const totalRows = nonEmptyFullNames.length + 1; // +1 for header row
    const dataRange = newSheet.getRange(1, 1, totalRows, headers.length);
    
    // Set up alternating row colors with header row
    dataRange.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY, // Use light grey theme
      true, // Show header
      true  // Show footer (not used but required parameter)
    );
    
    // Ensure header row styling is preserved (after banding)
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // Copy conditional formatting from roster sheet
    console.log('🎨 Copying conditional formatting...');
    copyConditionalFormattingFromRoster(newSheet, rosterSheet, headers, headerRow, totalRows);
    
    // Enable filtering on the data range
    console.log('🔍 Enabling filtering...');
    const filterRange = newSheet.getRange(1, 1, totalRows, headers.length);
    filterRange.createFilter();
    
    // Activate the new sheet
    console.log('🎯 Activating new sheet...');
    newSheet.activate();
    
    console.log(`✅ Sheet creation complete! "${sheetName}" with ${selectedColumns.length} columns and ${nonEmptyFullNames.length} students`);
    
    // Don't show alert here as it's handled by the HTML dialog success handler
    
  } catch (error) {
    console.error('Error creating custom sheet:', error);
    throw error; // Let the HTML dialog handle the error display
  }
}

/**
 * Copy column formatting from roster sheet to new custom sheet
 * Includes column widths, number formats, text wrapping, and alignment
 */
function copyColumnFormattingFromRoster(newSheet, rosterSheet, headers, rosterHeaderRow) {
  headers.forEach((columnName, newColumnIndex) => {
    // Find the column in the roster sheet
    const rosterColumnIndex = rosterHeaderRow.findIndex(name => name === columnName);
    
    if (rosterColumnIndex === -1) {
      console.warn(`Column "${columnName}" not found in roster sheet for formatting`);
      return;
    }
    
    const newColumn = newColumnIndex + 1; // Convert to 1-based
    const rosterColumn = rosterColumnIndex + 1; // Convert to 1-based
    
    try {
      // Copy column width
      const rosterColumnWidth = rosterSheet.getColumnWidth(rosterColumn);
      newSheet.setColumnWidth(newColumn, rosterColumnWidth);
      
      // Copy formatting from a data cell in the roster (row 6 = first data row)
      const rosterFormatCell = rosterSheet.getRange(FIRST_DATA_ROW, rosterColumn);
      const newFormatCell = newSheet.getRange(2, newColumn); // Row 2 = first data row in new sheet
      
      // Copy number format
      const numberFormat = rosterFormatCell.getNumberFormat();
      if (numberFormat) {
        const newColumnRange = newSheet.getRange(2, newColumn, newSheet.getMaxRows() - 1, 1);
        newColumnRange.setNumberFormat(numberFormat);
      }
      
      // Copy text wrapping
      const textWrapping = rosterFormatCell.getWrap();
      const newColumnRange = newSheet.getRange(2, newColumn, newSheet.getMaxRows() - 1, 1);
      newColumnRange.setWrap(textWrapping);
      
      // Copy horizontal alignment
      const horizontalAlignment = rosterFormatCell.getHorizontalAlignment();
      newColumnRange.setHorizontalAlignment(horizontalAlignment);
      
      // Copy vertical alignment
      const verticalAlignment = rosterFormatCell.getVerticalAlignment();
      newColumnRange.setVerticalAlignment(verticalAlignment);
      
      // Copy font family and size (but not color/weight as that might interfere with banding)
      const fontFamily = rosterFormatCell.getFontFamily();
      const fontSize = rosterFormatCell.getFontSize();
      newColumnRange.setFontFamily(fontFamily);
      newColumnRange.setFontSize(fontSize);
      
      console.log(`✅ Copied formatting for column "${columnName}" (width: ${rosterColumnWidth}px)`);
      
    } catch (error) {
      console.warn(`Could not copy formatting for column "${columnName}":`, error);
    }
  });
}

/**
 * Copy conditional formatting rules from roster sheet to new custom sheet
 * Applies all rules to the entire new sheet range (simple approach)
 */
function copyConditionalFormattingFromRoster(newSheet, rosterSheet, headers, rosterHeaderRow, totalRows) {
  try {
    // Get all conditional formatting rules from the roster sheet
    const rosterRules = rosterSheet.getConditionalFormatRules();
    
    if (rosterRules.length === 0) {
      console.log('No conditional formatting rules found in roster sheet');
      return;
    }
    
    console.log(`Found ${rosterRules.length} conditional formatting rules in roster sheet`);
    
    // Apply all rules to the entire new sheet range
    const entireSheetRange = newSheet.getRange(1, 1, totalRows, headers.length);
    const newRules = [];
    
    rosterRules.forEach((rule, ruleIndex) => {
      try {
        // Clone the rule and apply it to the entire new sheet
        const newRule = rule.copy().setRanges([entireSheetRange]);
        newRules.push(newRule);
        
        console.log(`✅ Applied conditional formatting rule ${ruleIndex + 1} to entire new sheet`);
        
      } catch (ruleError) {
        console.warn(`Could not copy conditional formatting rule ${ruleIndex + 1}:`, ruleError);
      }
    });
    
    // Set all rules on the new sheet
    if (newRules.length > 0) {
      newSheet.setConditionalFormatRules(newRules);
      console.log(`✅ Applied ${newRules.length} conditional formatting rules to new sheet`);
    }
    
  } catch (error) {
    console.warn('Could not copy conditional formatting:', error);
  }
}