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
  const rosterColumns = headerRow.filter(name => name && name.toString().trim() !== '');
  
  if (rosterColumns.length === 0) {
    SpreadsheetApp.getUi().alert('Error', 'No columns found in roster sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Discover native Attendance columns (columns that aren't copied from Roster)
  const attendanceColumns = discoverAttendanceColumns();
  
  // Show column selection dialog with both Roster and Attendance columns
  showColumnSelectionDialog(rosterColumns, attendanceColumns);
}

/**
 * Get the Attendance sheet from the current spreadsheet
 * @return {Sheet|null} The Attendance sheet or null if not accessible
 */
function getAttendanceSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName('Attendance');
    
    if (!attendanceSheet) {
      console.warn('Sheet named "Attendance" not found in the current spreadsheet');
      return null;
    }
    
    return attendanceSheet;
  } catch (error) {
    console.warn('Could not access Attendance sheet:', error);
    return null;
  }
}

/**
 * Discover native Attendance columns (columns that aren't copied from Roster)
 * @return {Array} Array of attendance column names that are native to attendance sheet
 */
function discoverAttendanceColumns() {
  const attendanceSheet = getAttendanceSheet();
  if (!attendanceSheet) {
    return [];
  }
  
  // Get header row (row 1)
  const headerRow = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
  const allAttendanceColumns = headerRow.filter(name => name && name.toString().trim() !== '');
  
  if (allAttendanceColumns.length === 0) {
    console.warn('No columns found in Attendance sheet');
    return [];
  }
  
  // Get row 2 formulas to identify which columns are copied from Roster
  const formulaRow = attendanceSheet.getRange(2, 1, 1, attendanceSheet.getLastColumn()).getFormulas()[0];
  
  const nativeAttendanceColumns = [];
  
  console.log(`Analyzing ${allAttendanceColumns.length} attendance columns...`);
  
  // Check each column - if row 2 does NOT have an XLOOKUP formula, it's a native attendance column
  allAttendanceColumns.forEach((columnName, index) => {
    const formula = formulaRow[index] || '';
    
    // Check if formula contains XLOOKUP (indicating it's copied from another sheet)
    const hasXlookup = formula.toUpperCase().includes('XLOOKUP');
    
    console.log(`Column ${index + 1} "${columnName}": formula="${formula}", hasXlookup=${hasXlookup}`);
    
    // Only include columns that don't have XLOOKUP formulas in row 2 and aren't Full Name
    if (!hasXlookup && columnName.trim() !== '' && columnName !== 'Full Name') {
      nativeAttendanceColumns.push(columnName);
    }
  });
  
  console.log(`Found ${nativeAttendanceColumns.length} native Attendance columns: ${nativeAttendanceColumns.join(', ')}`);
  return nativeAttendanceColumns;
}

/**
 * Show dialog for selecting columns and entering sheet name
 * Uses HTML dialog with checkboxes for better UX
 */
function showColumnSelectionDialog(rosterColumns, attendanceColumns) {
  // Create HTML content for the dialog with both column types
  const htmlContent = createColumnSelectionHtml(rosterColumns, attendanceColumns);
  
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
function createColumnSelectionHtml(rosterColumns, attendanceColumns) {
  // Create checkboxes for Roster columns (exclude Full Name as it's always included)
  const rosterCheckboxes = rosterColumns
    .filter(name => name !== 'Full Name')
    .map(name => `
      <div class="checkbox-item">
        <label>
          <input type="checkbox" name="columns" value="roster:${name}" data-source="roster"> ${name}
        </label>
      </div>
    `).join('');

  // Create checkboxes for native Attendance columns
  const attendanceCheckboxes = attendanceColumns
    .map(name => `
      <div class="checkbox-item">
        <label>
          <input type="checkbox" name="columns" value="attendance:${name}" data-source="attendance"> ${name}
        </label>
      </div>
    `).join('');

  // Combine sections
  const allCheckboxes = `
    ${rosterColumns.length > 1 ? `
      <h4 style="margin-top: 0; color: #4285f4;">📊 Roster Columns</h4>
      ${rosterCheckboxes}
    ` : ''}
    ${attendanceColumns.length > 0 ? `
      <h4 style="margin-top: 15px; color: #34a853;">📅 Attendance Columns</h4>
      ${attendanceCheckboxes}
    ` : ''}
  `;
  
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
            ${allCheckboxes}
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
          // Load saved selections when dialog opens
          window.onload = function() {
            restoreLastSelections();
          };
          
          function restoreLastSelections() {
            try {
              const savedSelections = localStorage.getItem('sheetBuilder_lastSelection');
              const savedSheetName = localStorage.getItem('sheetBuilder_lastSheetName');
              
              if (savedSheetName) {
                document.getElementById('sheetName').value = savedSheetName;
              }
              
              if (savedSelections) {
                const selections = JSON.parse(savedSelections);
                selections.forEach(value => {
                  const checkbox = document.querySelector('input[value="' + value + '"]');
                  if (checkbox) {
                    checkbox.checked = true;
                  }
                });
              }
            } catch (error) {
              console.log('Could not restore last selections:', error);
            }
          }
          
          function saveCurrentSelections() {
            try {
              const sheetName = document.getElementById('sheetName').value.trim();
              const checkboxes = document.querySelectorAll('input[name="columns"]:checked');
              const selectedColumns = Array.from(checkboxes).map(cb => cb.value);
              
              localStorage.setItem('sheetBuilder_lastSelection', JSON.stringify(selectedColumns));
              localStorage.setItem('sheetBuilder_lastSheetName', sheetName);
            } catch (error) {
              console.log('Could not save selections:', error);
            }
          }
          
          function createSheet() {
            const sheetName = document.getElementById('sheetName').value.trim();
            
            if (!sheetName) {
              alert('Please enter a sheet name');
              return;
            }
            
            const checkboxes = document.querySelectorAll('input[name="columns"]:checked');
            const selectedColumns = Array.from(checkboxes).map(cb => cb.value);
            
            // Save current selections for next time
            saveCurrentSelections();
            
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
  
  // Parse selected columns to separate roster and attendance columns
  const rosterColumns = [];
  const attendanceColumns = [];
  
  selectedColumns.forEach(columnSpec => {
    if (columnSpec.startsWith('roster:')) {
      rosterColumns.push(columnSpec.substring(7)); // Remove 'roster:' prefix
    } else if (columnSpec.startsWith('attendance:')) {
      attendanceColumns.push(columnSpec.substring(11)); // Remove 'attendance:' prefix
    } else {
      // Legacy format (assume roster column)
      rosterColumns.push(columnSpec);
    }
  });
  
  console.log(`Processing ${rosterColumns.length} roster columns and ${attendanceColumns.length} attendance columns`);
  
  // Create the sheet
  createCustomSheetWithColumns(sheetName, rosterColumns, attendanceColumns);
  
  // Return success message
  const totalColumns = rosterColumns.length + attendanceColumns.length;
  const message = totalColumns === 0 
    ? `Created "${sheetName}" with only the Full Name column.`
    : `Created "${sheetName}" with Full Name + ${totalColumns} additional columns (${rosterColumns.length} from Roster, ${attendanceColumns.length} from Attendance).`;
    
  return message;
}

/**
 * Create the custom sheet with selected columns
 */
function createCustomSheetWithColumns(sheetName, rosterColumns, attendanceColumns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  
  try {
    // Create new sheet
    console.log('📄 Creating new sheet...');
    const newSheet = ss.insertSheet(sheetName);
    
    // Find Full Name column in roster
    const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    const fullNameColIndex = rosterHeaderRow.findIndex(name => name === 'Full Name');
    
    if (fullNameColIndex === -1) {
      throw new Error('Full Name column not found in roster sheet');
    }
    
    // Get attendance sheet info if we have attendance columns
    let attendanceSheet = null;
    let attendanceHeaderRow = null;
    if (attendanceColumns.length > 0) {
      attendanceSheet = getAttendanceSheet();
      if (!attendanceSheet) {
        throw new Error('Cannot access Attendance sheet to create selected columns');
      }
      attendanceHeaderRow = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
    }
    
    // Set up headers: Full Name + roster columns + attendance columns
    console.log('📋 Setting up headers...');
    const allSelectedColumns = [...rosterColumns, ...attendanceColumns];
    const headers = ['Full Name', ...allSelectedColumns];
    
    // Set headers with direct cell references for column names
    const headerFormulas = headers.map((columnName, index) => {
      if (index === 0) {
        // First column is always Full Name
        return 'Full Name';
      } else {
        // Check if this is a roster column or attendance column
        const isRosterColumn = rosterColumns.includes(columnName);
        const isAttendanceColumn = attendanceColumns.includes(columnName);
        
        if (isRosterColumn) {
          // Use direct reference to roster column header
          const rosterColumnIndex = rosterHeaderRow.findIndex(name => name === columnName);
          if (rosterColumnIndex === -1) {
            return columnName; // Fallback to static name if not found
          }
          const rosterColumnLetter = getColumnLetter(rosterColumnIndex + 1);
          return `=${CONFIG.roster.sheetName}!${rosterColumnLetter}1`;
        } else if (isAttendanceColumn && attendanceHeaderRow) {
          // Use static name for attendance columns (they come from external sheet)
          return columnName;
        } else {
          return columnName; // Fallback
        }
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
    
    // Create XLOOKUP formulas for roster columns
    console.log('🔗 Creating XLOOKUP formulas for roster columns...');
    rosterColumns.forEach((columnName, columnIndex) => {
      const targetColumn = columnIndex + 2; // +2 because Full Name is column 1
      
      // Find the column in roster sheet
      const rosterColumnIndex = rosterHeaderRow.findIndex(name => name === columnName);
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
    
    // Create XLOOKUP formulas for attendance columns
    if (attendanceColumns && attendanceColumns.length > 0) {
      console.log('🔗 Creating XLOOKUP formulas for attendance columns...');
      const attendanceSheet = getAttendanceSheet();
      if (attendanceSheet) {
        const attendanceHeaderRow = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
        
        attendanceColumns.forEach((columnName, columnIndex) => {
          const targetColumn = rosterColumns.length + columnIndex + 2; // After roster columns
          
          // Find the column in attendance sheet
          const attendanceColumnIndex = attendanceHeaderRow.findIndex(name => name === columnName);
          if (attendanceColumnIndex === -1) {
            console.warn(`Column "${columnName}" not found in attendance sheet`);
            return;
          }
          
          const attendanceColumnLetter = getColumnLetter(attendanceColumnIndex + 1);
          const fullNameColumnLetter = getColumnLetter(fullNameColIndex + 1);
          
          // Create XLOOKUP formula referencing Attendance sheet
          const formula = `=IFERROR(XLOOKUP(A2,'Attendance'!A:A,'Attendance'!${attendanceColumnLetter}:${attendanceColumnLetter}),"")`;
          
          // Set formula in first data row
          newSheet.getRange(newSheetDataStartRow, targetColumn).setFormula(formula);
          
          // Copy formula down for all rows with Full Name data
          if (nonEmptyFullNames.length > 1) {
            const sourceRange = newSheet.getRange(newSheetDataStartRow, targetColumn, 1, 1);
            const targetRange = newSheet.getRange(newSheetDataStartRow + 1, targetColumn, nonEmptyFullNames.length - 1, 1);
            sourceRange.copyTo(targetRange);
          }
        });
      }
    }
    
    // Copy column formatting from roster sheet
    console.log('🎨 Copying column formatting...');
    copyColumnFormattingFromRoster(newSheet, rosterSheet, headers, rosterHeaderRow);
    
    // Apply Format Spruce Up formatting (alternating colors, filtering, vertical centering)
    console.log('✨ Applying Format Spruce Up formatting...');
    applySpruceUpFormatting(newSheet);
    
    // Ensure header row styling is preserved (after Format Spruce Up)
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    // Copy conditional formatting from roster sheet
    console.log('🎨 Copying conditional formatting...');
    const totalRows = nonEmptyFullNames.length + 1; // +1 for header row
    copyConditionalFormattingFromRoster(newSheet, rosterSheet, headers, rosterHeaderRow, totalRows);
    
    // Activate the new sheet
    console.log('🎯 Activating new sheet...');
    newSheet.activate();
    
    const totalColumns = rosterColumns.length + (attendanceColumns ? attendanceColumns.length : 0);
    console.log(`✅ Sheet creation complete! "${sheetName}" with ${totalColumns} columns (${rosterColumns.length} roster, ${attendanceColumns ? attendanceColumns.length : 0} attendance) and ${nonEmptyFullNames.length} students`);
    
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