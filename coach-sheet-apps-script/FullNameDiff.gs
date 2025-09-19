/**
 * Full Name Diff Module
 * Compare Full Name columns between different sheets to find differences
 */

/**
 * Main function to start the Full Name Diff process
 * Called from the menu
 */
function fullNameDiff() {
  console.log('🔍 Starting Full Name Diff...');
  
  try {
    // Discover all sheets with Full Name columns
    const sheetsWithFullName = discoverSheetsWithFullName();
    
    if (sheetsWithFullName.length < 2) {
      SpreadsheetApp.getUi().alert(
        'Insufficient Sheets',
        `Found ${sheetsWithFullName.length} sheet(s) with Full Name column. Need at least 2 sheets to compare.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Show sheet selection dialog
    showSheetSelectionDialog(sheetsWithFullName);
    
  } catch (error) {
    console.error('Error in Full Name Diff:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to start Full Name Diff: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Discover all sheets in the spreadsheet that have a "Full Name" column
 * @return {Array} Array of objects with sheet info: {name, fullNameColumn}
 */
function discoverSheetsWithFullName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const sheetsWithFullName = [];
  
  console.log(`🔍 Scanning ${allSheets.length} sheets for Full Name columns...`);
  
  allSheets.forEach(sheet => {
    try {
      const sheetName = sheet.getName();
      
      // Skip hidden sheets and sheets that are likely temporary
      if (sheet.isSheetHidden() || sheetName.startsWith('_')) {
        console.log(`⏭️ Skipping hidden/temp sheet: "${sheetName}"`);
        return;
      }
      
      // Get header row (row 1)
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) {
        console.log(`⏭️ Skipping empty sheet: "${sheetName}"`);
        return;
      }
      
      const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // Look for "Full Name" column (case-sensitive)
      const fullNameColumnIndex = headerRow.findIndex(header => 
        header && header.toString() === 'Full Name'
      );
      
      if (fullNameColumnIndex !== -1) {
        sheetsWithFullName.push({
          name: sheetName,
          fullNameColumn: fullNameColumnIndex + 1, // Convert to 1-based
          sheet: sheet
        });
        console.log(`✅ Found Full Name column in "${sheetName}" at column ${fullNameColumnIndex + 1}`);
      } else {
        console.log(`❌ No Full Name column in "${sheetName}"`);
      }
      
    } catch (error) {
      console.warn(`Error analyzing sheet "${sheet.getName()}":`, error);
    }
  });
  
  console.log(`🎯 Found ${sheetsWithFullName.length} sheets with Full Name columns`);
  return sheetsWithFullName;
}

/**
 * Show dialog for selecting two sheets to compare
 * @param {Array} sheetsWithFullName - Array of sheet objects with Full Name columns
 */
function showSheetSelectionDialog(sheetsWithFullName) {
  const htmlContent = createSheetSelectionHtml(sheetsWithFullName);
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(500)
    .setTitle('Full Name Diff - Select Sheets');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Full Name Diff');
}

/**
 * Create HTML content for sheet selection dialog
 * @param {Array} sheetsWithFullName - Array of sheet objects with Full Name columns
 * @return {string} HTML content
 */
function createSheetSelectionHtml(sheetsWithFullName) {
  const sheetOptions = sheetsWithFullName.map(sheet => 
    `<option value="${sheet.name}">${sheet.name}</option>`
  ).join('');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 20px; 
            background: #f8f9fa;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          h2 {
            color: #1a73e8;
            margin-top: 0;
            text-align: center;
          }
          .form-group { 
            margin-bottom: 20px; 
          }
          label { 
            font-weight: bold; 
            display: block; 
            margin-bottom: 8px; 
            color: #333;
          }
          select { 
            width: 100%; 
            padding: 10px; 
            font-size: 14px; 
            border: 1px solid #ddd;
            border-radius: 4px;
            background: white;
          }
          .buttons {
            text-align: center;
            margin-top: 30px;
          }
          .btn {
            padding: 12px 24px;
            font-size: 14px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin: 0 10px;
            font-weight: bold;
          }
          .btn-primary {
            background: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background: #1557b0;
          }
          .btn-secondary {
            background: #f1f3f4;
            color: #5f6368;
          }
          .btn-secondary:hover {
            background: #e8eaed;
          }
          .info {
            background: #e8f0fe;
            border: 1px solid #d2e3fc;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
            color: #1565c0;
          }
          .error {
            color: #d93025;
            font-size: 14px;
            margin-top: 10px;
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🔍 Full Name Diff</h2>
          
          <div class="info">
            <strong>Found ${sheetsWithFullName.length} sheets with Full Name columns.</strong><br>
            Select two sheets to compare their Full Name columns and see the differences.
          </div>
          
          <div class="form-group">
            <label for="sheet1">First Sheet:</label>
            <select id="sheet1" name="sheet1">
              <option value="">-- Select First Sheet --</option>
              ${sheetOptions}
            </select>
          </div>
          
          <div class="form-group">
            <label for="sheet2">Second Sheet:</label>
            <select id="sheet2" name="sheet2">
              <option value="">-- Select Second Sheet --</option>
              ${sheetOptions}
            </select>
          </div>
          
          <div class="error" id="errorMessage"></div>
          
          <div class="buttons">
            <button class="btn btn-primary" onclick="compareSheets()">Compare Sheets</button>
            <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>
        </div>
        
        <script>
          // Load saved selections when dialog opens
          window.onload = function() {
            restoreLastSelections();
          };
          
          function restoreLastSelections() {
            try {
              const savedSheet1 = localStorage.getItem('fullNameDiff_lastSheet1');
              const savedSheet2 = localStorage.getItem('fullNameDiff_lastSheet2');
              
              if (savedSheet1) {
                const sheet1Select = document.getElementById('sheet1');
                if (sheet1Select.querySelector('option[value="' + savedSheet1 + '"]')) {
                  sheet1Select.value = savedSheet1;
                }
              }
              
              if (savedSheet2) {
                const sheet2Select = document.getElementById('sheet2');
                if (sheet2Select.querySelector('option[value="' + savedSheet2 + '"]')) {
                  sheet2Select.value = savedSheet2;
                }
              }
            } catch (error) {
              console.log('Could not restore last selections:', error);
            }
          }
          
          function saveCurrentSelections(sheet1, sheet2) {
            try {
              localStorage.setItem('fullNameDiff_lastSheet1', sheet1);
              localStorage.setItem('fullNameDiff_lastSheet2', sheet2);
            } catch (error) {
              console.log('Could not save selections:', error);
            }
          }
          
          function compareSheets() {
            const sheet1 = document.getElementById('sheet1').value;
            const sheet2 = document.getElementById('sheet2').value;
            const errorEl = document.getElementById('errorMessage');
            
            // Validation
            if (!sheet1 || !sheet2) {
              showError('Please select both sheets.');
              return;
            }
            
            if (sheet1 === sheet2) {
              showError('Please select two different sheets.');
              return;
            }
            
            // Save current selections for next time
            saveCurrentSelections(sheet1, sheet2);
            
            // Hide error and call server function
            errorEl.style.display = 'none';
            
            google.script.run
              .withSuccessHandler(showResults)
              .withFailureHandler(showError)
              .performFullNameComparison(sheet1, sheet2);
          }
          
          function showError(error) {
            const errorEl = document.getElementById('errorMessage');
            errorEl.textContent = typeof error === 'string' ? error : error.message;
            errorEl.style.display = 'block';
          }
          
          function showResults(result) {
            // Close current dialog and show results
            google.script.host.close();
            
            // Results will be shown in a new dialog by the server function
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Server-side function to perform the Full Name comparison
 * @param {string} sheet1Name - Name of first sheet
 * @param {string} sheet2Name - Name of second sheet
 * @return {Object} Comparison results
 */
function performFullNameComparison(sheet1Name, sheet2Name) {
  console.log(`🔍 Comparing Full Names: "${sheet1Name}" vs "${sheet2Name}"`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName(sheet1Name);
    const sheet2 = ss.getSheetByName(sheet2Name);
    
    if (!sheet1 || !sheet2) {
      throw new Error('One or both selected sheets not found');
    }
    
    // Get Full Name data from both sheets
    const names1 = getFullNamesFromSheet(sheet1, sheet1Name);
    const names2 = getFullNamesFromSheet(sheet2, sheet2Name);
    
    // Perform comparison
    const comparison = compareFullNames(names1, names2, sheet1Name, sheet2Name);
    
    // Show results dialog
    showComparisonResults(comparison);
    
    return comparison;
    
  } catch (error) {
    console.error('Error performing Full Name comparison:', error);
    throw error;
  }
}

/**
 * Extract Full Name values from a sheet
 * @param {Sheet} sheet - The sheet to extract names from
 * @param {string} sheetName - Name of the sheet for logging
 * @return {Array} Array of clean Full Name values
 */
function getFullNamesFromSheet(sheet, sheetName) {
  // Find Full Name column (case-sensitive)
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fullNameColumnIndex = headerRow.findIndex(header => 
    header && header.toString() === 'Full Name'
  );
  
  if (fullNameColumnIndex === -1) {
    throw new Error(`Full Name column not found in sheet "${sheetName}"`);
  }
  
  // Get all values in Full Name column (skip header row)
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    console.log(`⚠️ Sheet "${sheetName}" has no data rows`);
    return [];
  }
  
  const fullNameData = sheet.getRange(2, fullNameColumnIndex + 1, lastRow - 1, 1).getValues();
  
  // Use strict string comparison - only filter out truly empty cells
  const strictNames = fullNameData
    .map(row => row[0])
    .filter(name => name !== null && name !== undefined && name !== '')
    .map(name => name.toString()); // Convert to string but don't trim
  
  console.log(`📊 Found ${strictNames.length} names in "${sheetName}" (strict comparison)`);
  return strictNames;
}

/**
 * Compare two arrays of Full Names and find differences
 * @param {Array} names1 - Names from first sheet
 * @param {Array} names2 - Names from second sheet
 * @param {string} sheet1Name - Name of first sheet
 * @param {string} sheet2Name - Name of second sheet
 * @return {Object} Comparison results
 */
function compareFullNames(names1, names2, sheet1Name, sheet2Name) {
  const set1 = new Set(names1);
  const set2 = new Set(names2);
  
  // Find names in sheet1 but not in sheet2
  const onlyInSheet1 = names1.filter(name => !set2.has(name));
  
  // Find names in sheet2 but not in sheet1
  const onlyInSheet2 = names2.filter(name => !set1.has(name));
  
  // Find common names
  const inBothSheets = names1.filter(name => set2.has(name));
  
  const results = {
    sheet1Name,
    sheet2Name,
    sheet1Count: names1.length,
    sheet2Count: names2.length,
    onlyInSheet1: [...new Set(onlyInSheet1)], // Remove duplicates
    onlyInSheet2: [...new Set(onlyInSheet2)], // Remove duplicates
    inBothSheets: [...new Set(inBothSheets)], // Remove duplicates
  };
  
  console.log(`📈 Comparison results:
    - ${sheet1Name}: ${results.sheet1Count} names
    - ${sheet2Name}: ${results.sheet2Count} names
    - Only in ${sheet1Name}: ${results.onlyInSheet1.length}
    - Only in ${sheet2Name}: ${results.onlyInSheet2.length}
    - In both sheets: ${results.inBothSheets.length}`);
  
  return results;
}

/**
 * Show comparison results in a new dialog
 * @param {Object} comparison - Comparison results object
 */
function showComparisonResults(comparison) {
  const htmlContent = createComparisonResultsHtml(comparison);
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(700)
    .setHeight(600)
    .setTitle('Full Name Diff - Results');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Full Name Diff Results');
}

/**
 * Create HTML content for comparison results
 * @param {Object} comparison - Comparison results object
 * @return {string} HTML content
 */
function createComparisonResultsHtml(comparison) {
  const formatNameList = (names) => {
    if (names.length === 0) {
      return '<em style="color: #666;">None</em>';
    }
    return names.map(name => `<div class="name-item">• ${name}</div>`).join('');
  };
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 20px; 
            background: #f8f9fa;
            margin: 0;
          }
          .container {
            max-width: 650px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          h2 {
            color: #1a73e8;
            margin-top: 0;
            text-align: center;
          }
          .summary {
            background: #e8f0fe;
            border: 1px solid #d2e3fc;
            padding: 20px;
            border-radius: 6px;
            margin-bottom: 30px;
          }
          .summary-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
          }
          .summary-item {
            text-align: center;
          }
          .summary-number {
            font-size: 24px;
            font-weight: bold;
            color: #1a73e8;
          }
          .summary-label {
            color: #5f6368;
            font-size: 14px;
          }
          .section {
            margin-bottom: 30px;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            overflow: hidden;
          }
          .section-header {
            padding: 15px 20px;
            font-weight: bold;
            color: white;
            font-size: 16px;
          }
          .section-content {
            padding: 20px;
            max-height: 200px;
            overflow-y: auto;
          }
          .only-in-1 .section-header {
            background: #ea4335;
          }
          .only-in-2 .section-header {
            background: #fbbc04;
            color: #333;
          }
          .in-both .section-header {
            background: #34a853;
          }
          .name-item {
            padding: 4px 0;
            border-bottom: 1px solid #f0f0f0;
          }
          .name-item:last-child {
            border-bottom: none;
          }
          .buttons {
            text-align: center;
            margin-top: 30px;
          }
          .btn {
            padding: 12px 24px;
            font-size: 14px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin: 0 10px;
            font-weight: bold;
          }
          .btn-secondary {
            background: #f1f3f4;
            color: #5f6368;
          }
          .btn-secondary:hover {
            background: #e8eaed;
          }
          .count {
            font-weight: bold;
            color: #1a73e8;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🔍 Full Name Diff Results</h2>
          
          <div class="summary">
            <div class="summary-grid">
              <div class="summary-item">
                <div class="summary-number">${comparison.sheet1Count}</div>
                <div class="summary-label">${comparison.sheet1Name}</div>
              </div>
              <div class="summary-item">
                <div class="summary-number">${comparison.sheet2Count}</div>
                <div class="summary-label">${comparison.sheet2Name}</div>
              </div>
            </div>
          </div>
          
          <div class="section only-in-1">
            <div class="section-header">
              Only in "${comparison.sheet1Name}" (<span class="count">${comparison.onlyInSheet1.length}</span>)
            </div>
            <div class="section-content">
              ${formatNameList(comparison.onlyInSheet1)}
            </div>
          </div>
          
          <div class="section only-in-2">
            <div class="section-header">
              Only in "${comparison.sheet2Name}" (<span class="count">${comparison.onlyInSheet2.length}</span>)
            </div>
            <div class="section-content">
              ${formatNameList(comparison.onlyInSheet2)}
            </div>
          </div>
          
          <div class="section in-both">
            <div class="section-header">
              In Both Sheets (<span class="count">${comparison.inBothSheets.length}</span>)
            </div>
            <div class="section-content">
              ${formatNameList(comparison.inBothSheets)}
            </div>
          </div>
          
          <div class="buttons">
            <button class="btn btn-secondary" onclick="google.script.host.close()">Close</button>
          </div>
        </div>
      </body>
    </html>
  `;
}