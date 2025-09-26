/**
 * Build Practice Roster Module
 * Creates practice-specific roster sheets with availability data
 */

/**
 * Main function to build a practice roster
 * Called from the menu
 */
function buildPracticeRoster() {
  console.log('🏅 Starting Build Practice Roster...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get practice dates from Practice Info sheet
    const practiceDates = getDatesFromInfoSheet(ss, PRACTICE_AVAILABILITY_CONFIG);
    
    if (practiceDates.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Practice Dates Found',
        'No practice dates found in "📍Practice Info" sheet. Please ensure the sheet exists and contains practice dates.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Find the next upcoming practice (including today)
    const defaultPracticeIndex = findNextUpcomingPractice(practiceDates);
    
    // Show date selection dialog
    showPracticeDateSelectionDialog(practiceDates, defaultPracticeIndex);
    
  } catch (error) {
    console.error('Error building practice roster:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to build practice roster: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Find the index of the next upcoming practice (including today)
 * @param {Array} practiceDates - Array of practice date objects
 * @return {number} Index of the next upcoming practice, or 0 if none found
 */
function findNextUpcomingPractice(practiceDates) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Reset to start of day for comparison
  
  for (let i = 0; i < practiceDates.length; i++) {
    const practiceDate = new Date(practiceDates[i].date);
    practiceDate.setHours(0, 0, 0, 0);
    
    if (practiceDate >= today) {
      console.log(`📅 Next upcoming practice: ${practiceDates[i].formattedDate} (index ${i})`);
      return i;
    }
  }
  
  console.log('📅 No upcoming practices found, defaulting to first practice');
  return 0; // Default to first practice if no upcoming ones
}

/**
 * Show the practice date selection dialog
 * @param {Array} practiceDates - Array of practice date objects
 * @param {number} defaultIndex - Index of the default selected practice
 */
function showPracticeDateSelectionDialog(practiceDates, defaultIndex) {
  const html = createPracticeDateSelectionHtml(practiceDates, defaultIndex);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(300);
    
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Build Practice Roster');
}

/**
 * Create HTML for practice date selection dialog
 * @param {Array} practiceDates - Array of practice date objects
 * @param {number} defaultIndex - Index of the default selected practice
 * @return {string} HTML content
 */
function createPracticeDateSelectionHtml(practiceDates, defaultIndex) {
  const defaultDate = practiceDates[defaultIndex].formattedDate;
  const defaultSheetName = `${defaultDate} Roster`;

  // Create dropdown options
  const dateOptions = practiceDates.map((pd, index) => {
    const selected = index === defaultIndex ? 'selected' : '';
    return `<option value="${pd.formattedDate}" ${selected}>${pd.formattedDate}</option>`;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            padding: 20px;
            margin: 0;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            font-weight: 500;
            margin-bottom: 8px;
            color: #202124;
            font-size: 14px;
          }
          select, input[type="text"], input[type="radio"] {
            padding: 10px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
          }
          select, input[type="text"] {
            width: 100%;
          }
          input[type="radio"] {
            width: auto;
            margin-right: 8px;
            padding: 0;
          }
          select:focus, input[type="text"]:focus {
            outline: none;
            border-color: #1a73e8;
          }
          .note {
            font-size: 12px;
            color: #5f6368;
            margin-top: 5px;
          }
          .radio-group {
            display: flex;
            gap: 20px;
            margin-bottom: 15px;
          }
          .radio-option {
            display: flex;
            align-items: center;
          }
          .buttons {
            display: flex;
            gap: 10px;
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
          }
          .btn {
            flex: 1;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: background-color 0.2s;
          }
          .btn-primary {
            background-color: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background-color: #1557b0;
          }
          .btn-secondary {
            background-color: #f8f9fa;
            color: #3c4043;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background-color: #f1f3f4;
          }
          .progress-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.95);
            z-index: 1000;
          }
          .progress-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .hidden {
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="practiceDate">Select Practice Date:</label>
          <select id="practiceDate" onchange="updateTargetSheet()">
            ${dateOptions}
          </select>
          <div class="note">Choose the practice date for this roster</div>
        </div>

        <div class="form-group">
          <label>Action:</label>
          <div class="radio-group">
            <div class="radio-option">
              <input type="radio" id="actionCreate" name="action" value="create" checked onchange="toggleActionFields()">
              <label for="actionCreate" style="margin-bottom: 0;">Create New Sheet</label>
            </div>
            <div class="radio-option">
              <input type="radio" id="actionUpdate" name="action" value="update" onchange="toggleActionFields()">
              <label for="actionUpdate" style="margin-bottom: 0;">Update Existing Sheet</label>
            </div>
          </div>
        </div>

        <div class="form-group" id="newSheetGroup">
          <label for="sheetName">Sheet Name:</label>
          <input type="text" id="sheetName" value="${defaultSheetName}">
          <div class="note">Name for the new practice roster sheet</div>
        </div>

        <div class="form-group hidden" id="existingSheetGroup">
          <label for="existingSheet">Select Sheet to Update:</label>
          <select id="existingSheet">
            <option value="">Loading sheets...</option>
          </select>
          <div class="note">Choose an existing sheet to update with new data</div>
        </div>

        <div class="buttons">
          <button class="btn btn-primary" onclick="processRoster()" id="actionButton">Create Roster</button>
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>

        <div class="progress-overlay" id="progressOverlay">
          <div class="progress-content">
            <div class="spinner"></div>
            <div style="font-size: 16px; font-weight: bold; color: #333;" id="progressTitle">
              Building Practice Roster...
            </div>
            <div style="font-size: 14px; color: #666; margin-top: 8px;" id="progressMessage">
              Please wait while we create your roster
            </div>
          </div>
        </div>

        <script>
          let activeSheetName = '';

          // Load existing sheets and restore last selection on page load
          window.onload = function() {
            loadActiveSheetName();
            loadExistingSheets();
            restoreLastSelection();
          };

          function loadActiveSheetName() {
            google.script.run
              .withSuccessHandler(function(sheetName) {
                activeSheetName = sheetName;
                console.log('Active sheet:', activeSheetName);
              })
              .withFailureHandler(function(error) {
                console.error('Failed to get active sheet name:', error);
              })
              .getActiveSheetName();
          }

          function loadExistingSheets() {
            google.script.run
              .withSuccessHandler(function(sheets) {
                const select = document.getElementById('existingSheet');
                select.innerHTML = '<option value="">Select a sheet...</option>';
                sheets.forEach(sheet => {
                  const option = document.createElement('option');
                  option.value = sheet;
                  option.textContent = sheet;
                  select.appendChild(option);
                });
              })
              .withFailureHandler(function(error) {
                console.error('Failed to load sheets:', error);
                document.getElementById('existingSheet').innerHTML = '<option value="">Error loading sheets</option>';
              })
              .getPracticeRosterSheets();
          }

          function restoreLastSelection() {
            const lastSheet = localStorage.getItem('lastPracticeRosterSheet');
            const lastAction = localStorage.getItem('lastPracticeRosterAction');

            if (lastAction === 'update') {
              document.getElementById('actionUpdate').checked = true;
              toggleActionFields();

              if (lastSheet) {
                // Wait a bit for sheets to load, then select
                setTimeout(() => {
                  const select = document.getElementById('existingSheet');
                  if (select.querySelector(\`option[value="\${lastSheet}"]\`)) {
                    select.value = lastSheet;
                  }
                }, 500);
              }
            }
          }

          function toggleActionFields() {
            const isCreate = document.getElementById('actionCreate').checked;
            const newSheetGroup = document.getElementById('newSheetGroup');
            const existingSheetGroup = document.getElementById('existingSheetGroup');
            const actionButton = document.getElementById('actionButton');

            if (isCreate) {
              newSheetGroup.classList.remove('hidden');
              existingSheetGroup.classList.add('hidden');
              actionButton.textContent = 'Create Roster';
            } else {
              newSheetGroup.classList.add('hidden');
              existingSheetGroup.classList.remove('hidden');
              actionButton.textContent = 'Update Roster';

              // When switching to update mode, default to active sheet if available
              const select = document.getElementById('existingSheet');
              if (activeSheetName && select.querySelector(\`option[value="\${activeSheetName}"]\`)) {
                select.value = activeSheetName;
              }
            }
          }

          function updateTargetSheet() {
            const practiceDate = document.getElementById('practiceDate').value;
            document.getElementById('sheetName').value = practiceDate + ' Roster';
          }

          function processRoster() {
            const practiceDate = document.getElementById('practiceDate').value;
            const isCreate = document.getElementById('actionCreate').checked;

            if (isCreate) {
              createNewRoster(practiceDate);
            } else {
              updateExistingRoster(practiceDate);
            }
          }

          function createNewRoster(practiceDate) {
            const sheetName = document.getElementById('sheetName').value.trim();

            if (!sheetName) {
              alert('Please enter a sheet name');
              return;
            }

            // Show progress
            showProgress('Creating Practice Roster...', 'Please wait while we create your roster');

            // Save action to localStorage
            localStorage.setItem('lastPracticeRosterAction', 'create');
            localStorage.setItem('lastPracticeRosterSheet', sheetName);

            // Check for duplicate sheet name first
            google.script.run
              .withSuccessHandler(function(isDuplicate) {
                if (isDuplicate) {
                  hideProgress();
                  alert('Sheet name "' + sheetName + '" already exists. Please choose a different name.');
                  return;
                }

                // If not duplicate, create the roster
                google.script.run
                  .withSuccessHandler(onSuccess)
                  .withFailureHandler(onFailure)
                  .createPracticeRosterSheet(sheetName, practiceDate);
              })
              .withFailureHandler(onFailure)
              .isSheetNameDuplicate(sheetName);
          }

          function updateExistingRoster(practiceDate) {
            const existingSheet = document.getElementById('existingSheet').value;

            if (!existingSheet) {
              alert('Please select a sheet to update');
              return;
            }

            // Show progress
            showProgress('Updating Practice Roster...', 'Please wait while we update your roster');

            // Save action and sheet to localStorage
            localStorage.setItem('lastPracticeRosterAction', 'update');
            localStorage.setItem('lastPracticeRosterSheet', existingSheet);

            // Update the existing roster
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .updatePracticeRosterSheet(existingSheet, practiceDate);
          }

          function showProgress(title, message) {
            document.getElementById('progressTitle').textContent = title;
            document.getElementById('progressMessage').textContent = message;
            document.getElementById('progressOverlay').style.display = 'block';
          }

          function hideProgress() {
            document.getElementById('progressOverlay').style.display = 'none';
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
 * Get all existing practice roster sheets for the dropdown
 * @return {Array} Array of sheet names that look like practice rosters
 */
function getPracticeRosterSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Filter for sheets that look like practice rosters (contain "Roster" in name)
  const practiceRosterSheets = sheets
    .map(sheet => sheet.getName())
    .filter(name => name.toLowerCase().includes('roster'))
    .sort();

  console.log(`Found ${practiceRosterSheets.length} potential practice roster sheets`);
  return practiceRosterSheets;
}

/**
 * Get the name of the currently active sheet
 * @return {string} Name of the active sheet
 */
function getActiveSheetName() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  return activeSheet.getName();
}

/**
 * Update an existing practice roster sheet with new data (content only, preserve formatting)
 * @param {string} sheetName - Name of the existing sheet to update
 * @param {string} practiceDate - Practice date in format "M/D"
 */
function updatePracticeRosterSheet(sheetName, practiceDate) {
  console.log(`🔄 Updating existing practice roster sheet: "${sheetName}" for date: ${practiceDate}`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheet = ss.getSheetByName(sheetName);

    if (!existingSheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Get source sheets
    const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
    const practiceAvailabilitySheet = ss.getSheetByName('Practice Availability');
    const gameAvailabilitySheet = ss.getSheetByName('Game Availability');

    if (!rosterSheet) {
      throw new Error('Roster sheet not found');
    }

    if (!practiceAvailabilitySheet) {
      throw new Error('Practice Availability sheet not found');
    }

    // Find the availability columns for this practice date
    const availColumns = findPracticeAvailabilityColumns(practiceAvailabilitySheet, practiceDate);

    if (!availColumns.availabilityColumn) {
      throw new Error(`Practice date "${practiceDate}" not found in Practice Availability sheet`);
    }

    // Find the next game date after this practice
    const nextGameInfo = findNextGameAfterPractice(ss, practiceDate);

    // Clear content from data area only (preserve headers and formatting)
    console.log('🧹 Clearing existing content while preserving formatting...');
    const lastRow = existingSheet.getLastRow();
    const lastColumn = existingSheet.getLastColumn();

    if (lastRow > 1 && lastColumn > 0) {
      // Clear content starting from row 2 (preserve header row)
      const dataRange = existingSheet.getRange(2, 1, lastRow - 1, lastColumn);
      dataRange.clearContent();

      // Clear borders from data area (but preserve other formatting)
      console.log('🧹 Clearing existing borders...');
      dataRange.setBorder(false, false, false, false, false, false);
    }

    // Update headers with new dates (in case the practice date changed)
    const headers = ['#', 'Full Name', 'Grade', 'Gender', 'Team', practiceDate, `${practiceDate} Note`];

    // Add next game columns if found
    if (nextGameInfo) {
      headers.push(nextGameInfo.formattedDate);
      headers.push(`${nextGameInfo.formattedDate} Note`);
    }

    // Update header row (preserve formatting but update text)
    const headerRange = existingSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    console.log(`📝 Updated ${headers.length} column headers: ${headers.join(', ')}`);

    if (nextGameInfo) {
      console.log(`🎮 Next game found: ${nextGameInfo.formattedDate}`);
    } else {
      console.log('🎮 No next game found after this practice');
    }

    // Populate with fresh data using existing functions
    const fullNameInfo = copyFullNameColumnToColumn(existingSheet, rosterSheet, 2, 2);
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);

    if (fullNameInfo.rowCount > 0) {
      const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];

      // Populate other columns with XLOOKUP formulas
      populatePracticeRosterData(existingSheet, rosterSheet, rosterHeaderRow, practiceAvailabilitySheet, availColumns, fullNameInfo.rowCount, gameAvailabilitySheet, nextGameInfo);

      // Force recalculation to ensure formulas are evaluated before sorting
      SpreadsheetApp.flush();

      // Sort the data AFTER formulas have been calculated
      sortPracticeRoster(existingSheet, fullNameInfo.rowCount, headers.length);

      // Populate # column AFTER sorting (so the formula references are correct)
      populateNumberColumn(existingSheet, fullNameInfo.rowCount);

      // Force calculation of # column formulas before adding borders
      SpreadsheetApp.flush();

      // Add borders at group changes (where # = 1)
      addGroupBorders(existingSheet, fullNameInfo.rowCount);
    }

    console.log(`✅ Practice roster "${sheetName}" updated successfully`);

    // Show success alert
    SpreadsheetApp.getUi().alert(
      'Practice Roster Updated!',
      `Successfully updated practice roster "${sheetName}" for ${practiceDate} with ${fullNameInfo.rowCount} students.\n\nFormatting and layout preserved.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return 'Success';

  } catch (error) {
    console.error('Error updating practice roster sheet:', error);
    throw new Error(`Failed to update practice roster: ${error.message}`);
  }
}

/**
 * Create the practice roster sheet with all data
 * @param {string} sheetName - Name for the new sheet
 * @param {string} practiceDate - Practice date in format "M/D"
 */
function createPracticeRosterSheet(sheetName, practiceDate) {
  console.log(`📋 Creating practice roster sheet: "${sheetName}" for date: ${practiceDate}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create new sheet
    const newSheet = ss.insertSheet(sheetName);
    console.log(`✅ Created new sheet: "${sheetName}"`);
    
    // Get source sheets
    const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
    const practiceAvailabilitySheet = ss.getSheetByName('Practice Availability');
    const gameAvailabilitySheet = ss.getSheetByName('Game Availability');
    
    if (!rosterSheet) {
      throw new Error('Roster sheet not found');
    }
    
    if (!practiceAvailabilitySheet) {
      throw new Error('Practice Availability sheet not found');
    }
    
    // Find the availability columns for this practice date
    const availColumns = findPracticeAvailabilityColumns(practiceAvailabilitySheet, practiceDate);
    
    if (!availColumns.availabilityColumn) {
      throw new Error(`Practice date "${practiceDate}" not found in Practice Availability sheet`);
    }
    
    // Find the next game date after this practice
    const nextGameInfo = findNextGameAfterPractice(ss, practiceDate);
    
    // Define column structure with dynamic names
    const headers = ['#', 'Full Name', 'Grade', 'Gender', 'Team', practiceDate, `${practiceDate} Note`];
    
    // Add next game columns if found
    if (nextGameInfo) {
      headers.push(nextGameInfo.formattedDate);
      headers.push(`${nextGameInfo.formattedDate} Note`);
    }
    
    // Set up headers
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    
    console.log(`📝 Set up ${headers.length} column headers: ${headers.join(', ')}`);
    
    if (nextGameInfo) {
      console.log(`🎮 Next game found: ${nextGameInfo.formattedDate}`);
    } else {
      console.log('🎮 No next game found after this practice');
    }
    
    console.log(`📍 Found availability columns: ${availColumns.availabilityColumn} and ${availColumns.noteColumn || 'none'}`);
    
    // Copy Full Name column to column B (column 2) from roster using shared utility
    const fullNameInfo = copyFullNameColumnToColumn(newSheet, rosterSheet, 2, 2); // startRow=2, targetColumn=2
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);
    
    const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
    const nonEmptyFullNames = {length: fullNameInfo.rowCount}; // For backward compatibility
    
    // Populate other columns with XLOOKUP formulas
    populatePracticeRosterData(newSheet, rosterSheet, rosterHeaderRow, practiceAvailabilitySheet, availColumns, fullNameInfo.rowCount, gameAvailabilitySheet, nextGameInfo);
    
    // Copy formatting from roster using shared utility
    console.log('🎨 Copying column formatting...');
    copyColumnFormatting(newSheet, rosterSheet, headers, rosterHeaderRow);
    
    // Apply Format Spruce Up silently
    console.log('✨ Applying Format Spruce Up formatting...');
    applySpruceUpFormatting(newSheet);
    
    // Ensure header row styling is preserved using shared utility
    styleHeaderRow(newSheet, headers.length);
    
    // Copy conditional formatting using shared utility
    console.log('🎨 Copying conditional formatting...');
    const totalRows = fullNameInfo.rowCount + 1;
    copyConditionalFormatting(newSheet, rosterSheet, totalRows, headers.length);
    
    // Copy data validation from Practice Availability for the Availability column using shared utility
    if (availColumns.availabilityColumn) {
      console.log('✅ Copying data validation from Practice Availability...');
      const availColIndex = practiceAvailabilitySheet.getRange(1, 1, 1, practiceAvailabilitySheet.getLastColumn())
        .getValues()[0].findIndex(h => h === practiceDate || 
          (h instanceof Date && `${h.getMonth() + 1}/${h.getDate()}` === practiceDate)) + 1;
      
      if (availColIndex > 0) {
        copyDataValidation(newSheet, practiceAvailabilitySheet, 
          [{sourceColumn: practiceDate, targetColumn: 6}], fullNameInfo.rowCount);
      }
    }
    
    // Force recalculation to ensure formulas are evaluated before sorting
    SpreadsheetApp.flush();
    
    // Sort the data AFTER formulas have been calculated
    if (fullNameInfo.rowCount > 0) {
      sortPracticeRoster(newSheet, fullNameInfo.rowCount, headers.length);
    }
    
    // Populate # column AFTER sorting (so the formula references are correct)
    if (fullNameInfo.rowCount > 0) {
      populateNumberColumn(newSheet, fullNameInfo.rowCount);
      
      // Force calculation of # column formulas before adding borders
      SpreadsheetApp.flush();
      
      // Add borders at group changes (where # = 1)
      addGroupBorders(newSheet, fullNameInfo.rowCount);
    }
    
    // Delete empty rows and columns to clean up the sheet
    console.log('🧹 Cleaning up empty rows and columns...');
    deleteEmptyRowsAndColumnsForSheet(newSheet);
    
    // Auto-resize specific columns
    console.log('📏 Auto-resizing columns...');
    newSheet.autoResizeColumn(1); // # column
    newSheet.autoResizeColumn(6); // Practice availability column
    if (nextGameInfo) {
      newSheet.autoResizeColumn(8); // Next game availability column
    }
    
    // Enable text wrapping for note columns
    console.log('📝 Enabling text wrap for note columns...');
    // Practice note column (column 7)
    const practiceNoteRange = newSheet.getRange(2, 7, fullNameInfo.rowCount, 1);
    practiceNoteRange.setWrap(true);
    
    // Next game note column (column 9) if it exists
    if (nextGameInfo) {
      const gameNoteRange = newSheet.getRange(2, 9, fullNameInfo.rowCount, 1);
      gameNoteRange.setWrap(true);
    }
    
    // Set print settings
    console.log('🖨️ Configuring print settings...');
    configurePrintSettings(newSheet);
    
    console.log(`✅ Practice roster "${sheetName}" created successfully`);
    
    // Show success alert AFTER all processing is complete
    SpreadsheetApp.getUi().alert(
      'Practice Roster Created!',
      `Successfully created practice roster for ${practiceDate} with ${fullNameInfo.rowCount} students.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return 'Success';
    
  } catch (error) {
    console.error('Error creating practice roster sheet:', error);
    throw new Error(`Failed to create practice roster: ${error.message}`);
  }
}

/**
 * Find the availability columns in Practice Availability sheet for a specific date
 * @param {Sheet} practiceAvailabilitySheet - The Practice Availability sheet
 * @param {string} practiceDate - Practice date in format "M/D"
 * @return {Object} Object with availabilityColumn and noteColumn letters
 */
function findPracticeAvailabilityColumns(practiceAvailabilitySheet, practiceDate) {
  return findAvailabilityColumns(practiceAvailabilitySheet, practiceDate, 'Practice Availability');
}

/**
 * Populate practice roster data with XLOOKUP formulas
 * @param {Sheet} newSheet - The new practice roster sheet
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {Array} rosterHeaderRow - Header row from roster sheet
 * @param {Sheet} practiceAvailabilitySheet - The practice availability sheet
 * @param {Object} availColumns - Availability column letters
 * @param {number} numRows - Number of data rows
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet (optional)
 * @param {Object} nextGameInfo - Next game information (optional)
 */
function populatePracticeRosterData(newSheet, rosterSheet, rosterHeaderRow, practiceAvailabilitySheet, availColumns, numRows, gameAvailabilitySheet = null, nextGameInfo = null) {
  if (numRows === 0) return;
  
  const rosterSheetName = CONFIG.roster.sheetName;
  const practiceAvailSheetName = 'Practice Availability';
  
  // Find "Full Name" column for XLOOKUP key (column B in the practice roster)
  const rosterFullNameColIndex = rosterHeaderRow.indexOf(CONFIG.columns.fullName) + 1;
  if (rosterFullNameColIndex === 0) {
    throw new Error(`${CONFIG.columns.fullName} column not found in Roster sheet`);
  }
  const rosterFullNameCol = getColumnLetter(rosterFullNameColIndex);
  console.log(`📍 Using ${CONFIG.columns.fullName} column ${rosterFullNameCol} for XLOOKUP key`);
  
  // Column 3: Grade
  const gradeColIndex = rosterHeaderRow.indexOf(CONFIG.columns.grade) + 1;
  if (gradeColIndex > 0) {
    const gradeCol = getColumnLetter(gradeColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${gradeCol}:${gradeCol}),"")`;
    newSheet.getRange(2, 3).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, 3).copyTo(newSheet.getRange(3, 3, numRows - 1, 1));
    }
    console.log(`✅ Populated Grade column with XLOOKUP`);
  }
  
  // Column 4: Gender (from "Gender Identification")
  const genderColIndex = rosterHeaderRow.indexOf(CONFIG.columns.genderIdentification) + 1;
  if (genderColIndex > 0) {
    const genderCol = getColumnLetter(genderColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${genderCol}:${genderCol}),"")`;
    newSheet.getRange(2, 4).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, 4).copyTo(newSheet.getRange(3, 4, numRows - 1, 1));
    }
    console.log(`✅ Populated Gender column with XLOOKUP from column ${genderCol}`);
  } else {
    console.warn(`⚠️ Gender Identification column not found in Roster sheet - available columns: ${rosterHeaderRow.join(', ')}`);
  }
  
  // Column 5: Team
  const teamColIndex = rosterHeaderRow.indexOf(CONFIG.columns.team) + 1;
  if (teamColIndex > 0) {
    const teamCol = getColumnLetter(teamColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${teamCol}:${teamCol}),"")`;
    newSheet.getRange(2, 5).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, 5).copyTo(newSheet.getRange(3, 5, numRows - 1, 1));
    }
    console.log(`✅ Populated Team column with XLOOKUP from column ${teamCol}`);
  } else {
    console.warn(`⚠️ Team column not found in Roster sheet - available columns: ${rosterHeaderRow.join(', ')}`);
  }
  
  // Column 6: Availability
  if (availColumns.availabilityColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${practiceAvailSheetName}'!A:A,'${practiceAvailSheetName}'!${availColumns.availabilityColumn}:${availColumns.availabilityColumn}),"")`;
    newSheet.getRange(2, 6).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, 6).copyTo(newSheet.getRange(3, 6, numRows - 1, 1));
    }
    console.log(`✅ Populated Availability column with XLOOKUP`);
  }
  
  // Column 7: Practice Availability Note
  if (availColumns.noteColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${practiceAvailSheetName}'!A:A,'${practiceAvailSheetName}'!${availColumns.noteColumn}:${availColumns.noteColumn}),"")`;
    newSheet.getRange(2, 7).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, 7).copyTo(newSheet.getRange(3, 7, numRows - 1, 1));
    }
    console.log(`✅ Populated Practice Availability Note column with XLOOKUP`);
  }
  
  // Add next game columns if available
  if (nextGameInfo && gameAvailabilitySheet) {
    const gameAvailSheetName = 'Game Availability';
    const nextGameColumns = findGameAvailabilityColumns(gameAvailabilitySheet, nextGameInfo.formattedDate);
    
    // Column 8: Next Game Availability
    if (nextGameColumns.availabilityColumn) {
      const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${nextGameColumns.availabilityColumn}:${nextGameColumns.availabilityColumn}),"")`;
      newSheet.getRange(2, 8).setFormula(formula);
      if (numRows > 1) {
        newSheet.getRange(2, 8).copyTo(newSheet.getRange(3, 8, numRows - 1, 1));
      }
      console.log(`✅ Populated Next Game Availability column (${nextGameInfo.formattedDate}) with XLOOKUP`);
    }
    
    // Column 9: Next Game Availability Note
    if (nextGameColumns.noteColumn) {
      const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${nextGameColumns.noteColumn}:${nextGameColumns.noteColumn}),"")`;
      newSheet.getRange(2, 9).setFormula(formula);
      if (numRows > 1) {
        newSheet.getRange(2, 9).copyTo(newSheet.getRange(3, 9, numRows - 1, 1));
      }
      console.log(`✅ Populated Next Game Availability Note column (${nextGameInfo.formattedDate} Note) with XLOOKUP`);
    }
  }
}

/**
 * Populate the # column with formulas that reset when Team or Gender changes
 * This must be called AFTER sorting to ensure correct formula references
 * @param {Sheet} sheet - The practice roster sheet
 * @param {number} numRows - Number of data rows
 */
function populateNumberColumn(sheet, numRows) {
  console.log(`🔢 Populating # column with reset formulas...`);
  
  // All rows get the formula: =IF(OR(D1<>D2,E1<>E2),1,A1+1)
  // If Team (col E) or Gender (col D) changes, reset to 1, otherwise increment
  const numberFormula = '=IF(OR(D1<>D2,E1<>E2),1,A1+1)';
  sheet.getRange(2, 1).setFormula(numberFormula);
  
  if (numRows > 1) {
    sheet.getRange(2, 1).copyTo(sheet.getRange(3, 1, numRows - 1, 1));
  }
  
  console.log(`✅ Populated # column with reset formula for ${numRows} rows`);
}

/**
 * Add black borders at the top of rows where groups change (# = 1)
 * @param {Sheet} sheet - The practice roster sheet
 * @param {number} numRows - Number of data rows
 */
function addGroupBorders(sheet, numRows) {
  console.log(`🎨 Adding group borders...`);
  
  // Get all values from the # column (already flushed before calling this function)
  const numberColumnValues = sheet.getRange(2, 1, numRows, 1).getValues();
  
  // Find rows where # = 1 (group starts)
  const groupStartRows = [];
  numberColumnValues.forEach((row, index) => {
    if (row[0] === 1) {
      groupStartRows.push(index + 2); // +2 because array is 0-based and data starts at row 2
    }
  });
  
  console.log(`Found ${groupStartRows.length} group starts at rows: ${groupStartRows.join(', ')}`);
  
  // Apply top border to each group start row (entire row)
  const numColumns = sheet.getLastColumn();
  groupStartRows.forEach(rowNum => {
    const range = sheet.getRange(rowNum, 1, 1, numColumns);
    range.setBorder(
      true, null, null, null,  // top border only
      null, null,              // no vertical borders
      'black',                 // color
      SpreadsheetApp.BorderStyle.SOLID // style
    );
  });
  
  console.log(`✅ Applied top borders to ${groupStartRows.length} group starts`);
}

/**
 * Sort the practice roster by Team, Gender, Full Name
 * Within each team, sorted by gender, and within each gender, sorted by name
 * Note: # column will automatically update after sort due to formula
 * @param {Sheet} sheet - The practice roster sheet
 * @param {number} numRows - Number of data rows
 * @param {number} numColumns - Number of columns
 */
function sortPracticeRoster(sheet, numRows, numColumns) {
  console.log(`🔄 Sorting ${numRows} rows by Team, Gender, Name...`);
  
  const dataRange = sheet.getRange(2, 1, numRows, numColumns);
  
  dataRange.sort([
    {column: 5, ascending: true},  // Team (column E - primary sort)
    {column: 4, ascending: true},  // Gender (column D - secondary sort)
    {column: 2, ascending: true}   // Full Name (column B - tertiary sort)
  ]);
  
  console.log('✅ Sorting complete');
}

/**
 * Find the next game date after a given practice date
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {string} practiceDate - Practice date in format "M/D"
 * @return {Object|null} Next game info object or null if not found
 */
function findNextGameAfterPractice(ss, practiceDate) {
  try {
    // Get game dates from Game Info sheet
    const gameDates = getDatesFromInfoSheet(ss, GAME_AVAILABILITY_CONFIG);
    
    if (gameDates.length === 0) {
      console.log('No game dates found in Game Info sheet');
      return null;
    }
    
    // Parse practice date
    const practiceDateParts = practiceDate.split('/');
    const practiceMonth = parseInt(practiceDateParts[0]);
    const practiceDay = parseInt(practiceDateParts[1]);
    
    // Create practice date object (assume current year)
    const currentYear = new Date().getFullYear();
    const practiceDateTime = new Date(currentYear, practiceMonth - 1, practiceDay);
    
    // Find the next game after this practice
    for (const gameDate of gameDates) {
      if (gameDate.date > practiceDateTime) {
        console.log(`Found next game: ${gameDate.formattedDate} after practice ${practiceDate}`);
        return gameDate;
      }
    }
    
    console.log(`No game found after practice ${practiceDate}`);
    return null;
    
  } catch (error) {
    console.error('Error finding next game after practice:', error);
    return null;
  }
}

/**
 * Find the availability columns in Game Availability sheet for a specific date
 * @param {Sheet} gameAvailabilitySheet - The Game Availability sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @return {Object} Object with availabilityColumn and noteColumn letters
 */
function findGameAvailabilityColumns(gameAvailabilitySheet, gameDate) {
  return findAvailabilityColumns(gameAvailabilitySheet, gameDate, 'Game Availability');
}

/**
 * Shared function to find availability columns in any availability sheet for a specific date
 * @param {Sheet} availabilitySheet - The availability sheet to search
 * @param {string} dateString - Date in format "M/D"
 * @param {string} sheetType - Type of sheet for logging (e.g., 'Practice Availability', 'Game Availability')
 * @return {Object} Object with availabilityColumn and noteColumn letters
 */
function findAvailabilityColumns(availabilitySheet, dateString, sheetType) {
  const headerRow = availabilitySheet.getRange(1, 1, 1, availabilitySheet.getLastColumn()).getValues()[0];
  
  let availabilityColumn = null;
  let noteColumn = null;
  
  console.log(`🔍 Looking for date "${dateString}" in ${sheetType} headers...`);
  
  headerRow.forEach((header, index) => {
    let headerStr = '';
    
    // Handle both Date objects and strings
    if (header instanceof Date) {
      // Convert Date object to M/D format
      const month = header.getMonth() + 1;
      const day = header.getDate();
      headerStr = `${month}/${day}`;
      console.log(`📅 Column ${index + 1}: Date object converted to "${headerStr}"`);
    } else {
      headerStr = header.toString().trim();
      console.log(`📅 Column ${index + 1}: "${headerStr}"`);
    }
    
    // Check for exact match with date
    if (headerStr === dateString) {
      availabilityColumn = getColumnLetter(index + 1);
      console.log(`✅ Found availability column: ${availabilityColumn} (${headerStr})`);
    }
    
    // Check for note column (date + " Note")
    if (headerStr === `${dateString} Note`) {
      noteColumn = getColumnLetter(index + 1);
      console.log(`✅ Found note column: ${noteColumn} (${headerStr})`);
    }
  });
  
  if (!availabilityColumn) {
    console.error(`❌ Date "${dateString}" not found in ${sheetType} sheet`);
    console.log(`Available headers: ${headerRow.map((h, i) => {
      if (h instanceof Date) {
        const m = h.getMonth() + 1;
        const d = h.getDate();
        return `${m}/${d}`;
      }
      return h.toString().trim();
    }).join(', ')}`);
  }
  
  return {
    availabilityColumn: availabilityColumn,
    noteColumn: noteColumn
  };
}

/**
 * Configure print settings for practice roster sheets
 * @param {Sheet} sheet - The practice roster sheet to configure
 */
function configurePrintSettings(sheet) {
  try {
    // Set print margins (in inches)
    sheet.setMargins(
      0.75,  // top margin
      0.25,  // bottom margin
      0.25,  // left margin
      0.25   // right margin
    );
    
    console.log('✅ Print margins configured: top=0.75", bottom=0.25", left=0.25", right=0.25"');
    
    // Note: Page orientation, scale, and alignment settings require manual configuration
    // in Google Sheets UI (File > Page setup) as they are not available via Apps Script API
    console.log('ℹ️ For optimal printing, manually set: Portrait orientation, Fit to height, Left/Top alignment');
    
  } catch (error) {
    console.warn('Could not configure print settings:', error);
  }
}

