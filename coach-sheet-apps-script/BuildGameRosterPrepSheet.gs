/**
 * Build Game Roster Prep Sheet Module
 * Creates game-specific prep sheets with availability data
 * Sorted by Gender > Availability > Name for optimal team organization
 */

/**
 * Main function to build a game roster prep sheet
 * Called from the menu
 */
function buildGameRosterPrepSheet() {
  console.log('🏆 Starting Build Game Roster Prep Sheet...');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get game dates from Game Info sheet
    const gameDates = getDatesFromInfoSheet(ss, GAME_AVAILABILITY_CONFIG);

    if (gameDates.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Game Dates Found',
        'No game dates found in "📍Game Info" sheet. Please ensure the sheet exists and contains game dates.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Find the next upcoming game (including today)
    const defaultGameIndex = findNextUpcomingGame(gameDates);

    // Show date selection dialog
    showGameDateSelectionDialog(gameDates, defaultGameIndex);

  } catch (error) {
    console.error('Error building game roster prep sheet:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to build game roster prep sheet: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Find the index of the next upcoming game (including today)
 * @param {Array} gameDates - Array of game date objects
 * @return {number} Index of the next upcoming game, or 0 if none found
 */
function findNextUpcomingGame(gameDates) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Reset to start of day for comparison

  for (let i = 0; i < gameDates.length; i++) {
    const gameDate = new Date(gameDates[i].date);
    gameDate.setHours(0, 0, 0, 0);

    if (gameDate >= today) {
      console.log(`📅 Next upcoming game: ${gameDates[i].formattedDate} (index ${i})`);
      return i;
    }
  }

  console.log('📅 No upcoming games found, defaulting to first game');
  return 0; // Default to first game if no upcoming ones
}

/**
 * Show the game date selection dialog
 * @param {Array} gameDates - Array of game date objects
 * @param {number} defaultIndex - Index of the default selected game
 */
function showGameDateSelectionDialog(gameDates, defaultIndex) {
  const html = createGameDateSelectionHtml(gameDates, defaultIndex);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(300);

  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Build Game Roster Prep Sheet');
}

/**
 * Create HTML for game date selection dialog
 * @param {Array} gameDates - Array of game date objects
 * @param {number} defaultIndex - Index of the default selected game
 * @return {string} HTML content
 */
function createGameDateSelectionHtml(gameDates, defaultIndex) {
  const defaultDate = gameDates[defaultIndex].formattedDate;
  const defaultSheetName = `${defaultDate} Game Roster Prep`;

  // Create dropdown options
  const dateOptions = gameDates.map((gd, index) => {
    const selected = index === defaultIndex ? 'selected' : '';
    return `<option value="${gd.formattedDate}" ${selected}>${gd.formattedDate}</option>`;
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
          select, input[type="text"] {
            width: 100%;
            padding: 10px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
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
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="gameDate">Select Game Date:</label>
          <select id="gameDate" onchange="updateSheetName()">
            ${dateOptions}
          </select>
          <div class="note">Choose the game date for this prep sheet</div>
        </div>

        <div class="form-group">
          <label for="sheetName">Sheet Name:</label>
          <input type="text" id="sheetName" value="${defaultSheetName}">
          <div class="note">Name for the new game roster prep sheet</div>
        </div>

        <div class="buttons">
          <button class="btn btn-primary" onclick="createGamePrepSheet()">Create Prep Sheet</button>
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>

        <div class="progress-overlay" id="progressOverlay">
          <div class="progress-content">
            <div class="spinner"></div>
            <div style="font-size: 16px; font-weight: bold; color: #333;">
              Building Game Roster Prep Sheet...
            </div>
            <div style="font-size: 14px; color: #666; margin-top: 8px;">
              Please wait while we create your prep sheet
            </div>
          </div>
        </div>

        <script>
          function updateSheetName() {
            const gameDate = document.getElementById('gameDate').value;
            document.getElementById('sheetName').value = gameDate + ' Game Roster Prep';
          }

          function createGamePrepSheet() {
            const gameDate = document.getElementById('gameDate').value;
            const sheetName = document.getElementById('sheetName').value.trim();

            if (!sheetName) {
              alert('Please enter a sheet name');
              return;
            }

            // Show progress
            document.getElementById('progressOverlay').style.display = 'block';

            // Check for duplicate sheet name first
            google.script.run
              .withSuccessHandler(function(isDuplicate) {
                if (isDuplicate) {
                  document.getElementById('progressOverlay').style.display = 'none';
                  alert('Sheet name "' + sheetName + '" already exists. Please choose a different name.');
                  return;
                }

                // If not duplicate, create the prep sheet
                google.script.run
                  .withSuccessHandler(onSuccess)
                  .withFailureHandler(onFailure)
                  .createGameRosterPrepSheet(sheetName, gameDate);
              })
              .withFailureHandler(onFailure)
              .isSheetNameDuplicate(sheetName);
          }

          function onSuccess(message) {
            document.getElementById('progressOverlay').style.display = 'none';
            google.script.host.close();
          }

          function onFailure(error) {
            document.getElementById('progressOverlay').style.display = 'none';
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Create the game roster prep sheet with all data
 * @param {string} sheetName - Name for the new sheet
 * @param {string} gameDate - Game date in format "M/D"
 */
function createGameRosterPrepSheet(sheetName, gameDate) {
  console.log(`🏆 Creating game roster prep sheet: "${sheetName}" for date: ${gameDate}`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Create new sheet
    const newSheet = ss.insertSheet(sheetName);
    console.log(`✅ Created new sheet: "${sheetName}"`);

    // Get source sheets
    const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
    const gameAvailabilitySheet = ss.getSheetByName('Game Availability');

    if (!rosterSheet) {
      throw new Error('Roster sheet not found');
    }

    if (!gameAvailabilitySheet) {
      throw new Error('Game Availability sheet not found');
    }

    // Find the availability columns for this game date
    const availColumns = findAvailabilityColumns(gameAvailabilitySheet, gameDate, 'Game Availability');

    if (!availColumns.availabilityColumn) {
      throw new Error(`Game date "${gameDate}" not found in Game Availability sheet`);
    }

    // Define column structure with shared base columns + dynamic availability columns
    const headers = [];

    // Add base columns in order defined by rosterPrintoutBaseColumns
    Object.keys(CONFIG.rosterPrintoutBaseColumns)
      .sort((a, b) => CONFIG.rosterPrintoutBaseColumns[a].index - CONFIG.rosterPrintoutBaseColumns[b].index)
      .forEach(key => {
        headers.push(CONFIG.rosterPrintoutBaseColumns[key].name);
      });

    // Add dynamic availability columns
    headers.push(gameDate);                              // Game availability
    headers.push(`${gameDate} Note`);                    // Game availability note

    // Set up headers
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');

    console.log(`📝 Set up ${headers.length} column headers: ${headers.join(', ')}`);
    console.log(`📍 Found availability columns: ${availColumns.availabilityColumn} and ${availColumns.noteColumn || 'none'}`);

    // Copy Full Name column to column B (column 2) from roster using shared utility
    const fullNameInfo = copyFullNameColumnToColumn(newSheet, rosterSheet, 2, 2);
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);

    const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];

    // Populate other columns with XLOOKUP formulas
    populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumns, fullNameInfo.rowCount);

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

    // Copy data validation from Game Availability for the Availability column using shared utility
    if (availColumns.availabilityColumn) {
      console.log('✅ Copying data validation from Game Availability...');
      const availColIndex = gameAvailabilitySheet.getRange(1, 1, 1, gameAvailabilitySheet.getLastColumn())
        .getValues()[0].findIndex(h => h === gameDate ||
          (h instanceof Date && `${h.getMonth() + 1}/${h.getDate()}` === gameDate)) + 1;

      if (availColIndex > 0) {
        const availabilityTargetColumn = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 1;
        copyDataValidation(newSheet, gameAvailabilitySheet,
          [{sourceColumn: gameDate, targetColumn: availabilityTargetColumn}], fullNameInfo.rowCount);
      }
    }

    // Force recalculation to ensure formulas are evaluated before sorting
    SpreadsheetApp.flush();

    // Sort the data AFTER formulas have been calculated (Gender > Availability > Name)
    if (fullNameInfo.rowCount > 0) {
      sortGameRosterPrep(newSheet, fullNameInfo.rowCount, headers.length);
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
    newSheet.autoResizeColumn(CONFIG.rosterPrintoutBaseColumns.number.index); // # column
    const gameAvailabilityColumnIndex = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 1;
    newSheet.autoResizeColumn(gameAvailabilityColumnIndex); // Game availability column

    // Enable text wrapping for note column
    console.log('📝 Enabling text wrap for note column...');
    const gameNoteColumnIndex = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 2;
    const noteColumnRange = newSheet.getRange(2, gameNoteColumnIndex, fullNameInfo.rowCount, 1);
    noteColumnRange.setWrap(true);

    // Set print settings
    console.log('🖨️ Configuring print settings...');
    configurePrintSettings(newSheet);

    console.log(`✅ Game roster prep sheet "${sheetName}" created successfully`);

    // Show success alert AFTER all processing is complete
    SpreadsheetApp.getUi().alert(
      'Game Roster Prep Sheet Created!',
      `Successfully created game roster prep sheet for ${gameDate} with ${fullNameInfo.rowCount} students.\n\nSorted by Gender > Availability > Name for optimal team organization.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return 'Success';

  } catch (error) {
    console.error('Error creating game roster prep sheet:', error);
    throw new Error(`Failed to create game roster prep sheet: ${error.message}`);
  }
}

/**
 * Populate game roster prep data with XLOOKUP formulas
 * @param {Sheet} newSheet - The new game roster prep sheet
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {Array} rosterHeaderRow - Header row from roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {Object} availColumns - Availability column letters
 * @param {number} numRows - Number of data rows
 */
function populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumns, numRows) {
  if (numRows === 0) return;

  const rosterSheetName = CONFIG.roster.sheetName;
  const gameAvailSheetName = 'Game Availability';

  // Find "Full Name" column for XLOOKUP key (column B in the game roster prep)
  const rosterFullNameColIndex = rosterHeaderRow.indexOf(CONFIG.columns.fullName) + 1;
  if (rosterFullNameColIndex === 0) {
    throw new Error(`${CONFIG.columns.fullName} column not found in Roster sheet`);
  }
  const rosterFullNameCol = getColumnLetter(rosterFullNameColIndex);
  console.log(`📍 Using ${CONFIG.columns.fullName} column ${rosterFullNameCol} for XLOOKUP key`);

  // Column at Team index: Team
  const teamColIndex = rosterHeaderRow.indexOf(CONFIG.columns.team) + 1;
  if (teamColIndex > 0) {
    const teamCol = getColumnLetter(teamColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${teamCol}:${teamCol}),"")`;
    newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.team.index).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.team.index).copyTo(newSheet.getRange(3, CONFIG.rosterPrintoutBaseColumns.team.index, numRows - 1, 1));
    }
    console.log(`✅ Populated Team column with XLOOKUP from column ${teamCol}`);
  } else {
    console.warn(`⚠️ Team column not found in Roster sheet - available columns: ${rosterHeaderRow.join(', ')}`);
  }

  // Column at Gender index: Gender (from "Gender Identification")
  const genderColIndex = rosterHeaderRow.indexOf(CONFIG.columns.genderIdentification) + 1;
  if (genderColIndex > 0) {
    const genderCol = getColumnLetter(genderColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${genderCol}:${genderCol}),"")`;
    newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.gender.index).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.gender.index).copyTo(newSheet.getRange(3, CONFIG.rosterPrintoutBaseColumns.gender.index, numRows - 1, 1));
    }
    console.log(`✅ Populated Gender column with XLOOKUP from column ${genderCol}`);
  } else {
    console.warn(`⚠️ Gender Identification column not found in Roster sheet - available columns: ${rosterHeaderRow.join(', ')}`);
  }

  // Column at Grade index: Grade
  const gradeColIndex = rosterHeaderRow.indexOf(CONFIG.columns.grade) + 1;
  if (gradeColIndex > 0) {
    const gradeCol = getColumnLetter(gradeColIndex);
    const formula = `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${gradeCol}:${gradeCol}),"")`;
    newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.grade.index).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, CONFIG.rosterPrintoutBaseColumns.grade.index).copyTo(newSheet.getRange(3, CONFIG.rosterPrintoutBaseColumns.grade.index, numRows - 1, 1));
    }
    console.log(`✅ Populated Grade column with XLOOKUP`);
  }

  // Game Availability column (first column after base columns)
  const availabilityColumnIndex = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 1;
  if (availColumns.availabilityColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.availabilityColumn}:${availColumns.availabilityColumn}),"")`;
    newSheet.getRange(2, availabilityColumnIndex).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, availabilityColumnIndex).copyTo(newSheet.getRange(3, availabilityColumnIndex, numRows - 1, 1));
    }
    console.log(`✅ Populated Game Availability column with XLOOKUP`);
  }

  // Game Availability Note column (second column after base columns)
  const noteColumnIndex = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 2;
  if (availColumns.noteColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.noteColumn}:${availColumns.noteColumn}),"")`;
    newSheet.getRange(2, noteColumnIndex).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, noteColumnIndex).copyTo(newSheet.getRange(3, noteColumnIndex, numRows - 1, 1));
    }
    console.log(`✅ Populated Game Availability Note column with XLOOKUP`);
  }
}

/**
 * Sort the game roster prep by Team ASC, Gender ASC, Availability ASC, Name ASC
 * @param {Sheet} sheet - The game roster prep sheet
 * @param {number} numRows - Number of data rows
 * @param {number} numColumns - Number of columns
 */
function sortGameRosterPrep(sheet, numRows, numColumns) {
  console.log(`🔄 Sorting ${numRows} rows by Team ASC, Gender ASC, Availability ASC, Name ASC...`);

  const dataRange = sheet.getRange(2, 1, numRows, numColumns);

  const availabilityColumnIndex = Object.keys(CONFIG.rosterPrintoutBaseColumns).length + 1; // First column after base columns

  dataRange.sort([
    {column: CONFIG.rosterPrintoutBaseColumns.team.index, ascending: true},       // Team - primary sort
    {column: CONFIG.rosterPrintoutBaseColumns.gender.index, ascending: true},     // Gender - secondary sort
    {column: availabilityColumnIndex, ascending: true},                           // Availability - tertiary sort
    {column: CONFIG.rosterPrintoutBaseColumns.fullName.index, ascending: true}    // Full Name - quaternary sort
  ]);

  console.log('✅ Sorting complete');
}

