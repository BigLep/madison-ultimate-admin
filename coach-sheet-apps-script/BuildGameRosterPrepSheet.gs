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
          .radio-group {
            display: flex;
            gap: 20px;
            margin-bottom: 15px;
          }
          .radio-option {
            display: flex;
            align-items: center;
            gap: 8px;
          }
          .radio-option input[type="radio"] {
            width: auto;
            margin: 0;
          }
          .radio-option label {
            margin: 0;
            font-weight: normal;
            display: inline;
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
          <label>Audience:</label>
          <div class="radio-group">
            <div class="radio-option">
              <input type="radio" id="coaches" name="audience" value="coaches" checked onchange="updateSheetName()">
              <label for="coaches">Coaches</label>
            </div>
            <div class="radio-option">
              <input type="radio" id="parents" name="audience" value="parents" onchange="updateSheetName()">
              <label for="parents">Parents</label>
            </div>
          </div>
          <div class="note">Choose the intended audience for this roster sheet</div>
        </div>

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
            const audience = document.querySelector('input[name="audience"]:checked').value;
            const suffix = audience === 'parents' ? ' Parent Roster' : ' Game Roster Prep';
            document.getElementById('sheetName').value = gameDate + suffix;
          }

          function createGamePrepSheet() {
            const gameDate = document.getElementById('gameDate').value;
            const sheetName = document.getElementById('sheetName').value.trim();
            const audience = document.querySelector('input[name="audience"]:checked').value;

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
                  .createGameRosterPrepSheet(sheetName, gameDate, audience);
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
 * @param {string} audience - Target audience: "coaches" or "parents"
 */
function createGameRosterPrepSheet(sheetName, gameDate, audience = 'coaches') {
  console.log(`🏆 Creating game roster prep sheet: "${sheetName}" for date: ${gameDate}, audience: ${audience}`);

  try {
    // Common setup for both audiences
    const { ss, newSheet, rosterSheet, gameAvailabilitySheet, availColumns } = setupGameRosterSheets(sheetName, gameDate);

    // Route to appropriate function based on audience
    if (audience === 'parents') {
      return buildParentGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumns);
    } else {
      return buildCoachGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumns);
    }
  } catch (error) {
    console.error('Error creating game roster prep sheet:', error);
    throw new Error(`Failed to create game roster prep sheet: ${error.message}`);
  }
}

/**
 * Get column layout for coach game roster prep based on CONFIG.gameRosterPrep.
 * Order: #, Full Name, [Team?], Gender, Grade, [$date Activation Status?], $date Availability, $date Note
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object} availColumns - From findAvailabilityColumns (headers for availability, note, activation)
 * @return {{ headers: string[], indices: Object }} headers array and 1-based column index for each logical column
 */
function getGameRosterPrepColumnLayout(gameDate, availColumns) {
  const hasTeam = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasTeam;
  const hasActivation = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasActivationStatus;
  const headers = [CONFIG.rosterPrintoutBaseColumns.number.name, CONFIG.rosterPrintoutBaseColumns.fullName.name];
  const indices = { number: 1, fullName: 2, team: null, gender: null, grade: null, activationStatus: null, availability: null, note: null };
  let col = 3;
  if (hasTeam) {
    headers.push(CONFIG.rosterPrintoutBaseColumns.team.name);
    indices.team = col++;
  }
  headers.push(CONFIG.rosterPrintoutBaseColumns.gender.name);
  indices.gender = col++;
  headers.push(CONFIG.rosterPrintoutBaseColumns.grade.name);
  indices.grade = col++;
  if (hasActivation && availColumns.activationHeader) {
    headers.push(availColumns.activationHeader);
    indices.activationStatus = col++;
  }
  headers.push(availColumns.availabilityHeader);
  indices.availability = col++;
  headers.push(availColumns.noteHeader);
  indices.note = col++;
  return { headers: headers, indices: indices };
}

/**
 * Common setup for game roster sheets
 * @param {string} sheetName - Name for the new sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @return {Object} Common resources needed by both roster types
 */
function setupGameRosterSheets(sheetName, gameDate) {
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

  return { ss, newSheet, rosterSheet, gameAvailabilitySheet, availColumns };
}

/**
 * Common cleanup and finalization for game roster sheets
 * @param {Sheet} sheet - The sheet to finalize
 * @param {number} rowCount - Number of data rows
 */
function finalizeGameRosterSheet(sheet, rowCount) {
  // Delete empty rows and columns to clean up the sheet
  console.log('🧹 Cleaning up empty rows and columns...');
  deleteEmptyRowsAndColumnsForSheet(sheet);

  console.log(`✅ Sheet finalized successfully with ${rowCount} students`);
}

/**
 * Apply availability data validation to game roster
 * @param {Sheet} newSheet - The sheet to apply validation to
 * @param {Sheet} gameAvailabilitySheet - Source sheet with validation
 * @param {Object} availColumns - From findAvailabilityColumns (must include availabilityHeader)
 * @param {number} targetColumn - Column index to apply validation to
 * @param {number} rowCount - Number of data rows
 */
function applyGameAvailabilityValidation(newSheet, gameAvailabilitySheet, availColumns, targetColumn, rowCount) {
  console.log('✅ Copying data validation from Game Availability...');
  if (!availColumns.availabilityHeader) return;
  copyDataValidation(newSheet, gameAvailabilitySheet,
    [{ sourceColumn: availColumns.availabilityHeader, targetColumn: targetColumn }], rowCount);
}

/**
 * Apply complete availability conditional formatting matching Game Availability sheet
 * @param {Sheet} sheet - The sheet to apply formatting to
 * @param {number} column - The column index with availability data
 * @param {number} rowCount - Number of data rows
 */
function applyAvailabilityConditionalFormatting(sheet, column, rowCount) {
  console.log('🎨 Applying availability conditional formatting...');

  const availabilityRange = sheet.getRange(2, column, rowCount, 1);
  const rules = [];

  // Light green for "👍 Planning to be there"
  const planningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('👍 Planning to be there')
    .setBackground('#b7e1cd')  // Light green matching Game Availability
    .setRanges([availabilityRange])
    .build();
  rules.push(planningRule);

  // Light red for "👎 Can't make it"
  const cantMakeItRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("👎 Can't make it")
    .setBackground('#f4c7c3')  // Light red matching Game Availability
    .setRanges([availabilityRange])
    .build();
  rules.push(cantMakeItRule);

  // Light gray for "❓ Not sure yet"
  const notSureRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('❓ Not sure yet')
    .setBackground('#cfe2f3')  // Light gray/blue matching Game Availability
    .setRanges([availabilityRange])
    .build();
  rules.push(notSureRule);

  // Green for "Was there"
  const wasThereRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Was there')
    .setBackground('#93c47d')  // Green matching Game Availability
    .setRanges([availabilityRange])
    .build();
  rules.push(wasThereRule);

  // Dark red for "Wasn't there"
  const wasntThereRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Wasn't there")
    .setBackground('#cc4125')  // Dark red matching Game Availability
    .setRanges([availabilityRange])
    .build();
  rules.push(wasntThereRule);

  // Apply all rules
  const existingRules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules(existingRules.concat(rules));
  console.log('✅ Applied conditional formatting for availability column');
}

/**
 * Apply Activation Status conditional formatting (Active=green, Inactive=red, TBD=grey) to a column.
 * Uses GAME_ACTIVATION_STATUS_OPTIONS from Availability.gs.
 * @param {Sheet} sheet - The sheet to format
 * @param {number} column - 1-based column index
 * @param {number} rowCount - Number of data rows
 */
function applyActivationStatusConditionalFormatting(sheet, column, rowCount) {
  const range = sheet.getRange(2, column, rowCount, 1);
  const rules = sheet.getConditionalFormatRules();
  GAME_ACTIVATION_STATUS_OPTIONS.forEach(function (opt) {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(opt.value)
      .setBackground(opt.backgroundColor)
      .setRanges([range])
      .build();
    rules.push(rule);
  });
  sheet.setConditionalFormatRules(rules);
  console.log('✅ Applied Activation Status conditional formatting');
}

/**
 * Build the coach version of game roster prep sheet
 * @param {Sheet} newSheet - The new sheet to populate
 * @param {Sheet} rosterSheet - The roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object} availColumns - Availability column info
 */
function buildCoachGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumns) {
  console.log(`🏆 Building COACH game roster for date: ${gameDate}`);

  try {
    const layout = getGameRosterPrepColumnLayout(gameDate, availColumns);
    const headers = layout.headers;
    const idx = layout.indices;

    // Set up headers
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');

    console.log(`📝 Set up ${headers.length} column headers: ${headers.join(', ')}`);
    console.log(`📍 Found availability columns: ${availColumns.availabilityColumn} and ${availColumns.noteColumn || 'none'}`);

    // Copy Full Name column to column 2 from roster using shared utility
    const fullNameInfo = copyFullNameColumnToColumn(newSheet, rosterSheet, 2, 2);
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);

    const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];

    // Populate other columns with XLOOKUP formulas (uses layout.indices)
    populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumns, layout.indices, fullNameInfo.rowCount);

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

    // Copy data validation from Game Availability for the Availability column
    if (availColumns.availabilityColumn && idx.availability) {
      applyGameAvailabilityValidation(newSheet, gameAvailabilitySheet, availColumns, idx.availability, fullNameInfo.rowCount);
      applyAvailabilityConditionalFormatting(newSheet, idx.availability, fullNameInfo.rowCount);
    }

    // Copy Activation Status validation/formatting from Game Availability if this season uses it
    if (idx.activationStatus && availColumns.activationStatusColumn) {
      copyDataValidation(newSheet, gameAvailabilitySheet,
        [{ sourceColumn: availColumns.activationHeader, targetColumn: idx.activationStatus }], fullNameInfo.rowCount);
      applyActivationStatusConditionalFormatting(newSheet, idx.activationStatus, fullNameInfo.rowCount);
    }

    // Force recalculation to ensure formulas are evaluated before sorting
    SpreadsheetApp.flush();

    // Sort: Activation Status (if present) > Gender > Availability > Name
    if (fullNameInfo.rowCount > 0) {
      sortGameRosterPrep(newSheet, fullNameInfo.rowCount, headers.length, layout.indices);
    }

    // Populate # column AFTER sorting (reset when Activation Status or Gender changes)
    if (fullNameInfo.rowCount > 0) {
      const groupByCols = [idx.activationStatus, idx.gender].filter(Boolean);
      populateNumberColumn(newSheet, fullNameInfo.rowCount, groupByCols);

      // Force calculation of # column formulas before adding borders
      SpreadsheetApp.flush();

      // Add borders at group changes (where # = 1)
      addGroupBorders(newSheet, fullNameInfo.rowCount);
    }

    // Common cleanup
    finalizeGameRosterSheet(newSheet, fullNameInfo.rowCount);

    // Auto-resize columns
    console.log('📏 Auto-resizing columns...');
    newSheet.autoResizeColumn(idx.number);
    if (idx.availability) newSheet.autoResizeColumn(idx.availability);
    if (idx.activationStatus) newSheet.autoResizeColumn(idx.activationStatus);

    // Enable text wrapping for note column
    if (idx.note) {
      console.log('📝 Enabling text wrap for note column...');
      const noteColumnRange = newSheet.getRange(2, idx.note, fullNameInfo.rowCount, 1);
      noteColumnRange.setWrap(true);
    }

    // Set print settings
    console.log('🖨️ Configuring print settings...');
    configurePrintSettings(newSheet);

    console.log(`✅ Game roster prep sheet created successfully`);

    const sortDesc = idx.activationStatus
      ? 'Activation Status > Gender > Availability > Name'
      : 'Gender > Availability > Name';
    SpreadsheetApp.getUi().alert(
      'Game Roster Prep Sheet Created!',
      `Successfully created game roster prep sheet for ${gameDate} with ${fullNameInfo.rowCount} students.\n\nSorted by ${sortDesc}.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return 'Success';

  } catch (error) {
    console.error('Error creating game roster prep sheet:', error);
    throw new Error(`Failed to create game roster prep sheet: ${error.message}`);
  }
}

/**
 * Populate game roster prep data with XLOOKUP formulas.
 * Uses indices from getGameRosterPrepColumnLayout (optional Team, optional Activation Status).
 * @param {Sheet} newSheet - The new game roster prep sheet
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {Array} rosterHeaderRow - Header row from roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {Object} availColumns - Availability column letters and headers
 * @param {Object} indices - 1-based column indices (number, fullName, team?, gender, grade, activationStatus?, availability, note)
 * @param {number} numRows - Number of data rows
 */
function populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumns, indices, numRows) {
  if (numRows === 0) return;

  const rosterSheetName = CONFIG.roster.sheetName;
  const gameAvailSheetName = 'Game Availability';

  const rosterFullNameColIndex = rosterHeaderRow.indexOf(CONFIG.columns.fullName) + 1;
  if (rosterFullNameColIndex === 0) {
    throw new Error(`${CONFIG.columns.fullName} column not found in Roster sheet`);
  }
  const rosterFullNameCol = getColumnLetter(rosterFullNameColIndex);
  console.log(`📍 Using ${CONFIG.columns.fullName} column ${rosterFullNameCol} for XLOOKUP key`);

  function setFormula(colIndex, formula) {
    if (!colIndex) return;
    newSheet.getRange(2, colIndex).setFormula(formula);
    if (numRows > 1) {
      newSheet.getRange(2, colIndex).copyTo(newSheet.getRange(3, colIndex, numRows - 1, 1));
    }
  }

  // Team (only if this season has team)
  if (indices.team) {
    const teamColIndex = rosterHeaderRow.indexOf(CONFIG.columns.team) + 1;
    if (teamColIndex > 0) {
      const teamCol = getColumnLetter(teamColIndex);
      setFormula(indices.team, `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${teamCol}:${teamCol}),"")`);
      console.log(`✅ Populated Team column with XLOOKUP`);
    }
  }

  // Gender
  if (indices.gender) {
    const genderColIndex = rosterHeaderRow.indexOf(CONFIG.columns.genderIdentification) + 1;
    if (genderColIndex > 0) {
      const genderCol = getColumnLetter(genderColIndex);
      setFormula(indices.gender, `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${genderCol}:${genderCol}),"")`);
      console.log(`✅ Populated Gender column with XLOOKUP`);
    }
  }

  // Grade
  if (indices.grade) {
    const gradeColIndex = rosterHeaderRow.indexOf(CONFIG.columns.grade) + 1;
    if (gradeColIndex > 0) {
      const gradeCol = getColumnLetter(gradeColIndex);
      setFormula(indices.grade, `=IFERROR(XLOOKUP(B2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${gradeCol}:${gradeCol}),"")`);
      console.log(`✅ Populated Grade column with XLOOKUP`);
    }
  }

  // Activation Status (only if this season has it)
  if (indices.activationStatus && availColumns.activationStatusColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.activationStatusColumn}:${availColumns.activationStatusColumn}),"")`;
    setFormula(indices.activationStatus, formula);
    console.log(`✅ Populated Activation Status column with XLOOKUP`);
  }

  // Game Availability
  if (indices.availability && availColumns.availabilityColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.availabilityColumn}:${availColumns.availabilityColumn}),"")`;
    setFormula(indices.availability, formula);
    console.log(`✅ Populated Game Availability column with XLOOKUP`);
  }

  // Game Note
  if (indices.note && availColumns.noteColumn) {
    const formula = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.noteColumn}:${availColumns.noteColumn}),"")`;
    setFormula(indices.note, formula);
    console.log(`✅ Populated Game Note column with XLOOKUP`);
  }
}

/**
 * Sort the game roster prep: Activation Status (if present) > Gender > Availability > Name.
 * @param {Sheet} sheet - The game roster prep sheet
 * @param {number} numRows - Number of data rows
 * @param {number} numColumns - Number of columns
 * @param {Object} indices - 1-based column indices from getGameRosterPrepColumnLayout
 */
function sortGameRosterPrep(sheet, numRows, numColumns, indices) {
  const sortSpec = [];
  if (indices.activationStatus) {
    sortSpec.push({ column: indices.activationStatus, ascending: true });
    console.log(`🔄 Sorting ${numRows} rows by Activation Status > Gender > Availability > Name...`);
  } else {
    console.log(`🔄 Sorting ${numRows} rows by Gender > Availability > Name...`);
  }
  sortSpec.push({ column: indices.gender, ascending: true });
  if (indices.availability) {
    sortSpec.push({ column: indices.availability, ascending: true });
  }
  sortSpec.push({ column: indices.fullName, ascending: true });

  const dataRange = sheet.getRange(2, 1, numRows, numColumns);
  dataRange.sort(sortSpec);
  console.log('✅ Sorting complete');
}

/**
 * Build a parent-friendly game roster with just names and availability
 * @param {Sheet} newSheet - The new sheet to populate
 * @param {Sheet} rosterSheet - The roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object} availColumns - Availability column info
 */
function buildParentGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumns) {
  console.log(`👨‍👩‍👧‍👦 Building PARENT game roster for date: ${gameDate}`);

  try {
    const hasTeam = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasTeam;
    const hasActivation = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasActivationStatus;

    // Column order: Full Name (1), [Activation Status?], Availability, [Team if hasTeam]
    const headers = ['Full Name'];
    const col = { fullName: 1, activation: null, availability: null, team: null };
    let c = 2;
    if (hasActivation && availColumns.activationHeader) {
      headers.push(availColumns.activationHeader);
      col.activation = c++;
    }
    headers.push(availColumns.availabilityHeader);
    col.availability = c++;
    if (hasTeam) {
      headers.push('Team');
      col.team = c++;
    }

    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');

    console.log(`📝 Set up ${headers.length} column headers: ${headers.join(', ')}`);

    const fullNameInfo = copyFullNameColumnToColumn(newSheet, rosterSheet, 2, 1);
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);

    if (fullNameInfo.rowCount > 0) {
      const gameAvailSheetName = 'Game Availability';
      const rosterSheetName = CONFIG.roster.sheetName;

      function setFormula(colIndex, formula) {
        if (!colIndex) return;
        newSheet.getRange(2, colIndex).setFormula(formula);
        if (fullNameInfo.rowCount > 1) {
          newSheet.getRange(2, colIndex).copyTo(newSheet.getRange(3, colIndex, fullNameInfo.rowCount - 1, 1));
        }
      }

      if (col.activation && availColumns.activationStatusColumn) {
        setFormula(col.activation, `=IFERROR(XLOOKUP(A2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.activationStatusColumn}:${availColumns.activationStatusColumn}),"")`);
        console.log(`✅ Populated Activation Status column with XLOOKUP`);
      }

      if (col.availability && availColumns.availabilityColumn) {
        setFormula(col.availability, `=IFERROR(XLOOKUP(A2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.availabilityColumn}:${availColumns.availabilityColumn}),"")`);
        console.log(`✅ Populated Game Availability column with XLOOKUP`);
      }

      if (col.team) {
        const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
        const rosterFullNameColIndex = rosterHeaderRow.indexOf(CONFIG.columns.fullName) + 1;
        const teamColIndex = rosterHeaderRow.indexOf(CONFIG.columns.team) + 1;
        if (teamColIndex > 0 && rosterFullNameColIndex > 0) {
          const rosterFullNameCol = getColumnLetter(rosterFullNameColIndex);
          const teamCol = getColumnLetter(teamColIndex);
          setFormula(col.team, `=IFERROR(XLOOKUP(A2,'${rosterSheetName}'!${rosterFullNameCol}:${rosterFullNameCol},'${rosterSheetName}'!${teamCol}:${teamCol}),"")`);
          console.log(`✅ Populated Team column with XLOOKUP`);
        }
      }
    }

    SpreadsheetApp.flush();

    const numCols = headers.length;
    if (fullNameInfo.rowCount > 0) {
      const sortSpec = col.activation
        ? [{ column: col.activation, ascending: true }, { column: 1, ascending: true }]
        : [{ column: 1, ascending: true }];
      console.log(col.activation
        ? `🔄 Sorting ${fullNameInfo.rowCount} rows by Activation Status > Player Name...`
        : `🔄 Sorting ${fullNameInfo.rowCount} rows by player name...`);
      const dataRange = newSheet.getRange(2, 1, fullNameInfo.rowCount, numCols);
      dataRange.sort(sortSpec);
      console.log('✅ Sorting complete');
    }

    // Filter to hide Practice Squad and Dropped only when we have a Team column
    if (hasTeam && fullNameInfo.rowCount > 0 && col.team) {
      console.log('🔍 Applying filter to hide Practice Squad and Dropped...');
      const fullDataRange = newSheet.getRange(1, 1, fullNameInfo.rowCount + 1, numCols);
      const filter = fullDataRange.createFilter();
      const criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(['Practice Squad', 'Dropped'])
        .build();
      filter.setColumnFilterCriteria(col.team, criteria);
      console.log('✅ Filter applied - Practice Squad and Dropped hidden');
    }

    if (fullNameInfo.rowCount > 0) {
      if (col.activation) {
        copyDataValidation(newSheet, gameAvailabilitySheet,
          [{ sourceColumn: availColumns.activationHeader, targetColumn: col.activation }], fullNameInfo.rowCount);
        applyActivationStatusConditionalFormatting(newSheet, col.activation, fullNameInfo.rowCount);
      }
      if (col.availability) {
        applyAvailabilityConditionalFormatting(newSheet, col.availability, fullNameInfo.rowCount);
        if (availColumns.availabilityColumn) {
          applyGameAvailabilityValidation(newSheet, gameAvailabilitySheet, availColumns, col.availability, fullNameInfo.rowCount);
        }
      }
    }

    finalizeGameRosterSheet(newSheet, fullNameInfo.rowCount);

    console.log('📏 Auto-resizing columns...');
    newSheet.autoResizeColumn(1);
    if (col.activation) newSheet.autoResizeColumn(col.activation);
    if (col.availability) newSheet.autoResizeColumn(col.availability);
    if (col.team) newSheet.autoResizeColumn(col.team);

    console.log(`✅ Parent game roster created successfully`);

    const alertDetail = col.activation
      ? (hasTeam ? 'Sorted by Activation Status > Player Name. Practice Squad and Dropped players are hidden (filter applied).' : 'Sorted by Activation Status > Player Name.')
      : (hasTeam ? 'Sorted by player name. Practice Squad and Dropped players are hidden (filter applied).' : 'Sorted by player name.');
    SpreadsheetApp.getUi().alert(
      'Parent Game Roster Created!',
      `Successfully created parent game roster for ${gameDate} with ${fullNameInfo.rowCount} students.\n\n${alertDetail}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return 'Success';

  } catch (error) {
    console.error('Error creating parent game roster:', error);
    throw new Error(`Failed to create parent game roster: ${error.message}`);
  }
}

