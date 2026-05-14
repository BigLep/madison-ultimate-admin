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
  const defaultGd = gameDates[defaultIndex];
  const defaultBase = defaultGd.formattedDate + (defaultGd.gameLabel ? ' ' + defaultGd.gameLabel : '');
  const defaultSheetName = `${defaultBase} Game Roster Prep`;

  const GAME_DATE_OPTIONS_JSON = JSON.stringify(gameDates.map(function (gd) {
    return {
      formattedDate: gd.formattedDate,
      gameLabel: gd.gameLabel || '',
      ordinalForDate: gd.ordinalForDate || 1
    };
  }));

  function escHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Create dropdown options (one row per Game Info line; double-headers get distinct entries)
  const dateOptions = gameDates.map((gd, index) => {
    const selected = index === defaultIndex ? 'selected' : '';
    let label = gd.formattedDate;
    if (gd.gameLabel) label += ' · ' + gd.gameLabel;
    if (gd.ordinalForDate > 1) label += ' (Game ' + gd.ordinalForDate + ')';
    return `<option value="${index}" ${selected}>${escHtml(label)}</option>`;
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
          <label for="gameDate">Select Game:</label>
          <select id="gameDate" onchange="updateSheetName()">
            ${dateOptions}
          </select>
          <div class="note">Pick any row for that calendar day—the prep sheet lists every game that day (activation, availability, and note for each).</div>
        </div>

        <div class="form-group">
          <label for="sheetName">Sheet Name:</label>
          <input type="text" id="sheetName" value="${escHtml(defaultSheetName)}">
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
          const GAME_DATE_OPTIONS = ${GAME_DATE_OPTIONS_JSON};

          function updateSheetName() {
            const idx = parseInt(document.getElementById('gameDate').value, 10);
            const gd = GAME_DATE_OPTIONS[idx];
            if (!gd) return;
            const audience = document.querySelector('input[name="audience"]:checked').value;
            const suffix = audience === 'parents' ? ' Parent Roster' : ' Game Roster Prep';
            var base = gd.formattedDate;
            if (gd.gameLabel) base += ' ' + gd.gameLabel;
            document.getElementById('sheetName').value = base + suffix;
          }

          function createGamePrepSheet() {
            const idx = parseInt(document.getElementById('gameDate').value, 10);
            const sheetName = document.getElementById('sheetName').value.trim();
            const audience = document.querySelector('input[name="audience"]:checked').value;

            if (!sheetName) {
              alert('Please enter a sheet name');
              return;
            }

            if (isNaN(idx)) {
              alert('Please select a game');
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
                  .createGameRosterPrepSheet(sheetName, idx, audience);
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
 * @param {number} gameRowIndex - Index into getDatesFromInfoSheet(...) result (one entry per Game Info row)
 * @param {string} audience - Target audience: "coaches" or "parents"
 */
function createGameRosterPrepSheet(sheetName, gameRowIndex, audience = 'coaches') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameDates = getDatesFromInfoSheet(ss, GAME_AVAILABILITY_CONFIG);
  const idx = typeof gameRowIndex === 'string' ? parseInt(gameRowIndex, 10) : gameRowIndex;
  if (isNaN(idx) || idx < 0 || idx >= gameDates.length) {
    throw new Error('Invalid game selection');
  }
  const selected = gameDates[idx];
  const gameAvailabilitySheet = ss.getSheetByName('Game Availability');
  if (!gameAvailabilitySheet) {
    throw new Error('Game Availability sheet not found');
  }
  const availColumnsList = buildAvailColumnsListForSelectedCalendarDay_(gameAvailabilitySheet, gameDates, idx);

  const sameDayCount = gameDates.filter(function (g) {
    return g.formattedDate === selected.formattedDate;
  }).length;
  var displayLabel = selected.formattedDate;
  if (sameDayCount > 1) {
    displayLabel += ' · ' + sameDayCount + ' games';
  } else if (selected.gameLabel) {
    displayLabel += ' · ' + selected.gameLabel;
  }

  console.log(`🏆 Creating game roster prep sheet: "${sheetName}" for ${displayLabel} (${availColumnsList.length} game column set(s)), audience: ${audience}`);

  try {
    const { newSheet, rosterSheet, gameAvailabilitySheet } = setupGameRosterSheets(sheetName, availColumnsList);

    if (audience === 'parents') {
      return buildParentGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, selected.formattedDate, availColumnsList, displayLabel);
    }
    return buildCoachGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, selected.formattedDate, availColumnsList, displayLabel);
  } catch (error) {
    console.error('Error creating game roster prep sheet:', error);
    throw new Error(`Failed to create game roster prep sheet: ${error.message}`);
  }
}

/**
 * All games on the same calendar day as gameDates[selectedIndex] (ordinal order). One findAvailabilityColumns result per game.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} gameAvailabilitySheet
 * @param {Array} gameDates from getDatesFromInfoSheet
 * @param {number} selectedIndex
 * @return {Object[]}
 */
function buildAvailColumnsListForSelectedCalendarDay_(gameAvailabilitySheet, gameDates, selectedIndex) {
  const selected = gameDates[selectedIndex];
  if (!selected) throw new Error('Invalid game selection');
  const dateKey = selected.formattedDate;
  const rows = [];
  for (var i = 0; i < gameDates.length; i++) {
    if (gameDates[i].formattedDate === dateKey) rows.push(gameDates[i]);
  }
  rows.sort(function (a, b) {
    return (a.ordinalForDate || 1) - (b.ordinalForDate || 1);
  });
  const list = [];
  for (var j = 0; j < rows.length; j++) {
    var ord = rows[j].ordinalForDate || 1;
    var ac = findAvailabilityColumns(gameAvailabilitySheet, dateKey, 'Game Availability', ord);
    if (!ac.availabilityColumn) {
      var expect = getAvailabilityColumnHeaders(dateKey, 'Game Availability', ord).availabilityHeader;
      throw new Error('Game "' + dateKey + '" (game ' + ord + ' on that date) not found in Game Availability. Expected column "' + expect + '". Run Build Game Availability first.');
    }
    list.push(ac);
  }
  return list;
}

/**
 * Get column layout for coach game roster prep based on CONFIG.gameRosterPrep.
 * Order: #, Full Name, [Team?], Gender, Grade, then for each game that calendar day:
 *   [$date Activation Status?], $date Availability, $date Note (and "(Game N)" variants).
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object[]} availColumnsList - One findAvailabilityColumns result per game on that day
 * @return {{ headers: string[], indices: Object }} headers array and 1-based column index for each logical column
 */
function getGameRosterPrepColumnLayout(gameDate, availColumnsList) {
  const hasTeam = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasTeam;
  const hasActivation = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasActivationStatus;
  const headers = [CONFIG.rosterPrintoutBaseColumns.number.name, CONFIG.rosterPrintoutBaseColumns.fullName.name];
  const indices = {
    number: 1,
    fullName: 2,
    team: null,
    gender: null,
    grade: null,
    games: [],
    activationStatus: null,
    availability: null,
    note: null
  };
  let col = 3;
  if (hasTeam) {
    headers.push(CONFIG.rosterPrintoutBaseColumns.team.name);
    indices.team = col++;
  }
  headers.push(CONFIG.rosterPrintoutBaseColumns.gender.name);
  indices.gender = col++;
  headers.push(CONFIG.rosterPrintoutBaseColumns.grade.name);
  indices.grade = col++;

  for (var gi = 0; gi < availColumnsList.length; gi++) {
    var availColumns = availColumnsList[gi];
    var gameIndices = { activation: null, availability: null, note: null };
    if (hasActivation && availColumns.activationHeader && availColumns.activationStatusColumn) {
      headers.push(availColumns.activationHeader);
      gameIndices.activation = col++;
    }
    headers.push(availColumns.availabilityHeader);
    gameIndices.availability = col++;
    headers.push(availColumns.noteHeader);
    gameIndices.note = col++;
    indices.games.push(gameIndices);
  }

  if (indices.games.length > 0) {
    indices.activationStatus = indices.games[0].activation;
    indices.availability = indices.games[0].availability;
    indices.note = indices.games[0].note;
  }
  return { headers: headers, indices: indices };
}

/**
 * Common setup for game roster sheets
 * @param {string} sheetName - Name for the new sheet
 * @param {Object[]} availColumnsList - One findAvailabilityColumns result per game on the selected calendar day
 * @return {Object} Common resources needed by both roster types
 */
function setupGameRosterSheets(sheetName, availColumnsList) {
  if (!availColumnsList || availColumnsList.length === 0) {
    throw new Error('No game availability columns to build roster.');
  }

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

  return { ss: ss, newSheet: newSheet, rosterSheet: rosterSheet, gameAvailabilitySheet: gameAvailabilitySheet, availColumnsList: availColumnsList };
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

/** Gender labels for coach roster activation summary COUNTIFS (must match Gender column values, e.g. Gender Identification). */
const COACH_GAME_ROSTER_SUMMARY_GENDERS = ['Bx', 'Gx'];

/**
 * Append a COUNTIFS summary below player rows: Active / Inactive / TBD × Bx / Gx.
 * Placed one blank row after data; table in columns B–D (labels in B, Bx in C, Gx in D).
 * No-op if Activation Status column is not on the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} numRows - Player data rows (rows 2 .. numRows+1)
 * @param {Object} idx - layout.indices from getGameRosterPrepColumnLayout
 * @return {number|undefined} Last row of this summary (for placing the next table), or undefined if skipped
 */
function addCoachGameRosterActivationSummary(sheet, numRows, idx) {
  if (!numRows || !idx.activationStatus || !idx.gender) {
    return undefined;
  }

  const lastDataRow = 1 + numRows;
  const genderLetter = getColumnLetter(idx.gender);
  const activationLetter = getColumnLetter(idx.activationStatus);
  const labelCol = 2; // B
  const bxCol = 3; // C
  const gxCol = 4; // D
  const labelColLetter = getColumnLetter(labelCol);
  const bxColLetter = getColumnLetter(bxCol);
  const gxColLetter = getColumnLetter(gxCol);

  const headerRow = numRows + 3;
  sheet.getRange(headerRow, labelCol).setValue('Independent of status');
  sheet.getRange(headerRow, bxCol).setValue(COACH_GAME_ROSTER_SUMMARY_GENDERS[0]);
  sheet.getRange(headerRow, gxCol).setValue(COACH_GAME_ROSTER_SUMMARY_GENDERS[1]);

  const statusLabels = GAME_ACTIVATION_STATUS_OPTIONS.map(function (opt) { return opt.value; });
  for (let i = 0; i < statusLabels.length; i++) {
    const row = headerRow + 1 + i;
    sheet.getRange(row, labelCol).setValue(statusLabels[i]);
    const formulaBx = '=COUNTIFS($' + genderLetter + '$2:$' + genderLetter + '$' + lastDataRow + ',' +
      '$' + bxColLetter + '$' + headerRow + ',$' + activationLetter + '$2:$' + activationLetter + '$' + lastDataRow + ',' +
      '$' + labelColLetter + '$' + row + ')';
    const formulaGx = '=COUNTIFS($' + genderLetter + '$2:$' + genderLetter + '$' + lastDataRow + ',' +
      '$' + gxColLetter + '$' + headerRow + ',$' + activationLetter + '$2:$' + activationLetter + '$' + lastDataRow + ',' +
      '$' + labelColLetter + '$' + row + ')';
    sheet.getRange(row, bxCol).setFormula(formulaBx);
    sheet.getRange(row, gxCol).setFormula(formulaGx);
  }

  const summaryEndRow = headerRow + statusLabels.length;
  sheet.getRange(headerRow, labelCol, headerRow, gxCol).setFontWeight('bold');
  sheet.getRange(headerRow + 1, labelCol, summaryEndRow, labelCol).setFontWeight('bold');
  console.log('✅ Added coach activation summary table at rows ' + headerRow + '–' + summaryEndRow);
  return summaryEndRow;
}

/**
 * Second summary: same activation × gender counts but only rows where game availability is not "👎 Can't make it"
 * (Planning, Not sure, blank, Was there, etc.). One blank row after the first summary table.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} numRows - Player data rows (rows 2 .. numRows+1)
 * @param {Object} idx - layout.indices from getGameRosterPrepColumnLayout
 * @param {number} firstSummaryEndRow - Last row of the Independent-of-status table (data row for TBD)
 */
function addCoachGameRosterExcludingCantMakeItSummary(sheet, numRows, idx, firstSummaryEndRow) {
  if (!numRows || !idx.activationStatus || !idx.gender || !idx.availability) {
    return;
  }

  const lastDataRow = 1 + numRows;
  const genderLetter = getColumnLetter(idx.gender);
  const activationLetter = getColumnLetter(idx.activationStatus);
  const availabilityLetter = getColumnLetter(idx.availability);
  const labelCol = 2;
  const bxCol = 3;
  const gxCol = 4;
  const labelColLetter = getColumnLetter(labelCol);
  const bxColLetter = getColumnLetter(bxCol);
  const gxColLetter = getColumnLetter(gxCol);

  const headerRow2 = firstSummaryEndRow + 2;
  sheet.getRange(headerRow2, labelCol).setValue('Status: Planning to be there');
  sheet.getRange(headerRow2, bxCol).setValue(COACH_GAME_ROSTER_SUMMARY_GENDERS[0]);
  sheet.getRange(headerRow2, gxCol).setValue(COACH_GAME_ROSTER_SUMMARY_GENDERS[1]);

  const statusLabels = GAME_ACTIVATION_STATUS_OPTIONS.map(function (opt) { return opt.value; });
  const availNotCantMakeIt = '"<>👎 Can\'t make it"';

  for (let i = 0; i < statusLabels.length; i++) {
    const row = headerRow2 + 1 + i;
    sheet.getRange(row, labelCol).setValue(statusLabels[i]);
    const formulaBx =
      '=COUNTIFS($' +
      genderLetter +
      '$2:$' +
      genderLetter +
      '$' +
      lastDataRow +
      ',$' +
      bxColLetter +
      '$' +
      headerRow2 +
      ',$' +
      activationLetter +
      '$2:$' +
      activationLetter +
      '$' +
      lastDataRow +
      ',$' +
      labelColLetter +
      '$' +
      row +
      ',$' +
      availabilityLetter +
      '$2:$' +
      availabilityLetter +
      '$' +
      lastDataRow +
      ',' +
      availNotCantMakeIt +
      ')';
    const formulaGx =
      '=COUNTIFS($' +
      genderLetter +
      '$2:$' +
      genderLetter +
      '$' +
      lastDataRow +
      ',$' +
      gxColLetter +
      '$' +
      headerRow2 +
      ',$' +
      activationLetter +
      '$2:$' +
      activationLetter +
      '$' +
      lastDataRow +
      ',$' +
      labelColLetter +
      '$' +
      row +
      ',$' +
      availabilityLetter +
      '$2:$' +
      availabilityLetter +
      '$' +
      lastDataRow +
      ',' +
      availNotCantMakeIt +
      ')';
    sheet.getRange(row, bxCol).setFormula(formulaBx);
    sheet.getRange(row, gxCol).setFormula(formulaGx);
  }

  const summary2EndRow = headerRow2 + statusLabels.length;
  sheet.getRange(headerRow2, labelCol, headerRow2, gxCol).setFontWeight('bold');
  sheet.getRange(headerRow2 + 1, labelCol, summary2EndRow, labelCol).setFontWeight('bold');
  console.log('✅ Added coach summary (excluding Can\'t make it) at rows ' + headerRow2 + '–' + summary2EndRow);
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
 * Build the coach version of game roster prep sheet
 * @param {Sheet} newSheet - The new sheet to populate
 * @param {Sheet} rosterSheet - The roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object} availColumns - Availability column info
 * @param {string} displayLabel - Short label for dialogs (date plus optional game label)
 */
function buildCoachGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumnsList, displayLabel) {
  const labelForUi = displayLabel || gameDate;
  console.log(`🏆 Building COACH game roster for: ${labelForUi}`);

  try {
    const layout = getGameRosterPrepColumnLayout(gameDate, availColumnsList);
    const headers = layout.headers;
    const idx = layout.indices;

    // Set up headers
    const headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');

    console.log(`📝 Set up ${headers.length} column headers: ${headers.join(', ')}`);
    console.log(`📍 ${availColumnsList.length} game(s) on this calendar day — availability columns: ${availColumnsList.map(function (a) { return a.availabilityColumn; }).join(', ')}`);

    // Copy Full Name column to column 2 from roster using shared utility
    const fullNameInfo = copyFullNameColumnToColumn(newSheet, rosterSheet, 2, 2);
    console.log(`📊 Copied ${fullNameInfo.rowCount} students from roster`);

    const rosterHeaderRow = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];

    // Populate other columns with XLOOKUP formulas (uses layout.indices)
    populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumnsList, layout.indices, fullNameInfo.rowCount);

    // Copy formatting from roster using shared utility
    console.log('🎨 Copying column formatting...');
    copyColumnFormatting(newSheet, rosterSheet, headers, rosterHeaderRow);

    // Apply Format Spruce Up silently
    console.log('✨ Applying Format Spruce Up formatting...');
    applySpruceUpFormatting(newSheet);

    // Ensure header row styling is preserved using shared utility
    styleHeaderRow(newSheet, headers.length);

    // Whole-sheet value colors (same managed rules as Game Availability)
    console.log('🎨 Applying managed availability / activation conditional formatting (whole sheet)...');
    refreshManagedAvailabilityAndActivationCfOnSheet(newSheet);

    for (var gi = 0; gi < availColumnsList.length; gi++) {
      var ac = availColumnsList[gi];
      var gix = idx.games[gi];
      if (!gix) continue;
      if (ac.availabilityColumn && gix.availability) {
        applyGameAvailabilityValidation(newSheet, gameAvailabilitySheet, ac, gix.availability, fullNameInfo.rowCount);
      }
      if (gix.activation && ac.activationStatusColumn) {
        copyDataValidation(newSheet, gameAvailabilitySheet,
          [{ sourceColumn: ac.activationHeader, targetColumn: gix.activation }], fullNameInfo.rowCount);
      }
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

    const firstSummaryEndRow = addCoachGameRosterActivationSummary(newSheet, fullNameInfo.rowCount, idx);
    if (firstSummaryEndRow !== undefined) {
      addCoachGameRosterExcludingCantMakeItSummary(newSheet, fullNameInfo.rowCount, idx, firstSummaryEndRow);
    }

    // Auto-resize columns
    console.log('📏 Auto-resizing columns...');
    newSheet.autoResizeColumn(idx.number);
    for (var ri = 0; ri < idx.games.length; ri++) {
      var gx = idx.games[ri];
      if (gx.availability) newSheet.autoResizeColumn(gx.availability);
      if (gx.activation) newSheet.autoResizeColumn(gx.activation);
    }

    for (var ni = 0; ni < idx.games.length; ni++) {
      var nx = idx.games[ni];
      if (nx.note) {
        newSheet.getRange(2, nx.note, fullNameInfo.rowCount, 1).setWrap(true);
      }
    }
    if (idx.games.some(function (g) { return g.note; })) {
      console.log('📝 Enabled text wrap for note column(s)');
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
      `Successfully created game roster prep sheet for ${labelForUi} with ${fullNameInfo.rowCount} students.\n\nSorted by ${sortDesc}.`,
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
 * @param {Object[]} availColumnsList - One findAvailabilityColumns result per game that day
 * @param {Object} indices - 1-based column indices (number, fullName, team?, gender, grade, games[], activationStatus/availability/note = first game)
 * @param {number} numRows - Number of data rows
 */
function populateGameRosterPrepData(newSheet, rosterSheet, rosterHeaderRow, gameAvailabilitySheet, availColumnsList, indices, numRows) {
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

  for (var gi = 0; gi < availColumnsList.length; gi++) {
    var availColumns = availColumnsList[gi];
    var gix = indices.games[gi];
    if (!gix) continue;

    if (gix.activation && availColumns.activationStatusColumn) {
      var fAct = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.activationStatusColumn}:${availColumns.activationStatusColumn}),"")`;
      setFormula(gix.activation, fAct);
      console.log(`✅ Populated Activation Status column (game ${gi + 1}) with XLOOKUP`);
    }

    if (gix.availability && availColumns.availabilityColumn) {
      var fAvail = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.availabilityColumn}:${availColumns.availabilityColumn}),"")`;
      setFormula(gix.availability, fAvail);
      console.log(`✅ Populated Game Availability column (game ${gi + 1}) with XLOOKUP`);
    }

    if (gix.note && availColumns.noteColumn) {
      var fNote = `=IFERROR(XLOOKUP(B2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${availColumns.noteColumn}:${availColumns.noteColumn}),"")`;
      setFormula(gix.note, fNote);
      console.log(`✅ Populated Game Note column (game ${gi + 1}) with XLOOKUP`);
    }
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
 * Build a parent-friendly game roster with names and per-game availability columns
 * @param {Sheet} newSheet - The new sheet to populate
 * @param {Sheet} rosterSheet - The roster sheet
 * @param {Sheet} gameAvailabilitySheet - The game availability sheet
 * @param {string} gameDate - Game date in format "M/D"
 * @param {Object[]} availColumnsList - One findAvailabilityColumns result per game that day
 * @param {string} displayLabel - Short label for dialogs
 */
function buildParentGameRoster(newSheet, rosterSheet, gameAvailabilitySheet, gameDate, availColumnsList, displayLabel) {
  const labelForUi = displayLabel || gameDate;
  console.log(`👨‍👩‍👧‍👦 Building PARENT game roster for: ${labelForUi}`);

  try {
    const hasTeam = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasTeam;
    const hasActivation = CONFIG.gameRosterPrep && CONFIG.gameRosterPrep.hasActivationStatus;

    const headers = ['Full Name'];
    const col = { fullName: 1, games: [], team: null };
    var c = 2;

    for (var gi = 0; gi < availColumnsList.length; gi++) {
      var ac = availColumnsList[gi];
      var gc = { activation: null, availability: null, note: null };
      if (hasActivation && ac.activationHeader && ac.activationStatusColumn) {
        headers.push(ac.activationHeader);
        gc.activation = c++;
      }
      headers.push(ac.availabilityHeader);
      gc.availability = c++;
      headers.push(ac.noteHeader);
      gc.note = c++;
      col.games.push(gc);
    }

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

      for (var pi = 0; pi < availColumnsList.length; pi++) {
        var pac = availColumnsList[pi];
        var pc = col.games[pi];
        if (!pc) continue;
        if (pc.activation && pac.activationStatusColumn) {
          setFormula(pc.activation, `=IFERROR(XLOOKUP(A2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${pac.activationStatusColumn}:${pac.activationStatusColumn}),"")`);
          console.log(`✅ Populated Activation Status (game ${pi + 1}) with XLOOKUP`);
        }
        if (pc.availability && pac.availabilityColumn) {
          setFormula(pc.availability, `=IFERROR(XLOOKUP(A2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${pac.availabilityColumn}:${pac.availabilityColumn}),"")`);
          console.log(`✅ Populated Game Availability (game ${pi + 1}) with XLOOKUP`);
        }
        if (pc.note && pac.noteColumn) {
          setFormula(pc.note, `=IFERROR(XLOOKUP(A2,'${gameAvailSheetName}'!A:A,'${gameAvailSheetName}'!${pac.noteColumn}:${pac.noteColumn}),"")`);
          console.log(`✅ Populated Game Note (game ${pi + 1}) with XLOOKUP`);
        }
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
      var sortSpec = [];
      var g0 = col.games[0];
      if (g0 && g0.activation) sortSpec.push({ column: g0.activation, ascending: true });
      if (g0 && g0.availability) sortSpec.push({ column: g0.availability, ascending: true });
      sortSpec.push({ column: 1, ascending: true });
      console.log(sortSpec.length > 1
        ? `🔄 Sorting ${fullNameInfo.rowCount} rows by first game status / availability > Player Name...`
        : `🔄 Sorting ${fullNameInfo.rowCount} rows by player name...`);
      const dataRange = newSheet.getRange(2, 1, fullNameInfo.rowCount, numCols);
      dataRange.sort(sortSpec);
      console.log('✅ Sorting complete');
    }

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

    styleHeaderRow(newSheet, headers.length);
    refreshManagedAvailabilityAndActivationCfOnSheet(newSheet);

    if (fullNameInfo.rowCount > 0) {
      for (var di = 0; di < availColumnsList.length; di++) {
        var dac = availColumnsList[di];
        var dc = col.games[di];
        if (!dc) continue;
        if (dc.activation) {
          copyDataValidation(newSheet, gameAvailabilitySheet,
            [{ sourceColumn: dac.activationHeader, targetColumn: dc.activation }], fullNameInfo.rowCount);
        }
        if (dc.availability && dac.availabilityColumn) {
          applyGameAvailabilityValidation(newSheet, gameAvailabilitySheet, dac, dc.availability, fullNameInfo.rowCount);
        }
      }
    }

    finalizeGameRosterSheet(newSheet, fullNameInfo.rowCount);

    console.log('📏 Auto-resizing columns...');
    newSheet.autoResizeColumn(1);
    for (var ai = 0; ai < col.games.length; ai++) {
      var ag = col.games[ai];
      if (ag.activation) newSheet.autoResizeColumn(ag.activation);
      if (ag.availability) newSheet.autoResizeColumn(ag.availability);
      if (ag.note) {
        newSheet.getRange(2, ag.note, fullNameInfo.rowCount, 1).setWrap(true);
      }
    }
    if (col.team) newSheet.autoResizeColumn(col.team);

    console.log(`✅ Parent game roster created successfully`);

    var firstAct = col.games[0] && col.games[0].activation;
    const alertDetail = firstAct
      ? (hasTeam ? 'Sorted using the first game’s activation and availability, then player name. Practice Squad and Dropped players are hidden (filter applied).' : 'Sorted using the first game’s activation and availability, then player name.')
      : (hasTeam ? 'Sorted by player name. Practice Squad and Dropped players are hidden (filter applied).' : 'Sorted by player name.');
    SpreadsheetApp.getUi().alert(
      'Parent Game Roster Created!',
      `Successfully created parent game roster for ${labelForUi} with ${fullNameInfo.rowCount} students.\n\n${alertDetail}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return 'Success';

  } catch (error) {
    console.error('Error creating parent game roster:', error);
    throw new Error(`Failed to create parent game roster: ${error.message}`);
  }
}

