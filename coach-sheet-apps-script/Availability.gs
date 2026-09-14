/**
 * Availability Builder Module
 * Creates availability tracking columns based on dates from info sheets
 * Supports both Practice and Game availability tracking
 *
 * AVAILABILITY_VALIDATION_OPTIONS and GAME_ACTIVATION_STATUS_OPTIONS are defined in ManagedConditionalFormatting.gs
 * (shared with roster prep sheets and whole-sheet CF).
 */

// Configuration for Practice Availability feature
const PRACTICE_AVAILABILITY_CONFIG = {
  type: 'practice',
  emoji: '🏃',
  infoSheet: '📍Practice Info',
  availabilitySheet: 'Practice Availability',
  validationOptions: AVAILABILITY_VALIDATION_OPTIONS,
  skipConfig: {
    columnName: 'note',
    skipCondition: 'startsWith',
    skipValue: 'Bye',
    // Also skip cancelled practices: no column in Practice Availability, and sync will delete from calendar.
    // Matching is case-insensitive (e.g. Bye, BYE, Cancelled, CANCELLED all work).
    skipValues: ['Bye', 'Cancelled']
  }
};

// Configuration for Game Availability feature
const GAME_AVAILABILITY_CONFIG = {
  type: 'game',
  emoji: '🎮',
  infoSheet: '📍Game Info',
  availabilitySheet: 'Game Availability',
  validationOptions: AVAILABILITY_VALIDATION_OPTIONS,
  skipConfig: {
    columnName: 'game #',
    skipCondition: 'equals',
    skipValue: 'Bye'
  },
  // Game-only: three columns per date — $date Availability, $date Activation Status, $date Note (free text)
  columnsPerDate: [
    { suffix: ' Availability', useAvailabilityValidation: true },
    { suffix: ' Activation Status', useAvailabilityValidation: false, useActivationStatusValidation: true },
    { suffix: ' Note', useAvailabilityValidation: false, isFreeText: true }
  ]
};

/**
 * Parse a Game/Practice Info date cell into a Date (local). Handles "5/9 Sat", sheet serials, ISO.
 * @param {*} dateValue
 * @return {Date|null}
 */
function parseDateFromInfoSheet(dateValue) {
  if (dateValue == null || dateValue === '') return null;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) return dateValue;
  var s = String(dateValue).trim();
  if (/\s/.test(s)) {
    var first = s.split(/\s+/)[0];
    var d1 = parseDateFromInfoSheet(first);
    if (d1) return d1;
  }
  var serial = Number(s);
  if (!isNaN(serial) && serial > 0) {
    var epoch = new Date((serial - 25569) * 86400 * 1000);
    if (!isNaN(epoch.getTime())) return epoch;
  }
  var parts = s.split('/');
  if (parts.length >= 2) {
    var month = parseInt(parts[0], 10);
    var day = parseInt(parts[1], 10);
    var year = parts[2] ? parseInt(parts[2], 10) : new Date().getFullYear();
    if (!isNaN(month) && !isNaN(day) && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(year, month - 1, day);
    }
  }
  var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
  }
  var fallback = new Date(s);
  return isNaN(fallback.getTime()) ? null : fallback;
}

/**
 * Canonical "M/D" key for counting games on the same calendar day (matches player portal toCanonicalDateKey).
 * @param {Date} dateObj
 * @return {string}
 */
function toCanonicalDateKeyFromDateObj(dateObj) {
  return formatDateForColumn(dateObj);
}

/**
 * Get the expected column headers for a date in an availability sheet.
 * Single source of truth for column naming so Build Game/Practice Roster and Build Game Roster Prep use the same names.
 * Game Availability: 2nd+ game on the same day uses " (Game N)" suffix (matches madison-ultimate findGameColumns).
 * @param {string} dateString - Date in format "M/D" (e.g. "3/7")
 * @param {string} sheetType - 'Practice Availability' or 'Game Availability'
 * @param {number} [ordinalForDate] - 1 = first game that day (default); 2+ = double-header columns
 * @return {{ availabilityHeader: string, noteHeader: string, activationHeader?: string }}
 */
function getAvailabilityColumnHeaders(dateString, sheetType, ordinalForDate) {
  ordinalForDate = ordinalForDate || 1;
  if (sheetType === 'Game Availability') {
    const gamePart = ordinalForDate <= 1 ? '' : (' (Game ' + ordinalForDate + ')');
    return {
      availabilityHeader: dateString + ' Availability' + gamePart,
      noteHeader: dateString + ' Note' + gamePart,
      activationHeader: dateString + ' Activation Status' + gamePart
    };
  }
  // Practice Availability: first column is just the date, second is " Note"
  return {
    availabilityHeader: dateString,
    noteHeader: dateString + ' Note'
  };
}

/**
 * Shared function to build availability columns for practice or game
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 */
function buildAvailability(config) {
  const typeCapitalized = config.type.charAt(0).toUpperCase() + config.type.slice(1);
  console.log(`${config.emoji} Starting Build ${typeCapitalized} Availability...`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get dates from the appropriate info sheet
    const dates = getDatesFromInfoSheet(ss, config);
    
    if (dates.length === 0) {
      SpreadsheetApp.getUi().alert(
        `No ${typeCapitalized} Dates Found`,
        `No ${config.type} dates found in "${config.infoSheet}" sheet. Please ensure the sheet exists and contains ${config.type} date information.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Build availability columns in the appropriate availability sheet
    const result = buildAvailabilityColumns(ss, dates, config);
    
    console.log(`✅ ${typeCapitalized} Availability build complete`);
    
    const rowLabel = config.type === 'game' ? 'game row' : `${config.type} date`;
    let message = `Successfully processed ${config.availabilitySheet} sheet for ${dates.length} ${rowLabel}(s).\n\n`;
    
    if (result.columnsCreated > 0) {
      message += `📊 ${result.columnSummary}`;
    }
    
    if (result.columnsSkipped > 0) {
      if (result.columnsCreated > 0) message += '\n\n';
      message += `⏭️ ${result.skippedSummary}`;
    }
    
    if (result.columnsCreated === 0 && result.columnsSkipped === 0) {
      message += 'No changes needed - all columns already exist.';
    }

    message += `\n\n👥 Player rows: ${result.rowsAdded} added from the Roster (Include In Generated Rosters TRUE), ${result.playerIdsFilled} existing row(s) given a PlayerID. Rows are never deleted; remove a cut player's row by hand if it was seeded before their Include flag was set FALSE.`;
    if (result.unmatchedRows.length > 0) {
      message += `\n\n⚠️ ${result.unmatchedRows.length} existing row(s) have no PlayerID because no single Roster player has that Full Name: ${result.unmatchedRows.join(', ')}. Fix the name or type the PlayerID.`;
    }
    
    message += '\n\n🎯 Data validation is applied in bulk (one rule type for all availability columns, one for all activation columns on game sheets). Conditional formatting uses one rule per status value across all matching columns.';
    
    SpreadsheetApp.getUi().alert(`${typeCapitalized} Availability Updated!`, message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error(`Error building ${typeCapitalized} Availability:`, error);
    SpreadsheetApp.getUi().alert('Error', `Failed to build ${typeCapitalized} Availability: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Main function to build practice availability columns
 * Called from the menu
 */
function buildPracticeAvailability() {
  buildAvailability(PRACTICE_AVAILABILITY_CONFIG);
}

/**
 * Main function to build game availability columns
 * Called from the menu
 */
function buildGameAvailability() {
  buildAvailability(GAME_AVAILABILITY_CONFIG);
}

/**
 * Shared function to extract dates from an info sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 * @return {Array} Array of date objects: {date, formattedDate, rowIndex}; for games also {ordinalForDate, gameLabel, team}
 */
function getDatesFromInfoSheet(ss, config) {
  const infoSheet = ss.getSheetByName(config.infoSheet);
  
  if (!infoSheet) {
    throw new Error(`${config.type.charAt(0).toUpperCase() + config.type.slice(1)} Info sheet "${config.infoSheet}" not found`);
  }
  
  console.log(`📅 Reading ${config.type} dates from "${config.infoSheet}"`);
  
  // Look for a Date column in the info sheet
  const headerRow = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
  const dateColumnIndex = headerRow.findIndex(header => 
    header && header.toString().toLowerCase().includes('date')
  );
  
  if (dateColumnIndex === -1) {
    throw new Error(`No date column found in "${config.infoSheet}" sheet. Please ensure there is a column with "date" in the header.`);
  }
  
  console.log(`📍 Found date column at index ${dateColumnIndex + 1}`);
  
  // Look for skip column if configured
  let skipColumnIndex = -1;
  if (config.skipConfig) {
    skipColumnIndex = headerRow.findIndex(header => 
      header && header.toString().toLowerCase().includes(config.skipConfig.columnName.toLowerCase())
    );
    
    if (skipColumnIndex !== -1) {
      console.log(`📍 Found skip column "${config.skipConfig.columnName}" at index ${skipColumnIndex + 1}`);
    } else {
      console.log(`⚠️ Skip column "${config.skipConfig.columnName}" not found - will not skip any rows`);
    }
  }
  
  // Games only: the Team column groups same-day games so each team gets one column triple per date.
  const teamColumnIndex = config.type === 'game' ? findTeamColumnIndex_(headerRow) : -1;
  if (config.type === 'game') {
    if (teamColumnIndex !== -1) {
      console.log(`📍 Found Team column at index ${teamColumnIndex + 1}`);
    } else {
      console.log('⚠️ No Team column in Game Info; every same-day row counts as another game for the same team');
    }
  }

  // Get all dates from the date column (skip header row)
  const lastRow = infoSheet.getLastRow();
  if (lastRow <= 1) {
    console.log(`⚠️ No ${config.type} data found in ${config.infoSheet} sheet`);
    return [];
  }
  
  // Get all data we need (date column and skip column if applicable)
  const columnsNeeded = Math.max(dateColumnIndex + 1, skipColumnIndex + 1, teamColumnIndex + 1);
  const allData = infoSheet.getRange(2, 1, lastRow, columnsNeeded).getValues();
  const dates = [];
  const isGame = config.type === 'game';
  const ordinalByCanonical = isGame ? {} : null;
  
  allData.forEach((row, index) => {
    const dateValue = row[dateColumnIndex];
    
    // Check if we should skip this row based on skip configuration
    if (config.skipConfig && skipColumnIndex !== -1) {
      const skipValue = row[skipColumnIndex];
      let shouldSkip = false;

      if (skipValue && skipValue.toString().trim() !== '') {
        const skipValueStr = skipValue.toString().trim().toLowerCase();
        const valuesToCheck = config.skipConfig.skipValues && config.skipConfig.skipValues.length
          ? config.skipConfig.skipValues
          : [config.skipConfig.skipValue];

        if (config.skipConfig.skipCondition === 'startsWith') {
          shouldSkip = valuesToCheck.some(function (v) {
            return v && skipValueStr.startsWith(String(v).toLowerCase());
          });
        } else if (config.skipConfig.skipCondition === 'equals') {
          shouldSkip = valuesToCheck.some(function (v) {
            return v && skipValueStr === String(v).toLowerCase();
          });
        }
      }

      if (shouldSkip) {
        console.log(`⏭️ Skipping row ${index + 2}: ${config.skipConfig.columnName} = "${skipValue}"`);
        return; // Skip this iteration
      }
    }
    
    if (dateValue && dateValue !== '') {
      try {
        const dateObj = parseDateFromInfoSheet(dateValue);
        if (dateObj && !isNaN(dateObj.getTime())) {
          const formattedDate = formatDateForColumn(dateObj);
          const entry = {
            date: dateObj,
            formattedDate: formattedDate,
            rowIndex: index + 2 // +2 for 1-based indexing and header row
          };
          if (isGame) {
            const canonical = toCanonicalDateKeyFromDateObj(dateObj);
            const team = teamColumnIndex !== -1 && row[teamColumnIndex] != null ? String(row[teamColumnIndex]).trim() : '';
            const ordinalForDate = nextGameOrdinal(ordinalByCanonical, canonical, team);
            entry.ordinalForDate = ordinalForDate;
            entry.team = team;
            if (skipColumnIndex !== -1 && row[skipColumnIndex] != null && String(row[skipColumnIndex]).trim() !== '') {
              entry.gameLabel = String(row[skipColumnIndex]).trim();
            }
            console.log(`${config.emoji} Found game row: ${formattedDate} (Game ${ordinalForDate} on ${canonical}${team ? ' for ' + team : ''}, row ${index + 2})`);
          } else {
            console.log(`${config.emoji} Found ${config.type} date: ${formattedDate} (row ${index + 2})`);
          }
          dates.push(entry);
        } else {
          console.warn(`⚠️ Invalid date in row ${index + 2}: "${dateValue}"`);
        }
      } catch (error) {
        console.warn(`⚠️ Error parsing date in row ${index + 2}: "${dateValue}" - ${error.message}`);
      }
    }
  });
  
  console.log(`🎯 Found ${dates.length} valid ${config.type} ${isGame ? 'rows' : 'dates'}`);
  return dates;
}

/**
 * Index of the Game Info "Team" header (exact, case-insensitive), or -1 when the season has none.
 * @param {Array} headerRow
 * @return {number}
 */
function findTeamColumnIndex_(headerRow) {
  return headerRow.findIndex(function (header) {
    return header != null && String(header).trim().toLowerCase() === 'team';
  });
}

/**
 * Group key for numbering same-day games: one group per calendar date per Team, so three teams
 * playing on 9/26 each get game 1 on that date and share the "9/26 Availability" column triple.
 * A blank Team is its own group (an all-team event). Pure; shared by Build Game Availability and
 * Build Game Roster Prep through getDatesFromInfoSheet, and mirrored by the player portal.
 * @param {string} canonicalDate "M/D"
 * @param {string} team Team cell, may be blank
 * @return {string}
 */
function gameOrdinalGroupKey(canonicalDate, team) {
  return canonicalDate + '|' + String(team == null ? '' : team).trim().toLowerCase();
}

/**
 * Assign the next ordinal for a game row within its date-plus-team group, mutating the counter map.
 * @param {Object} counters group key to count so far
 * @param {string} canonicalDate "M/D"
 * @param {string} team Team cell, may be blank
 * @return {number} 1 for the first game of that team on that date, 2 for a real double-header, and so on
 */
function nextGameOrdinal(counters, canonicalDate, team) {
  const key = gameOrdinalGroupKey(canonicalDate, team);
  counters[key] = (counters[key] || 0) + 1;
  return counters[key];
}

/**
 * Format a date for use in column headers (e.g., "9/26")
 * @param {Date} date - The date to format
 * @return {string} Formatted date string
 */
function formatDateForColumn(date) {
  const month = date.getMonth() + 1; // getMonth() returns 0-based month
  const day = date.getDate();
  return `${month}/${day}`;
}

/**
 * Build availability columns in the specified availability sheet (shared function for practice/game)
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {Array} dates - Array of date objects (practice or game dates)
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 * @return {Object} Result object with statistics
 */
function buildAvailabilityColumns(ss, dates, config) {
  let availabilitySheet = ss.getSheetByName(config.availabilitySheet);
  
  // Create the sheet if it doesn't exist
  if (!availabilitySheet) {
    console.log(`📋 Creating new "${config.availabilitySheet}" sheet`);
    availabilitySheet = ss.insertSheet(config.availabilitySheet);
    
    // Column A stays Full Name (the prep sheets XLOOKUP against A:A); the portal matches rows by PlayerID.
    const baseHeaders = [AVAILABILITY_ROW_HEADERS.fullName, AVAILABILITY_ROW_HEADERS.playerId, AVAILABILITY_ROW_HEADERS.grade, AVAILABILITY_ROW_HEADERS.genderIdentification];
    availabilitySheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
    availabilitySheet.getRange(1, 1, 1, baseHeaders.length).setFontWeight('bold');
  }
  
  console.log(`📊 Building availability columns in "${config.availabilitySheet}"`);
  
  // Get existing columns to check what already exists
  const existingColumns = getExistingColumns(availabilitySheet);
  console.log(`📋 Found ${Object.keys(existingColumns).length} existing columns`);
  
  const columnsCreated = [];
  const columnsSkipped = [];
  let validationRanges = []; // Track ranges that need availability dropdown
  let noteColumns = []; // Game only: Note columns (free text — clear any validation)

  // Find where to start adding new columns (after existing columns)
  let nextColumnIndex = Math.max(2, availabilitySheet.getLastColumn() + 1);

  const isGame = config.type === 'game' && config.columnsPerDate;

  // Process each date (practice or game)
  dates.forEach((dateInfo, index) => {
    const dateString = dateInfo.formattedDate;

    if (isGame) {
      const ord = dateInfo.ordinalForDate || 1;
      const hdr = getAvailabilityColumnHeaders(dateString, 'Game Availability', ord);
      const triple = [
        { header: hdr.availabilityHeader, colDef: config.columnsPerDate[0] },
        { header: hdr.activationHeader, colDef: config.columnsPerDate[1] },
        { header: hdr.noteHeader, colDef: config.columnsPerDate[2] }
      ];
      triple.forEach(function (item) {
        const header = item.header;
        const colDef = item.colDef;
        const existingCol = existingColumns[header];

        if (existingCol) {
          console.log(`⏭️ Column "${header}" already exists at column ${existingCol}`);
          columnsSkipped.push(header);
          if (colDef.useAvailabilityValidation) {
            const existingValidation = availabilitySheet.getRange(2, existingCol, 1, 1).getDataValidation();
            if (!existingValidation) {
              validationRanges.push({ column: existingCol, header: header, isExisting: true });
            }
          } else if (colDef.isFreeText) {
            noteColumns.push(existingCol);
          }
          return;
        }

        console.log(`➕ Creating new column "${header}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(header);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        columnsCreated.push(header);

        if (colDef.useAvailabilityValidation) {
          validationRanges.push({ column: nextColumnIndex, header: header, isExisting: false });
        } else if (colDef.isFreeText) {
          noteColumns.push(nextColumnIndex);
        }
        nextColumnIndex++;
      });
    } else {
      // Practice: two columns per date — dateString (availability), "$Date Note"
      const availabilityHeader = dateString;
      const notesHeader = dateString + ' Note';

      if (existingColumns[availabilityHeader]) {
        console.log(`⏭️ Column "${availabilityHeader}" already exists at column ${existingColumns[availabilityHeader]}`);
        columnsSkipped.push(availabilityHeader);
        const existingColumnIndex = existingColumns[availabilityHeader];
        const existingValidation = availabilitySheet.getRange(2, existingColumnIndex, 1, 1).getDataValidation();
        if (!existingValidation) {
          validationRanges.push({ column: existingColumnIndex, header: availabilityHeader, isExisting: true });
        }
      } else {
        console.log(`➕ Creating new column "${availabilityHeader}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(availabilityHeader);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        validationRanges.push({ column: nextColumnIndex, header: availabilityHeader, isExisting: false });
        columnsCreated.push(availabilityHeader);
        nextColumnIndex++;
      }

      if (existingColumns[notesHeader]) {
        console.log(`⏭️ Column "${notesHeader}" already exists at column ${existingColumns[notesHeader]}`);
        columnsSkipped.push(notesHeader);
      } else {
        console.log(`➕ Creating new column "${notesHeader}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(notesHeader);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        columnsCreated.push(notesHeader);
        nextColumnIndex++;
      }
    }
  });

  // Seed one row per Roster player (Include In Generated Rosters TRUE) not already present, keyed by PlayerID.
  const seedResult = seedAvailabilityRows_(ss, availabilitySheet);

  // Apply or extend data validation to availability columns (per-column; consolidated below for fewer rules)
  extendOrCreateDataValidation(availabilitySheet, validationRanges, config);

  // Game only: ensure Note columns are free text (no dropdown)
  noteColumns.forEach(function (col) {
    // getRange(row, column, numRows, numColumns) — 1 column only
    availabilitySheet.getRange(2, col, 999, 1).clearDataValidations();
  });

  if (config.type === 'game') {
    consolidateGameAvailabilityDataValidations(availabilitySheet);
  } else if (config.type === 'practice') {
    consolidatePracticeAvailabilityDataValidations(availabilitySheet);
  }

  // Enable text wrapping on the sheet so headers and cells wrap
  const lastCol = availabilitySheet.getLastColumn();
  const lastRow = Math.max(availabilitySheet.getLastRow(), 2);
  if (lastCol >= 1 && lastRow >= 1) {
    availabilitySheet.getRange(1, 1, lastRow, lastCol).setWrap(true);
  }

  // Apply Format Spruce Up silently (no modal)
  console.log('✨ Applying Format Spruce Up formatting...');
  try {
    applySpruceUpFormatting(availabilitySheet);
  } catch (error) {
    console.warn('⚠️ Could not apply Format Spruce Up formatting:', error.message);
  }

  refreshManagedAvailabilityAndActivationCfOnSheet(availabilitySheet);

  return {
    rowsAdded: seedResult.rowsAdded,
    playerIdsFilled: seedResult.playerIdsFilled,
    unmatchedRows: seedResult.unmatchedRows,
    columnsCreated: columnsCreated.length,
    columnsSkipped: columnsSkipped.length,
    columnSummary: columnsCreated.length > 0 ? 
      `Created: ${columnsCreated.join(', ')}` : 'No new columns created',
    skippedSummary: columnsSkipped.length > 0 ? 
      `Skipped existing: ${columnsSkipped.join(', ')}` : ''
  };
}

// Per-player (non-date) headers of an availability sheet. Full Name is column A by convention
// (the roster prep sheets XLOOKUP against A:A); everything else is found by header name.
const AVAILABILITY_ROW_HEADERS = {
  fullName: 'Full Name',
  playerId: 'PlayerID',
  grade: 'Grade',
  genderIdentification: 'Gender Identification'
};

/**
 * Seed and repair the per-player rows of an availability sheet from the Roster.
 *
 * Rule: every Roster player whose Include In Generated Rosters is TRUE and whose PlayerID is not
 * already in the sheet's PlayerID column gets one appended row (Full Name, PlayerID, Grade,
 * Gender Identification as plain values). Rows are never deleted or reordered, the same rule
 * Sync Extra Player Info follows, so flag cut players FALSE before the first build. An existing
 * row with a Full Name but a blank PlayerID is filled when exactly one Roster player has that
 * Full Name; otherwise it is left blank and reported. The PlayerID header is appended after the
 * last column when missing, and always resolved by name afterwards.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Practice Availability or Game Availability
 * @return {{rowsAdded: number, playerIdsFilled: number, unmatchedRows: string[]}}
 */
function seedAvailabilityRows_(ss, sheet) {
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  if (!rosterSheet) {
    throw new Error(`Roster sheet "${CONFIG.roster.sheetName}" not found; run "Generate Fresh Roster" first.`);
  }
  const table = readRosterTable(rosterSheet);
  const idCol = table.col(CONFIG.columns.playerId);
  const nameCol = table.col(CONFIG.columns.fullName);
  const includeCol = table.col(CONFIG.columns.includeInGeneratedRosters);
  const gradeCol = table.col(CONFIG.columns.grade);
  const genderCol = table.col(CONFIG.columns.genderIdentification);
  const cell = (row, index) => (row[index] === null || row[index] === undefined) ? '' : row[index].toString().trim();

  const rosterPlayers = table.rows
    .filter(row => cell(row, idCol) !== '')
    .map(row => ({
      playerId: cell(row, idCol),
      fullName: cell(row, nameCol),
      include: row[includeCol] === true || cell(row, includeCol).toUpperCase() === 'TRUE',
      grade: row[gradeCol] === null || row[gradeCol] === undefined ? '' : row[gradeCol],
      genderIdentification: cell(row, genderCol)
    }));
  const rosterIdsByFullName = {};
  rosterPlayers.forEach(p => {
    if (!p.fullName) return;
    if (!rosterIdsByFullName[p.fullName]) rosterIdsByFullName[p.fullName] = [];
    rosterIdsByFullName[p.fullName].push(p.playerId);
  });

  // Ensure the PlayerID header exists; append it after the last column when missing.
  let columns = getExistingColumns(sheet);
  if (!columns[AVAILABILITY_ROW_HEADERS.playerId]) {
    const newCol = sheet.getLastColumn() + 1;
    console.log(`➕ Adding "${AVAILABILITY_ROW_HEADERS.playerId}" header at column ${newCol}`);
    sheet.getRange(1, newCol).setValue(AVAILABILITY_ROW_HEADERS.playerId).setFontWeight('bold');
    columns = getExistingColumns(sheet);
  }
  const playerIdCol = columns[AVAILABILITY_ROW_HEADERS.playerId];
  const fullNameCol = columns[AVAILABILITY_ROW_HEADERS.fullName] || 1;
  const gradeAvailCol = columns[AVAILABILITY_ROW_HEADERS.grade];
  const genderAvailCol = columns[AVAILABILITY_ROW_HEADERS.genderIdentification];
  const lastCol = sheet.getLastColumn();

  // Existing rows: collect PlayerIDs, fill blanks by unambiguous Full Name.
  const existingIds = new Set();
  const unmatchedRows = [];
  let playerIdsFilled = 0;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, fullNameCol, lastRow - 1, 1).getValues();
    const ids = sheet.getRange(2, playerIdCol, lastRow - 1, 1).getValues();
    const fills = [];
    for (let i = 0; i < ids.length; i++) {
      const id = cell(ids[i], 0);
      const name = cell(names[i], 0);
      if (id) {
        existingIds.add(id);
        continue;
      }
      if (!name) continue;
      const candidates = rosterIdsByFullName[name] || [];
      if (candidates.length === 1 && !existingIds.has(candidates[0])) {
        fills.push({ row: i + 2, playerId: candidates[0] });
        existingIds.add(candidates[0]);
      } else {
        unmatchedRows.push(name);
      }
    }
    fills.forEach(f => sheet.getRange(f.row, playerIdCol).setValue(f.playerId));
    playerIdsFilled = fills.length;
  }

  // Append one row per included Roster player not yet present.
  const missing = rosterPlayers.filter(p => p.include && !existingIds.has(p.playerId));
  if (missing.length > 0) {
    const startRow = Math.max(lastRow, 1) + 1;
    const rows = missing.map(p => {
      const row = new Array(lastCol).fill('');
      row[fullNameCol - 1] = p.fullName;
      row[playerIdCol - 1] = p.playerId;
      if (gradeAvailCol) row[gradeAvailCol - 1] = p.grade;
      if (genderAvailCol) row[genderAvailCol - 1] = p.genderIdentification;
      return row;
    });
    if (sheet.getMaxRows() < startRow + rows.length - 1) {
      sheet.insertRowsAfter(sheet.getMaxRows(), startRow + rows.length - 1 - sheet.getMaxRows());
    }
    sheet.getRange(startRow, 1, rows.length, lastCol).setValues(rows);
  }

  console.log(`👥 Availability rows: added ${missing.length}, PlayerIDs filled ${playerIdsFilled}, unmatched ${unmatchedRows.length}`);
  return { rowsAdded: missing.length, playerIdsFilled: playerIdsFilled, unmatchedRows: unmatchedRows };
}

/**
 * Get a map of existing column headers to their column indices
 * @param {Sheet} sheet - The Practice Availability sheet
 * @return {Object} Map of header names to column indices
 */
function getExistingColumns(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return {};
  
  const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const columnMap = {};
  
  console.log('🔍 Existing columns found:');
  headerRow.forEach((header, index) => {
    if (header && header.toString().trim() !== '') {
      const cleanHeader = header.toString().trim();
      columnMap[cleanHeader] = index + 1; // 1-based column index
      console.log(`  Column ${index + 1}: "${cleanHeader}"`);
      
      // Also map date objects to formatted date strings for comparison
      if (header instanceof Date) {
        const formattedDate = formatDateForColumn(header);
        columnMap[formattedDate] = index + 1;
        console.log(`  Column ${index + 1} also mapped as: "${formattedDate}"`);
      }
    }
  });
  
  return columnMap;
}

/**
 * Apply or extend data validation to availability columns
 * @param {Sheet} sheet - The availability sheet (practice or game)
 * @param {Array} validationRanges - Array of column info for validation
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 */
function extendOrCreateDataValidation(sheet, validationRanges, config) {
  console.log('🎯 Applying or extending data validation to availability columns...');
  console.log(`📊 Processing ${validationRanges.length} validation ranges`);
  
  // Create validation options from config
  const validationValues = config.validationOptions.map(option => option.value);
  console.log(`🎯 Expected validation values: [${validationValues.join(', ')}]`);
  
  // Check if there's an existing data validation rule we can extend
  const existingValidation = findExistingDataValidation(sheet, validationValues);
  
  if (existingValidation) {
    console.log(`✅ Found existing compatible validation rule in column ${existingValidation.column}`);
  } else {
    console.log(`❌ No existing compatible validation rule found`);
  }
  
  validationRanges.forEach((rangeInfo, index) => {
    console.log(`\n📋 Processing range ${index + 1}/${validationRanges.length}:`);
    console.log(`   Column: ${rangeInfo.column}, Header: "${rangeInfo.header}", IsExisting: ${rangeInfo.isExisting}`);
    
    try {
      if (rangeInfo.isExisting && existingValidation) {
        // Extend existing validation rule
        console.log(`🔄 Extending existing data validation for column ${rangeInfo.column} (${rangeInfo.header})`);
        extendExistingValidation(sheet, rangeInfo.column, existingValidation);
      } else {
        // Create new validation rule
        const reason = !rangeInfo.isExisting ? 'new column' : 'no compatible existing validation';
        console.log(`➕ Creating new data validation for column ${rangeInfo.column} (${rangeInfo.header}) - ${reason}`);
        createNewValidation(sheet, rangeInfo.column, validationValues);
      }
      
    } catch (error) {
      console.warn(`⚠️ Could not apply data validation to column ${rangeInfo.column}: ${error.message}`);
    }
  });
}

/**
 * Find existing data validation rule that matches our requirements
 * @param {Sheet} sheet - The Practice Availability sheet
 * @param {Array} expectedValues - Expected validation values
 * @return {Object|null} Existing validation info or null
 */
function findExistingDataValidation(sheet, expectedValues) {
  console.log('🔍 Searching for existing data validation rules...');
  console.log(`   Checking columns 2 to ${Math.min(10, sheet.getLastColumn())}`);
  
  try {
    // Check a few columns for existing validation rules
    for (let col = 2; col <= Math.min(10, sheet.getLastColumn()); col++) {
      console.log(`   📋 Checking column ${col} for existing validation...`);
      
      const range = sheet.getRange(2, col, 1, 1);
      const validation = range.getDataValidation();
      
      if (validation) {
        console.log(`      ✅ Found validation rule in column ${col}`);
        const criteria = validation.getCriteriaType();
        console.log(`      📊 Criteria type: ${criteria}`);
        
        if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          console.log(`      📝 It's a VALUE_IN_LIST validation`);
          const values = validation.getCriteriaValues()[0];
          console.log(`      🎯 Existing values: [${values ? values.join(', ') : 'null'}]`);
          console.log(`      🎯 Expected values: [${expectedValues.join(', ')}]`);
          
          if (values) {
            const existingSorted = values.slice().sort();
            const expectedSorted = expectedValues.slice().sort();
            console.log(`      🔄 Sorted existing: [${existingSorted.join(', ')}]`);
            console.log(`      🔄 Sorted expected: [${expectedSorted.join(', ')}]`);
            
            if (arraysEqual(existingSorted, expectedSorted)) {
              console.log(`      ✅ Arrays match! Found compatible existing validation rule in column ${col}`);
              return {
                validation: validation,
                column: col
              };
            } else {
              console.log(`      ❌ Arrays don't match`);
            }
          } else {
            console.log(`      ❌ No values found in validation rule`);
          }
        } else {
          console.log(`      ❌ Not a VALUE_IN_LIST validation`);
        }
      } else {
        console.log(`      ❌ No validation rule found in column ${col}`);
      }
    }
  } catch (error) {
    console.warn('⚠️ Error checking for existing validation:', error.message);
  }
  
  console.log('❌ No compatible existing validation rule found');
  return null;
}

/**
 * Extend existing validation to a new column (copies criteria values from a compatible existing rule).
 * @param {Sheet} sheet - The Practice Availability sheet
 * @param {number} column - Column to apply validation to
 * @param {Object} existingValidation - Existing validation info
 */
function extendExistingValidation(sheet, column, existingValidation) {
  console.log(`🎨 Extending validation to column ${column}`);
  const validationRange = sheet.getRange(2, column, 1000, 1);
  
  // Create a new validation rule that matches the existing one
  const originalValidation = existingValidation.validation;
  const criteriaValues = originalValidation.getCriteriaValues()[0];
  
  const newValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(criteriaValues, true)
    .setAllowInvalid(originalValidation.getAllowInvalid())
    .setHelpText(originalValidation.getHelpText() || 'Select your availability')
    .build();
    
  validationRange.setDataValidation(newValidation);
  console.log(`✅ Validation extended to column ${column}`);
}

/**
 * Create new data validation rule
 * @param {Sheet} sheet - The availability sheet
 * @param {number} column - Column to apply validation to
 * @param {Array} validationValues - Values for validation
 */
function createNewValidation(sheet, column, validationValues) {
  console.log(`🎨 Creating new validation for column ${column}`);
  const validationRange = sheet.getRange(2, column, 1000, 1);
  
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(validationValues, true) // true = show dropdown
    .setAllowInvalid(false)
    .setHelpText('Select your availability')
    .build();
  
  validationRange.setDataValidation(validation);
  console.log(`✅ New validation created for column ${column}`);
}

/**
 * Apply the same data validation to multiple columns (row 2 downward, numRows tall).
 * RangeList supports clearDataValidations() but not setDataValidation(); loop per column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number[]} columns 1-based column indices
 * @param {number} numRows Height in rows (SpreadsheetApp getRange row/column/numRows/numColumns overload)
 * @param {GoogleAppsScript.Spreadsheet.DataValidation} validation
 */
function applyDataValidationToColumnRanges_(sheet, columns, numRows, validation) {
  columns.forEach(function (c) {
    sheet.getRange(2, c, numRows, 1).setDataValidation(validation);
  });
}

/**
 * Apply one VALUE_IN_LIST to all game availability columns and one to all activation columns (fewer DV entries than per-column).
 * @param {Sheet} sheet
 */
function consolidateGameAvailabilityDataValidations(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 100);
  var numRows = lastRow - 1;
  var idx = collectGameAvailabilityColumnIndices_(sheet);
  var valOpts = GAME_AVAILABILITY_CONFIG.validationOptions.map(function (o) {
    return o.value;
  });
  var actOpts = GAME_ACTIVATION_STATUS_OPTIONS.map(function (o) {
    return o.value;
  });

  if (idx.availabilityCols.length > 0) {
    var dvA = SpreadsheetApp.newDataValidation()
      .requireValueInList(valOpts, true)
      .setAllowInvalid(false)
      .setHelpText('Select your availability')
      .build();
    applyDataValidationToColumnRanges_(sheet, idx.availabilityCols, numRows, dvA);
  }
  if (idx.activationCols.length > 0) {
    var dvB = SpreadsheetApp.newDataValidation()
      .requireValueInList(actOpts, true)
      .setAllowInvalid(false)
      .setHelpText('Active / Inactive / TBD')
      .build();
    applyDataValidationToColumnRanges_(sheet, idx.activationCols, numRows, dvB);
  }
}

/**
 * One VALUE_IN_LIST applied to all practice date availability columns (same rule per column).
 * @param {Sheet} sheet
 */
function consolidatePracticeAvailabilityDataValidations(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 100);
  var numRows = lastRow - 1;
  var cols = collectPracticeAvailabilityColumnIndices_(sheet);
  if (cols.length === 0) return;

  var valOpts = PRACTICE_AVAILABILITY_CONFIG.validationOptions.map(function (o) {
    return o.value;
  });
  var dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(valOpts, true)
    .setAllowInvalid(false)
    .setHelpText('Select your availability')
    .build();
  applyDataValidationToColumnRanges_(sheet, cols, numRows, dv);
}

/**
 * Check if two arrays are equal (order-independent)
 * @param {Array} arr1 - First array
 * @param {Array} arr2 - Second array
 * @return {boolean} True if arrays contain same elements
 */
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }
  
  return true;
}

