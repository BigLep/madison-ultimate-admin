/**
 * Build Coach Availability: the Coaches tab (one row per Coach, keyed by CoachID) and the Coach
 * Availability tab (one row per Coach, one column pair per practice and per game).
 *
 * Coach Availability combines practices and games in one tab, and every Game Info row gets its own
 * pair regardless of Team (coaches are not assigned to teams). Headers, shared with the portal's
 * coach-availability reader (madison-ultimate src/lib/coach-availability.ts):
 *   practice:           "9/23 Practice"      and "9/23 Practice Note"
 *   game with a Team:   "9/26 Blue Game"     and "9/26 Blue Game Note"
 *   game, blank Team:   "10/17 Game"         and "10/17 Game Note"
 *   second game for the same Team that day: "9/26 Blue Game 2" and "9/26 Blue Game 2 Note"
 * Events come from getDatesFromInfoSheet, the same reader Build Practice/Game Availability use, so
 * cancelled practices are skipped and games are numbered per date per Team the same way.
 *
 * Like the player Availability Sheets (ADR 0004), CoachID in column A is the only typed per-coach
 * value; Name is a lookup formula on the Coaches tab. Rows and columns are never deleted.
 */

const COACH_COLUMNS = {
  coachId: 'CoachID',
  name: 'Name',
  email: 'Email',
  phone: 'Phone',
  about: 'About',
  photoDriveFileId: 'Photo Drive File ID'
};

const COACH_AVAILABILITY_ROW_HEADERS = {
  coachId: 'CoachID',
  name: 'Name'
};

// Same alphabet and length as the portal's PlayerID (src/lib/player-identity.ts): no 0/1/i/l/o.
const COACH_ID_ALPHABET = 'abcdefghjkmnpqrstuvwxyz23456789';
const COACH_ID_LENGTH = 5;

/**
 * Mint a short random opaque CoachID not already in use.
 * @param {Set<string>} taken
 * @param {function(): number} [random] - Math.random by default; injectable for tests
 * @return {string}
 */
function mintCoachId_(taken, random) {
  random = random || Math.random;
  for (;;) {
    let id = '';
    for (let i = 0; i < COACH_ID_LENGTH; i++) {
      id += COACH_ID_ALPHABET.charAt(Math.floor(random() * COACH_ID_ALPHABET.length));
    }
    if (!taken.has(id)) return id;
  }
}

/**
 * Column headers for one Coach Availability event. Pure; the portal mirrors it.
 * @param {{kind: string, formattedDate: string, team?: string, ordinalForDate?: number}} event
 * @return {{availabilityHeader: string, noteHeader: string}}
 */
function coachAvailabilityHeaders(event) {
  let base;
  if (event.kind === 'practice') {
    base = event.formattedDate + ' Practice';
  } else {
    const team = event.team ? String(event.team).trim() : '';
    const ordinal = event.ordinalForDate || 1;
    base = event.formattedDate + (team ? ' ' + team : '') + ' Game' + (ordinal > 1 ? ' ' + ordinal : '');
  }
  return { availabilityHeader: base, noteHeader: base + ' Note' };
}

/**
 * Every practice and game, in date order, practices before games on the same date, games in Game
 * Info row order.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<{kind: string, date: Date, formattedDate: string, team?: string, ordinalForDate?: number}>}
 */
function getCoachAvailabilityEvents_(ss) {
  const practices = getDatesFromInfoSheet(ss, PRACTICE_AVAILABILITY_CONFIG)
    .map(d => ({ kind: 'practice', date: d.date, formattedDate: d.formattedDate, rowIndex: d.rowIndex }));
  const games = getDatesFromInfoSheet(ss, GAME_AVAILABILITY_CONFIG)
    .map(d => ({ kind: 'game', date: d.date, formattedDate: d.formattedDate, team: d.team, ordinalForDate: d.ordinalForDate, rowIndex: d.rowIndex }));
  const kindRank = { practice: 0, game: 1 };
  return practices.concat(games).sort((a, b) =>
    (a.date.getTime() - b.date.getTime()) || (kindRank[a.kind] - kindRank[b.kind]) || (a.rowIndex - b.rowIndex));
}

/**
 * Get a sheet by name, creating it with the given header row when missing, and appending any
 * header that is missing from an existing sheet (never moving or removing columns).
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.log(`📋 Creating "${sheetName}" sheet`);
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sheet;
  }
  const existing = getExistingColumns(sheet);
  let next = sheet.getLastColumn() + 1;
  headers.forEach(h => {
    if (existing[h]) return;
    sheet.getRange(1, next).setValue(h).setFontWeight('bold');
    next++;
  });
  return sheet;
}

/**
 * Ensure the Coaches tab exists with its headers and give every named Coach a CoachID.
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, coachIds: string[], coachIdsMinted: number}}
 *   coachIds in sheet order (rows with neither a Name nor a CoachID are ignored)
 */
function ensureCoaches_(ss, random) {
  const sheet = ensureSheetWithHeaders_(ss, CONFIG.coaches.sheetName, Object.values(COACH_COLUMNS));
  const cols = getExistingColumns(sheet);
  const idCol = cols[COACH_COLUMNS.coachId];
  const nameCol = cols[COACH_COLUMNS.name];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { sheet: sheet, coachIds: [], coachIdsMinted: 0 };

  const numRows = lastRow - 1;
  const text = v => (v === null || v === undefined) ? '' : String(v).trim();
  const ids = sheet.getRange(2, idCol, numRows, 1).getValues().map(r => text(r[0]));
  const names = sheet.getRange(2, nameCol, numRows, 1).getValues().map(r => text(r[0]));
  const taken = new Set(ids.filter(Boolean));
  let minted = 0;
  for (let i = 0; i < numRows; i++) {
    if (ids[i] || !names[i]) continue;
    ids[i] = mintCoachId_(taken, random);
    taken.add(ids[i]);
    sheet.getRange(i + 2, idCol).setValue(ids[i]);
    minted++;
  }
  return { sheet: sheet, coachIds: ids.filter(Boolean), coachIdsMinted: minted };
}

/**
 * Everything Build Coach Availability changes except formatting (validation, conditional
 * formatting, Spruce Up), so it can run against a fake sheet in the offline harness.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {function(): number} [random]
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, eventCount: number, coachIdsMinted: number,
 *   rowsAdded: number, columnsCreated: string[], answersCarriedOver: number,
 *   availabilityColumns: number[], noteColumns: number[]}}
 */
function buildCoachAvailabilityCore_(ss, random) {
  const coaches = ensureCoaches_(ss, random);
  const events = getCoachAvailabilityEvents_(ss);
  const sheet = ensureSheetWithHeaders_(ss, CONFIG.coachAvailability.sheetName, Object.values(COACH_AVAILABILITY_ROW_HEADERS));

  // Columns: append any missing pair, in event order.
  const existing = getExistingColumns(sheet);
  const columnsCreated = [];
  const carryOvers = [];
  let next = sheet.getLastColumn() + 1;
  events.forEach(event => {
    const hdr = coachAvailabilityHeaders(event);
    const createdPair = !existing[hdr.availabilityHeader];
    [hdr.availabilityHeader, hdr.noteHeader].forEach(h => {
      if (existing[h]) return;
      sheet.getRange(1, next).setValue(h).setFontWeight('bold');
      existing[h] = next;
      columnsCreated.push(h);
      next++;
    });
    // A blank-Team game later split into per-team rows: carry each coach's all-teams answer into
    // the new team column (grill Q26). The old column stays; the portal just stops showing it.
    if (createdPair && event.kind === 'game' && event.team && (event.ordinalForDate || 1) === 1) {
      const allTeams = coachAvailabilityHeaders({ kind: 'game', formattedDate: event.formattedDate });
      if (existing[allTeams.availabilityHeader]) {
        carryOvers.push({ from: existing[allTeams.availabilityHeader], to: existing[hdr.availabilityHeader] });
      }
      if (existing[allTeams.noteHeader]) {
        carryOvers.push({ from: existing[allTeams.noteHeader], to: existing[hdr.noteHeader] });
      }
    }
  });

  // Rows: one per CoachID not yet present, in Coaches order. Name is a Coaches lookup formula.
  const coachesCols = getExistingColumns(coaches.sheet);
  const coachesIdLetter = getColumnLetter(coachesCols[COACH_COLUMNS.coachId]);
  const coachesNameLetter = getColumnLetter(coachesCols[COACH_COLUMNS.name]);
  const idCol = existing[COACH_AVAILABILITY_ROW_HEADERS.coachId];
  const nameCol = existing[COACH_AVAILABILITY_ROW_HEADERS.name];
  const idLetter = getColumnLetter(idCol);
  const nameFormula = row => playerIdLookupFormula(idLetter, row, CONFIG.coaches.sheetName, coachesIdLetter, coachesNameLetter);

  const lastRow = sheet.getLastRow();
  const present = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, idCol, lastRow - 1, 1).getValues().forEach(r => {
      const id = r[0] === null || r[0] === undefined ? '' : String(r[0]).trim();
      if (id) present.add(id);
    });
  }
  const missing = coaches.coachIds.filter(id => !present.has(id));
  missing.forEach((id, i) => {
    const row = Math.max(lastRow, 1) + 1 + i;
    sheet.getRange(row, idCol).setValue(id);
    sheet.getRange(row, nameCol).setFormula(nameFormula(row));
  });

  let answersCarriedOver = 0;
  const totalRows = sheet.getLastRow() - 1;
  if (totalRows > 0) {
    carryOvers.forEach(c => {
      const values = sheet.getRange(2, c.from, totalRows, 1).getValues();
      answersCarriedOver += values.filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined).length;
      sheet.getRange(2, c.to, totalRows, 1).setValues(values);
    });
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const availabilityColumns = [];
  const noteColumns = [];
  headers.forEach((h, i) => {
    if (/ (Practice|Game)( \d+)?$/.test(h)) availabilityColumns.push(i + 1);
    else if (/ (Practice|Game)( \d+)? Note$/.test(h)) noteColumns.push(i + 1);
  });

  return {
    sheet: sheet,
    eventCount: events.length,
    coachIdsMinted: coaches.coachIdsMinted,
    rowsAdded: missing.length,
    columnsCreated: columnsCreated,
    answersCarriedOver: answersCarriedOver,
    availabilityColumns: availabilityColumns,
    noteColumns: noteColumns
  };
}

/**
 * Menu entry point. Safe to re-run whenever Practice Info, Game Info, or Coaches change.
 */
function buildCoachAvailability() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = buildCoachAvailabilityCore_(ss);
    const sheet = result.sheet;

    const numRows = Math.max(sheet.getLastRow(), 100) - 1;
    if (result.availabilityColumns.length > 0) {
      const dv = SpreadsheetApp.newDataValidation()
        .requireValueInList(AVAILABILITY_VALIDATION_OPTIONS.map(o => o.value), true)
        .setAllowInvalid(false)
        .setHelpText('Select your availability')
        .build();
      applyDataValidationToColumnRanges_(sheet, result.availabilityColumns, numRows, dv);
    }
    result.noteColumns.forEach(c => sheet.getRange(2, c, numRows, 1).clearDataValidations());
    sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), sheet.getLastColumn()).setWrap(true);
    try {
      applySpruceUpFormatting(sheet);
    } catch (error) {
      console.warn('⚠️ Could not apply Format Spruce Up formatting:', error.message);
    }
    applyManagedAvailabilityCfRules_(sheet, result.availabilityColumns.length > 0, false);

    let message = `Processed ${result.eventCount} practice and game event(s).\n\n`;
    message += result.columnsCreated.length > 0
      ? `📊 Created: ${result.columnsCreated.join(', ')}`
      : 'No new columns needed.';
    message += `\n\n👥 Coach rows: ${result.rowsAdded} added. CoachIDs minted on "${CONFIG.coaches.sheetName}": ${result.coachIdsMinted}. Rows and columns are never deleted.`;
    if (result.answersCarriedOver > 0) {
      message += `\n\n↪️ ${result.answersCarriedOver} answer(s) and note(s) copied from all-teams game columns into new per-team columns.`;
    }
    ui.alert('Coach Availability Updated!', message, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Error building Coach Availability:', error);
    ui.alert('Error', `Failed to build Coach Availability: ${error.message}`, ui.ButtonSet.OK);
  }
}
