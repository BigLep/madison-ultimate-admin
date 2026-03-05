/**
 * Sync Game Info sheet to Google Calendar
 * Creates, updates, and deletes game and warmup events on the team calendar to match the sheet.
 */

const GAME_EVENT_TITLE_PREFIX = '🎯';
const GAME_TBD_TITLE = '🎯 TBD Game';
const GAME_WARMUP_TITLE = '🥏 Game Warmup';
const TBD_DEFAULT_TIME_HOUR = 9;
const TBD_DEFAULT_TIME_MINUTE = 0;

// Game Info column headers (must match sheet)
const GAME_INFO_COLUMNS = {
  DATE: 'Date',
  GAME_NUM: 'Game #',
  WARMUP_ARRIVAL: 'Warmup Arrival',
  GAME_START: 'Game Start',
  DONE_BY: 'Done By',
  FIELD_NAME: 'Field Name',
  FIELD_LOCATION: 'Field Location',
  GAME_NOTE: 'Game Note',
  OPPONENT: 'Opponent',
  OPPONENT_TEAM_PAGE: 'Oponent Team Page'  // exact header as in sheet
};

// Fields sheet column headers (lookup for Google Map URL and DiscNW URL by Field Name)
const FIELDS_SHEET_COLUMNS = {
  FIELD_NAME: 'Field Name',
  GOOGLE_MAP_URL: 'Google Map URL',
  DISCNW_URL: 'DiscNW URL'
};

/**
 * Menu entry point: sync Game Info sheet to calendar.
 */
function createGameCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.gameInfo.sheetName);
    if (!infoSheet) {
      ui.alert('Sheet not found', `"${CONFIG.gameInfo.sheetName}" sheet was not found.`, ui.ButtonSet.OK);
      return;
    }

    const eventSpecs = getGameEventSpecsFromSheet(ss, infoSheet);
    if (eventSpecs.length === 0) {
      ui.alert('No game rows', 'No valid game rows found in the sheet (check Date column and skip "Bye" rows).', ui.ButtonSet.OK);
      return;
    }

    const calendar = CalendarApp.getCalendarById(TEAM_CALENDAR_ID);
    if (!calendar) {
      ui.alert('Calendar not found', 'Could not open team calendar. Check that TEAM_CALENDAR_ID is correct and you have access.', ui.ButtonSet.OK);
      return;
    }

    const isGameEventOurs = function (e) {
      const title = e.getTitle();
      if (title.indexOf(GAME_EVENT_TITLE_PREFIX) === 0) return true;
      if (title === GAME_WARMUP_TITLE) return true;
      // Legacy or alternate-format TBD events (e.g. "TBD Game, 9am") so we can match/delete them
      if (title === 'TBD Game' || title.indexOf('TBD Game') === 0) return true;
      return false;
    };
    const result = syncEventsToCalendar(calendar, eventSpecs, isGameEventOurs);

    const message = 'Created: ' + result.created + '\nUpdated: ' + result.updated + '\nDeleted: ' + result.deleted;
    ui.alert('Calendar synced', message, ui.ButtonSet.OK);
  } catch (err) {
    console.error('createGameCalendarEvents', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

/**
 * Load Fields sheet into a lookup map: field name (trimmed, lowercased) -> { googleMapUrl, discNWUrl }.
 * If sheet is missing or has no data, returns {} so sync can continue without URLs.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @return {Object.<string, {googleMapUrl: string, discNWUrl: string}>}
 */
function getFieldsLookup(ss) {
  const sheet = CONFIG.fieldsSheet && ss.getSheetByName(CONFIG.fieldsSheet.sheetName);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = function (name) {
    const i = headerRow.findIndex(function (h) { return h && h.toString().trim() === name; });
    return i === -1 ? -1 : i;
  };
  const colFieldName = col(FIELDS_SHEET_COLUMNS.FIELD_NAME);
  const colGoogleMap = col(FIELDS_SHEET_COLUMNS.GOOGLE_MAP_URL);
  const colDiscNW = col(FIELDS_SHEET_COLUMNS.DISCNW_URL);
  if (colFieldName === -1) return {};
  const lastCol = Math.max(colFieldName + 1, colGoogleMap >= 0 ? colGoogleMap + 1 : 0, colDiscNW >= 0 ? colDiscNW + 1 : 0);
  const data = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  const lookup = {};
  data.forEach(function (row) {
    const name = row[colFieldName] != null ? String(row[colFieldName]).trim() : '';
    if (!name) return;
    const key = name.toLowerCase();
    lookup[key] = {
      googleMapUrl: colGoogleMap >= 0 && row[colGoogleMap] ? String(row[colGoogleMap]).trim() : '',
      discNWUrl: colDiscNW >= 0 && row[colDiscNW] ? String(row[colDiscNW]).trim() : ''
    };
  });
  return lookup;
}

/**
 * Build calendar event location string: plain text "Field Name (Field Location)" (no HTML; URLs go in description).
 */
function buildGameLocationString(fieldName, fieldLocation) {
  return fieldName + (fieldLocation ? ' (' + fieldLocation + ')' : '');
}

/**
 * Build calendar event description: Google Maps link, DiscNW field page, Opponent link, Game Note (omit empty lines).
 */
function buildGameDescription(googleMapUrl, discNWUrl, opponentName, opponentUrl, gameNote) {
  const parts = [];
  if (googleMapUrl) parts.push('Google Maps link: ' + googleMapUrl);
  if (discNWUrl) parts.push('DiscNW field page: ' + discNWUrl);
  if (opponentName || opponentUrl) {
    const link = opponentUrl ? '<a href="' + opponentUrl + '">' + (opponentName || opponentUrl) + '</a>' : (opponentName || '');
    parts.push('Opponent: ' + link);
  }
  if (gameNote) parts.push('Game Note: ' + gameNote);
  return parts.join('\n\n');
}

/**
 * Get game event specs from Game Info sheet using getDatesFromInfoSheet and Fields lookup.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {SpreadsheetApp.Sheet} infoSheet
 * @return {Array<{title: string, startTime: Date, endTime: Date, location: string, description: string, isAllDay: boolean}>}
 */
function getGameEventSpecsFromSheet(ss, infoSheet) {
  const fieldsLookup = getFieldsLookup(ss);
  const dateInfos = getDatesFromInfoSheet(ss, GAME_AVAILABILITY_CONFIG);
  const headerRow = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
  const col = function (name) {
    const i = headerRow.findIndex(function (h) {
      return h && h.toString().trim() === name;
    });
    return i === -1 ? -1 : i;
  };
  const colDate = col(GAME_INFO_COLUMNS.DATE);
  const colWarmupArrival = col(GAME_INFO_COLUMNS.WARMUP_ARRIVAL);
  const colGameStart = col(GAME_INFO_COLUMNS.GAME_START);
  const colDoneBy = col(GAME_INFO_COLUMNS.DONE_BY);
  const colFieldName = col(GAME_INFO_COLUMNS.FIELD_NAME);
  const colFieldLocation = col(GAME_INFO_COLUMNS.FIELD_LOCATION);
  const colGameNote = col(GAME_INFO_COLUMNS.GAME_NOTE);
  const colOpponent = col(GAME_INFO_COLUMNS.OPPONENT);
  const colOpponentTeamPage = col(GAME_INFO_COLUMNS.OPPONENT_TEAM_PAGE);

  if (colDate === -1) {
    return [];
  }
  const lastRow = infoSheet.getLastRow();
  const lastCol = Math.max(
    colDate + 1,
    colWarmupArrival >= 0 ? colWarmupArrival + 1 : 0,
    colGameStart >= 0 ? colGameStart + 1 : 0,
    colDoneBy >= 0 ? colDoneBy + 1 : 0,
    colFieldName >= 0 ? colFieldName + 1 : 0,
    colFieldLocation >= 0 ? colFieldLocation + 1 : 0,
    colGameNote >= 0 ? colGameNote + 1 : 0,
    colOpponent >= 0 ? colOpponent + 1 : 0,
    colOpponentTeamPage >= 0 ? colOpponentTeamPage + 1 : 0
  );
  const dataRange = infoSheet.getRange(2, 1, lastRow, lastCol);
  const data = dataRange.getValues();

  const specs = [];
  dateInfos.forEach(function (dateInfo) {
    const dataIndex = dateInfo.rowIndex - 2;
    if (dataIndex < 0 || dataIndex >= data.length) return;
    const row = data[dataIndex];
    const dateObj = dateInfo.date;

    const gameStartVal = colGameStart >= 0 ? row[colGameStart] : null;
    const opponentVal = colOpponent >= 0 ? row[colOpponent] : null;
    const opponentStr = opponentVal != null ? String(opponentVal).trim() : '';
    const gameStartStr = gameStartVal != null ? String(gameStartVal).trim() : '';
    const isTBD = gameStartStr === '' || gameStartStr.toLowerCase() === 'tbd';

    const fieldName = colFieldName >= 0 && row[colFieldName] ? String(row[colFieldName]).trim() : '';
    const fieldLocation = colFieldLocation >= 0 && row[colFieldLocation] ? String(row[colFieldLocation]).trim() : '';
    const gameNote = colGameNote >= 0 && row[colGameNote] ? String(row[colGameNote]).trim() : '';
    const opponentTeamPageUrl = colOpponentTeamPage >= 0 && row[colOpponentTeamPage] ? String(row[colOpponentTeamPage]).trim() : '';
    const fieldInfo = fieldName ? fieldsLookup[fieldName.toLowerCase()] : null;
    const googleMapUrl = fieldInfo ? fieldInfo.googleMapUrl : '';
    const discNWUrl = fieldInfo ? fieldInfo.discNWUrl : '';
    const locationStr = buildGameLocationString(fieldName, fieldLocation);
    const mainDescription = buildGameDescription(googleMapUrl, discNWUrl, opponentStr, opponentTeamPageUrl, gameNote);
    const warmupDescription = discNWUrl ? 'DiscNW field page: ' + discNWUrl : '';

    if (isTBD) {
      const tbdStart = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), TBD_DEFAULT_TIME_HOUR, TBD_DEFAULT_TIME_MINUTE, 0);
      specs.push({
        title: GAME_TBD_TITLE,
        startTime: tbdStart,
        endTime: tbdStart,
        location: '',
        description: '',
        isAllDay: false
      });
      return;
    }

    const gameStartTime = parseTimeOnDate(dateObj, gameStartVal);
    const doneByTime = colDoneBy >= 0 ? parseTimeOnDate(dateObj, row[colDoneBy]) : null;
    if (!gameStartTime) {
      const tbdStart = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), TBD_DEFAULT_TIME_HOUR, TBD_DEFAULT_TIME_MINUTE, 0);
      specs.push({
        title: GAME_TBD_TITLE,
        startTime: tbdStart,
        endTime: tbdStart,
        location: locationStr,
        description: mainDescription,
        isAllDay: false
      });
      return;
    }
    const gameEndTime = doneByTime && doneByTime > gameStartTime ? doneByTime : new Date(gameStartTime.getTime() + 60 * 60 * 1000);

    specs.push({
      title: '🎯 Game vs. ' + (opponentStr || 'TBD'),
      startTime: gameStartTime,
      endTime: gameEndTime,
      location: locationStr,
      description: mainDescription,
      isAllDay: false
    });

    const warmupVal = colWarmupArrival >= 0 ? row[colWarmupArrival] : null;
    if (warmupVal && (typeof warmupVal === 'string' ? warmupVal.trim() : warmupVal)) {
      const warmupStart = parseTimeOnDate(dateObj, warmupVal);
      if (warmupStart && warmupStart < gameStartTime) {
        specs.push({
          title: GAME_WARMUP_TITLE,
          startTime: warmupStart,
          endTime: gameStartTime,
          location: locationStr,
          description: warmupDescription,
          isAllDay: false
        });
      }
    }
  });
  return specs;
}

/**
 * Parse a time value (Date or string like "10:45" or "10:00 AM") and return a Date on the given day.
 * @param {Date} dateObj - The date (year/month/day)
 * @param {Date|string} timeValue - Time as Date or string
 * @return {Date|null}
 */
function parseTimeOnDate(dateObj, timeValue) {
  if (timeValue == null || timeValue === '') return null;
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    return combineDateAndTime(dateObj, timeValue);
  }
  const str = String(timeValue).trim().replace(/\u202f/g, ' ');
  if (!str) return null;
  // Try "10:45" or "10:45 AM" / "12:30 PM"
  const match = str.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (match) {
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const ampm = match[3] ? match[3].toUpperCase() : null;
    if (ampm === 'PM' && hours < 12) hours += 12;
    if (ampm === 'AM' && hours === 12) hours = 0;
    return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), hours, minutes, 0);
  }
  const asDate = new Date(timeValue);
  if (!isNaN(asDate.getTime())) {
    return combineDateAndTime(dateObj, asDate);
  }
  return null;
}
