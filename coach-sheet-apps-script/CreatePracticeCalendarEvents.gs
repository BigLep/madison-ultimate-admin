/**
 * Sync Practice Info sheet to Google Calendar
 * Creates, updates, and deletes "🏃 Practice" events on the MadDogs calendar to match the sheet.
 */

// Calendar and event title - change per season or calendar as needed
const PRACTICE_CALENDAR_ID = '21081b4ccff3c7ad50dc835ce259ff76a09e0f05d1a66d727fafff195a7af612@group.calendar.google.com';
const PRACTICE_EVENT_TITLE = '🥏🏃 Practice';

// Match tolerance in ms (2 minutes) when matching sheet row to existing calendar event
const START_TIME_MATCH_MS = 2 * 60 * 1000;

/**
 * Menu entry point: sync Practice Info sheet to calendar.
 */
function createPracticeCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.practiceInfo.sheetName);
    if (!infoSheet) {
      ui.alert('Sheet not found', `"${CONFIG.practiceInfo.sheetName}" sheet was not found.`, ui.ButtonSet.OK);
      return;
    }

    const practiceRows = getPracticeRowsFromSheet(ss, infoSheet);
    if (practiceRows.length === 0) {
      ui.alert('No practice rows', 'No valid practice rows found in the sheet (check Date column and skip "Bye" rows).', ui.ButtonSet.OK);
      return;
    }

    const calendar = CalendarApp.getCalendarById(PRACTICE_CALENDAR_ID);
    if (!calendar) {
      ui.alert('Calendar not found', `Could not open calendar. Check that PRACTICE_CALENDAR_ID is correct and you have access.`, ui.ButtonSet.OK);
      return;
    }

    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);
    const endDate = new Date(startOfToday);
    endDate.setMonth(endDate.getMonth() + 6);
    const existingEvents = calendar.getEvents(startOfToday, endDate).filter(function (e) {
      return e.getTitle() === PRACTICE_EVENT_TITLE;
    });

    let created = 0;
    let updated = 0;
    const matchedEventIds = {};

    practiceRows.forEach(function (row) {
      const matched = findMatchingEvent(existingEvents, row.startTime, matchedEventIds);
      if (matched) {
        updateCalendarEvent(matched, row);
        matchedEventIds[matched.getId()] = true;
        updated++;
      } else {
        createCalendarEvent(calendar, row);
        created++;
      }
    });

    let deleted = 0;
    existingEvents.forEach(function (e) {
      if (!matchedEventIds[e.getId()]) {
        e.deleteEvent();
        deleted++;
      }
    });

    const message = 'Created: ' + created + '\nUpdated: ' + updated + '\nDeleted: ' + deleted;
    ui.alert('Calendar synced', message, ui.ButtonSet.OK);
  } catch (err) {
    console.error('createPracticeCalendarEvents', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

/**
 * Get practice rows from sheet using getDatesFromInfoSheet, then enrich with Location, Start, End, etc.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {SpreadsheetApp.Sheet} infoSheet
 * @return {Array<{startTime: Date, endTime: Date, location: string, locationUrl: string}>}
 */
function getPracticeRowsFromSheet(ss, infoSheet) {
  const dateInfos = getDatesFromInfoSheet(ss, PRACTICE_AVAILABILITY_CONFIG);
  const headerRow = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
  const col = function (name) {
    const i = headerRow.findIndex(function (h) {
      return h && h.toString().trim().toLowerCase() === name.toLowerCase();
    });
    return i === -1 ? -1 : i;
  };
  const colDate = headerRow.findIndex(function (h) { return h && h.toString().toLowerCase().includes('date'); });
  const colLocation = col('Location');
  const colLocationUrl = col('Location URL');
  const colStart = col('Start');
  const colEnd = col('End');

  const lastCol = Math.max(colDate + 1, colLocation >= 0 ? colLocation + 1 : 0, colLocationUrl >= 0 ? colLocationUrl + 1 : 0, colStart >= 0 ? colStart + 1 : 0, colEnd >= 0 ? colEnd + 1 : 0);
  if (lastCol <= 0) {
    return [];
  }
  const lastRow = infoSheet.getLastRow();
  const dataRange = infoSheet.getRange(2, 1, lastRow, lastCol);
  const data = dataRange.getValues();

  const rows = [];
  dateInfos.forEach(function (dateInfo) {
    const dataIndex = dateInfo.rowIndex - 2; // sheet row 2 = data[0]
    if (dataIndex < 0 || dataIndex >= data.length) return;
    const row = data[dataIndex];
    const dateObj = dateInfo.date;
    let startTime = dateObj;
    let endTime = new Date(dateObj.getTime());
    endTime.setHours(endTime.getHours() + 1);
    let isAllDay = true;

    if (colStart >= 0 && row[colStart]) {
      const startVal = row[colStart];
      const startDate = startVal instanceof Date ? startVal : new Date(startVal);
      if (!isNaN(startDate.getTime())) {
        startTime = combineDateAndTime(dateObj, startDate);
        isAllDay = false;
      }
    }
    if (colEnd >= 0 && row[colEnd]) {
      const endVal = row[colEnd];
      const endDate = endVal instanceof Date ? endVal : new Date(endVal);
      if (!isNaN(endDate.getTime())) {
        endTime = combineDateAndTime(dateObj, endDate);
        isAllDay = false;
      }
    }
    if (endTime <= startTime) {
      endTime = new Date(startTime.getTime());
      endTime.setHours(endTime.getHours() + 1);
    }

    const location = colLocation >= 0 && row[colLocation] ? String(row[colLocation]).trim() : '';
    const locationUrl = colLocationUrl >= 0 && row[colLocationUrl] ? String(row[colLocationUrl]).trim() : '';
    const description = locationUrl ? locationUrl : '';

    rows.push({
      startTime: startTime,
      endTime: endTime,
      location: location,
      locationUrl: locationUrl,
      description: description,
      isAllDay: isAllDay
    });
  });
  return rows;
}

/**
 * Combine date from one Date (day) and time from another (time-of-day).
 * @param {Date} datePart - supplies year, month, date
 * @param {Date} timePart - supplies hours, minutes, seconds
 */
function combineDateAndTime(datePart, timePart) {
  return new Date(
    datePart.getFullYear(),
    datePart.getMonth(),
    datePart.getDate(),
    timePart.getHours(),
    timePart.getMinutes(),
    timePart.getSeconds()
  );
}

/**
 * Find an existing event that matches the given start time (within tolerance).
 * @param {CalendarEvent[]} events
 * @param {Date} startTime
 * @param {Object} alreadyMatched - set of event IDs already matched
 * @return {CalendarEvent|null}
 */
function findMatchingEvent(events, startTime, alreadyMatched) {
  const t = startTime.getTime();
  for (let i = 0; i < events.length; i++) {
    if (alreadyMatched[events[i].getId()]) continue;
    const diff = Math.abs(events[i].getStartTime().getTime() - t);
    if (diff <= START_TIME_MATCH_MS) return events[i];
  }
  return null;
}

/**
 * Create a new calendar event for a practice row.
 */
function createCalendarEvent(calendar, row) {
  if (row.isAllDay) {
    calendar.createAllDayEvent(PRACTICE_EVENT_TITLE, row.startTime, row.startTime, { description: row.description || '', location: row.location || '' });
  } else {
    calendar.createEvent(PRACTICE_EVENT_TITLE, row.startTime, row.endTime, { description: row.description || '', location: row.location || '' });
  }
}

/**
 * Update an existing calendar event to match the sheet row.
 */
function updateCalendarEvent(event, row) {
  event.setLocation(row.location || '');
  event.setDescription(row.description || '');
  if (!row.isAllDay) {
    event.setTime(row.startTime, row.endTime);
  }
}
