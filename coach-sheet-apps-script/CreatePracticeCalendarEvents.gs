/**
 * Sync Practice Info sheet to Google Calendar
 * Creates, updates, and deletes "🏃 Practice" events on the MadDogs calendar to match the sheet.
 */

// Shared team calendar for practice and game events - change per season as needed
const TEAM_CALENDAR_ID = '21081b4ccff3c7ad50dc835ce259ff76a09e0f05d1a66d727fafff195a7af612@group.calendar.google.com';
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

    const calendar = CalendarApp.getCalendarById(TEAM_CALENDAR_ID);
    if (!calendar) {
      ui.alert('Calendar not found', `Could not open calendar. Check that TEAM_CALENDAR_ID is correct and you have access.`, ui.ButtonSet.OK);
      return;
    }

    const eventSpecs = practiceRows.map(function (row) {
      return {
        title: PRACTICE_EVENT_TITLE,
        startTime: row.startTime,
        endTime: row.endTime,
        location: row.location || '',
        description: row.description || '',
        isAllDay: row.isAllDay,
        sheetEventId: row.sheetEventId || null,
        writeBack: row.writeBack || null
      };
    });
    const result = syncEventsToCalendar(calendar, eventSpecs, function (e) {
      return e.getTitle() === PRACTICE_EVENT_TITLE;
    });

    if (result.writeBacks && result.writeBacks.length > 0) {
      result.writeBacks.forEach(function (wb) {
        infoSheet.getRange(wb.row, wb.col).setValue(wb.value);
      });
    }

    const message = 'Created: ' + result.created + '\nUpdated: ' + result.updated + '\nDeleted: ' + result.deleted;
    ui.alert('Calendar synced', message, ui.ButtonSet.OK);
  } catch (err) {
    console.error('createPracticeCalendarEvents', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

/**
 * Shared sync: create/update/delete calendar events to match the given event specs.
 * Supports ID-based matching and writing event IDs back to the spreadsheet.
 * @param {Calendar} calendar - Google Calendar
 * @param {Array} eventSpecs - Each spec: { title, startTime, endTime, location, description, isAllDay, sheetEventId?, writeBack? }
 *   - sheetEventId: optional stored Google Calendar event ID; if present we match by ID first
 *   - writeBack: optional { row: number, col: number } (1-based); if present we add { row, col, value: eventId } to result.writeBacks
 * @param {function(CalendarEvent): boolean} isEventOurs - returns true if this event is managed by this sync
 * @return {{created: number, updated: number, deleted: number, writeBacks: Array<{row: number, col: number, value: string}>}}
 */
function syncEventsToCalendar(calendar, eventSpecs, isEventOurs) {
  const startOfToday = new Date();
  startOfToday.setHours(0, 0, 0, 0);
  const endDate = new Date(startOfToday);
  endDate.setMonth(endDate.getMonth() + 6);
  const existingEvents = calendar.getEvents(startOfToday, endDate).filter(isEventOurs);

  let created = 0;
  let updated = 0;
  const matchedEventIds = {};
  const writeBacks = [];

  eventSpecs.forEach(function (spec) {
    let matched = null;
    if (spec.sheetEventId && spec.sheetEventId.toString().trim()) {
      const id = spec.sheetEventId.toString().trim();
      matched = existingEvents.filter(function (e) { return e.getId() === id; })[0] || null;
    }
    if (!matched) {
      matched = findMatchingEventBySpec(existingEvents, spec, matchedEventIds);
    }
    if (matched) {
      matchedEventIds[matched.getId()] = true;
      if (!calendarEventMatchesSpec(matched, spec)) {
        updateCalendarEventFromSpec(matched, spec);
        updated++;
      }
      if (spec.writeBack) {
        writeBacks.push({ row: spec.writeBack.row, col: spec.writeBack.col, value: matched.getId() });
      }
    } else {
      const createdEvent = createCalendarEventFromSpec(calendar, spec);
      created++;
      if (spec.writeBack && createdEvent) {
        writeBacks.push({ row: spec.writeBack.row, col: spec.writeBack.col, value: createdEvent.getId() });
      }
    }
  });

  let deleted = 0;
  existingEvents.forEach(function (e) {
    if (!matchedEventIds[e.getId()]) {
      e.deleteEvent();
      deleted++;
    }
  });

  return { created: created, updated: updated, deleted: deleted, writeBacks: writeBacks };
}

/**
 * Find an existing event that matches the spec by start time (within tolerance) or by same day + same title.
 * @param {CalendarEvent[]} events
 * @param {Object} spec - has title, startTime
 * @param {Object} alreadyMatched - set of event IDs already matched
 * @return {CalendarEvent|null}
 */
function findMatchingEventBySpec(events, spec, alreadyMatched) {
  const startTime = spec.startTime;
  const specDayStart = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate()).getTime();
  for (let i = 0; i < events.length; i++) {
    if (alreadyMatched[events[i].getId()]) continue;
    const e = events[i];
    const eStart = e.getStartTime();
    const timeDiff = Math.abs(eStart.getTime() - startTime.getTime());
    if (timeDiff <= START_TIME_MATCH_MS) return e;
    const eDayStart = new Date(eStart.getFullYear(), eStart.getMonth(), eStart.getDate()).getTime();
    if (eDayStart === specDayStart && e.getTitle() === spec.title) return e;
  }
  return null;
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
  const colGoogleCalEventId = col('Google Calendar Event ID');

  const lastCol = Math.max(colDate + 1, colLocation >= 0 ? colLocation + 1 : 0, colLocationUrl >= 0 ? colLocationUrl + 1 : 0, colStart >= 0 ? colStart + 1 : 0, colEnd >= 0 ? colEnd + 1 : 0, colGoogleCalEventId >= 0 ? colGoogleCalEventId + 1 : 0);
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

    const eventIdVal = colGoogleCalEventId >= 0 ? row[colGoogleCalEventId] : null;
    const sheetEventId = eventIdVal != null && String(eventIdVal).trim() !== '' ? String(eventIdVal).trim() : null;
    const writeBack = (colGoogleCalEventId >= 0) ? { row: dateInfo.rowIndex, col: colGoogleCalEventId + 1 } : null;

    rows.push({
      startTime: startTime,
      endTime: endTime,
      location: location,
      locationUrl: locationUrl,
      description: description,
      isAllDay: isAllDay,
      sheetEventId: sheetEventId,
      writeBack: writeBack
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
 * Used when specs do not use findMatchingEventBySpec (e.g. Practice sync without writeBack).
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
 * Create a new calendar event from a generic event spec.
 * @param {Calendar} calendar
 * @param {{title: string, startTime: Date, endTime: Date, location: string, description: string, isAllDay: boolean}} spec
 * @return {CalendarEvent|null} The created event (for writing ID back to sheet)
 */
function createCalendarEventFromSpec(calendar, spec) {
  if (spec.isAllDay) {
    return calendar.createAllDayEvent(spec.title, spec.startTime, spec.startTime, { description: spec.description || '', location: spec.location || '' });
  }
  return calendar.createEvent(spec.title, spec.startTime, spec.endTime, { description: spec.description || '', location: spec.location || '' });
}

/**
 * Return true if the calendar event already matches the spec (no write needed).
 * @param {CalendarEvent} event
 * @param {{title: string, startTime: Date, endTime: Date, location: string, description: string, isAllDay: boolean}} spec
 * @return {boolean}
 */
function calendarEventMatchesSpec(event, spec) {
  if (event.getTitle() !== spec.title) return false;
  if ((event.getLocation() || '') !== (spec.location || '')) return false;
  if ((event.getDescription() || '') !== (spec.description || '')) return false;
  if (!spec.isAllDay) {
    const startDiff = Math.abs(event.getStartTime().getTime() - spec.startTime.getTime());
    const endDiff = Math.abs(event.getEndTime().getTime() - spec.endTime.getTime());
    if (startDiff > START_TIME_MATCH_MS || endDiff > START_TIME_MATCH_MS) return false;
  }
  return true;
}

/**
 * Update an existing calendar event to match an event spec.
 * @param {CalendarEvent} event
 * @param {{title: string, startTime: Date, endTime: Date, location: string, description: string, isAllDay: boolean}} spec
 */
function updateCalendarEventFromSpec(event, spec) {
  event.setTitle(spec.title);
  event.setLocation(spec.location || '');
  event.setDescription(spec.description || '');
  if (!spec.isAllDay) {
    event.setTime(spec.startTime, spec.endTime);
  }
}
