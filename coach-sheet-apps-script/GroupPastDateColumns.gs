/**
 * Group Past Date Columns: folds the active sheet's Past Date Columns (see CONTEXT.md) into one
 * collapsed column group, so finished practices and games are hidden behind the group toggle.
 * Generic, works on any sheet whose date headers start with "M/D" (Practice, Game, and Coach
 * Availability, and Generated Rosters). Never moves, deletes, or edits a column.
 *
 * Re-running replaces the previous group: every column group overlapping the date columns is
 * removed first, then one new group is created. Groups outside the date columns are left alone.
 * When there is nothing to group, or a future date sits before a past one, nothing changes.
 */

/** Menu entry point: groups the active sheet's Past Date Columns and reports with a toast. */
function groupPastDateColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  try {
    const result = groupPastDateColumnsOnSheet_(sheet, new Date());
    if (result.status === 'outOfOrder') {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Group Past Date Columns',
        `"${result.header}" is a future date but comes before a past date column, so nothing was grouped on ${sheet.getName()}. Move that column after the past dates and run it again.`,
        ui.ButtonSet.OK);
      return;
    }
    if (result.status === 'none') {
      ss.toast(`No past date columns on ${sheet.getName()}; nothing changed.`, '🗂️ Group Past Date Columns', 5);
      return;
    }
    ss.toast(`Grouped ${result.columnCount} columns (${result.firstLabel} to ${result.lastLabel}) on ${sheet.getName()}.`, '🗂️ Group Past Date Columns', 5);
  } catch (error) {
    console.error('Error grouping past date columns:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Failed to group past date columns: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Group one sheet's Past Date Columns: remove column groups overlapping the date columns, then
 * create one collapsed group over the past span.
 * @param {Sheet} sheet
 * @param {Date} today
 * @return {{status: string, header?: string, columnCount?: number, firstLabel?: string, lastLabel?: string}}
 */
function groupPastDateColumnsOnSheet_(sheet, today) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { status: 'none' };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const span = findPastDateColumnSpan_(headers, today);
  if (span.status !== 'ok') return span;

  // 1-based columns from here on.
  for (let col = span.dateFirst + 1; col <= span.dateLast + 1; col++) {
    let group;
    let guard = 0;
    while ((group = sheet.getColumnGroup(col, 1)) && guard++ < 10) group.remove();
  }
  const first = span.first + 1;
  const count = span.last - span.first + 1;
  sheet.getRange(1, first, 1, count).shiftColumnGroupDepth(1);
  sheet.getColumnGroup(first, 1).collapse();
  return { status: 'ok', columnCount: count, firstLabel: span.firstLabel, lastLabel: span.lastLabel };
}

/**
 * Find the Past Date Columns in a header row. Pure.
 * A header is a date column when it is a Date value, or a string starting with "M/D" followed by
 * end of text or whitespace (the year is today's). Past means strictly before today.
 * @param {Array} headers - Row 1 values
 * @param {Date} today
 * @return {{status: 'ok', first: number, last: number, dateFirst: number, dateLast: number, firstLabel: string, lastLabel: string}
 *   | {status: 'none'} | {status: 'outOfOrder', header: string}} 0-based column indices
 */
function findPastDateColumnSpan_(headers, today) {
  const todayKey = dateKey_(today.getFullYear(), today.getMonth() + 1, today.getDate());
  const dated = [];
  headers.forEach((h, i) => {
    const d = parseDateHeader_(h, today.getFullYear());
    if (d) dated.push({ index: i, key: dateKey_(d.year, d.month, d.day), label: `${d.month}/${d.day}`, header: String(h) });
  });
  const past = dated.filter(d => d.key < todayKey);
  if (past.length === 0) return { status: 'none' };
  const first = past[0];
  const last = past[past.length - 1];
  const early = dated.find(d => d.key >= todayKey && d.index < last.index);
  if (early) return { status: 'outOfOrder', header: early.header };
  return {
    status: 'ok',
    first: first.index,
    last: last.index,
    dateFirst: dated[0].index,
    dateLast: dated[dated.length - 1].index,
    firstLabel: first.label,
    lastLabel: last.label
  };
}

/**
 * Parse a header cell as a date: a Date value, or "M/D" at the start of a string followed by end
 * of text or whitespace ("9/26", "9/26 Note", "9/26 Availability (Game 2)", "9/26 Silver Game").
 * @param {*} header
 * @param {number} year - Year for "M/D" headers
 * @return {{year: number, month: number, day: number}|null}
 */
function parseDateHeader_(header, year) {
  if (Object.prototype.toString.call(header) === '[object Date]') {
    if (isNaN(header.getTime())) return null;
    return { year: header.getFullYear(), month: header.getMonth() + 1, day: header.getDate() };
  }
  const match = String(header == null ? '' : header).trim().match(/^(\d{1,2})\/(\d{1,2})(?:\s|$)/);
  if (!match) return null;
  const month = Number(match[1]);
  const day = Number(match[2]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { year: year, month: month, day: day };
}

/** Sortable number for a calendar day. */
function dateKey_(year, month, day) {
  return year * 10000 + month * 100 + day;
}
