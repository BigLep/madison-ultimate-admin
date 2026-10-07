/**
 * Sort Date Columns: reorders the active sheet's contiguous run of date columns (headers starting
 * with "M/D", parsed by parseDateHeader_ in GroupPastDateColumns.gs) into chronological order.
 * The sort is stable, so columns sharing a date (an event and its Note, a game's Availability,
 * Activation Status, and Note, or several team games on one day) keep their relative order and
 * stay together. Columns are moved whole with moveColumns, so values, formatting, validation, and
 * notes travel with them; nothing is deleted or edited.
 *
 * When date columns are split into more than one run by a non-date column, nothing changes, since
 * it is unclear which run the coach means.
 */

/** Menu entry point: sorts the active sheet's date columns and reports with a toast. */
function sortDateColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  try {
    const result = sortDateColumnsOnSheet_(sheet, new Date());
    if (result.status === 'split') {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Sort Date Columns',
        `The date columns on ${sheet.getName()} are split by "${result.header}", so nothing was sorted. Move the non-date column out of the dates and run it again.`,
        ui.ButtonSet.OK);
      return;
    }
    if (result.status === 'none') {
      ss.toast(`No date columns on ${sheet.getName()}; nothing changed.`, '📆 Sort Date Columns', 5);
      return;
    }
    if (result.status === 'sorted') {
      ss.toast(`The ${result.columnCount} date columns on ${sheet.getName()} are already in order.`, '📆 Sort Date Columns', 5);
      return;
    }
    ss.toast(`Sorted ${result.columnCount} date columns on ${sheet.getName()} (${result.moveCount} move${result.moveCount !== 1 ? 's' : ''}).`, '📆 Sort Date Columns', 5);
  } catch (error) {
    console.error('Error sorting date columns:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Failed to sort date columns: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Sort one sheet's date columns in place.
 * @param {Sheet} sheet
 * @param {Date} today - only its year is used, for "M/D" headers
 * @return {{status: string, header?: string, columnCount?: number, moveCount?: number}}
 */
function sortDateColumnsOnSheet_(sheet, today) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { status: 'none' };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const run = findDateColumnRun_(headers, today.getFullYear());
  if (run.status !== 'ok') return run;

  const moves = planDateColumnMoves_(run.keys);
  // 1-based columns from here on. Every move's destination is left of its source, and
  // moveColumns reads the destination in pre-move coordinates, so start + to + 1 is exact.
  moves.forEach(function (m) {
    sheet.moveColumns(sheet.getRange(1, run.start + m.from + 1, 1, m.count), run.start + m.to + 1);
  });
  return { status: moves.length ? 'ok' : 'sorted', columnCount: run.keys.length, moveCount: moves.length };
}

/**
 * Find the single contiguous run of date columns in a header row. Pure.
 * @param {Array} headers - Row 1 values
 * @param {number} year - Year for "M/D" headers
 * @return {{status: 'ok', start: number, keys: Array<number>} | {status: 'none'} | {status: 'split', header: string}}
 *   start is a 0-based column index; keys are dateKey_ values, one per column in the run
 */
function findDateColumnRun_(headers, year) {
  const keys = headers.map(function (h) {
    const d = parseDateHeader_(h, year);
    return d ? dateKey_(d.year, d.month, d.day) : null;
  });
  const start = keys.findIndex(function (k) { return k !== null; });
  if (start === -1) return { status: 'none' };
  let end = start;
  while (end + 1 < keys.length && keys[end + 1] !== null) end++;
  const later = keys.findIndex(function (k, i) { return i > end && k !== null; });
  if (later !== -1) return { status: 'split', header: String(headers[end + 1]) };
  return { status: 'ok', start: start, keys: keys.slice(start, end + 1) };
}

/**
 * Plan the column block moves that stable-sort a run by key. Pure.
 * Each move takes `count` columns at position `from` (in the order as it stands after the earlier
 * moves) and places them at position `to`, always with to < from. Consecutive columns that already
 * sit in their sorted order move together as one block, so an event and its Note take one move.
 * @param {Array<number>} keys - sort key per column, in current order
 * @return {Array<{from: number, to: number, count: number}>} positions relative to the run
 */
function planDateColumnMoves_(keys) {
  const desired = keys.map(function (k, i) { return i; })
    .sort(function (a, b) { return keys[a] - keys[b] || a - b; });
  const current = keys.map(function (k, i) { return i; });
  const moves = [];
  for (let t = 0; t < desired.length; t++) {
    if (current[t] === desired[t]) continue;
    const from = current.indexOf(desired[t]);
    let count = 1;
    while (from + count < current.length && t + count < desired.length && current[from + count] === desired[t + count]) count++;
    const block = current.splice(from, count);
    current.splice.apply(current, [t, 0].concat(block));
    moves.push({ from: from, to: t, count: count });
  }
  return moves;
}
