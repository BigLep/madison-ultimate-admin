/**
 * Draw Group Borders: a bottom border on the last row of each contiguous run of matching values
 * across one or more coach-chosen columns, within the current selection's rows. Generic, works on
 * any sheet and any columns; unlike the fixed Team/Gender (or Team/Activation Status/Gender)
 * grouping that addGroupBorders in BuildPracticeRoster.gs already draws for Practice Roster and
 * Game Roster Prep, this is the standalone version for every other sheet. See CONTEXT.md's
 * "Group Border" entry for the concept.
 */

const GROUP_BORDER_SETTINGS_PREFIX = 'groupBorders.';

const GROUP_BORDER_STYLES = [
  { value: 'SOLID', label: 'Solid (thin)' },
  { value: 'SOLID_MEDIUM', label: 'Solid (medium)' },
  { value: 'SOLID_THICK', label: 'Solid (thick)' },
  { value: 'DOTTED', label: 'Dotted' },
  { value: 'DASHED', label: 'Dashed' },
  { value: 'DOUBLE', label: 'Double' }
];
const GROUP_BORDER_DEFAULT_STYLE = 'SOLID_MEDIUM';
const GROUP_BORDER_DEFAULT_COLOR = '#000000';

/** Trim a cell value to a comparable string; null, undefined, and '' all become ''. */
function groupBorderCellText_(v) {
  if (v == null || v === '') return '';
  return String(v).trim();
}

/**
 * The group key for one row: normalized values of the given 0-based column indices, joined with a
 * separator that can never appear in a trimmed cell value.
 * @param {Array} row
 * @param {number[]} colIndices - 0-based indices into row
 * @return {string}
 */
function groupBorderRowKey_(row, colIndices) {
  return colIndices.map(function (i) { return groupBorderCellText_(row[i]); }).join('\u0001');
}

/**
 * Pure: the 0-based row indices (relative to `rows`) that are the last row of a contiguous group,
 * always including the final row. A group is any maximal run of consecutive rows sharing the same
 * key; if a value reappears later after a different value came between (the data isn't actually
 * sorted by these columns), that later run is simply its own group; there is no "is this sorted"
 * check, every value-change transition gets a border.
 * @param {Array<Array>} rows
 * @param {number[]} colIndices - 0-based column indices to group by
 * @return {number[]}
 */
function findGroupBoundaryRows_(rows, colIndices) {
  const boundaries = [];
  if (rows.length === 0 || colIndices.length === 0) return boundaries;
  for (let i = 1; i < rows.length; i++) {
    if (groupBorderRowKey_(rows[i], colIndices) !== groupBorderRowKey_(rows[i - 1], colIndices)) {
      boundaries.push(i - 1);
    }
  }
  boundaries.push(rows.length - 1);
  return boundaries;
}

/**
 * Pure: which header names should start checked in the dialog. A selection narrower than the
 * whole header row is a deliberate signal (the coach highlighted specific columns), so it wins;
 * a selection spanning every column (or touching no real headers) falls back to whatever this
 * sheet used last time.
 * @param {string[]} headers - full header row, in column order, blanks as ''
 * @param {number} selStartCol1 - 1-based selection start column
 * @param {number} selNumCols
 * @param {string[]} lastUsed - previously saved group-by header names for this sheet
 * @return {string[]}
 */
function defaultGroupByColumns_(headers, selStartCol1, selNumCols, lastUsed) {
  const spansWholeHeader = selNumCols >= headers.length;
  if (!spansWholeHeader) {
    const selected = [];
    for (let i = 0; i < selNumCols; i++) {
      const h = headers[selStartCol1 - 1 + i];
      if (h) selected.push(h);
    }
    if (selected.length > 0) return selected;
  }
  return lastUsed || [];
}

function groupBorderSettingsKey_(sheetName) {
  return GROUP_BORDER_SETTINGS_PREFIX + sheetName;
}

/** @return {{columns: string[], borderStyle: string, color: string}} */
function loadGroupBorderSettings_(sheetName) {
  const fallback = { columns: [], borderStyle: GROUP_BORDER_DEFAULT_STYLE, color: GROUP_BORDER_DEFAULT_COLOR };
  const raw = PropertiesService.getDocumentProperties().getProperty(groupBorderSettingsKey_(sheetName));
  if (!raw) return fallback;
  try {
    const parsed = JSON.parse(raw);
    return {
      columns: Array.isArray(parsed.columns) ? parsed.columns : fallback.columns,
      borderStyle: parsed.borderStyle || fallback.borderStyle,
      color: parsed.color || fallback.color
    };
  } catch (e) {
    return fallback;
  }
}

function saveGroupBorderSettings_(sheetName, settings) {
  PropertiesService.getDocumentProperties().setProperty(
    groupBorderSettingsKey_(sheetName),
    JSON.stringify({ columns: settings.columns, borderStyle: settings.borderStyle, color: settings.color })
  );
}

/** Menu entry point. */
function showDrawGroupBordersDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    const selection = sheet.getActiveRange();
    if (!selection) {
      ui.alert('No selection', 'Select the rows you want to draw group borders on, then run this again.', ui.ButtonSet.OK);
      return;
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      ui.alert('Empty sheet', 'This sheet has no columns.', ui.ButtonSet.OK);
      return;
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(groupBorderCellText_);

    // Data starts after the header row; a selection that includes row 1 is clamped to row 2+.
    let startRow = selection.getRow();
    let numRows = selection.getNumRows();
    if (startRow <= 1) {
      numRows -= (2 - startRow);
      startRow = 2;
    }
    if (numRows < 1) {
      ui.alert('Select a row', 'Select at least one data row (below the header row) to draw group borders on.', ui.ButtonSet.OK);
      return;
    }

    const saved = loadGroupBorderSettings_(sheetName);
    const defaultColumns = defaultGroupByColumns_(headers, selection.getColumn(), selection.getNumColumns(), saved.columns);

    const html = buildGroupBordersDialogHtml_(sheetName, startRow, numRows, headers, defaultColumns, saved.borderStyle, saved.color);
    ui.showModalDialog(html, 'Draw Group Borders');
  } catch (err) {
    console.error('showDrawGroupBordersDialog', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

function escapeHtmlGroupBorders_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * @param {string} sheetName
 * @param {number} startRow - 1-based, first data row the command will affect
 * @param {number} numRows
 * @param {string[]} headers - full header row (blanks as '')
 * @param {string[]} defaultColumns - header names to pre-check
 * @param {string} defaultStyle - a GROUP_BORDER_STYLES value
 * @param {string} defaultColor - '#rrggbb'
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function buildGroupBordersDialogHtml_(sheetName, startRow, numRows, headers, defaultColumns, defaultStyle, defaultColor) {
  const defaultSet = {};
  defaultColumns.forEach(function (h) { defaultSet[h] = true; });

  const checkboxesHtml = headers
    .filter(function (h) { return h; })
    .map(function (h) {
      const checked = defaultSet[h] ? ' checked' : '';
      const safe = escapeHtmlGroupBorders_(h);
      return '<label class="col"><input type="checkbox" value="' + safe + '"' + checked + '> ' + safe + '</label>';
    })
    .join('');

  const styleOptionsHtml = GROUP_BORDER_STYLES.map(function (s) {
    const selected = s.value === defaultStyle ? ' selected' : '';
    return '<option value="' + s.value + '"' + selected + '>' + s.label + '</option>';
  }).join('');

  const context = JSON.stringify({ sheetName: sheetName, startRow: startRow, numRows: numRows });

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<style>body{font-family:Google Sans,Arial,sans-serif;padding:16px;margin:0}' +
      'label.section{display:block;font-weight:500;margin-bottom:8px}' +
      '.cols{display:flex;flex-direction:column;gap:4px;max-height:200px;overflow-y:auto;border:1px solid #dadce0;border-radius:4px;padding:8px;margin-bottom:16px}' +
      '.col{font-size:13px;font-weight:normal}' +
      '.row{display:flex;gap:12px;align-items:center;margin-bottom:12px}' +
      '.row label{font-size:13px;color:#3c4043}' +
      'select,input[type=color]{padding:6px;border:1px solid #dadce0;border-radius:4px;font-size:14px}' +
      '.note{font-size:12px;color:#5f6368;margin-top:4px}' +
      '.buttons{display:flex;gap:10px;margin-top:20px;padding-top:16px;border-top:1px solid #e0e0e0}' +
      '.btn{flex:1;padding:10px 16px;border:none;border-radius:4px;font-size:14px;font-weight:500;cursor:pointer}' +
      '.btn-primary{background:#1a73e8;color:#fff}.btn-secondary{background:#f8f9fa;color:#3c4043;border:1px solid #dadce0}</style></head><body>' +
      '<label class="section">Group by (a border is drawn wherever any checked column changes)</label>' +
      '<div class="cols" id="cols">' + checkboxesHtml + '</div>' +
      '<div class="row"><label for="style">Style</label><select id="style">' + styleOptionsHtml + '</select>' +
      '<label for="color">Color</label><input type="color" id="color" value="' + escapeHtmlGroupBorders_(defaultColor) + '"></div>' +
      '<div class="note">Draws on rows ' + startRow + '–' + (startRow + numRows - 1) + ' of "' + escapeHtmlGroupBorders_(sheetName) +
      '", spanning every column. Clears this sheet’s previous group borders in that row range first.</div>' +
      '<div class="buttons">' +
      '<button class="btn btn-primary" onclick="runApply()">Draw Borders</button>' +
      '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
      '</div>' +
      '<script>' +
      'var GROUP_BORDER_CONTEXT = ' + context + ';' +
      'function runApply(){' +
      'var boxes = document.querySelectorAll("#cols input[type=checkbox]:checked");' +
      'var columns = []; for (var i = 0; i < boxes.length; i++) { columns.push(boxes[i].value); }' +
      'if (columns.length === 0) { alert("Choose at least one column to group by."); return; }' +
      'var payload = { sheetName: GROUP_BORDER_CONTEXT.sheetName, startRow: GROUP_BORDER_CONTEXT.startRow, ' +
      'numRows: GROUP_BORDER_CONTEXT.numRows, columns: columns, ' +
      'borderStyle: document.getElementById("style").value, color: document.getElementById("color").value };' +
      'google.script.run.withSuccessHandler(function(msg){ alert(msg); google.script.host.close(); })' +
      '.withFailureHandler(function(e){ alert(e.message || String(e)); })' +
      '.applyGroupBordersFromDialog(encodeURIComponent(JSON.stringify(payload)));' +
      '}' +
      '</script></body></html>'
  )
    .setWidth(420)
    .setHeight(420);
}

/**
 * Google Sheets border style enum, callable only inside Apps Script. Falls back to the default
 * style for anything unrecognized (a tampered or stale payload) rather than throwing.
 * @param {string} name - one of GROUP_BORDER_STYLES' values
 * @return {GoogleAppsScript.Spreadsheet.BorderStyle}
 */
function groupBorderStyleFromName_(name) {
  const known = {};
  GROUP_BORDER_STYLES.forEach(function (s) { known[s.value] = true; });
  const key = known[name] ? name : GROUP_BORDER_DEFAULT_STYLE;
  return SpreadsheetApp.BorderStyle[key];
}

/**
 * Draw (or redraw) group borders on rows [startRow, startRow + numRows) of sheet, grouped by the
 * given header names, spanning the sheet's full data width. Idempotent: clears every bottom
 * border already in that row range first, and only ever touches the bottom edge, so re-running
 * after rows shift never leaves stale borders and never disturbs any other formatting.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} startRow - 1-based, first data row
 * @param {number} numRows
 * @param {string[]} columnNames - header names to group by
 * @param {string} borderStyleName
 * @param {string} color
 * @return {{groupCount: number, rowCount: number}}
 */
function drawGroupBordersOnSheet_(sheet, startRow, numRows, columnNames, borderStyleName, color) {
  const numColumns = sheet.getLastColumn();
  if (numColumns < 1 || numRows < 1) return { groupCount: 0, rowCount: 0 };

  const headers = sheet.getRange(1, 1, 1, numColumns).getValues()[0].map(groupBorderCellText_);
  const colIndices = columnNames
    .map(function (name) { return headers.indexOf(name); })
    .filter(function (i) { return i >= 0; });
  if (colIndices.length === 0) {
    throw new Error('None of the chosen columns were found on "' + sheet.getName() + '".');
  }

  const rows = sheet.getRange(startRow, 1, numRows, numColumns).getValues();

  // Clear first: bottom border only, every row in range, before redrawing (see file header).
  for (let r = 0; r < numRows; r++) {
    sheet.getRange(startRow + r, 1, 1, numColumns).setBorder(null, null, false, null, null, null);
  }

  const boundaryRows = findGroupBoundaryRows_(rows, colIndices);
  const style = groupBorderStyleFromName_(borderStyleName);
  boundaryRows.forEach(function (i) {
    sheet.getRange(startRow + i, 1, 1, numColumns).setBorder(null, null, true, null, null, null, color, style);
  });

  return { groupCount: boundaryRows.length, rowCount: numRows };
}

/**
 * google.script.run entry point from the dialog.
 * @param {string} encodedPayload - encodeURIComponent(JSON.stringify({sheetName, startRow, numRows, columns, borderStyle, color}))
 * @return {string} summary message shown to the coach
 */
function applyGroupBordersFromDialog(encodedPayload) {
  let payload;
  try {
    payload = JSON.parse(decodeURIComponent(encodedPayload));
  } catch (e) {
    throw new Error('Invalid selection');
  }
  const sheetName = payload.sheetName;
  const columnNames = Array.isArray(payload.columns) ? payload.columns : [];
  const borderStyleName = payload.borderStyle || GROUP_BORDER_DEFAULT_STYLE;
  const color = payload.color || GROUP_BORDER_DEFAULT_COLOR;
  if (!sheetName) throw new Error('Missing sheet');
  if (columnNames.length === 0) throw new Error('Choose at least one column to group by.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  saveGroupBorderSettings_(sheetName, { columns: columnNames, borderStyle: borderStyleName, color: color });

  const result = drawGroupBordersOnSheet_(sheet, payload.startRow, payload.numRows, columnNames, borderStyleName, color);
  return 'Drew ' + result.groupCount + ' group border' + (result.groupCount === 1 ? '' : 's') +
    ' across ' + result.rowCount + ' row' + (result.rowCount === 1 ? '' : 's') + '.';
}
