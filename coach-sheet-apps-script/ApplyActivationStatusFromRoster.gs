/**
 * Copy Activation Status from a prep sheet (e.g. "3/28 Game Roster Prep") into Game Availability
 * for the same dated column header (e.g. "3/28 Activation Status").
 */

const GAME_ACTIVATION_STATUS_HEADER_SUFFIX = ' Activation Status';

/**
 * @param {string} headerStr
 * @return {string|null} Normalized "M/D" or null
 */
function parseGameActivationStatusHeader(headerStr) {
  if (!headerStr || typeof headerStr !== 'string') return null;
  const m = headerStr.trim().match(/^(\d{1,2})\/(\d{1,2}) Activation Status(?: \(Game \d+\))?$/);
  if (!m) return null;
  return String(parseInt(m[1], 10)) + '/' + String(parseInt(m[2], 10));
}

/** Header cell to comparable string (handles Date-typed headers in row 1). */
function headerText_(cell) {
  if (cell == null || cell === '') return '';
  if (cell instanceof Date) return formatDateForColumn(cell) + GAME_ACTIVATION_STATUS_HEADER_SUFFIX;
  return String(cell).trim();
}

/** 1-based column index, or -1 */
function findHeaderColumn1Based_(sheet, headerText) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headerText_(headers[i]) === headerText) return i + 1;
  }
  return -1;
}

function strCell_(v) {
  if (v == null || v === '') return '';
  return String(v).trim();
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<{ sheetName: string, activationHeader: string }>}
 */
function getSheetsWithGameActivationStatusColumn(ss) {
  const out = [];
  const sheets = ss.getSheets();
  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const name = sh.getName();
    if (name === PRACTICE_AVAILABILITY_CONFIG.availabilitySheet) continue;
    if (name === GAME_AVAILABILITY_CONFIG.availabilitySheet) continue;
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) continue;
    const row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const seen = {};
    for (let c = 0; c < row.length; c++) {
      const h = headerText_(row[c]);
      if (parseGameActivationStatusHeader(h) && seen[h] !== true) {
        seen[h] = true;
        out.push({ sheetName: name, activationHeader: h });
      }
    }
  }
  out.sort(function (a, b) {
    if (a.sheetName !== b.sheetName) return a.sheetName.localeCompare(b.sheetName);
    return a.activationHeader.localeCompare(b.activationHeader);
  });
  return out;
}

function showApplyActivationStatusDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const candidates = getSheetsWithGameActivationStatusColumn(ss);
    if (candidates.length === 0) {
      ui.alert(
        'No matching sheets',
        'No sheets found with a column like "3/28 Activation Status".',
        ui.ButtonSet.OK
      );
      return;
    }

    const optionsHtml = candidates.map(function (item) {
      const payload = encodeURIComponent(JSON.stringify({
        sheetName: item.sheetName,
        activationHeader: item.activationHeader
      }));
      return (
        '<option value="' +
        payload +
        '">' +
        escapeHtmlApplyActivation(item.sheetName + ' — ' + item.activationHeader) +
        '</option>'
      );
    }).join('');

    const html = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><head><meta charset="utf-8">' +
        '<style>body{font-family:Google Sans,Arial,sans-serif;padding:16px;margin:0}label{display:block;font-weight:500;margin-bottom:8px}' +
        'select{width:100%;padding:10px;border:1px solid #dadce0;border-radius:4px;font-size:14px;box-sizing:border-box}' +
        '.note{font-size:12px;color:#5f6368;margin-top:8px}.buttons{display:flex;gap:10px;margin-top:20px;padding-top:16px;border-top:1px solid #e0e0e0}' +
        '.btn{flex:1;padding:10px 16px;border:none;border-radius:4px;font-size:14px;font-weight:500;cursor:pointer}' +
        '.btn-primary{background:#1a73e8;color:#fff}.btn-secondary{background:#f8f9fa;color:#3c4043;border:1px solid #dadce0}</style></head><body>' +
        '<label for="src">Source sheet and Activation Status column</label>' +
        '<select id="src">' +
        optionsHtml +
        '</select>' +
        '<div class="note">Each Full Name on the source sheet updates that player’s cell in Game Availability for this column (Active, Inactive, or TBD only).</div>' +
        '<div class="buttons">' +
        '<button class="btn btn-primary" onclick="runApply()">Apply</button>' +
        '<button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
        '</div>' +
        '<script>' +
        'function runApply(){var sel=document.getElementById("src");var v=sel.value;if(!v)return;' +
        'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
        '.withFailureHandler(function(e){alert(e.message||String(e));})' +
        '.applyActivationStatusFromRosterPrepSelection(v);}' +
        '</script></body></html>'
    )
      .setWidth(480)
      .setHeight(280);

    ui.showModalDialog(html, 'Apply Activation Status');
  } catch (err) {
    console.error('showApplyActivationStatusDialog', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

function escapeHtmlApplyActivation(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function applyActivationStatusFromRosterPrepSelection(encodedPayload) {
  const ui = SpreadsheetApp.getUi();
  let payload;
  try {
    payload = JSON.parse(decodeURIComponent(encodedPayload));
  } catch (e) {
    throw new Error('Invalid selection');
  }
  const sheetName = payload.sheetName;
  const activationHeader = headerText_(payload.activationHeader);
  if (!sheetName || !activationHeader) throw new Error('Missing sheet or column');

  const result = applyActivationStatusFromRosterSheet(sheetName, activationHeader);
  ui.alert(
    'Done',
    'Updated ' +
      result.updated +
      ' cell(s) in Game Availability for "' +
      activationHeader +
      '".\n\n' +
      'Skipped rows: ' +
      result.skipped,
    ui.ButtonSet.OK
  );
}

/**
 * @param {string} sourceSheetName
 * @param {string} activationHeader e.g. "3/28 Activation Status"
 * @return {{ updated: number, skipped: number }}
 */
function applyActivationStatusFromRosterSheet(sourceSheetName, activationHeader) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  if (!parseGameActivationStatusHeader(activationHeader)) {
    throw new Error('Not a valid activation header: ' + activationHeader);
  }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) throw new Error('Source sheet not found: ' + sourceSheetName);

  const gaSheet = ss.getSheetByName(GAME_AVAILABILITY_CONFIG.availabilitySheet);
  if (!gaSheet) throw new Error('Game Availability sheet not found');

  const gaCol = findHeaderColumn1Based_(gaSheet, activationHeader);
  if (gaCol < 0) {
    throw new Error('Column "' + activationHeader + '" not found in Game Availability.');
  }

  const lastColSrc = sourceSheet.getLastColumn();
  const srcHeaders = sourceSheet.getRange(1, 1, 1, lastColSrc).getValues()[0];
  let colName = -1;
  let colAct = -1;
  for (let i = 0; i < srcHeaders.length; i++) {
    const h = strCell_(srcHeaders[i]);
    if (h === CONFIG.columns.fullName) colName = i + 1;
    if (headerText_(srcHeaders[i]) === activationHeader) colAct = i + 1;
  }
  if (colName < 0) throw new Error('Source sheet needs a "' + CONFIG.columns.fullName + '" column.');
  if (colAct < 0) throw new Error('Source sheet needs column "' + activationHeader + '".');

  const allowed = {};
  GAME_ACTIVATION_STATUS_OPTIONS.forEach(function (o) {
    allowed[o.value] = true;
  });

  const lastRowSrc = sourceSheet.getLastRow();
  if (lastRowSrc < 2) return { updated: 0, skipped: 0 };

  const names = sourceSheet.getRange(2, colName, lastRowSrc, colName).getValues();
  const acts = sourceSheet.getRange(2, colAct, lastRowSrc, colAct).getValues();

  const gaLast = gaSheet.getLastRow();
  if (gaLast < 2) throw new Error('Game Availability has no data rows.');

  const gaNames = gaSheet.getRange(2, 1, gaLast, 1).getValues();
  const rowByName = {};
  for (let i = 0; i < gaNames.length; i++) {
    const nm = strCell_(gaNames[i][0]);
    if (nm) rowByName[nm.toLowerCase()] = i;
  }

  const outCol = gaSheet.getRange(2, gaCol, gaLast, gaCol).getValues();

  let updated = 0;
  let skipped = 0;

  for (let r = 0; r < names.length; r++) {
    const name = strCell_(names[r][0]);
    const status = strCell_(acts[r][0]);
    if (!name || !status || !allowed[status]) {
      skipped++;
      continue;
    }
    const ix = rowByName[name.toLowerCase()];
    if (ix === undefined) {
      skipped++;
      continue;
    }
    if (outCol[ix][0] !== status) updated++;
    outCol[ix][0] = status;
  }

  gaSheet.getRange(2, gaCol, gaLast, gaCol).setValues(outCol);

  return { updated: updated, skipped: skipped };
}
