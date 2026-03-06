/**
 * Export Game Info sheet to a bulletted Markdown list.
 * Shows result in a dialog with Copy to clipboard button.
 * Uses Fields sheet for Google Map URL by field name.
 */

/**
 * Menu entry point: build markdown from Game Info and show in dialog.
 */
function exportGameInfoToMarkdown() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const infoSheet = ss.getSheetByName(CONFIG.gameInfo.sheetName);
    if (!infoSheet) {
      ui.alert('Sheet not found', '"' + CONFIG.gameInfo.sheetName + '" sheet was not found.', ui.ButtonSet.OK);
      return;
    }
    const markdown = getGameInfoMarkdown(ss, infoSheet);
    if (!markdown || markdown.trim() === '') {
      ui.alert('No game rows', 'No valid game rows found (check Date column and skip "Bye" rows).', ui.ButtonSet.OK);
      return;
    }
    showMarkdownDialog(markdown);
  } catch (err) {
    console.error('exportGameInfoToMarkdown', err);
    ui.alert('Error', err.message || String(err), ui.ButtonSet.OK);
  }
}

/**
 * Build markdown list from Game Info sheet using getDatesFromInfoSheet and Fields lookup.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {SpreadsheetApp.Sheet} infoSheet
 * @return {string}
 */
function getGameInfoMarkdown(ss, infoSheet) {
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
  const colFieldName = col(GAME_INFO_COLUMNS.FIELD_NAME);
  const colFieldLocation = col(GAME_INFO_COLUMNS.FIELD_LOCATION);
  const colGameNote = col(GAME_INFO_COLUMNS.GAME_NOTE);
  const colOpponent = col(GAME_INFO_COLUMNS.OPPONENT);
  const colOpponentTeamPage = col(GAME_INFO_COLUMNS.OPPONENT_TEAM_PAGE);

  if (colDate === -1) return '';
  const lastRow = infoSheet.getLastRow();
  const lastCol = Math.max(
    colDate + 1,
    colWarmupArrival >= 0 ? colWarmupArrival + 1 : 0,
    colGameStart >= 0 ? colGameStart + 1 : 0,
    colFieldName >= 0 ? colFieldName + 1 : 0,
    colFieldLocation >= 0 ? colFieldLocation + 1 : 0,
    colGameNote >= 0 ? colGameNote + 1 : 0,
    colOpponent >= 0 ? colOpponent + 1 : 0,
    colOpponentTeamPage >= 0 ? colOpponentTeamPage + 1 : 0
  );
  const dataRange = infoSheet.getRange(2, 1, lastRow, lastCol);
  const data = dataRange.getValues();

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const blocks = [];
  dateInfos.forEach(function (dateInfo) {
    const dataIndex = dateInfo.rowIndex - 2;
    if (dataIndex < 0 || dataIndex >= data.length) return;
    const row = data[dataIndex];
    const dateObj = dateInfo.date;

    const dateLabel = (dateObj.getMonth() + 1) + '/' + dateObj.getDate() + ' ' + dayNames[dateObj.getDay()];
    const gameStartVal = colGameStart >= 0 ? row[colGameStart] : null;
    const gameStartStr = gameStartVal != null ? String(gameStartVal).trim() : '';
    const isTBD = gameStartStr === '' || gameStartStr.toLowerCase() === 'tbd';

    const fieldName = colFieldName >= 0 && row[colFieldName] ? String(row[colFieldName]).trim() : '';
    const fieldLocation = colFieldLocation >= 0 && row[colFieldLocation] ? String(row[colFieldLocation]).trim() : '';
    const gameNote = colGameNote >= 0 && row[colGameNote] ? String(row[colGameNote]).trim() : '';
    const opponentStr = colOpponent >= 0 && row[colOpponent] ? String(row[colOpponent]).trim() : '';
    const opponentTeamPageUrl = colOpponentTeamPage >= 0 && row[colOpponentTeamPage] ? String(row[colOpponentTeamPage]).trim() : '';
    const fieldInfo = fieldName ? fieldsLookup[fieldName.toLowerCase()] : null;
    const googleMapUrl = fieldInfo ? fieldInfo.googleMapUrl : '';

    const locationDisplay = fieldName + (fieldLocation ? ' (' + fieldLocation + ')' : '');
    const locationLine = googleMapUrl
      ? '  * Location: [' + locationDisplay + '](' + googleMapUrl + ')'
      : (locationDisplay ? '  * Location: ' + locationDisplay : '');

    const lines = ['* ' + dateLabel];
    if (!isTBD && colWarmupArrival >= 0 && row[colWarmupArrival]) {
      const warmupTime = formatTimeForDisplay(parseTimeOnDate(dateObj, row[colWarmupArrival]));
      if (warmupTime) lines.push('  * Warmup: ' + warmupTime);
    }
    if (!isTBD && gameStartVal) {
      const gameStartTime = parseTimeOnDate(dateObj, gameStartVal);
      if (gameStartTime) lines.push('  * Game start: ' + formatTimeForDisplay(gameStartTime));
    }
    if (locationLine) lines.push(locationLine);
    if (opponentStr || opponentTeamPageUrl) {
      const opponentLine = opponentTeamPageUrl
        ? '  * Opponent: [' + (opponentStr || opponentTeamPageUrl) + '](' + opponentTeamPageUrl + ')'
        : '  * Opponent: ' + opponentStr;
      lines.push(opponentLine);
    }
    if (gameNote) lines.push('  * Note: ' + gameNote);
    blocks.push(lines.join('\n'));
  });
  return blocks.join('\n');
}

/**
 * Format a Date (time part) for display, e.g. "10:00 AM".
 * @param {Date} date
 * @return {string}
 */
function formatTimeForDisplay(date) {
  if (!date || isNaN(date.getTime())) return '';
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const ampm = hours >= 12 ? 'PM' : 'AM';
  const h = hours % 12 || 12;
  const m = minutes < 10 ? '0' + minutes : String(minutes);
  return h + ':' + m + ' ' + ampm;
}

/**
 * Escape HTML for safe embedding in dialog.
 * @param {string} s
 * @return {string}
 */
function escapeHtmlForDialog(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Show modal dialog with markdown and Copy to clipboard button.
 * @param {string} markdown
 */
function showMarkdownDialog(markdown) {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"></head><body>' +
    '<style>' +
    'body { font-family: Arial, sans-serif; padding: 12px; }' +
    'h3 { margin-top: 0; color: #1a73e8; }' +
    'textarea { width: 100%; height: 280px; font-family: monospace; font-size: 13px; padding: 8px; box-sizing: border-box; }' +
    'button { margin-top: 8px; padding: 8px 16px; background: #1a73e8; color: #fff; border: none; border-radius: 4px; cursor: pointer; }' +
    'button:hover { background: #1557b0; }' +
    '.hint { color: #5f6368; font-size: 12px; margin-top: 6px; }' +
    '</style>' +
    '<h3>Export Game Info to Markdown List</h3>' +
    '<div id="mdSource" style="display:none">' + escapeHtmlForDialog(markdown) + '</div>' +
    '<textarea id="md" readonly placeholder="Markdown will appear here..."></textarea>' +
    '<div><button type="button" id="copyBtn">Copy to clipboard</button></div>' +
    '<div class="hint" id="hint"></div>' +
    '<script>' +
    '(function() {' +
    'var src = document.getElementById("mdSource");' +
    'var ta = document.getElementById("md");' +
    'var btn = document.getElementById("copyBtn");' +
    'var hint = document.getElementById("hint");' +
    'if (src && ta) ta.value = src.textContent;' +
    'function copy() {' +
    '  ta.select();' +
    '  try { document.execCommand("copy"); hint.textContent = "Copied to clipboard."; }' +
    '  catch (e) { hint.textContent = "Copy failed. Select the text above and copy manually."; }' +
    '  setTimeout(function() { hint.textContent = ""; }, 2000);' +
    '}' +
    'if (btn) btn.addEventListener("click", copy);' +
    '})();' +
    '</script></body></html>'
  )
    .setWidth(520)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Export Game Info to Markdown List');
}
