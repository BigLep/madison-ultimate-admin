/**
 * Setup diagnostics: confirms this spreadsheet/script is configured correctly.
 * Run this after deploying to a new season's sheet, or any time something is
 * misbehaving, before digging further.
 */
function runDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  function checkSheet(label, sheetName, required) {
    const exists = !!ss.getSheetByName(sheetName);
    results.push({
      pass: exists,
      required: required !== false,
      line: exists
        ? `✅ Sheet "${sheetName}" (${label}) exists`
        : `${required === false ? '⚠️' : '❌'} Sheet "${sheetName}" (${label}) NOT found`
    });
  }

  function checkFolder(label, folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      results.push({ pass: true, required: true, line: `✅ ${label} folder accessible: "${folder.getName()}" (${folderId})` });
    } catch (e) {
      results.push({ pass: false, required: true, line: `❌ ${label} folder NOT accessible (${folderId}): ${e.message}` });
    }
  }

  function checkButtondownReadAccess() {
    const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.buttondown.apiKeyProperty);
    if (!apiKey) {
      results.push({ pass: false, required: true, line: `❌ Script property "${CONFIG.buttondown.apiKeyProperty}" NOT set (needed to read Buttondown subscribers)` });
      return;
    }
    try {
      const response = UrlFetchApp.fetch(`${CONFIG.buttondown.apiBase}/subscribers`, {
        headers: { Authorization: `Token ${apiKey}` },
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      if (code === 200) {
        results.push({ pass: true, required: true, line: `✅ Buttondown API key (property "${CONFIG.buttondown.apiKeyProperty}") can read subscribers` });
      } else if (code === 401 || code === 403) {
        results.push({ pass: false, required: true, line: `❌ Buttondown API key rejected (HTTP ${code}) - check the "${CONFIG.buttondown.apiKeyProperty}" script property value` });
      } else {
        results.push({ pass: false, required: true, line: `❌ Buttondown API returned HTTP ${code} when listing subscribers` });
      }
    } catch (e) {
      results.push({ pass: false, required: true, line: `❌ Buttondown API request failed: ${e.message}` });
    }
  }

  function checkSignupsMirror() {
    const sheet = ss.getSheetByName(CONFIG.signups.sheetName);
    if (!sheet) {
      results.push({ pass: false, required: true, line: `❌ Sheet "${CONFIG.signups.sheetName}" (Signups mirror) NOT found; add a tab with that name whose A1 is an IMPORTRANGE of the portal's Signups sheet` });
      return;
    }
    const a1Formula = sheet.getRange('A1').getFormula() || '';
    if (!/^=IMPORTRANGE\(/i.test(a1Formula.trim())) {
      results.push({ pass: false, required: true, line: `❌ Sheet "${CONFIG.signups.sheetName}" A1 is not an IMPORTRANGE formula (found: ${a1Formula || 'a plain value'})` });
    } else {
      const a2 = sheet.getRange('A2');
      const a2Value = a2.getValue();
      const a2Text = a2Value === null || a2Value === undefined ? '' : a2Value.toString();
      const errored = /^#(REF|N\/A|ERROR|VALUE|NAME)/.test(a2Text) || a2Text === '';
      results.push({
        pass: !errored,
        required: true,
        line: errored
          ? `❌ Sheet "${CONFIG.signups.sheetName}" IMPORTRANGE has not resolved (A2 is "${a2Text}"); open the tab and click "Allow access" if prompted`
          : `✅ Sheet "${CONFIG.signups.sheetName}" (Signups mirror) IMPORTRANGE resolved`
      });
    }
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    try {
      resolveSignupsColumns(headers);
      results.push({ pass: true, required: true, line: `✅ Sheet "${CONFIG.signups.sheetName}" has every header the Roster formulas reference (${Object.keys(SIGNUPS_HEADERS).length})` });
    } catch (e) {
      results.push({ pass: false, required: true, line: `❌ ${e.message}` });
    }
  }

  function checkRosterHeaderAndKey() {
    const sheet = ss.getSheetByName(CONFIG.roster.sheetName);
    if (!sheet) return; // reported by checkSheet above
    const headers = sheet.getRange(ROSTER_HEADER_ROW, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
      .map(h => h === null || h === undefined ? '' : h.toString().trim());
    const present = new Set(headers);
    const missingDefined = ROSTER_COLUMNS.map(c => c.name).filter(name => !present.has(name));
    const missingConfig = Object.values(CONFIG.columns).filter(name => !present.has(name));
    const missing = Array.from(new Set(missingDefined.concat(missingConfig)));
    results.push({
      pass: missing.length === 0,
      required: true,
      line: missing.length === 0
        ? `✅ Roster header row has every ROSTER_COLUMNS name and every CONFIG.columns value (${ROSTER_COLUMNS.length} columns)`
        : `❌ Roster header row is missing: ${missing.join(', ')}; run "Generate Fresh Roster"`
    });
    // Column A holds PlayerIDs as plain values and every other column a per-row
    // formula (ADR 0003), so the Roster is stale whenever Signups gained or lost a
    // Player since Generate Fresh Roster last ran.
    const lastRow = sheet.getLastRow();
    const rosterIds = lastRow < ROSTER_FIRST_DATA_ROW ? [] :
      sheet.getRange(ROSTER_FIRST_DATA_ROW, 1, lastRow - ROSTER_HEADER_ROW, 1).getValues()
        .map(r => (r[0] === null || r[0] === undefined ? '' : r[0].toString().trim()))
        .filter(id => id !== '');
    const firstFormula = lastRow < ROSTER_FIRST_DATA_ROW ? '' : (sheet.getRange(ROSTER_FIRST_DATA_ROW, 2).getFormula() || '');
    const shapeOk = rosterIds.length === 0 || firstFormula.startsWith('=IF($A' + ROSTER_FIRST_DATA_ROW + '=""');
    results.push({
      pass: shapeOk,
      required: true,
      line: shapeOk
        ? `✅ Roster rows are PlayerID values in column A plus per-row formulas (${rosterIds.length} Players)`
        : `❌ Roster B${ROSTER_FIRST_DATA_ROW} is not a per-row Roster formula (found: "${firstFormula || 'a plain value'}"); the sheet was edited or generated by an older version; run "Generate Fresh Roster"`
    });
    const signupsSheet = ss.getSheetByName(CONFIG.signups.sheetName);
    if (signupsSheet && signupsSheet.getLastRow() >= 1) {
      const signupsHeaders = signupsSheet.getRange(1, 1, 1, Math.max(1, signupsSheet.getLastColumn())).getValues()[0]
        .map(h => h === null || h === undefined ? '' : h.toString().trim());
      const idCol = signupsHeaders.indexOf(SIGNUPS_HEADERS.playerId);
      if (idCol >= 0) {
        const signupsIds = signupsSheet.getLastRow() < 2 ? [] :
          signupsSheet.getRange(2, idCol + 1, signupsSheet.getLastRow() - 1, 1).getValues()
            .map(r => (r[0] === null || r[0] === undefined ? '' : r[0].toString().trim()))
            .filter(id => id !== '');
        const rosterSet = new Set(rosterIds);
        const signupsSet = new Set(signupsIds);
        const missing = signupsIds.filter(id => !rosterSet.has(id));
        const extra = rosterIds.filter(id => !signupsSet.has(id));
        const fresh = missing.length === 0 && extra.length === 0;
        results.push({
          pass: fresh,
          required: true,
          line: fresh
            ? `✅ Roster is current: its ${rosterIds.length} PlayerIDs match Signups`
            : `❌ Roster is stale: ${missing.length} Signups PlayerID(s) not on the Roster, ${extra.length} Roster row(s) no longer in Signups; run "Generate Fresh Roster"`
        });
      }
    }
  }

  // Sheets the script reads from or writes into by exact name.
  checkSheet('Roster', CONFIG.roster.sheetName);
  checkRosterHeaderAndKey();
  checkSignupsMirror();
  checkSheet('Final Forms import target', CONFIG.finalForms.sheetName);
  checkSheet('Newsletter Subscribers import target', CONFIG.newsletterSubscribers.sheetName);
  results.push(Object.assign({ required: true }, checkExtraPlayerInfoSheet(ss)));
  checkSheet('Practice Info', CONFIG.practiceInfo.sheetName);
  checkSheet('Game Info', CONFIG.gameInfo.sheetName);
  checkSheet('Fields', CONFIG.fieldsSheet.sheetName);
  checkSheet('Practice Availability', CONFIG.practiceAvailability.sheetName);
  checkSheet('Game Availability', CONFIG.gameAvailability.sheetName);
  // Only written by "Convert to Actual Attendance"; that function already warns
  // gracefully if missing, so treat it as a warning here rather than a failure.
  checkSheet('Attendance (used by Convert to Actual Attendance)', CONFIG.attendance.sheetName, false);

  // Drive folders the "Update ..." importers read the newest CSV from.
  checkFolder('Final Forms exports', CONFIG.finalForms.folderId);

  // Script properties / external APIs (secrets set via Project Settings > Script
  // Properties, not committed to git). Add a check here whenever a new integration
  // starts reading one, so this stays the single "is everything set up" check.
  checkButtondownReadAccess();

  const failures = results.filter(r => !r.pass && r.required);
  const warnings = results.filter(r => !r.pass && !r.required);
  const passed = results.length - failures.length - warnings.length;

  const summaryLine = failures.length
    ? `❌ ${failures.length} check(s) FAILED, ${passed}/${results.length} passed` + (warnings.length ? `, ${warnings.length} warning(s)` : '')
    : warnings.length
      ? `✅ All required checks passed (${warnings.length} warning(s))`
      : `✅ All ${results.length} checks passed`;

  const message = `Script version: ${SCRIPT_VERSION}\n\n${summaryLine}\n\n${results.map(r => r.line).join('\n')}`;

  console.log(message);
  SpreadsheetApp.getUi().alert(
    failures.length ? '🩺 Diagnostics: issues found' : '🩺 Diagnostics: all good',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
