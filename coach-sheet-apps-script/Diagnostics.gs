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

  // Sheets the script reads from or writes into by exact name.
  checkSheet('Roster', CONFIG.roster.sheetName);
  checkSheet('Final Forms import target', CONFIG.finalForms.sheetName);
  checkSheet('Newsletter Subscribers import target', CONFIG.newsletterSubscribers.sheetName);
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
