/**
 * Update the Newsletter Subscribers sheet from the Buttondown API.
 *
 * Replaces the old Google Groups CSV mailing-list import (Google Groups was
 * retired as the mailing-list system of record in spring 2026; see
 * madison-ultimate/docs/fall-2026/signup-grill.md). This pulls the full
 * subscriber list (email + status) directly from Buttondown.
 *
 * Requires the "BUTTONDOWN_API_KEY" script property (Extensions > Apps Script >
 * Project Settings (gear) > Script Properties). Get the key from Buttondown:
 * Settings > Programming > API Keys — read access to subscribers is enough.
 */
function updateNewsletterSubscribers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CONFIG.newsletterSubscribers.sheetName);

  if (!sheet) {
    const existingSheetNames = ss.getSheets().map(s => s.getName()).join(', ');
    ui.alert('Error',
      `Sheet "${CONFIG.newsletterSubscribers.sheetName}" not found.\n\n` +
      `Fix: add a tab named exactly "${CONFIG.newsletterSubscribers.sheetName}" (blank is fine, this will populate it), then run this again.\n\n` +
      `Existing tabs: ${existingSheetNames}`,
      ui.ButtonSet.OK);
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.buttondown.apiKeyProperty);
  if (!apiKey) {
    ui.alert('Error',
      `Script property "${CONFIG.buttondown.apiKeyProperty}" is not set.\n\n` +
      `Fix: Extensions > Apps Script > Project Settings (gear) > Script Properties > Add script property.\n` +
      `Get the key from Buttondown: Settings > Programming > API Keys (read access to subscribers is enough).`,
      ui.ButtonSet.OK);
    return;
  }

  try {
    const subscribers = fetchAllButtondownSubscribers(apiKey);

    // "email_address" and "type" (subscriber status) are the fields we know are
    // reliable. Append any other scalar fields Buttondown returns so nothing is
    // silently dropped, without hardcoding field names we haven't verified.
    const baseFields = ['email_address', 'type'];
    const extraFieldSet = new Set();
    subscribers.forEach(sub => {
      Object.keys(sub).forEach(key => {
        const value = sub[key];
        if (!baseFields.includes(key) && (value === null || typeof value !== 'object')) {
          extraFieldSet.add(key);
        }
      });
    });
    const extraFields = Array.from(extraFieldSet).sort();
    const fields = baseFields.concat(extraFields);
    const header = ['Email', 'Status'].concat(extraFields);

    const rows = subscribers.map(sub =>
      fields.map(field => (sub[field] === undefined || sub[field] === null) ? '' : sub[field])
    );

    sheet.clear();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    }

    console.log(`✅ Updated Newsletter Subscribers: ${rows.length} subscribers`);
    ui.alert('Newsletter Subscribers Updated',
      `Successfully imported ${rows.length} subscribers from Buttondown.`,
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', `Could not update Newsletter Subscribers:\n${e.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * Fetch every subscriber from the Buttondown API, following pagination.
 * @param {string} apiKey
 * @returns {Array<Object>} raw subscriber objects as returned by the API
 */
function fetchAllButtondownSubscribers(apiKey) {
  const subscribers = [];
  let url = `${CONFIG.buttondown.apiBase}/subscribers`;
  let pageCount = 0;
  const maxPages = 50; // safety cap against an unexpected pagination loop

  while (url && pageCount < maxPages) {
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Token ${apiKey}` },
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();

    if (code === 401 || code === 403) {
      throw new Error(`Buttondown API rejected the request (HTTP ${code}). Check the "${CONFIG.buttondown.apiKeyProperty}" script property value.`);
    }
    if (code !== 200) {
      throw new Error(`Buttondown API error: HTTP ${code}`);
    }

    const data = JSON.parse(response.getContentText());
    (data.results || []).forEach(sub => subscribers.push(sub));
    url = data.next || null;
    pageCount++;
  }

  return subscribers;
}
