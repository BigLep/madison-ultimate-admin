/**
 * Buttondown API calls: creating a Draft and uploading an image.
 * https://docs.buttondown.com/api-emails-create, https://docs.buttondown.com/api-images-create
 */

/**
 * Create a Buttondown Draft from Markdown. Returns the parsed API response
 * (includes "id" when the request succeeds).
 */
function createButtondownDraft(apiKey, subject, bodyMarkdown) {
  // Without this, Buttondown auto-detects the body's format and switches to raw-HTML
  // ("fancy") interpretation as soon as it spots anything resembling an HTML tag
  // (which our own COPY_PASTE_IN_TABLE-style placeholders no longer do, but coach-
  // authored content might). Under fancy mode our Markdown syntax doesn't get
  // parsed at all: "#", "**", "_" show up as literal escaped characters instead of
  // headings/bold/italic. This must be the literal first line of the body.
  const modeComment = '<!-- buttondown-editor-mode: plaintext -->';
  // Buttondown also rejects a body starting with "---" as YAML frontmatter unless
  // told otherwise; moot now that the mode comment is always the real first line,
  // kept as defense in depth in case that check looks past leading comment lines.
  const safeMarkdown = bodyMarkdown.startsWith('---') ? `​\n${bodyMarkdown}` : bodyMarkdown;
  const body = `${modeComment}\n${safeMarkdown}`;

  const payload = { subject, body, status: 'draft' };
  const seasonTag = CONFIG.buttondown.currentSeasonTag;
  if (seasonTag && seasonTag.id) {
    // Defaults every draft's audience to the current season's roster tag (tag
    // membership is managed elsewhere, outside this script). Buttondown's filters
    // key subscribers by tag id, not name; only evaluated when the draft is sent,
    // so this is safe to set on a draft, not just a scheduled/sent email.
    payload.filters = {
      predicate: 'and',
      filters: [{ field: 'subscriber.tags', operator: 'contains', value: seasonTag.id }],
      groups: [] // required by the API even when empty (HTTP 422 without it)
    };
  }

  const response = UrlFetchApp.fetch(`${CONFIG.buttondown.apiBase}/emails`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Token ${apiKey}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code === 401 || code === 403) {
    throw new Error(`Buttondown API rejected the request (HTTP ${code}). Check the "${CONFIG.buttondown.apiKeyProperty}" script property value and that it has write access.`);
  }
  if (code < 200 || code >= 300) {
    throw new Error(`Buttondown API error creating draft: HTTP ${code} - ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}

/**
 * Upload one image to Buttondown, returning the parsed API response (includes
 * "image", the hosted URL, when the request succeeds).
 */
function uploadButtondownImage(apiKey, blob) {
  const response = UrlFetchApp.fetch(`${CONFIG.buttondown.apiBase}/images`, {
    method: 'post',
    headers: { Authorization: `Token ${apiKey}` },
    payload: { image: blob },
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Buttondown API error uploading image: HTTP ${code} - ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText());
}
