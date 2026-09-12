/**
 * Buttondown API calls: creating a Draft and uploading an image.
 * https://docs.buttondown.com/api-emails-create, https://docs.buttondown.com/api-images-create
 */

/**
 * Create a Buttondown Draft from Markdown. Returns the parsed API response
 * (includes "id" when the request succeeds).
 */
function createButtondownDraft(apiKey, subject, bodyMarkdown) {
  // Buttondown rejects a body starting with "---" as YAML frontmatter unless told
  // otherwise. A Newsletter Block body would only start that way by coincidence
  // (e.g. opening with a horizontal rule), so guard against it rather than fail.
  const safeBody = bodyMarkdown.startsWith('---') ? `​\n${bodyMarkdown}` : bodyMarkdown;

  const response = UrlFetchApp.fetch(`${CONFIG.buttondown.apiBase}/emails`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Token ${apiKey}` },
    payload: JSON.stringify({ subject, body: safeBody, status: 'draft' }),
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
