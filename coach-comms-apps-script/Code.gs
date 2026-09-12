/**
 * Madison Ultimate coach comms: converts a Newsletter Block in the season's
 * Communications Doc into a Buttondown Draft.
 *
 * See CONTEXT.md for the Communications Doc / Newsletter Block / Buttondown Draft
 * vocabulary. Menu action operates on whichever Newsletter Block the cursor is
 * currently inside; see NewsletterBlock.gs for how that block is located, validated,
 * and converted to Markdown, and ButtondownApi.gs for the Buttondown API calls.
 */

// Script Version - Increment this number when making changes
const SCRIPT_VERSION = '1.4';

const CONFIG = {
  buttondown: {
    apiBase: 'https://api.buttondown.com/v1',
    // Script property (Extensions > Apps Script > Project Settings > Script Properties),
    // not committed here. Can reuse the same Buttondown API key already set for
    // coach-sheet-apps-script's Newsletter Subscribers sync (drafting requires the
    // same account, though a scoped key with write access is preferable if available).
    apiKeyProperty: 'BUTTONDOWN_API_KEY'
  },
  newsletterBlock: {
    // The Insert > Building Blocks > Email Draft address that marks a block as
    // meant for Buttondown, rather than an ordinary drafted-but-unaddressed email.
    toAddress: 'drafts@mg.buttondown.email',
    rowCount: 5,
    toRow: 0,
    subjectRow: 3,
    bodyRow: 4
  }
};

/**
 * Simple trigger: adds the menu whenever the Communications Doc is opened.
 */
function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  DocumentApp.getUi()
    .createMenu(`🥏 Madison Ultimate (v${SCRIPT_VERSION})`)
    .addItem('📬 Send Newsletter Block to Buttondown', 'sendNewsletterBlockToButtondown')
    .addToUi();
}

/**
 * Convert whichever Newsletter Block the cursor is in to a Buttondown Draft.
 * Click anywhere inside the Insert > Building Blocks > Email Draft block addressed
 * to drafts@mg.buttondown.email, then run this from the menu.
 */
function sendNewsletterBlockToButtondown() {
  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi();

  const table = findEnclosingTable(doc);
  if (!table) {
    ui.alert('No Newsletter Block Found',
      'Click inside a Newsletter Block (Insert > Building Blocks > Email Draft) before running this, then try again.',
      ui.ButtonSet.OK);
    return;
  }

  const cfg = CONFIG.newsletterBlock;
  if (table.getNumRows() !== cfg.rowCount) {
    ui.alert('Not a Newsletter Block',
      'The table your cursor is in doesn\'t look like an Email Draft building block. Click inside the To/Cc/Bcc/Subject/body block you want to send and try again.',
      ui.ButtonSet.OK);
    return;
  }

  const toEmail = findPersonEmailInRow(table.getRow(cfg.toRow));
  if (toEmail !== cfg.toAddress) {
    ui.alert('Not a Newsletter Block',
      `This email block's To field isn't "${cfg.toAddress}" (found: ${toEmail || '(none)'}). ` +
      'Point your cursor at the block addressed to Buttondown and try again.',
      ui.ButtonSet.OK);
    return;
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.buttondown.apiKeyProperty);
  if (!apiKey) {
    ui.alert('Error',
      `Script property "${CONFIG.buttondown.apiKeyProperty}" is not set.\n\n` +
      'Fix: Extensions > Apps Script > Project Settings (gear) > Script Properties > Add script property.\n' +
      'Get the key from Buttondown: Settings > Programming > API Keys (needs write access to create drafts).',
      ui.ButtonSet.OK);
    return;
  }

  const subject = table.getRow(cfg.subjectRow).getCell(1).getText().trim();
  const bodyMarkdown = rowToMarkdown(table.getRow(cfg.bodyRow), apiKey);

  try {
    const draft = createButtondownDraft(apiKey, subject, bodyMarkdown);
    const draftUrl = draft.id ? `https://buttondown.com/emails/${draft.id}` : null;
    showDraftCreatedDialog(subject, draftUrl, bodyMarkdown);
  } catch (e) {
    ui.alert('Error', `Could not create the Buttondown draft:\n${e.toString()}`, ui.ButtonSet.OK);
  }
}

/**
 * ui.alert() can't render a clickable link or a scrollable text area, so this uses
 * a small HTML dialog instead: the draft link plus the exact Markdown that was sent,
 * so a coach can sanity-check the conversion without leaving the Doc.
 */
function showDraftCreatedDialog(subject, draftUrl, bodyMarkdown) {
  const ui = DocumentApp.getUi();
  if (!draftUrl) {
    ui.alert('Buttondown Draft Created',
      `Subject: ${subject}\n\nDraft created, but no id came back to build a link. Check the Buttondown dashboard.`,
      ui.ButtonSet.OK);
    return;
  }
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5;">` +
    `<p><strong>Subject:</strong> ${escapeHtml(subject)}</p>` +
    `<p><a href="${escapeHtml(draftUrl)}" target="_blank">Review and send it in Buttondown</a></p>` +
    `<p style="margin-bottom:4px;"><strong>Markdown sent:</strong></p>` +
    `<textarea readonly style="width:100%;height:280px;box-sizing:border-box;font-family:monospace;font-size:12px;">${escapeHtml(bodyMarkdown)}</textarea>` +
    `</div>`
  ).setWidth(480).setHeight(440);
  ui.showModalDialog(html, 'Buttondown Draft Created');
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Walk up from the cursor (or the start of the current selection) to the nearest
 * enclosing Table, or null if the cursor isn't inside one.
 */
function findEnclosingTable(doc) {
  let element = null;
  const cursor = doc.getCursor();
  if (cursor) {
    element = cursor.getElement();
  } else {
    const selection = doc.getSelection();
    const ranges = selection ? selection.getRangeElements() : [];
    if (ranges.length > 0) element = ranges[0].getElement();
  }
  while (element && element.getType() !== DocumentApp.ElementType.TABLE) {
    element = element.getParent();
  }
  return element ? element.asTable() : null;
}
