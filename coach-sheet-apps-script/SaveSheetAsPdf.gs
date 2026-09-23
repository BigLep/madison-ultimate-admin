/**
 * Save Sheet as PDF: exports the active sheet as a single-page-tall PDF (fit to height, Letter,
 * portrait, top/left aligned, gridlines shown; 0.5" top margin, 0.25" on the other three sides)
 * and opens it in a new browser tab. Generic, works on any sheet.
 *
 * SpreadsheetApp has no page-setup API for scale/fit-to-page/alignment (Sheet.setMargins() only
 * covers margins, and even that has no effect on the interactive File > Print dialog in practice
 * — see configurePrintSettings's note in BuildPracticeRoster.gs). This bypasses that dialog
 * entirely by fetching the spreadsheet's PDF export endpoint directly, which supports fit/margin/
 * alignment via (undocumented but long-stable) query parameters.
 */

const SAVE_AS_PDF_MARGINS = { top: 0.5, bottom: 0.25, left: 0.25, right: 0.25 };

/** Menu entry point: exports the active sheet and shows a dialog that opens the PDF. */
function showSaveSheetAsPdfDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const result = exportSheetAsPdf_(sheet);
    const html = createSaveSheetAsPdfHtml_(result.base64, result.filename);
    const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(340).setHeight(140);
    ui.showModalDialog(htmlOutput, 'Save Sheet as PDF');
  } catch (error) {
    console.error('Error saving sheet as PDF:', error);
    ui.alert('Error', `Failed to save sheet as PDF: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Fetch one sheet's PDF from the spreadsheet's export endpoint, fit to page height with the
 * coach's standard print settings (SAVE_AS_PDF_MARGINS): Letter, portrait, top/left aligned,
 * gridlines shown. Hidden rows/columns (e.g. a printout's hidden PlayerID key column) are excluded
 * automatically by the export endpoint, same as the interactive Print dialog.
 * @param {Sheet} sheet
 * @return {{base64: string, filename: string}}
 */
function exportSheetAsPdf_(sheet) {
  const ss = sheet.getParent();
  const params = {
    format: 'pdf',
    size: 'letter',
    portrait: 'true',
    scale: '3', // Fit Height
    top_margin: String(SAVE_AS_PDF_MARGINS.top),
    bottom_margin: String(SAVE_AS_PDF_MARGINS.bottom),
    left_margin: String(SAVE_AS_PDF_MARGINS.left),
    right_margin: String(SAVE_AS_PDF_MARGINS.right),
    horizontal_alignment: 'LEFT',
    vertical_alignment: 'TOP',
    gridlines: 'true',
    printtitle: 'false',
    sheetnames: 'false',
    pagenum: 'UNDEFINED',
    gid: String(sheet.getSheetId())
  };
  const query = Object.keys(params)
    .map(function (key) { return key + '=' + encodeURIComponent(params[key]); })
    .join('&');
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?${query}`;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  const blob = response.getBlob();

  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: sanitizePdfFilename_(sheet.getName()) + '.pdf'
  };
}

/**
 * Turn a sheet tab name into a filesystem-friendly PDF filename: replaces characters invalid (or
 * awkward) in filenames on Windows/macOS (\ / : * ? " < > |) with "-", collapses whitespace, and
 * otherwise leaves the name (including emoji) alone.
 * @param {string} sheetName
 * @return {string}
 */
function sanitizePdfFilename_(sheetName) {
  return sheetName.replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
}

/**
 * HTML for the small result dialog: opens the PDF in a new tab as soon as it loads (browsers
 * preview a PDF opened this way rather than force-downloading it, which is fine here), plus a
 * fallback link in case the popup is blocked.
 * @param {string} base64Pdf
 * @param {string} filename
 * @return {string}
 */
function createSaveSheetAsPdfHtml_(base64Pdf, filename) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: 'Google Sans', Arial, sans-serif; padding: 16px; text-align: center; }
          .btn {
            display: inline-block; margin-top: 12px; padding: 10px 20px; background-color: #1a73e8;
            color: white; border: none; border-radius: 4px; font-size: 14px; font-weight: 500;
            cursor: pointer; text-decoration: none;
          }
          .note { font-size: 12px; color: #5f6368; margin-top: 10px; word-break: break-word; }
        </style>
      </head>
      <body>
        <div id="status">Opening PDF…</div>
        <a class="btn" id="openLink" style="display:none;" target="_blank" rel="noopener">Open PDF</a>
        <div class="note">${filename}</div>
        <script>
          const base64 = "${base64Pdf}";
          const byteChars = atob(base64);
          const byteNumbers = new Array(byteChars.length);
          for (let i = 0; i < byteChars.length; i++) byteNumbers[i] = byteChars.charCodeAt(i);
          const blob = new Blob([new Uint8Array(byteNumbers)], { type: 'application/pdf' });
          const blobUrl = URL.createObjectURL(blob);

          const link = document.getElementById('openLink');
          link.href = blobUrl;

          const opened = window.open(blobUrl, '_blank');
          if (!opened) {
            document.getElementById('status').textContent = 'Pop-up blocked. Click below to open the PDF:';
            link.style.display = 'inline-block';
            link.textContent = 'Open PDF';
          } else {
            document.getElementById('status').textContent = 'PDF opened in a new tab.';
            link.style.display = 'inline-block';
            link.textContent = 'Open again';
          }
        </script>
      </body>
    </html>
  `;
}
