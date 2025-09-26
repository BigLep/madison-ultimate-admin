/**
 * Build Email List Module
 * Creates email lists from Full Names by looking up parent/caregiver emails
 */

/**
 * Main function to build email list
 * Called from the menu
 */
function buildEmailList() {
  console.log('📧 Starting Build Email List...');

  try {
    showEmailListDialog();
  } catch (error) {
    console.error('Error building email list:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to build email list: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show the email list builder dialog
 */
function showEmailListDialog() {
  const html = createEmailListHtml();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Build Email List');
}

/**
 * Create HTML for email list dialog
 * @return {string} HTML content
 */
function createEmailListHtml() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            padding: 20px;
            margin: 0;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            font-weight: 500;
            margin-bottom: 8px;
            color: #202124;
            font-size: 14px;
          }
          textarea {
            width: 100%;
            padding: 10px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
            font-family: monospace;
            resize: vertical;
          }
          textarea:focus {
            outline: none;
            border-color: #1a73e8;
          }
          .note {
            font-size: 12px;
            color: #5f6368;
            margin-top: 5px;
          }
          .buttons {
            display: flex;
            gap: 10px;
            margin-top: 25px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
          }
          .btn {
            flex: 1;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: background-color 0.2s;
          }
          .btn-primary {
            background-color: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background-color: #1557b0;
          }
          .btn-secondary {
            background-color: #f8f9fa;
            color: #3c4043;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background-color: #f1f3f4;
          }
          .result-section {
            display: none;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
          }
          .copy-btn {
            background-color: #34a853;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            font-size: 12px;
            cursor: pointer;
            margin-top: 8px;
          }
          .copy-btn:hover {
            background-color: #2d8f46;
          }
          .progress-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.95);
            z-index: 1000;
          }
          .progress-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .success-message {
            background-color: #e8f5e8;
            color: #2d7d32;
            padding: 10px;
            border-radius: 4px;
            margin-bottom: 10px;
          }
        </style>
      </head>
      <body>
        <div class="form-group">
          <label for="fullNames">Full Names (one per line):</label>
          <textarea id="fullNames" rows="8" placeholder="Paste or type full names here, one per line:

John Smith
Jane Doe
Bob Johnson"></textarea>
          <div class="note">Enter student full names exactly as they appear in the roster</div>
        </div>

        <div class="buttons">
          <button class="btn btn-primary" onclick="buildEmailList()">Build Email List</button>
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>

        <div class="result-section" id="resultSection">
          <div id="successMessage" class="success-message"></div>
          <div class="form-group">
            <label for="emailResults">Email List Results:</label>
            <textarea id="emailResults" rows="10" readonly></textarea>
            <button class="copy-btn" onclick="copyToClipboard()">Copy to Clipboard</button>
            <div class="note">Emails are sorted and deduplicated. Click copy button to copy to clipboard.</div>
          </div>
        </div>

        <div class="progress-overlay" id="progressOverlay">
          <div class="progress-content">
            <div class="spinner"></div>
            <div style="font-size: 16px; font-weight: bold; color: #333;">
              Building Email List...
            </div>
            <div style="font-size: 14px; color: #666; margin-top: 8px;">
              Looking up parent/caregiver emails
            </div>
          </div>
        </div>

        <script>
          function buildEmailList() {
            const fullNames = document.getElementById('fullNames').value.trim();

            if (!fullNames) {
              alert('Please enter at least one full name');
              return;
            }

            // Show progress
            document.getElementById('progressOverlay').style.display = 'block';

            // Call the server-side function
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .processEmailList(fullNames);
          }

          function onSuccess(result) {
            document.getElementById('progressOverlay').style.display = 'none';

            // Show success message
            document.getElementById('successMessage').textContent =
              \`Found \${result.emailCount} unique emails for \${result.studentCount} students\`;

            // Show results
            document.getElementById('emailResults').value = result.emailList;
            document.getElementById('resultSection').style.display = 'block';
          }

          function onFailure(error) {
            document.getElementById('progressOverlay').style.display = 'none';
            alert('Error: ' + error.message);
          }

          function copyToClipboard() {
            const emailResults = document.getElementById('emailResults');
            emailResults.select();
            emailResults.setSelectionRange(0, 99999); // For mobile devices

            try {
              document.execCommand('copy');
              // Temporarily change button text to show success
              const btn = event.target;
              const originalText = btn.textContent;
              btn.textContent = 'Copied!';
              btn.style.backgroundColor = '#2d8f46';
              setTimeout(() => {
                btn.textContent = originalText;
                btn.style.backgroundColor = '#34a853';
              }, 2000);
            } catch (err) {
              console.error('Failed to copy to clipboard:', err);
              alert('Failed to copy to clipboard. Please select the text and copy manually.');
            }
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Process the email list request
 * @param {string} fullNamesText - Newline-separated list of full names
 * @return {Object} Result object with email list and counts
 */
function processEmailList(fullNamesText) {
  console.log('📧 Processing email list request...');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);

    if (!rosterSheet) {
      throw new Error('Roster sheet not found');
    }

    // Parse full names from input
    const fullNames = fullNamesText.split('\n')
      .map(name => name.trim())
      .filter(name => name.length > 0);

    if (fullNames.length === 0) {
      throw new Error('No valid full names provided');
    }

    console.log(`📋 Processing ${fullNames.length} full names: ${fullNames.join(', ')}`);

    // Get roster data
    const rosterData = rosterSheet.getDataRange().getValues();
    const headerRow = rosterData[0];

    // Find required column indices
    const fullNameColIndex = headerRow.indexOf(CONFIG.columns.fullName);
    const parent1EmailIndex = headerRow.indexOf(CONFIG.columns.parent1Email);
    const parent2EmailIndex = headerRow.indexOf(CONFIG.columns.parent2Email);

    if (fullNameColIndex === -1) {
      throw new Error(`${CONFIG.columns.fullName} column not found in roster`);
    }
    if (parent1EmailIndex === -1) {
      throw new Error(`${CONFIG.columns.parent1Email} column not found in roster`);
    }
    if (parent2EmailIndex === -1) {
      throw new Error(`${CONFIG.columns.parent2Email} column not found in roster`);
    }

    console.log(`📍 Found columns - Full Name: ${fullNameColIndex}, Parent 1 Email: ${parent1EmailIndex}, Parent 2 Email: ${parent2EmailIndex}`);

    // Collect emails for each student
    const emailSet = new Set();
    const foundStudents = [];
    const notFoundStudents = [];

    for (const targetName of fullNames) {
      let found = false;

      // Search through roster data (skip header row)
      for (let i = 1; i < rosterData.length; i++) {
        const row = rosterData[i];
        const rosterFullName = row[fullNameColIndex];

        if (rosterFullName && rosterFullName.toString().trim() === targetName) {
          found = true;
          foundStudents.push(targetName);

          // Get parent emails
          const parent1Email = row[parent1EmailIndex];
          const parent2Email = row[parent2EmailIndex];

          // Add valid emails to set (automatically deduplicates)
          if (parent1Email && parent1Email.toString().trim() && isValidEmail(parent1Email.toString().trim())) {
            emailSet.add(parent1Email.toString().trim().toLowerCase());
          }
          if (parent2Email && parent2Email.toString().trim() && isValidEmail(parent2Email.toString().trim())) {
            emailSet.add(parent2Email.toString().trim().toLowerCase());
          }

          break; // Found the student, move to next one
        }
      }

      if (!found) {
        notFoundStudents.push(targetName);
      }
    }

    // Convert set to sorted array
    const sortedEmails = Array.from(emailSet).sort();

    console.log(`✅ Found ${foundStudents.length} students, ${notFoundStudents.length} not found`);
    console.log(`📧 Collected ${sortedEmails.length} unique emails`);

    if (notFoundStudents.length > 0) {
      console.warn(`⚠️ Students not found: ${notFoundStudents.join(', ')}`);
    }

    // Create result
    const emailList = sortedEmails.join('\n');

    const result = {
      emailList: emailList,
      emailCount: sortedEmails.length,
      studentCount: foundStudents.length,
      foundStudents: foundStudents,
      notFoundStudents: notFoundStudents
    };

    console.log(`📋 Email list created successfully`);
    return result;

  } catch (error) {
    console.error('Error processing email list:', error);
    throw new Error(`Failed to process email list: ${error.message}`);
  }
}

/**
 * Simple email validation
 * @param {string} email - Email to validate
 * @return {boolean} Whether email is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}