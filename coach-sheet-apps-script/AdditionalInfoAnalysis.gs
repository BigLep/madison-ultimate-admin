/**
 * Additional Info Analysis Module
 *
 * This module handles analysis of Additional Info responses, creating a separate
 * analysis sheet to avoid modifying the original Additional Info data.
 */

/**
 * Analyze Additional Info responses for matches, suggestions, and potential duplicates
 * Creates a separate "Additional Info Analysis" sheet with the results
 */
function analyzeAdditionalInfoResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  const additionalInfoSheet = ss.getSheetByName(CONFIG.additionalInfo.sheetName);

  if (!rosterSheet || !additionalInfoSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find Roster or Additional Info sheets');
    return;
  }

  // Get current timestamp for analysis
  const analysisTimestamp = new Date().toISOString();

  // Create or get the analysis sheet
  const analysisSheetName = 'Additional Info Analysis';
  let analysisSheet = ss.getSheetByName(analysisSheetName);

  if (!analysisSheet) {
    // Create new analysis sheet next to Additional Info sheet
    const additionalInfoIndex = additionalInfoSheet.getIndex();
    analysisSheet = ss.insertSheet(analysisSheetName, additionalInfoIndex + 1);
  } else {
    // Clear existing content
    analysisSheet.clear();
  }

  // Set up the analysis sheet headers
  const headers = [
    'Player Name (First & Last)',
    `Roster Match Status as of ${analysisTimestamp}`,
    'Suggested Match',
    'Potential Duplicate?'
  ];

  analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = analysisSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');

  // Get roster data for matching
  const rosterData = getRosterDataForMatching(rosterSheet);

  // Get Additional Info data
  const additionalInfoData = getAdditionalInfoData(additionalInfoSheet);

  if (!additionalInfoData || additionalInfoData.length === 0) {
    SpreadsheetApp.getUi().alert('No Additional Info data found');
    return;
  }

  // Analyze each response and populate the analysis sheet
  const analysisResults = analyzeResponses(additionalInfoData, rosterData);

  // Write results to analysis sheet
  if (analysisResults.length > 0) {
    const resultsRange = analysisSheet.getRange(2, 1, analysisResults.length, headers.length);
    resultsRange.setValues(analysisResults);

    // Apply conditional formatting
    applyConditionalFormatting(analysisSheet, analysisResults.length + 1);
  }

  // Auto-resize columns
  analysisSheet.autoResizeColumns(1, headers.length);

  // Apply Format Spruce Up formatting silently
  applySpruceUpFormatting(analysisSheet);

  // Show summary and activate the analysis sheet
  showAnalysisSummary(analysisResults, analysisTimestamp);
  analysisSheet.activate();
}

/**
 * Get roster data for matching purposes
 */
function getRosterDataForMatching(rosterSheet) {
  const headers = rosterSheet.getRange(1, 1, 1, rosterSheet.getMaxColumns()).getValues()[0];
  const fullNameCol = headers.indexOf('Full Name') + 1;
  const lastNameCol = headers.indexOf('Last Name') + 1;

  if (!fullNameCol) {
    throw new Error('Could not find "Full Name" column in roster');
  }

  if (!lastNameCol) {
    throw new Error('Could not find "Last Name" column in roster');
  }

  const rosterLastRow = rosterSheet.getLastRow();
  const rosterNames = new Set();
  const rosterLastNames = new Map(); // Map last names to full names for suggestions

  for (let row = FIRST_DATA_ROW; row <= rosterLastRow; row++) {
    const fullName = rosterSheet.getRange(row, fullNameCol).getValue();
    const lastName = rosterSheet.getRange(row, lastNameCol).getValue();

    if (fullName) {
      const exactFullName = fullName.toString();
      rosterNames.add(exactFullName);

      // Store last name mapping for suggestions
      if (lastName) {
        const normalizedLastName = lastName.toString().trim().toLowerCase();
        if (!rosterLastNames.has(normalizedLastName)) {
          rosterLastNames.set(normalizedLastName, []);
        }
        rosterLastNames.get(normalizedLastName).push(fullName.toString().trim());
      }
    }
  }

  return {
    rosterNames: rosterNames,
    rosterLastNames: rosterLastNames
  };
}

/**
 * Get Additional Info data
 */
function getAdditionalInfoData(additionalInfoSheet) {
  const additionalInfoData = additionalInfoSheet.getDataRange().getValues();
  if (additionalInfoData.length < 2) {
    return null;
  }

  const headers = additionalInfoData[0];
  const playerNameCol = headers.indexOf('Player Name (First & Last)');

  if (playerNameCol === -1) {
    throw new Error('Could not find "Player Name (First & Last)" column in Additional Info');
  }

  // Extract just the player names and their row indices
  const playerData = [];
  for (let i = 1; i < additionalInfoData.length; i++) {
    const playerName = additionalInfoData[i][playerNameCol];
    if (playerName) {
      playerData.push({
        rowIndex: i,
        playerName: playerName.toString()
      });
    }
  }

  return playerData;
}

/**
 * Analyze responses for matches, suggestions, and duplicates
 */
function analyzeResponses(additionalInfoData, rosterData) {
  const { rosterNames, rosterLastNames } = rosterData;

  // Detect potential duplicates by tracking player names
  const playerNameCounts = new Map();
  additionalInfoData.forEach(item => {
    const exactName = item.playerName;
    playerNameCounts.set(exactName, (playerNameCounts.get(exactName) || 0) + 1);
  });

  // Analyze each response
  const results = [];

  additionalInfoData.forEach(item => {
    const playerName = item.playerName;

    const isMatched = rosterNames.has(playerName);
    const suggestion = isMatched ? '' : findLastNameSuggestion(playerName, rosterLastNames);
    const isDuplicate = playerNameCounts.get(playerName) > 1;

    results.push([
      playerName,
      isMatched ? 'MATCHED' : 'UNMATCHED',
      suggestion || '',
      isDuplicate ? 'YES' : 'NO'
    ]);
  });

  return results;
}

/**
 * Apply conditional formatting to the analysis sheet
 */
function applyConditionalFormatting(sheet, totalRows) {
  // Format MATCHED/UNMATCHED column
  const statusRange = sheet.getRange(2, 2, totalRows - 1, 1);

  // Green for MATCHED
  const matchedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MATCHED')
    .setBackground('#d9ead3')
    .setFontColor('#0d652d')
    .setRanges([statusRange])
    .build();

  // Red for UNMATCHED
  const unmatchedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('UNMATCHED')
    .setBackground('#fce5cd')
    .setFontColor('#cc4125')
    .setRanges([statusRange])
    .build();

  // Yellow for potential duplicates
  const duplicateRange = sheet.getRange(2, 4, totalRows - 1, 1);
  const duplicateRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('YES')
    .setBackground('#fff2cc')
    .setFontColor('#bf9000')
    .setRanges([duplicateRange])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(matchedRule, unmatchedRule, duplicateRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Show analysis summary dialog
 */
function showAnalysisSummary(results, timestamp) {
  let matchedCount = 0;
  let unmatchedCount = 0;
  let duplicateCount = 0;

  results.forEach(result => {
    if (result[1] === 'MATCHED') {
      matchedCount++;
    } else {
      unmatchedCount++;
    }

    if (result[3] === 'YES') {
      duplicateCount++;
    }
  });

  SpreadsheetApp.getUi().alert(
    'Analysis Complete',
    `Analysis completed at ${timestamp}

Results:
• ${matchedCount} responses matched roster entries
• ${unmatchedCount} responses did not match roster entries
• ${duplicateCount} potential duplicate names detected

A new "Additional Info Analysis" sheet has been created with the detailed results.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Helper function to find potential roster matches based on last name
 */
function findLastNameSuggestion(playerName, rosterLastNames) {
  // Extract last name from the player name (assume "First Last" format)
  const nameParts = playerName.trim().split(/\s+/);
  if (nameParts.length < 2) {
    return null; // Can't extract last name
  }

  const lastName = nameParts[nameParts.length - 1].toLowerCase();

  // Look for matching last names in roster
  if (rosterLastNames.has(lastName)) {
    const matches = rosterLastNames.get(lastName);
    if (matches.length === 1) {
      return matches[0]; // Single suggestion
    } else if (matches.length > 1) {
      return matches.join(' OR '); // Multiple suggestions
    }
  }

  return null; // No suggestion found
}