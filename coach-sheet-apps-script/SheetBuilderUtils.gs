/**
 * Sheet Builder Utilities
 * Shared functions for building custom sheets with XLOOKUP formulas
 */

/**
 * Copy Full Name column from roster to a new sheet
 * @param {Sheet} targetSheet - The sheet to copy Full Name to
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {number} startRow - The row to start copying data (defaults to 2)
 * @return {Object} Object with fullNameColIndex and rowCount
 */
function copyFullNameColumn(targetSheet, rosterSheet, startRow = 2) {
  return copyFullNameColumnToColumn(targetSheet, rosterSheet, startRow, 1);
}

/**
 * Whether a roster row should appear on Build Practice Roster / Build Game Roster Prep.
 * Column missing or empty → include (backward compatible). Explicit FALSE (boolean or string) → exclude.
 * @param {*} cellValue - Value from "Include In Generated Rosters" column
 * @return {boolean}
 */
function isIncludedInGeneratedRosters(cellValue) {
  if (cellValue === false) return false;
  if (cellValue === true) return true;
  const s = cellValue != null ? String(cellValue).trim().toUpperCase() : '';
  if (s === 'FALSE') return false;
  return true;
}

/**
 * House-style lookup formula for one cell of a PlayerID-keyed sheet, the same shape as the
 * Roster's own derived columns: blank when the row has no key, blank when the key is not found,
 * otherwise the looked-up value. Row 2 keyed by A, looking up Full Name (F) in the Roster:
 *   =IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))
 * @param {string} keyLetter - Column letter (on the sheet holding the formula) of the lookup key
 * @param {number} row - 1-based sheet row the formula lives on
 * @param {string} sheetName - Sheet to look up in (unquoted; quotes are added and escaped here)
 * @param {string} sheetKeyLetter - That sheet's key column letter
 * @param {string} sheetValueLetter - That sheet's value column letter
 * @return {string}
 */
function playerIdLookupFormula(keyLetter, row, sheetName, sheetKeyLetter, sheetValueLetter) {
  const quoted = `'${String(sheetName).replace(/'/g, "''")}'`;
  const key = `$${keyLetter}${row}`;
  return `=IF(${key}="","",IFERROR(XLOOKUP(${key},${quoted}!$${sheetKeyLetter}:$${sheetKeyLetter},${quoted}!$${sheetValueLetter}:$${sheetValueLetter}),""))`;
}

/**
 * Fill one printout column (rows 2..numRows+1) with per-row playerIdLookupFormula formulas.
 * No-op when colIndex or sheetValueLetter is missing (the column is not part of this build).
 * @return {boolean} whether anything was written
 */
function fillLookupColumn(sheet, colIndex, numRows, keyLetter, sheetName, sheetKeyLetter, sheetValueLetter) {
  if (!colIndex || !sheetValueLetter || numRows <= 0) return false;
  const formulas = [];
  for (let i = 0; i < numRows; i++) {
    formulas.push([playerIdLookupFormula(keyLetter, i + 2, sheetName, sheetKeyLetter, sheetValueLetter)]);
  }
  sheet.getRange(2, colIndex, numRows, 1).setFormulas(formulas);
  return true;
}

/**
 * How a printout joins to an availability sheet: on PlayerID when the availability sheet has a
 * PlayerID column (every sheet built by 3.28 or later), otherwise on Full Name (an older sheet).
 * @param {Object} availColumns - From findAvailabilityColumns (playerIdColumn, fullNameColumn)
 * @param {string} playerIdLetter - The printout's PlayerID column letter
 * @param {string} fullNameLetter - The printout's Full Name column letter
 * @return {{keyLetter: string, sheetKeyLetter: string}}
 */
function availabilityJoin(availColumns, playerIdLetter, fullNameLetter) {
  if (availColumns.playerIdColumn) {
    return { keyLetter: playerIdLetter, sheetKeyLetter: availColumns.playerIdColumn };
  }
  console.warn(`⚠️ Availability sheet has no "${AVAILABILITY_ROW_HEADERS.playerId}" column; joining by Full Name. Run Build Practice/Game Availability to add it.`);
  return { keyLetter: fullNameLetter, sheetKeyLetter: availColumns.fullNameColumn };
}

/**
 * Seed the player rows of a printout (practice roster, game roster prep) from the Roster: one row
 * per Player whose Include In Generated Rosters is not FALSE, PlayerID written as a value in
 * playerIdColumn and Full Name as a Roster formula keyed by it in fullNameColumn. PlayerID is the
 * only per-player value a printout holds; everything else is a lookup on it (ADR 0004).
 * @param {Sheet} targetSheet
 * @param {Sheet} rosterSheet
 * @param {number} startRow - First data row on the printout (2)
 * @param {number} playerIdColumn - 1-based printout column for PlayerID
 * @param {number} fullNameColumn - 1-based printout column for Full Name
 * @return {{rowCount: number, playerIdLetter: string, fullNameLetter: string, rosterIdLetter: string}}
 */
function seedPrintoutPlayerRows(targetSheet, rosterSheet, startRow, playerIdColumn, fullNameColumn) {
  const table = readRosterTable(rosterSheet);
  const idCol = table.col(CONFIG.columns.playerId);
  const includeCol = table.headers.indexOf(CONFIG.columns.includeInGeneratedRosters);
  const ids = [];
  table.rows.forEach(row => {
    const id = (row[idCol] === null || row[idCol] === undefined) ? '' : String(row[idCol]).trim();
    if (!id) return;
    if (includeCol !== -1 && !isIncludedInGeneratedRosters(row[includeCol])) return;
    ids.push(id);
  });
  if (ids.length === 0) {
    throw new Error('No student data found in roster');
  }

  const playerIdLetter = getColumnLetter(playerIdColumn);
  const rosterIdLetter = getColumnLetter(table.col(CONFIG.columns.playerId) + 1);
  const rosterNameLetter = getColumnLetter(table.col(CONFIG.columns.fullName) + 1);
  targetSheet.getRange(startRow, playerIdColumn, ids.length, 1).setValues(ids.map(id => [id]));
  targetSheet.getRange(startRow, fullNameColumn, ids.length, 1).setFormulas(
    ids.map((id, i) => [playerIdLookupFormula(playerIdLetter, startRow + i, CONFIG.roster.sheetName, rosterIdLetter, rosterNameLetter)])
  );
  return { rowCount: ids.length, playerIdLetter: playerIdLetter, fullNameLetter: getColumnLetter(fullNameColumn), rosterIdLetter: rosterIdLetter };
}

/**
 * Sort rank of a Team value: its position in CONFIG.teams; blank or unlisted values rank after
 * every listed Team so unassigned Players sit at the bottom of a printout.
 * @param {*} value
 * @return {number}
 */
function teamSortRank(value) {
  const v = value === null || value === undefined ? '' : String(value).trim();
  const i = v === '' ? -1 : CONFIG.teams.indexOf(v);
  return i === -1 ? CONFIG.teams.length : i;
}

/**
 * Append a temporary numeric sort-key column to a printout (after insertAfterCol) holding
 * rankFn(value) for each data row of sourceCol, so Range.sort() can order by a custom list.
 * The caller sorts, then removes the column with sheet.deleteColumn(returned index), highest
 * temp column first when it added more than one.
 * @param {Sheet} sheet
 * @param {number} sourceCol - 1-based column whose values are ranked
 * @param {number} numRows - Data rows (from row 2)
 * @param {number} insertAfterCol - Column after which the temp column is inserted
 * @param {function(*): number} rankFn
 * @return {number} 1-based index of the temp column
 */
function insertSortRankColumn(sheet, sourceCol, numRows, insertAfterCol, rankFn) {
  sheet.insertColumnAfter(insertAfterCol);
  const rankCol = insertAfterCol + 1;
  const values = sheet.getRange(2, sourceCol, numRows, 1).getValues();
  sheet.getRange(2, rankCol, numRows, 1).setValues(values.map(function (row) { return [rankFn(row[0])]; }));
  return rankCol;
}

/**
 * Copy Full Name column from roster to a specific column in new sheet
 * @param {Sheet} targetSheet - The sheet to copy Full Name to
 * @param {Sheet} rosterSheet - The source roster sheet
 * @param {number} startRow - The row to start copying data
 * @param {number} targetColumn - The column to copy Full Name to (1-based)
 * @return {Object} Object with fullNameColIndex and rowCount
 */
function copyFullNameColumnToColumn(targetSheet, rosterSheet, startRow, targetColumn) {
  const rosterHeaderRow = rosterSheet.getRange(ROSTER_HEADER_ROW, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const fullNameColIndex = rosterHeaderRow.findIndex(name => name === CONFIG.columns.fullName);

  if (fullNameColIndex === -1) {
    throw new Error(`${CONFIG.columns.fullName} column not found in roster sheet`);
  }

  const includeColName = CONFIG.columns.includeInGeneratedRosters;
  const includeColIndex = rosterHeaderRow.findIndex(function (name) {
    return name === includeColName;
  });

  // Use configurable first data row so roster layout (1 header vs 5 metadata rows) can change
  const lastRow = rosterSheet.getLastRow();
  const numDataRows = Math.max(0, lastRow - ROSTER_FIRST_DATA_ROW + 1);
  const rosterDataRange = rosterSheet.getRange(ROSTER_FIRST_DATA_ROW, fullNameColIndex + 1, numDataRows, 1);
  const fullNameValues = rosterDataRange.getValues();

  let includeValues = null;
  if (includeColIndex !== -1) {
    includeValues = rosterSheet.getRange(ROSTER_FIRST_DATA_ROW, includeColIndex + 1, numDataRows, 1).getValues();
  }

  const nonEmptyFullNames = [];
  for (let i = 0; i < fullNameValues.length; i++) {
    const nameCell = fullNameValues[i][0];
    if (!nameCell || String(nameCell).trim() === '') continue;
    if (includeValues && !isIncludedInGeneratedRosters(includeValues[i][0])) continue;
    nonEmptyFullNames.push([nameCell]);
  }

  if (nonEmptyFullNames.length === 0) {
    throw new Error('No student data found in roster');
  }

  targetSheet.getRange(startRow, targetColumn, nonEmptyFullNames.length, 1).setValues(nonEmptyFullNames);

  return {
    fullNameColIndex: fullNameColIndex,
    rowCount: nonEmptyFullNames.length
  };
}

/**
 * Create XLOOKUP formulas for columns from roster sheet
 * @param {Sheet} targetSheet - The sheet to add formulas to
 * @param {Array} columnNames - Array of column names to create formulas for
 * @param {number} startColumn - The column index to start adding formulas (1-based)
 * @param {number} rowCount - Number of rows to copy formulas down
 * @param {string} lookupSheetName - Name of the sheet to lookup from (defaults to CONFIG.roster.sheetName)
 */
function createRosterXlookupFormulas(targetSheet, columnNames, startColumn, rowCount, lookupSheetName = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(lookupSheetName || CONFIG.roster.sheetName);
  const isRoster = !lookupSheetName || lookupSheetName === CONFIG.roster.sheetName;
  const headerRowNum = isRoster ? ROSTER_HEADER_ROW : 1;
  const sourceHeaderRow = sourceSheet.getRange(headerRowNum, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  const fullNameColIndex = sourceHeaderRow.findIndex(name => name === CONFIG.columns.fullName);
  if (fullNameColIndex === -1) {
    throw new Error(`${CONFIG.columns.fullName} column not found in source sheet`);
  }
  const fullNameColumnLetter = getColumnLetter(fullNameColIndex + 1);
  
  columnNames.forEach((columnName, index) => {
    const targetColumn = startColumn + index;
    
    const sourceColumnIndex = sourceHeaderRow.findIndex(name => name === columnName);
    if (sourceColumnIndex === -1) {
      console.warn(`Column "${columnName}" not found in ${lookupSheetName || CONFIG.roster.sheetName}`);
      return;
    }
    
    const sourceColumnLetter = getColumnLetter(sourceColumnIndex + 1);
    const sheetName = lookupSheetName || CONFIG.roster.sheetName;
    
    const formula = `=IFERROR(XLOOKUP(A2,'${sheetName}'!${fullNameColumnLetter}:${fullNameColumnLetter},'${sheetName}'!${sourceColumnLetter}:${sourceColumnLetter}),"")`;
    
    targetSheet.getRange(2, targetColumn).setFormula(formula);
    
    if (rowCount > 1) {
      const sourceRange = targetSheet.getRange(2, targetColumn, 1, 1);
      const targetRange = targetSheet.getRange(3, targetColumn, rowCount - 1, 1);
      sourceRange.copyTo(targetRange);
    }
  });
}

/**
 * Copy column formatting from source sheet to target sheet
 * @param {Sheet} targetSheet - The sheet to apply formatting to
 * @param {Sheet} sourceSheet - The source sheet to copy formatting from
 * @param {Array} headers - Array of column headers in target sheet
 * @param {Array} sourceHeaderRow - Header row from source sheet
 */
function copyColumnFormatting(targetSheet, sourceSheet, headers, sourceHeaderRow) {
  headers.forEach((columnName, newColumnIndex) => {
    const sourceColumnIndex = sourceHeaderRow.findIndex(name => name === columnName);
    
    if (sourceColumnIndex === -1) {
      console.warn(`Column "${columnName}" not found in source sheet for formatting`);
      return;
    }
    
    const newColumn = newColumnIndex + 1;
    const sourceColumn = sourceColumnIndex + 1;
    
    try {
      const sourceColumnWidth = sourceSheet.getColumnWidth(sourceColumn);
      targetSheet.setColumnWidth(newColumn, sourceColumnWidth);
      
      const sourceFormatCell = sourceSheet.getRange(ROSTER_FIRST_DATA_ROW, sourceColumn);
      const newFormatCell = targetSheet.getRange(2, newColumn);
      
      const numberFormat = sourceFormatCell.getNumberFormat();
      if (numberFormat) {
        const newColumnRange = targetSheet.getRange(2, newColumn, targetSheet.getMaxRows() - 1, 1);
        newColumnRange.setNumberFormat(numberFormat);
      }
      
      const textWrapping = sourceFormatCell.getWrap();
      const newColumnRange = targetSheet.getRange(2, newColumn, targetSheet.getMaxRows() - 1, 1);
      newColumnRange.setWrap(textWrapping);
      
      const horizontalAlignment = sourceFormatCell.getHorizontalAlignment();
      newColumnRange.setHorizontalAlignment(horizontalAlignment);
      
      const verticalAlignment = sourceFormatCell.getVerticalAlignment();
      newColumnRange.setVerticalAlignment(verticalAlignment);
      
      const fontFamily = sourceFormatCell.getFontFamily();
      const fontSize = sourceFormatCell.getFontSize();
      newColumnRange.setFontFamily(fontFamily);
      newColumnRange.setFontSize(fontSize);
      
      console.log(`✅ Copied formatting for column "${columnName}" (width: ${sourceColumnWidth}px)`);
      
    } catch (error) {
      console.warn(`Could not copy formatting for column "${columnName}":`, error);
    }
  });
}

/**
 * Copy conditional formatting from source sheet to target sheet
 * @param {Sheet} targetSheet - The sheet to apply conditional formatting to
 * @param {Sheet} sourceSheet - The source sheet to copy formatting from
 * @param {number} totalRows - Total number of rows in target sheet
 * @param {number} totalColumns - Total number of columns in target sheet
 */
function copyConditionalFormatting(targetSheet, sourceSheet, totalRows, totalColumns) {
  try {
    const sourceRules = sourceSheet.getConditionalFormatRules();
    
    if (sourceRules.length === 0) {
      console.log('No conditional formatting rules found in source sheet');
      return;
    }
    
    console.log(`Found ${sourceRules.length} conditional formatting rules in source sheet`);
    
    const entireSheetRange = targetSheet.getRange(1, 1, totalRows, totalColumns);
    const newRules = [];
    
    sourceRules.forEach((rule, ruleIndex) => {
      try {
        const newRule = rule.copy().setRanges([entireSheetRange]);
        newRules.push(newRule);
        console.log(`✅ Applied conditional formatting rule ${ruleIndex + 1} to target sheet`);
      } catch (ruleError) {
        console.warn(`Could not copy conditional formatting rule ${ruleIndex + 1}:`, ruleError);
      }
    });
    
    if (newRules.length > 0) {
      targetSheet.setConditionalFormatRules(newRules);
      console.log(`✅ Applied ${newRules.length} conditional formatting rules to target sheet`);
    }
    
  } catch (error) {
    console.warn('Could not copy conditional formatting:', error);
  }
}

/**
 * Copy data validation from source sheet to target sheet for specific columns
 * @param {Sheet} targetSheet - The sheet to apply data validation to
 * @param {Sheet} sourceSheet - The source sheet to copy validation from
 * @param {Array} columnMappings - Array of {sourceColumn, targetColumn} objects
 * @param {number} rowCount - Number of rows to apply validation to
 */
function copyDataValidation(targetSheet, sourceSheet, columnMappings, rowCount) {
  const sourceHeaderRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  columnMappings.forEach(mapping => {
    const sourceColIndex = sourceHeaderRow.indexOf(mapping.sourceColumn) + 1;
    if (sourceColIndex === 0) {
      console.warn(`Source column "${mapping.sourceColumn}" not found`);
      return;
    }
    
    const sourceCell = sourceSheet.getRange(ROSTER_FIRST_DATA_ROW, sourceColIndex);
    const validation = sourceCell.getDataValidation();
    
    if (!validation) {
      console.log(`No data validation found for column "${mapping.sourceColumn}"`);
      return;
    }
    
    const targetRange = targetSheet.getRange(2, mapping.targetColumn, rowCount, 1);
    targetRange.setDataValidation(validation);
    
    console.log(`✅ Copied data validation from "${mapping.sourceColumn}" to column ${mapping.targetColumn}`);
  });
}

/**
 * Style header row with standard formatting
 * @param {Sheet} sheet - The sheet to style
 * @param {number} columnCount - Number of columns in header
 */
function styleHeaderRow(sheet, columnCount) {
  const headerRange = sheet.getRange(1, 1, 1, columnCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
}