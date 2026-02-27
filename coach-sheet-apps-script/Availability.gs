/**
 * Availability Builder Module
 * Creates availability tracking columns based on dates from info sheets
 * Supports both Practice and Game availability tracking
 */

// Shared validation options for all availability tracking (colors managed separately via conditional formatting)
const AVAILABILITY_VALIDATION_OPTIONS = [
  { value: '👍 Planning to be there' },
  { value: '👎 Can\'t make it' },
  { value: '❓ Not sure yet' },
  { value: 'Was there' },
  { value: 'Wasn\'t there' }
];

// Configuration for Practice Availability feature
const PRACTICE_AVAILABILITY_CONFIG = {
  type: 'practice',
  emoji: '🏃',
  infoSheet: '📍Practice Info',
  availabilitySheet: 'Practice Availability',
  validationOptions: AVAILABILITY_VALIDATION_OPTIONS,
  skipConfig: {
    columnName: 'note',
    skipCondition: 'startsWith',
    skipValue: 'Bye'
  }
};

// Game Activation Status (game availability only): dropdown + background colors
const GAME_ACTIVATION_STATUS_OPTIONS = [
  { value: 'Active', backgroundColor: '#d9ead3' },   // light green
  { value: 'Inactive', backgroundColor: '#f4cccc' }, // light red
  { value: 'TBD', backgroundColor: '#efefef' }      // light grey / white
];

// Configuration for Game Availability feature
const GAME_AVAILABILITY_CONFIG = {
  type: 'game',
  emoji: '🎮',
  infoSheet: '📍Game Info',
  availabilitySheet: 'Game Availability',
  validationOptions: AVAILABILITY_VALIDATION_OPTIONS,
  skipConfig: {
    columnName: 'game #',
    skipCondition: 'equals',
    skipValue: 'Bye'
  },
  // Game-only: three columns per date — $date Availability, $date Activation Status, $date Note (free text)
  columnsPerDate: [
    { suffix: ' Availability', useAvailabilityValidation: true },
    { suffix: ' Activation Status', useAvailabilityValidation: false, useActivationStatusValidation: true },
    { suffix: ' Note', useAvailabilityValidation: false, isFreeText: true }
  ]
};

/**
 * Shared function to build availability columns for practice or game
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 */
function buildAvailability(config) {
  const typeCapitalized = config.type.charAt(0).toUpperCase() + config.type.slice(1);
  console.log(`${config.emoji} Starting Build ${typeCapitalized} Availability...`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get dates from the appropriate info sheet
    const dates = getDatesFromInfoSheet(ss, config);
    
    if (dates.length === 0) {
      SpreadsheetApp.getUi().alert(
        `No ${typeCapitalized} Dates Found`,
        `No ${config.type} dates found in "${config.infoSheet}" sheet. Please ensure the sheet exists and contains ${config.type} date information.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Build availability columns in the appropriate availability sheet
    const result = buildAvailabilityColumns(ss, dates, config);
    
    console.log(`✅ ${typeCapitalized} Availability build complete`);
    
    let message = `Successfully processed ${config.availabilitySheet} sheet for ${dates.length} ${config.type} dates.\n\n`;
    
    if (result.columnsCreated > 0) {
      message += `📊 ${result.columnSummary}`;
    }
    
    if (result.columnsSkipped > 0) {
      if (result.columnsCreated > 0) message += '\n\n';
      message += `⏭️ ${result.skippedSummary}`;
    }
    
    if (result.columnsCreated === 0 && result.columnsSkipped === 0) {
      message += 'No changes needed - all columns already exist.';
    }
    
    message += '\n\n🎯 Data validation applied only to new columns - existing validation and colors preserved.';
    
    SpreadsheetApp.getUi().alert(`${typeCapitalized} Availability Updated!`, message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error(`Error building ${typeCapitalized} Availability:`, error);
    SpreadsheetApp.getUi().alert('Error', `Failed to build ${typeCapitalized} Availability: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Main function to build practice availability columns
 * Called from the menu
 */
function buildPracticeAvailability() {
  buildAvailability(PRACTICE_AVAILABILITY_CONFIG);
}

/**
 * Main function to build game availability columns
 * Called from the menu
 */
function buildGameAvailability() {
  buildAvailability(GAME_AVAILABILITY_CONFIG);
}

/**
 * Shared function to extract dates from an info sheet
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 * @return {Array} Array of date objects: {date, formattedDate}
 */
function getDatesFromInfoSheet(ss, config) {
  const infoSheet = ss.getSheetByName(config.infoSheet);
  
  if (!infoSheet) {
    throw new Error(`${config.type.charAt(0).toUpperCase() + config.type.slice(1)} Info sheet "${config.infoSheet}" not found`);
  }
  
  console.log(`📅 Reading ${config.type} dates from "${config.infoSheet}"`);
  
  // Look for a Date column in the info sheet
  const headerRow = infoSheet.getRange(1, 1, 1, infoSheet.getLastColumn()).getValues()[0];
  const dateColumnIndex = headerRow.findIndex(header => 
    header && header.toString().toLowerCase().includes('date')
  );
  
  if (dateColumnIndex === -1) {
    throw new Error(`No date column found in "${config.infoSheet}" sheet. Please ensure there is a column with "date" in the header.`);
  }
  
  console.log(`📍 Found date column at index ${dateColumnIndex + 1}`);
  
  // Look for skip column if configured
  let skipColumnIndex = -1;
  if (config.skipConfig) {
    skipColumnIndex = headerRow.findIndex(header => 
      header && header.toString().toLowerCase().includes(config.skipConfig.columnName.toLowerCase())
    );
    
    if (skipColumnIndex !== -1) {
      console.log(`📍 Found skip column "${config.skipConfig.columnName}" at index ${skipColumnIndex + 1}`);
    } else {
      console.log(`⚠️ Skip column "${config.skipConfig.columnName}" not found - will not skip any rows`);
    }
  }
  
  // Get all dates from the date column (skip header row)
  const lastRow = infoSheet.getLastRow();
  if (lastRow <= 1) {
    console.log(`⚠️ No ${config.type} data found in ${config.infoSheet} sheet`);
    return [];
  }
  
  // Get all data we need (date column and skip column if applicable)
  const columnsNeeded = skipColumnIndex !== -1 ? 
    Math.max(dateColumnIndex + 1, skipColumnIndex + 1) : 
    dateColumnIndex + 1;
  const allData = infoSheet.getRange(2, 1, lastRow - 1, columnsNeeded).getValues();
  const dates = [];
  
  allData.forEach((row, index) => {
    const dateValue = row[dateColumnIndex];
    
    // Check if we should skip this row based on skip configuration
    if (config.skipConfig && skipColumnIndex !== -1) {
      const skipValue = row[skipColumnIndex];
      let shouldSkip = false;
      
      if (skipValue && skipValue.toString().trim() !== '') {
        const skipValueStr = skipValue.toString().trim();
        
        if (config.skipConfig.skipCondition === 'startsWith') {
          shouldSkip = skipValueStr.toLowerCase().startsWith(config.skipConfig.skipValue.toLowerCase());
        } else if (config.skipConfig.skipCondition === 'equals') {
          shouldSkip = skipValueStr.toLowerCase() === config.skipConfig.skipValue.toLowerCase();
        }
      }
      
      if (shouldSkip) {
        console.log(`⏭️ Skipping row ${index + 2}: ${config.skipConfig.columnName} = "${skipValue}"`);
        return; // Skip this iteration
      }
    }
    
    if (dateValue && dateValue !== '') {
      try {
        // Handle both Date objects and date strings
        let dateObj;
        if (dateValue instanceof Date) {
          dateObj = dateValue;
        } else {
          dateObj = new Date(dateValue);
        }
        
        // Validate that it's a valid date
        if (!isNaN(dateObj.getTime())) {
          const formattedDate = formatDateForColumn(dateObj);
          dates.push({
            date: dateObj,
            formattedDate: formattedDate,
            rowIndex: index + 2 // +2 for 1-based indexing and header row
          });
          console.log(`${config.emoji} Found ${config.type} date: ${formattedDate} (row ${index + 2})`);
        } else {
          console.warn(`⚠️ Invalid date in row ${index + 2}: "${dateValue}"`);
        }
      } catch (error) {
        console.warn(`⚠️ Error parsing date in row ${index + 2}: "${dateValue}" - ${error.message}`);
      }
    }
  });
  
  console.log(`🎯 Found ${dates.length} valid ${config.type} dates`);
  return dates;
}

/**
 * Format a date for use in column headers (e.g., "9/26")
 * @param {Date} date - The date to format
 * @return {string} Formatted date string
 */
function formatDateForColumn(date) {
  const month = date.getMonth() + 1; // getMonth() returns 0-based month
  const day = date.getDate();
  return `${month}/${day}`;
}

/**
 * Build availability columns in the specified availability sheet (shared function for practice/game)
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {Array} dates - Array of date objects (practice or game dates)
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 * @return {Object} Result object with statistics
 */
function buildAvailabilityColumns(ss, dates, config) {
  let availabilitySheet = ss.getSheetByName(config.availabilitySheet);
  
  // Create the sheet if it doesn't exist
  if (!availabilitySheet) {
    console.log(`📋 Creating new "${config.availabilitySheet}" sheet`);
    availabilitySheet = ss.insertSheet(config.availabilitySheet);
    
    // Set up basic structure with Full Name column
    availabilitySheet.getRange(1, 1).setValue('Full Name');
    availabilitySheet.getRange(1, 1).setFontWeight('bold');
  }
  
  console.log(`📊 Building availability columns in "${config.availabilitySheet}"`);
  
  // Get existing columns to check what already exists
  const existingColumns = getExistingColumns(availabilitySheet);
  console.log(`📋 Found ${Object.keys(existingColumns).length} existing columns`);
  
  const columnsCreated = [];
  const columnsSkipped = [];
  let validationRanges = []; // Track ranges that need availability dropdown
  let activationStatusColumns = []; // Game only: columns that need Activation Status dropdown + colors
  let noteColumns = []; // Game only: Note columns (free text — clear any validation)

  // Find where to start adding new columns (after existing columns)
  let nextColumnIndex = Math.max(2, availabilitySheet.getLastColumn() + 1);

  const isGame = config.type === 'game' && config.columnsPerDate;

  // Process each date (practice or game)
  dates.forEach((dateInfo, index) => {
    const dateString = dateInfo.formattedDate;

    if (isGame) {
      // Game: three columns per date — $date Availability, $date Activation Status, $date Note (free text)
      config.columnsPerDate.forEach((colDef) => {
        const header = dateString + colDef.suffix;
        const existingCol = existingColumns[header];

        if (existingCol) {
          console.log(`⏭️ Column "${header}" already exists at column ${existingCol}`);
          columnsSkipped.push(header);
          if (colDef.useAvailabilityValidation) {
            const existingValidation = availabilitySheet.getRange(2, existingCol, 1, 1).getDataValidation();
            if (!existingValidation) {
              validationRanges.push({ column: existingCol, header: header, isExisting: true });
            }
          } else if (colDef.useActivationStatusValidation) {
            activationStatusColumns.push({ column: existingCol, header: header });
          } else if (colDef.isFreeText) {
            noteColumns.push(existingCol);
          }
          return;
        }

        console.log(`➕ Creating new column "${header}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(header);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        columnsCreated.push(header);

        if (colDef.useAvailabilityValidation) {
          validationRanges.push({ column: nextColumnIndex, header: header, isExisting: false });
        } else if (colDef.useActivationStatusValidation) {
          activationStatusColumns.push({ column: nextColumnIndex, header: header });
        } else if (colDef.isFreeText) {
          noteColumns.push(nextColumnIndex);
        }
        nextColumnIndex++;
      });
    } else {
      // Practice: two columns per date — dateString (availability), "$Date Note"
      const availabilityHeader = dateString;
      const notesHeader = dateString + ' Note';

      if (existingColumns[availabilityHeader]) {
        console.log(`⏭️ Column "${availabilityHeader}" already exists at column ${existingColumns[availabilityHeader]}`);
        columnsSkipped.push(availabilityHeader);
        const existingColumnIndex = existingColumns[availabilityHeader];
        const existingValidation = availabilitySheet.getRange(2, existingColumnIndex, 1, 1).getDataValidation();
        if (!existingValidation) {
          validationRanges.push({ column: existingColumnIndex, header: availabilityHeader, isExisting: true });
        }
      } else {
        console.log(`➕ Creating new column "${availabilityHeader}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(availabilityHeader);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        validationRanges.push({ column: nextColumnIndex, header: availabilityHeader, isExisting: false });
        columnsCreated.push(availabilityHeader);
        nextColumnIndex++;
      }

      if (existingColumns[notesHeader]) {
        console.log(`⏭️ Column "${notesHeader}" already exists at column ${existingColumns[notesHeader]}`);
        columnsSkipped.push(notesHeader);
      } else {
        console.log(`➕ Creating new column "${notesHeader}" at column ${nextColumnIndex}`);
        availabilitySheet.getRange(1, nextColumnIndex).setValue(notesHeader);
        availabilitySheet.getRange(1, nextColumnIndex).setFontWeight('bold');
        columnsCreated.push(notesHeader);
        nextColumnIndex++;
      }
    }
  });

  // Apply or extend data validation to availability columns
  extendOrCreateDataValidation(availabilitySheet, validationRanges, config);

  // Game only: apply Activation Status dropdown + conditional formatting (Active=green, Inactive=red, TBD=grey)
  if (activationStatusColumns.length > 0) {
    applyActivationStatusValidationAndFormatting(availabilitySheet, activationStatusColumns);
  }

  // Game only: ensure Note columns are free text (no dropdown)
  noteColumns.forEach(function (col) {
    // getRange(row, column, numRows, numColumns) — 1 column only
    availabilitySheet.getRange(2, col, 999, 1).clearDataValidations();
  });

  // Enable text wrapping on the sheet so headers and cells wrap
  const lastCol = availabilitySheet.getLastColumn();
  const lastRow = Math.max(availabilitySheet.getLastRow(), 2);
  if (lastCol >= 1 && lastRow >= 1) {
    availabilitySheet.getRange(1, 1, lastRow, lastCol).setWrap(true);
  }

  // Apply Format Spruce Up silently (no modal)
  console.log('✨ Applying Format Spruce Up formatting...');
  try {
    applySpruceUpFormatting(availabilitySheet);
  } catch (error) {
    console.warn('⚠️ Could not apply Format Spruce Up formatting:', error.message);
  }

  return {
    columnsCreated: columnsCreated.length,
    columnsSkipped: columnsSkipped.length,
    columnSummary: columnsCreated.length > 0 ? 
      `Created: ${columnsCreated.join(', ')}` : 'No new columns created',
    skippedSummary: columnsSkipped.length > 0 ? 
      `Skipped existing: ${columnsSkipped.join(', ')}` : ''
  };
}

/**
 * Get a map of existing column headers to their column indices
 * @param {Sheet} sheet - The Practice Availability sheet
 * @return {Object} Map of header names to column indices
 */
function getExistingColumns(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return {};
  
  const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const columnMap = {};
  
  console.log('🔍 Existing columns found:');
  headerRow.forEach((header, index) => {
    if (header && header.toString().trim() !== '') {
      const cleanHeader = header.toString().trim();
      columnMap[cleanHeader] = index + 1; // 1-based column index
      console.log(`  Column ${index + 1}: "${cleanHeader}"`);
      
      // Also map date objects to formatted date strings for comparison
      if (header instanceof Date) {
        const formattedDate = formatDateForColumn(header);
        columnMap[formattedDate] = index + 1;
        console.log(`  Column ${index + 1} also mapped as: "${formattedDate}"`);
      }
    }
  });
  
  return columnMap;
}

/**
 * Apply or extend data validation to availability columns
 * @param {Sheet} sheet - The availability sheet (practice or game)
 * @param {Array} validationRanges - Array of column info for validation
 * @param {Object} config - Configuration object (PRACTICE_AVAILABILITY_CONFIG or GAME_AVAILABILITY_CONFIG)
 */
function extendOrCreateDataValidation(sheet, validationRanges, config) {
  console.log('🎯 Applying or extending data validation to availability columns...');
  console.log(`📊 Processing ${validationRanges.length} validation ranges`);
  
  // Create validation options from config
  const validationValues = config.validationOptions.map(option => option.value);
  console.log(`🎯 Expected validation values: [${validationValues.join(', ')}]`);
  
  // Check if there's an existing data validation rule we can extend
  const existingValidation = findExistingDataValidation(sheet, validationValues);
  
  if (existingValidation) {
    console.log(`✅ Found existing compatible validation rule in column ${existingValidation.column}`);
  } else {
    console.log(`❌ No existing compatible validation rule found`);
  }
  
  validationRanges.forEach((rangeInfo, index) => {
    console.log(`\n📋 Processing range ${index + 1}/${validationRanges.length}:`);
    console.log(`   Column: ${rangeInfo.column}, Header: "${rangeInfo.header}", IsExisting: ${rangeInfo.isExisting}`);
    
    try {
      if (rangeInfo.isExisting && existingValidation) {
        // Extend existing validation rule
        console.log(`🔄 Extending existing data validation for column ${rangeInfo.column} (${rangeInfo.header})`);
        extendExistingValidation(sheet, rangeInfo.column, existingValidation);
      } else {
        // Create new validation rule
        const reason = !rangeInfo.isExisting ? 'new column' : 'no compatible existing validation';
        console.log(`➕ Creating new data validation for column ${rangeInfo.column} (${rangeInfo.header}) - ${reason}`);
        createNewValidation(sheet, rangeInfo.column, validationValues);
      }
      
    } catch (error) {
      console.warn(`⚠️ Could not apply data validation to column ${rangeInfo.column}: ${error.message}`);
    }
  });
}

/**
 * Find existing data validation rule that matches our requirements
 * @param {Sheet} sheet - The Practice Availability sheet
 * @param {Array} expectedValues - Expected validation values
 * @return {Object|null} Existing validation info or null
 */
function findExistingDataValidation(sheet, expectedValues) {
  console.log('🔍 Searching for existing data validation rules...');
  console.log(`   Checking columns 2 to ${Math.min(10, sheet.getLastColumn())}`);
  
  try {
    // Check a few columns for existing validation rules
    for (let col = 2; col <= Math.min(10, sheet.getLastColumn()); col++) {
      console.log(`   📋 Checking column ${col} for existing validation...`);
      
      const range = sheet.getRange(2, col, 1, 1);
      const validation = range.getDataValidation();
      
      if (validation) {
        console.log(`      ✅ Found validation rule in column ${col}`);
        const criteria = validation.getCriteriaType();
        console.log(`      📊 Criteria type: ${criteria}`);
        
        if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          console.log(`      📝 It's a VALUE_IN_LIST validation`);
          const values = validation.getCriteriaValues()[0];
          console.log(`      🎯 Existing values: [${values ? values.join(', ') : 'null'}]`);
          console.log(`      🎯 Expected values: [${expectedValues.join(', ')}]`);
          
          if (values) {
            const existingSorted = values.slice().sort();
            const expectedSorted = expectedValues.slice().sort();
            console.log(`      🔄 Sorted existing: [${existingSorted.join(', ')}]`);
            console.log(`      🔄 Sorted expected: [${expectedSorted.join(', ')}]`);
            
            if (arraysEqual(existingSorted, expectedSorted)) {
              console.log(`      ✅ Arrays match! Found compatible existing validation rule in column ${col}`);
              return {
                validation: validation,
                column: col
              };
            } else {
              console.log(`      ❌ Arrays don't match`);
            }
          } else {
            console.log(`      ❌ No values found in validation rule`);
          }
        } else {
          console.log(`      ❌ Not a VALUE_IN_LIST validation`);
        }
      } else {
        console.log(`      ❌ No validation rule found in column ${col}`);
      }
    }
  } catch (error) {
    console.warn('⚠️ Error checking for existing validation:', error.message);
  }
  
  console.log('❌ No compatible existing validation rule found');
  return null;
}

/**
 * Extend existing validation to a new column (preserves custom colors)
 * @param {Sheet} sheet - The Practice Availability sheet
 * @param {number} column - Column to apply validation to
 * @param {Object} existingValidation - Existing validation info
 */
function extendExistingValidation(sheet, column, existingValidation) {
  console.log(`🎨 Extending validation to column ${column} while preserving any custom colors`);
  const validationRange = sheet.getRange(2, column, 1000, 1);
  
  // Create a new validation rule that matches the existing one
  // This preserves any custom conditional formatting/colors that may exist
  const originalValidation = existingValidation.validation;
  const criteriaValues = originalValidation.getCriteriaValues()[0];
  
  const newValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(criteriaValues, true)
    .setAllowInvalid(originalValidation.getAllowInvalid())
    .setHelpText(originalValidation.getHelpText() || 'Select your availability')
    .build();
    
  validationRange.setDataValidation(newValidation);
  console.log(`✅ Validation extended to column ${column} without overriding colors`);
}

/**
 * Create new data validation rule (without setting colors - preserves conditional formatting)
 * @param {Sheet} sheet - The Practice Availability sheet
 * @param {number} column - Column to apply validation to
 * @param {Array} validationValues - Values for validation
 */
function createNewValidation(sheet, column, validationValues) {
  console.log(`🎨 Creating new validation for column ${column} without setting colors`);
  const validationRange = sheet.getRange(2, column, 1000, 1);
  
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(validationValues, true) // true = show dropdown
    .setAllowInvalid(false)
    .setHelpText('Select your availability')
    .build();
  
  validationRange.setDataValidation(validation);
  console.log(`✅ New validation created for column ${column} - colors can be set via conditional formatting`);
}

/**
 * Apply Activation Status dropdown (Active / Inactive / TBD) and background colors to game columns.
 * @param {Sheet} sheet - Game Availability sheet
 * @param {Array<{column: number, header: string}>} activationStatusColumns - Column indices and headers
 */
function applyActivationStatusValidationAndFormatting(sheet, activationStatusColumns) {
  const values = GAME_ACTIVATION_STATUS_OPTIONS.map(function (o) { return o.value; });
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const numRows = lastRow - 1; // rows 2 through lastRow inclusive

  activationStatusColumns.forEach(function (info) {
    const col = info.column;
    // getRange(row, column, numRows, numColumns) — use 1 column so we only affect this Activation Status column
    const range = sheet.getRange(2, col, numRows, 1);

    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .setHelpText('Active / Inactive / TBD')
      .build();
    range.setDataValidation(validation);

    const rules = sheet.getConditionalFormatRules();
    GAME_ACTIVATION_STATUS_OPTIONS.forEach(function (opt) {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(opt.value)
        .setBackground(opt.backgroundColor)
        .setRanges([range])
        .build();
      rules.push(rule);
    });
    sheet.setConditionalFormatRules(rules);
  });
}

/**
 * Check if two arrays are equal (order-independent)
 * @param {Array} arr1 - First array
 * @param {Array} arr2 - Second array
 * @return {boolean} True if arrays contain same elements
 */
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }
  
  return true;
}

