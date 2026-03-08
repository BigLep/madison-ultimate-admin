/**
 * Convert to Attendance Module
 * Converts availability responses to actual attendance records
 */

/** Activation Status value that counts as "active" for two-column conversion (game roster). */
const ACTIVATION_STATUS_ACTIVE = 'Active';

/**
 * Main function to convert selected cells from availability to attendance
 * Called from the menu
 *
 * One column selected (Availability only):
 * - "👎 Can't make it" → "Wasn't there"
 * - "👍 Planning to be there" → "Was there"
 * - Empty cells → "Was there"
 * - "❓ Not sure yet", "Was there", "Wasn't there", other → unchanged
 *
 * Two columns selected (Availability + Activation Status):
 * - Only mark "Was there" when Availability is "👍 Planning to be there" AND Activation Status is "Active".
 * - "👎 Can't make it" (in first column) → "Wasn't there"
 * - Otherwise first column unchanged. Second column is never written.
 */
function convertToActualAttendance() {
  console.log('🔄 Starting Convert to Actual Attendance...');

  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const selection = sheet.getActiveRange();

    if (!selection) {
      SpreadsheetApp.getUi().alert(
        'No Selection',
        'Please select the cells you want to convert to actual attendance.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    console.log(`📍 Selected range: ${selection.getA1Notation()}`);

    const values = selection.getValues();
    const numRows = values.length;
    const numCols = values[0].length;
    let convertedCount = 0;

    const planningToBeThere = AVAILABILITY_VALIDATION_OPTIONS[0].value;
    const cantMakeIt = AVAILABILITY_VALIDATION_OPTIONS[1].value;
    const wasThere = AVAILABILITY_VALIDATION_OPTIONS[3].value;
    const wasntThere = AVAILABILITY_VALIDATION_OPTIONS[4].value;

    let newValues;
    let messageLines;

    if (numCols === 2) {
      // Two columns: Availability (col 0) + Activation Status (col 1). Only write to col 0. "Was there" only if Planning + Active.
      newValues = values.map(function (row) {
        const avail = row[0] != null ? String(row[0]).trim() : '';
        const activation = row[1] != null ? String(row[1]).trim() : '';
        let newFirst = row[0];
        if (avail === cantMakeIt) {
          convertedCount++;
          newFirst = wasntThere;
        } else if (avail === planningToBeThere && activation === ACTIVATION_STATUS_ACTIVE) {
          convertedCount++;
          newFirst = wasThere;
        }
        return [newFirst, row[1]];
      });
      messageLines = [
        `• "${cantMakeIt}" → "${wasntThere}"`,
        `• "${planningToBeThere}" + "${ACTIVATION_STATUS_ACTIVE}" → "${wasThere}" (only when both)`,
        '• Otherwise first column unchanged. Second column (Activation Status) never modified.'
      ];
    } else {
      // One column: original per-cell logic
      newValues = values.map(function (row) {
        return row.map(function (cell) {
          const cellValue = cell != null ? String(cell).trim() : '';
          if (cellValue === '') {
            convertedCount++;
            return wasThere;
          }
          if (cellValue === cantMakeIt) {
            convertedCount++;
            return wasntThere;
          }
          if (cellValue === planningToBeThere) {
            convertedCount++;
            return wasThere;
          }
          return cell;
        });
      });
      messageLines = [
        `• "${cantMakeIt}" → "${wasntThere}"`,
        `• "${planningToBeThere}" → "${wasThere}"`,
        '• Empty cells → "' + wasThere + '"',
        '• "❓ Not sure yet" / "Was there" / "Wasn\'t there" / other → unchanged'
      ];
    }

    selection.setValues(newValues);

    console.log(`✅ Converted ${convertedCount} cells to actual attendance`);

    SpreadsheetApp.getUi().alert(
      'Conversion Complete!',
      'Successfully converted ' + convertedCount + ' cell' + (convertedCount !== 1 ? 's' : '') + ' to actual attendance.\n\n' +
      'Mapping applied:\n' + messageLines.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    console.error('Error converting to actual attendance:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to convert to actual attendance: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}