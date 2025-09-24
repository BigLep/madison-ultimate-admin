/**
 * Convert to Attendance Module
 * Converts availability responses to actual attendance records
 */

/**
 * Main function to convert selected cells from availability to attendance
 * Called from the menu
 * 
 * Mapping:
 * - "👎 Can't make it" → "Wasn't there"
 * - "👍 Planning to be there" → "Was there"
 * - "❓ Not sure yet" → unchanged
 * - "Was there" and "Wasn't there" → unchanged
 * - Empty cells → "Was there"
 * - Any other values → unchanged
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
    
    const newValues = values.map(row => {
      return row.map(cell => {
        const cellValue = cell.toString().trim();
        
        if (cellValue === '') {
          convertedCount++;
          return AVAILABILITY_VALIDATION_OPTIONS[3].value;
        }
        
        if (cellValue === AVAILABILITY_VALIDATION_OPTIONS[1].value) {
          convertedCount++;
          return AVAILABILITY_VALIDATION_OPTIONS[4].value;
        }
        
        if (cellValue === AVAILABILITY_VALIDATION_OPTIONS[0].value) {
          convertedCount++;
          return AVAILABILITY_VALIDATION_OPTIONS[3].value;
        }
        
        return cell;
      });
    });
    
    selection.setValues(newValues);
    
    console.log(`✅ Converted ${convertedCount} cells to actual attendance`);
    
    const planningToBeThere = AVAILABILITY_VALIDATION_OPTIONS[0].value;
    const cantMakeIt = AVAILABILITY_VALIDATION_OPTIONS[1].value;
    const notSureYet = AVAILABILITY_VALIDATION_OPTIONS[2].value;
    const wasThere = AVAILABILITY_VALIDATION_OPTIONS[3].value;
    const wasntThere = AVAILABILITY_VALIDATION_OPTIONS[4].value;
    
    SpreadsheetApp.getUi().alert(
      'Conversion Complete!',
      `Successfully converted ${convertedCount} cell${convertedCount !== 1 ? 's' : ''} to actual attendance.\n\n` +
      `Mapping applied:\n` +
      `• "${cantMakeIt}" → "${wasntThere}"\n` +
      `• "${planningToBeThere}" → "${wasThere}"\n` +
      `• Empty cells → "${wasThere}"\n` +
      `• "${notSureYet}" → unchanged\n` +
      `• "${wasThere}" / "${wasntThere}" → unchanged\n` +
      `• Other values → unchanged`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error converting to actual attendance:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to convert to actual attendance: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}