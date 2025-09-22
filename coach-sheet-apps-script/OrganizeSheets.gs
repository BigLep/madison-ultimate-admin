/**
 * Organize Sheets Module
 * Provides drag-and-drop interface for reordering sheets and controlling visibility
 */

/**
 * Main function to show the Organize Sheets dialog
 * Called from the menu
 */
function organizeSheets() {
  console.log('📋 Starting Organize Sheets...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    if (sheets.length === 0) {
      SpreadsheetApp.getUi().alert('No Sheets Found', 'No sheets found in this spreadsheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get sheet information
    const sheetInfo = sheets.map((sheet, index) => ({
      name: sheet.getName(),
      index: index,
      isHidden: sheet.isSheetHidden(),
      sheetId: sheet.getSheetId()
    }));
    
    console.log(`Found ${sheetInfo.length} sheets`);
    
    // Show the organize dialog
    showOrganizeSheetsDialog(sheetInfo);
    
  } catch (error) {
    console.error('Error in organizeSheets:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to load sheet organizer: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show the HTML dialog for organizing sheets
 * @param {Array} sheetInfo - Array of sheet information objects
 */
function showOrganizeSheetsDialog(sheetInfo) {
  const html = createOrganizeSheetsHtml(sheetInfo);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600);
    
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Organize Sheets');
}

/**
 * Create the HTML content for the organize sheets dialog
 * @param {Array} sheetInfo - Array of sheet information objects
 * @return {string} HTML content
 */
function createOrganizeSheetsHtml(sheetInfo) {
  // Create the sheet items HTML
  const sheetItems = sheetInfo.map((sheet, index) => `
    <div class="sheet-item" data-sheet-name="${sheet.name}" data-original-index="${index}">
      <div class="drag-handle">⋮⋮</div>
      <div class="sheet-info">
        <div class="sheet-name">${sheet.name}</div>
      </div>
      <div class="visibility-control">
        <label class="checkbox-label">
          <input type="checkbox" ${sheet.isHidden ? '' : 'checked'}>
          Visible
        </label>
      </div>
    </div>
  `).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>Organize Sheets</title>
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f8f9fa;
          }
          .container {
            max-width: 100%;
            background: white;
            border-radius: 8px;
            padding: 16px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-height: 90vh;
            overflow-y: auto;
          }
          .header {
            text-align: center;
            margin-bottom: 16px;
          }
          .header h2 {
            margin: 0 0 8px 0;
            color: #1a73e8;
            font-size: 22px;
          }
          .instructions {
            background-color: #e8f0fe;
            border-left: 4px solid #1a73e8;
            padding: 10px 14px;
            margin-bottom: 16px;
            font-size: 13px;
            line-height: 1.3;
          }
          .sheets-container {
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            background: #fafafa;
            margin-bottom: 16px;
          }
          .sheet-item {
            display: flex;
            align-items: center;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            margin: 6px;
            padding: 8px 12px;
            cursor: move;
            transition: all 0.2s ease;
            min-height: 36px;
          }
          .sheet-item:hover {
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            border-color: #1a73e8;
          }
          .sheet-item.dragging {
            opacity: 0.5;
            transform: rotate(2deg);
          }
          .sheet-item.drag-over {
            border-top: 3px solid #1a73e8;
            background-color: #e8f0fe;
          }
          .drag-handle {
            font-size: 14px;
            color: #9aa0a6;
            margin-right: 10px;
            cursor: grab;
            user-select: none;
            line-height: 1;
          }
          .drag-handle:active {
            cursor: grabbing;
          }
          .sheet-info {
            flex: 1;
            min-width: 0;
          }
          .sheet-name {
            font-weight: 500;
            font-size: 15px;
            color: #202124;
            word-break: break-word;
          }
          .visibility-control {
            margin-left: 10px;
          }
          .checkbox-label {
            display: flex;
            align-items: center;
            cursor: pointer;
            font-size: 13px;
            color: #202124;
          }
          .checkbox-label input[type="checkbox"] {
            margin: 0 6px 0 0;
            width: 14px;
            height: 14px;
          }
          .buttons {
            display: flex;
            justify-content: center;
            gap: 12px;
            margin-top: 16px;
            padding-top: 16px;
            border-top: 1px solid #e0e0e0;
          }
          .btn {
            padding: 10px 24px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            min-width: 120px;
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
          .progress-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.9);
            z-index: 1000;
          }
          .progress-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            padding: 30px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <div class="container">
          
          <div class="instructions">
            <strong>How to use:</strong><br>
            • Drag sheets up/down to reorder them<br>
            • Check/uncheck "Visible" to show/hide sheets<br>
            • Click "Apply Changes" to save your organization
          </div>
          
          <div class="sheets-container" id="sheetsContainer">
            ${sheetItems}
          </div>
          
          <div class="buttons">
            <button class="btn btn-primary" onclick="applyChanges()">Apply Changes</button>
            <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>
        </div>
        
        <!-- Progress Overlay -->
        <div class="progress-overlay" id="progressOverlay">
          <div class="progress-content">
            <div class="spinner"></div>
            <div style="font-size: 16px; font-weight: bold; color: #333;">
              Organizing Sheets...
            </div>
            <div style="font-size: 14px; color: #666; margin-top: 8px;">
              Please wait while we apply your changes
            </div>
          </div>
        </div>
        
        <script>
          let draggedElement = null;
          
          // Initialize drag and drop functionality
          document.addEventListener('DOMContentLoaded', function() {
            initializeDragAndDrop();
          });
          
          function initializeDragAndDrop() {
            const sheetItems = document.querySelectorAll('.sheet-item');
            
            sheetItems.forEach(item => {
              item.draggable = true;
              
              item.addEventListener('dragstart', function(e) {
                draggedElement = this;
                this.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
              });
              
              item.addEventListener('dragend', function(e) {
                this.classList.remove('dragging');
                draggedElement = null;
                // Remove drag-over class from all items
                document.querySelectorAll('.sheet-item').forEach(item => {
                  item.classList.remove('drag-over');
                });
              });
              
              item.addEventListener('dragover', function(e) {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                
                if (this !== draggedElement) {
                  this.classList.add('drag-over');
                }
              });
              
              item.addEventListener('dragleave', function(e) {
                this.classList.remove('drag-over');
              });
              
              item.addEventListener('drop', function(e) {
                e.preventDefault();
                this.classList.remove('drag-over');
                
                if (this !== draggedElement) {
                  const container = document.getElementById('sheetsContainer');
                  const afterElement = getDragAfterElement(container, e.clientY);
                  
                  if (afterElement == null) {
                    container.appendChild(draggedElement);
                  } else {
                    container.insertBefore(draggedElement, afterElement);
                  }
                }
              });
            });
          }
          
          function getDragAfterElement(container, y) {
            const draggableElements = [...container.querySelectorAll('.sheet-item:not(.dragging)')];
            
            return draggableElements.reduce((closest, child) => {
              const box = child.getBoundingClientRect();
              const offset = y - box.top - box.height / 2;
              
              if (offset < 0 && offset > closest.offset) {
                return { offset: offset, element: child };
              } else {
                return closest;
              }
            }, { offset: Number.NEGATIVE_INFINITY }).element;
          }
          
          function applyChanges() {
            // Get the current order and visibility
            const sheetItems = document.querySelectorAll('.sheet-item');
            const sheetOrder = [];
            
            sheetItems.forEach(item => {
              const sheetName = item.getAttribute('data-sheet-name');
              const checkbox = item.querySelector('input[type="checkbox"]');
              const isVisible = checkbox.checked;
              
              sheetOrder.push({
                name: sheetName,
                visible: isVisible
              });
            });
            
            if (sheetOrder.length === 0) {
              alert('No sheets to organize');
              return;
            }
            
            // Show progress
            document.getElementById('progressOverlay').style.display = 'block';
            
            // Call server-side function
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .applySheetsOrganization(sheetOrder);
          }
          
          function onSuccess(message) {
            document.getElementById('progressOverlay').style.display = 'none';
            google.script.host.close();
          }
          
          function onFailure(error) {
            document.getElementById('progressOverlay').style.display = 'none';
            alert('Error organizing sheets: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Apply the sheet organization changes (reorder and visibility)
 * @param {Array} sheetOrder - Array of objects with name and visible properties
 */
function applySheetsOrganization(sheetOrder) {
  console.log('🔄 Applying sheet organization changes...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // Get original sheet information for comparison
    const originalOrder = sheets.map((sheet, index) => ({
      name: sheet.getName(),
      originalIndex: index,
      isHidden: sheet.isSheetHidden()
    }));
    
    console.log('📋 Original sheet order:', originalOrder.map(s => s.name).join(', '));
    console.log('📋 Target sheet order:', sheetOrder.map(s => s.name).join(', '));
    
    let visibilityChanges = 0;
    let reorderMoves = 0;
    
    // First, set visibility for sheets that changed
    sheetOrder.forEach((sheetInfo, index) => {
      const originalSheet = originalOrder.find(s => s.name === sheetInfo.name);
      const sheet = ss.getSheetByName(sheetInfo.name);
      
      if (sheet && originalSheet) {
        const shouldBeVisible = sheetInfo.visible;
        const currentlyVisible = !originalSheet.isHidden;
        
        if (shouldBeVisible !== currentlyVisible) {
          if (shouldBeVisible) {
            sheet.showSheet();
            console.log(`👁️ Showing sheet: ${sheetInfo.name}`);
          } else {
            sheet.hideSheet();
            console.log(`🙈 Hiding sheet: ${sheetInfo.name}`);
          }
          visibilityChanges++;
        }
      }
    });
    
    // Smart reordering - only move sheets that changed position
    // Create a mapping of where each sheet should be
    const targetPositions = {};
    sheetOrder.forEach((sheetInfo, index) => {
      targetPositions[sheetInfo.name] = index;
    });
    
    // Find sheets that actually need to move
    const sheetsToMove = [];
    originalOrder.forEach(originalSheet => {
      const targetIndex = targetPositions[originalSheet.name];
      if (targetIndex !== undefined && targetIndex !== originalSheet.originalIndex) {
        sheetsToMove.push({
          name: originalSheet.name,
          currentIndex: originalSheet.originalIndex,
          targetIndex: targetIndex
        });
      }
    });
    
    console.log(`📊 Found ${sheetsToMove.length} sheets that need to move`);
    
    // Sort by target position to avoid conflicts
    sheetsToMove.sort((a, b) => a.targetIndex - b.targetIndex);
    
    // Apply moves one by one, recalculating positions as we go
    sheetsToMove.forEach(moveInfo => {
      const sheet = ss.getSheetByName(moveInfo.name);
      if (sheet) {
        const currentPosition = sheet.getIndex(); // Get current position (1-based)
        const targetPosition = moveInfo.targetIndex + 1; // Convert to 1-based
        
        if (currentPosition !== targetPosition) {
          ss.setActiveSheet(sheet);
          ss.moveActiveSheet(targetPosition);
          console.log(`📋 Moved "${moveInfo.name}" from position ${currentPosition} to ${targetPosition}`);
          reorderMoves++;
        }
      }
    });
    
    console.log(`✅ Sheet organization complete: ${visibilityChanges} visibility changes, ${reorderMoves} position changes`);
    return `Sheets organized successfully! ${visibilityChanges} visibility changes, ${reorderMoves} moves applied.`;
    
  } catch (error) {
    console.error('Error applying sheet organization:', error);
    throw new Error(`Failed to organize sheets: ${error.message}`);
  }
}