/**
 * Shared whole-sheet conditional formatting for availability / activation values.
 * Used by Build Game/Practice Availability, Build Game/Practice Roster Prep, and any sheet
 * whose row 1 includes Practice-style date columns (M/D) and/or Game Availability-style headers.
 */

// Same values and colors as Practice & Game Availability builders (single source for CF + DV labels elsewhere)
var AVAILABILITY_VALIDATION_OPTIONS = [
  { value: '👍 Planning to be there', backgroundColor: '#d5e8d4' },
  { value: '👎 Can\'t make it', backgroundColor: '#f4c7c3' },
  { value: '❓ Not sure yet', backgroundColor: '#fce5cd' },
  { value: 'Was there', backgroundColor: '#38761d' },
  { value: 'Wasn\'t there', backgroundColor: '#cc0000' }
];

var GAME_ACTIVATION_STATUS_OPTIONS = [
  { value: 'Active', backgroundColor: '#38761d' },
  { value: 'Inactive', backgroundColor: '#cc0000' },
  { value: 'TBD', backgroundColor: '#e8eaed' }
];

/**
 * Normalize header cell for pattern matching (handles Date-typed headers in row 1).
 * @param {*} cell
 * @return {string}
 */
function normalizeAvailabilityHeaderForCf_(cell) {
  if (cell == null || cell === '') return '';
  if (cell instanceof Date && !isNaN(cell.getTime())) {
    return cell.getMonth() + 1 + '/' + cell.getDate();
  }
  return String(cell).trim();
}

/** @return {Object.<string, boolean>} */
function managedAvailabilityCfTextSet_() {
  var o = {};
  AVAILABILITY_VALIDATION_OPTIONS.forEach(function (x) {
    o[x.value] = true;
  });
  GAME_ACTIVATION_STATUS_OPTIONS.forEach(function (x) {
    o[x.value] = true;
  });
  return o;
}

/**
 * Drop CF rules we manage (text-equals for availability or activation values) so rebuilds do not stack.
 * @param {GoogleAppsScript.Spreadsheet.ConditionalFormatRule[]} rules
 * @param {Object.<string, boolean>} valueSet
 * @return {GoogleAppsScript.Spreadsheet.ConditionalFormatRule[]}
 */
function removeManagedTextEqualsCfRules_(rules, valueSet) {
  return rules.filter(function (rule) {
    try {
      var bc = rule.getBooleanCondition();
      if (!bc) return true;
      var vals = bc.getCriteriaValues();
      if (!vals || vals.length === 0) return true;
      var t = vals[0];
      if (typeof t === 'string' && valueSet[t]) return false;
      return true;
    } catch (err) {
      return true;
    }
  });
}

/**
 * @return {{ availabilityCols: number[], activationCols: number[] }}
 */
function collectGameAvailabilityColumnIndices_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { availabilityCols: [], activationCols: [] };
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var reAvail = /^\d{1,2}\/\d{1,2} Availability(?: \(Game \d+\))?$/;
  var reAct = /^\d{1,2}\/\d{1,2} Activation Status(?: \(Game \d+\))?$/;
  var availabilityCols = [];
  var activationCols = [];
  for (var i = 0; i < headers.length; i++) {
    var s = normalizeAvailabilityHeaderForCf_(headers[i]);
    if (reAvail.test(s)) availabilityCols.push(i + 1);
    else if (reAct.test(s)) activationCols.push(i + 1);
  }
  return { availabilityCols: availabilityCols, activationCols: activationCols };
}

/**
 * Practice availability columns are headers that are exactly M/D (not "M/D Note").
 * @return {number[]}
 */
function collectPracticeAvailabilityColumnIndices_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var reDateOnly = /^\d{1,2}\/\d{1,2}$/;
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    var s = normalizeAvailabilityHeaderForCf_(headers[i]);
    if (reDateOnly.test(s)) out.push(i + 1);
  }
  return out;
}

/**
 * Entire sheet grid — used for conditional formatting so we do not maintain per-column range lists.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {GoogleAppsScript.Spreadsheet.Range}
 */
function getWholeSheetRangeForCf_(sheet) {
  return sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
}

/**
 * @param {{ value: string, backgroundColor: string }} opt
 * @param {GoogleAppsScript.Spreadsheet.Range} whole
 * @return {GoogleAppsScript.Spreadsheet.ConditionalFormatRule}
 */
function buildTextEqualsWholeSheetCfRule_(opt, whole) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(opt.value)
    .setBackground(opt.backgroundColor)
    .setRanges([whole])
    .build();
}

/**
 * One rule per availability value and one per activation value; each applies to the whole sheet.
 * Adds availability coloring if the sheet has any practice-style M/D column and/or game "M/D Availability" column.
 * Adds activation coloring if the sheet has any "M/D Activation Status" column.
 * Strips prior managed rules (same value set) so rules do not stack.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function refreshManagedAvailabilityAndActivationCfOnSheet(sheet) {
  var gameIdx = collectGameAvailabilityColumnIndices_(sheet);
  var practiceCols = collectPracticeAvailabilityColumnIndices_(sheet);
  var hasAvailabilityTargets = gameIdx.availabilityCols.length > 0 || practiceCols.length > 0;
  var hasActivationTargets = gameIdx.activationCols.length > 0;

  var managed = managedAvailabilityCfTextSet_();
  var rules = removeManagedTextEqualsCfRules_(sheet.getConditionalFormatRules(), managed);
  var whole = getWholeSheetRangeForCf_(sheet);

  AVAILABILITY_VALIDATION_OPTIONS.forEach(function (opt) {
    if (!hasAvailabilityTargets) return;
    rules.push(buildTextEqualsWholeSheetCfRule_(opt, whole));
  });

  GAME_ACTIVATION_STATUS_OPTIONS.forEach(function (opt) {
    if (!hasActivationTargets) return;
    rules.push(buildTextEqualsWholeSheetCfRule_(opt, whole));
  });

  sheet.setConditionalFormatRules(rules);
}
