// Offline regression harness for the coach sheet Apps Script. Loads the .gs files into a Node
// vm with a fake Sheet, then asserts the shapes the builds produce: PlayerID-keyed rows, the
// house-style lookup formulas, availability seeding, Team and availability sort order, the
// season switches (Activation Status, Team), and that onOpen builds the menu without throwing.
// Run from anywhere: node coach-sheet-apps-script/test/harness.js. Exit code 1 on any failure.
// Required before every clasp push (see AGENTS.md).
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..');
// Fake sheet: records writes, serves a header row and roster rows.
function FakeSheet(name, rows) {
  this.name = name; this.rows = rows.map(r => r.slice()); this.writes = []; this.hidden = [];
}
FakeSheet.prototype.getName = function () { return this.name; };
FakeSheet.prototype.getLastRow = function () { return this.rows.length; };
FakeSheet.prototype.getLastColumn = function () { return Math.max(...this.rows.map(r => r.length)); };
FakeSheet.prototype.getMaxRows = function () { return 1000; };
FakeSheet.prototype.getMaxColumns = function () { return 26; };
FakeSheet.prototype.hideColumns = function (c) { this.hidden.push(c); };
FakeSheet.prototype.insertColumnAfter = function (c) { this.rows.forEach(r => { while (r.length < c) r.push(''); r.splice(c, 0, ''); }); };
FakeSheet.prototype.deleteColumn = function (c) { this.rows.forEach(r => r.splice(c - 1, 1)); };
FakeSheet.prototype.getRange = function (r, c, nr, nc) {
  const sheet = this; nr = nr || 1; nc = nc || 1;
  return {
    getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (sheet.rows[r - 1 + i] || [])[c - 1 + j] ?? '')),
    getFormulas: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = (sheet.rows[r - 1 + i] || [])[c - 1 + j]; return typeof v === 'string' && v.startsWith('=') ? v : ''; })),
    setValues: (vals) => { vals.forEach((row, i) => row.forEach((v, j) => { while (sheet.rows.length < r + i) sheet.rows.push([]); sheet.rows[r - 1 + i][c - 1 + j] = v; })); sheet.writes.push({ r, c, vals }); return this; },
    setFormulas: function (vals) { return this.setValues(vals); },
    setFormula: function (v) { return this.setValues([[v]]); },
    setValue: function (v) { return this.setValues([[v]]); },
    sort: (spec) => { const block = sheet.rows.slice(r - 1, r - 1 + nr); block.sort((a, b) => { for (const k of spec) { const x = a[k.column - 1] ?? '', y = b[k.column - 1] ?? ''; if (x < y) return -1; if (x > y) return 1; } return 0; }); block.forEach((row, i) => { sheet.rows[r - 1 + i] = row; }); },
    setFontWeight: () => ({}), clearDataValidations: () => ({}), setWrap: () => ({}),
    setBorder: (top, left, bottom, right, vertical, horizontal, color, style) => {
      sheet.borders = sheet.borders || [];
      sheet.borders.push({ r, c, nr, nc, top, left, bottom, right, vertical, horizontal, color, style });
      return this;
    }
  };
};
const src = ['Code.gs', 'ManagedConditionalFormatting.gs', 'SheetBuilderUtils.gs', 'Availability.gs', 'BuildPracticeRoster.gs', 'BuildGameRosterPrepSheet.gs', 'CreatePracticeCalendarEvents.gs', 'CreateGameCalendarEvents.gs', 'GroupBorders.gs']
  .map(f => fs.readFileSync(dir + '/' + f, 'utf8')).join('\n');
const sandbox = { console: { log() {}, warn() {}, error() {} }, SpreadsheetApp: { getUi: () => ({}), flush() {}, BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK', DOTTED: 'DOTTED', DASHED: 'DASHED', DOUBLE: 'DOUBLE' } } };
const vm = require('vm'); vm.createContext(sandbox);
vm.runInContext(src + `
  ;module = { seedPrintoutPlayerRows, populatePracticeRosterData, populateGameRosterPrepData, getGameRosterPrepColumnLayout, findAvailabilityColumns, playerIdLookupFormula, CONFIG, seedAvailabilityRows_, ROSTER_COLUMNS, getGameEventSpecsFromSheet, gameTeamLabel, findGroupBoundaryRows_, defaultGroupByColumns_, groupBorderCellText_, drawGroupBordersOnSheet_, groupPositions_, applyGroupBordersAndNumbering_, numberAndBorderPracticeRosterGroups_, capturePracticeRosterGroupValues_, restorePracticeRosterGroupValues_, readPracticeRosterDate_, PRACTICE_ROSTER_COLUMNS };`, sandbox);
const m = sandbox.module;
m.CONFIG.gameRosterPrep.hasActivationStatus = true;
m.CONFIG.gameRosterPrep.hasTeam = false; // earlier game prep cases were written without a Team column
// Roster: PlayerID A, Full Name F, Grade O, Gender Identification T, Team U, Include In Generated Rosters somewhere.
const hdr = m.ROSTER_COLUMNS.map(c => c.name);
const idx = n => hdr.indexOf(n);
function rosterRow(id, name, grade, gender, team, include) { const r = new Array(hdr.length).fill(''); r[idx('PlayerID')] = id; r[idx('Full Name')] = name; r[idx('Grade')] = grade; r[idx('Gender Identification')] = gender; r[idx('Team')] = team; r[idx('Include In Generated Rosters')] = include; return r; }
const roster = new FakeSheet('📋 Roster', [hdr, rosterRow('e9jpt', 'Ann Able', 7, 'Gx', 'Blue', true), rosterRow('hjkzg', 'Bob Baker', 8, 'Bx', 'Gold', false), rosterRow('vwrk2', 'Cy Cole', 6, 'Bx', 'Blue', true)]);
console.log('Roster letters: PlayerID', String.fromCharCode(65 + idx('PlayerID')), 'Full Name', String.fromCharCode(65 + idx('Full Name')), 'Grade', String.fromCharCode(65 + idx('Grade')), 'Gender Identification', String.fromCharCode(65 + idx('Gender Identification')), 'Team', String.fromCharCode(65 + idx('Team')));
let fails = 0; const eq = (a, b, what) => { if (a !== b) { fails++; console.log('FAIL', what, '\n got:', a, '\n want:', b); } };

// 1. Practice roster
const pa = new FakeSheet('Practice Availability', [['PlayerID', 'Full Name', 'Grade', 'Gender Identification', '9/9', '9/9 Note']]);
const ga = new FakeSheet('Game Availability', [['PlayerID', 'Full Name', 'Grade', 'Gender Identification', '9/13 Availability', '9/13 Activation Status', '9/13 Note']]);
const pr = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note', '9/13 Activation Status', '9/13 Availability', '9/13 Note']]);
const seeded = m.seedPrintoutPlayerRows(pr, roster, 2, 1, 4);
eq(seeded.rowCount, 2, 'include filter drops Bob');
eq(pr.rows[1][0], 'e9jpt', 'PlayerID value in A2'); eq(pr.rows[2][0], 'vwrk2', 'PlayerID value in A3');
eq(pr.rows[2][3], `=IF($A3="","",IFERROR(XLOOKUP($A3,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`, 'Full Name formula row 3');
const availCols = m.findAvailabilityColumns(pa, '9/9', 'Practice Availability');
eq(availCols.playerIdColumn, 'A', 'pa playerIdColumn'); eq(availCols.fullNameColumn, 'B', 'pa fullNameColumn'); eq(availCols.availabilityColumn, 'E', 'pa availabilityColumn');
sandbox.findNextGameAfterPractice = () => null;
m.populatePracticeRosterData(pr, roster, hdr, pa, availCols, 2, ga, { formattedDate: '9/13', ordinalForDate: 1 });
eq(pr.rows[1][4], `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$U:$U),""))`, 'Team formula');
eq(pr.rows[2][6], `=IF($A3="","",IFERROR(XLOOKUP($A3,'📋 Roster'!$A:$A,'📋 Roster'!$O:$O),""))`, 'Grade formula row 3');
eq(pr.rows[1][7], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Practice Availability'!$A:$A,'Practice Availability'!$E:$E),""))`, 'practice availability formula');
eq(pr.rows[1][8], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Practice Availability'!$A:$A,'Practice Availability'!$F:$F),""))`, 'practice note formula');
eq(pr.rows[1][9], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$F:$F),""))`, 'next game activation formula');
eq(pr.rows[1][10], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$E:$E),""))`, 'next game availability formula');
eq(pr.rows[1][11], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$G:$G),""))`, 'next game note formula');

// 2. Older availability sheet without PlayerID: falls back to Full Name join
const oldPa = new FakeSheet('Practice Availability', [['Full Name', 'Grade', 'Gender Identification', '9/9', '9/9 Note']]);
const oldCols = m.findAvailabilityColumns(oldPa, '9/9', 'Practice Availability');
eq(oldCols.playerIdColumn, null, 'old sheet has no playerIdColumn');
const pr2 = new FakeSheet('Practice Roster', [pr.rows[0].slice(0, 9)]);
m.seedPrintoutPlayerRows(pr2, roster, 2, 1, 4);
m.populatePracticeRosterData(pr2, roster, hdr, oldPa, oldCols, 2, null, null);
eq(pr2.rows[1][7], `=IF($D2="","",IFERROR(XLOOKUP($D2,'Practice Availability'!$A:$A,'Practice Availability'!$D:$D),""))`, 'fallback join by Full Name');

// 3. Game roster prep layout + populate (hasTeam false, activation true)
const gCols = [m.findAvailabilityColumns(ga, '9/13', 'Game Availability', 1)];
const layout = m.getGameRosterPrepColumnLayout('9/13', gCols);
eq(layout.headers.join('|'), 'PlayerID|#|Full Name|Gender|Grade|9/13 Activation Status|9/13 Availability|9/13 Note', 'game prep headers');
eq(layout.indices.playerId, 1, 'playerId index'); eq(layout.indices.gender, 4, 'gender index'); eq(layout.indices.games[0].activation, 6, 'activation index');
const gp = new FakeSheet('Game Roster Prep', [layout.headers]);
const gs = m.seedPrintoutPlayerRows(gp, roster, 2, layout.indices.playerId, layout.indices.fullName);
m.populateGameRosterPrepData(gp, roster, hdr, ga, gCols, layout.indices, gs.rowCount);
eq(gp.rows[1][2], `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`, 'game prep Full Name');
eq(gp.rows[1][3], `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$T:$T),""))`, 'game prep Gender');
eq(gp.rows[2][5], `=IF($A3="","",IFERROR(XLOOKUP($A3,'Game Availability'!$A:$A,'Game Availability'!$F:$F),""))`, 'game prep activation row 3');
eq(gp.rows[1][7], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$G:$G),""))`, 'game prep note');

// 4. Availability seeding still emits the same shape via the shared helper
const pa2 = new FakeSheet('Practice Availability', [['PlayerID', 'Full Name', 'Grade', 'Gender Identification', '9/9']]);
const ssStub = { getSheetByName: n => (n === '📋 Roster' ? roster : null) };
sandbox.getExistingColumns = (sheet) => { const o = {}; sheet.rows[0].forEach((h, i) => { if (h) o[h] = i + 1; }); return o; };
const res = vm.runInContext('seedAvailabilityRows_', sandbox)(ssStub, pa2);
eq(res.rowsAdded, 2, 'availability seeds included players');
eq(pa2.rows[1][0], 'e9jpt', 'availability PlayerID value');
eq(pa2.rows[1][1], `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`, 'availability Full Name formula');
eq(pa2.rows[2][3], `=IF($A3="","",IFERROR(XLOOKUP($A3,'📋 Roster'!$A:$A,'📋 Roster'!$T:$T),""))`, 'availability Gender formula row 3');
// 5. "#" column is a plain value (ADR 0005) that resets on Team+Gender changes, via the shared
// applyGroupBordersAndNumbering_ core (numberAndBorderPracticeRosterGroups_ in BuildPracticeRoster.gs)
const np = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade'], ['x', '', '', '', 'Blue', 'Gx', 7], ['y', '', '', '', 'Blue', 'Bx', 8]]);
vm.runInContext('numberAndBorderPracticeRosterGroups_', sandbox)(np, 2);
eq(np.rows[1][1], 1, '# is a plain value, 1 for row 2 (its own group: Blue/Gx)');
eq(np.rows[2][1], 1, '# resets to 1 for row 3 (its own group: Blue/Bx, Gender changed)');
eq(np.borders.filter(b => b.bottom === true).map(b => b.r).join(','), '2,3', 'bottom border on both rows (each row is its own group)');
// 5b. Build Game Roster Prep Sheet's call shape: applyGroupBordersAndNumbering_ with the same
// [team, activationStatus, gender] group-by set BuildGameRosterPrepSheet.gs passes, 0-based indices.
const grp = new FakeSheet('Game Roster Prep', [
  ['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', 'Activation Status', '9/13 Availability', '9/13 Note'],
  ['a', '', '', 'Blue', 'Gx', 7, 'Active', '', ''],
  ['b', '', '', 'Blue', 'Gx', 7, 'Active', '', ''],
  ['c', '', '', 'Blue', 'Gx', 7, 'Inactive', '', '']
]);
const grpColIndices = [3, 6, 4]; // team, activationStatus, gender (0-based), matching idx.team/idx.activationStatus/idx.gender order
const grpResult = m.applyGroupBordersAndNumbering_(grp, 2, 3, grpColIndices, 1, sandbox.SpreadsheetApp.BorderStyle.SOLID, '#000000');
eq(grpResult.groupCount, 2, 'game roster prep grouping: 2 groups (Active, then Inactive)');
eq(grpResult.numbered, true, 'game roster prep grouping: numbered:true when a "#" column and numberColIndex are given');
eq(grp.rows.slice(1).map(r => r[1]).join(','), '1,2,1', '# is a plain value resetting on Team/Activation Status/Gender, matching BuildGameRosterPrepSheet.gs’s call shape');
eq(grp.borders.filter(b => b.bottom === true).map(b => b.r).join(','), '3,4', 'bottom border on the last row of each group (rows 3 and 4)');
// 6. Season without Activation Status: next-game block is two columns, game availability build skips the column
m.CONFIG.gameRosterPrep.hasActivationStatus = false;
const ng = vm.runInContext('practiceRosterNextGameColumns', sandbox)();
eq(ng.activation, null, 'no activation column'); eq(ng.availability, 10, 'next game availability at base+3'); eq(ng.note, 11, 'next game note at base+4');
const pr3 = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note', '9/13 Availability', '9/13 Note']]);
m.seedPrintoutPlayerRows(pr3, roster, 2, 1, 4);
m.populatePracticeRosterData(pr3, roster, hdr, pa, availCols, 2, ga, { formattedDate: '9/13', ordinalForDate: 1 });
eq(pr3.rows[1][9], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$E:$E),""))`, 'next game availability without activation');
eq(pr3.rows[1][10], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$G:$G),""))`, 'next game note without activation');
eq(pr3.rows[1][11], undefined, 'nothing written past the note column');
const layout2 = m.getGameRosterPrepColumnLayout('9/13', gCols);
eq(layout2.headers.join('|'), 'PlayerID|#|Full Name|Gender|Grade|9/13 Availability|9/13 Note', 'game prep headers without activation');
// 7. Team sort order on the practice roster and coach game prep
const mkRow = (id, name, team, gender, grade, avail) => [id, '', '', name, team, gender, grade, avail, ''];
const sp = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note'],
  mkRow('a', 'Ann', 'Practice Squad', 'Gx', 7, ''), mkRow('b', 'Bea', 'TBD', 'Gx', 7, ''), mkRow('c', 'Cal', 'Silver', 'Bx', 8, ''),
  mkRow('d', 'Dee', '', 'Gx', 6, ''), mkRow('e', 'Eli', 'Gold', 'Bx', 8, ''), mkRow('f', 'Fay', 'Blue', 'Gx', 7, '👎 Can\'t make it'), mkRow('g', 'Gus', 'Blue', 'Gx', 7, '')]);
vm.runInContext('sortPracticeRoster', sandbox)(sp, 7, 9);
eq(sp.rows.slice(1).map(r => r[3]).join(','), 'Gus,Fay,Eli,Cal,Bea,Ann,Dee', 'practice roster: Blue, Gold, Silver, TBD, Practice Squad, blank; availability within team');
eq(sp.rows[0].length, 9, 'temp sort columns removed');

// 7b. Checkin sort mode: Gender > Name only, ignoring Team
const spCheckin = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note'],
  mkRow('a', 'Ann', 'Practice Squad', 'Gx', 7, ''), mkRow('b', 'Bea', 'TBD', 'Gx', 7, ''), mkRow('c', 'Cal', 'Silver', 'Bx', 8, ''), mkRow('d', 'Dee', 'Blue', 'Bx', 6, '')]);
vm.runInContext('sortPracticeRoster', sandbox)(spCheckin, 4, 9, 'checkin');
eq(spCheckin.rows.slice(1).map(r => r[3]).join(','), 'Cal,Dee,Ann,Bea', 'checkin mode: Gender then Name, ignoring Team');
eq(spCheckin.rows[0].length, 9, 'checkin mode: no temp sort columns added');

// 7c. practiceRosterGroupByColumns_: team mode groups by Team+Gender, checkin mode by Gender only
const groupByCols = vm.runInContext('practiceRosterGroupByColumns_', sandbox);
eq(groupByCols().join(','), '5,6', 'team mode (default) groups by Team, Gender');
eq(groupByCols('team').join(','), '5,6', 'team mode groups by Team, Gender');
eq(groupByCols('checkin').join(','), '6', 'checkin mode groups by Gender only');

// 7d. readPracticeRosterDate_: reads the date header right after Grade, whether or not Group exists
eq(m.readPracticeRosterDate_(['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note']), '9/9', 'readPracticeRosterDate_ with Group column');
eq(m.readPracticeRosterDate_(['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', '9/13', '9/13 Note']), '9/13', 'readPracticeRosterDate_ on a legacy sheet without Group');
eq(m.readPracticeRosterDate_(['PlayerID', '#', 'Full Name', 'Team', 'Gender']), null, 'readPracticeRosterDate_ with no date column after Grade');

// 7e. capturePracticeRosterGroupValues_ / restorePracticeRosterGroupValues_: Group survives a
// PlayerID-keyed round trip regardless of row order, and is skipped when there's no Group column.
const gHeader = ['PlayerID', '#', 'Group', 'Full Name', 'Team', 'Gender', 'Grade'];
const gSheet = new FakeSheet('Practice Roster', [gHeader,
  ['a', 1, 'Red', 'Ann', 'Blue', 'Gx', 7],
  ['b', 2, '', 'Bea', 'Blue', 'Gx', 7],
  ['c', 3, 'Blue', 'Cal', 'Gold', 'Bx', 8]]);
const captured = m.capturePracticeRosterGroupValues_(gSheet, gHeader);
eq(JSON.stringify(captured), JSON.stringify({ a: 'Red', c: 'Blue' }), 'capturePracticeRosterGroupValues_ keys by PlayerID, skips blanks');

const gRebuilt = new FakeSheet('Practice Roster', [gHeader, ['c', '', '', 'Cal', '', '', ''], ['a', '', '', 'Ann', '', '', ''], ['newkid', '', '', 'Zed', '', '', '']]);
m.restorePracticeRosterGroupValues_(gRebuilt, 3, captured);
eq(gRebuilt.rows.slice(1).map(r => r[2]).join(','), 'Blue,Red,', 'restorePracticeRosterGroupValues_ follows PlayerID, blank for a player with none captured');

const noGroupHeader = ['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade'];
const noGroupSheet = new FakeSheet('Practice Roster', [noGroupHeader, ['a', 1, 'Ann', 'Blue', 'Gx', 7]]);
eq(JSON.stringify(m.capturePracticeRosterGroupValues_(noGroupSheet, noGroupHeader)), '{}', 'capturePracticeRosterGroupValues_ on a legacy sheet without Group yields nothing');
const gpS = new FakeSheet('Game Roster Prep', [['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', '9/13 Availability', '9/13 Note'],
  ['a', '', 'Ann', 'TBD', 'Bx', 7, '', ''], ['b', '', 'Bea', 'Blue', 'Gx', 7, '', ''], ['c', '', 'Cal', 'Silver', 'Bx', 8, '', ''], ['d', '', 'Dee', 'Gold', 'Bx', 8, '', '']]);
vm.runInContext('sortGameRosterPrep', sandbox)(gpS, 4, 8, { team: 4, activationStatus: null, gender: 5, availability: 7, fullName: 3 });
eq(gpS.rows.slice(1).map(r => r[2]).join(','), 'Bea,Dee,Cal,Ann', 'coach game prep sorted Blue, Gold, Silver, TBD');
eq(gpS.rows[0].length, 8, 'game prep temp column removed');
m.CONFIG.gameRosterPrep.hasTeam = true;
const layout3 = m.getGameRosterPrepColumnLayout('9/13', gCols);
eq(layout3.headers.join('|'), 'PlayerID|#|Full Name|Team|Gender|Grade|9/13 Availability|9/13 Note', 'game prep headers with Team, no activation');
// 8. onOpen builds the menu (a ReferenceError here hid the whole menu in 3.31)
const menuItems = [];
const stubMenu = { addItem: (l) => { menuItems.push(l); return stubMenu; }, addSeparator: () => stubMenu, addToUi: () => { menuItems.push('<added>'); } };
sandbox.SpreadsheetApp.getUi = () => ({ createMenu: () => stubMenu });
m.CONFIG.gameRosterPrep.hasActivationStatus = false;
try { vm.runInContext('onOpen', sandbox)(); } catch (e) { fails++; console.log('FAIL onOpen threw:', e.message); }
eq(menuItems[menuItems.length - 1], '<added>', 'menu added to UI');
eq(menuItems.includes('⬆️ Apply Activation Status'), false, 'no Apply Activation Status item when the season has none');
m.CONFIG.gameRosterPrep.hasActivationStatus = true; menuItems.length = 0;
try { vm.runInContext('onOpen', sandbox)(); } catch (e) { fails++; console.log('FAIL onOpen threw:', e.message); }
eq(menuItems.includes('⬆️ Apply Activation Status'), true, 'Apply Activation Status item when the season has it');
// 9. Sync Game Info to Calendar: per-team title with emoji, blank Team (all-team event) gets no team segment
eq(m.gameTeamLabel('Blue'), '🟦 Blue', 'gameTeamLabel known team');
eq(m.gameTeamLabel('tbd'), '', 'gameTeamLabel TBD is blank');
eq(m.gameTeamLabel(''), '', 'gameTeamLabel blank stays blank');
eq(m.gameTeamLabel('Mystery'), 'Mystery', 'gameTeamLabel unlisted team falls back to raw text');
const giHeader = ['Date', 'Game #', 'Team', 'Warmup Arrival', 'Game Start', 'Done By', 'Field Name', 'Field Location', 'Game Note', 'Opponent', 'Oponent Team Page', 'Google Calendar Event ID', 'Google Calendar Warmup Event ID'];
const gi = new FakeSheet('📍Game Info', [giHeader,
  ['9/26 Sat', 'Game 1', 'Blue', '1:15 PM', '2:00 PM', '3:45 PM', 'Garfield HS Field', 'East', '', 'Salmon Bay Panthers 8th', '', '', ''],
  ['9/26 Sat', 'Game 1', '', '2:15 PM', '3:00 PM', '4:45 PM', 'Franklin HS Field', 'W', '', 'Some Rival', '', '', '']]);
const giSsStub = { getSheetByName: n => (n === '📍Game Info' ? gi : null) };
const gameSpecs = m.getGameEventSpecsFromSheet(giSsStub, gi);
eq(gameSpecs.length, 4, 'game + warmup spec for both rows');
eq(gameSpecs[0].title, '🎯 🟦 Blue vs. Salmon Bay Panthers 8th', 'Blue game title has emoji and team name');
eq(gameSpecs[1].title, '🥏 🟦 Blue Warmup', 'Blue warmup title has emoji and team name');
eq(gameSpecs[2].title, '🎯 Game vs. Some Rival', 'blank Team (all-team event) has no team segment');
eq(gameSpecs[3].title, '🥏 Game Warmup', 'blank Team warmup falls back to the plain title');

// 10. Draw Group Borders: pure group-boundary and column-default logic
eq(m.groupBorderCellText_(null), '', 'groupBorderCellText_ null is blank');
eq(m.groupBorderCellText_(undefined), '', 'groupBorderCellText_ undefined is blank');
eq(m.groupBorderCellText_('  Blue  '), 'Blue', 'groupBorderCellText_ trims');
eq(m.groupBorderCellText_(0), '0', 'groupBorderCellText_ zero is not blank');

eq(m.findGroupBoundaryRows_([], [0]).join(','), '', 'findGroupBoundaryRows_ empty rows');
eq(m.findGroupBoundaryRows_([['Blue']], [0]).join(','), '0', 'findGroupBoundaryRows_ single row is its own group');
eq(
  m.findGroupBoundaryRows_([['Blue', 'Gx'], ['Blue', 'Gx'], ['Blue', 'Bx'], ['Gold', 'Gx']], [0]).join(','),
  '2,3',
  'findGroupBoundaryRows_ groups by Team only'
);
eq(
  m.findGroupBoundaryRows_([['Blue', 'Gx'], ['Blue', 'Gx'], ['Blue', 'Bx'], ['Gold', 'Gx']], [0, 1]).join(','),
  '1,2,3',
  'findGroupBoundaryRows_ groups by Team + Gender'
);
eq(
  m.findGroupBoundaryRows_([['Blue'], ['Gold'], ['Blue']], [0]).join(','),
  '0,1,2',
  'findGroupBoundaryRows_ non-contiguous groups: every transition borders, no sortedness check'
);
eq(
  m.findGroupBoundaryRows_([['', 'Gx'], ['', 'Gx'], ['Blue', 'Gx']], [0]).join(','),
  '1,2',
  'findGroupBoundaryRows_ blank Team is its own group'
);

const gbHeaders = ['PlayerID', 'Full Name', 'Team', 'Gender', 'Grade'];
eq(m.defaultGroupByColumns_(gbHeaders, 3, 1, []).join(','), 'Team', 'defaultGroupByColumns_ narrow selection wins');
eq(m.defaultGroupByColumns_(gbHeaders, 3, 2, []).join(','), 'Team,Gender', 'defaultGroupByColumns_ multi-column selection');
eq(m.defaultGroupByColumns_(gbHeaders, 1, 5, ['Gender', 'Team']).join(','), 'Gender,Team', 'defaultGroupByColumns_ whole-row selection falls back to last used');
eq(m.defaultGroupByColumns_(gbHeaders, 1, 5, []).join(','), '', 'defaultGroupByColumns_ whole-row selection with no memory yet');
eq(m.defaultGroupByColumns_(['PlayerID', '', 'Team', 'Gender'], 1, 2, ['Team']).join(','), 'PlayerID', 'defaultGroupByColumns_ skips blank headers in the selection');

// 11. Draw Group Borders: sheet-level draw is idempotent and finds columns by header, not position
const gbSheet = new FakeSheet('Extra Player Info', [
  ['PlayerID', 'Full Name', 'Team', 'Returning'],
  ['a', 'Ann', 'Blue', 'TRUE'],
  ['b', 'Bea', 'Blue', 'FALSE'],
  ['c', 'Cal', 'Gold', 'TRUE']
]);
const gbResult1 = m.drawGroupBordersOnSheet_(gbSheet, 2, 3, ['Team'], 'SOLID_MEDIUM', '#000000');
eq(gbResult1.groupCount, 2, 'drawGroupBordersOnSheet_ finds 2 groups (Blue, Gold)');
const bottomBorderRows1 = gbSheet.borders.filter(b => b.bottom === true).map(b => b.r);
eq(bottomBorderRows1.join(','), '3,4', 'drawGroupBordersOnSheet_ borders rows 3 (end of Blue) and 4 (end of Gold)');
eq(gbSheet.borders.every(b => b.top == null && b.left == null && b.right == null), true, 'drawGroupBordersOnSheet_ only ever touches the bottom edge');

// Re-run after a row is removed (Gold's row is gone): must not leave row 4's stale border behind
gbSheet.rows.splice(3, 1);
gbSheet.borders = [];
const gbResult2 = m.drawGroupBordersOnSheet_(gbSheet, 2, 2, ['Team'], 'SOLID_MEDIUM', '#000000');
eq(gbResult2.groupCount, 1, 'drawGroupBordersOnSheet_ re-run after a row is removed finds 1 group');
const bottomBorderRows2 = gbSheet.borders.filter(b => b.bottom === true).map(b => b.r);
eq(bottomBorderRows2.join(','), '3', 'drawGroupBordersOnSheet_ re-run draws only the current last row, no stale row 4 border');
const clearedRows = gbSheet.borders.filter(b => b.bottom === false).map(b => b.r).sort();
eq(clearedRows.join(','), '2,3', 'drawGroupBordersOnSheet_ clears bottom borders on every row in range before redrawing');

// 12. A 1-row selection is valid (no artificial minimum): it's its own group, bordered once.
const gbOneRow = new FakeSheet('Extra Player Info', [
  ['PlayerID', 'Full Name', 'Team', 'Returning'],
  ['a', 'Ann', 'Blue', 'TRUE']
]);
const gbResult3 = m.drawGroupBordersOnSheet_(gbOneRow, 2, 1, ['Team'], 'SOLID_MEDIUM', '#000000');
eq(gbResult3.groupCount, 1, 'drawGroupBordersOnSheet_ single-row selection is its own group');
eq(gbOneRow.borders.filter(b => b.bottom === true).map(b => b.r).join(','), '2', 'drawGroupBordersOnSheet_ borders the single selected row');

// 13. Draw Group Borders "#" column: pure per-row position-in-group logic
eq(m.groupPositions_([], [0]).join(','), '', 'groupPositions_ empty rows');
eq(m.groupPositions_([['Blue']], [0]).join(','), '1', 'groupPositions_ single row starts at 1');
eq(
  m.groupPositions_([['Blue', 'Gx'], ['Blue', 'Gx'], ['Blue', 'Bx'], ['Gold', 'Gx']], [0]).join(','),
  '1,2,3,1',
  'groupPositions_ resets when Team changes, ignores Gender'
);
eq(
  m.groupPositions_([['Blue', 'Gx'], ['Blue', 'Gx'], ['Blue', 'Bx'], ['Gold', 'Gx']], [0, 1]).join(','),
  '1,2,1,1',
  'groupPositions_ resets on Team + Gender'
);
eq(
  m.groupPositions_([['Blue'], ['Gold'], ['Blue']], [0]).join(','),
  '1,1,1',
  'groupPositions_ non-contiguous groups each restart at 1, no sortedness check'
);
eq(
  m.groupPositions_([['', 'Gx'], ['', 'Gx'], ['Blue', 'Gx']], [0]).join(','),
  '1,2,1',
  'groupPositions_ blank Team is its own group'
);

// 14. Draw Group Borders "#" column: sheet-level write is a plain value, full range, header-found
const gbNum = new FakeSheet('Extra Player Info', [
  ['PlayerID', 'Full Name', 'Team', '#'],
  ['a', 'Ann', 'Blue', ''],
  ['b', 'Bea', 'Blue', ''],
  ['c', 'Cal', 'Gold', '']
]);
const gbNumResult1 = m.drawGroupBordersOnSheet_(gbNum, 2, 3, ['Team'], 'SOLID_MEDIUM', '#000000', true);
eq(gbNumResult1.numbered, true, 'drawGroupBordersOnSheet_ reports numbered:true when "#" exists and requested');
eq(gbNum.rows.slice(1).map(r => r[3]).join(','), '1,2,1', 'drawGroupBordersOnSheet_ writes plain 1-based positions into "#", resetting on Team');

// Re-run after a row is removed: full overwrite, no stale leftover value
gbNum.rows.splice(3, 1);
const gbNumResult2 = m.drawGroupBordersOnSheet_(gbNum, 2, 2, ['Team'], 'SOLID_MEDIUM', '#000000', true);
eq(gbNumResult2.numbered, true, 'drawGroupBordersOnSheet_ re-run after a row is removed still numbers');
eq(gbNum.rows.slice(1).map(r => r[3]).join(','), '1,2', 'drawGroupBordersOnSheet_ re-run fully overwrites "#" for the current range');

// numberColumn requested but not checked: never writes "#", never errors
const gbNoCheckbox = new FakeSheet('Extra Player Info', [
  ['PlayerID', 'Full Name', 'Team', '#'],
  ['a', 'Ann', 'Blue', 'stale']
]);
const gbNoCheckboxResult = m.drawGroupBordersOnSheet_(gbNoCheckbox, 2, 1, ['Team'], 'SOLID_MEDIUM', '#000000', false);
eq(gbNoCheckboxResult.numbered, false, 'drawGroupBordersOnSheet_ reports numbered:false when not requested');
eq(gbNoCheckbox.rows[1][3], 'stale', 'drawGroupBordersOnSheet_ leaves "#" untouched when numbering not requested');

// No "#" column on the sheet at all: numbering requested but silently skipped, no error, border still drawn
const gbNoHashColumn = new FakeSheet('Extra Player Info', [
  ['PlayerID', 'Full Name', 'Team'],
  ['a', 'Ann', 'Blue'],
  ['b', 'Bea', 'Gold']
]);
const gbNoHashResult = m.drawGroupBordersOnSheet_(gbNoHashColumn, 2, 2, ['Team'], 'SOLID_MEDIUM', '#000000', true);
eq(gbNoHashResult.numbered, false, 'drawGroupBordersOnSheet_ numbered:false when the sheet has no "#" column');
eq(gbNoHashResult.groupCount, 2, 'drawGroupBordersOnSheet_ still draws the border when there is no "#" column to number');

console.log(fails === 0 ? 'ALL ASSERTIONS PASSED' : `${fails} FAILURES`);
process.exit(fails ? 1 : 0);
