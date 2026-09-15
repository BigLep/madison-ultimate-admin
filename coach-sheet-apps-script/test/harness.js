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
    setFontWeight: () => ({}), clearDataValidations: () => ({}), setWrap: () => ({}), setBorder: () => ({})
  };
};
const src = ['Code.gs', 'ManagedConditionalFormatting.gs', 'SheetBuilderUtils.gs', 'Availability.gs', 'BuildPracticeRoster.gs', 'BuildGameRosterPrepSheet.gs']
  .map(f => fs.readFileSync(dir + '/' + f, 'utf8')).join('\n');
const sandbox = { console: { log() {}, warn() {}, error() {} }, SpreadsheetApp: { getUi: () => ({}), flush() {} } };
const vm = require('vm'); vm.createContext(sandbox);
vm.runInContext(src + `
  ;module = { seedPrintoutPlayerRows, populatePracticeRosterData, populateGameRosterPrepData, getGameRosterPrepColumnLayout, findAvailabilityColumns, playerIdLookupFormula, CONFIG, seedAvailabilityRows_, ROSTER_COLUMNS };`, sandbox);
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
const pr = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note', '9/13 Activation Status', '9/13 Availability', '9/13 Note']]);
const seeded = m.seedPrintoutPlayerRows(pr, roster, 2, 1, 3);
eq(seeded.rowCount, 2, 'include filter drops Bob');
eq(pr.rows[1][0], 'e9jpt', 'PlayerID value in A2'); eq(pr.rows[2][0], 'vwrk2', 'PlayerID value in A3');
eq(pr.rows[2][2], `=IF($A3="","",IFERROR(XLOOKUP($A3,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`, 'Full Name formula row 3');
const availCols = m.findAvailabilityColumns(pa, '9/9', 'Practice Availability');
eq(availCols.playerIdColumn, 'A', 'pa playerIdColumn'); eq(availCols.fullNameColumn, 'B', 'pa fullNameColumn'); eq(availCols.availabilityColumn, 'E', 'pa availabilityColumn');
sandbox.findNextGameAfterPractice = () => null;
m.populatePracticeRosterData(pr, roster, hdr, pa, availCols, 2, ga, { formattedDate: '9/13', ordinalForDate: 1 });
eq(pr.rows[1][3], `=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$U:$U),""))`, 'Team formula');
eq(pr.rows[2][5], `=IF($A3="","",IFERROR(XLOOKUP($A3,'📋 Roster'!$A:$A,'📋 Roster'!$O:$O),""))`, 'Grade formula row 3');
eq(pr.rows[1][6], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Practice Availability'!$A:$A,'Practice Availability'!$E:$E),""))`, 'practice availability formula');
eq(pr.rows[1][7], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Practice Availability'!$A:$A,'Practice Availability'!$F:$F),""))`, 'practice note formula');
eq(pr.rows[1][8], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$F:$F),""))`, 'next game activation formula');
eq(pr.rows[1][9], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$E:$E),""))`, 'next game availability formula');
eq(pr.rows[1][10], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$G:$G),""))`, 'next game note formula');

// 2. Older availability sheet without PlayerID: falls back to Full Name join
const oldPa = new FakeSheet('Practice Availability', [['Full Name', 'Grade', 'Gender Identification', '9/9', '9/9 Note']]);
const oldCols = m.findAvailabilityColumns(oldPa, '9/9', 'Practice Availability');
eq(oldCols.playerIdColumn, null, 'old sheet has no playerIdColumn');
const pr2 = new FakeSheet('Practice Roster', [pr.rows[0].slice(0, 8)]);
m.seedPrintoutPlayerRows(pr2, roster, 2, 1, 3);
m.populatePracticeRosterData(pr2, roster, hdr, oldPa, oldCols, 2, null, null);
eq(pr2.rows[1][6], `=IF($C2="","",IFERROR(XLOOKUP($C2,'Practice Availability'!$A:$A,'Practice Availability'!$D:$D),""))`, 'fallback join by Full Name');

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
// 5. # column formula lives in column B and chains on B
const np = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade'], ['x', '', '', 'Blue', 'Gx', 7], ['y', '', '', 'Blue', 'Bx', 8]]);
np.getRange = (function (orig) { return function (r, c, nr, nc) { const rg = orig.call(this, r, c, nr, nc); rg.copyTo = (dest) => { dest.setValues([[np.rows[1][1]]]); }; return rg; }; })(np.getRange);
vm.runInContext('populateNumberColumn', sandbox)(np, 2);
eq(np.rows[1][1], '=IF(OR(D1<>D2,E1<>E2),1,B1+1)', '# formula in column B chains on B');
// 6. Season without Activation Status: next-game block is two columns, game availability build skips the column
m.CONFIG.gameRosterPrep.hasActivationStatus = false;
const ng = vm.runInContext('practiceRosterNextGameColumns', sandbox)();
eq(ng.activation, null, 'no activation column'); eq(ng.availability, 9, 'next game availability at base+3'); eq(ng.note, 10, 'next game note at base+4');
const pr3 = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note', '9/13 Availability', '9/13 Note']]);
m.seedPrintoutPlayerRows(pr3, roster, 2, 1, 3);
m.populatePracticeRosterData(pr3, roster, hdr, pa, availCols, 2, ga, { formattedDate: '9/13', ordinalForDate: 1 });
eq(pr3.rows[1][8], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$E:$E),""))`, 'next game availability without activation');
eq(pr3.rows[1][9], `=IF($A2="","",IFERROR(XLOOKUP($A2,'Game Availability'!$A:$A,'Game Availability'!$G:$G),""))`, 'next game note without activation');
eq(pr3.rows[1][10], undefined, 'nothing written past the note column');
const layout2 = m.getGameRosterPrepColumnLayout('9/13', gCols);
eq(layout2.headers.join('|'), 'PlayerID|#|Full Name|Gender|Grade|9/13 Availability|9/13 Note', 'game prep headers without activation');
// 7. Team sort order on the practice roster and coach game prep
const mkRow = (id, name, team, gender, grade, avail) => [id, '', name, team, gender, grade, avail, ''];
const sp = new FakeSheet('Practice Roster', [['PlayerID', '#', 'Full Name', 'Team', 'Gender', 'Grade', '9/9', '9/9 Note'],
  mkRow('a', 'Ann', 'Practice Squad', 'Gx', 7, ''), mkRow('b', 'Bea', 'TBD', 'Gx', 7, ''), mkRow('c', 'Cal', 'Silver', 'Bx', 8, ''),
  mkRow('d', 'Dee', '', 'Gx', 6, ''), mkRow('e', 'Eli', 'Gold', 'Bx', 8, ''), mkRow('f', 'Fay', 'Blue', 'Gx', 7, '👎 Can\'t make it'), mkRow('g', 'Gus', 'Blue', 'Gx', 7, '')]);
vm.runInContext('sortPracticeRoster', sandbox)(sp, 7, 8);
eq(sp.rows.slice(1).map(r => r[2]).join(','), 'Gus,Fay,Eli,Cal,Bea,Ann,Dee', 'practice roster: Blue, Gold, Silver, TBD, Practice Squad, blank; availability within team');
eq(sp.rows[0].length, 8, 'temp sort columns removed');
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
console.log(fails === 0 ? 'ALL ASSERTIONS PASSED' : `${fails} FAILURES`);
process.exit(fails ? 1 : 0);
