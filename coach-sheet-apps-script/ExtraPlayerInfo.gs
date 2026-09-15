/**
 * Sync Extra Player Info: the one tab where coaches author per-player facts the
 * family cannot (Team, Returning, Number of Past Seasons, Tryout Group, Tryout
 * ID, Signup Grade, and the Include In Generated Rosters override). Full Name
 * and Signup Playing Experience are per-row XLOOKUP formulas, not authored; the
 * latter exists purely so a coach can see the Signups text next to Number of
 * Past Seasons while filling it in. Assign Tryout IDs (assignTryoutIds, its own
 * menu item) fills blank Tryout ID cells for Players with a Tryout Group set.
 * Keyed by PlayerID; one row per Player, rows added by this sync rather than typed.
 * Never deletes or reorders rows.
 */

const EXTRA_PLAYER_INFO_BOOLEAN_CHOICES = ['TRUE', 'FALSE'];

/**
 * Menu entry: create the tab if missing, apply dropdowns, append a row for every
 * Signups PlayerID not already present, and report the counts.
 */
function syncExtraPlayerInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const signupsSheet = ss.getSheetByName(CONFIG.signups.sheetName);
  if (!signupsSheet) {
    ui.alert('Error', `Sheet "${CONFIG.signups.sheetName}" not found. Extra Player Info rows come from its PlayerID column.`, ui.ButtonSet.OK);
    return;
  }

  const signupsHeaders = signupsSheet.getRange(1, 1, 1, Math.max(1, signupsSheet.getLastColumn())).getValues()[0];
  const signupsLetters = resolveSignupsColumns(signupsHeaders);
  const playingExperienceLetter = signupsLetters[SIGNUPS_HEADERS.playingExperience];
  const playerIdLetter = signupsLetters[SIGNUPS_HEADERS.playerId];
  const signupsName = `'${CONFIG.signups.sheetName.replace(/'/g, "''")}'`;

  const players = readSignupsPlayers(signupsSheet);
  const sheet = ensureExtraPlayerInfoSheet(ss);
  applyExtraPlayerInfoValidations(sheet);

  const lastRow = sheet.getLastRow();
  const existingIds = new Set();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(row => {
      const id = row[0] === null || row[0] === undefined ? '' : row[0].toString().trim();
      if (id) existingIds.add(id);
    });
  }

  const missing = players.filter(p => !existingIds.has(p.playerId));
  if (missing.length > 0) {
    const rosterLetters = resolveRosterColumns();
    const rosterName = `'${CONFIG.roster.sheetName.replace(/'/g, "''")}'`;
    const idLetter = rosterLetters[CONFIG.columns.playerId];
    const fullNameLetter = rosterLetters[CONFIG.columns.fullName];
    const startRow = lastRow + 1;
    const rows = missing.map((p, i) => {
      const row = startRow + i;
      return [
        p.playerId,
        `=IFERROR(XLOOKUP($A${row},${rosterName}!$${idLetter}:$${idLetter},${rosterName}!$${fullNameLetter}:$${fullNameLetter}),"")`,
        '', '', '',
        `=IFERROR(XLOOKUP($A${row},${signupsName}!$${playerIdLetter}:$${playerIdLetter},${signupsName}!$${playingExperienceLetter}:$${playingExperienceLetter}),"")`,
        '', '', '', ''
      ];
    });
    if (sheet.getMaxRows() < startRow + rows.length - 1) {
      sheet.insertRowsAfter(sheet.getMaxRows(), startRow + rows.length - 1 - sheet.getMaxRows());
    }
    sheet.getRange(startRow, 1, rows.length, EXTRA_PLAYER_INFO_HEADERS.length).setValues(rows);
  }

  SpreadsheetApp.flush();
  const total = existingIds.size + missing.length;
  console.log(`Extra Player Info: added ${missing.length}, total ${total}`);
  ui.alert('Extra Player Info Synced',
    `Added ${missing.length} PlayerID row(s); ${total} row(s) total.\n\n` +
    `Fill in Team, Returning, Number of Past Seasons, Tryout Group, Tryout ID, Signup Grade, and Include In Generated Rosters here (Signup Playing Experience is a read-only reference for spot-checking Number of Past Seasons). Blank Include means "included". Rows are never deleted or reordered by this sync.`,
    ui.ButtonSet.OK);
}

/**
 * Menu entry: fill blank Tryout ID cells in Extra Player Info for Players who
 * have a Tryout Group but no Tryout ID yet. An existing Tryout ID is never
 * touched or renumbered (it may already be printed or handed out); a new
 * Player in a Grade/gender bucket gets the next offset after that bucket's
 * current highest id, and when several new Players land in the same bucket
 * at once they're assigned in Full Name order. Needs the Roster's Grade and
 * Gender Identification, so a Player missing either is reported and skipped.
 */
function assignTryoutIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(CONFIG.extraPlayerInfo.sheetName);
  if (!sheet) {
    ui.alert('Error', `Sheet "${CONFIG.extraPlayerInfo.sheetName}" not found. Run "Sync Extra Player Info" first.`, ui.ButtonSet.OK);
    return;
  }
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  if (!rosterSheet) {
    ui.alert('Error', `Sheet "${CONFIG.roster.sheetName}" not found. Tryout ID needs its Grade and Gender Identification; run "Generate Fresh Roster" first.`, ui.ButtonSet.OK);
    return;
  }

  const text = (v) => (v === null || v === undefined) ? '' : v.toString().trim();

  const table = readRosterTable(rosterSheet);
  const rIdCol = table.col(CONFIG.columns.playerId);
  const rFullNameCol = table.col(CONFIG.columns.fullName);
  const rGradeCol = table.col(CONFIG.columns.grade);
  const rGenderCol = table.col(CONFIG.columns.genderIdentification);
  const rosterByPlayerId = {};
  table.rows.forEach(row => {
    const pid = text(row[rIdCol]);
    if (!pid) return;
    rosterByPlayerId[pid] = { fullName: text(row[rFullNameCol]), grade: text(row[rGradeCol]), gender: text(row[rGenderCol]) };
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No Players', 'Extra Player Info has no data rows. Run "Sync Extra Player Info" first.', ui.ButtonSet.OK);
    return;
  }
  const width = EXTRA_PLAYER_INFO_HEADERS.length;
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const playerIdIdx = EXTRA_PLAYER_INFO_HEADERS.indexOf('PlayerID');
  const tryoutGroupIdx = EXTRA_PLAYER_INFO_HEADERS.indexOf('Tryout Group');
  const tryoutIdIdx = EXTRA_PLAYER_INFO_HEADERS.indexOf('Tryout ID');
  const VALID_GRADES = ['6', '7', '8'];
  const BUCKET_SIZE = 50; // offsets 0-49 for Bx, 50-99 for Gx within a Grade

  // Highest existing offset already used in each Grade/gender bucket, on a
  // shared 0-49 local scale (a Gx id's local offset is its id mod 100, minus 50).
  const maxLocalOffset = {};
  values.forEach(row => {
    const num = Number(row[tryoutIdIdx]);
    if (row[tryoutIdIdx] === '' || row[tryoutIdIdx] === null || !Number.isFinite(num)) return;
    const grade = Math.floor(num / 100);
    const withinHundred = num - grade * 100;
    const gender = withinHundred < BUCKET_SIZE ? 'Bx' : 'Gx';
    const localOffset = withinHundred < BUCKET_SIZE ? withinHundred : withinHundred - BUCKET_SIZE;
    const key = `${grade}-${gender}`;
    if (maxLocalOffset[key] === undefined || localOffset > maxLocalOffset[key]) maxLocalOffset[key] = localOffset;
  });

  // Candidates: a Tryout Group set, no Tryout ID yet, and a Roster Grade/Gender
  // Identification that resolves to 6/7/8 and Bx/Gx.
  const candidates = [];
  const skipped = [];
  values.forEach((row, i) => {
    const playerId = text(row[playerIdIdx]);
    if (!playerId) return;
    const hasTryoutId = row[tryoutIdIdx] !== '' && row[tryoutIdIdx] !== null && row[tryoutIdIdx] !== undefined;
    if (text(row[tryoutGroupIdx]) === '' || hasTryoutId) return;
    const info = rosterByPlayerId[playerId] || { fullName: playerId, grade: '', gender: '' };
    if (VALID_GRADES.indexOf(info.grade) === -1 || ['Bx', 'Gx'].indexOf(info.gender) === -1) {
      skipped.push(info);
      return;
    }
    candidates.push({ rowIndex: i, fullName: info.fullName, grade: info.grade, gender: info.gender });
  });

  if (candidates.length === 0) {
    ui.alert('No New Tryout IDs Needed',
      skipped.length === 0
        ? 'Every Player with a Tryout Group already has a Tryout ID.'
        : `Every Player with a Tryout Group and a resolvable Grade/Gender already has a Tryout ID. ${skipped.length} Player(s) still need a Grade and/or Gender Identification before an id can be assigned: ${skipped.map(s => s.fullName).join(', ')}.`,
      ui.ButtonSet.OK);
    return;
  }

  // Group by Grade/gender, sort each group by Full Name, assign the next local
  // offset after that bucket's current highest; never renumber an existing id.
  const byBucket = {};
  candidates.forEach(c => (byBucket[`${c.grade}-${c.gender}`] = byBucket[`${c.grade}-${c.gender}`] || []).push(c));

  const overflow = [];
  Object.keys(byBucket).forEach(key => {
    const [grade, gender] = key.split('-');
    const base = Number(grade) * 100 + (gender === 'Bx' ? 0 : BUCKET_SIZE);
    let nextLocalOffset = (maxLocalOffset[key] === undefined ? -1 : maxLocalOffset[key]) + 1;
    byBucket[key]
      .sort((a, b) => a.fullName.localeCompare(b.fullName, undefined, { sensitivity: 'base' }))
      .forEach(c => {
        if (nextLocalOffset >= BUCKET_SIZE) { overflow.push(c); return; }
        c.tryoutId = base + nextLocalOffset;
        values[c.rowIndex][tryoutIdIdx] = c.tryoutId;
        nextLocalOffset++;
      });
  });

  const assigned = candidates.filter(c => c.tryoutId !== undefined);
  if (assigned.length > 0) {
    sheet.getRange(2, tryoutIdIdx + 1, values.length, 1).setValues(values.map(row => [row[tryoutIdIdx]]));
    SpreadsheetApp.flush();
  }

  console.log(`Assigned ${assigned.length} Tryout ID(s), skipped ${skipped.length}, overflow ${overflow.length}`);
  let message = assigned.length > 0
    ? `Assigned ${assigned.length} new Tryout ID(s):\n\n${assigned.sort((a, b) => a.tryoutId - b.tryoutId).map(c => `${c.tryoutId}: ${c.fullName}`).join('\n')}`
    : 'No Tryout IDs could be assigned.';
  if (skipped.length > 0) {
    message += `\n\n${skipped.length} Player(s) skipped, no resolvable Grade/Gender yet: ${skipped.map(s => s.fullName).join(', ')}.`;
  }
  if (overflow.length > 0) {
    message += `\n\n${overflow.length} Player(s) could not be assigned, their Grade/gender bucket is full (past 49 players): ${overflow.map(c => c.fullName).join(', ')}.`;
  }
  ui.alert('Tryout IDs Assigned', message, ui.ButtonSet.OK);
}

/**
 * Every Signups Player as { playerId, lastName, preferredFirstName }, in Roster
 * order (Last Name, then Preferred First Name).
 */
function readSignupsPlayers(signupsSheet) {
  const lastRow = signupsSheet.getLastRow();
  const lastCol = Math.max(1, signupsSheet.getLastColumn());
  const headers = signupsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  resolveSignupsColumns(headers); // throws naming any missing header
  const indexOf = (key) => headers.findIndex(h => h !== null && h !== undefined && h.toString().trim() === SIGNUPS_HEADERS[key]);
  const idCol = indexOf('playerId');
  const lastCol0 = indexOf('lastName');
  const prefCol = indexOf('preferredFirstName');
  if (lastRow < 2) return [];
  const values = signupsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const players = values
    .map(row => ({
      playerId: row[idCol] === null || row[idCol] === undefined ? '' : row[idCol].toString().trim(),
      lastName: (row[lastCol0] || '').toString(),
      preferredFirstName: (row[prefCol] || '').toString()
    }))
    .filter(p => p.playerId);
  players.sort((a, b) => a.lastName.localeCompare(b.lastName, undefined, { sensitivity: 'base' }) || a.preferredFirstName.localeCompare(b.preferredFirstName, undefined, { sensitivity: 'base' }));
  return players;
}

/**
 * Get the Extra Player Info tab, creating it with the header row when missing.
 * An existing tab with a different header row is rewritten only when it has no
 * data rows; once coaches have authored rows the mismatch is an error, because a
 * silent header rewrite could change what their values mean.
 */
function ensureExtraPlayerInfoSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.extraPlayerInfo.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.extraPlayerInfo.sheetName);
    console.log(`Created sheet "${CONFIG.extraPlayerInfo.sheetName}"`);
  }
  const width = EXTRA_PLAYER_INFO_HEADERS.length;
  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  const current = sheet.getRange(1, 1, 1, width).getValues()[0].map(h => h === null || h === undefined ? '' : h.toString().trim());
  const matches = current.every((h, i) => h === EXTRA_PLAYER_INFO_HEADERS[i]);
  if (!matches) {
    if (sheet.getLastRow() > 1) {
      throw new Error(`"${CONFIG.extraPlayerInfo.sheetName}" header row is [${current.join(', ')}] but the script expects [${EXTRA_PLAYER_INFO_HEADERS.join(', ')}] and the tab already has data rows. Fix the header (or move the data) and run again.`);
    }
    sheet.getRange(1, 1, 1, width).setValues([EXTRA_PLAYER_INFO_HEADERS]);
    console.log(`Rewrote "${CONFIG.extraPlayerInfo.sheetName}" header row`);
  }
  sheet.getRange(1, 1, 1, width).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Dropdowns on the coach-authored columns. Blank stays allowed everywhere (a
 * list rule only judges non-empty input), which is what lets a blank Include mean
 * "included" (coach sheet ADR 0002). No checkboxes: a checkbox cannot be blank.
 *
 * The Team list is seeded once from CONFIG.teams and left alone afterwards so coaches
 * can edit the dropdown in the sheet without the next sync undoing it.
 */
function applyExtraPlayerInfoValidations(sheet) {
  const rows = Math.max(1, sheet.getMaxRows() - 1);
  const column = (letter) => sheet.getRange(`${letter}2:${letter}${rows + 1}`);

  const teamRange = column(EXTRA_PLAYER_INFO_LETTERS.team);
  const hasTeamRule = teamRange.getDataValidations().some(row => row[0] !== null);
  if (!hasTeamRule) {
    teamRange.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.teams, true)
      .setAllowInvalid(false)
      .setHelpText('Team for the season (CONFIG.teams). Edit this dropdown list if the teams change.')
      .build());
  }

  const booleanRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(EXTRA_PLAYER_INFO_BOOLEAN_CHOICES, true)
    .setAllowInvalid(false)
    .build();
  column(EXTRA_PLAYER_INFO_LETTERS.returning).setDataValidation(booleanRule);
  column(EXTRA_PLAYER_INFO_LETTERS.include).setDataValidation(booleanRule);

  column(EXTRA_PLAYER_INFO_LETTERS.numberOfPastSeasons).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('Seasons of prior organized Ultimate play, read from Signup Playing Experience. Leave blank if not yet reviewed.')
    .build());

  column(EXTRA_PLAYER_INFO_LETTERS.tryoutId).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(600, 899)
    .setAllowInvalid(false)
    .setHelpText('Three-digit tryout id: first digit is Grade (6/7/8), then 00-49 for Bx or 50-99 for Gx, assigned alphabetically by Full Name within that Grade and gender. Leave blank if not yet assigned.')
    .build());
}

/**
 * Diagnostics helper: the tab exists and its header row is exactly EXTRA_PLAYER_INFO_HEADERS.
 * Returns { pass, line }.
 */
function checkExtraPlayerInfoSheet(ss) {
  const sheet = ss.getSheetByName(CONFIG.extraPlayerInfo.sheetName);
  if (!sheet) {
    return { pass: false, line: `❌ Sheet "${CONFIG.extraPlayerInfo.sheetName}" (Extra Player Info) NOT found; run "Sync Extra Player Info" to create it` };
  }
  const width = EXTRA_PLAYER_INFO_HEADERS.length;
  const current = sheet.getRange(1, 1, 1, Math.min(width, sheet.getMaxColumns())).getValues()[0].map(h => h === null || h === undefined ? '' : h.toString().trim());
  const missing = EXTRA_PLAYER_INFO_HEADERS.filter((h, i) => current[i] !== h);
  if (missing.length > 0) {
    return { pass: false, line: `❌ Sheet "${CONFIG.extraPlayerInfo.sheetName}" header row should be [${EXTRA_PLAYER_INFO_HEADERS.join(', ')}]; missing or misplaced: ${missing.join(', ')}` };
  }
  return { pass: true, line: `✅ Sheet "${CONFIG.extraPlayerInfo.sheetName}" (Extra Player Info) exists with the expected header row` };
}
