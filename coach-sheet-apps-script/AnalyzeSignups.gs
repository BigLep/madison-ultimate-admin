/**
 * Analyze Signups: the report of Players whose Sources disagree or are incomplete.
 * It only reports; the portal's Final Forms Backfill does the joining.
 *
 * Reads the Signups and Final Forms tabs directly (not the Roster), so it works
 * before the Roster is generated.
 */

const ANALYZE_SIGNUPS_SHEET_NAME = 'Analyze Signups';

// Final Forms export columns this report reads (fixed positions, 0-based).
const ANALYZE_FINAL_FORMS_INDEX = {
  studentId: 0,   // A
  firstName: 3,   // D
  lastName: 4,    // E
  grade: 22       // W
};

/**
 * Menu entry: write or replace the "Analyze Signups" sheet and show the counts.
 */
function analyzeSignups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const signupsSheet = ss.getSheetByName(CONFIG.signups.sheetName);
  if (!signupsSheet) {
    ui.alert('Error', `Sheet "${CONFIG.signups.sheetName}" not found.`, ui.ButtonSet.OK);
    return;
  }
  const finalFormsSheet = ss.getSheetByName(CONFIG.finalForms.sheetName);
  if (!finalFormsSheet) {
    ui.alert('Error', `Sheet "${CONFIG.finalForms.sheetName}" not found. Run "Update Final Forms" after adding the tab.`, ui.ButtonSet.OK);
    return;
  }

  const signupsValues = signupsSheet.getDataRange().getValues();
  const finalFormsValues = finalFormsSheet.getDataRange().getValues();
  const result = analyzeSignupsData(signupsValues[0], signupsValues.slice(1), finalFormsValues);

  writeAnalyzeSignupsSheet(ss, result);

  const summary = result.sections.map(s => `${s.rows.length}: ${s.title}`).join('\n');
  ui.alert('Analyze Signups', `Report written to "${ANALYZE_SIGNUPS_SHEET_NAME}".\n\n${summary}`, ui.ButtonSet.OK);
}

/**
 * Normalization, as the portal defines it: trim, lowercase, strip internal
 * whitespace, strip apostrophes, fold accents to plain letters, keep hyphens.
 */
function normalizeName(value) {
  return (value === null || value === undefined ? '' : value.toString())
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .replace(/['\u2018\u2019]/g, '');
}

/**
 * A birthdate as yyyy-mm-dd text whether it arrived as ISO text, US text, or a Date.
 */
function normalizeBirthdate(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    const pad = (n) => (n < 10 ? '0' : '') + n;
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }
  const text = value.toString().trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  const us = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (us) return `${us[3]}-${us[1].padStart(2, '0')}-${us[2].padStart(2, '0')}`;
  return text;
}

/**
 * Pure analysis. Takes the Signups header row, the Signups data rows, and the
 * whole Final Forms tab (header included). Returns { sections: [{ title, headers, rows }] }.
 */
function analyzeSignupsData(signupsHeaders, signupsRows, finalFormsValues) {
  const text = (v) => (v === null || v === undefined) ? '' : v.toString().trim();
  const headerIndex = {};
  signupsHeaders.forEach((h, i) => { const name = text(h); if (name && headerIndex[name] === undefined) headerIndex[name] = i; });
  const col = (key) => {
    const name = SIGNUPS_HEADERS[key];
    if (headerIndex[name] === undefined) throw new Error(`Signups tab is missing header "${name}"`);
    return headerIndex[name];
  };
  const idCol = col('playerId'), spsCol = col('spsStudentId'), prefCol = col('preferredFirstName'), lastCol = col('lastName');
  const dobCol = col('dateOfBirth'), gradeCol = col('grade'), c1EmailCol = col('caretaker1Email');

  const players = signupsRows
    .map(row => ({
      playerId: text(row[idCol]),
      spsStudentId: text(row[spsCol]),
      fullName: `${text(row[prefCol])} ${text(row[lastCol])}`.trim(),
      lastName: text(row[lastCol]),
      birthdate: normalizeBirthdate(row[dobCol]),
      grade: text(row[gradeCol]),
      caretaker1Email: text(row[c1EmailCol])
    }))
    .filter(p => p.playerId);

  const finalForms = finalFormsValues.slice(1)
    .map(row => ({
      studentId: text(row[ANALYZE_FINAL_FORMS_INDEX.studentId]),
      fullName: `${text(row[ANALYZE_FINAL_FORMS_INDEX.firstName])} ${text(row[ANALYZE_FINAL_FORMS_INDEX.lastName])}`.trim(),
      grade: text(row[ANALYZE_FINAL_FORMS_INDEX.grade])
    }))
    .filter(f => f.studentId);
  const finalFormsIds = new Set(finalForms.map(f => f.studentId));
  const signupIds = new Set(players.map(p => p.spsStudentId).filter(Boolean));

  // 1. Signups with no SPS Student ID (the Final Forms Join has not happened yet).
  const noSpsId = players.filter(p => !p.spsStudentId)
    .map(p => [p.playerId, p.fullName, [p.birthdate && `DOB ${p.birthdate}`, p.grade && `Grade ${p.grade}`].filter(Boolean).join(', ')]);

  // 2. Final Forms students whose StudentID is on no signup row.
  const noSignup = finalForms.filter(f => !signupIds.has(f.studentId))
    .map(f => [f.studentId, f.fullName, f.grade]);

  // 3. Signups whose SPS Student ID is missing from the Final Forms tab.
  const notInFinalForms = players.filter(p => p.spsStudentId && !finalFormsIds.has(p.spsStudentId))
    .map(p => [p.playerId, p.fullName, p.spsStudentId]);

  // 4. Suspected duplicates: same normalized last name and birthdate, or same SPS Student ID.
  const duplicateRows = [];
  const groupBy = (keyFn, label) => {
    const groups = new Map();
    players.forEach(p => { const k = keyFn(p); if (!k) return; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(p); });
    groups.forEach((group, key) => {
      if (group.length < 2) return;
      group.forEach(p => {
        const others = group.filter(o => o !== p).map(o => `${o.playerId} (${o.fullName})`).join(', ');
        duplicateRows.push([p.playerId, p.fullName, `${label(key)} as ${others}`]);
      });
    });
  };
  groupBy(p => (p.birthdate && normalizeName(p.lastName)) ? `${normalizeName(p.lastName)}|${p.birthdate}` : '', () => 'same last name + birthdate');
  groupBy(p => p.spsStudentId, key => `same SPS Student ID ${key}`);

  // 5. Not Profile Complete: which of Grade, Date of Birth, Caretaker 1 Email are missing.
  const incomplete = players
    .map(p => {
      const missing = [];
      if (!p.grade) missing.push('Grade');
      if (!p.birthdate) missing.push('Date of Birth');
      if (!p.caretaker1Email) missing.push('Caretaker 1 Email');
      return missing.length ? [p.playerId, p.fullName, missing.join(', ')] : null;
    })
    .filter(Boolean);

  return {
    sections: [
      { title: 'Signups with no SPS Student ID (Final Forms Join pending)', headers: ['PlayerID', 'Full Name', 'Signup details'], rows: noSpsId },
      { title: 'Final Forms students with no signup', headers: ['SPS Student ID', 'Full Name (Final Forms)', 'Grade'], rows: noSignup },
      { title: 'Signups whose SPS Student ID is not in Final Forms', headers: ['PlayerID', 'Full Name', 'SPS Student ID'], rows: notInFinalForms },
      { title: 'Suspected duplicate signups', headers: ['PlayerID', 'Full Name', 'Why'], rows: duplicateRows },
      { title: 'Signups not Profile Complete', headers: ['PlayerID', 'Full Name', 'Missing'], rows: incomplete }
    ]
  };
}

/**
 * Write the report: a timestamp, then each section as a heading, a header row,
 * and its rows (or "none").
 */
function writeAnalyzeSignupsSheet(ss, result) {
  let sheet = ss.getSheetByName(ANALYZE_SIGNUPS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANALYZE_SIGNUPS_SHEET_NAME);
  } else {
    sheet.clear();
  }

  const out = [];
  const bold = [];
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  out.push([`Analyze Signups, generated ${stamp}`, '', '']);
  bold.push(out.length);
  out.push(['Reads Signups and Final Forms directly. Fix a wrong value in its Source; the portal\'s Final Forms Backfill does the joining.', '', '']);

  result.sections.forEach((section, index) => {
    out.push(['', '', '']);
    out.push([`${index + 1}. ${section.title} (${section.rows.length})`, '', '']);
    bold.push(out.length);
    out.push(section.headers.slice());
    bold.push(out.length);
    if (section.rows.length === 0) {
      out.push(['none', '', '']);
    } else {
      section.rows.forEach(row => out.push(row.map(v => v === null || v === undefined ? '' : v.toString())));
    }
  });

  sheet.getRange(1, 1, out.length, 3).setValues(out);
  bold.forEach(rowNumber => sheet.getRange(rowNumber, 1, 1, 3).setFontWeight('bold'));
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 520);
  sheet.setFrozenRows(1);
}
