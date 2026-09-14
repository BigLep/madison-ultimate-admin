/**
 * Madison Ultimate coach sheet: roster generation and shared configuration.
 *
 * The 📋 Roster is a formula-only view keyed by the Signups PlayerID (see
 * docs/adr/0001-roster-keyed-by-signups-playerid.md and ADR 0003). Generate Fresh
 * Roster writes one header row, then one row per Player: the PlayerID as a plain
 * value in column A and a per-row formula in every other column. Nothing is
 * authored in the Roster itself. Every fact lives in exactly one Source: Signups,
 * Extra Player Info, Final Forms, or Newsletter Subscribers. Re-run Generate Fresh
 * Roster to pick up new or removed Signups.
 */

// Script Version - Increment this number when making changes
const SCRIPT_VERSION = '3.27';

// Roster layout: row 1 holds the column headers, data rows start at row 2 (one
// per Player, each self-contained so the sheet can be sorted and filtered freely).
// Readers (Build Practice Roster, Game Roster Prep, Full Name Diff, analysis) use
// these two constants only.
const ROSTER_HEADER_ROW = 1;
const ROSTER_FIRST_DATA_ROW = 2;

// Configuration
const CONFIG = {
  finalForms: {
    // Per-season: update to this season's FinalForms exports folder (same folder the
    // finalforms-export automation's DRIVE_FOLDER_ID uploads into) before the season starts.
    folderId: '1WgD4hY0fIZlQEBt7ekOlHIECA-HgOMIZ', // 2026 Fall Final Forms
    sheetName: 'Final Forms'
  },
  // Read-only IMPORTRANGE mirror of the portal's Signups sheet (one row per Player,
  // keyed by PlayerID). The Roster's key column and most passthrough columns read it.
  signups: {
    sheetName: '2026 Fall Signups'
  },
  // Coach-authored per-player facts (Team, Returning, Include override), keyed by PlayerID.
  extraPlayerInfo: {
    sheetName: 'Extra Player Info'
  },
  // Google Groups was retired as the mailing-list system of record in spring 2026.
  // Buttondown (newsletterSubscribers/buttondown below) is the real source now.
  newsletterSubscribers: {
    sheetName: 'Newsletter Subscribers'
  },
  buttondown: {
    apiBase: 'https://api.buttondown.com/v1',
    // Script property (Extensions > Apps Script > Project Settings > Script Properties),
    // not committed here. Read access to subscribers is enough.
    apiKeyProperty: 'BUTTONDOWN_API_KEY'
  },
  roster: {
    sheetName: '📋 Roster'
  },
  attendance: {
    sheetName: '🏃 Attendance'
  },
  practiceInfo: {
    sheetName: '📍Practice Info'
  },
  gameInfo: {
    sheetName: '📍Game Info'
  },
  fieldsSheet: {
    sheetName: '📍Fields'
  },
  practiceAvailability: {
    sheetName: 'Practice Availability'
  },
  gameAvailability: {
    sheetName: 'Game Availability'
  },

  // Roster column headers other files look up by name. Every value here must be a
  // ROSTER_COLUMNS name (Run Diagnostics checks the live header row for all of them).
  columns: {
    playerId: 'PlayerID',
    spsStudentId: 'SPS Student ID',
    preferredFirstName: 'Preferred First Name',
    lastName: 'Last Name',
    fullName: 'Full Name',
    grade: 'Grade',
    genderIdentification: 'Gender Identification',
    team: 'Team',
    returning: 'Returning',
    includeInGeneratedRosters: 'Include In Generated Rosters',
    profileComplete: 'Profile Complete?',
    areAllFormsParentSigned: 'Are All Forms Parent Signed',
    areAllFormsStudentSigned: 'Are All Forms Student Signed',
    physicalCleared: 'Physical Cleared',
    finalFormsCleared: 'Final Forms Cleared?',
    dateOfBirth: 'Date of Birth',
    studentPersonalEmail: 'Student Personal Email',
    studentNewsletterStatus: 'Student Newsletter Status',
    caretaker1Name: 'Caretaker 1 Name',
    caretaker1Email: 'Caretaker 1 Email',
    caretaker1NewsletterStatus: 'Caretaker 1 Newsletter Status',
    caretaker2Name: 'Caretaker 2 Name',
    caretaker2Email: 'Caretaker 2 Email',
    caretaker2NewsletterStatus: 'Caretaker 2 Newsletter Status'
  },

  // Shared base column structure for roster printouts (Practice Roster and Game Roster Prep)
  // Header order is driven by rosterPrintoutBaseColumnKeys (not Object.keys order).
  rosterPrintoutBaseColumnKeys: ['number', 'fullName', 'team', 'gender', 'grade'],
  rosterPrintoutBaseColumns: {
    // Base columns (always present): # | Full Name | Team | Gender | Grade
    number: { name: '#', index: 1 },
    fullName: { name: 'Full Name', index: 2 },
    team: { name: 'Team', index: 3 },
    gender: { name: 'Gender', index: 4 },
    grade: { name: 'Grade', index: 5 }
    // Additional columns (availability, notes) are added dynamically after these base columns
  },

  // Game Roster Prep sheet: season-specific columns (change per season)
  gameRosterPrep: {
    hasTeam: false,           // If true, include Team column; if false, omit it
    hasActivationStatus: true // If true, include $date Activation Status column and sort by it first
  }
};

// Signups headers the Roster formulas reference. Resolved to column letters by name
// at generation time, so the portal can reorder or add Signups columns freely.
const SIGNUPS_HEADERS = {
  playerId: 'PlayerID',
  spsStudentId: 'SPS Student ID',
  preferredFirstName: 'Preferred First Name',
  legalFirstName: 'Legal First Name',
  lastName: 'Last Name',
  grade: 'Grade',
  elementarySchool: 'Elementary School',
  genderIdentification: 'Gender Identification',
  pronouns: 'Pronouns',
  dateOfBirth: 'Date of Birth',
  studentSpsEmail: 'Student SPS Email',
  studentPersonalEmail: 'Student Personal Email',
  caretaker1Name: 'Caretaker 1 Name',
  caretaker1Email: 'Caretaker 1 Email',
  caretaker2Name: 'Caretaker 2 Name',
  caretaker2Email: 'Caretaker 2 Email',
  allergies: 'Allergies',
  competingSports: 'Competing Sports and Activities',
  jerseySize: 'Jersey Size',
  playingExperience: 'Playing Experience',
  hopes: 'Hopes',
  otherInfo: 'Other Info',
  mediaOptOut: 'Media Opt-Out',
  photoDriveFileId: 'Photo Drive File ID',
  seededAt: 'Seeded At',
  profileComplete: 'Profile Complete'
};

// Final Forms export columns at fixed positions (validated by finalforms-export).
const FINAL_FORMS_LETTERS = {
  studentId: 'A',
  parentSigned: 'P',
  studentSigned: 'Q',
  gender: 'U',
  grade: 'W',
  physicalClearance: 'AB'
};

// Extra Player Info header row, in column order. Sync Extra Player Info creates it.
// Tryout Group and Signup Grade were added directly in the live sheet (coach-typed
// during tryouts) before the code caught up; this list now matches that live order.
const EXTRA_PLAYER_INFO_HEADERS = ['PlayerID', 'Full Name', 'Team', 'Returning', 'Number of Past Seasons', 'Signup Playing Experience', 'Tryout Group', 'Tryout ID', 'Signup Grade', 'Include In Generated Rosters'];

// Column letters of the Extra Player Info tab, derived from the header order above.
const EXTRA_PLAYER_INFO_LETTERS = {
  playerId: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('PlayerID') + 1),
  fullName: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Full Name') + 1),
  team: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Team') + 1),
  returning: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Returning') + 1),
  numberOfPastSeasons: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Number of Past Seasons') + 1),
  signupPlayingExperience: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Signup Playing Experience') + 1),
  tryoutGroup: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Tryout Group') + 1),
  tryoutId: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Tryout ID') + 1),
  signupGrade: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Signup Grade') + 1),
  include: getColumnLetter(EXTRA_PLAYER_INFO_HEADERS.indexOf('Include In Generated Rosters') + 1)
};

// Source labels used in header notes and Diagnostics output.
const ROSTER_SOURCE = {
  signups: 'Signups',
  extraPlayerInfo: 'Extra Player Info',
  finalForms: 'Final Forms',
  newsletter: 'Newsletter Subscribers',
  derived: 'Derived'
};

/**
 * Roster column definitions: the single source of truth for the 📋 Roster.
 *
 * Order here is column order in the sheet. Each entry has the header name, a type
 * and source for the header note, an explanation, and a formula builder that
 * receives a context of resolved column letters and returns the formula for one
 * data row. The PlayerID key column has no builder (formula: null): Generate Fresh
 * Roster writes it as plain values, one per Player, and every other column's
 * formula looks that row's PlayerID up. Adding or reordering a column means
 * editing this list and running Generate Fresh Roster.
 *
 * Formula builder context (see buildRosterFormulas):
 *   s(key)   Signups range like 'Signups'!$X:$X for a SIGNUPS_HEADERS key
 *   f(key)   Final Forms range for a FINAL_FORMS_LETTERS key
 *   e(key)   Extra Player Info range for an EXTRA_PLAYER_INFO_LETTERS key
 *   r(name)  Roster sibling cell on the same row, like $X2, for a ROSTER_COLUMNS name
 *   lookupSignups(key), lookupFinalForms(key, whenMissing), lookupExtra(key),
 *   newsletterStatus(name), and rowFormula(body) wrap the shared patterns.
 */
const ROSTER_COLUMNS = [
  {
    name: 'PlayerID',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Key: every non-empty PlayerID in Signups, written as plain values by Generate Fresh Roster (initially sorted by Last Name then Preferred First Name). Sort and filter the sheet freely; re-run Generate Fresh Roster to pick up new or removed Signups.',
    formula: null
  },
  {
    name: 'SPS Student ID',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Set on the Signups row by the portal once the Final Forms Join succeeds. Blank until then, which leaves every Final Forms column blank or FALSE.',
    formula: (c) => c.rowFormula(c.lookupSignups('spsStudentId'))
  },
  {
    name: 'Preferred First Name',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'The first name the Player goes by, chosen by the family at signup.',
    formula: (c) => c.rowFormula(c.lookupSignups('preferredFirstName'))
  },
  {
    name: 'Legal First Name',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Only filled when it differs from Preferred First Name.',
    formula: (c) => c.rowFormula(c.lookupSignups('legalFirstName'))
  },
  {
    name: 'Last Name',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('lastName'))
  },
  {
    name: 'Full Name',
    type: 'String',
    source: ROSTER_SOURCE.derived,
    note: 'TRIM(Preferred First Name & " " & Last Name). The key every Generated Roster uses to refer to a Player, and column A of the availability sheets (which the portal matches by PlayerID).',
    formula: (c) => c.rowFormula(`TRIM(${c.r('Preferred First Name')}&" "&${c.r('Last Name')})`)
  },
  {
    name: 'Elementary School',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('elementarySchool'))
  },
  {
    name: 'Date of Birth',
    type: 'Date',
    source: ROSTER_SOURCE.signups,
    note: 'Signups Date of Birth (ISO text) converted to a real date.',
    formula: (c) => c.rowFormula(`LET(v,${c.lookupSignups('dateOfBirth')},IF(v="","",IF(ISNUMBER(v),v,IFERROR(DATEVALUE(v),""))))`)
  },
  {
    name: 'Player Allergies',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Signups "Allergies".',
    formula: (c) => c.rowFormula(c.lookupSignups('allergies'))
  },
  {
    name: 'Competing Sports and Activities',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('competingSports'))
  },
  {
    name: 'Jersey Size',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('jerseySize'))
  },
  {
    name: 'Playing Experience',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('playingExperience'))
  },
  {
    name: 'Player hopes for the season',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Signups "Hopes".',
    formula: (c) => c.rowFormula(c.lookupSignups('hopes'))
  },
  {
    name: 'Other Player Info',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Signups "Other Info".',
    formula: (c) => c.rowFormula(c.lookupSignups('otherInfo'))
  },
  {
    name: 'Grade',
    type: 'Number',
    source: `${ROSTER_SOURCE.finalForms}, ${ROSTER_SOURCE.signups} fallback`,
    note: 'Final Forms Grade when the SPS Student ID resolves, else the Signups Grade.',
    formula: (c) => c.rowFormula(`LET(ff,${c.lookupFinalForms('grade', '""')},IF(ff="",${c.lookupSignups('grade')},ff))`)
  },
  {
    name: 'Final Forms Gender',
    type: 'Enum',
    source: ROSTER_SOURCE.finalForms,
    note: 'Final Forms Gender (Male or Female). Blank when the SPS Student ID is missing or not in the export.',
    formula: (c) => c.rowFormula(c.lookupFinalForms('gender', '""'))
  },
  {
    name: 'Signup Gender',
    type: 'Enum',
    source: ROSTER_SOURCE.signups,
    note: 'Signups Gender Identification collapsed: Girl or Gx to "Gx", Boy or Bx to "Bx", else blank.',
    formula: (c) => c.rowFormula(`LET(g,TO_TEXT(${c.lookupSignups('genderIdentification')}),IF(REGEXMATCH(g,"Girl|Gx"),"Gx",IF(REGEXMATCH(g,"Boy|Bx"),"Bx","")))`)
  },
  {
    name: 'Pronouns',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups (semicolon-joined when several were chosen).',
    formula: (c) => c.rowFormula(c.lookupSignups('pronouns'))
  },
  {
    name: 'Gender Default Handling',
    type: 'Boolean',
    source: ROSTER_SOURCE.derived,
    note: 'TRUE when nothing about gender needs a coach\'s attention. FALSE when Final Forms Gender (Female, Male, or anything else such as Non-Binary) and Signup Gender disagree, or when Pronouns include anything outside he/him for a Bx or she/her for a Gx: a prompt to check in with the Player, not a verdict.',
    formula: (c) => {
      const ff = c.r('Final Forms Gender');
      const sg = c.r('Signup Gender');
      const gi = c.r('Gender Identification');
      const pronouns = c.r('Pronouns');
      // Strip the expected pronouns and separators; anything left over flags. Plain substrings on purpose (no regex escapes):
      // with tokens he, him, she, her, they, them, every unexpected token leaves a residue.
      const leftover = (expected) => `(REGEXREPLACE(p,"${expected}|[;,/ ]","")<>"")`;
      // Final Forms Female/Male map to Gx/Bx; any other non-blank value (for example Non-Binary) counts as a disagreement with a Gx or Bx signup.
      return c.rowFormula(`LET(ff,IF(${ff}="","",IF(${ff}="Female","Gx",IF(${ff}="Male","Bx","Other"))),p,LOWER(${pronouns}),((ff<>"")*(${sg}<>"")*(ff<>${sg})+(${gi}="Bx")*${leftover('he|him')}+(${gi}="Gx")*${leftover('she|her')})=0)`);
    }
  },
  {
    name: 'Gender Identification',
    type: 'Enum',
    source: ROSTER_SOURCE.derived,
    note: 'Gx or Bx: Signup Gender when set, else Final Forms Gender mapped Female to Gx and Male to Bx, else blank. The value Generated Rosters print.',
    formula: (c) => {
      const sg = c.r('Signup Gender');
      const fg = c.r('Final Forms Gender');
      return c.rowFormula(`IF(${sg}<>"",${sg},IF(${fg}="Female","Gx",IF(${fg}="Male","Bx","")))`);
    }
  },
  {
    name: 'Team',
    type: 'Enum',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'Coach-assigned squad for the season, authored in Extra Player Info after tryouts.',
    formula: (c) => c.rowFormula(c.lookupExtra('team'))
  },
  {
    name: 'Returning',
    type: 'Boolean',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'Whether the Player played for Madison Ultimate in a prior season, authored in Extra Player Info.',
    formula: (c) => c.rowFormula(c.lookupExtra('returning'))
  },
  {
    name: 'Number of Past Seasons',
    type: 'Number',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'Coach-authored count of seasons of prior organized Ultimate play, read from the Signups "Playing Experience" text and entered in Extra Player Info (which also carries that text alongside it for spot-checking). Blank means not yet reviewed, not zero.',
    formula: (c) => c.rowFormula(c.lookupExtra('numberOfPastSeasons'))
  },
  {
    name: 'Has Playing Experience',
    type: 'Boolean',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'TRUE when Returning is TRUE or Number of Past Seasons is greater than zero, else FALSE (including when Number of Past Seasons is blank/not yet reviewed).',
    formula: (c) => c.rowFormula(`(${c.r('Returning')}=TRUE)+(N(${c.r('Number of Past Seasons')})>0)>0`)
  },
  {
    name: 'Include In Generated Rosters',
    type: 'Boolean',
    source: ROSTER_SOURCE.derived,
    note: 'Extra Player Info value when one is set, else TRUE: every Player is on Generated Rosters unless a coach says otherwise (coach sheet ADR 0002).',
    formula: (c) => c.rowFormula(`LET(e,${c.lookupExtra('include')},IF(e="",TRUE,e))`)
  },
  {
    name: 'Profile Complete?',
    type: 'Boolean',
    source: ROSTER_SOURCE.signups,
    note: 'Defined and written by the portal (Player Info, Caretaker Info, and Photo Upload all done; portal ADR 0006). Passed through, never computed here. A Seeded Signup starts FALSE until the family finishes.',
    formula: (c) => c.rowFormula(`UPPER(TO_TEXT(${c.lookupSignups('profileComplete')}))="TRUE"`)
  },
  {
    name: 'Are All Forms Parent Signed',
    type: 'Boolean',
    source: ROSTER_SOURCE.finalForms,
    note: 'Final Forms "Are All Forms Parent Signed" equals TRUE. FALSE when the SPS Student ID is missing or not in the export.',
    formula: (c) => c.rowFormula(c.finalFormsFlag('parentSigned'))
  },
  {
    name: 'Are All Forms Student Signed',
    type: 'Boolean',
    source: ROSTER_SOURCE.finalForms,
    note: 'Final Forms "Are All Forms Student Signed" equals TRUE. FALSE when the SPS Student ID is missing or not in the export.',
    formula: (c) => c.rowFormula(c.finalFormsFlag('studentSigned'))
  },
  {
    name: 'Physical Cleared',
    type: 'Boolean',
    source: ROSTER_SOURCE.finalForms,
    note: 'Final Forms "Physical Clearance" equals "Cleared". FALSE when the SPS Student ID is missing or not in the export.',
    formula: (c) => c.rowFormula(`IF(${c.r('SPS Student ID')}="",FALSE,${c.lookupFinalFormsRaw('physicalClearance')}="Cleared")`)
  },
  {
    name: 'Final Forms Cleared?',
    type: 'Boolean',
    source: ROSTER_SOURCE.derived,
    note: 'TRUE only when all forms are parent signed, all forms are student signed, and the physical is cleared.',
    formula: (c) => c.rowFormula(`(${c.r('Are All Forms Parent Signed')}=TRUE)*(${c.r('Are All Forms Student Signed')}=TRUE)*(${c.r('Physical Cleared')}=TRUE)=1`)
  },
  {
    name: 'Caretaker 1 Name',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('caretaker1Name'))
  },
  {
    name: 'Caretaker 1 Email',
    type: 'Email',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('caretaker1Email'))
  },
  {
    name: 'Caretaker 1 Newsletter Status',
    type: 'Enum',
    source: ROSTER_SOURCE.newsletter,
    note: 'Buttondown status for Caretaker 1 Email ("regular" = subscribed, "unactivated" = pending confirmation, "unsubscribed"), "not a member" when absent, blank when there is no email.',
    formula: (c) => c.rowFormula(c.newsletterStatus('Caretaker 1 Email'))
  },
  {
    name: 'Caretaker 2 Name',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('caretaker2Name'))
  },
  {
    name: 'Caretaker 2 Email',
    type: 'Email',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('caretaker2Email'))
  },
  {
    name: 'Caretaker 2 Newsletter Status',
    type: 'Enum',
    source: ROSTER_SOURCE.newsletter,
    note: 'Buttondown status for Caretaker 2 Email ("regular" = subscribed, "unactivated" = pending confirmation, "unsubscribed"), "not a member" when absent, blank when there is no email.',
    formula: (c) => c.rowFormula(c.newsletterStatus('Caretaker 2 Email'))
  },
  {
    name: 'Student SPS Email',
    type: 'Email',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups.',
    formula: (c) => c.rowFormula(c.lookupSignups('studentSpsEmail'))
  },
  {
    name: 'Student Personal Email',
    type: 'Email',
    source: ROSTER_SOURCE.signups,
    note: 'Passthrough from Signups, blanked when the domain is seattleschools.org.',
    formula: (c) => c.rowFormula(`LET(v,${c.lookupSignups('studentPersonalEmail')},IF(REGEXMATCH(LOWER(v),"@seattleschools\\.org$"),"",v))`)
  },
  {
    name: 'Student Newsletter Status',
    type: 'Enum',
    source: ROSTER_SOURCE.newsletter,
    note: 'Buttondown status for Student Personal Email ("regular" = subscribed, "unactivated" = pending confirmation, "unsubscribed"), "not a member" when absent, blank when there is no email.',
    formula: (c) => c.rowFormula(c.newsletterStatus('Student Personal Email'))
  },
  {
    name: 'Media OK',
    type: 'Boolean',
    source: ROSTER_SOURCE.signups,
    note: 'TRUE when photos of the Player may appear in team communications. FALSE when the family declared a Media Opt-Out. Does not affect the Player Photo.',
    formula: (c) => c.rowFormula(`LOWER(TO_TEXT(${c.lookupSignups('mediaOptOut')}))<>"true"`)
  },
  {
    name: 'Photo Drive File ID',
    type: 'String',
    source: ROSTER_SOURCE.signups,
    note: 'Drive file id of the Player Photo the family uploaded through the portal.',
    formula: (c) => c.rowFormula(c.lookupSignups('photoDriveFileId'))
  },
  {
    name: 'Photo Link',
    type: 'Hyperlink',
    source: ROSTER_SOURCE.derived,
    note: 'Link to the full-size Player Photo when a Photo Drive File ID is present.',
    formula: (c) => {
      const id = c.r('Photo Drive File ID');
      return c.rowFormula(`IF(${id}="","",HYPERLINK("https://drive.google.com/uc?id="&${id},"photo"))`);
    }
  },
  {
    name: 'Tryout Group',
    type: 'Number',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'Coach-assigned tryout group, authored in Extra Player Info.',
    formula: (c) => c.rowFormula(c.lookupExtra('tryoutGroup'))
  },
  {
    name: 'Tryout ID',
    type: 'Number',
    source: ROSTER_SOURCE.extraPlayerInfo,
    note: 'Three-digit id for referencing a Player on tryout printouts, authored in Extra Player Info: first digit is Grade (6, 7, or 8), remaining two digits are a sequence assigned within that Grade and gender (00-49 for Bx, 50-99 for Gx) by Full Name alphabetically. A fixed assignment, not recomputed by the Roster, so it stays stable once printed.',
    formula: (c) => c.rowFormula(c.lookupExtra('tryoutId'))
  }
];

/**
 * Map the Signups header row to column letters for every SIGNUPS_HEADERS name.
 * Pure: takes the header row values, returns { headerName: letter }.
 * Throws naming every missing header so the fix is obvious.
 */
function resolveSignupsColumns(headerRowValues) {
  const letters = {};
  headerRowValues.forEach((header, index) => {
    const name = header === null || header === undefined ? '' : header.toString().trim();
    if (name && letters[name] === undefined) letters[name] = getColumnLetter(index + 1);
  });
  const missing = Object.values(SIGNUPS_HEADERS).filter(name => letters[name] === undefined);
  if (missing.length > 0) {
    throw new Error(`Signups tab "${CONFIG.signups.sheetName}" is missing header(s): ${missing.join(', ')}. The Roster formulas need every one of these in row 1.`);
  }
  const resolved = {};
  Object.values(SIGNUPS_HEADERS).forEach(name => { resolved[name] = letters[name]; });
  return resolved;
}

/**
 * Column letters of every ROSTER_COLUMNS entry, in code order: { name: letter }.
 */
function resolveRosterColumns() {
  const letters = {};
  ROSTER_COLUMNS.forEach((col, index) => { letters[col.name] = getColumnLetter(index + 1); });
  return letters;
}

/**
 * Build the formula for every ROSTER_COLUMNS entry for one data row, in code
 * order; null for the PlayerID key column, which holds a plain value.
 * Pure: takes the resolved Signups letters (from resolveSignupsColumns) and the
 * sheet row number. Every reference to a sibling Roster cell is row-relative
 * ($X<row>), so sorting the sheet in place keeps each row consistent.
 */
function buildRosterFormulas(signupsLetters, row) {
  const rosterLetters = resolveRosterColumns();
  if (!Number.isInteger(row) || row < ROSTER_FIRST_DATA_ROW) throw new Error(`buildRosterFormulas needs a data row number (>= ${ROSTER_FIRST_DATA_ROW}), got ${row}`);
  const quote = (sheetName) => `'${sheetName.replace(/'/g, "''")}'`;
  const S = quote(CONFIG.signups.sheetName);
  const F = quote(CONFIG.finalForms.sheetName);
  const E = quote(CONFIG.extraPlayerInfo.sheetName);
  const N = quote(CONFIG.newsletterSubscribers.sheetName);

  const signupsLetter = (key) => {
    const header = SIGNUPS_HEADERS[key];
    if (!header || !signupsLetters[header]) throw new Error(`Unknown Signups column key "${key}"`);
    return signupsLetters[header];
  };
  const rosterLetter = (name) => {
    if (!rosterLetters[name]) throw new Error(`Unknown Roster column "${name}"`);
    return rosterLetters[name];
  };
  const keyRange = `$${rosterLetter('PlayerID')}${row}`;
  const spsRange = `$${rosterLetter('SPS Student ID')}${row}`;

  const c = {
    s: (key) => { const l = signupsLetter(key); return `${S}!$${l}:$${l}`; },
    f: (key) => { const l = FINAL_FORMS_LETTERS[key]; return `${F}!$${l}:$${l}`; },
    e: (key) => { const l = EXTRA_PLAYER_INFO_LETTERS[key]; return `${E}!$${l}:$${l}`; },
    r: (name) => { const l = rosterLetter(name); return `$${l}${row}`; },
    // Wrap a body so a row whose PlayerID was cleared stays blank. Plain per-row
    // formulas, no ARRAYFORMULA: that means no scalar function may be applied to a
    // whole-column range (Sheets would implicitly intersect it with the current
    // row); pass ranges straight to XLOOKUP and convert the single lookup key instead.
    rowFormula: (body) => `=IF(${keyRange}="","",${body})`,
    // Signups value for this row's PlayerID, blank when missing.
    lookupSignups: (key) => `IFERROR(XLOOKUP(${keyRange},${c.s('playerId')},${c.s(key)}),"")`,
    // Final Forms value for this row's SPS Student ID, blank when missing. Signups
    // delivers the ID as text (IMPORTRANGE) while the Final Forms import stores
    // StudentID as a number, so look up the numeric form first and fall back to the
    // text form in case a future export stores IDs as text.
    lookupFinalFormsRaw: (key) => `IFERROR(XLOOKUP(VALUE(${spsRange}),${c.f('studentId')},${c.f(key)}),IFERROR(XLOOKUP(${spsRange},${c.f('studentId')},${c.f(key)}),""))`,
    // Same, but a blank SPS Student ID yields whenMissing instead of matching a blank export row.
    lookupFinalForms: (key, whenMissing) => `IF(${spsRange}="",${whenMissing},${c.lookupFinalFormsRaw(key)})`,
    // Final Forms TRUE/FALSE flag as a real boolean whether the import stored text or boolean.
    finalFormsFlag: (key) => `IF(${spsRange}="",FALSE,UPPER(TO_TEXT(${c.lookupFinalFormsRaw(key)}))="TRUE")`,
    // Extra Player Info value for this row's PlayerID, blank when missing.
    lookupExtra: (key) => `IFERROR(XLOOKUP(${keyRange},${c.e('playerId')},${c.e(key)}),"")`,
    // Buttondown status for the email in a sibling Roster column (XLOOKUP matches
    // case-insensitively, so no LOWER is needed).
    newsletterStatus: (emailColumnName) => {
      const email = c.r(emailColumnName);
      return `IF(${email}="","",IFERROR(XLOOKUP(${email},${N}!$A$2:$A,${N}!$B$2:$B),"not a member"))`;
    }
  };

  return ROSTER_COLUMNS.map(col => (col.formula ? col.formula(c) : null));
}

/**
 * Every non-empty PlayerID in Signups, sorted by Last Name then Preferred First
 * Name (case-insensitive, like the sheet's own SORT). Pure: takes the Signups
 * data rows (without the header) as plain values and the resolved Signups letters.
 */
function sortedSignupsPlayerIds(signupsRows, signupsLetters) {
  const index = (key) => columnLetterToIndex(signupsLetters[SIGNUPS_HEADERS[key]]) - 1;
  const idCol = index('playerId');
  const lastCol = index('lastName');
  const prefCol = index('preferredFirstName');
  const text = (v) => (v === null || v === undefined ? '' : v.toString().trim());
  const compare = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });
  return signupsRows
    .map(r => ({ id: text(r[idCol]), last: text(r[lastCol]), pref: text(r[prefCol]) }))
    .filter(r => r.id !== '')
    .sort((a, b) => compare(a.last, b.last) || compare(a.pref, b.pref) || compare(a.id, b.id))
    .map(r => r.id);
}

/**
 * Header note text for a ROSTER_COLUMNS entry.
 */
function rosterHeaderNote(col) {
  return `Type: ${col.type} / Source: ${col.source} / ${col.note}`;
}

/**
 * Generate Fresh Roster: rewrite the 📋 Roster as one header row plus one row per
 * Player in Signups: the PlayerID as a plain value in column A and a per-row
 * formula in every other column, all derived from Signups, Extra Player Info,
 * Final Forms, and Newsletter Subscribers. Creates the tab if missing. Existing
 * conditional formatting and filter views are left untouched. Because column A is
 * values, the Roster only learns about new or removed Signups when this runs again.
 */
function generateRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const signupsSheet = ss.getSheetByName(CONFIG.signups.sheetName);
  if (!signupsSheet) {
    ui.alert('Error',
      `Sheet "${CONFIG.signups.sheetName}" not found.\n\n` +
      `The Roster is derived from it. Add a tab with that exact name whose A1 is an IMPORTRANGE of the portal's Signups sheet, then run this again.`,
      ui.ButtonSet.OK);
    return;
  }
  const signupsValues = signupsSheet.getRange(1, 1, Math.max(1, signupsSheet.getLastRow()), Math.max(1, signupsSheet.getLastColumn())).getValues();
  const signupsLetters = resolveSignupsColumns(signupsValues[0]);
  const playerIds = sortedSignupsPlayerIds(signupsValues.slice(1), signupsLetters);
  const headers = ROSTER_COLUMNS.map(col => col.name);
  const notes = ROSTER_COLUMNS.map(rosterHeaderNote);
  if (ROSTER_COLUMNS.findIndex(col => !col.formula) !== 0 || ROSTER_COLUMNS.filter(col => !col.formula).length !== 1) {
    throw new Error('ROSTER_COLUMNS must have exactly one value column (PlayerID) and it must be first');
  }
  // One formula row per Player; every other column's formula references its own row.
  const formulaRows = playerIds.map((_, i) => buildRosterFormulas(signupsLetters, ROSTER_FIRST_DATA_ROW + i).slice(1));

  let rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(CONFIG.roster.sheetName);
    console.log(`Created sheet "${CONFIG.roster.sheetName}"`);
  }

  // Nothing is authored in the Roster, so clearing loses nothing: contents, notes,
  // and leftover data validations go; conditional formatting stays.
  rosterSheet.clearContents();
  rosterSheet.clearNotes();
  rosterSheet.getRange(1, 1, rosterSheet.getMaxRows(), rosterSheet.getMaxColumns()).clearDataValidations();
  rosterSheet.setFrozenRows(0);

  const neededRows = ROSTER_HEADER_ROW + Math.max(1, playerIds.length);
  if (rosterSheet.getMaxRows() < neededRows) {
    rosterSheet.insertRowsAfter(rosterSheet.getMaxRows(), neededRows - rosterSheet.getMaxRows());
  }
  if (rosterSheet.getMaxColumns() < headers.length) {
    rosterSheet.insertColumnsAfter(rosterSheet.getMaxColumns(), headers.length - rosterSheet.getMaxColumns());
  }

  const headerRange = rosterSheet.getRange(ROSTER_HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setNotes([notes]);
  headerRange.setFontWeight('bold');
  if (playerIds.length > 0) {
    rosterSheet.getRange(ROSTER_FIRST_DATA_ROW, 1, playerIds.length, 1).setValues(playerIds.map(id => [id]));
    rosterSheet.getRange(ROSTER_FIRST_DATA_ROW, 2, playerIds.length, headers.length - 1).setFormulas(formulaRows);
  }
  rosterSheet.setFrozenRows(ROSTER_HEADER_ROW);

  ROSTER_COLUMNS.forEach((col, index) => {
    if (col.type === 'Date' && playerIds.length > 0) {
      rosterSheet.getRange(ROSTER_FIRST_DATA_ROW, index + 1, playerIds.length, 1).setNumberFormat('yyyy-mm-dd');
    }
  });

  // Finish with Format Spruce Up's plain worker (not formatSpruceUp() itself, which
  // targets whatever sheet is active and ends in its own UI alert) so a fresh Roster
  // always comes out with banding, a filter, and frozen row/column already applied.
  applySpruceUpFormatting(rosterSheet);

  SpreadsheetApp.flush();
  console.log(`Roster generated: ${playerIds.length} Players, ${headers.length} columns, per-row formulas from row ${ROSTER_FIRST_DATA_ROW}`);

  ui.alert('Roster Generated',
    `"${CONFIG.roster.sheetName}" now has ${playerIds.length} Players and ${headers.length} columns: a header row (hover a header for its type, Source, and rule), PlayerIDs in column A, and a formula in every other cell keyed by that row's PlayerID. Format Spruce Up ran automatically: alternating row colors, a data filter, vertical centering, and frozen row 1 / column A.\n\n` +
    `Nothing is authored here; fix a wrong value in its Source (Signups, Extra Player Info, Final Forms, Newsletter Subscribers).\n\n` +
    `Sort and filter freely (filter views, the basic filter, or Data > Sort range); every row stays self-contained.\n\n` +
    `New or removed Signups only show up when you run Generate Fresh Roster again (Run Diagnostics tells you when the Roster is stale).\n\n` +
    `Next: run "Sync Extra Player Info" so every PlayerID has a row for Team, Returning, and the Include override.`,
    ui.ButtonSet.OK);
}

/**
 * Convert column letter(s) to a 1-based column number (inverse of getColumnLetter).
 */
function columnLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/**
 * Convert column number to letter(s)
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const modulo = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    columnNumber = Math.floor((columnNumber - modulo) / 26);
  }
  return letter;
}

/**
 * Create custom menu for easy access
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(`🥏 Madison Ultimate (v${SCRIPT_VERSION})`)
    .addItem('🩺 Run Diagnostics', 'runDiagnostics')
    .addSeparator()
    .addItem('📝 Generate Fresh Roster', 'generateRoster')
    .addItem('🧩 Sync Extra Player Info', 'syncExtraPlayerInfo')
    .addItem('🔢 Assign Tryout IDs', 'assignTryoutIds')
    .addSeparator()
    .addItem('🔄 Refresh All Data', 'refreshAllData')
    .addItem('📊 Update Final Forms', 'updateFinalForms')
    .addItem('📬 Update Newsletter Subscribers', 'updateNewsletterSubscribers')
    .addSeparator()
    .addItem('🏗️ Build Custom Sheet', 'buildCustomSheet')
    .addItem('🏅 Build Practice Roster', 'buildPracticeRoster')
    .addItem('🏆 Build Game Roster Prep Sheet', 'buildGameRosterPrepSheet')
    .addItem('⬆️ Apply Activation Status', 'showApplyActivationStatusDialog')
    .addItem('📧 Build Email List', 'buildEmailList')
    .addItem('🎨 Format Spruce Up', 'formatSpruceUp')
    .addItem('🧹 Delete Empty Rows & Columns', 'deleteEmptyRowsAndColumns')
    .addItem('🏃 Build Practice Availability', 'buildPracticeAvailability')
    .addItem('🎮 Build Game Availability', 'buildGameAvailability')
    .addItem('✅ Convert to Actual Attendance', 'convertToActualAttendance')
    .addItem('📅 Sync Practice Info to Calendar', 'createPracticeCalendarEvents')
    .addItem('📅 Sync Game Info to Calendar', 'createGameCalendarEvents')
    .addItem('📄 Export Game Info to Markdown List', 'exportGameInfoToMarkdown')
    .addItem('📋 Organize Sheets', 'organizeSheets')
    .addSeparator()
    .addItem('📈 Show Statistics', 'showStatistics')
    .addItem('🔍 Find Emails Not Subscribed to Newsletter', 'findMissingEmails')
    .addItem('👥 Caretakers Not Subscribed to Newsletter', 'findPendingParents')
    .addItem('🔎 Analyze Signups', 'analyzeSignups')
    .addItem('🔀 Full Name Diff', 'fullNameDiff')
    .addToUi();
}

/**
 * Refresh all data sources
 */
function refreshAllData() {
  updateFinalForms();
  updateNewsletterSubscribers();
  // Signups refreshes on its own through IMPORTRANGE; the Roster formulas pick everything up.
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Data Refreshed', 'Final Forms and Newsletter Subscribers have been updated. Signups refreshes on its own through IMPORTRANGE.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Helper function to find the most recent CSV file in a folder
 * @param {string} folderId - The Google Drive folder ID
 * @returns {File|null} - The most recent CSV file, or null if none found
 */
function findMostRecentCsvFile(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  let mostRecentFile = null;
  let mostRecentDate = new Date(0); // Start with epoch time
  
  // Find the most recent CSV file
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().toLowerCase();
    
    // Only consider CSV files
    if (fileName.endsWith('.csv')) {
      const fileDate = file.getLastUpdated();
      if (fileDate > mostRecentDate) {
        mostRecentDate = fileDate;
        mostRecentFile = file;
      }
    }
  }
  
  return mostRecentFile;
}

/**
 * Helper function to calculate and display differences between old and new data
 * @param {Array} oldData - Previous data array
 * @param {Array} newData - New data array  
 * @param {string} dataType - Type of data (for logging)
 */
function reportDataDifferences(oldData, newData, dataType) {
  const oldCount = oldData ? oldData.length - 1 : 0; // Subtract 1 for header
  const newCount = newData.length - 1; // Subtract 1 for header
  
  const difference = newCount - oldCount;
  const diffSign = difference > 0 ? '+' : '';
  
  console.log(`📊 ${dataType} Import Summary:`);
  console.log(`   Previous: ${oldCount} rows`);
  console.log(`   New: ${newCount} rows`);
  console.log(`   Change: ${diffSign}${difference} rows`);
  
  // Show detailed row changes for debugging
  if (oldData && newData.length > 1) {
    console.log(`🔍 ${dataType} Row Changes:`);
    
    // Create maps using row key (first non-empty column) for comparison
    const oldRowMap = new Map();
    const newRowMap = new Map();
    
    // Build old data map (skip header row)
    if (oldData.length > 1) {
      for (let i = 1; i < oldData.length; i++) {
        const row = oldData[i];
        const key = getRowKey(row, i);
        oldRowMap.set(key, { data: row, index: i });
      }
    }
    
    // Build new data map (skip header row)
    for (let i = 1; i < newData.length; i++) {
      const row = newData[i];
      const key = getRowKey(row, i);
      newRowMap.set(key, { data: row, index: i });
    }
    
    const addedRows = [];
    const removedRows = [];
    const modifiedRows = [];
    
    // Find added and modified rows
    for (const [key, newRow] of newRowMap) {
      if (!oldRowMap.has(key)) {
        // New row
        addedRows.push({ key, data: newRow.data });
      } else {
        // Check if row content changed
        const oldRow = oldRowMap.get(key);
        if (JSON.stringify(oldRow.data) !== JSON.stringify(newRow.data)) {
          modifiedRows.push({
            key,
            oldData: oldRow.data,
            newData: newRow.data
          });
        }
      }
    }
    
    // Find removed rows
    for (const [key, oldRow] of oldRowMap) {
      if (!newRowMap.has(key)) {
        removedRows.push({ key, data: oldRow.data });
      }
    }
    
    // Log added rows
    if (addedRows.length > 0) {
      console.log(`   ➕ Added ${addedRows.length} rows:`);
      addedRows.forEach(({ key, data }) => {
        console.log(`      + ${key}`);
      });
    }
    
    // Log removed rows
    if (removedRows.length > 0) {
      console.log(`   ➖ Removed ${removedRows.length} rows:`);
      removedRows.forEach(({ key, data }) => {
        console.log(`      - ${key}`);
      });
    }
    
    // Log modified rows with before/after
    if (modifiedRows.length > 0) {
      console.log(`   ✏️ Modified ${modifiedRows.length} rows:`);
      modifiedRows.forEach(({ key, oldData, newData }) => {
        console.log(`      📝 ${key}:`);
        
        // Compare each column to show what changed
        const maxColumns = Math.max(oldData.length, newData.length);
        for (let col = 0; col < maxColumns; col++) {
          const oldValue = oldData[col] || '';
          const newValue = newData[col] || '';
          
          if (oldValue !== newValue) {
            console.log(`         Column ${col + 1}: "${oldValue}" → "${newValue}"`);
          }
        }
      });
    }
    
    if (addedRows.length === 0 && removedRows.length === 0 && modifiedRows.length === 0) {
      console.log(`   ✅ No changes detected - data is identical`);
    }
  } else if (!oldData) {
    console.log(`   🆕 Initial import - no previous data to compare`);
  }
  
  return {
    oldCount,
    newCount, 
    difference
  };
}

/**
 * Helper function to get a unique key for a row (used for row matching)
 * @param {Array} row - The row data array
 * @param {number} index - Row index as fallback
 * @returns {string} - Unique key for the row
 */
function getRowKey(row, index) {
  // Try to find a good identifier from the first few columns
  for (let i = 0; i < Math.min(3, row.length); i++) {
    const value = row[i];
    if (value && value.toString().trim()) {
      return value.toString().trim();
    }
  }
  
  // Fallback to row index if no good identifier found
  return `Row_${index}`;
}


/**
 * Update Final Forms data
 */
function updateFinalForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.finalForms.sheetName);
  
  if (!sheet) {
    const existingSheetNames = ss.getSheets().map(s => s.getName()).join(', ');
    SpreadsheetApp.getUi().alert('Error',
      `Sheet "${CONFIG.finalForms.sheetName}" not found.\n\n` +
      `This sheet is required: the roster's XLOOKUP formulas read from it, and this menu item writes the imported CSV into it.\n\n` +
      `Fix: add a tab named exactly "${CONFIG.finalForms.sheetName}" (blank is fine, this will populate it), then run this again.\n\n` +
      `Existing tabs: ${existingSheetNames}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  try {
    // Get existing data for comparison
    const lastRow = sheet.getLastRow();
    const oldData = lastRow > 0 ? sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues() : null;
    
    // Get the most recent CSV file from the Final Forms folder
    const mostRecentFile = findMostRecentCsvFile(CONFIG.finalForms.folderId);
    
    if (!mostRecentFile) {
      SpreadsheetApp.getUi().alert('Error', 'No CSV files found in the Final Forms folder.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Using most recent Final Forms file: ${mostRecentFile.getName()} (${mostRecentFile.getLastUpdated()})`);
    
    const csvData = mostRecentFile.getBlob().getDataAsString();
    const csvArray = Utilities.parseCsv(csvData);
    
    // Report differences
    const diff = reportDataDifferences(oldData, csvArray, 'Final Forms');
    
    sheet.clear();
    if (csvArray.length > 0) {
      sheet.getRange(1, 1, csvArray.length, csvArray[0].length).setValues(csvArray);
    }
    
    const fileName = mostRecentFile.getName();
    const studentCount = csvArray.length - 1; // Subtract 1 for header row
    
    console.log(`✅ Updated Final Forms from: ${fileName}`);
    
    SpreadsheetApp.getUi().alert('Final Forms Updated', 
      `Successfully imported ${studentCount} students from:\n${fileName}\n\nChange: ${diff.difference >= 0 ? '+' : ''}${diff.difference} students\n\nThe Roster's Final Forms columns update on their own (they look up each row's SPS Student ID).`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (e) {
    console.error('Error updating Final Forms:', e);
    SpreadsheetApp.getUi().alert('Error', 'Could not update Final Forms data. Check the folder and file permissions.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Read the Roster header row plus every data row in one call.
 * Returns { headers, rows, col(name) } where col(name) is the 0-based index of a
 * header (throws when a required header is missing).
 */
function readRosterTable(rosterSheet) {
  const lastRow = rosterSheet.getLastRow();
  const lastCol = rosterSheet.getLastColumn();
  const values = lastRow >= ROSTER_HEADER_ROW && lastCol > 0
    ? rosterSheet.getRange(ROSTER_HEADER_ROW, 1, lastRow - ROSTER_HEADER_ROW + 1, lastCol).getValues()
    : [[]];
  const headers = values[0].map(h => (h === null || h === undefined) ? '' : h.toString().trim());
  const rows = values.slice(ROSTER_FIRST_DATA_ROW - ROSTER_HEADER_ROW);
  const col = (name) => {
    const index = headers.indexOf(name);
    if (index === -1) throw new Error(`Cannot find required Roster column "${name}". Run "Generate Fresh Roster" first.`);
    return index;
  };
  return { headers, rows, col };
}

/**
 * Show statistics about the roster
 */
function showStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);

  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Roster sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const table = readRosterTable(rosterSheet);
  const playerIdCol = table.col(CONFIG.columns.playerId);
  const players = table.rows.filter(row => row[playerIdCol] !== '' && row[playerIdCol] !== null);
  const total = players.length;

  const countTrue = (name) => {
    const index = table.col(name);
    return players.filter(row => row[index] === true).length;
  };
  const countRegular = (name) => {
    const index = table.col(name);
    return players.filter(row => row[index] === 'regular').length;
  };
  const pct = (n) => total ? `${Math.round(n / total * 100)}%` : '0%';

  const stats = {
    profileComplete: countTrue(CONFIG.columns.profileComplete),
    finalFormsCleared: countTrue(CONFIG.columns.finalFormsCleared),
    parentSigned: countTrue(CONFIG.columns.areAllFormsParentSigned),
    studentSigned: countTrue(CONFIG.columns.areAllFormsStudentSigned),
    physicalCleared: countTrue(CONFIG.columns.physicalCleared),
    includeInGeneratedRosters: countTrue(CONFIG.columns.includeInGeneratedRosters),
    caretaker1Regular: countRegular(CONFIG.columns.caretaker1NewsletterStatus),
    caretaker2Regular: countRegular(CONFIG.columns.caretaker2NewsletterStatus),
    grades: {}
  };
  const gradeCol = table.col(CONFIG.columns.grade);
  players.forEach(row => {
    const grade = row[gradeCol];
    if (grade !== '' && grade !== null) stats.grades[grade] = (stats.grades[grade] || 0) + 1;
  });

  let message = `📊 Roster Statistics\n\n`;
  message += `Total Players (Signups rows): ${total}\n`;
  message += `Profile Complete: ${stats.profileComplete} (${pct(stats.profileComplete)})\n`;
  message += `Include In Generated Rosters: ${stats.includeInGeneratedRosters} (${pct(stats.includeInGeneratedRosters)})\n\n`;

  message += `Final Forms:\n`;
  message += `  Parent Signed: ${stats.parentSigned} (${pct(stats.parentSigned)})\n`;
  message += `  Student Signed: ${stats.studentSigned} (${pct(stats.studentSigned)})\n`;
  message += `  Physical Cleared: ${stats.physicalCleared} (${pct(stats.physicalCleared)})\n`;
  message += `  Final Forms Cleared: ${stats.finalFormsCleared} (${pct(stats.finalFormsCleared)})\n\n`;

  message += `Newsletter (status "regular"):\n`;
  message += `  Caretaker 1 subscribed: ${stats.caretaker1Regular}\n`;
  message += `  Caretaker 2 subscribed: ${stats.caretaker2Regular}\n\n`;

  message += `Grade Distribution:\n`;
  Object.keys(stats.grades).sort((a, b) => Number(a) - Number(b)).forEach(grade => {
    message += `  Grade ${grade}: ${stats.grades[grade]} players\n`;
  });

  SpreadsheetApp.getUi().alert('Madison Ultimate Roster Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Find all email addresses in roster that are not Buttondown newsletter subscribers
 * Excludes Seattle School email addresses
 */
function findMissingEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);
  const mailingSheet = ss.getSheetByName(CONFIG.newsletterSubscribers.sheetName);

  if (!rosterSheet || !mailingSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find Roster or Newsletter Subscribers sheets');
    return;
  }

  // Get all email addresses from Newsletter Subscribers (column A, header in row 1, data from row 2)
  const mailingListData = mailingSheet.getRange('A2:A').getValues();
  const mailingListEmails = new Set(
    mailingListData
      .flat()
      .filter(email => email && email.toString().trim())
      .map(email => email.toString().toLowerCase().trim())
  );

  console.log(`Found ${mailingListEmails.size} emails in Newsletter Subscribers`);

  const table = readRosterTable(rosterSheet);
  const fullNameCol = table.col(CONFIG.columns.fullName);

  // Email columns are the ones whose header ends with "Email" (the "... Newsletter Status" columns hold statuses, not addresses)
  const emailColumns = [];
  table.headers.forEach((header, index) => {
    if (/email$/i.test(header)) {
      emailColumns.push(index);
      console.log(`Found email column: ${header} at column ${index + 1}`);
    }
  });

  const uniqueRosterEmails = new Set();
  const missingEmails = [];

  emailColumns.forEach(colIndex => {
    table.rows.forEach((row, rowIndex) => {
      const email = row[colIndex];
      if (!email || !email.toString().trim()) return;
      const emailStr = email.toString().trim();
      const emailLower = emailStr.toLowerCase();

      // Skip Seattle School email addresses
      if (emailLower.includes('@seattleschools.org')) return;

      if (!mailingListEmails.has(emailLower) && !uniqueRosterEmails.has(emailLower)) {
        uniqueRosterEmails.add(emailLower);
        missingEmails.push({
          email: emailStr,
          name: (row[fullNameCol] || '').toString().trim(),
          source: table.headers[colIndex],
          row: rowIndex + ROSTER_FIRST_DATA_ROW
        });
      }
    });
  });

  // Sort missing emails alphabetically
  missingEmails.sort((a, b) => a.email.localeCompare(b.email));

  // Display results
  if (missingEmails.length === 0) {
    SpreadsheetApp.getUi().alert(
      'All Emails Subscribed',
      'Every roster email address is a Newsletter subscriber (Seattle Schools emails excluded).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    let message = `Found ${missingEmails.length} email addresses not on the Newsletter:\n\n`;

    const bySource = {};
    missingEmails.forEach(item => {
      if (!bySource[item.source]) bySource[item.source] = [];
      bySource[item.source].push(item);
    });

    Object.keys(bySource).sort().forEach(source => {
      message += `\n${source}:\n`;
      bySource[source].forEach(item => {
        message += `  • ${item.email} (${item.name})\n`;
      });
    });

    message += '\n\nYou can copy these addresses to invite them to the Newsletter.';

    const emailList = missingEmails.map(item => item.email).join(', ');
    message += `\n\nComma-separated list for easy copying:\n${emailList}`;

    // Show in a dialog (alert has size limits, so using custom HTML dialog for long lists)
    if (missingEmails.length > 10) {
      showMissingEmailsDialog(missingEmails, emailList);
    } else {
      SpreadsheetApp.getUi().alert(
        'Emails Not on the Newsletter',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

/**
 * Show missing emails in a scrollable HTML dialog for long lists
 */
function showMissingEmailsDialog(missingEmails, emailList) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 10px; }
      h3 { color: #1a73e8; }
      .email-group { margin-bottom: 20px; }
      .email-item { margin: 5px 0; padding: 5px; background: #f8f9fa; }
      .copy-section { 
        margin-top: 20px; 
        padding: 10px; 
        background: #e8f0fe; 
        border-radius: 5px;
      }
      textarea { 
        width: 100%; 
        height: 100px; 
        margin-top: 10px;
        font-family: monospace;
      }
      .stats { color: #5f6368; margin-bottom: 15px; }
    </style>
    <div>
      <h3>Emails Not on the Newsletter</h3>
      <div class="stats">Found ${missingEmails.length} email addresses (Seattle Schools excluded)</div>
      
      <div class="copy-section">
        <strong>All emails (comma-separated):</strong>
        <textarea readonly onclick="this.select()">${emailList}</textarea>
      </div>
      
      <h4>Detailed List by Source:</h4>
      ${Object.entries(
        missingEmails.reduce((acc, item) => {
          if (!acc[item.source]) acc[item.source] = [];
          acc[item.source].push(item);
          return acc;
        }, {})
      ).map(([source, items]) => `
        <div class="email-group">
          <strong>${source}:</strong>
          ${items.map(item => `
            <div class="email-item">
              ${item.email} - ${item.name} (Row ${item.row})
            </div>
          `).join('')}
        </div>
      `).join('')}
    </div>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Emails Not on the Newsletter');
}


/**
 * Find every Caretaker who is not an active Buttondown Newsletter subscriber
 * Shows those with any status other than "regular" (includes "unactivated", "unsubscribed", and "not a member")
 */
function findPendingParents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName(CONFIG.roster.sheetName);

  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find Roster sheet');
    return;
  }

  const table = readRosterTable(rosterSheet);
  const fullNameCol = table.col(CONFIG.columns.fullName);
  const caretakers = [
    { label: 'Caretaker 1', name: table.col(CONFIG.columns.caretaker1Name), email: table.col(CONFIG.columns.caretaker1Email), status: table.col(CONFIG.columns.caretaker1NewsletterStatus) },
    { label: 'Caretaker 2', name: table.col(CONFIG.columns.caretaker2Name), email: table.col(CONFIG.columns.caretaker2Email), status: table.col(CONFIG.columns.caretaker2NewsletterStatus) }
  ];

  const pendingParents = [];
  const seenEmails = new Set(); // Avoid duplicates

  table.rows.forEach(row => {
    const student = (row[fullNameCol] || '').toString().trim();
    if (!student) return;

    caretakers.forEach(caretaker => {
      const status = row[caretaker.status];
      const email = row[caretaker.email];
      if (!status || status === 'regular' || !email) return;
      const emailStr = email.toString().trim();
      if (!emailStr || seenEmails.has(emailStr.toLowerCase())) return;
      seenEmails.add(emailStr.toLowerCase());
      pendingParents.push({
        name: (row[caretaker.name] || '').toString().trim(),
        email: emailStr,
        status: status.toString(),
        student: student,
        parentType: caretaker.label
      });
    });
  });

  // Sort by caretaker name, then email
  pendingParents.sort((a, b) => {
    const nameCompare = a.name.localeCompare(b.name);
    return nameCompare !== 0 ? nameCompare : a.email.localeCompare(b.email);
  });

  if (pendingParents.length === 0) {
    SpreadsheetApp.getUi().alert(
      'All Caretakers Subscribed',
      'Every Caretaker email has status "regular" on the Newsletter.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    showPendingParentsDialog(pendingParents);
  }
}

/**
 * Show Caretakers not subscribed to the Newsletter in a modal dialog with HTML table for easy copy/paste
 */
function showPendingParentsDialog(pendingParents) {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 15px; }
      h3 { color: #1a73e8; margin-bottom: 5px; }
      .stats { color: #5f6368; margin-bottom: 20px; font-size: 14px; }
      
      .table-container { 
        background: white;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        overflow: hidden;
        margin-bottom: 20px;
      }
      
      table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
      }
      
      th {
        background: #f8f9fa;
        color: #202124;
        font-weight: 600;
        padding: 12px 16px;
        text-align: left;
        border-bottom: 2px solid #e8eaed;
      }
      
      td {
        padding: 10px 16px;
        border-bottom: 1px solid #e8eaed;
        vertical-align: top;
      }
      
      tr:hover {
        background: #f8f9fa;
      }
      
      .status-invited {
        background: #fff3cd;
        color: #856404;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
      }
      
      .status-not-member {
        background: #f8d7da;
        color: #721c24;
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 500;
      }
      
      .instructions {
        background: #e8f0fe;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 15px;
        font-size: 13px;
      }
      
      .copy-instructions {
        color: #5f6368;
        font-style: italic;
      }
    </style>
    <div>
      <h3>Caretakers Not Subscribed to Newsletter</h3>
      <div class="stats">Found ${pendingParents.length} Caretaker emails whose Newsletter status is not "regular"</div>
      
      <div class="instructions">
        <strong>Instructions:</strong> Select the table below and copy (Ctrl+C / Cmd+C) to paste into spreadsheets or emails.
        <br><span class="copy-instructions">The table will copy with proper formatting and can be pasted directly into Excel, Google Sheets, or email.</span>
      </div>
      
      <div class="table-container">
        <table id="parentTable">
          <thead>
            <tr>
              <th>Caretaker</th>
              <th>Email Address</th>
              <th>Status</th>
              <th>Player</th>
              <th>Which</th>
            </tr>
          </thead>
          <tbody>
            ${pendingParents.map(parent => `
              <tr>
                <td>${parent.name}</td>
                <td>${parent.email}</td>
                <td><span class="status-${parent.status.replace(' ', '-')}">${parent.status}</span></td>
                <td>${parent.student}</td>
                <td>${parent.parentType}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      
      <div style="margin-top: 15px; font-size: 12px; color: #5f6368;">
        <strong>Tip:</strong> You can select individual rows or the entire table and copy to paste elsewhere.
      </div>
    </div>
    
    <script>
      // Auto-select table when clicked for easy copying
      document.getElementById('parentTable').addEventListener('click', function() {
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(this);
        selection.removeAllRanges();
        selection.addRange(range);
      });
    </script>
  `)
    .setWidth(700)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Caretakers Not Subscribed to Newsletter');
}


/**
 * Run on spreadsheet open to create menu
 */
function onOpen() {
  createCustomMenu();
}
