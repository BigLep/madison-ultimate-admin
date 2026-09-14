# Coach Sheet

The coach-facing Google Sheets workbook for a Madison Ultimate season and the Apps Script menu bound to it. It joins data mastered elsewhere (the family portal's Signups sheet, the district's Final Forms export, Buttondown) with a small set of coach-authored facts, and derives the roster and every printout from that join. Shared terms (Player, PlayerID, SPS Student ID, Caretaker, Newsletter, Final Forms Join, Final Forms Backfill, Seed Signups from Final Forms, Seeded Signup, Seeded Field, Profile Complete, Player Photo, Media Opt-Out) are defined in the portal's glossary at `../../madison-ultimate/CONTEXT.md` and mean the same thing here.

## Language

### Sources

**Source**:
One of the four places an authored fact lives: Signups (family-authored through the portal), Extra Player Info (coach-authored), Final Forms (district), Newsletter Subscribers (Buttondown). Every value in the Roster traces to exactly one Source; nothing is authored in the Roster itself.
_Avoid_: import, raw data

**Signups**:
The coach sheet's read-only mirror of the portal's Signups sheet: one row per Player, keyed by PlayerID. The list of Players for the season is, by definition, the set of rows here.
_Avoid_: Additional Info (the retired Google Form this replaces), signup form

**Extra Player Info**:
The one tab where coaches author per-player facts the family cannot: Team, Returning, and the Include In Generated Rosters override. Keyed by PlayerID; one row per Player, rows added by sync rather than typed.
_Avoid_: manual columns, roster overrides

**Final Forms**:
The latest nightly export of the district registration system, keyed by SPS Student ID. The only Source for signature and physical clearance facts, and the preferred Source for Grade.

**Newsletter Subscribers**:
The current Buttondown subscriber list (email and status), the Source for every newsletter status column.
_Avoid_: mailing list, Google Group

### Roster and derived sheets

**Roster**:
The derived, formula-only view of every Player, one row per Signups row, joined to the other Sources by PlayerID and SPS Student ID. Coaches read it and never edit it; a wrong value is fixed in its Source.
_Avoid_: master list, player list

**Full Name**:
Preferred First Name followed by Last Name. The human-readable key every Generated Roster uses to refer to a Player. On an Availability Sheet it is a per-row Roster formula keyed by the row's PlayerID, never a typed value, so it cannot drift from the Roster; the prep sheets find it by header. The portal never matches on it: it reads and writes availability cells by PlayerID.
_Avoid_: name, display name

**Availability Sheet**:
Practice Availability or Game Availability: one row per Player, PlayerID in column A as the only typed per-player value, then Full Name, Grade, and Gender Identification as Roster lookup formulas in the Roster's own shape (`=IF($A2="","",IFERROR(XLOOKUP($A2,'📋 Roster'!$A:$A,'📋 Roster'!$F:$F),""))`), then the date columns families fill in through the portal. Rows are seeded by Build Practice/Game Availability and never deleted or reordered. See ADR 0004.
_Avoid_: availability tracker, roster copy

**Generated Roster**:
A practice roster or game roster prep sheet built from the Roster for one date, keyed by Full Name. Includes a Player only when Include In Generated Rosters resolves to TRUE.
_Avoid_: printout, roster copy

**Include In Generated Rosters**:
Whether a Player appears on Generated Rosters. Resolves to the coach's Extra Player Info value when one is set, otherwise TRUE: every Player is included until a coach says otherwise.
_Avoid_: active, dropped (those describe per-game Activation Status, a different concept)

**Profile Complete**:
As defined by the portal: Player Info, Caretaker Info, and Photo Upload all done, written on the Signups row and passed through here, never computed in the coach sheet. A row abandoned at step 0, or a Seeded Signup the family has not finished, is not Profile Complete and stays visible in the Roster so a coach can follow up.
_Avoid_: registered, signed up

**Team**:
The coach-assigned squad for the season, authored in Extra Player Info after tryouts. Fall 2026 values: Blue, Gold, Silver, Practice Squad; TBD until assigned. A Game Info row carries the Team it belongs to, blank meaning every team.

**Analyze Signups**:
The coach sheet's report of Players whose Sources disagree or are incomplete: signups with no SPS Student ID, Final Forms students not yet seeded or joined, signups whose SPS Student ID has left Final Forms, suspected duplicate signups, signups not Profile Complete, and Seeded Signups the family has not finished. It only reports; the portal's Seed Signups from Final Forms does the joining and seeding.
_Avoid_: reconciliation, Additional Info Analysis (the retired name-matching report this replaces)

**Returning**:
Whether the Player played for Madison Ultimate in a prior season, authored in Extra Player Info.
_Avoid_: veteran, alumni

### Eligibility and identity facts

**Final Forms Cleared**:
TRUE only when Final Forms shows all forms parent signed, all forms student signed, and physical clearance Cleared. Any missing SPS Student ID or missing export row makes it FALSE.
_Avoid_: eligible, cleared (bare, which also names an unrelated Final Forms export column)

**Signup Gender**:
The family's Gender Identification from Signups collapsed to Gx or Bx, blank when not supplied.

**Gender Default Handling**:
A Roster flag, TRUE when nothing about gender needs a coach's attention. FALSE when Final Forms Gender and Signup Gender disagree, or when the Player's Pronouns include anything outside he/him for a Bx or she/her for a Gx: a prompt to check in with the Player, not a verdict about them. Like every Roster Boolean, TRUE means all is well and FALSE means something needs attention.
_Avoid_: gender mismatch, gender flag, Gender Special Attention (the inverted earlier name)

**Media OK**:
TRUE when photos of the Player may appear in team communications; FALSE when the family declared a Media Opt-Out.
_Avoid_: Media Opt-Out as a column name (the Signups field keeps that name; the Roster column is inverted so TRUE means all is well)

**Gender Identification**:
Gx or Bx for the Player: Signup Gender when supplied, otherwise Final Forms gender mapped Female to Gx and Male to Bx, otherwise blank. The value Generated Rosters print.
_Avoid_: gender (ambiguous between the Final Forms field and this derived value)
