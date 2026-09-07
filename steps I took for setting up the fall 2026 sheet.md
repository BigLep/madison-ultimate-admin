## Steps I took for setting up the fall sheet

These are the steps I took for setting up the fall sheet. I will want an agent to actually process this and turn into better docs. It's just my dumping ground for now.

duplicated spring sheet
- this copied over the apps script
I deleted sheets that clearly didn't make sense (e.g., date specific sheets)
I also deleted spring reign sheets which arne't relevant for fall
- BUG this caused: the "Final Forms" tab got deleted in this cleanup, but it's a required data sheet, not leftover: updateFinalForms() in Code.gs writes the imported CSV into a sheet with that exact name, and the Roster sheet's XLOOKUP formulas (Code.gs, the finalForms column definitions) read from 'Final Forms'!A:A etc. Running "Update Final Forms" failed with "Final Forms sheet not found." Fixed by re-adding a blank "Final Forms" tab via `gog sheets add-tab`; Update Final Forms then repopulates it from the latest CSV.
- LESSON for future seasons: before deleting any sheet during season setup cleanup, check Code.gs for sheetName references (grep for the tab name) to make sure it isn't a required data sheet the script writes into or formulas read from.
I filled in the practice info sheet
I filled in basic Game info sheet
- I didn't initially set it up for multiple teams (gold and blue). Fall 2025 had that structure I can draw from in future.

I reviewed and update the columns for the roster

Decided this new season is v3.0 of the coach sheet script (currently at v2.45 in coach-sheet-apps-script/Code.gs).
It should be installed on the fall 2026 sheet: https://docs.google.com/spreadsheets/d/1kMg6OgT8gzZyZChS6MtobLZ6GjmcrieV_VRUTTiuinE/edit (Drive file name: "2026 Fall Coach Sheets")
DONE: updated coach-sheet-apps-script/.clasp.json scriptId to the fall 2026 sheet's bound script: 1mK-vQmzJ95TbtMkQAd3kgOViet4-7GqPbq6JBU88pJKwYZ08B2HUuvtS
- Its Apps Script project is still internally titled "2026 Spring Coach Sheet Admin" since duplicating the sheet duplicated the script project as-is without renaming it. Should rename that project title to something like "2026 Fall Coach Sheet Admin" for clarity (this is separate from the SCRIPT_VERSION constant).

RECURRING STEP for future seasons: after duplicating a coach sheet, you must manually grab the new bound script's scriptId and put it in .clasp.json before clasp push will target the right sheet.
How to get it: open the new sheet > Extensions > Apps Script. This opens the bound project's editor in a new tab; the scriptId is the long ID segment in that tab's URL (https://script.google.com/u/0/home/projects/<scriptId>/edit).
Why gog/API access can't get it: gog's Apps Script commands (get, content, run, create, pull, deployments, versions) all require a scriptId you already have; there is no "look up the script bound to spreadsheet X" lookup. A container-bound script is not exposed as a queryable child of its spreadsheet via the Drive API either, so there's no Drive search that surfaces it. The Apps Script editor UI is the only way to discover a new bound script's ID.
- Bump SCRIPT_VERSION in Code.gs to '3.0' as part of the next clasp push for this season (done)

DONE: fixed CONFIG.finalForms.folderId in coach-sheet-apps-script/Code.gs. It was still pointing at last season's "Final Forms" Drive folder (1SnWCxDIn3FxJCvd1JcWyoeoOMscEsQcW). Updated it to the fall 2026 folder, "2026 Fall Final Forms" (1WgD4hY0fIZlQEBt7ekOlHIECA-HgOMIZ), which matches the DRIVE_FOLDER_ID the finalforms-export GitHub Action already uploads nightly CSVs into. Bumped SCRIPT_VERSION to 3.1 and clasp pushed.

RECURRING STEP for future seasons: CONFIG.finalForms.folderId in Code.gs must be updated to the new season's FinalForms exports folder at the start of each season, in lockstep with the finalforms-export automation's DRIVE_FOLDER_ID (repo variable in BigLep/madison-ultimate-admin). These two values need to point at the same folder or "Update Final Forms" silently imports stale/wrong-season data.

REMINDER (standing rule now, not season-specific): every clasp push must be preceded by bumping SCRIPT_VERSION in Code.gs first.

DONE: updated the README's "Deploy to a new season's spreadsheet" section and added a new "Script Properties (secrets)" section, covering everything learned above (scriptId discovery has no API/CLI path, finalForms folder id must move in lockstep with finalforms-export, don't delete sheets without grepping Code.gs first, Script Properties aren't copied when duplicating a spreadsheet).

DONE: added a "🩺 Run Diagnostics" menu item (Diagnostics.gs) that checks required sheets exist, configured Drive folders/spreadsheets are reachable, and required Script Properties (Buttondown key) actually work. Intended as the last step of season setup: "run diagnostics to confirm everything is configured correctly." First real run against the fall 2026 sheet caught the missing "Mailing List" sheet as an (expected, legacy) warning, everything else passed.

DONE: added "📬 Update Newsletter Subscribers" (new file NewsletterSubscribers.gs), replacing the Google Groups CSV mailing-list import as the real mailing-list source of truth (Google Groups was retired spring 2026; see madison-ultimate/docs/fall-2026/signup-grill.md). Pulls the full subscriber list from the Buttondown API (paginated, GET https://api.buttondown.com/v1/subscribers) into a "Newsletter Subscribers" sheet (columns: Email, Status, plus any other scalar fields Buttondown returns, discovered dynamically rather than hardcoded). Requires the "BUTTONDOWN_API_KEY" script property (Steven set this himself directly in the Apps Script editor's Script Properties, not via chat, to keep the key out of the conversation transcript). Diagnostics now live-tests that this key can actually read subscribers, not just that the property is set.
- The old "📧 Update Mailing List (legacy)" menu item, its CONFIG.mailingList Drive-folder CSV import, findMissingEmails, findPendingParents, and the roster's "MailingList Email address" XLOOKUP formula columns were intentionally left untouched in this pass (Steven's call: "just get my read on the import first").

DONE: stripped all Google Groups CSV mailing-list plumbing out of Code.gs and rewired the two useful reports + roster columns to read from Newsletter Subscribers/Buttondown instead (Steven's call: "delete and replace," not delete-with-no-replacement). Specifically:
- Removed CONFIG.mailingList entirely (Drive folderId + sheetName), and the 3 dead CONFIG.columns.*OnMailingList constants (defined but never read anywhere).
- Removed updateMailingList() and the "📧 Update Mailing List (legacy)" menu item outright.
- The 3 roster columns ("Student Personal/Parent 1/Parent 2 Email On Mailing List?") now VLOOKUP against 'Newsletter Subscribers'!$A$2:$B instead of 'Mailing List'!$A$3:$C. Column names kept the same on purpose (same coach-facing feature), but the values they return changed: Buttondown's "regular"/"unactivated"/"unsubscribed" instead of Google Groups' "member"/"invited"/"not a member".
- findMissingEmails() now reads the Newsletter Subscribers sheet instead of Mailing List.
- findPendingParents()'s "pending" check changed from `status !== 'member'` to `status !== 'regular'`.
- BONUS FIX found while touching this: showStatistics()'s parent1OnList/parent2OnList counters compared the status column to the boolean `true`, which that column (a status string) never returns, so those counts were always silently 0 even before this migration. Fixed to compare against `'regular'`.
- Diagnostics.gs: removed the now-meaningless "Mailing List sheet" legacy warning check (CONFIG.mailingList no longer exists).
- README: removed all "Mailing List (Google Groups)" / legacy-import documentation, replaced with a "Newsletter Subscribers (Buttondown)" data source section; updated Required Sheets table, Menu Functions doc, Column Source Types table, Troubleshooting section, File Structure table.
- Bumped SCRIPT_VERSION to 3.5 and clasp pushed.
- Verified with `grep -rn "updateMailingList\|CONFIG\.mailingList" *.gs` returning zero matches.

STILL OPEN (don't lose track):
- Broader README refresh based on recent changes (menu docs, etc.) — Steven to specify what