# Madison Ultimate Admin

Tools for managing the Madison Middle School Ultimate Frisbee team roster and communications.

## Projects

### [Coach Sheet Apps Script](./coach-sheet-apps-script/)
Google Apps Script tools for managing the team roster spreadsheet. Handles Signups and Final Forms registration data, Caretaker contacts, Newsletter subscription status, practice/game availability, and various roster views.

**Multiple events on the same calendar day:** The coach sheet supports **more than one game on a single calendar date** (often called a *double-header*). Use **one Game Info row per game** (same date repeated), run **Build Game Availability** so Game Availability gets `M/D …` columns for the first game and `M/D … (Game 2)` (etc.) for each additional game that day. **Build Game Roster Prep** still lists one picker row per Game Info line, but the prep sheet always includes **every game that calendar day** (activation, availability, and note for each). Full workflow: [coach-sheet-apps-script/README.md — Multiple events on one calendar day](./coach-sheet-apps-script/README.md#multiple-events-on-the-same-calendar-day-double-headers).

## Development setup

One-time per clone: `git config core.hooksPath .githooks` activates the pre-commit hook that runs the coach sheet regression harness (`node coach-sheet-apps-script/test/harness.js`) against staged changes whenever its script or test files are staged, blocking the commit on failure.
