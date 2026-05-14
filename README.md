# Madison Ultimate Admin

Tools for managing the Madison Middle School Ultimate Frisbee team roster and communications.

## Projects

### [Coach Sheet Apps Script](./coach-sheet-apps-script/)
Google Apps Script tools for managing the team roster spreadsheet. Handles player registration data, parent contacts, mailing list status, practice/game availability, and various roster views.

**Multiple events on the same calendar day:** The coach sheet supports **more than one game on a single calendar date** (often called a *double-header*). Use **one Game Info row per game** (same date repeated), run **Build Game Availability** so Game Availability gets `M/D …` columns for the first game and `M/D … (Game 2)` (etc.) for each additional game that day. **Build Game Roster Prep** still lists one picker row per Game Info line, but the prep sheet always includes **every game that calendar day** (activation, availability, and note for each). Full workflow: [coach-sheet-apps-script/README.md — Multiple events on one calendar day](./coach-sheet-apps-script/README.md#multiple-events-on-the-same-calendar-day-double-headers).

### [Photo Mapper](./photo-mapper/)
Web application for mapping team photos to players. Uses a NextJS frontend with a Flask backend to match photos from Google Drive to roster data.
