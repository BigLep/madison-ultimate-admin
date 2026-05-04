# Madison Ultimate Admin

Tools for managing the Madison Middle School Ultimate Frisbee team roster and communications.

## Projects

### [Coach Sheet Apps Script](./coach-sheet-apps-script/)
Google Apps Script tools for managing the team roster spreadsheet. Handles player registration data, parent contacts, mailing list status, practice/game availability, and various roster views.

**Double headers:** Two (or more) games on the same calendar day are supported. Use one row per game in **Game Info**, run **Build Game Availability** to create `M/D …` columns for the first game and `M/D … (Game 2)` (etc.) for later games that day, and use **Build Game Roster Prep** with the picker to choose which game. Details: [coach-sheet-apps-script/README.md — Double headers](./coach-sheet-apps-script/README.md#double-headers).

### [Photo Mapper](./photo-mapper/)
Web application for mapping team photos to players. Uses a NextJS frontend with a Flask backend to match photos from Google Drive to roster data.
