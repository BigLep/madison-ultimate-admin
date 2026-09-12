# Context Map

## Contexts

- [Coach Sheet](./coach-sheet-apps-script/CONTEXT.md): the coach-facing roster workbook and its Apps Script menu, joining Signups, Extra Player Info, Final Forms, and Newsletter Subscribers into the derived Roster
- [Coach Comms](./coach-comms-apps-script/CONTEXT.md): converts a coach's drafted email in the season's Communications Doc into a Buttondown Draft

## Relationships

- **Coach Sheet ↔ Coach Comms**: both integrate with Buttondown, but for different purposes. Coach Sheet reads the Newsletter Subscribers list (who receives newsletters); Coach Comms creates Buttondown Drafts (what gets sent). Neither reads the other's data.
