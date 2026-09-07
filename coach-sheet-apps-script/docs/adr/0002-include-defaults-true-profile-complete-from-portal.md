---
status: accepted
date: 2026-09-07
---

# Include In Generated Rosters defaults to TRUE; Profile Complete comes from the portal

ADR 0001 made Profile Complete a Roster formula (Grade, Date of Birth, and Caretaker 1 Email present) and used it as the fallback for Include In Generated Rosters. The portal's Seed Signups from Final Forms (portal ADR 0006) fills all three of those fields from Final Forms when it creates a Seeded Signup, so under that formula every seeded row would be complete the moment it existed and the column would stop meaning anything. We decided that Profile Complete is defined once, by the portal (Player Info, Caretaker Info, and Photo Upload all done), written on the Signups row, and passed through into the Roster; the coach sheet no longer computes its own. Separately, Include In Generated Rosters no longer references Profile Complete at all: it resolves to the coach's Extra Player Info value when one is set and otherwise to TRUE, so every Player, seeded or not, is on Generated Rosters until a coach says otherwise.

## Considered options

- **Keep the lenient formula under a new name and let Include fall back to it.** Rejected as one more state to explain; a coach who wants a Player off a roster sets Include to FALSE, which already exists.
- **Let Include fall back to the portal's stricter Profile Complete.** Rejected for now: a family that has not uploaded a photo or answered every profile question should not keep their player off a practice roster. Coupling the two can be revisited without touching the portal.

## Consequences

- Blank Include In Generated Rosters now means "included"; the Extra Player Info header note and dropdown help say so.
- Analyze Signups section 5 reads the passthrough, and a new section 6 lists Seeded Signups the family has not finished, the outreach follow-up list.
- The Signups mirror must carry the portal's `Seeded At` and `Profile Complete` headers; Run Diagnostics checks for them.
