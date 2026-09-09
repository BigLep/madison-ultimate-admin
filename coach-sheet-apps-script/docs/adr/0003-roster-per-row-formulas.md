---
status: proposed
date: 2026-09-08
---

# Roster rows are per-row formulas over a written PlayerID column, not one array formula per column

ADR 0001 made the 📋 Roster a formula-only view keyed by Signups PlayerID and, in its first form, implemented it as a `SORT(FILTER(...))` key formula in A2 plus one `ARRAYFORMULA` per column in row 2. In use the next day the coach found that filter views did not work on the Roster at all and the sheet could not be sorted: Google Sheets cannot sort or filter rows whose values spill from a single array formula, and the basic filter's sort or Data > Sort range would physically move the key formula. Sorting and filtering the Roster is a daily need (by team, grade, gender, clearance, newsletter status), so the array-formula layout was replaced.

We decided the Roster keeps every other property of ADR 0001 (keyed by PlayerID, nothing authored, one Source per fact, columns defined once in code) but is written one row per Player: Generate Fresh Roster reads Signups, writes each non-empty PlayerID as a plain value in column A (sorted by Last Name then Preferred First Name), and writes a formula in every other cell of that row that looks up the PlayerID on its own row. Sibling references are row-relative (`$F2` on row 2), so an in-place sort moves a whole row and every cell stays consistent. The formulas are plain, fall 2025 style, with one rule: never apply a scalar function to a whole-column range, because a plain formula implicitly intersects it with its own row and every lookup misses (the first 3.21 push did exactly that with `TO_TEXT('Final Forms'!$A:$A)` and `LOWER('Newsletter Subscribers'!$A$2:$A)`, so every Final Forms and newsletter column came back empty). Instead the single lookup key is converted: Signups delivers the SPS Student ID as text and the Final Forms import stores StudentID as a number, so Final Forms lookups try `VALUE($B<row>)` first and fall back to the text key; newsletter lookups pass the email straight in, since XLOOKUP already matches case-insensitively.

## Considered options

- **Keep the array formulas and tell coaches to use filter views.** Rejected: filter views do not work on array-formula output either, which was the trigger for this change.
- **Per-row formulas with a formula-derived key** (`INDEX(SORT(FILTER(...)), n)` in each column A cell). Rejected: new Signups would appear without a script run, but every row's Player shifts whenever a signup is added or removed, so a sorted sheet silently rescrambles. A written key is stable.
- **Per-row formulas with a written key** (this decision; the spring 2026 approach, keyed by PlayerID instead of Final Forms StudentID). ADR 0001 rejected it because it "requires a script run to extend when players are added and leaves stale rows behind when players are removed". Both are addressed: Generate Fresh Roster rewrites every row (no stale rows survive a run), and Run Diagnostics compares the Roster's PlayerIDs with Signups and fails when they differ, so a stale Roster is visible and one menu click away from fixed.

## Consequences

- Coaches sort and filter freely: filter views, the basic filter, and Data > Sort range all work.
- The Roster is a snapshot of which Players exist. After a new signup or a removal, run Generate Fresh Roster; the row order resets to Last Name then Preferred First Name and any in-place sort is lost (filter views keep their own sort and survive the rewrite).
- Run Diagnostics no longer checks for the `=SORT(FILTER(` key formula; it checks that data rows have the per-row formula shape and that the Roster's PlayerIDs match Signups.
- `buildRosterFormulas(signupsLetters, row)` builds one row at a time; the PlayerID column's builder is `null`. Downstream readers are unchanged: they still find columns by header name from `ROSTER_FIRST_DATA_ROW`.
