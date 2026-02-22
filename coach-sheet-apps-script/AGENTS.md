# Agent Instructions for Apps Script Deployment

## Before Deploying

**ALWAYS** increment the `SCRIPT_VERSION` constant in `Code.gs` before running `clasp push`.

Use version format `2.x`; increment x for each release (e.g. `'2.0'` → `'2.1'`).

## Deployment Process

1. Update `SCRIPT_VERSION` in `Code.gs`
2. Run `clasp push` to deploy changes
3. Test the deployed functionality

## Commits

**All changes** to this project must be committed using [Conventional Commits](https://www.conventionalcommits.org/): use a type and optional scope (e.g. `feat(menu): add X`, `fix(roster): correct Y`, `chore(coach-sheet): Z`), and add a short body when it helps.

## Notes

- Do NOT modify the version field in `appsscript.json` - use `SCRIPT_VERSION` in `Code.gs` instead
- The version helps track which deployment is currently active
