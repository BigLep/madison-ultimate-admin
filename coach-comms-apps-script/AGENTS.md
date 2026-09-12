# Agent Instructions for Apps Script Deployment

## Before Deploying

**ALWAYS** increment the `SCRIPT_VERSION` constant in `Code.gs` before running `clasp push`. With every change you deploy, increment the minor version (x in `1.x`) first (e.g. `'1.0'` → `'1.1'`).

Use version format `1.x`; increment x for each release.

## Deployment Process

1. Update `SCRIPT_VERSION` in `Code.gs`
2. Run `clasp push` to deploy changes
3. Test the deployed functionality (open the Communications Doc, reload it so `onOpen` re-adds the menu, run "Send Newsletter Block to Buttondown" against a test block)

## Commits

**All changes** to this project must be committed using [Conventional Commits](https://www.conventionalcommits.org/): use a type and optional scope (e.g. `feat(comms): add X`, `fix(comms): correct Y`, `chore(coach-comms): Z`), and add a short body when it helps.

## Notes

- Do NOT modify the version field in `appsscript.json` - use `SCRIPT_VERSION` in `Code.gs` instead
- The version helps track which deployment is currently active
