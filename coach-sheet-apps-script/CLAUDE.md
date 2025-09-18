# Claude Instructions for Apps Script Deployment

## Before Deploying

**ALWAYS** increment the `SCRIPT_VERSION` constant in `Code.gs` before running `clasp push`.

The version number should be incremented as a string (e.g., '53' → '54').

## Deployment Process

1. Update `SCRIPT_VERSION` in `Code.gs`
2. Run `clasp push` to deploy changes
3. Test the deployed functionality

## Notes

- Do NOT modify the version field in `appsscript.json` - use `SCRIPT_VERSION` in `Code.gs` instead
- The version helps track which deployment is currently active