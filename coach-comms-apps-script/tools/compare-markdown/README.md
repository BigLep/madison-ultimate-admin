# compare-markdown

A QA tool for `../../NewsletterBlock.gs`'s Markdown converter: runs a by-hand mirror of its logic against a real Communications Doc's actual content, side by side with real Turndown run on the same content's HTML export. Used to find and verify fixes to the converter (see the repo's commit history for `coach-comms-apps-script` for the bugs this caught: missing paragraph separation, missing headings, tight-vs-loose spacing).

## Why the doc content isn't checked in

`doc_raw.json` and `doc.html` are full dumps of the real Communications Doc, which contains other coach-authored content beyond Newsletter Blocks, including real family email addresses in at least one block. They're gitignored (see `.gitignore` in this folder) and must be fetched fresh each time, not committed.

## Usage

1. Fetch the two input files this needs, from the repo root (or anywhere `gog` is configured for the `madisonultimate@gmail.com` account):
   ```
   gog docs raw <communicationsDocId> --json > coach-comms-apps-script/tools/compare-markdown/doc_raw.json
   gog docs export <communicationsDocId> --format html --out coach-comms-apps-script/tools/compare-markdown/doc.html
   ```
2. Install dependencies (once): `npm install` from this folder.
3. Run the comparison: `npm run compare`, or `NEWSLETTER_BLOCK_TO=someone@example.com npm run compare` to compare Newsletter Blocks addressed elsewhere.
4. Delete `doc_raw.json`/`doc.html` when done if you're concerned about leaving real doc content on disk; they're gitignored either way.

## Keeping the mirror in sync

`compare.js`'s `rowToMarkdownMirror` is a manual port of `NewsletterBlock.gs`'s conversion logic to plain Node (operating on the raw Docs API JSON instead of `DocumentApp`, since Apps Script's object model doesn't exist outside Apps Script). It does not run the real script, so update it by hand alongside any real change to `NewsletterBlock.gs`, or it'll silently compare against stale logic.
