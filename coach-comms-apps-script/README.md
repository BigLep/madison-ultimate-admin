# Coach Comms

Apps Script bound to the season's Communications Doc. Adds a menu item that converts whichever Newsletter Block the cursor is inside into a Buttondown Draft, so a coach never has to copy/paste rich text out of Google Docs by hand. See `CONTEXT.md` for the vocabulary (Communications Doc, Newsletter Block, Buttondown Draft) and `../CONTEXT-MAP.md` for how this relates to `coach-sheet-apps-script`.

## Setup

1. Bind this script to the Communications Doc: open the doc, Extensions > Apps Script, and either paste these files in or `clasp push` from here once `.clasp.json`'s `scriptId` points at that bound project.
2. Set the `BUTTONDOWN_API_KEY` script property (Extensions > Apps Script > Project Settings (gear) > Script Properties). Get the key from Buttondown: Settings > Programming > API Keys. It needs write access to create drafts (a read-only key, like the one `coach-sheet-apps-script` uses for the subscriber sync, will fail with a 401/403).
3. Reload the doc so the `🥏 Coach Comms` menu appears (bound scripts add their menu via the `onOpen` simple trigger, which only runs when the doc is opened or reloaded).

## Usage

Click anywhere inside a Newsletter Block (Insert > Building Blocks > Email Draft, addressed to `drafts@mg.buttondown.email`), then run `🥏 Coach Comms > 📬 Send Newsletter Block to Buttondown`. It creates an unsent Buttondown Draft and shows a link to review/send it there.

## Known limitations (v1)

- **Cursor-based, not a picker.** The menu action always acts on the block your cursor is currently inside; it doesn't scan the doc for other Newsletter Blocks. Simplest to build first per the grilling session; a picker (list every Newsletter Block, pick one) was the discussed alternative if this turns out to be error-prone in practice.
- **No "already sent" tracking.** Running this twice on the same block just creates a second Buttondown Draft. Cheap to notice and delete in Buttondown's dashboard; deliberately not built.
- **Images are attempted, not guaranteed.** An inline image is uploaded to Buttondown's `/v1/images` endpoint and referenced by its hosted URL. The exact multipart field name that endpoint expects (`image` here) is inferred from Buttondown's docs, not yet verified against a real upload. If it fails, the image is replaced with the placeholder text `<COPY PASTE IN IMAGE>` instead of erroring out — check for that placeholder in the draft before sending, and fix the field name in `ButtondownApi.gs`'s `uploadButtondownImage` once you've confirmed the correct one from a real failed attempt's error message.
- **No Markdown escaping.** Plain prose containing literal `*`, `_`, `[`, or `]` will be misinterpreted as Markdown formatting by Buttondown. Not seen in either real Newsletter Block sampled while designing this, but worth knowing if a draft renders oddly.
- **Draft URL is inferred.** The link shown after creating a draft (`https://buttondown.com/emails/<id>`) is a guess at Buttondown's dashboard URL pattern, not confirmed against their docs. If it 404s, the draft still exists; find it from the Buttondown dashboard's Drafts list instead.
