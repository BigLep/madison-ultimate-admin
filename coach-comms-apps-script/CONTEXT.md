# Coach Comms

Converts a coach's drafted email, written inline in the season's Communications Doc, into a Buttondown Draft, so a coach never has to copy/paste rich text out of Google Docs by hand.

## Language

**Communications Doc**:
The season's running Google Doc holding every email a coach drafts to parents, one after another over the season. There is no other name for it; coaches just call it the doc.

**Newsletter Block**:
A Google Docs table inserted via Insert → Building Blocks → Email Draft, containing To/Cc/Bcc/Subject/body fields, whose To field is `drafts@mg.buttondown.email`. Represents one authored email meant to become a Buttondown Draft. A Newsletter Block is a native Docs table, not a special embedded object; its To/Cc/Bcc cells hold a personProperties chip when a recipient is filled in.
_Avoid_: Email Block (too generic; the doc also has building blocks addressed elsewhere or left with empty recipients, which are not Newsletter Blocks)

**Buttondown Draft**:
The unsent draft created in Buttondown via its API (`POST /v1/emails`, `status: "draft"`) from one Newsletter Block's content. Distinct from Newsletter Subscribers (defined in [Coach Sheet](../coach-sheet-apps-script/CONTEXT.md)), which is the recipient list, not a piece of content.
_Avoid_: newsletter (ambiguous with Newsletter Subscribers), email (too generic)
