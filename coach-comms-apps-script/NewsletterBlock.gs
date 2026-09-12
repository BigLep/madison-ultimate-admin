/**
 * Reading a Newsletter Block: finding the To field's person chip, and converting
 * the body cell (paragraphs, bullet lists, bold/italic/links, inline images) to
 * Markdown for the Buttondown Draft body.
 */

const IMAGE_UPLOAD_PLACEHOLDER = '<COPY PASTE IN IMAGE>';

/**
 * The email address of the first Person chip found in a table row, or null.
 * Used on the To row: a filled recipient is a Person chip; an empty slot is a
 * placeholder glyph (plain text), which has no Person to find.
 */
function findPersonEmailInRow(row) {
  for (let c = 0; c < row.getNumCells(); c++) {
    const email = findPersonEmailInElement(row.getCell(c));
    if (email) return email;
  }
  return null;
}

function findPersonEmailInElement(element) {
  if (element.getType() === DocumentApp.ElementType.PERSON) {
    return element.asPerson().getEmail();
  }
  const numChildren = typeof element.getNumChildren === 'function' ? element.getNumChildren() : 0;
  for (let i = 0; i < numChildren; i++) {
    const email = findPersonEmailInElement(element.getChild(i));
    if (email) return email;
  }
  return null;
}

/**
 * Markdown for a Newsletter Block's body row. The body cell spans both table
 * columns, but iterates every cell in the row defensively rather than assuming
 * which index holds the content.
 */
function rowToMarkdown(row, apiKey) {
  const lines = [];
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    for (let i = 0; i < cell.getNumChildren(); i++) {
      const child = cell.getChild(i);
      const type = child.getType();
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        lines.push(containerInlineMarkdown(child.asParagraph(), apiKey).trimEnd());
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        const item = child.asListItem();
        const indent = '  '.repeat(item.getNestingLevel());
        lines.push(`${indent}- ${containerInlineMarkdown(item, apiKey).trimEnd()}`);
      } else if (type === DocumentApp.ElementType.TABLE) {
        // Neither sampled Newsletter Block has a nested table in its body; flagged
        // rather than silently dropped if one ever does.
        lines.push('<UNSUPPORTED NESTED TABLE, COPY/PASTE THIS SECTION MANUALLY>');
      }
      // Anything else (horizontal rule, page break) is skipped silently.
    }
  }
  // Collapse runs of blank lines (from consecutive empty paragraphs) down to one.
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Markdown for every child of a Paragraph or ListItem: text runs (bold/italic/link)
 * and inline images, concatenated in order.
 */
function containerInlineMarkdown(container, apiKey) {
  let out = '';
  for (let i = 0; i < container.getNumChildren(); i++) {
    out += elementInlineMarkdown(container.getChild(i), apiKey);
  }
  return out;
}

function elementInlineMarkdown(element, apiKey) {
  const type = element.getType();
  if (type === DocumentApp.ElementType.TEXT) {
    return textRunsToMarkdown(element.asText());
  }
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    return inlineImageToMarkdown(element.asInlineImage(), apiKey);
  }
  return '';
}

/**
 * A Text element's runs, each wrapped for bold/italic/link. Does not escape
 * Markdown-special characters in plain prose (*, _, [, ]); fine for the sentence-
 * style content these blocks actually contain, but a known limitation.
 */
function textRunsToMarkdown(text) {
  const content = text.getText();
  if (!content) return '';
  const indices = text.getTextAttributeIndices();
  let out = '';
  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    const end = i + 1 < indices.length ? indices[i + 1] : content.length;
    const chunk = content.substring(start, end);
    if (!chunk) continue;
    let piece = chunk;
    if (text.isBold(start)) piece = `**${piece}**`;
    if (text.isItalic(start)) piece = `_${piece}_`;
    const link = text.getLinkUrl(start);
    if (link) piece = `[${piece}](${link})`;
    out += piece;
  }
  return out;
}

/**
 * An inline image, uploaded to Buttondown and referenced by its hosted URL.
 * Falls back to a placeholder (never throws) if the upload fails, since the exact
 * multipart field name Buttondown's /v1/images endpoint expects is unverified
 * against a real image as of this writing; see README.md.
 */
function inlineImageToMarkdown(image, apiKey) {
  try {
    const uploaded = uploadButtondownImage(apiKey, image.getBlob());
    return uploaded && uploaded.image ? `![](${uploaded.image})` : IMAGE_UPLOAD_PLACEHOLDER;
  } catch (e) {
    console.error(`Buttondown image upload failed, using placeholder: ${e}`);
    return IMAGE_UPLOAD_PLACEHOLDER;
  }
}
