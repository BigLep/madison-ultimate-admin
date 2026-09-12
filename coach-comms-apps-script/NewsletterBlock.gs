/**
 * Reading a Newsletter Block: finding the To field's person chip, and converting
 * the body cell (paragraphs, bullet lists, bold/italic/links, inline images) to
 * Markdown for the Buttondown Draft body.
 */

// No angle brackets: Buttondown's editor parses "<...>" as an HTML tag even in
// Markdown mode, which swallows surrounding content into a bogus element.
const IMAGE_UPLOAD_PLACEHOLDER = 'COPY_PASTE_IN_IMAGE';
const NESTED_TABLE_PLACEHOLDER = 'COPY_PASTE_IN_TABLE';

// Markdown heading prefix for each DocumentApp.ParagraphHeading value that isn't
// NORMAL. TITLE/SUBTITLE map to h1/h2: Buttondown drafts don't have a separate
// title field, so a Doc's Title/Subtitle paragraph is just the biggest heading.
const HEADING_PREFIXES = {
  [DocumentApp.ParagraphHeading.TITLE]: '# ',
  [DocumentApp.ParagraphHeading.SUBTITLE]: '## ',
  [DocumentApp.ParagraphHeading.HEADING1]: '# ',
  [DocumentApp.ParagraphHeading.HEADING2]: '## ',
  [DocumentApp.ParagraphHeading.HEADING3]: '### ',
  [DocumentApp.ParagraphHeading.HEADING4]: '#### ',
  [DocumentApp.ParagraphHeading.HEADING5]: '##### ',
  [DocumentApp.ParagraphHeading.HEADING6]: '###### '
};

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
 *
 * Builds a list of block-level chunks (a heading/paragraph, or a run of
 * consecutive list items collapsed into one list) and joins them with a blank
 * line, since Markdown requires a blank line between paragraphs regardless of
 * whether the source Doc happened to have an empty paragraph there. Empty
 * paragraphs are dropped rather than preserved: blank-line spacing is now
 * structural, not copied from the Doc.
 */
function rowToMarkdown(row, apiKey) {
  const blocks = [];
  let listBuffer = [];
  const flushList = () => {
    if (listBuffer.length > 0) blocks.push(listBuffer.join('\n'));
    listBuffer = [];
  };

  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    for (let i = 0; i < cell.getNumChildren(); i++) {
      const child = cell.getChild(i);
      const type = child.getType();
      if (type === DocumentApp.ElementType.PARAGRAPH) {
        flushList();
        const paragraph = child.asParagraph();
        const text = containerInlineMarkdown(paragraph, apiKey);
        if (text.trim() === '') continue; // empty Doc paragraph: spacing only, not content
        const headingPrefix = HEADING_PREFIXES[paragraph.getHeading()];
        blocks.push(headingPrefix ? `${headingPrefix}${text}` : text);
      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        const item = child.asListItem();
        const indent = '  '.repeat(item.getNestingLevel());
        listBuffer.push(`${indent}- ${containerInlineMarkdown(item, apiKey)}`);
      } else if (type === DocumentApp.ElementType.TABLE) {
        // A nested table in the body isn't converted; flagged with a visible
        // placeholder rather than silently dropped.
        flushList();
        blocks.push(NESTED_TABLE_PLACEHOLDER);
      }
      // Anything else (horizontal rule, page break) is skipped silently.
    }
  }
  flushList();
  return blocks.join('\n\n').trim();
}

/**
 * Markdown for every child of a Paragraph or ListItem: text runs (bold/italic/link)
 * and inline images, concatenated in order. A Paragraph/ListItem is one logical
 * line, so any literal "\n" that shows up (Google's paragraph-terminator character,
 * occasionally carried inside a styled run) is stripped rather than treated as content.
 */
function containerInlineMarkdown(container, apiKey) {
  let out = '';
  for (let i = 0; i < container.getNumChildren(); i++) {
    out += elementInlineMarkdown(container.getChild(i), apiKey);
  }
  return out.replace(/\n/g, '');
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
