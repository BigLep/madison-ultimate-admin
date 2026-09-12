#!/usr/bin/env node
/**
 * Compares NewsletterBlock.gs's hand-rolled Docs-to-Markdown walker against real
 * Turndown-on-HTML-export, over the actual Newsletter Blocks in a real
 * Communications Doc. See README.md in this folder for how to fetch the two
 * input files this needs, and why they're gitignored rather than committed.
 *
 * The "mirror" conversion below is a best-effort, by-hand port of
 * NewsletterBlock.gs's logic to run outside Apps Script (against the raw Docs API
 * JSON model instead of DocumentApp). Keep it in sync manually when that file
 * changes; it does not run the real script and can drift.
 */
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');
const TurndownService = require('turndown');
const { gfm } = require('turndown-plugin-gfm');

const DOC_RAW_PATH = path.join(__dirname, 'doc_raw.json');
const DOC_HTML_PATH = path.join(__dirname, 'doc.html');
const TO_ADDRESS = process.env.NEWSLETTER_BLOCK_TO || 'drafts@mg.buttondown.email';

if (!fs.existsSync(DOC_RAW_PATH) || !fs.existsSync(DOC_HTML_PATH)) {
  console.error(`Missing doc_raw.json and/or doc.html in ${__dirname}.\nSee README.md in this folder for how to fetch them first.`);
  process.exit(1);
}

// ---- Mirror of NewsletterBlock.gs (raw-JSON version) ----------------------

const HEADING_PREFIXES = {
  TITLE: '# ', SUBTITLE: '## ',
  HEADING_1: '# ', HEADING_2: '## ', HEADING_3: '### ',
  HEADING_4: '#### ', HEADING_5: '##### ', HEADING_6: '###### '
};

function rgbColorToHex(rgbColor) {
  const channel = (v) => Math.round((v || 0) * 255).toString(16).padStart(2, '0');
  return `#${channel(rgbColor.red)}${channel(rgbColor.green)}${channel(rgbColor.blue)}`;
}

function textRunsToMarkdown(elements) {
  let out = '';
  for (const el of elements) {
    if (!el.textRun) continue;
    let piece = el.textRun.content;
    const ts = el.textRun.textStyle || {};
    if (ts.bold) piece = `**${piece}**`;
    if (ts.italic) piece = `_${piece}_`;
    if (ts.link && ts.link.url) piece = `[${piece}](${ts.link.url})`;
    if (ts.backgroundColor && ts.backgroundColor.color && ts.backgroundColor.color.rgbColor) {
      const hex = rgbColorToHex(ts.backgroundColor.color.rgbColor);
      piece = `<mark style="background-color:${hex}">${piece}</mark>`;
    }
    out += piece;
  }
  return out.replace(/\n/g, '');
}

function rowToMarkdownMirror(row) {
  const blocks = [];
  let paragraphLines = [];
  let listBuffer = [];
  const flushParagraph = () => { if (paragraphLines.length) blocks.push(paragraphLines.join('  \n')); paragraphLines = []; };
  const flushList = () => { if (listBuffer.length) blocks.push(listBuffer.join('\n')); listBuffer = []; };

  for (const cell of row.tableCells) {
    for (const item of cell.content) {
      const p = item.paragraph;
      if (!p) continue;
      const text = textRunsToMarkdown(p.elements);
      if (p.bullet) {
        flushParagraph();
        listBuffer.push(`- ${text}`);
        continue;
      }
      if (text.trim() === '') { flushParagraph(); flushList(); continue; }
      flushList();
      const style = (p.paragraphStyle && p.paragraphStyle.namedStyleType) || 'NORMAL_TEXT';
      const prefix = HEADING_PREFIXES[style];
      if (prefix) {
        flushParagraph();
        blocks.push(`${prefix}${text}`);
      } else {
        paragraphLines.push(text);
      }
    }
  }
  flushParagraph();
  flushList();
  return blocks.join('\n\n').trim();
}

function findEmail(cell) {
  for (const item of cell.content) {
    const p = item.paragraph;
    if (!p) continue;
    for (const el of p.elements) {
      if (el.person) return el.person.personProperties.email;
    }
  }
  return null;
}

// ---- Turndown-on-HTML-export ----------------------------------------------

const turndownService = new TurndownService({ headingStyle: 'atx', bulletListMarker: '-' });
turndownService.use(gfm);

function htmlBodyToMarkdown($, rows) {
  const bodyHtml = $(rows[4]).find('td').first().html();
  return turndownService.turndown(bodyHtml || '');
}

// ---- Run --------------------------------------------------------------

const doc = JSON.parse(fs.readFileSync(DOC_RAW_PATH, 'utf8'));
const tables = doc.body.content.filter(e => e.table).map(e => e.table);
const html = fs.readFileSync(DOC_HTML_PATH, 'utf8');
const $ = cheerio.load(html);
const htmlTables = $('table').toArray().filter(t => $(t).find('> tbody > tr, > tr').length === 5);

let matched = 0;
tables.forEach((table, i) => {
  const toEmail = findEmail(table.tableRows[0].tableCells[1]);
  if (toEmail !== TO_ADDRESS) return;
  matched++;

  const subjectCell = table.tableRows[3].tableCells[1];
  const subject = subjectCell.content
    .flatMap(item => (item.paragraph ? item.paragraph.elements : []))
    .map(el => (el.textRun ? el.textRun.content : ''))
    .join('').trim();

  console.log('='.repeat(80));
  console.log(`Newsletter Block #${i}: "${subject}"`);
  console.log('='.repeat(80));

  console.log('\n--- mirror (NewsletterBlock.gs logic, raw Docs JSON) ---\n');
  console.log(rowToMarkdownMirror(table.tableRows[4]));

  const htmlTable = htmlTables.find(t => $(t).find('tr').eq(3).find('td').eq(1).text().trim() === subject);
  if (htmlTable) {
    console.log('\n--- turndown (real library, HTML export) ---\n');
    console.log(htmlBodyToMarkdown($, $(htmlTable).find('tr').toArray()));
  } else {
    console.log('\n(could not find matching table in doc.html by subject text for Turndown comparison)');
  }
  console.log('');
});

if (matched === 0) {
  console.log(`No Newsletter Blocks found addressed to "${TO_ADDRESS}". Set NEWSLETTER_BLOCK_TO to compare a different address.`);
}
