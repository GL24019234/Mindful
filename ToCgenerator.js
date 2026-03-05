import { getLevel } from './colorConfig.js';

export async function generateTOC(context, body) {

  // 1. Collect highlighted strings and their levels
  context.load(body, 'paragraphs');
  await context.sync();

  const entries = [];

  for (const para of body.paragraphs.items) {
    const runs = para.getTextRanges([' '], false);
    context.load(runs, 'items');
    await context.sync();

    for (const run of runs.items) {
      context.load(run, 'text, font');
    }
    await context.sync();

    // Check paragraph-level highlight first (whole paragraph highlighted)
    context.load(para, 'text, font');
    await context.sync();

    const paraLevel = getLevel(para.font.highlightColor);
    if (paraLevel && para.text.trim().length > 0) {
      entries.push({ text: para.text.trim(), level: paraLevel });
      continue;
    }

    // Fall back to run-level for partial highlights
    let chunk = '';
    let chunkLevel = null;
    for (const run of runs.items) {
      const level = getLevel(run.font.highlightColor);
      if (level) {
        chunk += run.text;
        chunkLevel = level;
      } else if (chunk.trim().length > 0) {
        entries.push({ text: chunk.trim(), level: chunkLevel });
        chunk = '';
        chunkLevel = null;
      }
    }
    if (chunk.trim().length > 0) {
      entries.push({ text: chunk.trim(), level: chunkLevel });
    }
  }

  // 2. Insert TOC at top of document
  const indentMap = { 1: 0, 2: 20, 3: 40 };
  const fontSizeMap = { 1: 14, 2: 12, 3: 11 };

  // Insert title first
  const title = body.insertParagraph('Table of Contents', Word.InsertLocation.start);
  title.styleBuiltIn = Word.BuiltInStyleName.heading1;

  // Insert entries in reverse order since each inserts at start
  for (const entry of [...entries].reverse()) {
    const para = body.insertParagraph(entry.text, Word.InsertLocation.start);
    para.style = 'Normal';
    para.leftIndent = indentMap[entry.level];
    para.font.size = fontSizeMap[entry.level];
    if (entry.level === 1) para.font.bold = true;
  }

  await context.sync();
}