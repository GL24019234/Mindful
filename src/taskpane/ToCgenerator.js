import { getLevel } from './colorConfig.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

const TOC_BOOKMARK = '«TOC_START»';

async function clearExistingTOC(context, body) {
  try {
    const results = body.search(TOC_BOOKMARK, { matchCase: true });
    context.load(results, 'items');
    await context.sync();

    if (results.items.length === 0) return;

    // load all paragraph text before trying to delete
    const allParas = body.paragraphs;
    context.load(allParas, 'items');
    await context.sync();

    for (const para of allParas.items) {
      context.load(para, 'text');
    }
    await context.sync();

    let inTOC = false;
    for (const para of allParas.items) {
      if (para.text.includes(TOC_BOOKMARK)) {
        inTOC = !inTOC;
        para.delete();
        continue;
      }
      if (inTOC) para.delete();
    }

    await context.sync();

  } catch (err) {
    // no existing TOC found — safe to continue
    console.log('No existing TOC to clear:', err.message);
  }
}
// ── Main ──────────────────────────────────────────────────────────────────────

export async function generateTOC(context, body) {

  // 1. Clear any existing TOC first
  await clearExistingTOC(context, body);

  // 2. Collect highlighted entries
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

    context.load(para, 'text, font');
    await context.sync();

    // Paragraph-level check first
    const paraLevel = getLevel(para.font.highlightColor);
    if (paraLevel && para.text.trim().length > 0) {
      entries.push({ text: para.text.trim(), level: paraLevel });
      continue;
    }

    // Run-level fallback
    let currentChunk = '';
    let chunkLevel = null;
    for (const run of runs.items) {
      const level = getLevel(run.font.highlightColor);
      if (level) {
        currentChunk += run.text;
        chunkLevel = level;
      } else {
        if (currentChunk.trim().length > 0) {
          entries.push({ text: currentChunk.trim(), level: chunkLevel });
          currentChunk = '';
          chunkLevel = null;
        }
      }
    }
    if (currentChunk.trim().length > 0) {
      entries.push({ text: currentChunk.trim(), level: chunkLevel });
    }
  }

  // 3. Insert TOC — bottom to top since each insert goes to start
  const indentMap  = { 1: 0,  2: 20, 3: 40 };
  const sizeMap    = { 1: 14, 2: 12, 3: 11 };

  // End marker — invisible to user, used to find TOC boundary on next run
  const endMarker = body.insertParagraph(TOC_BOOKMARK, Word.InsertLocation.start);
  endMarker.font.hidden = true;

  // Insert entries in reverse so they appear in correct order
  for (const entry of [...entries].reverse()) {
    const para = body.insertParagraph(entry.text, Word.InsertLocation.start);
    para.style = 'Normal';
    para.leftIndent = indentMap[entry.level];
    para.font.size = sizeMap[entry.level];
    if (entry.level === 1) para.font.bold = true;
  }

  // Title
  const title = body.insertParagraph('Table of Contents', Word.InsertLocation.start);
  title.styleBuiltIn = Word.BuiltInStyleName.heading1;

  // Start marker — also hidden
  const startMarker = body.insertParagraph(TOC_BOOKMARK, Word.InsertLocation.start);
  startMarker.font.hidden = true;

  await context.sync();
}
/*

---

## How the clear logic works

The key problem with deleting a TOC is knowing **where it starts and ends** in the document. The solution here is wrapping the TOC in two hidden bookmark paragraphs containing a unique string `«TOC_START»` when it's first inserted:
```
«TOC_START»          ← hidden, start marker
Table of Contents    ← TOC title
Introduction         ← H1 entry
  Background         ← H2 entry
  Methodology        ← H2 entry
«TOC_START»          ← hidden, end marker
[rest of document]

*/