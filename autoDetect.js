import { headingColors } from './colorConfig.js';

// 1. Word styles — Heading 1/2/3
export async function detectByStyle(context, body) {
  const styleMap = {
    'Heading 1': headingColors.section,
    'Heading 2': headingColors.pointHeader,
    'Heading 3': headingColors.subPoint
  };

  context.load(body, 'paragraphs');
  await context.sync();

  for (const para of body.paragraphs.items) {
    context.load(para, 'text, style');
  }
  await context.sync();

  for (const para of body.paragraphs.items) {
    const color = styleMap[para.style];
    if (color) {
      para.font.highlightColor = color;
    }
  }
  await context.sync();
}

// 2. Bold paragraphs — treated as Point Headers
export async function detectByBold(context, body) {
  for (const para of body.paragraphs.items) {
    context.load(para, 'text, font');
  }
  await context.sync();

  for (const para of body.paragraphs.items) {
    if (para.font.bold === true && para.text.trim().length > 0) {
      para.font.highlightColor = headingColors.pointHeader;
    }
  }
  await context.sync();
}

// 3. Pattern detection — Roman numerals, Arabic numerals, letters followed by period
export async function detectByPattern(context, body) {
  const patterns = [
    // Roman numerals:  I. II. III. IV. etc
    { regex: /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)\./i, color: headingColors.section },
    // Arabic numerals: 1. 2. 3. etc
    { regex: /^\d+\./, color: headingColors.pointHeader },
    // Letters:         A. B. C. etc
    { regex: /^[A-Z]\./, color: headingColors.subPoint }
  ];

  for (const para of body.paragraphs.items) {
    context.load(para, 'text');
  }
  await context.sync();

  for (const para of body.paragraphs.items) {
    const text = para.text.trim();
    for (const pattern of patterns) {
      if (pattern.regex.test(text)) {
        para.font.highlightColor = pattern.color;
        break; // stop at first match
      }
    }
  }
  await context.sync();
}