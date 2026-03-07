import { headingColors, getColor } from './colorConfig.js';

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
    // Uppercase Roman numerals: I. II. III. etc
    { regex: /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)\./, color: headingColors.section },
    // Lowercase Roman numerals: i. ii. iii. etc
    { regex: /^(i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii)\./, color: headingColors.section },
    // Arabic numerals: 1. 2. 3. etc
    { regex: /^\d+\./, color: headingColors.pointHeader },
    // Uppercase letters: A. B. C. etc
    { regex: /^[A-Z]\./, color: headingColors.subPoint },
    // Lowercase letters: a. b. c. etc
    { regex: /^[a-z]\./, color: headingColors.subPoint }
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
        break;
      }
    }
  }
  await context.sync();
}

// Master function — runs all detectors with optional toggles
export async function runAllDetection(context, body, options = {}) {
  const {
    useStyle   = true,
    useBold    = true,
    usePattern = true
  } = options;

  // style runs first — highest priority
  // bold and pattern will not overwrite an already highlighted paragraph
  if (useStyle)   await detectByStyle(context, body);
  if (useBold)    await detectByBold(context, body);
  if (usePattern) await detectByPattern(context, body);
}
export async function clearHighlights(context, body) {
  context.load(body, 'paragraphs');
  await context.sync();

  for (const para of body.paragraphs.items) {
    context.load(para, 'font');
  }
  await context.sync();

  for (const para of body.paragraphs.items) {
    const level = getLevel(para.font.highlightColor);
    if (level) {
      para.font.highlightColor = 'None';
    }
  }
  await context.sync();
}