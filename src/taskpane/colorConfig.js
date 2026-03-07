// Default color → heading level mapping
export const headingColors = {
  section:        'Yellow',
  pointHeader:    'Cyan',
  subPoint:       'Pink'
};

// Update a color assignment — called from the UI
export function setColor(level, color) {
  if (!(level in headingColors)) {
    console.error(`Unknown level: ${level}`);
    return;
  }
  headingColors[level] = color;
}

// Given a highlight color, return its heading level (1, 2, 3 or null)
export function getLevel(color) {
  if (color === headingColors.section)        return 1;
  if (color === headingColors.pointHeader)    return 2;
  if (color === headingColors.subPoint)       return 3;
  return null;
}

// Given a heading level, return its color
export function getColor(level) {
  const map = { 1: headingColors.section, 2: headingColors.pointHeader, 3: headingColors.subPoint };
  return map[level] ?? null;
}

// Human-readable label for each level
export function getLabel(level) {
  const map = { 1: 'Section', 2: 'Point Header', 3: 'Sub-point Header' };
  return map[level] ?? 'Unknown';
}