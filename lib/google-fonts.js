'use strict';

// These are local/system fonts, not public Google Fonts. Including licensed
// families such as Georgia in a CSS API request can reject the whole stylesheet
// when a browser loads it from an ordinary website. Keep their document CSS;
// only omit them from the web-font request.
const SYSTEM_FONTS = new Set([
  'arial', 'arial black', 'calibri', 'cambria', 'cambria math',
  'comic sans ms', 'consolas', 'courier', 'courier new', 'georgia',
  'helvetica', 'helvetica neue', 'impact', 'lucida console',
  'lucida sans unicode', 'palatino', 'palatino linotype', 'segoe ui',
  'symbol', 'tahoma', 'times', 'times new roman', 'trebuchet ms',
  'verdana', 'webdings', 'wingdings',
  'serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui'
]);

function trackFont(usedFonts, style) {
  const family = style.weightedFontFamily?.fontFamily;
  if (!family) return;
  const weight = style.weightedFontFamily.weight || 400;
  usedFonts.add(`${family}:${weight}${style.italic ? ':italic' : ''}`);
}

function buildGoogleFontsLink(fonts) {
  const families = new Map();
  for (const font of fonts || []) {
    const [name, weight = '400', style] = font.split(':');
    const family = name.trim();
    if (!family || SYSTEM_FONTS.has(family.toLowerCase())) continue;
    if (!families.has(family)) {
      families.set(family, { weights: new Set([400, 700]), italic: false });
    }
    const entry = families.get(family);
    entry.weights.add(Number(weight));
    entry.italic ||= style === 'italic';
  }
  if (!families.size) return '';

  const query = [...families].map(([family, { weights, italic }]) => {
    const name = encodeURIComponent(family).replace(/%20/g, '+');
    const sorted = [...weights].sort((a, b) => a - b);
    // Keep upright faces when italic runs inherit their family. Request both
    // regular and bold italics; the browser only downloads faces it uses.
    const variants = italic
      ? `ital,wght@${[0, 1].flatMap(i => sorted.map(w => `${i},${w}`)).join(';')}`
      : `wght@${sorted.join(';')}`;
    return `family=${name}:${variants}`;
  }).join('&');
  return `https://fonts.googleapis.com/css2?${query}&display=block`;
}

module.exports = { trackFont, buildGoogleFontsLink };
