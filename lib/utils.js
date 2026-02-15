/**
 * Utility functions for Google Docs export
 */

/**
 * Escape HTML special characters.
 *
 * @param {string} str - String to escape
 * @returns {string} Escaped string
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  const strValue = String(str);
  return strValue
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Convert points to pixels (1pt ≈ 1.3333px).
 *
 * @param {number} pts - Points value
 * @returns {number} Pixels value
 */
function ptToPx(pts) {
  if (typeof pts !== 'number' || isNaN(pts)) return 0;
  return Math.round(pts * 1.3333);
}

/**
 * Convert RGB values (0-1 range) to hex color.
 *
 * @param {number} r - Red (0-1)
 * @param {number} g - Green (0-1)
 * @param {number} b - Blue (0-1)
 * @returns {string} Hex color string
 */
function rgbToHex(r, g, b) {
  const clamp = (val) => Math.max(0, Math.min(1, val || 0));
  const nr = Math.round(clamp(r) * 255);
  const ng = Math.round(clamp(g) * 255);
  const nb = Math.round(clamp(b) * 255);
  return '#' + [nr, ng, nb].map(x => x.toString(16).padStart(2, '0')).join('');
}

/**
 * Deep merge two objects.
 *
 * @param {object} base - Base object to merge into
 * @param {object} overlay - Object to merge from
 */
function deepMerge(base, overlay) {
  for (const k in overlay) {
    if (
      typeof overlay[k] === 'object' &&
      overlay[k] !== null &&
      !Array.isArray(overlay[k])
    ) {
      if (!base[k]) base[k] = {};
      deepMerge(base[k], overlay[k]);
    } else {
      base[k] = overlay[k];
    }
  }
}

/**
 * Deep copy an object via JSON serialization.
 *
 * @param {object} obj - Object to copy
 * @returns {object} Deep copy of the object
 */
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * Safely get a value with a default fallback.
 *
 * @param {*} value - The value to check
 * @param {*} defaultValue - The default value to return if value is null/undefined
 * @returns {*} The value or default
 */
function getOrDefault(value, defaultValue) {
  return (value !== undefined && value !== null) ? value : defaultValue;
}

/**
 * Format a border style for CSS.
 *
 * @param {string} side - Border side (top, bottom, left, right)
 * @param {object} border - Border object from Google Docs
 * @returns {string} CSS border string
 */
function formatBorder(side, border, borderStyleMap) {
  if (!border || !border.width || !border.width.magnitude) return '';
  const width = ptToPx(border.width.magnitude);
  const style = borderStyleMap[border.dashStyle] || 'solid';
  let color = '#000000';
  if (border.color?.color?.rgbColor) {
    const rgb = border.color.color.rgbColor;
    color = rgbToHex(rgb.red || 0, rgb.green || 0, rgb.blue || 0);
  }
  return `border-${side}:${width}px ${style} ${color};`;
}

module.exports = {
  escapeHtml,
  ptToPx,
  rgbToHex,
  deepMerge,
  deepCopy,
  getOrDefault,
  formatBorder
};
