/**
 * Constants used throughout the Google Docs exporter
 */

// Indentation thresholds for inferring nesting levels (in points)
const INDENT_LEVEL_0_MAX = 40;  // Items with indent <= 40pt are level 0
const INDENT_LEVEL_1_MIN = 60;  // Items with indent >= 60pt are level 1

// List glyph types
const NUMBERED_GLYPH_TYPES = [
  'DECIMAL', 'ALPHA', 'ROMAN',
  'UPPER_ALPHA', 'UPPER_ROMAN',
  'LOWER_ALPHA', 'LOWER_ROMAN'
];

// Basic alignment map for LTR paragraphs
const ALIGNMENT_MAP_LTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// Border style map
const BORDER_STYLE_MAP = {
  SOLID: 'solid',
  DOTTED: 'dotted',
  DASHED: 'dashed',
  DOUBLE: 'double'
};

// Bullet style mappings
const BULLET_STYLE_MAP = {
  '●': 'disc',
  '○': 'circle',
  '■': 'square',
  '-': 'dash'
};

// Number style mappings
const NUMBER_STYLE_MAP = {
  'DECIMAL': 'decimal',
  'UPPER_ALPHA': 'upper-alpha',
  'LOWER_ALPHA': 'lower-alpha',
  'ALPHA': 'lower-alpha',
  'UPPER_ROMAN': 'upper-roman',
  'ROMAN': 'upper-roman',
  'LOWER_ROMAN': 'lower-roman'
};

module.exports = {
  INDENT_LEVEL_0_MAX,
  INDENT_LEVEL_1_MIN,
  NUMBERED_GLYPH_TYPES,
  ALIGNMENT_MAP_LTR,
  BORDER_STYLE_MAP,
  BULLET_STYLE_MAP,
  NUMBER_STYLE_MAP
};
