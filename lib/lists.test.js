const {
  inferNestingLevel,
  detectBulletStyle,
  detectNumberStyle,
  isNumberedList,
  closeAllLists
} = require('./lists');

describe('inferNestingLevel', () => {
  test('returns explicit nesting level if provided', () => {
    const bullet = { nestingLevel: 2 };
    expect(inferNestingLevel(bullet, {}, 0)).toBe(2);
  });

  test('infers level 0 from small indentation', () => {
    const bullet = {};
    const paragraphStyle = { indentStart: { magnitude: 0 } };
    expect(inferNestingLevel(bullet, paragraphStyle, -1)).toBe(0);
  });

  test('infers level 1 from large indentation', () => {
    const bullet = {};
    const paragraphStyle = { indentStart: { magnitude: 60 } };
    expect(inferNestingLevel(bullet, paragraphStyle, -1)).toBe(1);
  });

  test('continues at previous level if indentation is ambiguous', () => {
    const bullet = {};
    const paragraphStyle = { indentStart: { magnitude: 50 } }; // Between 40 and 60
    expect(inferNestingLevel(bullet, paragraphStyle, 1)).toBe(1);
  });

  test('defaults to level 0 if no previous level', () => {
    const bullet = {};
    const paragraphStyle = { indentStart: { magnitude: 25 } };
    expect(inferNestingLevel(bullet, paragraphStyle, -1)).toBe(0);
  });
});

describe('detectBulletStyle', () => {
  test('detects disc style (default)', () => {
    const glyph = { glyphSymbol: '●' };
    expect(detectBulletStyle(glyph)).toBe('disc');
  });

  test('detects circle style', () => {
    const glyph = { glyphSymbol: '○' };
    expect(detectBulletStyle(glyph)).toBe('circle');
  });

  test('detects square style', () => {
    const glyph = { glyphSymbol: '■' };
    expect(detectBulletStyle(glyph)).toBe('square');
  });

  test('detects dash style', () => {
    const glyph = { glyphSymbol: '-' };
    expect(detectBulletStyle(glyph)).toBe('dash');
  });

  test('returns disc as default for unknown symbols', () => {
    const glyph = { glyphSymbol: '★' };
    expect(detectBulletStyle(glyph)).toBe('disc');
  });

  test('returns disc if no glyph symbol', () => {
    expect(detectBulletStyle({})).toBe('disc');
    expect(detectBulletStyle(null)).toBe('disc');
  });
});

describe('detectNumberStyle', () => {
  test('detects decimal style', () => {
    const glyph = { glyphType: 'DECIMAL' };
    expect(detectNumberStyle(glyph)).toBe('decimal');
  });

  test('detects upper-alpha style', () => {
    const glyph = { glyphType: 'UPPER_ALPHA' };
    expect(detectNumberStyle(glyph)).toBe('upper-alpha');
  });

  test('detects lower-alpha style', () => {
    const glyph = { glyphType: 'LOWER_ALPHA' };
    expect(detectNumberStyle(glyph)).toBe('lower-alpha');
  });

  test('detects upper-roman style', () => {
    const glyph = { glyphType: 'UPPER_ROMAN' };
    expect(detectNumberStyle(glyph)).toBe('upper-roman');
  });

  test('detects lower-roman style', () => {
    const glyph = { glyphType: 'LOWER_ROMAN' };
    expect(detectNumberStyle(glyph)).toBe('lower-roman');
  });

  test('returns decimal as default', () => {
    expect(detectNumberStyle({})).toBe('decimal');
    expect(detectNumberStyle(null)).toBe('decimal');
  });
});

describe('isNumberedList', () => {
  test('returns false if glyphSymbol is present (bullet list)', () => {
    const glyph = { glyphSymbol: '●' };
    expect(isNumberedList(glyph, 'list1', 0, {})).toBe(false);
  });

  test('returns true if glyphType is numbered', () => {
    const glyph = { glyphType: 'DECIMAL' };
    expect(isNumberedList(glyph, 'list1', 0, {})).toBe(true);

    const glyph2 = { glyphType: 'UPPER_ALPHA' };
    expect(isNumberedList(glyph2, 'list1', 0, {})).toBe(true);
  });

  test('uses item count heuristic for unspecified glyph type', () => {
    const glyph = { glyphType: 'GLYPH_TYPE_UNSPECIFIED' };

    // Single-item list → numbered
    const counts1 = { 'list1:0': 1 };
    expect(isNumberedList(glyph, 'list1', 0, counts1)).toBe(true);

    // Multi-item list → bullet
    const counts2 = { 'list1:0': 3 };
    expect(isNumberedList(glyph, 'list1', 0, counts2)).toBe(false);
  });

  test('defaults to bullet list', () => {
    const glyph = { glyphType: 'UNKNOWN_TYPE' };
    expect(isNumberedList(glyph, 'list1', 0, {})).toBe(false);
  });
});

describe('closeAllLists', () => {
  test('closes all lists in stack', () => {
    const listStack = ['ul:0', 'ol:1', 'ul:2'];
    const htmlLines = [];
    closeAllLists(listStack, htmlLines);
    expect(htmlLines).toEqual(['</ul>', '</ol>', '</ul>']);
    expect(listStack).toHaveLength(0);
  });

  test('closes unordered lists with </ul>', () => {
    const listStack = ['ul:0', 'ul_circle:1'];
    const htmlLines = [];
    closeAllLists(listStack, htmlLines);
    expect(htmlLines).toEqual(['</ul>', '</ul>']);
  });

  test('closes ordered lists with </ol>', () => {
    const listStack = ['ol:0', 'ol_upper_alpha:1'];
    const htmlLines = [];
    closeAllLists(listStack, htmlLines);
    expect(htmlLines).toEqual(['</ol>', '</ol>']);
  });

  test('handles empty stack', () => {
    const listStack = [];
    const htmlLines = [];
    closeAllLists(listStack, htmlLines);
    expect(htmlLines).toEqual([]);
  });

  test('handles null/undefined inputs gracefully', () => {
    expect(() => closeAllLists(null, [])).not.toThrow();
    expect(() => closeAllLists([], null)).not.toThrow();
  });
});
