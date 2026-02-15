const {
  escapeHtml,
  ptToPx,
  rgbToHex,
  deepMerge,
  deepCopy,
  getOrDefault,
  formatBorder
} = require('./utils');

describe('escapeHtml', () => {
  test('escapes HTML special characters', () => {
    expect(escapeHtml('<script>alert("XSS")</script>')).toBe('&lt;script&gt;alert(&quot;XSS&quot;)&lt;/script&gt;');
    expect(escapeHtml('Tom & Jerry')).toBe('Tom &amp; Jerry');
    expect(escapeHtml('"Hello"')).toBe('&quot;Hello&quot;');
  });

  test('handles null and undefined', () => {
    expect(escapeHtml(null)).toBe('');
    expect(escapeHtml(undefined)).toBe('');
  });

  test('handles empty string', () => {
    expect(escapeHtml('')).toBe('');
  });

  test('handles strings without special characters', () => {
    expect(escapeHtml('Hello World')).toBe('Hello World');
  });
});

describe('ptToPx', () => {
  test('converts points to pixels', () => {
    expect(ptToPx(12)).toBe(16); // 12 * 1.3333 ≈ 16
    expect(ptToPx(0)).toBe(0);
    expect(ptToPx(72)).toBe(96); // 72 * 1.3333 ≈ 96
  });

  test('handles invalid inputs', () => {
    expect(ptToPx(NaN)).toBe(0);
    expect(ptToPx('invalid')).toBe(0);
    expect(ptToPx(null)).toBe(0);
  });

  test('rounds to nearest integer', () => {
    expect(ptToPx(10)).toBe(13); // 10 * 1.3333 = 13.333 → 13
  });
});

describe('rgbToHex', () => {
  test('converts RGB to hex', () => {
    expect(rgbToHex(1, 0, 0)).toBe('#ff0000'); // Red
    expect(rgbToHex(0, 1, 0)).toBe('#00ff00'); // Green
    expect(rgbToHex(0, 0, 1)).toBe('#0000ff'); // Blue
    expect(rgbToHex(1, 1, 1)).toBe('#ffffff'); // White
    expect(rgbToHex(0, 0, 0)).toBe('#000000'); // Black
  });

  test('handles values between 0 and 1', () => {
    expect(rgbToHex(0.5, 0.5, 0.5)).toBe('#808080'); // Gray
  });

  test('clamps values outside 0-1 range', () => {
    expect(rgbToHex(2, 0, 0)).toBe('#ff0000'); // Clamped to 1
    expect(rgbToHex(-1, 0, 0)).toBe('#000000'); // Clamped to 0
  });

  test('handles null/undefined as 0', () => {
    expect(rgbToHex(null, null, null)).toBe('#000000');
  });
});

describe('deepMerge', () => {
  test('merges objects deeply', () => {
    const base = { a: 1, b: { c: 2 } };
    const overlay = { b: { d: 3 }, e: 4 };
    deepMerge(base, overlay);
    expect(base).toEqual({ a: 1, b: { c: 2, d: 3 }, e: 4 });
  });

  test('overwrites primitive values', () => {
    const base = { a: 1 };
    const overlay = { a: 2 };
    deepMerge(base, overlay);
    expect(base).toEqual({ a: 2 });
  });

  test('handles nested objects', () => {
    const base = { a: { b: { c: 1 } } };
    const overlay = { a: { b: { d: 2 } } };
    deepMerge(base, overlay);
    expect(base).toEqual({ a: { b: { c: 1, d: 2 } } });
  });

  test('handles arrays as primitives (overwrites)', () => {
    const base = { a: [1, 2] };
    const overlay = { a: [3, 4] };
    deepMerge(base, overlay);
    expect(base).toEqual({ a: [3, 4] });
  });
});

describe('deepCopy', () => {
  test('creates deep copy of object', () => {
    const obj = { a: 1, b: { c: 2 } };
    const copy = deepCopy(obj);
    expect(copy).toEqual(obj);
    expect(copy).not.toBe(obj); // Different reference
    expect(copy.b).not.toBe(obj.b); // Nested objects also copied
  });

  test('handles arrays', () => {
    const arr = [1, 2, { a: 3 }];
    const copy = deepCopy(arr);
    expect(copy).toEqual(arr);
    expect(copy).not.toBe(arr);
  });

  test('copies nested structures', () => {
    const obj = { a: { b: { c: [1, 2, 3] } } };
    const copy = deepCopy(obj);
    copy.a.b.c.push(4);
    expect(obj.a.b.c).toEqual([1, 2, 3]); // Original unchanged
    expect(copy.a.b.c).toEqual([1, 2, 3, 4]);
  });
});

describe('getOrDefault', () => {
  test('returns value if defined', () => {
    expect(getOrDefault(5, 10)).toBe(5);
    expect(getOrDefault('hello', 'default')).toBe('hello');
    expect(getOrDefault(0, 10)).toBe(0); // 0 is a valid value
    expect(getOrDefault(false, true)).toBe(false);
  });

  test('returns default if null or undefined', () => {
    expect(getOrDefault(null, 10)).toBe(10);
    expect(getOrDefault(undefined, 10)).toBe(10);
  });
});

describe('formatBorder', () => {
  const borderStyleMap = {
    'SOLID': 'solid',
    'DASHED': 'dashed',
    'DOTTED': 'dotted'
  };

  test('formats border with all properties', () => {
    const border = {
      width: { magnitude: 1 },
      dashStyle: 'SOLID',
      color: { color: { rgbColor: { red: 1, green: 0, blue: 0 } } }
    };
    const result = formatBorder('top', border, borderStyleMap);
    expect(result).toMatch(/^border-top:/);
    expect(result).toContain('solid');
    expect(result).toContain('#ff0000');
  });

  test('returns empty string if no border', () => {
    expect(formatBorder('top', null, borderStyleMap)).toBe('');
    expect(formatBorder('top', {}, borderStyleMap)).toBe('');
  });

  test('returns empty string if no width', () => {
    const border = { dashStyle: 'SOLID' };
    expect(formatBorder('top', border, borderStyleMap)).toBe('');
  });

  test('uses default color if not specified', () => {
    const border = {
      width: { magnitude: 1 },
      dashStyle: 'SOLID'
    };
    const result = formatBorder('top', border, borderStyleMap);
    expect(result).toContain('#000000');
  });

  test('handles different border sides', () => {
    const border = {
      width: { magnitude: 1 },
      dashStyle: 'SOLID'
    };
    expect(formatBorder('top', border, borderStyleMap)).toContain('border-top:');
    expect(formatBorder('bottom', border, borderStyleMap)).toContain('border-bottom:');
    expect(formatBorder('left', border, borderStyleMap)).toContain('border-left:');
    expect(formatBorder('right', border, borderStyleMap)).toContain('border-right:');
  });
});
