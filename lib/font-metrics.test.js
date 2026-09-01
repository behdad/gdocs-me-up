'use strict';

const {
  collectFontFamilies,
  naturalLineHeight,
  lineHeightFor
} = require('./font-metrics');

test('collects font families from document styles and runs', () => {
  const families = collectFontFamilies({
    namedStyles: { styles: [{ textStyle: { weightedFontFamily: { fontFamily: 'Anton' } } }] },
    body: { content: [{ paragraph: { elements: [{ textRun: {
      textStyle: { weightedFontFamily: { fontFamily: 'Lalezar' } }
    } }] } }] }
  });
  expect([...families].sort()).toEqual(['Anton', 'Lalezar']);
});

test('derives line height from OpenType vertical metrics', () => {
  const metric = { unitsPerEm: 1000, ascent: 979, descent: -588, lineGap: 0 };
  expect(naturalLineHeight(metric)).toBe(1.567);
  expect(lineHeightFor('Example', 115, new Map([['Example', metric]]))).toBe(1.80205);
});

test('uses the Docs spacing multiplier when metrics are unavailable', () => {
  expect(lineHeightFor('Unavailable', 115, new Map())).toBe(1.15);
});
