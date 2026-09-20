const { buildGoogleFontsLink } = require('./google-fonts');
const { renderParagraph } = require('../gdocs-me-up');
const { StyleRegistry } = require('./styles');

function families(fonts) {
  return new URL(buildGoogleFontsLink(fonts)).searchParams.getAll('family');
}

test('loads upright faces alongside regular and bold italics in numeric tuple order', () => {
  expect(families(['PT Sans:500:italic', 'PT Sans:400', 'Anton:400'])).toEqual([
    'PT Sans:ital,wght@0,400;0,500;0,700;1,400;1,500;1,700',
    'Anton:wght@400;700'
  ]);
});

test('keeps local fonts out of otherwise valid Google Fonts requests', () => {
  expect(families(['Georgia:700:italic', 'PT Sans:400', 'Comic Sans MS:400', 'Arial:400'])).toEqual([
    'PT Sans:wght@400;700'
  ]);
  expect(buildGoogleFontsLink(['Consolas:400', 'arial:400'])).toBe('');
  expect(buildGoogleFontsLink([])).toBe('');
});

test('escapes font family names without adding URL parameters', () => {
  const url = new URL(buildGoogleFontsLink(['Rock & Roll:400']));
  expect(url.searchParams.getAll('family')).toEqual(['Rock & Roll:wght@400;700']);
  expect([...url.searchParams.keys()]).toEqual(['family', 'display']);
  expect(url.searchParams.get('display')).toBe('block');
});

test.each([false, true])('tracks italic runs with inherited fonts (paragraph italic: %s)', async italic => {
  const fonts = new Set();
  const { html } = await renderParagraph({
    paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
    elements: [
      { textRun: { content: 'plain ', textStyle: { italic: false } } },
      { textRun: { content: 'italic ', textStyle: { italic: true } } },
      { textRun: { content: 'bold italic\n', textStyle: { italic: true, bold: true } } }
    ]
  }, {}, fonts, [], null, '', '', {
    NORMAL_TEXT: {
      paragraphStyle: {},
      textStyle: { weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }, italic }
    }
  }, new StyleRegistry());
  expect(html).toContain('bold italic');
  expect(families(fonts)).toEqual(['PT Sans:ital,wght@0,400;0,700;1,400;1,700']);
});

test('retains a local family in document CSS without requesting it from Google', async () => {
  const fonts = new Set();
  const registry = new StyleRegistry();
  await renderParagraph({
    paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
    elements: [{ textRun: { content: 'Subtitle\n', textStyle: {} } }]
  }, {}, fonts, [], null, '', '', {
    NORMAL_TEXT: {
      paragraphStyle: {},
      textStyle: { weightedFontFamily: { fontFamily: 'Georgia', weight: 400 }, italic: true }
    }
  }, registry);
  expect(registry.toCSS()).toContain("font-family:'Georgia',sans-serif;");
  expect(buildGoogleFontsLink(fonts)).toBe('');
});
