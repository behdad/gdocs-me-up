const {
  generateGlobalCSS,
  renderParagraph,
  inlineImageVerticalMarginPx
} = require('../gdocs-me-up');
const { StyleRegistry } = require('./styles');

describe('paragraph text-style inheritance', () => {
  test('does not leak a leading link style into the rest of the paragraph', async () => {
    const black = { color: { rgbColor: { red: 0, green: 0, blue: 0 } } };
    const blue = { color: { rgbColor: { red: 0.067, green: 0.333, blue: 0.8 } } };
    const namedStyles = {
      NORMAL_TEXT: {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT', lineSpacing: 115 },
        textStyle: {
          fontSize: { magnitude: 12, unit: 'PT' },
          weightedFontFamily: { fontFamily: 'Noto Naskh Arabic', weight: 400 },
          foregroundColor: black
        }
      }
    };
    const paragraph = {
      paragraphStyle: { namedStyleType: 'NORMAL_TEXT', direction: 'RIGHT_TO_LEFT' },
      elements: [
        {
          textRun: {
            content: 'نیما جفرودی',
            textStyle: {
              link: { url: 'https://example.com/' },
              underline: true,
              foregroundColor: blue
            }
          }
        },
        { textRun: { content: ' متن ساده\n', textStyle: {} } }
      ]
    };
    const registry = new StyleRegistry();

    const { html } = await renderParagraph(
      paragraph,
      {},
      new Set(),
      [],
      null,
      '',
      '',
      namedStyles,
      registry
    );

    expect(html).toContain('<a href="https://example.com/"');
    expect(html).toContain('<u>نیما جفرودی</u></a> متن ساده');
    expect(registry.toCSS()).toMatch(/\.p1\{[^}]*color:#000000;/);
    expect(registry.toCSS()).not.toMatch(/\.p1\{[^}]*(?:color:#1155cc|text-decoration-line:underline)/);
  });
});

describe('paragraph layout', () => {
  test('uses normal block flow so adjacent Google Docs margins collapse', () => {
    const css = generateGlobalCSS({}, 688);
    const contentRule = css.match(/\.doc-content\s*\{([^}]*)\}/)?.[1] || '';

    expect(contentRule).not.toMatch(/display\s*:\s*flex/);
    expect(css).toContain('.positioned-image { width: 100%; }');
    expect(css).toMatch(/\.doc-content li > p\s*\{[^}]*margin-block:\s*0;/s);
    expect(css).toMatch(/img\s*\{[^}]*vertical-align:\s*text-bottom;/s);
  });

  test('treats equal start and first-line indents as a whole-block indent', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: {
          namedStyleType: 'NORMAL_TEXT',
          direction: 'RIGHT_TO_LEFT',
          indentStart: { magnitude: 36, unit: 'PT' },
          indentFirstLine: { magnitude: 36, unit: 'PT' }
        },
        elements: [{ textRun: { content: 'quoted paragraph\n', textStyle: {} } }]
      },
      {},
      new Set(),
      [],
      null,
      '',
      '',
      { NORMAL_TEXT: { paragraphStyle: {}, textStyle: {} } },
      registry
    );

    expect(html).toContain('class="p1"');
    expect(registry.toCSS()).toContain('margin-inline-start:48px;');
    expect(registry.toCSS()).not.toContain('text-indent:');
  });

  test('preserves Docs title leading and a full blank line box', async () => {
    const namedStyles = {
      TITLE: {
        paragraphStyle: {
          namedStyleType: 'TITLE',
          spaceBelow: { magnitude: 6, unit: 'PT' }
        },
        textStyle: { fontSize: { magnitude: 24, unit: 'PT' } }
      },
      NORMAL_TEXT: {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT', lineSpacing: 115 },
        textStyle: {
          fontSize: { magnitude: 12, unit: 'PT' },
          weightedFontFamily: { fontFamily: 'Noto Naskh Arabic', weight: 400 }
        }
      }
    };
    const titleRegistry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [{ textRun: { content: 'Title\n', textStyle: {} } }]
      },
      {}, new Set(), [], null, '', '', namedStyles, titleRegistry
    );
    expect(titleRegistry.toCSS()).toMatch(/margin-bottom:8px;[^}]*padding-bottom:6px;/);

    const blankRegistry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{ textRun: { content: '\n', textStyle: {} } }]
      },
      {}, new Set(), [], null, '', '', namedStyles, blankRegistry
    );
    expect(blankRegistry.toCSS()).toContain('min-height:32px;');
  });

  test('scales inline-image wrap clearance to Docs inline spacing', () => {
    expect(inlineImageVerticalMarginPx(9)).toBe(4);
  });

  test('preserves an explicit LTR paragraph inside an RTL document', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: {
          namedStyleType: 'NORMAL_TEXT',
          direction: 'LEFT_TO_RIGHT',
          alignment: 'JUSTIFIED',
          indentStart: { magnitude: 36, unit: 'PT' },
          indentFirstLine: { magnitude: 36, unit: 'PT' }
        },
        elements: [{
          textRun: {
            content: '“You thought that it could never happen\n',
            textStyle: { italic: true }
          }
        }]
      },
      {}, new Set(), [], null, '', '',
      { NORMAL_TEXT: { paragraphStyle: {}, textStyle: {} } },
      registry
    );

    expect(html).toContain('<p dir="ltr"');
    expect(registry.toCSS()).toContain('margin-inline-start:48px;');
  });
});
