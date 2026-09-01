const {
  generateGlobalCSS,
  renderParagraph,
  inlineImageVerticalMarginPx,
  buildNamedStylesMap
} = require('../gdocs-me-up');
const { StyleRegistry } = require('./styles');

describe('paragraph text-style inheritance', () => {
  test('neutralizes browser link decoration when Docs does not underline a link', () => {
    const css = generateGlobalCSS({}, 688);
    expect(css).toContain('a { color: inherit; text-decoration: inherit; }');
  });

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

  test('uses an inline font override natural line metrics', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{
          textRun: {
            content: 'فارسی\n',
            textStyle: { weightedFontFamily: { fontFamily: 'Noto Naskh Arabic', weight: 400 } }
          }
        }]
      },
      {}, new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: { lineSpacing: 115 },
          textStyle: {
            fontSize: { magnitude: 12, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
          }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain("font-family:'Noto Naskh Arabic',sans-serif;line-height:1.940625;");
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
    expect(css).toMatch(/sup\s*\{[^}]*vertical-align:\s*baseline;[^}]*line-height:\s*0;/s);
    expect(css).toMatch(/sub\s*\{[^}]*vertical-align:\s*baseline;[^}]*line-height:\s*0;/s);
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
    expect(titleRegistry.toCSS()).toContain('margin-bottom:8px;');
    expect(titleRegistry.toCSS()).not.toContain('padding-bottom:');

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

  test('preserves bottom leading for the compact Docs title preset', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [{ textRun: { content: 'Compact title\n', textStyle: {} } }]
      },
      {}, new Set(), [], null, '', '',
      {
        TITLE: {
          paragraphStyle: {
            spaceAbove: { magnitude: 10, unit: 'PT' },
            spaceBelow: { magnitude: 6, unit: 'PT' }
          },
          textStyle: { fontSize: { magnitude: 24, unit: 'PT' } }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('padding-bottom:4px;margin-bottom:8px;');
  });

  test('keeps explicitly expanded title leading below the glyph baseline', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE', lineSpacing: 150 },
        elements: [{
          textRun: {
            content: 'Expanded title\n',
            textStyle: { weightedFontFamily: { fontFamily: 'Lalezar', weight: 400 } }
          }
        }]
      },
      {}, new Set(), [], null, '', '',
      {
        TITLE: {
          paragraphStyle: {},
          textStyle: { fontSize: { magnitude: 24, unit: 'PT' } }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('line-height:2.1195600000000003;position:relative;top:-8px;');
  });

  test('does not deduct paragraph margins from an intentional blank line', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{ textRun: { content: '\n', textStyle: {} } }]
      },
      {}, new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: {
            lineSpacing: 115,
            spaceAbove: { magnitude: 12, unit: 'PT' },
            spaceBelow: { magnitude: 12, unit: 'PT' }
          },
          textStyle: {
            fontSize: { magnitude: 12, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'Noto Naskh Arabic', weight: 400 }
          }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('margin-top:16px;margin-bottom:16px;');
    expect(registry.toCSS()).toContain('min-height:32px;');
  });

  test('accounts for blank-line spacing in positioned-object anchor flow', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{ textRun: { content: '\n', textStyle: {} } }]
      },
      { positionedObjects: { anchored: {} } }, new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: {
            lineSpacing: 138,
            spaceAbove: { magnitude: 10, unit: 'PT' },
            spaceBelow: { magnitude: 6, unit: 'PT' }
          },
          textStyle: {
            fontSize: { magnitude: 12, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
          }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('min-height:13px;');
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

  test('cascades Normal text properties into headings', () => {
    const styles = buildNamedStylesMap({
      namedStyles: {
        styles: [
          {
            namedStyleType: 'NORMAL_TEXT',
            paragraphStyle: { lineSpacing: 115, spacingMode: 'NEVER_COLLAPSE' },
            textStyle: {
              fontSize: { magnitude: 12, unit: 'PT' },
              weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
            }
          },
          {
            namedStyleType: 'HEADING_1',
            paragraphStyle: { spaceAbove: { magnitude: 24, unit: 'PT' } },
            textStyle: { bold: true, fontSize: { magnitude: 18, unit: 'PT' } }
          },
          {
            namedStyleType: 'TITLE',
            paragraphStyle: { spaceBelow: { magnitude: 6, unit: 'PT' } },
            textStyle: { fontSize: { magnitude: 24, unit: 'PT' } }
          }
        ]
      }
    });

    expect(styles.HEADING_1.paragraphStyle).toMatchObject({
      lineSpacing: 115,
      spacingMode: 'NEVER_COLLAPSE',
      spaceAbove: { magnitude: 24, unit: 'PT' }
    });
    expect(styles.HEADING_1.textStyle).toMatchObject({
      bold: true,
      fontSize: { magnitude: 18, unit: 'PT' },
      weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
    });
    expect(styles.TITLE.paragraphStyle).toEqual({
      spaceBelow: { magnitude: 6, unit: 'PT' }
    });
    expect(styles.TITLE.paragraphStyle).not.toHaveProperty('lineSpacing');
    expect(styles.TITLE.textStyle).not.toHaveProperty('weightedFontFamily');
  });
});
