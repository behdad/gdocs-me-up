const {
  generateGlobalCSS,
  renderParagraph,
  inlineImageVerticalMarginPx,
  buildNamedStylesMap,
  computeDocContainerWidth
} = require('../gdocs-me-up');
const { StyleRegistry } = require('./styles');

const FONT_METRICS = new Map([
  ['Lalezar', { unitsPerEm: 1000, ascent: 979, descent: -588, lineGap: 0 }],
  ['Noto Naskh Arabic', { unitsPerEm: 1000, ascent: 1069, descent: -634, lineGap: 0 }],
  ['PT Sans', { unitsPerEm: 1000, ascent: 1018, descent: -276, lineGap: 0 }]
]);

function metricDoc(properties = {}) {
  return { ...properties, ___fontMetrics: FONT_METRICS };
}

describe('paragraph text-style inheritance', () => {
  test('leaves ordinary link presentation to the browser', () => {
    const css = generateGlobalCSS({}, 688);
    expect(css).not.toMatch(/a \{/);
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
    expect(html).toContain('<a href="https://example.com/">نیما جفرودی</a> متن ساده');
    expect(registry.toCSS()).toMatch(/\.p1\{[^}]*color:#000000;/);
    expect(registry.toCSS()).not.toContain('#1155cc');
    expect(registry.toCSS()).not.toMatch(/\.p1\{[^}]*(?:color:#1155cc|text-decoration-line:underline)/);
  });

  test('emits no-underline styling only when the link says so explicitly', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{
          textRun: {
            content: 'plain link\n',
            textStyle: {
              link: { url: 'https://example.com/' },
              underline: false
            }
          }
        }]
      },
      {}, new Set(), [], null, '', '',
      { NORMAL_TEXT: { paragraphStyle: {}, textStyle: {} } },
      registry
    );

    expect(html).toMatch(/<a href="https:\/\/example\.com\/" class="t\d+">plain link<\/a>/);
    expect(registry.toCSS()).toMatch(/\.t\d+\{text-decoration-line:none;}/);
  });

  test('preserves a custom link color even when it matches inherited text', async () => {
    const black = { color: { rgbColor: { red: 0, green: 0, blue: 0 } } };
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{
          textRun: {
            content: 'custom link\n',
            textStyle: {
              link: { url: 'https://example.com/' },
              foregroundColor: black
            }
          }
        }]
      },
      {}, new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: {},
          textStyle: { foregroundColor: black }
        }
      },
      registry
    );

    expect(html).toMatch(/<a href="https:\/\/example\.com\/" class="t\d+">custom link<\/a>/);
    expect(registry.toCSS()).toMatch(/\.t\d+\{color:#000000;}/);
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
      metricDoc(), new Set(), [], null, '', '',
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

    expect(registry.toCSS()).toContain("line-height:1.95845;font-size:12pt;font-family:'Noto Naskh Arabic',sans-serif;");
    expect(registry.toCSS()).not.toContain('.t1{');
  });
});

describe('paragraph layout', () => {
  test('adds pageless width only when the document is not in page mode', () => {
    const documentStyle = {
      pageSize: { width: { magnitude: 612 } },
      marginLeft: { magnitude: 72 },
      marginRight: { magnitude: 72 }
    };
    expect(computeDocContainerWidth({
      documentStyle: {
        ...documentStyle,
        documentFormat: { documentMode: 'PAGES' }
      }
    })).toBe(624);
    expect(computeDocContainerWidth({
      documentStyle: {
        ...documentStyle,
        documentFormat: { documentMode: 'PAGELESS' }
      }
    })).toBe(690);
    expect(computeDocContainerWidth({ documentStyle })).toBe(690);
  });

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
      metricDoc(), new Set(), [], null, '', '', namedStyles, blankRegistry
    );
    expect(blankRegistry.toCSS()).toContain('min-height:32px;');
  });

  test('does not insert a line-height-inflating glyph between styled soft-break runs', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [
          {
            textRun: {
              content: 'Title\u000b',
              textStyle: {
                fontSize: { magnitude: 24, unit: 'PT' },
                weightedFontFamily: { fontFamily: 'Lalezar', weight: 400 }
              }
            }
          },
          {
            textRun: {
              content: 'Author\n',
              textStyle: {
                fontSize: { magnitude: 12, unit: 'PT' },
                weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
              }
            }
          }
        ]
      },
      metricDoc(), new Set(), [], null, '', '',
      { TITLE: { paragraphStyle: {}, textStyle: {} } },
      registry
    );

    expect(html).toContain('Title</span><br><span');
    expect(html).not.toContain('<br>&#8203;');
    expect(registry.toCSS()).toContain("font-family:'PT Sans',sans-serif;line-height:1.4881;");
  });

  test('keeps paragraph line spacing for homogeneous soft-break text', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT', lineSpacing: 150 },
        elements: [{ textRun: { content: 'First\u000bSecond\n', textStyle: {} } }]
      },
      metricDoc(), new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: {},
          textStyle: {
            fontSize: { magnitude: 12, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
          }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('line-height:1.941;');
    expect(registry.toCSS()).not.toContain('line-height:0;');
  });

  test('gives a leading blank soft-break line its own metrics', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT', lineSpacing: 115 },
        elements: [
          { textRun: { content: '\u000b', textStyle: {} } },
          {
            textRun: {
              content: 'Body\n',
              textStyle: { fontSize: { magnitude: 12, unit: 'PT' } }
            }
          }
        ]
      },
      metricDoc(), new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: {},
          textStyle: {
            fontSize: { magnitude: 11, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'PT Sans', weight: 400 }
          }
        }
      },
      registry
    );

    expect(html).toContain('>&#8203;<br><span');
    expect(registry.toCSS()).not.toContain('line-height:0;');
    expect(registry.toCSS()).toContain('line-height:1.15;');
    expect(registry.toCSS()).toContain('font-size:12pt;line-height:16.3691pt;');
  });

  test('does not invent a blank line for a break-only run after text', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [
          { textRun: { content: 'Title\u000b', textStyle: { weightedFontFamily: { fontFamily: 'Oswald' } } } },
          { textRun: { content: 'Author', textStyle: { weightedFontFamily: { fontFamily: 'Consolas' } } } },
          { textRun: { content: '\u000b', textStyle: { weightedFontFamily: { fontFamily: 'Georgia' } } } },
          { textRun: { content: 'Date\n', textStyle: { weightedFontFamily: { fontFamily: 'PT Sans' } } } }
        ]
      },
      metricDoc(), new Set(), [], null, '', '',
      { TITLE: { paragraphStyle: {}, textStyle: { fontSize: { magnitude: 24, unit: 'PT' } } } },
      registry
    );

    expect(html).toContain('Author</span><br><span');
    expect(html).not.toContain('&#8203;');
  });

  test('uses default Docs leading for generic runs in a mixed-font paragraph', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'SUBTITLE' },
        elements: [
          { textRun: { content: 'Author\u000b', textStyle: {} } },
          {
            textRun: {
              content: 'address@example.com\n',
              textStyle: { weightedFontFamily: { fontFamily: 'Consolas' } }
            }
          }
        ]
      },
      metricDoc(), new Set(), [], null, '', '',
      { SUBTITLE: { paragraphStyle: {}, textStyle: { fontSize: { magnitude: 14, unit: 'PT' } } } },
      registry
    );

    expect(html).toMatch(/^<h2[^>]*><span class="t\d+">Author<\/span><br>/);
    expect(registry.toCSS()).toContain('line-height:1.15;');
    expect(registry.toCSS()).toContain("font-family:'Consolas',sans-serif;line-height:");
  });

  test('sizes implicit-leading title lines independently when their font sizes differ', async () => {
    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [
          { textRun: { content: 'Title\u000b', textStyle: {} } },
          { textRun: { content: 'Tagline\n', textStyle: { fontSize: { magnitude: 12, unit: 'PT' } } } }
        ]
      },
      metricDoc(), new Set(), [], null, '', '',
      {
        TITLE: {
          paragraphStyle: {},
          textStyle: {
            fontSize: { magnitude: 24, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'Lalezar', weight: 400 }
          }
        }
      },
      registry
    );

    expect(html).toMatch(/<span class="t\d+">Title<\/span><br><span class="t\d+">Tagline<\/span>/);
    expect(registry.toCSS()).toContain('line-height:0;');
    expect(registry.toCSS()).toContain('font-size:12pt;line-height:');
  });

  test('retains metric-derived trailing leading for a title using body spacing', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'TITLE' },
        elements: [{ textRun: { content: 'Compact title\n', textStyle: {} } }]
      },
      metricDoc(), new Set(), [], null, '', '',
      {
        NORMAL_TEXT: {
          paragraphStyle: { spaceAbove: { magnitude: 10, unit: 'PT' } },
          textStyle: {}
        },
        TITLE: {
          paragraphStyle: {
            spaceAbove: { magnitude: 10, unit: 'PT' },
            spaceBelow: { magnitude: 6, unit: 'PT' }
          },
          textStyle: {
            fontSize: { magnitude: 24, unit: 'PT' },
            weightedFontFamily: { fontFamily: 'Lalezar', weight: 400 }
          }
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
      metricDoc(), new Set(), [], null, '', '',
      {
        TITLE: {
          paragraphStyle: {},
          textStyle: { fontSize: { magnitude: 24, unit: 'PT' } }
        }
      },
      registry
    );

    expect(registry.toCSS()).toContain('line-height:2.3505;position:relative;top:-9px;');
  });

  test('does not deduct paragraph margins from an intentional blank line', async () => {
    const registry = new StyleRegistry();
    await renderParagraph(
      {
        paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
        elements: [{ textRun: { content: '\n', textStyle: {} } }]
      },
      metricDoc(), new Set(), [], null, '', '',
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
      metricDoc({ positionedObjects: { anchored: {} } }), new Set(), [], null, '', '',
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

    expect(registry.toCSS()).toContain('min-height:16px;');
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
