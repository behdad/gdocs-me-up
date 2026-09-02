const {
  generateGlobalCSS,
  prepareTableOfContentsLinks,
  renderParagraph
} = require('../gdocs-me-up');
const { StyleRegistry } = require('./styles');

describe('table of contents', () => {
  test('retains TOC paragraph geometry instead of overriding it globally', () => {
    const css = generateGlobalCSS({}, 624);
    expect(css).not.toMatch(/\.doc-toc p\s*\{/);
    expect(css).not.toContain('.toc-level-');
  });

  test('connects empty generated TOC links to their ordered headings', async () => {
    const tocParagraph = {
      paragraphStyle: {
        namedStyleType: 'NORMAL_TEXT',
        indentStart: { magnitude: 18, unit: 'PT' },
        indentFirstLine: { magnitude: 18, unit: 'PT' }
      },
      elements: [{
        textRun: {
          content: 'Repeated heading\n',
          textStyle: { link: { headingId: '' }, underline: true }
        }
      }]
    };
    const targetParagraph = {
      paragraphStyle: { namedStyleType: 'HEADING_2' },
      elements: [{ textRun: { content: 'Repeated heading\n', textStyle: {} } }]
    };
    const doc = {
      body: {
        content: [
          { tableOfContents: { content: [{ paragraph: tocParagraph }] } },
          { startIndex: 42, paragraph: targetParagraph }
        ]
      }
    };

    prepareTableOfContentsLinks(doc);

    expect(tocParagraph.elements[0].textRun.textStyle.link.headingId).toBe('toc-42');
    expect(targetParagraph.paragraphStyle.headingId).toBe('toc-42');

    const registry = new StyleRegistry();
    const { html } = await renderParagraph(
      tocParagraph,
      doc,
      new Set(),
      [],
      null,
      '',
      '',
      { NORMAL_TEXT: { paragraphStyle: {}, textStyle: {} } },
      registry
    );
    expect(html).toContain('<a href="#heading-toc-42">Repeated heading</a>');
    expect(registry.toCSS()).toContain('margin-inline-start:24px;');
  });
});
