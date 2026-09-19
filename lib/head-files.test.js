const fs = require('fs');
const os = require('os');
const path = require('path');

const mockGetDocument = jest.fn();
jest.mock('googleapis', () => ({
  google: {
    auth: {
      GoogleAuth: jest.fn(() => ({ getClient: async () => ({}) }))
    },
    docs: jest.fn(() => ({ documents: { get: mockGetDocument } }))
  }
}));

const { exportDocToHTML, parseCliArguments } = require('../gdocs-me-up');

describe('head fragments in exported documents', () => {
  let directory;
  let title;

  beforeEach(() => {
    directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gdocs-head-files-'));
    title = `خانه & "Home" <draft> 'quoted' $& {{title}}`;
    mockGetDocument.mockReset();
    mockGetDocument.mockImplementation(async () => ({
      data: {
        title,
        body: { content: [{ paragraph: {
          elements: [{ textRun: { content: 'Body stays the same.\n' } }]
        } }] }
      }
    }));
    jest.spyOn(console, 'log').mockImplementation(() => {});
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
    fs.rmSync(directory, { recursive: true, force: true });
  });

  test('adds ordered fragments while preserving the rest of the export exactly', async () => {
    const output = path.join(directory, 'output');
    await exportDocToHTML('document-id', output);
    const baseline = fs.readFileSync(path.join(output, 'index.html'), 'utf8');
    const firstFile = path.join(directory, 'shared head.html');
    const secondFile = path.join(directory, 'extra.html');
    fs.writeFileSync(firstFile,
      '<meta property="og:title" content="{{title}}">\n' +
      "<meta name='description' content='{{ title }}'>"
    );
    const secondFragment = '<!-- Shared metadata -->\n' +
      '<meta property="og:image" content="https://example.com/card.png?x=1&amp;y=2">';
    fs.writeFileSync(secondFile, secondFragment);

    const options = parseCliArguments([
      'document-id', output,
      '--head-file', firstFile,
      `--head-file=${secondFile}`
    ]);
    await exportDocToHTML(options.docId, options.outDir, options);
    const html = fs.readFileSync(path.join(output, 'index.html'), 'utf8');
    const escapedTitle = 'خانه &amp; &quot;Home&quot; &lt;draft&gt; &#39;quoted&#39; $&amp; {{title}}';
    const expected = `<meta property="og:title" content="${escapedTitle}">\n` +
      `<meta name='description' content='${escapedTitle}'>\n${secondFragment}\n`;
    expect(html).toBe(baseline.replace('</head>', () => `${expected}</head>`));
  });

  test('reuses one fragment for different document titles', async () => {
    const headFile = path.join(directory, 'shared.html');
    fs.writeFileSync(headFile, '<meta property="og:title" content="{{title}}">');
    const output = path.join(directory, 'output');
    for (const nextTitle of ['First piece', 'قطعهٔ بعدی']) {
      title = nextTitle;
      await exportDocToHTML('document-id', output, { headFiles: [headFile] });
      const html = fs.readFileSync(path.join(output, 'index.html'), 'utf8');
      expect(html).toContain(`<meta property="og:title" content="${nextTitle}">`);
    }
  });

  test('reports missing fragments before fetching or creating output', async () => {
    const output = path.join(directory, 'output');
    const headFile = path.join(directory, 'missing.html');
    await expect(exportDocToHTML('document-id', output, { headFiles: [headFile] }))
      .rejects.toThrow(`Could not read head file "${headFile}"`);
    expect(mockGetDocument).not.toHaveBeenCalled();
    expect(fs.existsSync(output)).toBe(false);
  });
});
