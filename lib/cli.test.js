const {
  parseCliArguments,
  renderExternalStylesheetLink,
  documentRequestParameters
} = require('../gdocs-me-up');

test('exports the document without flattening suggestions into its content', () => {
  expect(documentRequestParameters('document-id')).toEqual({
    documentId: 'document-id',
    suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS'
  });
});

describe('custom stylesheet option', () => {
  test('parses a relative stylesheet href', () => {
    expect(parseCliArguments([
      'document-id',
      'output',
      '--stylesheet',
      '../style.css'
    ])).toEqual({
      docId: 'document-id',
      outDir: 'output',
      stylesheet: '../style.css',
      htmlName: undefined,
      imagesDir: undefined,
      help: false
    });
  });

  test('accepts the equals form', () => {
    expect(parseCliArguments([
      'document-id',
      'output',
      '--stylesheet=../style.css'
    ]).stylesheet).toBe('../style.css');
  });

  test('rejects a missing stylesheet href', () => {
    expect(() => parseCliArguments([
      'document-id',
      'output',
      '--stylesheet'
    ])).toThrow('--stylesheet requires an href');
  });

  test('parses custom HTML and images names', () => {
    const options = parseCliArguments([
      'document-id',
      'output',
      '--html-name',
      'article.html',
      '--images-dir=assets'
    ]);

    expect(options.htmlName).toBe('article.html');
    expect(options.imagesDir).toBe('assets');
  });

  test('keeps generated files inside the output directory', () => {
    expect(() => parseCliArguments([
      'document-id',
      'output',
      '--html-name',
      '../article.html'
    ])).toThrow('Invalid HTML filename provided');
    expect(() => parseCliArguments([
      'document-id',
      'output',
      '--images-dir',
      '../assets'
    ])).toThrow('Invalid images directory provided');
  });

  test('escapes the href for an HTML attribute', () => {
    expect(renderExternalStylesheetLink('../style.css?one=1&two="2"')).toBe(
      '  <link rel="stylesheet" href="../style.css?one=1&amp;two=&quot;2&quot;">'
    );
  });
});
