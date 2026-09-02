const {
  parseCliArguments,
  renderExternalStylesheetLinks,
  renderExternalScriptTags,
  documentRequestParameters,
  resolveServiceAccountKeyFile
} = require('../gdocs-me-up');
const path = require('path');

test('resolves authentication independently of the working directory', () => {
  expect(resolveServiceAccountKeyFile({})).toBe(
    path.join(__dirname, '..', 'service_account.json')
  );
  expect(resolveServiceAccountKeyFile({ SERVICE_ACCOUNT_KEY_FILE: '/keys/docs.json' }))
    .toBe('/keys/docs.json');
});

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
      stylesheets: ['../style.css'],
      scripts: [],
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
    ]).stylesheets).toEqual(['../style.css']);
  });

  test('rejects a missing stylesheet href', () => {
    expect(() => parseCliArguments([
      'document-id',
      'output',
      '--stylesheet'
    ])).toThrow('--stylesheet requires an href');
  });

  test('preserves the order of repeated stylesheets and scripts', () => {
    const options = parseCliArguments([
      'document-id',
      'output',
      '--stylesheet',
      'base.css',
      '--script=first.js',
      '--stylesheet=theme.css',
      '--script',
      'last.js'
    ]);

    expect(options.stylesheets).toEqual(['base.css', 'theme.css']);
    expect(options.scripts).toEqual(['first.js', 'last.js']);
  });

  test('rejects a missing script src', () => {
    expect(() => parseCliArguments([
      'document-id',
      'output',
      '--script'
    ])).toThrow('--script requires a src');
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

  test('renders resources in their supplied order and escapes attributes', () => {
    expect(renderExternalStylesheetLinks(['one.css', 'two.css?x=1&y=2'])).toEqual([
      '  <link rel="stylesheet" href="one.css">',
      '  <link rel="stylesheet" href="two.css?x=1&amp;y=2">'
    ]);
    expect(renderExternalScriptTags(['one.js', 'two.js?x="yes"'])).toEqual([
      '  <script src="one.js"></script>',
      '  <script src="two.js?x=&quot;yes&quot;"></script>'
    ]);
  });
});
