const fs = require('fs');
const os = require('os');
const path = require('path');
const sharp = require('sharp');
const cheerio = require('cheerio');

const mockGetDocument = jest.fn();
const mockRequest = jest.fn();
jest.mock('googleapis', () => ({ google: {
  auth: { GoogleAuth: jest.fn(() => ({ getClient: async () => ({ request: mockRequest }) })) },
  docs: jest.fn(() => ({ documents: { get: mockGetDocument } }))
} }));
const { exportDocToHTML } = require('../gdocs-me-up');

let directory;
beforeEach(() => {
  directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gdocs-images-'));
  jest.spyOn(console, 'log').mockImplementation(() => {});
});
afterEach(() => {
  jest.restoreAllMocks();
  fs.rmSync(directory, { recursive: true, force: true });
});

test.each(['inline', 'positioned'])('exports %s image crops at the page width', async kind => {
  const source = await sharp({ create: {
    width: 2000, height: 2000, channels: 3, background: '#336699'
  } }).jpeg().toBuffer();
  mockRequest.mockResolvedValue({ data: source });
  const embeddedObject = {
    title: 'A & B',
    size: { width: { magnitude: 750 }, height: { magnitude: 750 } },
    imageProperties: { contentUri: 'https://example.test/image', cropProperties: {
      offsetLeft: .125, offsetRight: .375, offsetTop: .375, offsetBottom: .125
    } }
  };
  const paragraph = { elements: [{ textRun: { content: '\n' } }] };
  const doc = { title: 'Images', body: { content: [{ paragraph }] }, documentStyle: {
    pageSize: { width: { magnitude: 444 } },
    marginLeft: { magnitude: 72 }, marginRight: { magnitude: 72 },
    documentFormat: { documentMode: 'PAGES' }
  } };
  if (kind === 'inline') {
    doc.inlineObjects = { example: { inlineObjectProperties: { embeddedObject } } };
    paragraph.elements.unshift({ inlineObjectElement: { inlineObjectId: 'example' } });
  } else {
    doc.positionedObjects = { example: { positionedObjectProperties: { embeddedObject } } };
    paragraph.positionedObjectIds = ['example'];
  }
  mockGetDocument.mockResolvedValue({ data: doc });
  await exportDocToHTML('example', directory);
  const $ = cheerio.load(fs.readFileSync(path.join(directory, 'index.html'), 'utf8'));
  expect($('img')).toHaveLength(1);
  expect($('img').attr('alt')).toBe('A & B');
  expect($('img').parent().is('span')).toBe(true);
  expect($('style').text()).toContain('width:200%;height:200%;left:-25%;top:-75%;');
  expect($('style').text()).toContain('position:relative;overflow:hidden;');
  const metadata = await sharp(path.join(directory, $('img').attr('src'))).metadata();
  // A 400px page, a half-width crop, and two pixels per CSS pixel.
  expect(metadata).toMatchObject({ width: 1600, height: 1600 });
});
