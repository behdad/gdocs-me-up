const sharp = require('sharp');
const { optimizeImage } = require('./images');

describe('optimizeImage', () => {
  test('preserves JPEG data and reports the correct extension', async () => {
    const source = await sharp({
      create: { width: 10, height: 10, channels: 3, background: '#336699' }
    }).jpeg().toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('jpg');
    expect(result.mimeType).toBe('image/jpeg');
    expect(result.converted).toBe(false);
    expect(result.buffer).toEqual(source);
  });

  test('preserves PNG images with transparency', async () => {
    const source = await sharp({
      create: { width: 10, height: 10, channels: 4, background: { r: 0, g: 0, b: 0, alpha: 0.5 } }
    }).png().toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('png');
    expect(result.converted).toBe(false);
  });

  test('preserves an opaque flat-color PNG when JPEG is not a useful saving', async () => {
    const source = await sharp({
      create: { width: 64, height: 64, channels: 3, background: '#ffffff' }
    }).png().toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('png');
    expect(result.converted).toBe(false);
    expect(result.buffer).toEqual(source);
  });

  test('can convert an opaque PNG whose file includes an alpha channel', async () => {
    const width = 256;
    const height = 256;
    const pixels = Buffer.alloc(width * height * 4);
    let state = 0x87654321;
    for (let i = 0; i < pixels.length; i += 4) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      pixels[i] = state & 0xff;
      pixels[i + 1] = (state >>> 8) & 0xff;
      pixels[i + 2] = (state >>> 16) & 0xff;
      pixels[i + 3] = 255;
    }
    const source = await sharp(pixels, { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('jpg');
    expect(result.converted).toBe(true);
  });

  test('uses JPEG when it materially reduces an opaque photographic PNG', async () => {
    const width = 256;
    const height = 256;
    const pixels = Buffer.alloc(width * height * 3);
    let state = 0x12345678;
    for (let i = 0; i < pixels.length; i++) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      pixels[i] = state & 0xff;
    }
    const source = await sharp(pixels, { raw: { width, height, channels: 3 } })
      .png()
      .toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('jpg');
    expect(result.converted).toBe(true);
    expect(result.buffer.length).toBeLessThan(source.length * 0.8);
  });
});
