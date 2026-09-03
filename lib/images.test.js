const sharp = require('sharp');
const { optimizeImage } = require('./images');

describe('optimizeImage', () => {
  test('converts JPEG data to WebP', async () => {
    const source = await sharp({
      create: { width: 10, height: 10, channels: 3, background: '#336699' }
    }).jpeg().toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('webp');
    expect(result.mimeType).toBe('image/webp');
    expect(result.converted).toBe(true);
    expect((await sharp(result.buffer).metadata()).format).toBe('webp');
  });

  test('preserves PNG images with transparency', async () => {
    const source = await sharp({
      create: { width: 10, height: 10, channels: 4, background: { r: 0, g: 0, b: 0, alpha: 0.5 } }
    }).png().toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('png');
    expect(result.converted).toBe(false);
  });

  test('preserves an opaque flat-color PNG', async () => {
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
    expect(result.extension).toBe('webp');
    expect(result.converted).toBe(true);
  });

  test('converts a photographic PNG with a tiny alpha fringe', async () => {
    const width = 512;
    const height = 512;
    const pixels = Buffer.alloc(width * height * 4);
    let state = 0x2468ace0;
    for (let i = 0, pixel = 0; i < pixels.length; i += 4, pixel++) {
      state ^= state << 13;
      state ^= state >>> 17;
      state ^= state << 5;
      pixels[i] = state & 0xff;
      pixels[i + 1] = (state >>> 8) & 0xff;
      pixels[i + 2] = (state >>> 16) & 0xff;
      pixels[i + 3] = pixel < 500 ? 220 : 255;
    }
    const source = await sharp(pixels, { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('webp');
    expect(result.converted).toBe(true);
  });

  test('preserves a PNG with sparse but meaningful transparency', async () => {
    const width = 256;
    const height = 256;
    const pixels = Buffer.alloc(width * height * 4, 255);
    for (let pixel = 0; pixel < 700; pixel++) {
      pixels[pixel * 4 + 3] = 0;
    }
    const source = await sharp(pixels, { raw: { width, height, channels: 4 } })
      .png()
      .toBuffer();

    const result = await optimizeImage(source);
    expect(result.extension).toBe('png');
    expect(result.converted).toBe(false);
  });

  test('uses WebP when it materially reduces an opaque photographic PNG', async () => {
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
    expect(result.extension).toBe('webp');
    expect(result.converted).toBe(true);
    expect(result.buffer.length).toBeLessThan(source.length * 0.8);
  });

  test('caps JPEG dimensions at twice the display size', async () => {
    const source = await sharp({
      create: { width: 1200, height: 800, channels: 3, background: '#336699' }
    }).jpeg().toBuffer();

    const result = await optimizeImage(source, {
      displayWidth: 300,
      displayHeight: 200
    });
    const metadata = await sharp(result.buffer).metadata();

    expect(result.extension).toBe('webp');
    expect(result.converted).toBe(true);
    expect(metadata.width).toBe(600);
    expect(metadata.height).toBe(400);
  });

  test('converts without enlarging an image smaller than twice its display size', async () => {
    const source = await sharp({
      create: { width: 500, height: 300, channels: 3, background: '#336699' }
    }).jpeg().toBuffer();

    const result = await optimizeImage(source, {
      displayWidth: 300,
      displayHeight: 200
    });

    const metadata = await sharp(result.buffer).metadata();
    expect(result.extension).toBe('webp');
    expect(result.converted).toBe(true);
    expect(metadata.width).toBe(500);
    expect(metadata.height).toBe(300);
  });
});
