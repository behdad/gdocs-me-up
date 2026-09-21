const sharp = require('sharp');
const { optimizeImage } = require('./images');

const rgba = buffer => sharp(buffer).toColourspace('srgb').ensureAlpha().raw().toBuffer();

async function expectLosslessPNG(source, result) {
  expect(result.extension).toBe('png');
  expect(result.buffer.length).toBeLessThanOrEqual(source.length);
  expect((await rgba(result.buffer)).equals(await rgba(source))).toBe(true);
}

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
    await expectLosslessPNG(source, result);
  });

  test('preserves an opaque flat-color PNG', async () => {
    const source = await sharp({
      create: { width: 64, height: 64, channels: 3, background: '#ffffff' }
    }).png().toBuffer();

    const result = await optimizeImage(source);
    await expectLosslessPNG(source, result);
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
    await expectLosslessPNG(source, result);
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

describe('image size and lossless compression regressions', () => {
  test('caps oversized document images to twice the content width', async () => {
    const source = await sharp({ create: {
      width: 2048, height: 1536, channels: 3, background: '#336699'
    } }).jpeg().toBuffer();
    const result = await optimizeImage(source, {
      displayWidth: 2048, displayHeight: 1536, maxDisplayWidth: 690
    });
    expect(await sharp(result.buffer).metadata()).toMatchObject({ width: 1380, height: 1035 });
  });

  test('retains two source pixels per displayed pixel inside a crop', async () => {
    const source = await sharp({ create: {
      width: 2000, height: 2000, channels: 3, background: '#336699'
    } }).jpeg().toBuffer();
    const result = await optimizeImage(source, {
      displayWidth: 500, displayHeight: 500, maxDisplayWidth: 300,
      cropProperties: { offsetLeft: 0.25, offsetRight: 0.25, offsetTop: 0.25, offsetBottom: 0.25 }
    });
    expect(await sharp(result.buffer).metadata()).toMatchObject({ width: 1200, height: 1200 });
  });

  test.each(['gray', 'color', 'alpha'])('compresses a %s diagram without changing decoded pixels', async kind => {
    const width = 240, height = 240;
    const pixels = Buffer.alloc(width * height * 4);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = 4 * (y * width + x);
        const value = ((x >> 3) + (y >> 3)) % 8 * 31;
        pixels[i] = value;
        pixels[i + 1] = kind === 'color' ? 255 - value : value;
        pixels[i + 2] = value;
        pixels[i + 3] = kind === 'alpha' ? value : 255;
      }
    }
    const source = await sharp(pixels, { raw: { width, height, channels: 4 } })
      .png({ compressionLevel: 0 }).toBuffer();
    const result = await optimizeImage(source);
    await expectLosslessPNG(source, result);
    expect(result.buffer.length).toBeLessThan(source.length / 10);
  });

  test('does not rewrite an existing WebP unless resizing saves bytes', async () => {
    const source = await sharp({ create: {
      width: 100, height: 100, channels: 3, background: '#336699'
    } }).webp().toBuffer();
    expect((await optimizeImage(source)).buffer).toEqual(source);
  });

  test('keeps sufficient crop resolution when the frame stretches one axis', async () => {
    const source = await sharp({ create: {
      width: 2000, height: 2000, channels: 3, background: '#336699'
    } }).jpeg().toBuffer();
    const result = await optimizeImage(source, {
      displayWidth: 300, displayHeight: 100,
      cropProperties: { offsetLeft: .25, offsetRight: .25 }
    });
    expect(await sharp(result.buffer).metadata()).toMatchObject({ width: 1200, height: 1200 });
  });
});

test('keeps the original resolution when resampling a palette diagram costs more bytes', async () => {
  const width = 574, height = 378;
  const pixels = Buffer.alloc(width * height * 3, 255);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (x % 19 < 3 || y % 23 < 3) {
        const i = 3 * (y * width + x);
        pixels[i] = 0; pixels[i + 1] = 0; pixels[i + 2] = 0;
      }
    }
  }
  const source = await sharp(pixels, { raw: { width, height, channels: 3 } })
    .png({ palette: true }).toBuffer();
  const result = await optimizeImage(source, {
    displayWidth: width, displayHeight: height, maxDisplayWidth: 260
  });
  await expectLosslessPNG(source, result);
  expect(await sharp(result.buffer).metadata()).toMatchObject({ width, height });
});
