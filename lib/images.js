/**
 * Image handling functions for Google Docs export
 */

const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const { imageCrop } = require('./image-crop');

const FORMAT_INFO = {
  avif: { extension: 'avif', mimeType: 'image/avif' },
  gif: { extension: 'gif', mimeType: 'image/gif' },
  heif: { extension: 'heic', mimeType: 'image/heic' },
  jpeg: { extension: 'jpg', mimeType: 'image/jpeg' },
  jpg: { extension: 'jpg', mimeType: 'image/jpeg' },
  png: { extension: 'png', mimeType: 'image/png' },
  svg: { extension: 'svg', mimeType: 'image/svg+xml' },
  tiff: { extension: 'tiff', mimeType: 'image/tiff' },
  webp: { extension: 'webp', mimeType: 'image/webp' }
};

const EFFECTIVE_OPACITY = {
  // Google Docs sometimes adds a faint antialiased alpha fringe to an
  // otherwise opaque photograph. Treat that encoding artifact as opaque, but
  // keep images with even a small amount of intentional transparency as PNG.
  minAlpha: 253,
  maxTranslucentFraction: 0.0025,
  minMeanAlpha: 0.999
};

async function isEffectivelyOpaque(buffer) {
  const { data } = await sharp(buffer)
    .ensureAlpha()
    .extractChannel('alpha')
    .raw()
    .toBuffer({ resolveWithObject: true });

  const maxTranslucentPixels = Math.floor(
    data.length * EFFECTIVE_OPACITY.maxTranslucentFraction
  );
  let translucentPixels = 0;
  let alphaTotal = 0;

  for (const alpha of data) {
    alphaTotal += alpha;
    if (alpha < EFFECTIVE_OPACITY.minAlpha) {
      translucentPixels++;
      if (translucentPixels > maxTranslucentPixels) return false;
    }
  }

  return alphaTotal / (data.length * 255) >= EFFECTIVE_OPACITY.minMeanAlpha;
}

const PHOTOGRAPHIC_ENTROPY_THRESHOLD = 4.5;

async function isLikelyPhotographic(buffer) {
  const { entropy } = await sharp(buffer).stats();
  return entropy >= PHOTOGRAPHIC_ENTROPY_THRESHOLD;
}

// Budget source pixels for the visible frame, including pixels outside a crop.
function getResizeOptions(metadata, options) {
  const pixelRatio = options.pixelRatio || 2;
  const frameScale = options.displayWidth && options.maxDisplayWidth
    ? Math.min(1, options.maxDisplayWidth / options.displayWidth)
    : 1;
  const crop = imageCrop(options.cropProperties);
  const maxWidth = options.displayWidth
    ? Math.ceil(options.displayWidth * frameScale * pixelRatio / (crop?.width || 1))
    : Math.min(metadata.width, (options.maxDisplayWidth || metadata.width) * pixelRatio);
  const maxHeight = options.displayHeight
    ? Math.ceil(options.displayHeight * frameScale * pixelRatio / (crop?.height || 1))
    : metadata.height;

  if (!metadata.width || !metadata.height) return null;
  // A crop can stretch the source differently along each axis. Keep enough
  // pixels for the more enlarged axis instead of undersampling that direction.
  const scale = Math.min(1, (crop ? Math.max : Math.min)(
    maxWidth / metadata.width, maxHeight / metadata.height
  ));
  if (scale >= 1) return null;
  return {
    width: Math.max(1, Math.round(metadata.width * scale)),
    height: Math.max(1, Math.round(metadata.height * scale)),
    fit: 'fill'
  };
}

function imagePipeline(buffer, resize) {
  const pipeline = sharp(buffer);
  return resize ? pipeline.resize(resize) : pipeline;
}

function webpPipeline(buffer, resize) {
  return imagePipeline(buffer, resize)
    .webp({ quality: 84, effort: 6, smartSubsample: true });
}

async function optimizePNG(buffer, resize, metadata) {
  // Preserve higher bit depths when there is no need to resample them.
  if (metadata.depth !== 'uchar') {
    if (!resize) return buffer;
    return imagePipeline(buffer, resize).toColourspace('rgb16')
      .png({ compressionLevel: 9, adaptiveFiltering: true }).toBuffer();
  }

  const { data, info } = await imagePipeline(buffer, resize)
    .toColourspace('srgb').ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  let opaque = true;
  let gray = true;
  const colors = new Set();
  for (let i = 0; i < data.length; i += 4) {
    opaque &&= data[i + 3] === 255;
    gray &&= data[i] === data[i + 1] && data[i] === data[i + 2];
    if (colors.size <= 256) colors.add(data.readUInt32BE(i));
  }
  const pipeline = () => {
    let image = sharp(data, { raw: { width: info.width, height: info.height, channels: 4 } });
    if (opaque) image = image.removeAlpha();
    if (gray) image = image.toColourspace('b-w');
    return image;
  };
  let best = resize ? null : buffer;
  for (const adaptiveFiltering of [false, true]) {
    const candidate = await pipeline().png({ compressionLevel: 9, adaptiveFiltering }).toBuffer();
    if (!best || candidate.length < best.length) best = candidate;
  }
  // An 8-bit grayscale PNG already uses one byte per pixel. Only try a
  // palette there when it can lower the bit depth to four bits or fewer.
  if (colors.size <= 256 && (!gray || colors.size <= 16)) {
    const candidate = await pipeline().png({
      compressionLevel: 9, palette: true, colours: Math.max(2, colors.size),
      quality: 100, dither: 0, effort: 10
    }).toBuffer();
    if (candidate.length < best.length) {
      // Quantizers may merge even existing colours. Only accept a palette if
      // every decoded RGBA sample is unchanged, including transparent pixels.
      const decoded = await sharp(candidate).toColourspace('srgb').ensureAlpha().raw().toBuffer();
      if (decoded.equals(data)) best = candidate;
    }
  }
  return best;
}

async function optimizeImage(buffer, options = {}) {
  const pipeline = sharp(buffer, { animated: true });
  const metadata = await pipeline.metadata();
  const source = FORMAT_INFO[metadata.format] || {
    extension: metadata.format || 'bin',
    mimeType: `image/${metadata.format || 'octet-stream'}`
  };

  if (!['jpeg', 'png', 'webp'].includes(metadata.format) || metadata.pages > 1) {
    return { buffer, ...source, converted: false };
  }

  const resize = getResizeOptions(metadata, options);

  if (metadata.format === 'webp') {
    if (!resize) return { buffer, ...source, converted: false };
    // Resizing should not add another generation of lossy compression.
    const webp = await imagePipeline(buffer, resize).webp({ lossless: true, effort: 6 }).toBuffer();
    return webp.length < buffer.length
      ? { buffer: webp, ...source, converted: true }
      : { buffer, ...source, converted: false };
  }

  if (metadata.format === 'jpeg') {
    const webp = await webpPipeline(buffer, resize).toBuffer();
    return {
      buffer: webp,
      extension: 'webp',
      mimeType: 'image/webp',
      converted: true
    };
  }

  let png = await optimizePNG(buffer, resize, metadata);
  if (resize) {
    // Resampling line art introduces intermediate colours and can make a
    // small palette image much larger. Keep the sharper original if its
    // losslessly optimized encoding is smaller than the resized version.
    const originalPNG = await optimizePNG(buffer, null, metadata);
    if (originalPNG.length < png.length) png = originalPNG;
  }
  if (metadata.hasAlpha && !(await isEffectivelyOpaque(buffer))) {
    return { buffer: png, ...source, converted: !png.equals(buffer) };
  }

  if (!(await isLikelyPhotographic(buffer))) {
    return { buffer: png, ...source, converted: !png.equals(buffer) };
  }

  const webp = await imagePipeline(buffer, resize)
    .flatten({ background: '#ffffff' })
    .webp({ quality: 84, effort: 6, smartSubsample: true })
    .toBuffer();

  if (webp.length <= png.length * 0.8) {
    return {
      buffer: webp,
      extension: 'webp',
      mimeType: 'image/webp',
      converted: true
    };
  }

  return { buffer: png, ...source, converted: !png.equals(buffer) };
}

async function writeOptimizedImage(buffer, imagesDir, baseName, options = {}) {
  const metadata = await sharp(buffer).metadata();
  const optimized = await optimizeImage(buffer, options);
  const fileName = `${baseName}.${optimized.extension}`;
  const filePath = path.join(imagesDir, fileName);
  for (const existing of fs.readdirSync(imagesDir)) {
    if (existing !== fileName && existing.startsWith(`${baseName}.`)) {
      fs.unlinkSync(path.join(imagesDir, existing));
    }
  }
  fs.writeFileSync(filePath, optimized.buffer);
  return { ...optimized, fileName, filePath, sourceWidth: metadata.width, sourceHeight: metadata.height };
}

module.exports = {
  optimizeImage,
  writeOptimizedImage
};
