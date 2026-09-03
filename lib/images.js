/**
 * Image handling functions for Google Docs export
 */

const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

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

/**
 * Preserve already-efficient formats and try a high-quality WebP for JPEGs and opaque PNGs.
 * Google commonly serves pasted photographs as JPEGs or very large PNGs. The WebP candidate is
 * selected only when it is at least 20% smaller, avoiding pointless lossy conversion of
 * screenshots, diagrams, and other PNGs that already compress well.
 */
function getResizeOptions(metadata, options) {
  const pixelRatio = options.pixelRatio || 2;
  const maxWidth = options.displayWidth
    ? Math.ceil(options.displayWidth * pixelRatio)
    : metadata.width;
  const maxHeight = options.displayHeight
    ? Math.ceil(options.displayHeight * pixelRatio)
    : metadata.height;

  if (!metadata.width || !metadata.height) return null;
  const scale = Math.min(1, maxWidth / metadata.width, maxHeight / metadata.height);
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

async function optimizeImage(buffer, options = {}) {
  const pipeline = sharp(buffer, { animated: true });
  const metadata = await pipeline.metadata();
  const source = FORMAT_INFO[metadata.format] || {
    extension: metadata.format || 'bin',
    mimeType: `image/${metadata.format || 'octet-stream'}`
  };

  if (!['jpeg', 'png'].includes(metadata.format) || metadata.pages > 1) {
    return { buffer, ...source, converted: false };
  }

  const resize = getResizeOptions(metadata, options);

  if (metadata.format === 'jpeg') {
    const webp = await webpPipeline(buffer, resize).toBuffer();
    return {
      buffer: webp,
      extension: 'webp',
      mimeType: 'image/webp',
      converted: true
    };
  }

  if (metadata.hasAlpha && !(await isEffectivelyOpaque(buffer))) {
    if (!resize) return { buffer, ...source, converted: false };
    const png = await imagePipeline(buffer, resize)
      .png({ compressionLevel: 9, adaptiveFiltering: true })
      .toBuffer();
    return { buffer: png, ...source, converted: true };
  }

  if (!(await isLikelyPhotographic(buffer))) {
    if (!resize) return { buffer, ...source, converted: false };
    const png = await imagePipeline(buffer, resize)
      .png({ compressionLevel: 9, adaptiveFiltering: true })
      .toBuffer();
    return { buffer: png, ...source, converted: true };
  }

  const webp = await imagePipeline(buffer, resize)
    .flatten({ background: '#ffffff' })
    .webp({ quality: 84, effort: 6, smartSubsample: true })
    .toBuffer();

  const png = resize
    ? await imagePipeline(buffer, resize)
      .png({ compressionLevel: 9, adaptiveFiltering: true })
      .toBuffer()
    : buffer;

  if (webp.length <= png.length * 0.8) {
    return {
      buffer: webp,
      extension: 'webp',
      mimeType: 'image/webp',
      converted: true
    };
  }

  return { buffer: png, ...source, converted: Boolean(resize) };
}

async function writeOptimizedImage(buffer, imagesDir, baseName, options = {}) {
  const optimized = await optimizeImage(buffer, options);
  const fileName = `${baseName}.${optimized.extension}`;
  const filePath = path.join(imagesDir, fileName);
  for (const existing of fs.readdirSync(imagesDir)) {
    if (existing !== fileName && existing.startsWith(`${baseName}.`)) {
      fs.unlinkSync(path.join(imagesDir, existing));
    }
  }
  fs.writeFileSync(filePath, optimized.buffer);
  return { ...optimized, fileName, filePath };
}

module.exports = {
  optimizeImage,
  writeOptimizedImage
};
