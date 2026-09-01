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

/**
 * Preserve already-efficient formats and try a high-quality JPEG for opaque PNGs.
 * Google commonly serves pasted photographs as very large PNGs. The JPEG candidate is
 * selected only when it is at least 20% smaller, avoiding pointless lossy conversion of
 * screenshots, diagrams, and other PNGs that already compress well.
 */
async function optimizeImage(buffer) {
  const pipeline = sharp(buffer, { animated: true });
  const metadata = await pipeline.metadata();
  const source = FORMAT_INFO[metadata.format] || {
    extension: metadata.format || 'bin',
    mimeType: `image/${metadata.format || 'octet-stream'}`
  };

  if (metadata.format !== 'png' || metadata.pages > 1) {
    return { buffer, ...source, converted: false };
  }

  if (metadata.hasAlpha) {
    const stats = await pipeline.stats();
    if (!stats.isOpaque) return { buffer, ...source, converted: false };
  }

  const jpeg = await sharp(buffer)
    .flatten({ background: '#ffffff' })
    .jpeg({ quality: 92, chromaSubsampling: '4:4:4', mozjpeg: true })
    .toBuffer();

  if (jpeg.length <= buffer.length * 0.8) {
    return {
      buffer: jpeg,
      extension: 'jpg',
      mimeType: 'image/jpeg',
      converted: true
    };
  }

  return { buffer, ...source, converted: false };
}

async function writeOptimizedImage(buffer, imagesDir, baseName) {
  const optimized = await optimizeImage(buffer);
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
