'use strict';

// Docs specifies a rectangle in the original image, followed by a clockwise
// rotation of that rectangle around its centre. CSS clips the inverse transform
// of the image through a fixed, responsive viewport.
function imageCrop(crop = {}) {
  const values = ['offsetLeft', 'offsetTop', 'offsetRight', 'offsetBottom', 'angle']
    .map(key => Number(crop?.[key] ?? 0))
    .map(value => Math.abs(value) < 1e-10 ? 0 : value);
  if (!values.every(Number.isFinite)) throw new Error('Invalid image crop');
  const [left, top, right, bottom, angle] = values;
  if (!values.some(Boolean)) return null;
  const width = 1 - left - right;
  const height = 1 - top - bottom;
  if (width <= 0 || height <= 0) throw new Error('Empty image crop');
  return { left, top, width, height, angle };
}

function cropImageStyle(crop, { displayWidth, displayHeight, sourceWidth, sourceHeight }) {
  const n = value => String(Math.round(value * 1e8) / 1e8);
  let css = 'position:absolute;max-width:none;max-height:none;';
  css += `width:${n(100 / crop.width)}%;height:${n(100 / crop.height)}%;`;
  css += `left:${n(-100 * crop.left / crop.width)}%;top:${n(-100 * crop.top / crop.height)}%;`;
  if (crop.angle) {
    // S R S^-1 also handles a document frame whose aspect ratio differs from
    // the source crop. The origin is the crop centre in image coordinates.
    const ratio = (displayWidth / (crop.width * sourceWidth)) /
      (displayHeight / (crop.height * sourceHeight));
    const cos = Math.cos(crop.angle);
    const sin = Math.sin(crop.angle);
    css += `transform-origin:${n(100 * (crop.left + crop.width / 2))}% ${n(100 * (crop.top + crop.height / 2))}%;`;
    css += `transform:matrix(${n(cos)},${n(-sin / ratio)},${n(sin * ratio)},${n(cos)},0,0);`;
  }
  return css;
}

module.exports = { imageCrop, cropImageStyle };
