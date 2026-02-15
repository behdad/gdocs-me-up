/**
 * Image handling functions for Google Docs export
 */

const fs = require('fs');
const path = require('path');
const { escapeHtml, ptToPx } = require('./utils');

/**
 * Fetch image data as base64.
 *
 * @param {string} url - Image URL
 * @param {object} authClient - Auth client for API calls
 * @returns {Promise<string|null>} Base64 encoded image data or null on error
 */
async function fetchAsBase64(url, authClient) {
  try {
    if (!url) {
      throw new Error('No URL provided');
    }
    const resp = await authClient.request({
      url,
      method: 'GET',
      responseType: 'arraybuffer'
    });
    if (!resp.data) {
      throw new Error('No data received');
    }
    return Buffer.from(resp.data, 'binary').toString('base64');
  } catch (error) {
    console.error(`Failed to fetch image from ${url}:`, error.message);
    return null;
  }
}

/**
 * Render an inline image object to HTML.
 *
 * @param {string} objectId - Image object ID
 * @param {object} doc - Full document object
 * @param {object} authClient - Auth client for API calls
 * @param {string} outputDir - Output directory path
 * @param {string} imagesDir - Images directory path
 * @returns {Promise<string>} HTML string for the image
 */
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir) {
  try {
    const inlineObj = doc.inlineObjects?.[objectId];
    if (!inlineObj) return '';

    const embedded = inlineObj.inlineObjectProperties?.embeddedObject;
    if (!embedded?.imageProperties) return '';

    const { imageProperties, size } = embedded;
    const { contentUri, cropProperties } = imageProperties;

    if (!contentUri) {
      console.warn(`Image ${objectId} has no content URI, skipping`);
      return '';
    }

    let scaleX = 1, scaleY = 1;
    let translateX = 0, translateY = 0;
    if (embedded.transform) {
      scaleX = embedded.transform.scaleX || 1;
      scaleY = embedded.transform.scaleY || 1;
      translateX = embedded.transform.translateX || 0;
      translateY = embedded.transform.translateY || 0;
    }

    const base64Data = await fetchAsBase64(contentUri, authClient);
    if (!base64Data) {
      console.warn(`Failed to fetch image ${objectId}`);
      return '';
    }

    const buffer = Buffer.from(base64Data, 'base64');
    const fileName = `image_${objectId}.png`;
    const filePath = path.join(imagesDir, fileName);
    fs.writeFileSync(filePath, buffer);

    const imgSrc = path.relative(outputDir, filePath);

    let style = '';
    if (size?.width?.magnitude && size?.height?.magnitude) {
      const wPx = Math.round(size.width.magnitude * 1.3333 * scaleX);
      const hPx = Math.round(size.height.magnitude * 1.3333 * scaleY);
      style = `max-width:${wPx}px; max-height:${hPx}px;`;
    }

    // Handle cropping - using object-fit and object-position
    if (cropProperties) {
      const { offsetLeft, offsetTop, offsetRight, offsetBottom } = cropProperties;
      if (offsetLeft || offsetTop || offsetRight || offsetBottom) {
        style += `object-fit:cover;`;
        // Calculate the visible portion
        const left = (offsetLeft || 0) * 100;
        const top = (offsetTop || 0) * 100;
        style += `object-position:${-left}% ${-top}%;`;
      }
    }

    // Handle image positioning/translation
    if (translateX !== 0 || translateY !== 0) {
      const txPx = Math.round(translateX * 1.3333);
      const tyPx = Math.round(translateY * 1.3333);
      style += `transform:translate(${txPx}px, ${tyPx}px);`;
    }

    // Image margins
    if (embedded.marginTop?.magnitude) {
      style += `margin-top:${ptToPx(embedded.marginTop.magnitude)}px;`;
    }
    if (embedded.marginBottom?.magnitude) {
      style += `margin-bottom:${ptToPx(embedded.marginBottom.magnitude)}px;`;
    }
    if (embedded.marginLeft?.magnitude) {
      style += `margin-left:${ptToPx(embedded.marginLeft.magnitude)}px;`;
    }
    if (embedded.marginRight?.magnitude) {
      style += `margin-right:${ptToPx(embedded.marginRight.magnitude)}px;`;
    }

    const alt = embedded.title || embedded.description || '';
    return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}" style="${style}" />`;
  } catch (error) {
    console.error(`Error rendering image ${objectId}:`, error.message);
    return `<!-- Image ${objectId} failed to render -->`;
  }
}

module.exports = {
  fetchAsBase64,
  renderInlineObject
};
