/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Restores the behavior that images never exceed their doc size nor the layout width.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ------------------ CONFIG ------------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images in "images/" folder instead of base64

// Alignments for LTR
const alignmentMapLTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// ---------------------------------------
// MAIN EXPORT
// ---------------------------------------
async function exportDocToHTML(docId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & docs
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId: docId });
  console.log(`Exporting document: ${doc.title}`);

  // Named styles
  const namedStylesMap = buildNamedStylesMap(doc);

  // Section style => col info
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle);

  const usedFonts = new Set();
  let htmlLines = [];

  // Basic HTML
  htmlLines.push('<!DOCTYPE html>');
  htmlLines.push('<html lang="en">');
  htmlLines.push('<head>');
  htmlLines.push('  <meta charset="UTF-8">');
  htmlLines.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlLines.push('  <style>');
  htmlLines.push(generateGlobalCSS(doc, colInfo));
  htmlLines.push('  </style>');
  htmlLines.push('</head>');
  htmlLines.push('<body>');
  htmlLines.push('<div class="doc-content">');

  let listStack = [];
  const bodyContent = doc.body?.content || [];

  for (const element of bodyContent) {
    if (element.sectionBreak) {
      htmlLines.push('<div class="section-break"></div>');
      continue;
    }
    if (element.tableOfContents) {
      closeAllLists(listStack, htmlLines);
      const tocHtml = await renderTableOfContents(
        element.tableOfContents,
        doc,
        usedFonts,
        authClient,
        outputDir,
        namedStylesMap
      );
      htmlLines.push(tocHtml);
      continue;
    }
    if (element.paragraph) {
      const { html, listChange } = await renderParagraph(
        element.paragraph,
        doc,
        usedFonts,
        listStack,
        authClient,
        outputDir,
        imagesDir,
        namedStylesMap
      );
      if (listChange) {
        handleListState(listChange, listStack, htmlLines);
      }
      if (listStack.length > 0) {
        htmlLines.push(`<li>${html}</li>`);
      } else {
        htmlLines.push(html);
      }
      continue;
    }
    if (element.table) {
      closeAllLists(listStack, htmlLines);
      const tableHtml = await renderTable(
        element.table,
        doc,
        usedFonts,
        authClient,
        outputDir,
        imagesDir,
        namedStylesMap
      );
      htmlLines.push(tableHtml);
      continue;
    }
  }

  // close lists
  closeAllLists(listStack, htmlLines);

  htmlLines.push('</div>');
  htmlLines.push('</body>');
  htmlLines.push('</html>');

  // Insert Google Fonts link if needed
  const fontLink = buildGoogleFontsLink(Array.from(usedFonts));
  if (fontLink) {
    const idx = htmlLines.findIndex(l => l.includes('</title>'));
    if (idx >= 0) {
      htmlLines.splice(idx + 1, 0, `  <link rel="stylesheet" href="${fontLink}">`);
    }
  }

  // Write index.html
  const indexPath = path.join(outputDir, 'index.html');
  fs.writeFileSync(indexPath, htmlLines.join('\n'), 'utf8');
  console.log(`HTML exported to: ${indexPath}`);

  // Write .htaccess
  const htaccessPath = path.join(outputDir, '.htaccess');
  fs.writeFileSync(htaccessPath, 'DirectoryIndex index.html\n', 'utf8');
  console.log(`.htaccess written to: ${htaccessPath}`);
}

// ---------------------------------------
// Named Styles
// ---------------------------------------
function buildNamedStylesMap(doc) {
  const map = {};
  const named = doc.namedStyles?.styles || [];
  for (const s of named) {
    map[s.namedStyleType] = {
      paragraphStyle: s.paragraphStyle || {},
      textStyle: s.textStyle || {}
    };
  }
  return map;
}

// ---------------------------------------
// findFirstSectionStyle & extractColumnInfo
// ---------------------------------------
function findFirstSectionStyle(doc) {
  const content = doc.body?.content || [];
  for (const c of content) {
    if (c.sectionBreak?.sectionStyle) {
      return c.sectionBreak.sectionStyle;
    }
  }
  return null;
}
function extractColumnInfo(sectionStyle) {
  if (!sectionStyle) return null;
  const cols = sectionStyle.columnProperties;
  if (!cols || cols.length === 0) return null;
  const first = cols[0];
  const colW = first.width?.magnitude || 0;
  const colP = first.padding?.magnitude || 0;
  return {
    colWidthPx: ptToPx(colW),
    colPaddingPx: ptToPx(colP)
  };
}

// ---------------------------------------
// TOC
// ---------------------------------------
async function renderTableOfContents(
  toc,
  doc,
  usedFonts,
  authClient,
  outputDir,
  namedStylesMap
) {
  let html = '<div class="doc-toc">\n<h2>Table of Contents</h2>\n';
  if (toc.content) {
    for (const c of toc.content) {
      if (c.paragraph) {
        const { html: pHtml } = await renderParagraph(
          c.paragraph,
          doc,
          usedFonts,
          [],
          authClient,
          outputDir,
          null,
          namedStylesMap
        );
        html += pHtml + '\n';
      }
    }
  }
  html += '</div>\n';
  return html;
}

// ---------------------------------------
// Paragraph
// ---------------------------------------
async function renderParagraph(
  paragraph,
  doc,
  usedFonts,
  listStack,
  authClient,
  outputDir,
  imagesDir,
  namedStylesMap
) {
  const style = paragraph.paragraphStyle || {};
  const namedType = style.namedStyleType || 'NORMAL_TEXT';

  // Merge doc-level style
  let mergedParaStyle = {};
  let mergedTextStyle = {};
  if (namedStylesMap[namedType]) {
    mergedParaStyle = deepCopy(namedStylesMap[namedType].paragraphStyle);
    mergedTextStyle = deepCopy(namedStylesMap[namedType].textStyle);
  }
  deepMerge(mergedParaStyle, style);

  const isRTL = (mergedParaStyle.direction === 'RIGHT_TO_LEFT');

  let listChange = null;
  if (paragraph.bullet) {
    listChange = detectListChange(paragraph.bullet, doc, listStack, isRTL);
  } else {
    if (listStack.length > 0) {
      const top = listStack[listStack.length - 1];
      listChange = `end${top.toUpperCase()}`;
    }
  }

  // Tag
  let tag = 'p';
  if (namedType === 'TITLE') {
    tag = 'h1';
  } else if (namedType === 'SUBTITLE') {
    tag = 'h2';
  } else if (namedType.startsWith('HEADING_')) {
    const lvl = parseInt(namedType.replace('HEADING_', ''), 10);
    if (lvl >= 1 && lvl <= 6) tag = 'h' + lvl;
  }
  let headingIdAttr = '';
  if (style.headingId) {
    headingIdAttr = ` id="heading-${escapeHtml(style.headingId)}"`;
  }

  // inline style + dir
  let dirAttr = '';
  if (mergedParaStyle.direction === 'RIGHT_TO_LEFT') {
    dirAttr = ' dir="rtl"';
  }
  // Flip START/END if RTL
  let align = mergedParaStyle.alignment;
  if (isRTL && (align === 'START' || align === 'END')) {
    align = (align === 'START') ? 'END' : 'START';
  }
  let inlineStyle = '';
  if (align && alignmentMapLTR[align]) {
    inlineStyle += `text-align: ${alignmentMapLTR[align]};`;
  }
  if (mergedParaStyle.lineSpacing) {
    inlineStyle += `line-height: ${mergedParaStyle.lineSpacing / 100};`;
  }
  if (mergedParaStyle.indentFirstLine) {
    inlineStyle += `text-indent: ${ptToPx(mergedParaStyle.indentFirstLine.magnitude)}px;`;
  } else if (mergedParaStyle.indentStart) {
    inlineStyle += `margin-left: ${ptToPx(mergedParaStyle.indentStart.magnitude)}px;`;
  }
  if (mergedParaStyle.spaceAbove) {
    inlineStyle += `margin-top: ${ptToPx(mergedParaStyle.spaceAbove.magnitude)}px;`;
  }
  if (mergedParaStyle.spaceBelow) {
    inlineStyle += `margin-bottom: ${ptToPx(mergedParaStyle.spaceBelow.magnitude)}px;`;
  }

  // Merge text runs
  const mergedRuns = mergeTextRuns(paragraph.elements || []);

  let innerHtml = '';
  for (const r of mergedRuns) {
    if (r.inlineObjectElement) {
      const objId = r.inlineObjectElement.inlineObjectId;
      innerHtml += await renderInlineObject(
        objId,
        doc,
        authClient,
        outputDir,
        imagesDir
      );
    } else if (r.textRun) {
      innerHtml += renderTextRun(r.textRun, usedFonts, mergedTextStyle);
    }
  }

  let paragraphHtml = `<${tag}${headingIdAttr}${dirAttr}`;
  if (inlineStyle) paragraphHtml += ` style="${inlineStyle}"`;
  paragraphHtml += `>${innerHtml}</${tag}>`;

  return { html: paragraphHtml, listChange };
}

// ---------------------------------------
// Merge Text Runs
// ---------------------------------------
function mergeTextRuns(elements) {
  const merged = [];
  let last = null;

  for (const e of elements) {
    if (e.inlineObjectElement) {
      merged.push({ inlineObjectElement: e.inlineObjectElement });
      last = null;
    } else if (e.textRun) {
      const style = e.textRun.textStyle || {};
      const content = e.textRun.content || '';
      if (last && last.textRun && isSameTextStyle(last.textRun.textStyle, style)) {
        last.textRun.content += content;
      } else {
        merged.push({ textRun: { content, textStyle: deepCopy(style) } });
        last = merged[merged.length - 1];
      }
    }
  }
  return merged;
}

function isSameTextStyle(a, b) {
  const fields = [
    'bold', 'italic', 'underline', 'strikethrough',
    'baselineOffset', 'fontSize', 'weightedFontFamily',
    'foregroundColor', 'link'
  ];
  for (const f of fields) {
    if (JSON.stringify(a[f] || null) !== JSON.stringify(b[f] || null)) {
      return false;
    }
  }
  return true;
}

// ---------------------------------------
// Render a single text run
// ---------------------------------------
function renderTextRun(textRun, usedFonts, baseStyle) {
  const finalStyle = deepCopy(baseStyle || {});
  deepMerge(finalStyle, textRun.textStyle || {});

  let content = textRun.content || '';
  content = content.replace(/\n$/, '');

  const cssClasses = [];
  let inlineStyle = '';

  if (finalStyle.bold) cssClasses.push('bold');
  if (finalStyle.italic) cssClasses.push('italic');
  if (finalStyle.underline) cssClasses.push('underline');
  if (finalStyle.strikethrough) cssClasses.push('strikethrough');

  if (finalStyle.baselineOffset === 'SUPERSCRIPT') {
    cssClasses.push('superscript');
  } else if (finalStyle.baselineOffset === 'SUBSCRIPT') {
    cssClasses.push('subscript');
  }
  if (finalStyle.fontSize?.magnitude) {
    inlineStyle += `font-size: ${finalStyle.fontSize.magnitude}pt;`;
  }
  if (finalStyle.weightedFontFamily?.fontFamily) {
    const fam = finalStyle.weightedFontFamily.fontFamily;
    usedFonts.add(fam);
    inlineStyle += `font-family: '${fam}', sans-serif;`;
  }
  if (finalStyle.foregroundColor?.color?.rgbColor) {
    const rgb = finalStyle.foregroundColor.color.rgbColor;
    const hex = rgbToHex(rgb.red || 0, rgb.green || 0, rgb.blue || 0);
    inlineStyle += `color: ${hex};`;
  }

  let openTag = '<span';
  if (cssClasses.length > 0) {
    openTag += ` class="${cssClasses.join(' ')}"`;
  }
  if (inlineStyle) {
    openTag += ` style="${inlineStyle}"`;
  }
  openTag += '>';
  let closeTag = '</span>';

  // Link?
  if (finalStyle.link) {
    let linkHref = '';
    if (finalStyle.link.headingId) {
      linkHref = `#heading-${escapeHtml(finalStyle.link.headingId)}`;
    } else if (finalStyle.link.url) {
      linkHref = finalStyle.link.url;
    }
    if (linkHref) {
      openTag = `<a href="${escapeHtml(linkHref)}"`;
      if (cssClasses.length > 0) {
        openTag += ` class="${cssClasses.join(' ')}"`;
      }
      if (inlineStyle) {
        openTag += ` style="${inlineStyle}"`;
      }
      openTag += '>';
      closeTag = '</a>';
    }
  }

  return openTag + escapeHtml(content) + closeTag;
}

// ---------------------------------------
// Render Inline Object (Images)
//   -> Now we do "max-width: Xpx; max-height: Ypx;"
// ---------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir) {
  const inlineObj = doc.inlineObjects?.[objectId];
  if (!inlineObj) return '';

  const embedded = inlineObj.inlineObjectProperties?.embeddedObject;
  if (!embedded?.imageProperties) return '';

  const { imageProperties } = embedded;
  const { contentUri, size } = imageProperties;

  if (EMBED_IMAGES_AS_BASE64) {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    const dataUrl = `data:image/*;base64,${base64Data}`;
    return buildImageTag(dataUrl, size, embedded);
  } else {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    const buffer = Buffer.from(base64Data, 'base64');
    const fileName = `image_${objectId}.png`;
    const filePath = path.join(imagesDir, fileName);
    fs.writeFileSync(filePath, buffer);

    const imgSrc = path.relative(outputDir, filePath);
    return buildImageTag(imgSrc, size, embedded);
  }
}

/**
 * Build an <img> that won't exceed doc's reported size (max-width / max-height)
 * nor exceed container (img { max-width: 100%; height: auto; }).
 */
function buildImageTag(src, size, embedded) {
  let style = '';
  if (size?.width?.magnitude && size?.height?.magnitude) {
    const wPx = Math.round(size.width.magnitude * 1.3333);
    const hPx = Math.round(size.height.magnitude * 1.3333);

    // We use max-width & max-height so it can't exceed doc size,
    // but also won't overflow container's width due to CSS: img {max-width:100%;}
    style = `max-width:${wPx}px; max-height:${hPx}px;`;
  }
  const alt = embedded.title || embedded.description || '';
  return `<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// ---------------------------------------
// Lists
// ---------------------------------------
function detectListChange(bullet, doc, listStack, isRTL) {
  const listId = bullet.listId;
  const nestingLevel = bullet.nestingLevel || 0;
  const listDef = doc.lists?.[listId];
  if (!listDef?.listProperties?.nestingLevels) return null;

  const glyph = listDef.listProperties.nestingLevels[nestingLevel];
  const isNumbered = glyph?.glyphType?.toLowerCase().includes('number');
  const top = listStack[listStack.length - 1];

  const startType = isNumbered ? 'OL' : 'UL';
  const rtlFlag = isRTL ? '_RTL' : '';

  if (!top || !top.startsWith(startType.toLowerCase())) {
    if (top) {
      return `end${top.toUpperCase()}|start${startType}${rtlFlag}`;
    } else {
      return `start${startType}${rtlFlag}`;
    }
  }
  return null;
}

function handleListState(listChange, listStack, htmlLines) {
  const actions = listChange.split('|');
  for (const action of actions) {
    if (action.startsWith('start')) {
      if (action.includes('UL_RTL')) {
        htmlLines.push('<ul dir="rtl">');
        listStack.push('ul_rtl');
      } else if (action.includes('OL_RTL')) {
        htmlLines.push('<ol dir="rtl">');
        listStack.push('ol_rtl');
      } else if (action.includes('UL')) {
        htmlLines.push('<ul>');
        listStack.push('ul');
      } else {
        htmlLines.push('<ol>');
        listStack.push('ol');
      }
    } else if (action.startsWith('end')) {
      const top = listStack.pop();
      if (top.startsWith('u')) {
        htmlLines.push('</ul>');
      } else {
        htmlLines.push('</ol>');
      }
    }
  }
}

function closeAllLists(listStack, htmlLines) {
  while (listStack.length > 0) {
    const top = listStack.pop();
    if (top.startsWith('u')) htmlLines.push('</ul>');
    else htmlLines.push('</ol>');
  }
}

// ---------------------------------------
// Table
// ---------------------------------------
async function renderTable(
  table,
  doc,
  usedFonts,
  authClient,
  outputDir,
  imagesDir,
  namedStylesMap
) {
  let html = '<table class="doc-table" style="border-collapse: collapse; border: 1px solid #ccc;">';
  for (const row of table.tableRows || []) {
    html += '<tr>';
    for (const cell of row.tableCells || []) {
      html += '<td style="border: 1px solid #ccc; padding: 0.5em;">';
      for (const c of cell.content || []) {
        if (c.paragraph) {
          const { html: pHtml } = await renderParagraph(
            c.paragraph,
            doc,
            usedFonts,
            [],
            authClient,
            outputDir,
            imagesDir,
            namedStylesMap
          );
          html += pHtml;
        }
      }
      html += '</td>';
    }
    html += '</tr>';
  }
  html += '</table>';
  return html;
}

// ---------------------------------------
// GLOBAL CSS
//   Reintroduce img { max-width: 100%; height: auto; }
//   to ensure images won't overflow container
// ---------------------------------------
function generateGlobalCSS(doc, colInfo) {
  const lines = [];
  lines.push(`
body {
  margin: 0;
  font-family: sans-serif;
  line-height: 1.5;
}
.doc-content {
  margin: 1em auto;
}
p, li {
  margin: 0.5em 0;
}
ul, ol {
  margin: 0.5em 0 0.5em 2em;
  padding: 0;
}
h1, h2, h3, h4, h5, h6 {
  margin: 0.8em 0;
  font-family: inherit;
}
/* CRUCIAL: allow images to scale down if container is narrower than doc size */
img {
  display: inline-block;
  max-width: 100%;
  height: auto;
}
.bold { font-weight: bold; }
.italic { font-style: italic; }
.underline { text-decoration: underline; }
.strikethrough { text-decoration: line-through; }
.superscript {
  vertical-align: super;
  font-size: 0.8em;
}
.subscript {
  vertical-align: sub;
  font-size: 0.8em;
}
.section-break {
  page-break-before: always;
}
.doc-toc {
  margin: 1em 0;
  padding: 1em;
  border: 1px dashed #666;
  background: #f9f9f9;
}
.doc-toc h2 {
  margin: 0 0 0.5em 0;
  font-size: 1.2em;
  font-weight: bold;
}
.doc-table {
  border-collapse: collapse;
  margin: 0.5em 0;
}
`);

  if (colInfo?.colWidthPx) {
    const pad = colInfo.colPaddingPx || 0;
    lines.push(`
.doc-content {
  width: ${colInfo.colWidthPx}px;
  padding-left: ${pad}px;
  padding-right: ${pad}px;
}
    `);
  } else {
    lines.push(`
.doc-content {
  max-width: 800px;
  padding: 0 20px;
}
    `);
  }

  if (doc.documentStyle?.pageSize) {
    const { width, height } = doc.documentStyle.pageSize;
    if (width && height) {
      const wIn = (width.magnitude || 612) / 72;
      const hIn = (height.magnitude || 792) / 72;
      const topM = doc.documentStyle.marginTop
        ? doc.documentStyle.marginTop.magnitude / 72
        : 1;
      const rightM = doc.documentStyle.marginRight
        ? doc.documentStyle.marginRight.magnitude / 72
        : 1;
      const botM = doc.documentStyle.marginBottom
        ? doc.documentStyle.marginBottom.magnitude / 72
        : 1;
      const leftM = doc.documentStyle.marginLeft
        ? doc.documentStyle.marginLeft.magnitude / 72
        : 1;

      lines.push(`
@page {
  size: ${wIn}in ${hIn}in;
  margin: ${topM}in ${rightM}in ${botM}in ${leftM}in;
}
      `);
    }
  }
  return lines.join('\n');
}

// ---------------------------------------
// MISC UTILS
// ---------------------------------------
function deepMerge(base, overlay) {
  for (const k in overlay) {
    if (
      typeof overlay[k] === 'object' &&
      overlay[k] !== null &&
      !Array.isArray(overlay[k])
    ) {
      if (!base[k]) base[k] = {};
      deepMerge(base[k], overlay[k]);
    } else {
      base[k] = overlay[k];
    }
  }
}
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}
function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function ptToPx(pts) {
  return Math.round(pts * 1.3333);
}
function rgbToHex(r, g, b) {
  const nr = Math.round(r * 255);
  const ng = Math.round(g * 255);
  const nb = Math.round(b * 255);
  return '#' + [nr, ng, nb].map(x => x.toString(16).padStart(2, '0')).join('');
}
function buildGoogleFontsLink(fontFamilies) {
  if (!fontFamilies || fontFamilies.length === 0) return '';
  const unique = Array.from(new Set(fontFamilies));
  const familiesParam = unique.map(f => f.trim().replace(/\s+/g, '+')).join('&family=');
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}

async function getAuthClient() {
  const auth = new google.auth.GoogleAuth({
    keyFile: SERVICE_ACCOUNT_KEY_FILE,
    scopes: [
      'https://www.googleapis.com/auth/documents.readonly',
      'https://www.googleapis.com/auth/drive.readonly'
    ]
  });
  return auth.getClient();
}

async function fetchAsBase64(url, authClient) {
  const resp = await authClient.request({
    url,
    method: 'GET',
    responseType: 'arraybuffer'
  });
  return Buffer.from(resp.data, 'binary').toString('base64');
}

// CLI
if (require.main === module) {
  const docId = process.argv[2];
  const outDir = process.argv[3];
  if (!docId || !outDir) {
    console.error('Usage: node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>');
    process.exit(1);
  }
  exportDocToHTML(docId, outDir).catch(err => console.error('Export error:', err));
}

