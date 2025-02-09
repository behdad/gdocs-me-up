/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Addresses:
 *   - SUBTITLE with its own <h2 class="doc-subtitle">
 *   - Some images at full column width => now enforced with max-width: Xpx, plus container-based scale-down
 *   - Reduced bullet-list spacing
 *   - Subtract doc margins from the final .doc-content width so doc is not too wide
 *   - Ensures lineSpacing is recognized if doc actually sets it in paragraphStyle.lineSpacing
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// -------------- CONFIG --------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images externally

// LTR alignment
const alignmentMapLTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// -------------------------------------
// MAIN EXPORT
// -------------------------------------
async function exportDocToHTML(docId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({
    documentId: docId
  });
  console.log(`Exporting document: ${doc.title}`);

  // Named styles
  const namedStylesMap = buildNamedStylesMap(doc);

  // Section => col info
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle, doc.documentStyle);

  const usedFonts = new Set();
  const htmlLines = [];

  // Basic skeleton
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

  closeAllLists(listStack, htmlLines);

  htmlLines.push('</div>');
  htmlLines.push('</body>');
  htmlLines.push('</html>');

  // Insert Google Fonts
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

// -------------------------------------
// Named Styles
// -------------------------------------
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

// -------------------------------------
// findFirstSectionStyle
// -------------------------------------
function findFirstSectionStyle(doc) {
  const content = doc.body?.content || [];
  for (const c of content) {
    if (c.sectionBreak?.sectionStyle) {
      return c.sectionBreak.sectionStyle;
    }
  }
  return null;
}

/**
 * Extract column info from the first column. Also subtract doc margins
 * from doc.documentStyle if we want final .doc-content narrower.
 */
function extractColumnInfo(sectionStyle, docStyle) {
  if (!sectionStyle) return null;
  const cols = sectionStyle.columnProperties;
  if (!cols || cols.length === 0) return null;

  const first = cols[0];
  const colW = first.width?.magnitude || 0;
  const colP = first.padding?.magnitude || 0;
  let colWidthPx = ptToPx(colW);
  let colPaddingPx = ptToPx(colP);

  // Subtract doc margins if they exist
  // So the final container width is doc's column width minus left/right margin
  if (docStyle) {
    const leftM = docStyle.marginLeft?.magnitude || 72;  // 1in default
    const rightM = docStyle.marginRight?.magnitude || 72;
    // Subtract from column width so we replicate the final text layout
    // This is approximate. If you don't want to subtract margins, remove this.
    colWidthPx -= ptToPx(leftM + rightM);
    if (colWidthPx < 200) colWidthPx = 200; // clamp min
  }

  return {
    colWidthPx,
    colPaddingPx
  };
}

// -------------------------------------
// TOC
// -------------------------------------
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

// -------------------------------------
// Paragraph
// -------------------------------------
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

  // Bullet logic
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
    // If we want a custom style, e.g. <div class="doc-subtitle">, do that
    // We'll do <h2> to show bigger text, plus a .doc-subtitle
    tag = 'h2 class="doc-subtitle"';
  } else if (namedType.startsWith('HEADING_')) {
    const lvl = parseInt(namedType.replace('HEADING_', ''), 10);
    if (lvl >= 1 && lvl <= 6) tag = `h${lvl}`;
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

  // Build innerHtml
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
  // If we used e.g. `tag = 'h2 class="doc-subtitle"'`, we have to handle that carefully:
  // Let's do a small fix to ensure we can still add style:
  const spaceIdx = paragraphHtml.indexOf(' ');
  // e.g. <h2 class="doc-subtitle"...
  // We'll insert style right before the '>' or after the class if no space
  if (spaceIdx < 0) {
    // normal <p>
    if (inlineStyle) paragraphHtml += ` style="${inlineStyle}"`;
  } else {
    // substring in
    let firstPart = paragraphHtml.slice(0, spaceIdx);
    let secondPart = paragraphHtml.slice(spaceIdx);
    if (inlineStyle) secondPart += ` style="${inlineStyle}"`;
    paragraphHtml = firstPart + secondPart;
  }
  paragraphHtml += `>${innerHtml}</${tag.split(' ')[0]}>`; // close e.g. </h2>

  return { html: paragraphHtml, listChange };
}

// -------------------------------------
// Merge text runs
// -------------------------------------
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

// -------------------------------------
// Render a single text run
// -------------------------------------
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

// -------------------------------------
// Render Inline Objects (Images)
// -------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir) {
  const inlineObj = doc.inlineObjects?.[objectId];
  if (!inlineObj) return '';

  const embedded = inlineObj.inlineObjectProperties?.embeddedObject;
  if (!embedded?.imageProperties) return '';

  const { imageProperties } = embedded;
  const { contentUri, size } = imageProperties;

  // fetch image
  const base64Data = await fetchAsBase64(contentUri, authClient);
  const buffer = Buffer.from(base64Data, 'base64');
  const fileName = `image_${objectId}.png`;
  const filePath = path.join(imagesDir, fileName);
  fs.writeFileSync(filePath, buffer);

  const imgSrc = path.relative(outputDir, filePath);

  // use doc size as max
  let style = '';
  if (size?.width?.magnitude && size?.height?.magnitude) {
    const wPx = ptToPx(size.width.magnitude);
    const hPx = ptToPx(size.height.magnitude);
    // We'll do "max-width" so it doesn't exceed doc size, plus container scale-down
    style = `max-width:${wPx}px; max-height:${hPx}px;`;
  }
  const alt = embedded.title || embedded.description || '';

  return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// -------------------------------------
// Lists
// -------------------------------------
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
    if (top.startsWith('u')) {
      htmlLines.push('</ul>');
    } else {
      htmlLines.push('</ol>');
    }
  }
}

// -------------------------------------
// Table
// -------------------------------------
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
      html += '<td style="border: 1px solid #ccc; padding: 0.3em;">';
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

// -------------------------------------
// CSS
// -------------------------------------
function generateGlobalCSS(doc, colInfo) {
  // We reduce spacing for bullet-lists, also add .doc-subtitle
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
  margin: 0.3em 0 0.3em 1.5em; /* reduce spacing */
  padding: 0;
}
li {
  margin: 0.2em 0; /* reduce vertical gap between items */
}
h1, h2, h3, h4, h5, h6 {
  margin: 0.6em 0; /* slightly less than 0.8em */
  font-family: inherit;
}
.doc-subtitle {
  font-size: 1.3em;
  display: block; /* ensure multiline is okay */
  margin: 0.4em 0 0.8em 0;
}
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
  margin: 0.5em 0;
  padding: 0.5em;
  border: 1px dashed #666;
  background: #f9f9f9;
}
.doc-toc h2 {
  margin: 0 0 0.3em 0;
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

  // If doc has pageSize => @page
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

// -------------------------------------
// UTILS
// -------------------------------------
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

