/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Exports a Google Doc to HTML + CSS with:
 *   - Title/SubTitle => <h1 class="doc-title"> / <h2 class="doc-subtitle">
 *   - Headings (H1..H6) with anchor IDs (for TOC links)
 *   - Right-to-left support for paragraphs & bullet lists (using dir="rtl")
 *   - External images in /images/
 *   - Table of contents linking to headings
 *   - Google Fonts link
 *   - Minimal .htaccess with DirectoryIndex
 *   - Pagination if doc is paginated
 *   - Justification, bold, italic, underline, strikethrough, etc.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// -------- CONFIG ---------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images in separate folder by default

// Basic mapping for LTR alignment
const alignmentMapLTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};
// We'll invert START/END if direction=RTL.

// ---------------------------------------
// Main Export Function
// ---------------------------------------
async function exportDocToHTML(docId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & fetch doc
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId: docId });
  console.log(`Exporting document: ${doc.title}`);

  // Column / section style
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle);

  const usedFonts = new Set();
  const htmlLines = [];

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

  // List stack: track something like ["ul", "ol_rtl", ...]
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
        outputDir
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
        imagesDir
      );
      // If a list starts or ends, handle that
      if (listChange) {
        handleListState(listChange, listStack, htmlLines);
      }
      if (listStack.length > 0) {
        // We are inside a list
        htmlLines.push(`<li>${html}</li>`);
      } else {
        // Normal block
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
        imagesDir
      );
      htmlLines.push(tableHtml);
      continue;
    }
  }

  closeAllLists(listStack, htmlLines);

  // End doc-content, body, html
  htmlLines.push('</div>');
  htmlLines.push('</body>');
  htmlLines.push('</html>');

  // Insert Google Fonts if needed
  const fontLink = buildGoogleFontsLink(Array.from(usedFonts));
  if (fontLink) {
    const insertIndex = htmlLines.findIndex(l => l.includes('</title>'));
    if (insertIndex >= 0) {
      htmlLines.splice(
        insertIndex + 1,
        0,
        `  <link rel="stylesheet" href="${fontLink}">`
      );
    }
  }

  // Write index.html
  const indexPath = path.join(outputDir, 'index.html');
  fs.writeFileSync(indexPath, htmlLines.join('\n'), 'utf-8');
  console.log(`HTML exported to: ${indexPath}`);

  // Write .htaccess
  const htaccessPath = path.join(outputDir, '.htaccess');
  fs.writeFileSync(htaccessPath, 'DirectoryIndex index.html\n', 'utf-8');
  console.log(`.htaccess written to: ${htaccessPath}`);
}

// ---------------------------------------
// TOC
// ---------------------------------------
async function renderTableOfContents(toc, doc, usedFonts, authClient, outputDir) {
  let html = `<div class="doc-toc">\n<h2>Table of Contents</h2>\n`;
  if (toc.content) {
    for (const c of toc.content) {
      if (c.paragraph) {
        // Empty list stack for the TOC
        const { html: pHtml } = await renderParagraph(
          c.paragraph,
          doc,
          usedFonts,
          [],
          authClient,
          outputDir
        );
        html += pHtml + '\n';
      }
    }
  }
  html += '</div>\n';
  return html;
}

// ---------------------------------------
// Paragraph Rendering (with RTL logic)
// ---------------------------------------
async function renderParagraph(
  paragraph,
  doc,
  usedFonts,
  listStack,
  authClient,
  outputDir,
  imagesDir
) {
  const style = paragraph.paragraphStyle || {};
  const namedStyleType = style.namedStyleType || 'NORMAL_TEXT';

  // Check if bullet list
  let listChange = null;
  let isRTL = (style.direction === 'RIGHT_TO_LEFT');

  if (paragraph.bullet) {
    // We detect if we should open/close a UL/OL, possibly with _RTL
    const listId = paragraph.bullet.listId;
    const nestingLevel = paragraph.bullet.nestingLevel || 0;
    const listDef = doc.lists?.[listId];
    if (listDef?.listProperties?.nestingLevels) {
      const glyph = listDef.listProperties.nestingLevels[nestingLevel];
      const isNumbered = glyph?.glyphType?.toLowerCase().includes('number');
      const top = listStack[listStack.length - 1];

      // We decide if we want "startUL" vs. "startUL_RTL" or "startOL" vs. "startOL_RTL"
      const startType = isNumbered ? 'OL' : 'UL';
      const startRTL = isRTL ? `_RTL` : ``;

      // If the top is not matching, we close the old list and open a new one
      if (!top || !top.startsWith(startType.toLowerCase())) {
        // e.g. "endUL|startUL_RTL"
        if (top) {
          listChange = `end${top.toUpperCase()}|start${startType}${startRTL}`;
        } else {
          listChange = `start${startType}${startRTL}`;
        }
      }
    }
  } else {
    // If we had a list open, close it
    if (listStack.length > 0) {
      const top = listStack[listStack.length - 1];
      listChange = `end${top.toUpperCase()}`;
    }
  }

  // TITLE / SUBTITLE / HEADINGS or normal <p>
  let tag = 'p';
  let headingIdAttr = '';
  let classAttr = '';

  if (namedStyleType === 'TITLE') {
    tag = 'h1';
    classAttr = ' class="doc-title"';
  } else if (namedStyleType === 'SUBTITLE') {
    tag = 'h2';
    classAttr = ' class="doc-subtitle"';
  } else if (namedStyleType.startsWith('HEADING_')) {
    const level = parseInt(namedStyleType.replace('HEADING_', ''), 10);
    if (level >= 1 && level <= 6) {
      tag = 'h' + level;
    }
    if (style.headingId) {
      headingIdAttr = ` id="heading-${escapeHtml(style.headingId)}"`;
    }
  }

  // Build inline styles + dir attribute
  let inlineStyle = '';
  let dirAttr = '';
  if (style.direction === 'RIGHT_TO_LEFT') {
    // We'll do dir="rtl"
    dirAttr = ' dir="rtl"';
  }

  // Align
  let align = style.alignment; // e.g. START, END, CENTER, JUSTIFIED
  if (isRTL && (align === 'START' || align === 'END')) {
    // Flip them
    align = (align === 'START') ? 'END' : 'START';
  }
  if (align && alignmentMapLTR[align]) {
    inlineStyle += `text-align: ${alignmentMapLTR[align]};`;
  }

  // lineSpacing, indent, spaceAbove/Below
  if (style.lineSpacing) {
    inlineStyle += `line-height: ${style.lineSpacing / 100};`;
  }
  if (style.indentFirstLine) {
    inlineStyle += `text-indent: ${ptToPx(style.indentFirstLine.magnitude)}px;`;
  } else if (style.indentStart) {
    inlineStyle += `margin-left: ${ptToPx(style.indentStart.magnitude)}px;`;
  }
  if (style.spaceAbove) {
    inlineStyle += `margin-top: ${ptToPx(style.spaceAbove.magnitude)}px;`;
  }
  if (style.spaceBelow) {
    inlineStyle += `margin-bottom: ${ptToPx(style.spaceBelow.magnitude)}px;`;
  }

  // Render elements
  let innerHtml = '';
  if (paragraph.elements) {
    for (const elem of paragraph.elements) {
      if (elem.textRun) {
        innerHtml += renderTextRun(elem.textRun, usedFonts);
      } else if (elem.inlineObjectElement) {
        const objId = elem.inlineObjectElement.inlineObjectId;
        innerHtml += await renderInlineObject(
          objId,
          doc,
          authClient,
          outputDir,
          imagesDir
        );
      }
    }
  }

  // Final paragraph
  let paragraphHtml = `<${tag}${headingIdAttr}${classAttr}${dirAttr}`;
  if (inlineStyle) {
    paragraphHtml += ` style="${inlineStyle}"`;
  }
  paragraphHtml += `>${innerHtml}</${tag}>`;

  return { html: paragraphHtml, listChange };
}

// ---------------------------------------
// Text Run
// ---------------------------------------
function renderTextRun(textRun, usedFonts) {
  let { content, textStyle } = textRun;
  if (!content) return '';
  content = content.replace(/\n$/, '');

  let cssClasses = [];
  let inlineStyle = '';

  if (textStyle) {
    // bold, italic, etc.
    if (textStyle.bold) cssClasses.push('bold');
    if (textStyle.italic) cssClasses.push('italic');
    if (textStyle.underline) cssClasses.push('underline');
    if (textStyle.strikethrough) cssClasses.push('strikethrough');

    if (textStyle.baselineOffset === 'SUPERSCRIPT') {
      cssClasses.push('superscript');
    } else if (textStyle.baselineOffset === 'SUBSCRIPT') {
      cssClasses.push('subscript');
    }

    // fontSize
    if (textStyle.fontSize?.magnitude) {
      inlineStyle += `font-size: ${textStyle.fontSize.magnitude}pt;`;
    }
    // fontFamily
    if (textStyle.weightedFontFamily?.fontFamily) {
      const fam = textStyle.weightedFontFamily.fontFamily;
      usedFonts.add(fam);
      inlineStyle += `font-family: '${fam}', sans-serif;`;
    }
    // color
    if (textStyle.foregroundColor?.color?.rgbColor) {
      const rgb = textStyle.foregroundColor.color.rgbColor;
      const hex = rgbToHex(rgb.red || 0, rgb.green || 0, rgb.blue || 0);
      inlineStyle += `color: ${hex};`;
    }
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

  // link
  if (textStyle?.link) {
    let linkHref = '';
    if (textStyle.link.headingId) {
      linkHref = `#heading-${escapeHtml(textStyle.link.headingId)}`;
    } else if (textStyle.link.url) {
      linkHref = textStyle.link.url;
    }
    if (linkHref) {
      openTag = `<a href="${escapeHtml(linkHref)}"`;
      // no target="_blank"
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
// Inline Objects (Images)
// ---------------------------------------
async function renderInlineObject(objId, doc, authClient, outputDir, imagesDir) {
  const inlineObj = doc.inlineObjects?.[objId];
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
    const fileName = `image_${objId}.png`;
    const filePath = path.join(imagesDir, fileName);
    fs.writeFileSync(filePath, buffer);

    const imgSrc = path.relative(outputDir, filePath);
    return buildImageTag(imgSrc, size, embedded);
  }
}

function buildImageTag(src, size, embedded) {
  let style = '';
  if (size?.width?.magnitude && size?.height?.magnitude) {
    const wPx = ptToPx(size.width.magnitude);
    const hPx = ptToPx(size.height.magnitude);
    style = `width:${wPx}px; height:${hPx}px;`;
  }
  const alt = embedded.title || embedded.description || '';
  return `<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// ---------------------------------------
// Tables
// ---------------------------------------
async function renderTable(
  table,
  doc,
  usedFonts,
  authClient,
  outputDir,
  imagesDir
) {
  let html = '<table class="doc-table" style="border-collapse: collapse; border: 1px solid #ccc;">';
  for (const row of table.tableRows) {
    html += '<tr>';
    for (const cell of row.tableCells) {
      html += '<td style="border: 1px solid #ccc; padding: 0.5em;">';
      for (const c of cell.content) {
        if (c.paragraph) {
          const { html: pHtml } = await renderParagraph(
            c.paragraph,
            doc,
            usedFonts,
            [],
            authClient,
            outputDir,
            imagesDir
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
// List State
// ---------------------------------------
function handleListState(listChange, listStack, htmlLines) {
  // e.g. "startUL_RTL", "endUL", "startOL_RTL", etc.
  const actions = listChange.split('|');
  for (const action of actions) {
    if (action.startsWith('start')) {
      // Could be "startUL", "startUL_RTL", "startOL", "startOL_RTL"
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
      // Could be "endUL", "endUL_RTL", "endOL", "endOL_RTL"
      const top = listStack.pop(); // "ul", "ul_rtl", "ol", "ol_rtl"
      if (top?.startsWith('u')) {
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

// ---------------------------------------
// Doc Section / Column Info
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
  if (cols && cols.length > 0) {
    const first = cols[0];
    const colW = first.width?.magnitude || 0;
    const colPad = first.padding?.magnitude || 0;
    return {
      colWidthPx: ptToPx(colW),
      colPaddingPx: ptToPx(colPad)
    };
  }
  return null;
}

// ---------------------------------------
// Global CSS
// ---------------------------------------
function generateGlobalCSS(doc, colInfo) {
  let lines = [];
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
h1, h2, h3, h4, h5, h6 {
  margin: 0.8em 0;
  font-family: inherit;
}
ul, ol {
  margin: 0.5em 0 0.5em 2em;
  padding: 0;
}
img {
  display: inline-block;
  max-width: 100%;
}
/* Basic classes for style runs */
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
/* Title / Subtitle styling */
.doc-title {
  font-size: 2.2em;
  font-weight: bold;
  margin: 0.6em 0 0.3em 0;
}
.doc-subtitle {
  font-size: 1.4em;
  font-weight: 400;
  color: #444;
  margin: 0 0 0.8em 0;
}
`);

  // If col width info
  if (colInfo?.colWidthPx) {
    let pad = colInfo.colPaddingPx || 0;
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

  // Paginated => @page
  if (doc.documentStyle?.pageSize) {
    const { width, height } = doc.documentStyle.pageSize;
    if (width && height) {
      const wIn = (width.magnitude || 612) / 72;
      const hIn = (height.magnitude || 792) / 72;
      const topMargin = doc.documentStyle.marginTop
        ? doc.documentStyle.marginTop.magnitude / 72
        : 1;
      const rightMargin = doc.documentStyle.marginRight
        ? doc.documentStyle.marginRight.magnitude / 72
        : 1;
      const bottomMargin = doc.documentStyle.marginBottom
        ? doc.documentStyle.marginBottom.magnitude / 72
        : 1;
      const leftMargin = doc.documentStyle.marginLeft
        ? doc.documentStyle.marginLeft.magnitude / 72
        : 1;

      lines.push(`
@page {
  size: ${wIn}in ${hIn}in;
  margin: ${topMargin}in ${rightMargin}in ${bottomMargin}in ${leftMargin}in;
}
      `);
    }
  }

  return lines.join('\n');
}

// ---------------------------------------
// Google Fonts
// ---------------------------------------
function buildGoogleFontsLink(fontFamilies) {
  if (!fontFamilies || fontFamilies.length === 0) return '';
  const unique = Array.from(new Set(fontFamilies));
  const familiesParam = unique.map(f => f.trim().replace(/\s+/g, '+')).join('&family=');
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}

// ---------------------------------------
// Utils
// ---------------------------------------
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

async function fetchAsBase64(url, authClient) {
  const resp = await authClient.request({
    url,
    method: 'GET',
    responseType: 'arraybuffer'
  });
  return Buffer.from(resp.data, 'binary').toString('base64');
}

// ---------------------------------------
// CLI Entry
// ---------------------------------------
if (require.main === module) {
  const docId = process.argv[2];
  const outputDir = process.argv[3];
  if (!docId || !outputDir) {
    console.error('Usage: node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>');
    process.exit(1);
  }
  exportDocToHTML(docId, outputDir).catch(err => console.error('Export error:', err));
}

