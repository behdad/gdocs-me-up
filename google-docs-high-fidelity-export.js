/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Exports a Google Doc to HTML + CSS with:
 *   - Accurate column width (from section/doc style)
 *   - Exact image sizing
 *   - Headings (H1..H6) with anchor IDs
 *   - Lists (<ul>, <ol>)
 *   - Justified text (text-align: justify)
 *   - Pagination (@page)
 *   - Table of contents (links to headings, in the same tab)
 *   - Google Fonts link
 *   - External images in /images/
 *   - Minimal .htaccess setting DirectoryIndex index.html
 *
 * No footnotes, comments, or other extras.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// --------------------------------------
// CONFIG
// --------------------------------------

const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images in "images/" folder

// For mapping paragraph alignment to CSS
const alignmentMap = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// --------------------------------------
// MAIN EXPORT FUNCTION
// --------------------------------------
async function exportDocToHTML(documentId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & fetch doc
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId });
  console.log(`Exporting document: ${doc.title}`);

  // Determine column/padding from first section
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle);

  // Prepare to build HTML
  const usedFonts = new Set();
  const htmlLines = [];

  // HTML skeleton
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

  // Track lists
  let listStack = [];
  const bodyContent = doc.body && doc.body.content ? doc.body.content : [];

  // Render each element
  for (const element of bodyContent) {
    if (element.sectionBreak) {
      htmlLines.push('<div class="section-break"></div>');
      continue;
    }
    // Table of Contents
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
    // Paragraph
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
    // Table
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
    // other element types if needed
  }

  // Close any open lists
  closeAllLists(listStack, htmlLines);

  // Finish HTML
  htmlLines.push('</div>');
  htmlLines.push('</body>');
  htmlLines.push('</html>');

  // Insert Google Fonts link if needed
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
  fs.writeFileSync(indexPath, htmlLines.join('\n'), 'utf8');
  console.log(`HTML exported to: ${indexPath}`);

  // .htaccess
  const htaccessPath = path.join(outputDir, '.htaccess');
  fs.writeFileSync(htaccessPath, 'DirectoryIndex index.html\n', 'utf8');
  console.log(`.htaccess written to: ${htaccessPath}`);
}

// --------------------------------------
// TOC RENDERING
// --------------------------------------
async function renderTableOfContents(
  toc,
  doc,
  usedFonts,
  authClient,
  outputDir
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
          outputDir
        );
        html += pHtml + '\n';
      }
    }
  }
  html += '</div>\n';
  return html;
}

// --------------------------------------
// PARAGRAPH RENDERING
// --------------------------------------
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

  // Check if list
  let listChange = null;
  if (paragraph.bullet) {
    const listId = paragraph.bullet.listId;
    const nestingLevel = paragraph.bullet.nestingLevel || 0;
    const listDef = doc.lists && doc.lists[listId];
    if (listDef && listDef.listProperties && listDef.listProperties.nestingLevels) {
      const glyph = listDef.listProperties.nestingLevels[nestingLevel];
      const isNumbered = glyph && glyph.glyphType && glyph.glyphType.toLowerCase().includes('number');
      const topOfStack = listStack[listStack.length - 1];
      if (isNumbered && topOfStack !== 'ol') {
        listChange = (topOfStack ? `end${topOfStack.toUpperCase()}` : '') + '|startOL';
      } else if (!isNumbered && topOfStack !== 'ul') {
        listChange = (topOfStack ? `end${topOfStack.toUpperCase()}` : '') + '|startUL';
      }
    }
  } else {
    if (listStack.length > 0) {
      const top = listStack[listStack.length - 1];
      listChange = `end${top.toUpperCase()}`;
    }
  }

  // Heading vs normal
  let tag = 'p';
  let headingIdAttr = '';
  if (namedStyleType.startsWith('HEADING_')) {
    const level = parseInt(namedStyleType.replace('HEADING_', ''), 10);
    if (level >= 1 && level <= 6) {
      tag = 'h' + level;
    }
    // If docs gives a headingId, anchor it
    if (style.headingId) {
      headingIdAttr = ` id="heading-${escapeHtml(style.headingId)}"`;
    }
  }

  // Paragraph style
  let inlineStyle = '';
  if (style.alignment && alignmentMap[style.alignment]) {
    inlineStyle += `text-align: ${alignmentMap[style.alignment]};`;
  }
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

  // Inner HTML
  let innerHtml = '';
  if (paragraph.elements) {
    for (const elem of paragraph.elements) {
      if (elem.textRun) {
        innerHtml += renderTextRun(elem.textRun, usedFonts);
      } else if (elem.inlineObjectElement) {
        const objectId = elem.inlineObjectElement.inlineObjectId;
        innerHtml += await renderInlineObject(
          objectId,
          doc,
          authClient,
          outputDir,
          imagesDir
        );
      }
    }
  }

  let paragraphHtml = `<${tag}${headingIdAttr}`;
  if (inlineStyle) {
    paragraphHtml += ` style="${inlineStyle}"`;
  }
  paragraphHtml += `>${innerHtml}</${tag}>`;

  return { html: paragraphHtml, listChange };
}

// --------------------------------------
// TEXT RUN RENDERING
// --------------------------------------
function renderTextRun(textRun, usedFonts) {
  let { content, textStyle } = textRun;
  if (!content) return '';
  content = content.replace(/\n$/, '');

  const cssClasses = [];
  let inlineStyle = '';

  if (textStyle) {
    if (textStyle.bold) cssClasses.push('bold');
    if (textStyle.italic) cssClasses.push('italic');
    if (textStyle.underline) cssClasses.push('underline');
    if (textStyle.strikethrough) cssClasses.push('strikethrough');

    if (textStyle.baselineOffset === 'SUPERSCRIPT') {
      cssClasses.push('superscript');
    } else if (textStyle.baselineOffset === 'SUBSCRIPT') {
      cssClasses.push('subscript');
    }

    if (textStyle.fontSize && textStyle.fontSize.magnitude) {
      inlineStyle += `font-size: ${textStyle.fontSize.magnitude}pt;`;
    }
    if (textStyle.weightedFontFamily && textStyle.weightedFontFamily.fontFamily) {
      const fam = textStyle.weightedFontFamily.fontFamily;
      usedFonts.add(fam);
      inlineStyle += `font-family: '${fam}', sans-serif;`;
    }
    if (
      textStyle.foregroundColor &&
      textStyle.foregroundColor.color &&
      textStyle.foregroundColor.color.rgbColor
    ) {
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

  // Links: no "target=_blank" now
  if (textStyle && textStyle.link) {
    let linkHref = '';
    if (textStyle.link.headingId) {
      linkHref = `#heading-${escapeHtml(textStyle.link.headingId)}`;
    } else if (textStyle.link.url) {
      linkHref = textStyle.link.url;
    }
    if (linkHref) {
      openTag = `<a href="${escapeHtml(linkHref)}"`;
      // No new window
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

// --------------------------------------
// INLINE OBJECT (IMAGE)
// --------------------------------------
async function renderInlineObject(
  objectId,
  doc,
  authClient,
  outputDir,
  imagesDir
) {
  const inlineObj = doc.inlineObjects[objectId];
  if (!inlineObj) return '';

  const embedded = inlineObj.inlineObjectProperties.embeddedObject;
  if (!embedded || !embedded.imageProperties) return '';

  const { imageProperties } = embedded;
  const { contentUri, size } = imageProperties;

  if (EMBED_IMAGES_AS_BASE64) {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    const dataUrl = `data:image/*;base64,${base64Data}`;
    return buildImageTag(dataUrl, size, embedded);
  } else {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    const buffer = Buffer.from(base64Data, 'base64');
    const imgFileName = `image_${objectId}.png`;
    const imgFilePath = path.join(imagesDir, imgFileName);
    fs.writeFileSync(imgFilePath, buffer);

    const imgSrc = path.relative(outputDir, imgFilePath);
    return buildImageTag(imgSrc, size, embedded);
  }
}

function buildImageTag(src, size, embedded) {
  let style = '';
  if (size && size.width && size.height) {
    const wPx = ptToPx(size.width.magnitude);
    const hPx = ptToPx(size.height.magnitude);
    style = `width:${wPx}px; height:${hPx}px;`;
  }
  const alt = embedded.title || embedded.description || '';
  return `<img src="${escapeHtml(src)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// --------------------------------------
// TABLE RENDERING
// --------------------------------------
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

// --------------------------------------
// LIST STATE
// --------------------------------------
function handleListState(listChange, listStack, htmlOutput) {
  const actions = listChange.split('|');
  actions.forEach(action => {
    if (action.startsWith('start')) {
      const type = action.replace('start', '').toLowerCase();
      if (type === 'ul') {
        htmlOutput.push('<ul>');
        listStack.push('ul');
      } else {
        htmlOutput.push('<ol>');
        listStack.push('ol');
      }
    } else if (action.startsWith('end')) {
      const current = listStack.pop();
      if (current === 'ul') htmlOutput.push('</ul>');
      else htmlOutput.push('</ol>');
    }
  });
}

function closeAllLists(listStack, htmlOutput) {
  while (listStack.length > 0) {
    const current = listStack.pop();
    if (current === 'ul') htmlOutput.push('</ul>');
    else htmlOutput.push('</ol>');
  }
}

// --------------------------------------
// FIND SECTION & COLUMN
// --------------------------------------
function findFirstSectionStyle(doc) {
  const content = doc.body && doc.body.content ? doc.body.content : [];
  for (const c of content) {
    if (c.sectionBreak && c.sectionBreak.sectionStyle) {
      return c.sectionBreak.sectionStyle;
    }
  }
  return null;
}

function extractColumnInfo(sectionStyle) {
  if (!sectionStyle) return null;
  const colProps = sectionStyle.columnProperties;
  if (colProps && colProps.length > 0) {
    const firstCol = colProps[0];
    const colWidthPts = firstCol.width?.magnitude || 0;
    const colPaddingPts = firstCol.padding?.magnitude || 0;
    return {
      colWidthPx: ptToPx(colWidthPts),
      colPaddingPx: ptToPx(colPaddingPts)
    };
  }
  return null;
}

// --------------------------------------
// GLOBAL CSS
// --------------------------------------
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
.bold {
  font-weight: bold;
}
.italic {
  font-style: italic;
}
.underline {
  text-decoration: underline;
}
.strikethrough {
  text-decoration: line-through;
}
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
`);

  if (colInfo && colInfo.colWidthPx) {
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

  if (doc.documentStyle && doc.documentStyle.pageSize) {
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

// --------------------------------------
// GOOGLE FONTS
// --------------------------------------
function buildGoogleFontsLink(fontFamilies) {
  if (!fontFamilies || fontFamilies.length === 0) return '';
  const uniqueFamilies = Array.from(new Set(fontFamilies));
  const familiesParam = uniqueFamilies
    .map(f => f.trim().replace(/\s+/g, '+'))
    .join('&family=');
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}

// --------------------------------------
// UTILS
// --------------------------------------
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

function rgbToHex(r, g, b) {
  const nr = Math.round(r * 255);
  const ng = Math.round(g * 255);
  const nb = Math.round(b * 255);
  return '#' + [nr, ng, nb].map(x => x.toString(16).padStart(2, '0')).join('');
}

function ptToPx(pts) {
  return Math.round(pts * 1.3333);
}

async function fetchAsBase64(url, authClient) {
  const resp = await authClient.request({
    url,
    method: 'GET',
    responseType: 'arraybuffer'
  });
  return Buffer.from(resp.data, 'binary').toString('base64');
}

// --------------------------------------
// CMD-LINE
// --------------------------------------
if (require.main === module) {
  const docId = process.argv[2];
  const outputDir = process.argv[3];
  if (!docId || !outputDir) {
    console.error('Usage: node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>');
    process.exit(1);
  }
  exportDocToHTML(docId, outputDir).catch(err => console.error('Export error:', err));
}

