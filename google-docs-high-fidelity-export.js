/**
 * google-docs-high-fidelity-export.js
 *
 * Exports a Google Doc to HTML + CSS with:
 *   - Accurate column width (from section or doc style)
 *   - Exact image sizing (no extra scaling)
 *   - Headings (H1..H6)
 *   - Lists (<ul>, <ol>)
 *   - Pagination (via @page if doc is paginated)
 *   - Fonts (Google Fonts link)
 *   - Text attributes (bold, italic, underline, color, size)
 *   - Tables
 *
 * No footnotes, comments, or other extras.
 */

const fs = require('fs');
const { google } = require('googleapis');

// ------------ Configuration ------------
const OUTPUT_HTML = 'index.html';   // Where to save the resulting HTML

// Toggle embedding images vs. saving locally:
const EMBED_IMAGES_AS_BASE64 = false;
const IMAGE_DIR = './images'; // used if EMBED_IMAGES_AS_BASE64 = false

// Service account or OAuth credentials (JSON file):
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';

// ---------------------------------------
// Main Export Function
// ---------------------------------------
async function exportDocToHTML(documentId) {
  // Auth & Setup
  const authClient = await getAuthClient();

  // Get the document structure
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId });
  console.log(`Exporting document: ${doc.title}`);

  // Extract column and padding info (single column).
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle);

  // Prepare sets, arrays, and accumulators
  const usedFonts = new Set();     // to build a single Google Fonts link
  let htmlOutput = [];

  // Basic HTML skeleton
  htmlOutput.push('<!DOCTYPE html>');
  htmlOutput.push('<html lang="en">');
  htmlOutput.push('<head>');
  htmlOutput.push('  <meta charset="UTF-8">');
  htmlOutput.push(`  <title>${escapeHtml(doc.title)}</title>`);
  // We’ll insert Google Fonts link later if any fonts are used
  htmlOutput.push('  <style>');
  htmlOutput.push(generateGlobalCSS(doc, colInfo));
  htmlOutput.push('  </style>');
  htmlOutput.push('</head>');
  htmlOutput.push('<body>');

  // Wrap all content in a container matching the doc’s column width
  htmlOutput.push('<div class="doc-content">');

  // We’ll track current list nesting so we know when to open/close <ul>/<ol>
  let listStack = [];

  // Traverse each structural element in doc body
  const bodyContent = doc.body && doc.body.content ? doc.body.content : [];
  for (const element of bodyContent) {
    // Section breaks (page or column breaks in paginated docs)
    if (element.sectionBreak) {
      // Insert a visual marker or page-break
      htmlOutput.push('<div class="section-break"></div>');
      continue;
    }

    if (element.paragraph) {
      // Render paragraph (or heading)
      const { html, listChange } = await renderParagraph(
        element.paragraph,
        doc,
        usedFonts,
        listStack,
        authClient
      );

      // Handle list start/end logic
      if (listChange) {
        handleListState(listChange, listStack, htmlOutput);
      }

      // If inside a list, wrap paragraph HTML in <li>; otherwise just output the paragraph
      if (listStack.length > 0) {
        htmlOutput.push(`<li>${html}</li>`);
      } else {
        htmlOutput.push(html);
      }

    } else if (element.table) {
      // Render a table
      const tableHtml = await renderTable(element.table, doc, usedFonts, authClient);
      // Ensure any open lists are closed before rendering a table
      closeAllLists(listStack, htmlOutput);
      htmlOutput.push(tableHtml);

    } else {
      // Other element types can be handled if needed
    }
  }

  // Close any open lists
  closeAllLists(listStack, htmlOutput);

  // Close .doc-content div
  htmlOutput.push('</div>');

  // Close body & html
  htmlOutput.push('</body>');
  htmlOutput.push('</html>');

  // Insert <link> for Google Fonts if needed
  const fontLink = buildGoogleFontsLink(Array.from(usedFonts));
  if (fontLink) {
    // Insert after the </title> or near the top
    const insertIndex = htmlOutput.findIndex(l => l.includes('</title>'));
    if (insertIndex >= 0) {
      htmlOutput.splice(
        insertIndex + 1,
        0,
        `  <link rel="stylesheet" href="${fontLink}">`
      );
    }
  }

  // Write to file
  fs.writeFileSync(OUTPUT_HTML, htmlOutput.join('\n'), 'utf8');
  console.log(`HTML exported to: ${OUTPUT_HTML}`);
}

// ---------------------------------------
// Section / Column Info
// ---------------------------------------
function findFirstSectionStyle(doc) {
  const content = doc.body && doc.body.content ? doc.body.content : [];
  for (const c of content) {
    if (c.sectionBreak && c.sectionBreak.sectionStyle) {
      return c.sectionBreak.sectionStyle;
    }
  }
  // Fallback to doc.documentStyle if no explicit section
  return null;
}

function extractColumnInfo(sectionStyle) {
  if (!sectionStyle) {
    return null;
  }
  const colProps = sectionStyle.columnProperties;
  if (colProps && colProps.length > 0) {
    const firstCol = colProps[0];
    const colWidthPts = (firstCol.width && firstCol.width.magnitude) || 0;
    const colPaddingPts = (firstCol.padding && firstCol.padding.magnitude) || 0;
    return {
      colWidthPx: ptToPx(colWidthPts),
      colPaddingPx: ptToPx(colPaddingPts),
    };
  }
  return null;
}

// ---------------------------------------
// Paragraph Rendering
// ---------------------------------------
async function renderParagraph(paragraph, doc, usedFonts, listStack, authClient) {
  const style = paragraph.paragraphStyle || {};
  const namedStyleType = style.namedStyleType || 'NORMAL_TEXT';

  // Check if this paragraph belongs to a list
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

  let tag = 'p';
  if (namedStyleType.startsWith('HEADING_')) {
    const levelStr = namedStyleType.replace('HEADING_', '');
    const level = parseInt(levelStr, 10);
    if (level >= 1 && level <= 6) {
      tag = 'h' + level;
    }
  }

  let inlineStyle = '';
  if (style.alignment) {
    inlineStyle += `text-align: ${style.alignment.toLowerCase()};`;
  }
  if (style.lineSpacing) {
    const lineHeight = style.lineSpacing / 100;
    inlineStyle += `line-height: ${lineHeight};`;
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

  let innerHtml = '';
  for (const elem of paragraph.elements) {
    if (elem.textRun) {
      innerHtml += renderTextRun(elem.textRun, usedFonts);
    } else if (elem.inlineObjectElement) {
      const objectId = elem.inlineObjectElement.inlineObjectId;
      innerHtml += await renderInlineObject(objectId, doc, authClient);
    }
  }

  let paragraphHtml = `<${tag}`;
  if (inlineStyle) {
    paragraphHtml += ` style="${inlineStyle}"`;
  }
  paragraphHtml += `>${innerHtml}</${tag}>`;

  return { html: paragraphHtml, listChange };
}

// ---------------------------------------
// Text Run Rendering
// ---------------------------------------
function renderTextRun(textRun, usedFonts) {
  let { content, textStyle } = textRun;
  if (!content) return '';
  content = content.replace(/\n$/, '');

  let cssClasses = [];
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

  if (textStyle && textStyle.link && textStyle.link.url) {
    openTag = `<a href="${escapeHtml(textStyle.link.url)}" target="_blank"`;
    if (cssClasses.length > 0) {
      openTag += ` class="${cssClasses.join(' ')}"`;
    }
    if (inlineStyle) {
      openTag += ` style="${inlineStyle}"`;
    }
    openTag += '>';
    closeTag = '</a>';
  }

  return openTag + escapeHtml(content) + closeTag;
}

// ---------------------------------------
// Inline Objects (Images)
// ---------------------------------------
async function renderInlineObject(objectId, doc, authClient) {
  const inlineObj = doc.inlineObjects[objectId];
  if (!inlineObj) return '';

  const embedded = inlineObj.inlineObjectProperties.embeddedObject;
  if (!embedded || !embedded.imageProperties) return '';

  const { imageProperties } = embedded;
  let { contentUri, size } = imageProperties;

  let imgSrc = '';
  if (EMBED_IMAGES_AS_BASE64) {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    imgSrc = `data:image/*;base64,${base64Data}`;
  } else {
    const base64Data = await fetchAsBase64(contentUri, authClient);
    const buffer = Buffer.from(base64Data, 'base64');
    if (!fs.existsSync(IMAGE_DIR)) {
      fs.mkdirSync(IMAGE_DIR, { recursive: true });
    }
    const fileName = `image_${objectId}.png`;
    fs.writeFileSync(`${IMAGE_DIR}/${fileName}`, buffer);
    imgSrc = `${IMAGE_DIR.replace('./', '')}/${fileName}`;
  }

  let style = '';
  if (size && size.width && size.height) {
    const wPx = ptToPx(size.width.magnitude);
    const hPx = ptToPx(size.height.magnitude);
    style = `width:${wPx}px; height:${hPx}px;`;
  }

  const alt = embedded.title || embedded.description || '';

  return `<img src="${imgSrc}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// ---------------------------------------
// Table Rendering (Basic)
// ---------------------------------------
async function renderTable(table, doc, usedFonts, authClient) {
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
            authClient
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
// Helper: Manage List State
// ---------------------------------------
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
      const type = action.replace('end', '').toLowerCase();
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

// ---------------------------------------
// Helper: Build Global CSS
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
img {
  max-width: 100%;
  display: inline-block;
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
    // fallback if no column info
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

function buildGoogleFontsLink(fontFamilies) {
  if (!fontFamilies || fontFamilies.length === 0) return '';
  const uniqueFamilies = Array.from(new Set(fontFamilies));
  const familiesParam = uniqueFamilies
    .map(f => f.trim().replace(/\s+/g, '+'))
    .join('&family=');
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}

// ---------------------------------------
// Utility Functions
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

// ---------------------------------------
// Run the exporter if called directly
// ---------------------------------------
if (require.main === module) {
  // Read docId from command line
  const docId = process.argv[2];
  if (!docId) {
    console.error('Usage: node google-docs-high-fidelity-export.js <DOC_ID>');
    process.exit(1);
  }

  exportDocToHTML(docId).catch(err => console.error('Export error:', err));
}

