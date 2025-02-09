/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Exports a Google Doc to HTML + CSS with:
 *   - Accurate column width (from section/doc style)
 *   - Exact image sizing (no extra scaling)
 *   - Headings (H1..H6)
 *   - Lists (<ul>, <ol>)
 *   - Justified text (text-align: justify)
 *   - Pagination (via @page if doc is paginated)
 *   - Google Fonts link
 *   - Text attributes (bold, italic, underline, color, size)
 *   - Tables
 *   - Table of contents (if doc has one)
 *
 * Creates:
 *   - <OUTPUT_DIR>/index.html
 *   - <OUTPUT_DIR>/.htaccess (sets DirectoryIndex to index.html)
 *   - <OUTPUT_DIR>/images/ (contains image_<id>.png if any images)
 *
 * No footnotes, comments, or other extras.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ------------------ CONFIG ------------------

// Path to your service account JSON (or OAuth creds).
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';

// Default to storing images externally in "images/" folder.
const EMBED_IMAGES_AS_BASE64 = false;

// Maps Docs alignment to CSS
const alignmentMap = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// ---------------------------------------
// Main Export Function
// ---------------------------------------
async function exportDocToHTML(documentId, outputDir) {
  // Ensure the output directory exists
  fs.mkdirSync(outputDir, { recursive: true });

  // Also prepare an images subfolder
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & Setup
  const authClient = await getAuthClient();

  // Fetch the document
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId });
  console.log(`Exporting document: ${doc.title}`);

  // Find column/padding info
  const sectionStyle = findFirstSectionStyle(doc);
  const colInfo = extractColumnInfo(sectionStyle);

  // Prepare sets and arrays
  const usedFonts = new Set(); // track fonts for Google Fonts link
  let htmlOutput = [];

  // Basic HTML skeleton
  htmlOutput.push('<!DOCTYPE html>');
  htmlOutput.push('<html lang="en">');
  htmlOutput.push('<head>');
  htmlOutput.push('  <meta charset="UTF-8">');
  // Only in the <title> tag for the browser
  htmlOutput.push(`  <title>${escapeHtml(doc.title)}</title>`);
  // We'll insert Google Fonts link later if fonts are used
  htmlOutput.push('  <style>');
  htmlOutput.push(generateGlobalCSS(doc, colInfo));
  htmlOutput.push('  </style>');
  htmlOutput.push('</head>');
  htmlOutput.push('<body>');

  // .doc-content container
  htmlOutput.push('<div class="doc-content">');

  // We’ll track current list nesting (for <ul>, <ol>)
  let listStack = [];

  // Go through the doc body content
  const bodyContent = doc.body && doc.body.content ? doc.body.content : [];
  for (const element of bodyContent) {
    if (element.sectionBreak) {
      htmlOutput.push('<div class="section-break"></div>');
      continue;
    }

    // If there's a table of contents, render it
    if (element.tableOfContents) {
      const tocHtml = await renderTableOfContents(
        element.tableOfContents, 
        doc, 
        usedFonts, 
        authClient, 
        outputDir
      );
      // Make sure no open list is pending
      closeAllLists(listStack, htmlOutput);
      htmlOutput.push(tocHtml);
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
        handleListState(listChange, listStack, htmlOutput);
      }
      if (listStack.length > 0) {
        htmlOutput.push(`<li>${html}</li>`);
      } else {
        htmlOutput.push(html);
      }
      continue;
    }

    // Table
    if (element.table) {
      closeAllLists(listStack, htmlOutput);
      const tableHtml = await renderTable(
        element.table, 
        doc, 
        usedFonts, 
        authClient, 
        outputDir,
        imagesDir
      );
      htmlOutput.push(tableHtml);
      continue;
    }

    // other elements can be handled if needed
  }

  // Close any open lists
  closeAllLists(listStack, htmlOutput);

  // Close doc-content & body
  htmlOutput.push('</div>'); // .doc-content
  htmlOutput.push('</body>');
  htmlOutput.push('</html>');

  // Insert <link> for Google Fonts if needed
  const fontLink = buildGoogleFontsLink(Array.from(usedFonts));
  if (fontLink) {
    const insertIndex = htmlOutput.findIndex(l => l.includes('</title>'));
    if (insertIndex >= 0) {
      htmlOutput.splice(
        insertIndex + 1,
        0,
        `  <link rel="stylesheet" href="${fontLink}">`
      );
    }
  }

  // Write index.html
  const indexPath = path.join(outputDir, 'index.html');
  fs.writeFileSync(indexPath, htmlOutput.join('\n'), 'utf8');
  console.log(`HTML exported to: ${indexPath}`);

  // Write .htaccess (only sets directory index)
  const htaccessPath = path.join(outputDir, '.htaccess');
  fs.writeFileSync(htaccessPath, 'DirectoryIndex index.html\n', 'utf8');
  console.log(`.htaccess written to: ${htaccessPath}`);
}

// ---------------------------------------
// findFirstSectionStyle (Missing in your version)
function findFirstSectionStyle(doc) {
  const content = doc.body && doc.body.content ? doc.body.content : [];
  for (const c of content) {
    if (c.sectionBreak && c.sectionBreak.sectionStyle) {
      return c.sectionBreak.sectionStyle;
    }
  }
  // Fallback if no explicit section
  return null;
}

// ---------------------------------------
// extractColumnInfo
function extractColumnInfo(sectionStyle) {
  if (!sectionStyle) return null;
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
// Table of Contents Rendering
// ---------------------------------------
async function renderTableOfContents(
  toc,
  doc,
  usedFonts,
  authClient,
  outputDir
) {
  // The TOC is similar to paragraphs. We'll wrap in <div class="doc-toc">.
  // Inside, each 'content' item might have paragraphs that link to headings.
  // We'll just reuse our paragraph renderer for each paragraph, ignoring lists for now.

  let html = '<div class="doc-toc">\n';
  html += '<h2>Table of Contents</h2>\n';

  if (toc.content) {
    for (const c of toc.content) {
      if (c.paragraph) {
        // We pass an empty listStack to avoid messing up normal doc lists.
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

  html += '</div>';
  return html;
}

// ---------------------------------------
// Paragraph Rendering
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
    if (listStack && listStack.length > 0) {
      const top = listStack[listStack.length - 1];
      listChange = `end${top.toUpperCase()}`;
    }
  }

  // Check if it's a heading
  let tag = 'p';
  if (namedStyleType.startsWith('HEADING_')) {
    const level = parseInt(namedStyleType.replace('HEADING_', ''), 10);
    if (level >= 1 && level <= 6) {
      tag = 'h' + level;
    }
  }

  // Build paragraph-level style
  let inlineStyle = '';

  // Map alignment from Docs to CSS
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

  // Render paragraph elements
  let innerHtml = '';
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

    // Font size
    if (textStyle.fontSize && textStyle.fontSize.magnitude) {
      inlineStyle += `font-size: ${textStyle.fontSize.magnitude}pt;`;
    }
    // Font family
    if (textStyle.weightedFontFamily && textStyle.weightedFontFamily.fontFamily) {
      const fam = textStyle.weightedFontFamily.fontFamily;
      usedFonts.add(fam);
      inlineStyle += `font-family: '${fam}', sans-serif;`;
    }
    // Color
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

  // Hyperlink
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

  // We store as separate files in images/
  const base64Data = await fetchAsBase64(contentUri, authClient);
  const buffer = Buffer.from(base64Data, 'base64');

  // Make a consistent name
  const imgFileName = `image_${objectId}.png`;
  const imgFilePath = path.join(imagesDir, imgFileName);
  fs.writeFileSync(imgFilePath, buffer);

  // Create relative path from index.html to the images folder
  const imgSrc = path.relative(outputDir, imgFilePath);

  // Use exact doc sizing
  let style = '';
  if (size && size.width && size.height) {
    const wPx = ptToPx(size.width.magnitude);
    const hPx = ptToPx(size.height.magnitude);
    style = `width:${wPx}px; height:${hPx}px;`;
  }

  // Alt text
  const alt = embedded.title || embedded.description || '';

  return `<img src="${imgSrc}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// ---------------------------------------
// Table Rendering (Basic)
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
// Helper: Manage List State
// ---------------------------------------
function handleListState(listChange, listStack, htmlOutput) {
  // e.g. "startUL", "endUL", "startOL", "endOL" or combined with '|'
  const actions = listChange.split('|');
  actions.forEach(action => {
    if (action.startsWith('start')) {
      const type = action.replace('start', '').toLowerCase(); // ul or ol
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
// Global CSS Generator
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
h1, h2, h3, h4, h5, h6 {
  margin: 0.8em 0;
  font-family: inherit; /* preserve doc’s font if set in text runs */
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

  // If we have colInfo, set the container width precisely
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
    // fallback if no column data
    lines.push(`
.doc-content {
  max-width: 800px;
  padding: 0 20px;
}
    `);
  }

  // Paginated doc => add @page rule
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

// ---------------------------------------
// Build Google Fonts Link
// ---------------------------------------
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
  // 1pt ~ 1.3333px
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
// Command-line Entry
// ---------------------------------------
if (require.main === module) {
  const docId = process.argv[2];
  const outputDir = process.argv[3];

  if (!docId || !outputDir) {
    console.error('Usage:');
    console.error('  node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>');
    process.exit(1);
  }

  exportDocToHTML(docId, outputDir).catch(err => console.error('Export error:', err));
}

