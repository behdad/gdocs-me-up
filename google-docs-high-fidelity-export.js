/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Forces:
 *   - Narrow column (max-width: 600px)
 *   - line-height: 1.15 for p, li
 *
 * Preserves:
 *   - doc-based alignment (including CENTER, JUSTIFIED, flipping START/END in RTL)
 *   - doc-based direction = RIGHT_TO_LEFT => dir="rtl"
 *   - bullet logic & bullet-lists in RTL if needed
 *   - Title, Subtitle, heading detection
 *   - Table of Contents linking to headings
 *   - Images sized by doc's reported dimension
 *   - Merging identical text runs
 *
 * Also includes .htaccess, minimal rewriting of doc-based lineSpacing logic.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ------------- CONFIG -------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images in "images/" folder

// Basic alignment map for LTR
const alignmentMapLTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// -----------------------------------
// MAIN EXPORT
// -----------------------------------
async function exportDocToHTML(docId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & fetch doc
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({
    documentId: docId
  });
  console.log(`Exporting: ${doc.title}`);

  // Build named styles (Title, Subtitle, headings)
  const namedStylesMap = buildNamedStylesMap(doc);

  // We'll ignore doc-based lineSpacing & column widths, but keep alignment/direction
  const usedFonts = new Set();
  let htmlLines = [];

  // Basic HTML skeleton
  htmlLines.push('<!DOCTYPE html>');
  htmlLines.push('<html lang="en">');
  htmlLines.push('<head>');
  htmlLines.push('  <meta charset="UTF-8">');
  htmlLines.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlLines.push('  <style>');
  // Force narrower column + line-height=1.15
  htmlLines.push(generateGlobalCSS());
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
      // We define renderTableOfContents below
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

  // Insert Google Fonts if needed
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
  fs.writeFileSync(path.join(outputDir, '.htaccess'), 'DirectoryIndex index.html\n', 'utf8');
  console.log(`.htaccess written.`);
}

// -----------------------------------
// generateGlobalCSS => forced narrow column + line-height=1.15
// -----------------------------------
function generateGlobalCSS() {
  // We forcibly set .doc-content to 600px & p/li line-height=1.15
  return `
/* Force narrower column: 600px. Force line-height=1.15 for p,li. */
body {
  margin: 0;
  font-family: sans-serif;
}
.doc-content {
  max-width: 600px;
  margin: 1em auto;
  padding: 0 1em;
}

/* All paragraphs and list items => line-height:1.15 */
p, li {
  margin: 0.5em 0;
  line-height: 1.15;
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

/* Subtitle example */
.doc-subtitle {
  display: block;
  font-size: 1.2em;
  margin: 0.5em 0;
}

/* Table of contents style */
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
`;
}

// -----------------------------------
// TOC function
// -----------------------------------
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
        // We pass an empty listStack so we don't break main bullet logic
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

// -----------------------------------
// Named Styles
// -----------------------------------
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

// -----------------------------------
// Paragraph
// -----------------------------------
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

  // Merge doc-based style for alignment/direction, color, bold, etc.
  let mergedParaStyle = {};
  let mergedTextStyle = {};
  if (namedStylesMap[namedType]) {
    mergedParaStyle = deepCopy(namedStylesMap[namedType].paragraphStyle);
    mergedTextStyle = deepCopy(namedStylesMap[namedType].textStyle);
  }
  deepMerge(mergedParaStyle, style);

  // Bullet
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

  // Title => h1, Subtitle => h2 class="doc-subtitle"
  let tag = 'p';
  if (namedType === 'TITLE') {
    tag = 'h1';
  } else if (namedType === 'SUBTITLE') {
    tag = 'h2 class="doc-subtitle"';
  } else if (namedType.startsWith('HEADING_')) {
    const lvl = parseInt(namedType.replace('HEADING_', ''), 10);
    if (lvl>=1 && lvl<=6) tag=`h${lvl}`;
  }

  let headingIdAttr = '';
  if (style.headingId) {
    headingIdAttr = ` id="heading-${escapeHtml(style.headingId)}"`;
  }

  // We keep alignment flipping for RTL
  let align = mergedParaStyle.alignment; // e.g. CENTER
  if (isRTL && (align === 'START' || align === 'END')) {
    align = (align === 'START')? 'END':'START';
  }

  // Merge text runs
  const mergedRuns = mergeTextRuns(paragraph.elements || []);

  // Build innerHtml
  let innerHtml = '';
  for (const r of mergedRuns) {
    if (r.inlineObjectElement) {
      const objId = r.inlineObjectElement.inlineObjectId;
      innerHtml += await renderInlineObject(objId, doc, authClient, outputDir, imagesDir);
    } else if (r.textRun) {
      innerHtml += renderTextRun(r.textRun, usedFonts, mergedTextStyle);
    }
  }

  // Construct final paragraph
  // We do not forcibly set lineSpacing here—it's done in CSS
  // We do set dir="rtl" if direction=RIGHT_TO_LEFT
  let dirAttr = (isRTL)? ' dir="rtl"' : '';
  let inlineStyle = ''; 
  if (align && alignmentMapLTR[align]) {
    inlineStyle += `text-align: ${alignmentMapLTR[align]};`;
  }

  let paragraphHtml = `<${tag}${headingIdAttr}${dirAttr}`;
  // If we used e.g. "h2 class="doc-subtitle"", handle that carefully
  const spaceIdx = paragraphHtml.indexOf(' ');
  paragraphHtml += `>`;
  paragraphHtml += innerHtml;
  paragraphHtml += `</${tag.split(' ')[0]}>`;

  // If we want to apply inline styles for alignment, we can do so by injecting
  // them into the tag. But you can rely on the p, li, or heading style if you prefer.
  // e.g. if (inlineStyle) apply it. Let’s do that for completeness:
  if (inlineStyle) {
    // Rebuild with inline style
    // E.g. <h2 class="doc-subtitle" dir="rtl" style="text-align:right;">
    const closeBracket = paragraphHtml.indexOf('>');
    if (closeBracket > 0) {
      paragraphHtml = paragraphHtml.slice(0, closeBracket)
        + ` style="${inlineStyle}"` 
        + paragraphHtml.slice(closeBracket);
    }
  }

  return { html: paragraphHtml, listChange };
}

// -----------------------------------
// Merge text runs
// -----------------------------------
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
    'bold','italic','underline','strikethrough',
    'baselineOffset','fontSize','weightedFontFamily',
    'foregroundColor','link'
  ];
  for (const f of fields) {
    if (JSON.stringify(a[f]||null) !== JSON.stringify(b[f]||null)) {
      return false;
    }
  }
  return true;
}

// -----------------------------------
// Render Text Run
// -----------------------------------
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
    const hex = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
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
      if (cssClasses.length>0) {
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

// -----------------------------------
// Inline Objects (Images)
// -----------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir) {
  const inlineObj = doc.inlineObjects?.[objectId];
  if (!inlineObj) return '';

  const embedded = inlineObj.inlineObjectProperties?.embeddedObject;
  if (!embedded?.imageProperties) return '';

  const { imageProperties } = embedded;
  const { contentUri, size } = imageProperties;

  const base64Data = await fetchAsBase64(contentUri, authClient);
  const buffer = Buffer.from(base64Data, 'base64');
  const fileName = `image_${objectId}.png`;
  const filePath = path.join(imagesDir, fileName);
  fs.writeFileSync(filePath, buffer);

  const imgSrc = path.relative(outputDir, filePath);

  let style = '';
  if (size?.width?.magnitude && size?.height?.magnitude) {
    const wPx = Math.round(size.width.magnitude * 1.3333);
    const hPx = Math.round(size.height.magnitude * 1.3333);
    // doc-based dimension => max
    style = `max-width:${wPx}px; max-height:${hPx}px;`;
  }

  const alt = embedded.title || embedded.description || '';
  return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// -----------------------------------
// Lists
// -----------------------------------
function detectListChange(bullet, doc, listStack, isRTL) {
  const listId = bullet.listId;
  const nestingLevel = bullet.nestingLevel||0;
  const listDef = doc.lists?.[listId];
  if(!listDef?.listProperties?.nestingLevels) return null;

  const glyph = listDef.listProperties.nestingLevels[nestingLevel];
  const isNumbered = glyph?.glyphType?.toLowerCase().includes('number');
  const top = listStack[listStack.length-1];

  const startType = isNumbered?'OL':'UL';
  const rtlFlag = isRTL?'_RTL':'';

  if(!top || !top.startsWith(startType.toLowerCase())) {
    if(top) {
      return `end${top.toUpperCase()}|start${startType}${rtlFlag}`;
    } else {
      return `start${startType}${rtlFlag}`;
    }
  }
  return null;
}

function handleListState(listChange, listStack, htmlLines) {
  const actions = listChange.split('|');
  for(const action of actions) {
    if(action.startsWith('start')) {
      if(action.includes('UL_RTL')) {
        htmlLines.push('<ul dir="rtl">');
        listStack.push('ul_rtl');
      } else if(action.includes('OL_RTL')) {
        htmlLines.push('<ol dir="rtl">');
        listStack.push('ol_rtl');
      } else if(action.includes('UL')) {
        htmlLines.push('<ul>');
        listStack.push('ul');
      } else {
        htmlLines.push('<ol>');
        listStack.push('ol');
      }
    } else if(action.startsWith('end')) {
      const top = listStack.pop();
      if(top.startsWith('u')) htmlLines.push('</ul>');
      else htmlLines.push('</ol>');
    }
  }
}

function closeAllLists(listStack, htmlLines) {
  while(listStack.length>0) {
    const top = listStack.pop();
    if(top.startsWith('u')) htmlLines.push('</ul>');
    else htmlLines.push('</ol>');
  }
}

// -----------------------------------
// Table
// -----------------------------------
async function renderTable(
  table, doc, usedFonts, authClient, outputDir, imagesDir, namedStylesMap
) {
  let html = '<table class="doc-table" style="border-collapse: collapse; border: 1px solid #ccc;">';
  for(const row of table.tableRows||[]) {
    html += '<tr>';
    for(const cell of row.tableCells||[]) {
      html += '<td style="border:1px solid #ccc; padding:0.5em;">';
      for(const c of cell.content||[]) {
        if(c.paragraph) {
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

// -----------------------------------
// Utility
// -----------------------------------
function deepMerge(base, overlay) {
  for(const k in overlay) {
    if(
      typeof overlay[k]==='object' &&
      overlay[k]!==null &&
      !Array.isArray(overlay[k])
    ) {
      if(!base[k]) base[k]={};
      deepMerge(base[k], overlay[k]);
    } else {
      base[k]=overlay[k];
    }
  }
}
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}
function escapeHtml(str) {
  return str
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
function ptToPx(pts) {
  return Math.round(pts*1.3333);
}
function rgbToHex(r,g,b){
  const nr=Math.round(r*255);
  const ng=Math.round(g*255);
  const nb=Math.round(b*255);
  return '#'+[nr,ng,nb].map(x=>x.toString(16).padStart(2,'0')).join('');
}
function buildGoogleFontsLink(fontFamilies) {
  if(!fontFamilies||fontFamilies.length===0)return'';
  const unique=Array.from(new Set(fontFamilies));
  const familiesParam=unique.map(f=>f.trim().replace(/\s+/g,'+')).join('&family=');
  return`https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}
async function getAuthClient(){
  const auth=new google.auth.GoogleAuth({
    keyFile: SERVICE_ACCOUNT_KEY_FILE,
    scopes:[
      'https://www.googleapis.com/auth/documents.readonly',
      'https://www.googleapis.com/auth/drive.readonly'
    ]
  });
  return auth.getClient();
}
async function fetchAsBase64(url, authClient) {
  const resp=await authClient.request({
    url,
    method:'GET',
    responseType:'arraybuffer'
  });
  return Buffer.from(resp.data,'binary').toString('base64');
}

// CLI
if(require.main===module){
  const docId=process.argv[2];
  const outDir=process.argv[3];
  if(!docId||!outDir){
    console.error('Usage: node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>');
    process.exit(1);
  }
  exportDocToHTML(docId, outDir).catch(err=>console.error('Export error:',err));
}

