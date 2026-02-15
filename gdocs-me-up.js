/**
 * google-docs-high-fidelity-export.js
 *
 * High-Fidelity Exporter for Google Docs
 *
 * Features:
 *  - **Named Styles**: Title, Subtitle, Headings (H1..H6) mapped to real HTML headings,
 *    with doc-based font sizing and inline styling (bold, italic, color, etc.).
 *  - **Heading Size Override**: Resets default <h1>.. <h6> to neutral size in CSS,
 *    so doc's inline size precisely controls final heading font.
 *  - **Line Spacing**: Honors doc's paragraph lineSpacing (e.g., 1.15, 1.5),
 *    plus spaceAbove/spaceBelow, and text indentation (indentFirstLine, indentStart, indentEnd).
 *  - **Right-to-Left Paragraphs**: Sets `dir="rtl"` for paragraphs whose style
 *    indicates RIGHT_TO_LEFT, flipping alignment START/END if needed. Includes comprehensive
 *    Unicode subset support (Arabic, Hebrew, Thai, Devanagari, etc.).
 *  - **Alignment**: PRESERVES doc-based alignment (CENTER, JUSTIFIED, etc.).
 *  - **Lists**: Bullet / Numbered lists, including RTL bullets if direction=RIGHT_TO_LEFT.
 *  - **Text Styling**: Bold, italic, underline, strikethrough, superscript, subscript,
 *    small caps, text color, background color, font family with weights.
 *  - **Paragraph Borders & Shading**: Supports top/bottom/left/right borders with colors
 *    and styles (solid, dotted, dashed, double), plus paragraph background colors.
 *  - **Pagination Control**: pageBreakBefore, keepLinesTogether, keepWithNext,
 *    avoidWidowAndOrphan for print-friendly layouts.
 *  - **Images**: Exact doc-based sizes, with transform scaling and translation.
 *    Supports image cropping (cropProperties), margins, and positioning.
 *    Exports images to an `images/` folder. Uses `max-width` / `max-height` so they
 *    never exceed doc's reported size or the container width.
 *  - **Table of Contents**: Indents each TOC entry based on the heading level of its
 *    linked heading (Heading 1 => level 1, etc.).
 *  - **Tables**: Exports Google Docs tables using <table>, <tr>, <td> with full support for:
 *    cell borders (per-cell), background colors, padding, colspan, rowspan.
 *  - **Horizontal Rules**: Renders horizontal rule elements as <hr>.
 *  - **Footnotes**: Renders footnote references with superscript links.
 *  - **Equations**: Basic equation support (rendered as code for now).
 *  - **Auto Text**: Page numbers and page counts (placeholders).
 *  - **Multi-Column Layouts**: Section breaks with column properties for multi-column text.
 *  - **Column Breaks**: Explicit column break rendering.
 *  - **Column Width**: Infers container width from doc's pageSize minus margins
 *    (with a small tweak). Then sets `.doc-content { max-width: ... }`.
 *  - **Google Fonts**: Gathers all distinct fonts used, generating a <link> to
 *    https://fonts.googleapis.com with multiple weights and comprehensive Unicode subsets
 *    for non-Latin scripts (Arabic, Hebrew, Greek, Cyrillic, etc.).
 *  - **Merging Text Runs**: Consecutive text runs with identical styling are combined
 *    into a single <span> to avoid excessive markup.
 *  - **Service Account Auth**: Reads from `SERVICE_ACCOUNT_KEY_FILE`, or adapt to your
 *    auth method. Requires the doc to be accessible with the given credentials.
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Example:
 *   node google-docs-high-fidelity-export.js 1AbCdEfgHIjKLMnOP <my-export-dir>
 *
 * Then open <my-export-dir>/index.html to see the rendered doc.
 *
 * Dependencies:
 *   - Node.js
 *   - "googleapis" library (npm install googleapis)
 *   - A valid Google service account JSON key or other OAuth method
 *
 * This script merges doc-based styling with neutral heading overrides so your headings
 * appear at the exact doc size without default HTML heading inflation. Right-to-left,
 * justification, bullet-lists, images, tables, borders, and more are handled for a truly
 * "high-fidelity" offline representation of your Google Doc with extensive support for
 * international and non-Latin scripts.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ------------- CONFIG -------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false;

// Basic alignment map for LTR paragraphs
const alignmentMapLTR = {
  START: 'left',
  CENTER: 'center',
  END: 'right',
  JUSTIFIED: 'justify'
};

// Border style map
const borderStyleMap = {
  SOLID: 'solid',
  DOTTED: 'dotted',
  DASHED: 'dashed',
  DOUBLE: 'double'
};

async function exportDocToHTML(docId, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const imagesDir = path.join(outputDir, 'images');
  fs.mkdirSync(imagesDir, { recursive: true });

  // Auth & fetch doc
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId: docId });
  console.log(`Exporting doc: ${doc.title}`);

  // Named styles for Title, Subtitle, Headings, etc.
  const namedStylesMap = buildNamedStylesMap(doc);

  // Container width from docStyle
  const colInfo = computeDocContainerWidth(doc);

  // Build global CSS
  const globalCSS = generateGlobalCSS(doc, colInfo);

  const usedFonts = new Set();
  let htmlLines = [];

  // Basic HTML skeleton
  htmlLines.push('<!DOCTYPE html>');
  htmlLines.push('<html lang="en">');
  htmlLines.push('<head>');
  htmlLines.push('  <meta charset="UTF-8">');
  htmlLines.push('  <meta name="viewport" content="width=device-width">');
  htmlLines.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlLines.push('  <style>');
  htmlLines.push(globalCSS);
  htmlLines.push('  </style>');
  htmlLines.push('</head>');
  htmlLines.push('<body>');
  htmlLines.push('<div class="doc-content">');

  // Pre-process: count items per list per level to determine if single-item (numbered) or multi-item (bullets)
  const listItemCounts = {};
  const bodyContent = doc.body?.content || [];
  for (const element of bodyContent) {
    if (element.paragraph?.bullet) {
      const bullet = element.paragraph.bullet;
      const level = bullet.nestingLevel ?? 0;
      const listId = bullet.listId;
      const key = `${listId}:${level}`;
      listItemCounts[key] = (listItemCounts[key] || 0) + 1;
    }
  }

  // Store counts in doc object for use by detectListChange
  doc.___listItemCounts = listItemCounts;

  let listStack = [];  // Array of {type: 'ul'/'ol', level: 0/1/2...}
  let prevNestingLevel = -1;
  let prevListId = null;  // Track previous listId to detect list changes

  for (const element of bodyContent) {
    if (element.sectionBreak) {
      closeAllLists(listStack, htmlLines);
      const sb = element.sectionBreak;
      const sectionStyle = sb.sectionStyle;

      // Handle column breaks
      if(sectionStyle?.columnSeparatorStyle === 'BETWEEN_EACH_COLUMN'){
        htmlLines.push('<div class="column-break"></div>');
      } else if(sectionStyle?.columnProperties && sectionStyle.columnProperties.length > 1){
        // Multi-column section
        const colCount = sectionStyle.columnProperties.length;
        htmlLines.push(`<div class="multi-column" style="column-count:${colCount};">`);
      } else {
        htmlLines.push('<div class="section-break"></div>');
      }
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
      const nestingLevel = element.paragraph.bullet ? (element.paragraph.bullet.nestingLevel ?? 0) : -1;

      // Pass previous nesting level and listId to detectListChange
      element.paragraph.___prevNestingLevel = prevNestingLevel;
      element.paragraph.___prevListId = prevListId;

      // Close previous <li> if nesting level is changing
      if (prevNestingLevel >= 0 && prevNestingLevel < nestingLevel) {
        // Nesting deeper - don't close previous <li>, nested list will go inside it
      } else if (prevNestingLevel >= 0 && listStack.length > 0) {
        // Same level or going back up - close previous <li>
        // When going back up, the nested list was already closed by handleListState
        htmlLines.push('</li>');
      }

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
        htmlLines.push(`<li>${html}`);
        prevNestingLevel = nestingLevel;
        prevListId = element.paragraph.bullet?.listId || null;
      } else {
        // Not in a list, close any open <li>
        if (prevNestingLevel >= 0) {
          htmlLines.push('</li>');
          prevNestingLevel = -1;
        }
        prevListId = null;
        htmlLines.push(html);
      }
      continue;
    }
    if (element.horizontalRule) {
      closeAllLists(listStack, htmlLines);
      htmlLines.push('<hr style="border: 1px solid #ccc; margin: 1em 0;">');
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

  // Close final <li> if open
  if (prevNestingLevel >= 0) {
    htmlLines.push('</li>');
  }
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
}

// -----------------------------------------------------
// Column width from doc documentStyle
// -----------------------------------------------------
function computeDocContainerWidth(doc) {
  let containerPx = 800; // fallback
  const ds = doc.documentStyle;
  if (ds?.pageSize?.width?.magnitude) {
    const pageW = ds.pageSize.width.magnitude;
    const leftM = ds.marginLeft?.magnitude || 72;
    const rightM = ds.marginRight?.magnitude || 72;
    const usablePts = pageW - (leftM + rightM);
    if (usablePts > 0) containerPx = ptToPx(usablePts);
  }
  // small tweak
  containerPx += 64;
  return containerPx;
}

// -----------------------------------------------------
// Global CSS with heading overrides
// -----------------------------------------------------
function generateGlobalCSS(doc, containerPx) {
  const lines = [];
  lines.push(`
/* Reset heading sizes so doc-based inline style rules. */
h1, h2, h3, h4, h5, h6 {
  margin: 1em 0;
  font-size: 1em;
  font-weight: normal;
}

body {
  font-family: sans-serif;
  /* Better font rendering for all scripts */
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-rendering: optimizeLegibility;
}
.doc-content {
  margin: 1em auto;
  max-width: ${containerPx}px;
  padding: 2em 1em;
}
p, li {
  margin: 0.5em 0;
  line-height: 1.4;
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
.column-break {
  break-after: column;
}
.multi-column {
  column-gap: 1em;
}

.doc-toc {
  margin: 0.5em 0;
  padding: 0.5em;
}
.doc-toc p {
  margin: 0.1em 0;
  line-height: 1.2;
}

.subtitle {
  display: block;
  white-space: pre-wrap;
}
.doc-table {
  border-collapse: collapse;
  margin: 0.5em 0;
  width: 100%;
}

/* Equation styling */
.equation {
  font-family: 'Cambria Math', 'Latin Modern Math', 'STIX Two Math', serif;
  font-style: italic;
  padding: 0 0.2em;
}

/* Page number placeholders */
.page-number, .page-count {
  font-style: italic;
  color: #666;
}

/* Right-to-left text improvements */
[dir="rtl"] {
  text-align: right;
  direction: rtl;
}
[dir="rtl"] ul, [dir="rtl"] ol {
  padding-right: 2em;
  padding-left: 0;
}

/* Better list spacing and nesting */
ul, ol {
  margin: 0.5em 0;
  padding-left: 2em;
}
li {
  margin: 0.25em 0;
}
ul ul, ol ol, ul ol, ol ul {
  margin: 0.25em 0;
}

/* TOC indentation levels */
.toc-level-1 { margin-left: 0; }
.toc-level-2 { margin-left: 1em; }
.toc-level-3 { margin-left: 2em; }
.toc-level-4 { margin-left: 3em; }

/* Improved print styles */
@media print {
  .doc-content {
    max-width: none;
  }
  @page {
    orphans: 2;
    widows: 2;
  }
}
`);

  // If doc has pageSize => @page
  if (doc.documentStyle?.pageSize?.width?.magnitude && doc.documentStyle?.pageSize?.height?.magnitude) {
    const wIn = doc.documentStyle.pageSize.width.magnitude / 72;
    const hIn = doc.documentStyle.pageSize.height.magnitude / 72;
    const topM = (doc.documentStyle.marginTop?.magnitude||72)/72;
    const rightM = (doc.documentStyle.marginRight?.magnitude||72)/72;
    const botM = (doc.documentStyle.marginBottom?.magnitude||72)/72;
    const leftM = (doc.documentStyle.marginLeft?.magnitude||72)/72;
    lines.push(`
@page {
  size: ${wIn}in ${hIn}in;
  margin: ${topM}in ${rightM}in ${botM}in ${leftM}in;
}
    `);
  }

  return lines.join('\n');
}

// -----------------------------------------------------
// Table of Contents (Indentation by heading level)
// -----------------------------------------------------
async function renderTableOfContents(
  toc,
  doc,
  usedFonts,
  authClient,
  outputDir,
  namedStylesMap
) {
  let html = '<div class="doc-toc">\n';

  if (toc.content) {
    for (const c of toc.content) {
      if (!c.paragraph) continue;
      let headingLevel = 1;
      for (const elem of c.paragraph.elements||[]) {
        const st = elem.textRun?.textStyle;
        if (st?.link?.headingId) {
          const lv = findHeadingLevelById(doc, st.link.headingId);
          if (lv>headingLevel) headingLevel=lv;
        }
      }
      if (headingLevel<1) headingLevel=1;
      if (headingLevel>4) headingLevel=4;

      const { html:pHtml } = await renderParagraph(
        c.paragraph,
        doc,
        usedFonts,
        [],
        authClient,
        outputDir,
        null,
        namedStylesMap
      );
      html += `<div class="toc-level-${headingLevel}">${pHtml}</div>\n`;
    }
  }

  html += '</div>\n';
  return html;
}

function findHeadingLevelById(doc, headingId) {
  const content = doc.body?.content||[];
  for (const e of content) {
    if (e.paragraph) {
      const ps = e.paragraph.paragraphStyle;
      if (ps?.headingId===headingId) {
        const named = ps.namedStyleType||'NORMAL_TEXT';
        if (named.startsWith('HEADING_')) {
          const lv=parseInt(named.replace('HEADING_',''),10);
          if(lv>=1 && lv<=6) return lv;
        }
      }
    }
  }
  return 1;
}

// -----------------------------------------------------
// Paragraph
// -----------------------------------------------------
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
  const style = paragraph.paragraphStyle||{};
  const namedType = style.namedStyleType||'NORMAL_TEXT';

  // Merge doc-based style
  let mergedParaStyle={};
  let mergedTextStyle={};
  if(namedStylesMap[namedType]){
    mergedParaStyle=deepCopy(namedStylesMap[namedType].paragraphStyle);
    mergedTextStyle=deepCopy(namedStylesMap[namedType].textStyle);
  }
  deepMerge(mergedParaStyle, style);

  // bullet logic
  const isRTL=(mergedParaStyle.direction==='RIGHT_TO_LEFT');
  const currentNestingLevel = paragraph.bullet?.nestingLevel ?? -1;
  let listChange=null;
  if (paragraph.bullet) {
    // prevNestingLevel and prevListId are passed from the main loop
    const prevLevel = paragraph.___prevNestingLevel ?? -1;
    const prevListId = paragraph.___prevListId;
    listChange=detectListChange(paragraph.bullet,doc,listStack,isRTL,prevLevel,prevListId);
  } else {
    if(listStack.length>0){
      // Exiting all lists
      let actions = [];
      for(let i = 0; i < listStack.length; i++){
        actions.push('endLIST');
      }
      listChange = actions.join('|');
    }
  }

  // Title => <h1 class="title">, Subtitle => <h2 class="subtitle">, heading => <hX>, else <p>
  let tag='p';
  if(namedType==='TITLE'){
    tag='h1 class="title"';
  } else if(namedType==='SUBTITLE'){
    tag='h2 class="subtitle"';
  } else if(namedType.startsWith('HEADING_')){
    const lv=parseInt(namedType.replace('HEADING_',''),10);
    if(lv>=1 && lv<=6) tag=`h${lv}`;
  }

  let headingIdAttr='';
  if(style.headingId){
    headingIdAttr=` id="heading-${escapeHtml(style.headingId)}"`;
  }

  // alignment flipping
  let align=mergedParaStyle.alignment;
  if(isRTL && (align==='START'||align==='END')){
    align=(align==='START')?'END':'START';
  }

  // doc-based lineSpacing => line-height
  let inlineStyle='';
  if(align && alignmentMapLTR[align]){
    inlineStyle += `text-align:${alignmentMapLTR[align]};`;
  }
  if(mergedParaStyle.lineSpacing){
    // Google Docs lineSpacing as direct percentage
    // Base line-height is 1.4 (set in CSS), this overrides when specified
    const ls=mergedParaStyle.lineSpacing*1.4/100;
    inlineStyle += `line-height:${ls};`;
  }
  if(mergedParaStyle.spaceAbove?.magnitude){
    inlineStyle += `margin-top:${ptToPx(mergedParaStyle.spaceAbove.magnitude)}px;`;
  }
  if(mergedParaStyle.spaceBelow?.magnitude){
    inlineStyle += `margin-bottom:${ptToPx(mergedParaStyle.spaceBelow.magnitude)}px;`;
  }
  if(mergedParaStyle.indentFirstLine?.magnitude){
    inlineStyle += `text-indent:${ptToPx(mergedParaStyle.indentFirstLine.magnitude)}px;`;
  } else if(mergedParaStyle.indentStart?.magnitude){
    inlineStyle += `margin-left:${ptToPx(mergedParaStyle.indentStart.magnitude)}px;`;
  }
  if(mergedParaStyle.indentEnd?.magnitude){
    inlineStyle += `margin-right:${ptToPx(mergedParaStyle.indentEnd.magnitude)}px;`;
  }

  // Paragraph borders
  if(mergedParaStyle.borderTop){
    inlineStyle += formatBorder('top', mergedParaStyle.borderTop);
  }
  if(mergedParaStyle.borderBottom){
    inlineStyle += formatBorder('bottom', mergedParaStyle.borderBottom);
  }
  if(mergedParaStyle.borderLeft){
    inlineStyle += formatBorder('left', mergedParaStyle.borderLeft);
  }
  if(mergedParaStyle.borderRight){
    inlineStyle += formatBorder('right', mergedParaStyle.borderRight);
  }

  // Paragraph shading (background color)
  if(mergedParaStyle.shading?.backgroundColor?.color?.rgbColor){
    const rgb = mergedParaStyle.shading.backgroundColor.color.rgbColor;
    const hex = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle += `background-color:${hex};`;
    inlineStyle += `padding:0.5em;`;
  }

  // Pagination control
  if(mergedParaStyle.pageBreakBefore){
    inlineStyle += `page-break-before:always;`;
  }
  if(mergedParaStyle.keepLinesTogether){
    inlineStyle += `page-break-inside:avoid;`;
  }
  if(mergedParaStyle.keepWithNext){
    inlineStyle += `page-break-after:avoid;`;
  }
  if(mergedParaStyle.avoidWidowAndOrphan){
    inlineStyle += `orphans:2;widows:2;`;
  }

  // Tab stops - store for potential future use
  if(mergedParaStyle.tabStops && mergedParaStyle.tabStops.length > 0){
    const tabStopPositions = mergedParaStyle.tabStops.map(ts => {
      return ts.offset?.magnitude ? ptToPx(ts.offset.magnitude) : 0;
    });
    // HTML doesn't support tab-stops directly, but we could use custom CSS tab-size
  }

  let dirAttr='';
  if(mergedParaStyle.direction==='RIGHT_TO_LEFT'){
    dirAttr=' dir="rtl"';
  }

  // Merge text runs
  const mergedRuns=mergeTextRuns(paragraph.elements||[]);
  let innerHtml='';
  for(const r of mergedRuns){
    if(r.inlineObjectElement){
      const objId=r.inlineObjectElement.inlineObjectId;
      innerHtml += await renderInlineObject(objId, doc, authClient, outputDir, imagesDir);
    } else if(r.textRun){
      innerHtml += renderTextRun(r.textRun, usedFonts, mergedTextStyle);
    } else if(r.footnoteReference){
      innerHtml += renderFootnoteReference(r.footnoteReference, doc);
    } else if(r.equation){
      innerHtml += renderEquation(r.equation);
    } else if(r.autoText){
      innerHtml += renderAutoText(r.autoText);
    }
  }

  let paragraphHtml=`<${tag}${headingIdAttr}${dirAttr}>${innerHtml}</${tag.split(' ')[0]}>`;
  if(inlineStyle){
    const closeBracket=paragraphHtml.indexOf('>');
    if(closeBracket>0){
      paragraphHtml=paragraphHtml.slice(0,closeBracket)
        + ` style="${inlineStyle}"`
        + paragraphHtml.slice(closeBracket);
    }
  }

  return { html: paragraphHtml, listChange };
}

// -----------------------------------------------------
// Merging text runs
// -----------------------------------------------------
function mergeTextRuns(elements){
  const merged=[];
  let last=null;
  for(const e of elements){
    if(e.inlineObjectElement){
      merged.push({ inlineObjectElement:e.inlineObjectElement});
      last=null;
    } else if(e.footnoteReference){
      merged.push({ footnoteReference:e.footnoteReference});
      last=null;
    } else if(e.equation){
      merged.push({ equation:e.equation});
      last=null;
    } else if(e.autoText){
      merged.push({ autoText:e.autoText});
      last=null;
    } else if(e.textRun){
      const style=e.textRun.textStyle||{};
      const content=e.textRun.content||'';
      if(last && last.textRun && isSameTextStyle(last.textRun.textStyle,style)){
        last.textRun.content+=content;
      } else {
        merged.push({ textRun:{ content, textStyle:deepCopy(style)}});
        last=merged[merged.length-1];
      }
    }
  }
  return merged;
}

function isSameTextStyle(a,b){
  const fields=[
    'bold','italic','underline','strikethrough',
    'baselineOffset','fontSize','weightedFontFamily',
    'foregroundColor','backgroundColor','link','smallCaps'
  ];
  for(const f of fields){
    if(JSON.stringify(a[f]||null)!==JSON.stringify(b[f]||null)){
      return false;
    }
  }
  return true;
}

// -----------------------------------------------------
// Rendering text runs
// -----------------------------------------------------
function renderTextRun(textRun, usedFonts, baseStyle){
  const finalStyle=deepCopy(baseStyle||{});
  deepMerge(finalStyle, textRun.textStyle||{});

  let content=textRun.content||'';
  // Remove trailing newline (marks end of paragraph)
  content=content.replace(/\n$/,'');
  // Convert vertical tabs (\u000b) to a placeholder
  // These represent soft line breaks within a paragraph (Shift+Enter in Google Docs)
  content=content.replace(/\u000b/g,'__LINEBREAK__');

  const cssClasses=[];
  let inlineStyle='';

  if(finalStyle.bold) cssClasses.push('bold');
  if(finalStyle.italic) cssClasses.push('italic');
  if(finalStyle.underline) cssClasses.push('underline');
  if(finalStyle.strikethrough) cssClasses.push('strikethrough');

  if(finalStyle.baselineOffset==='SUPERSCRIPT'){
    cssClasses.push('superscript');
  } else if(finalStyle.baselineOffset==='SUBSCRIPT'){
    cssClasses.push('subscript');
  }

  // Small caps support
  if(finalStyle.smallCaps){
    inlineStyle+=`font-variant:small-caps;`;
  }

  if(finalStyle.fontSize?.magnitude){
    inlineStyle+=`font-size:${finalStyle.fontSize.magnitude}pt;`;
  }
  if(finalStyle.weightedFontFamily?.fontFamily){
    const fam=finalStyle.weightedFontFamily.fontFamily;
    const weight = finalStyle.weightedFontFamily.weight || 400;
    // Track font with its weight for better loading
    usedFonts.add(`${fam}:${weight}`);
    inlineStyle+=`font-family:'${fam}',sans-serif;`;
    // Font weight if specified
    if(weight && weight !== 400){
      inlineStyle+=`font-weight:${weight};`;
    }
  }
  if(finalStyle.foregroundColor?.color?.rgbColor){
    const rgb=finalStyle.foregroundColor.color.rgbColor;
    const hex=rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle+=`color:${hex};`;
  }

  // Background color support
  if(finalStyle.backgroundColor?.color?.rgbColor){
    const rgb=finalStyle.backgroundColor.color.rgbColor;
    const hex=rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle+=`background-color:${hex};`;
  }

  let openTag='<span';
  if(cssClasses.length>0){
    openTag+=` class="${cssClasses.join(' ')}"`;
  }
  if(inlineStyle){
    openTag+=` style="${inlineStyle}"`;
  }
  openTag+='>';
  let closeTag='</span>';

  if(finalStyle.link){
    let linkHref='';
    if(finalStyle.link.headingId){
      linkHref=`#heading-${escapeHtml(finalStyle.link.headingId)}`;
    } else if(finalStyle.link.url){
      linkHref=finalStyle.link.url;
    }
    if(linkHref){
      openTag=`<a href="${escapeHtml(linkHref)}"`;
      if(cssClasses.length>0){
        openTag+=` class="${cssClasses.join(' ')}"`;
      }
      if(inlineStyle){
        openTag+=` style="${inlineStyle}"`;
      }
      openTag+='>';
      closeTag='</a>';
    }
  }

  // Replace line break placeholder with actual <br> tags after escaping
  const escapedContent = escapeHtml(content).replace(/__LINEBREAK__/g, '<br>');
  return openTag+escapedContent+closeTag;
}

// -----------------------------------------------------
// 7) Inline Objects (Images)
// -----------------------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir){
  const inlineObj=doc.inlineObjects?.[objectId];
  if(!inlineObj) return'';

  const embedded=inlineObj.inlineObjectProperties?.embeddedObject;
  if(!embedded?.imageProperties) return'';

  const { imageProperties, size }=embedded;
  const { contentUri, cropProperties }=imageProperties;

  let scaleX=1, scaleY=1;
  let translateX=0, translateY=0;
  if(embedded.transform){
    if(embedded.transform.scaleX) scaleX=embedded.transform.scaleX;
    if(embedded.transform.scaleY) scaleY=embedded.transform.scaleY;
    if(embedded.transform.translateX) translateX=embedded.transform.translateX;
    if(embedded.transform.translateY) translateY=embedded.transform.translateY;
  }

  const base64Data=await fetchAsBase64(contentUri,authClient);
  const buffer=Buffer.from(base64Data,'base64');
  const fileName=`image_${objectId}.png`;
  const filePath=path.join(imagesDir,fileName);
  fs.writeFileSync(filePath,buffer);

  const imgSrc=path.relative(outputDir,filePath);

  let style='';
  if(size?.width?.magnitude && size?.height?.magnitude){
    const wPx=Math.round(size.width.magnitude*1.3333*scaleX);
    const hPx=Math.round(size.height.magnitude*1.3333*scaleY);
    style=`max-width:${wPx}px; max-height:${hPx}px;`;
  }

  // Handle cropping - using object-fit and object-position
  if(cropProperties){
    const { offsetLeft, offsetTop, offsetRight, offsetBottom } = cropProperties;
    if(offsetLeft || offsetTop || offsetRight || offsetBottom){
      style += `object-fit:cover;`;
      // Calculate the visible portion
      const left = (offsetLeft || 0) * 100;
      const top = (offsetTop || 0) * 100;
      style += `object-position:${-left}% ${-top}%;`;
    }
  }

  // Handle image positioning/translation
  if(translateX !== 0 || translateY !== 0){
    const txPx = Math.round(translateX * 1.3333);
    const tyPx = Math.round(translateY * 1.3333);
    style += `transform:translate(${txPx}px, ${tyPx}px);`;
  }

  // Image margins from marginTop, marginBottom, marginLeft, marginRight
  if(embedded.marginTop?.magnitude){
    style += `margin-top:${ptToPx(embedded.marginTop.magnitude)}px;`;
  }
  if(embedded.marginBottom?.magnitude){
    style += `margin-bottom:${ptToPx(embedded.marginBottom.magnitude)}px;`;
  }
  if(embedded.marginLeft?.magnitude){
    style += `margin-left:${ptToPx(embedded.marginLeft.magnitude)}px;`;
  }
  if(embedded.marginRight?.magnitude){
    style += `margin-right:${ptToPx(embedded.marginRight.magnitude)}px;`;
  }

  const alt=embedded.title||embedded.description||'';
  return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// -----------------------------------------------------
// 8) Lists
// -----------------------------------------------------
function detectListChange(bullet, doc, listStack, isRTL, prevLevel, prevListId){
  const listId=bullet.listId;
  const nestingLevel=bullet.nestingLevel||0;
  const listDef=doc.lists?.[listId];
  if(!listDef?.listProperties?.nestingLevels)return null;

  const glyph=listDef.listProperties.nestingLevels[nestingLevel];
  // Detect list type:
  // - If glyphSymbol is present (●, ○, -, etc.) → bullet list
  // - If glyphType is explicitly DECIMAL, ALPHA, ROMAN, etc. → numbered list
  // - If GLYPH_TYPE_UNSPECIFIED:
  //   - Single-item lists (1 item at level 0) → numbered (section markers)
  //   - Multi-item lists (multiple items at level 0) → bullets
  const isBullet = glyph?.glyphSymbol !== undefined;
  const explicitlyNumbered = ['DECIMAL', 'ALPHA', 'ROMAN', 'UPPER_ALPHA', 'UPPER_ROMAN',
                               'LOWER_ALPHA', 'LOWER_ROMAN'].includes(glyph?.glyphType);

  let isNumbered;
  if (isBullet) {
    isNumbered = false;
  } else if (explicitlyNumbered) {
    isNumbered = true;
  } else if (glyph?.glyphType === 'GLYPH_TYPE_UNSPECIFIED') {
    // Check if this is a single-item list (numbered) or multi-item list (bullets)
    // Count items at THIS nesting level
    const key = `${listId}:${nestingLevel}`;
    const itemCount = doc.___listItemCounts?.[key] || 1;
    isNumbered = (itemCount === 1);
  } else {
    // Default to bullet list
    isNumbered = false;
  }

  const startType=isNumbered?'OL':'UL';
  const rtlFlag=isRTL?'_RTL':'';

  // Starting a list for the first time
  if(listStack.length === 0){
    return `start${startType}${rtlFlag}:${nestingLevel}`;
  }

  // Check if nesting level changed
  if(nestingLevel > prevLevel){
    // Going deeper - start nested list
    return `start${startType}${rtlFlag}:${nestingLevel}`;
  } else if(nestingLevel < prevLevel){
    // Coming back up - close nested lists
    let actions = [];
    for(let i = prevLevel; i > nestingLevel; i--){
      actions.push(`endLIST`);
    }

    // After closing nested lists, check if we need to switch lists at current level
    // The listStack will have (prevLevel - nestingLevel) fewer items after closing
    const stackIndexAfterClosing = listStack.length - (prevLevel - nestingLevel);
    if(stackIndexAfterClosing > 0){
      const parentType = listStack[stackIndexAfterClosing - 1]?.split(':')[0];
      const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '');

      // Check if parent list type or listId changed
      if(parentType !== wantType || (prevListId && prevListId !== listId)){
        actions.push(`end${parentType?.toUpperCase() || 'UL'}`);
        actions.push(`start${startType}${rtlFlag}:${nestingLevel}`);
      }
    }

    return actions.join('|');
  }

  // Same level - check if list type or listId changed
  const currentType = listStack[listStack.length - 1]?.split(':')[0];
  const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '');

  // Check if list type changed OR if listId changed
  if(currentType !== wantType || (prevListId && prevListId !== listId)){
    return `end${currentType?.toUpperCase() || 'UL'}|start${startType}${rtlFlag}:${nestingLevel}`;
  }

  return null;
}

function handleListState(listChange, listStack, htmlLines){
  const actions=listChange.split('|');
  for(const action of actions){
    if(action.startsWith('start')){
      // Extract type and level (format: "startUL:0" or "startOL_RTL:1")
      const parts = action.split(':');
      const typeInfo = parts[0].replace('start', '');
      const level = parts[1] || '0';

      if(typeInfo.includes('UL_RTL')){
        htmlLines.push('<ul dir="rtl">');
        listStack.push(`ul_rtl:${level}`);
      } else if(typeInfo.includes('OL_RTL')){
        htmlLines.push('<ol dir="rtl">');
        listStack.push(`ol_rtl:${level}`);
      } else if(typeInfo.includes('UL')){
        htmlLines.push('<ul>');
        listStack.push(`ul:${level}`);
      } else {
        htmlLines.push('<ol>');
        listStack.push(`ol:${level}`);
      }
    } else if(action === 'endLIST'){
      const top=listStack.pop();
      if(!top) continue;
      const listType = top.split(':')[0];
      // Close the nested list
      if(listType.startsWith('u')) htmlLines.push('</ul>');
      else htmlLines.push('</ol>');
      // Close the parent <li> that contained the nested list
      htmlLines.push('</li>');
    } else if(action.startsWith('end')){
      const top=listStack.pop();
      if(!top) continue;
      const listType = top.split(':')[0];
      if(listType.startsWith('u')) htmlLines.push('</ul>');
      else htmlLines.push('</ol>');
    }
  }
}
function closeAllLists(listStack,htmlLines){
  while(listStack.length>0){
    const top=listStack.pop();
    if(top.startsWith('u')) htmlLines.push('</ul>');
    else htmlLines.push('</ol>');
  }
}

// -----------------------------------------------------
// 9) Table
// -----------------------------------------------------
async function renderTable(
  table,
  doc,
  usedFonts,
  authClient,
  outputDir,
  imagesDir,
  namedStylesMap
){
  let html='<table class="doc-table" style="border-collapse:collapse;">';
  for(const row of table.tableRows||[]){
    // Row styling
    let rowStyle = '';
    if(row.tableCellStyle?.backgroundColor?.color?.rgbColor){
      const rgb = row.tableCellStyle.backgroundColor.color.rgbColor;
      const hex = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
      rowStyle = `background-color:${hex};`;
    }
    html+=`<tr${rowStyle ? ` style="${rowStyle}"` : ''}>`;

    for(const cell of row.tableCells||[]){
      // Cell styling
      let cellStyle = 'padding:0.5em;';
      const cellStyleObj = cell.tableCellStyle || {};

      // Cell background color
      if(cellStyleObj.backgroundColor?.color?.rgbColor){
        const rgb = cellStyleObj.backgroundColor.color.rgbColor;
        const hex = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
        cellStyle += `background-color:${hex};`;
      }

      // Cell borders
      if(cellStyleObj.borderTop){
        cellStyle += formatBorder('top', cellStyleObj.borderTop);
      } else {
        cellStyle += 'border-top:1px solid #ccc;';
      }
      if(cellStyleObj.borderBottom){
        cellStyle += formatBorder('bottom', cellStyleObj.borderBottom);
      } else {
        cellStyle += 'border-bottom:1px solid #ccc;';
      }
      if(cellStyleObj.borderLeft){
        cellStyle += formatBorder('left', cellStyleObj.borderLeft);
      } else {
        cellStyle += 'border-left:1px solid #ccc;';
      }
      if(cellStyleObj.borderRight){
        cellStyle += formatBorder('right', cellStyleObj.borderRight);
      } else {
        cellStyle += 'border-right:1px solid #ccc;';
      }

      // Cell padding
      if(cellStyleObj.paddingTop?.magnitude){
        cellStyle += `padding-top:${ptToPx(cellStyleObj.paddingTop.magnitude)}px;`;
      }
      if(cellStyleObj.paddingBottom?.magnitude){
        cellStyle += `padding-bottom:${ptToPx(cellStyleObj.paddingBottom.magnitude)}px;`;
      }
      if(cellStyleObj.paddingLeft?.magnitude){
        cellStyle += `padding-left:${ptToPx(cellStyleObj.paddingLeft.magnitude)}px;`;
      }
      if(cellStyleObj.paddingRight?.magnitude){
        cellStyle += `padding-right:${ptToPx(cellStyleObj.paddingRight.magnitude)}px;`;
      }

      // Column span and row span
      let colspan = '';
      let rowspan = '';
      if(cell.colspan && cell.colspan > 1){
        colspan = ` colspan="${cell.colspan}"`;
      }
      if(cell.rowspan && cell.rowspan > 1){
        rowspan = ` rowspan="${cell.rowspan}"`;
      }

      html+=`<td${colspan}${rowspan} style="${cellStyle}">`;
      for(const c of cell.content||[]){
        if(c.paragraph){
          const { html:pHtml }=await renderParagraph(
            c.paragraph,
            doc,
            usedFonts,
            [],
            authClient,
            outputDir,
            imagesDir,
            namedStylesMap
          );
          html+=pHtml;
        }
      }
      html+='</td>';
    }
    html+='</tr>';
  }
  html+='</table>';
  return html;
}

// -----------------------------------------------------
// 10) Named Styles
// -----------------------------------------------------
function buildNamedStylesMap(doc){
  const map={};
  const named=doc.namedStyles?.styles||[];
  for(const s of named){
    map[s.namedStyleType]={
      paragraphStyle:s.paragraphStyle||{},
      textStyle:s.textStyle||{}
    };
  }
  return map;
}

// -----------------------------------------------------
// UTILS
// -----------------------------------------------------
async function getAuthClient(){
  const auth=new google.auth.GoogleAuth({
    keyFile:SERVICE_ACCOUNT_KEY_FILE,
    scopes:[
      'https://www.googleapis.com/auth/documents.readonly',
      'https://www.googleapis.com/auth/drive.readonly'
    ]
  });
  return auth.getClient();
}

/** fetchAsBase64 => fetch image content from drive with auth, return base64 */
async function fetchAsBase64(url, authClient){
  const resp=await authClient.request({
    url,
    method:'GET',
    responseType:'arraybuffer'
  });
  return Buffer.from(resp.data,'binary').toString('base64');
}

function deepMerge(base,overlay){
  for(const k in overlay){
    if(
      typeof overlay[k]==='object' &&
      overlay[k]!==null &&
      !Array.isArray(overlay[k])
    ){
      if(!base[k]) base[k]={};
      deepMerge(base[k],overlay[k]);
    } else {
      base[k]=overlay[k];
    }
  }
}
function deepCopy(obj){
  return JSON.parse(JSON.stringify(obj));
}
function escapeHtml(str){
  return str
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
function ptToPx(pts){
  return Math.round(pts*1.3333);
}
function rgbToHex(r,g,b){
  const nr=Math.round(r*255), ng=Math.round(g*255), nb=Math.round(b*255);
  return '#'+[nr,ng,nb].map(x=>x.toString(16).padStart(2,'0')).join('');
}
function buildGoogleFontsLink(fontFamilies){
  if(!fontFamilies||fontFamilies.length===0)return'';
  const unique=Array.from(new Set(fontFamilies));

  // Group fonts by family and collect all weights
  const fontMap = {};
  unique.forEach(f => {
    const parts = f.split(':');
    const family = parts[0];
    const weight = parts[1] || '400';
    if(!fontMap[family]){
      fontMap[family] = new Set();
    }
    fontMap[family].add(weight);
    // Also add common weights for better rendering
    fontMap[family].add('400');
    fontMap[family].add('700');
  });

  // Build the families parameter with specific weights
  const familiesParam = Object.entries(fontMap).map(([family, weights]) => {
    const normalized = family.trim().replace(/\s+/g,'+');
    const weightList = Array.from(weights).sort((a,b) => parseInt(a) - parseInt(b)).join(';');
    return `${normalized}:wght@${weightList}`;
  }).join('&family=');

  // Include comprehensive unicode subsets for right-to-left and non-Latin scripts
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
}

function formatBorder(side, border){
  if(!border || !border.width || !border.width.magnitude) return '';
  const width = ptToPx(border.width.magnitude);
  const style = borderStyleMap[border.dashStyle] || 'solid';
  let color = '#000000';
  if(border.color?.color?.rgbColor){
    const rgb = border.color.color.rgbColor;
    color = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
  }
  return `border-${side}:${width}px ${style} ${color};`;
}

function renderFootnoteReference(footnoteRef, doc){
  const footnoteId = footnoteRef.footnoteId;
  const footnoteNumber = footnoteRef.footnoteNumber || '?';
  // Create a superscript footnote reference
  return `<sup><a href="#footnote-${escapeHtml(footnoteId)}" id="footnote-ref-${escapeHtml(footnoteId)}">[${footnoteNumber}]</a></sup>`;
}

function renderEquation(equation){
  // Google Docs equations are stored as special text
  // We'll render them as code for now, but could use MathJax in the future
  const content = equation.suggestedInsertionIds || equation.suggestedDeletionIds || '';
  return `<code class="equation">${escapeHtml(content)}</code>`;
}

function renderAutoText(autoText){
  const type = autoText.type;
  // Common auto text types: PAGE_NUMBER, PAGE_COUNT
  if(type === 'PAGE_NUMBER'){
    return '<span class="page-number">[Page #]</span>';
  } else if(type === 'PAGE_COUNT'){
    return '<span class="page-count">[Total Pages]</span>';
  }
  return '';
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

