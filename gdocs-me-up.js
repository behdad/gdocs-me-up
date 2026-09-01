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
const { StyleRegistry, joinClasses } = require('./lib/styles');
const { writeOptimizedImage } = require('./lib/images');

// ------------- CONFIG -------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
// ------------- CONSTANTS -------------
// Indentation thresholds for inferring nesting levels (in points)
const INDENT_LEVEL_0_MAX = 40;  // Items with indent <= 40pt are level 0
const INDENT_LEVEL_1_MIN = 60;  // Items with indent >= 60pt are level 1

// List glyph types
const NUMBERED_GLYPH_TYPES = [
  'DECIMAL', 'ALPHA', 'ROMAN',
  'UPPER_ALPHA', 'UPPER_ROMAN',
  'LOWER_ALPHA', 'LOWER_ROMAN'
];

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

// Bullet style mappings
const BULLET_STYLE_MAP = {
  '●': 'disc',
  '○': 'circle',
  '■': 'square',
  '-': 'dash'
};

// Number style mappings
const NUMBER_STYLE_MAP = {
  'DECIMAL': 'decimal',
  'UPPER_ALPHA': 'upper-alpha',
  'LOWER_ALPHA': 'lower-alpha',
  'ALPHA': 'lower-alpha',
  'UPPER_ROMAN': 'upper-roman',
  'ROMAN': 'upper-roman',
  'LOWER_ROMAN': 'lower-roman'
};

// ------------- HELPER FUNCTIONS -------------

/**
 * Infer nesting level when not explicitly provided by the API.
 * Uses indentation as a signal to determine the correct level.
 *
 * @param {object} bullet - The bullet object from the paragraph
 * @param {object} paragraphStyle - The paragraph style containing indentStart
 * @param {number} prevLevel - The previous item's nesting level
 * @returns {number} The inferred nesting level
 */
function inferNestingLevel(bullet, paragraphStyle, prevLevel) {
  if (bullet?.nestingLevel !== undefined) {
    return bullet.nestingLevel;
  }

  const indentStart = paragraphStyle?.indentStart?.magnitude || 0;

  // Use indentation heuristics
  if (indentStart <= INDENT_LEVEL_0_MAX) {
    return 0;
  } else if (indentStart >= INDENT_LEVEL_1_MIN) {
    return 1;
  }

  // Fallback: continue at previous level if in a list
  return (prevLevel >= 0) ? prevLevel : 0;
}

/**
 * Safely get a value with a default fallback.
 *
 * @param {*} value - The value to check
 * @param {*} defaultValue - The default value to return if value is null/undefined
 * @returns {*} The value or default
 */
function getOrDefault(value, defaultValue) {
  return (value !== undefined && value !== null) ? value : defaultValue;
}

/**
 * Detect the bullet style for an unordered list.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @returns {string} The bullet style name (disc, circle, square, dash)
 */
function detectBulletStyle(glyph) {
  if (!glyph?.glyphSymbol) return 'disc';
  return BULLET_STYLE_MAP[glyph.glyphSymbol] || 'disc';
}

/**
 * Detect the numbering style for an ordered list.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @returns {string} The numbering style name (decimal, upper-alpha, etc.)
 */
function detectNumberStyle(glyph) {
  if (!glyph?.glyphType) return 'decimal';
  return NUMBER_STYLE_MAP[glyph.glyphType] || 'decimal';
}

/**
 * Determine if a list is numbered based on glyph properties and item counts.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @param {string} listId - The list identifier
 * @param {number} nestingLevel - The nesting level
 * @param {object} listItemCounts - The item count map
 * @returns {boolean} True if the list is numbered
 */
function isNumberedList(glyph, listId, nestingLevel, listItemCounts) {
  // If glyphSymbol is present (●, ○, -, etc.) → bullet list
  const isBullet = glyph?.glyphSymbol !== undefined;
  if (isBullet) return false;

  // If glyphType is explicitly a numbered type → numbered list
  const explicitlyNumbered = NUMBERED_GLYPH_TYPES.includes(glyph?.glyphType);
  if (explicitlyNumbered) return true;

  // If GLYPH_TYPE_UNSPECIFIED: use item count heuristic
  // Single-item lists → numbered (section markers)
  // Multi-item lists → bullets
  if (glyph?.glyphType === 'GLYPH_TYPE_UNSPECIFIED') {
    const key = `${listId}:${nestingLevel}`;
    const itemCount = listItemCounts?.[key] || 1;
    return (itemCount === 1);
  }

  // Default to bullet list
  return false;
}

async function exportDocToHTML(docId, outputDir) {
  try {
    // Validate inputs
    if (!docId || typeof docId !== 'string') {
      throw new Error('Invalid document ID provided');
    }
    if (!outputDir || typeof outputDir !== 'string') {
      throw new Error('Invalid output directory provided');
    }

    // Create output directories
    fs.mkdirSync(outputDir, { recursive: true });
    const imagesDir = path.join(outputDir, 'images');
    fs.mkdirSync(imagesDir, { recursive: true });

    // Auth & fetch doc
    const authClient = await getAuthClient();
    const docs = google.docs({ version: 'v1', auth: authClient });
    const { data: doc } = await docs.documents.get({ documentId: docId });

    if (!doc) {
      throw new Error('Failed to fetch document from Google Docs API');
    }

    console.log(`Exporting doc: ${doc.title || 'Untitled'}`);

  // Named styles for Title, Subtitle, Headings, etc.
  const namedStylesMap = buildNamedStylesMap(doc);

  // Container width from docStyle
  const colInfo = computeDocContainerWidth(doc);

  // Build global CSS
  const globalCSS = generateGlobalCSS(doc, colInfo);

  const usedFonts = new Set();
  const styleRegistry = new StyleRegistry();
  const documentLanguage = detectDocumentLanguage(doc);
  let htmlLines = [];

  // Basic HTML skeleton
  htmlLines.push('<!DOCTYPE html>');
  htmlLines.push(`<html lang="${documentLanguage.lang}" dir="${documentLanguage.direction}">`);
  htmlLines.push('<head>');
  htmlLines.push('  <meta charset="UTF-8">');
  htmlLines.push('  <meta name="viewport" content="width=device-width">');
  htmlLines.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlLines.push('  <style>');
  htmlLines.push(globalCSS);
  htmlLines.push('/* Document-specific styles (deduplicated from the generated markup). */');
  htmlLines.push('__DOCUMENT_STYLES__');
  htmlLines.push('  </style>');
  htmlLines.push('</head>');
  htmlLines.push('<body>');
  htmlLines.push('<div class="doc-content">');

  // Pre-process positioned objects (fetch and prepare HTML, but don't insert yet)
  const positionedObjectsHTML = new Map();
  if (doc.positionedObjects) {
    for (const [objId, posObj] of Object.entries(doc.positionedObjects)) {
      try {
        const props = posObj.positionedObjectProperties;
        if (!props) continue;

        const embedded = props.embeddedObject;
        if (!embedded?.imageProperties) continue;

        const { contentUri } = embedded.imageProperties;
        if (!contentUri) continue;

        // Fetch and save the image
        const base64Data = await fetchAsBase64(contentUri, authClient);
        if (!base64Data) {
          console.warn(`Failed to fetch positioned image ${objId}`);
          continue;
        }

        const buffer = Buffer.from(base64Data, 'base64');
        const { filePath } = await writeOptimizedImage(buffer, imagesDir, `positioned_${objId}`);

        const imgSrc = path.relative(outputDir, filePath);

        // Build styles based on positioning properties and size
        let style = 'max-width:100%;height:auto;display:block;';

        // Check for size information (prefer imageProperties.size, fallback to embedded.size)
        const size = embedded.imageProperties.size || embedded.size;
        if (size?.width?.magnitude && size?.height?.magnitude) {
          const scaleX = embedded.transform?.scaleX || 1;
          const scaleY = embedded.transform?.scaleY || 1;
          const wPx = Math.round(size.width.magnitude * 1.3333 * scaleX);
          const hPx = Math.round(size.height.magnitude * 1.3333 * scaleY);
          style += `width:${wPx}px;height:${hPx}px;`;
        }

        for (const [property, cssProperty] of [
          ['marginTop', 'margin-top'],
          ['marginRight', 'margin-right'],
          ['marginLeft', 'margin-left']
        ]) {
          if (embedded[property]?.magnitude) {
            style += `${cssProperty}:${ptToPx(embedded[property].magnitude)}px;`;
          }
        }

        // Record the layout without inventing generic margins; Docs supplies exact ones.
        const positioning = props.positioning;
        if (positioning?.layout === 'WRAP_TEXT') {
          style += 'shape-outside:margin-box;';
        } else if (positioning?.layout === 'BREAK_BOTH' || positioning?.layout === 'BREAK_LEFT' || positioning?.layout === 'BREAK_RIGHT') {
          style += 'margin-inline:auto;';
        }

        const alt = embedded.title || embedded.description || '';
        const imageClass = styleRegistry.add('i', style);
        const offset = positioning?.leftOffset?.magnitude || 0;
        const topOffset = positioning?.topOffset?.magnitude || 0;
        let wrapperStyle = '';
        if (offset) wrapperStyle += `margin-left:${ptToPx(offset)}px;`;
        if (topOffset) wrapperStyle += `margin-top:${ptToPx(topOffset)}px;`;
        const wrapperClass = joinClasses('positioned-image', styleRegistry.add('o', wrapperStyle));
        positionedObjectsHTML.set(
          objId,
          `<figure class="${wrapperClass}"><img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}"${imageClass ? ` class="${imageClass}"` : ''}></figure>`
        );
      } catch (error) {
        console.error(`Error rendering positioned object ${objId}:`, error.message);
      }
    }
  }

  // Pre-process: count items per list per level to determine if single-item (numbered) or multi-item (bullets)
  const listItemCounts = {};
  const bodyContent = doc.body?.content || [];
  doc.___footnoteNumbers = buildFootnoteNumberMap(bodyContent);
  let prevLevel = -1;

  for (const element of bodyContent) {
    if (element.paragraph?.bullet) {
      const bullet = element.paragraph.bullet;
      const level = inferNestingLevel(bullet, element.paragraph.paragraphStyle, prevLevel);
      const listId = bullet.listId;
      const key = `${listId}:${level}`;
      listItemCounts[key] = (listItemCounts[key] || 0) + 1;
      prevLevel = level;
    } else {
      prevLevel = -1;
    }
  }

  // Store counts in doc object for use by detectListChange
  doc.___listItemCounts = listItemCounts;

  let listStack = [];  // Array of {type: 'ul'/'ol', level: 0/1/2...}
  let prevNestingLevel = -1;
  let prevListId = null;  // Track previous listId to detect list changes
  const insertedPositionedObjects = new Set();

  for (const element of bodyContent) {
    if (element.sectionBreak) {
      if(prevNestingLevel>=0){ htmlLines.push('</li>'); prevNestingLevel=-1; prevListId=null; }
      closeAllLists(listStack, htmlLines);
      const sb = element.sectionBreak;
      const sectionStyle = sb.sectionStyle;

      // Handle column breaks
      if(sectionStyle?.columnSeparatorStyle === 'BETWEEN_EACH_COLUMN'){
        htmlLines.push('<div class="column-break"></div>');
      } else if(sectionStyle?.columnProperties && sectionStyle.columnProperties.length > 1){
        // Multi-column section
        const colCount = sectionStyle.columnProperties.length;
        const columnClass=styleRegistry.add('s', `column-count:${colCount};`);
        htmlLines.push(`<div class="${joinClasses('multi-column', columnClass)}"></div>`);
      } else {
        htmlLines.push('<div class="section-break"></div>');
      }
      continue;
    }
    if (element.tableOfContents) {
      if(prevNestingLevel>=0){ htmlLines.push('</li>'); prevNestingLevel=-1; prevListId=null; }
      closeAllLists(listStack, htmlLines);
      const tocHtml = await renderTableOfContents(
        element.tableOfContents,
        doc,
        usedFonts,
        authClient,
        outputDir,
        namedStylesMap,
        styleRegistry
      );
      htmlLines.push(tocHtml);
      continue;
    }
    if (element.paragraph) {
      // Infer nesting level using helper function
      const nestingLevel = element.paragraph.bullet
        ? inferNestingLevel(element.paragraph.bullet, element.paragraph.paragraphStyle, prevNestingLevel)
        : -1;

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

      try {
        const { html, listChange } = await renderParagraph(
          element.paragraph,
          doc,
          usedFonts,
          listStack,
          authClient,
          outputDir,
          imagesDir,
          namedStylesMap,
          styleRegistry
        );
        if (listChange) {
          handleListState(listChange, listStack, htmlLines, styleRegistry);
        }
        if (listStack.length > 0) {
          htmlLines.push(`<li>${html}`);
          prevNestingLevel = nestingLevel;
          prevListId = element.paragraph.bullet?.listId || null;
        } else {
          // Not in a list. The previous item was closed before rendering this
          // paragraph; only reset the state here (emitting another </li> corrupts
          // root lists when the list stack has just been closed).
          if (prevNestingLevel >= 0) {
            prevNestingLevel = -1;
          }
          prevListId = null;
          htmlLines.push(html);

          // Positioned objects are anchored to a paragraph in the Docs API. Emitting
          // them directly after that paragraph preserves cover layouts and text flow.
          for (const objectId of element.paragraph.positionedObjectIds || []) {
            const objectHtml = positionedObjectsHTML.get(objectId);
            if (objectHtml) {
              htmlLines.push(objectHtml);
              insertedPositionedObjects.add(objectId);
            }
          }
        }
      } catch (error) {
        console.error('Error rendering paragraph:', error.message);
        // Continue processing remaining elements
      }
      continue;
    }
    if (element.horizontalRule) {
      if(prevNestingLevel>=0){ htmlLines.push('</li>'); prevNestingLevel=-1; prevListId=null; }
      closeAllLists(listStack, htmlLines);
      htmlLines.push('<hr>');
      continue;
    }
    if (element.table) {
      if(prevNestingLevel>=0){ htmlLines.push('</li>'); prevNestingLevel=-1; prevListId=null; }
      closeAllLists(listStack, htmlLines);
      const tableHtml = await renderTable(
        element.table,
        doc,
        usedFonts,
        authClient,
        outputDir,
        imagesDir,
        namedStylesMap,
        styleRegistry
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

  // If positioned objects weren't inserted yet (no title/subtitle in doc), insert at end
  for (const [objectId, objectHtml] of positionedObjectsHTML) {
    if (!insertedPositionedObjects.has(objectId)) htmlLines.push(objectHtml);
  }

  const footnotesHtml = await renderFootnotes(
    doc,
    usedFonts,
    authClient,
    outputDir,
    imagesDir,
    namedStylesMap,
    styleRegistry
  );
  if (footnotesHtml) htmlLines.push(footnotesHtml);

  htmlLines.push('</div>');
  htmlLines.push('</body>');
  htmlLines.push('</html>');

  const documentStyles = styleRegistry.toCSS();
  const stylePlaceholder = htmlLines.indexOf('__DOCUMENT_STYLES__');
  if (stylePlaceholder >= 0) htmlLines[stylePlaceholder] = documentStyles;

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
  } catch (error) {
    console.error('Export failed:', error.message);
    if (error.code === 'ENOENT') {
      console.error('File not found. Check your service account key file path.');
    } else if (error.code === 'EACCES') {
      console.error('Permission denied. Check file/directory permissions.');
    } else if (error.response?.status === 404) {
      console.error('Document not found. Check the document ID and access permissions.');
    } else if (error.response?.status === 403) {
      console.error('Access forbidden. Ensure the service account has access to the document.');
    }
    throw error;
  }
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
/* Google Docs styles, rather than browser defaults, control block geometry. */
h1, h2, h3, h4, h5, h6, p, figure {
  margin: 0;
  font-size: 1em;
  font-weight: normal;
}

*, *::before, *::after { box-sizing: border-box; }
body {
  margin: 8px;
  font-family: sans-serif;
  /* Better font rendering for all scripts */
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-rendering: optimizeLegibility;
}
a { color: inherit; text-decoration: inherit; }
.doc-content {
  box-sizing: content-box;
  margin: 1em auto;
  max-width: ${containerPx}px;
  padding: 2em 1em;
}
p, h1, h2, h3, h4, h5, h6 {
  white-space: pre-wrap;
  overflow-wrap: break-word;
}
img {
  display: inline-block;
  max-width: 100%;
  height: auto;
  vertical-align: text-bottom;
}
sup {
  vertical-align: baseline;
  position: relative;
  top: -0.4em;
  line-height: 0;
  font-size: 0.8em;
}
sub {
  vertical-align: baseline;
  position: relative;
  top: 0.2em;
  line-height: 0;
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
hr {
  width: 100%;
  border: 0;
  border-top: 1px solid #ccc;
  margin: 1em 0;
}

.doc-toc {
  margin: 0.5em 0;
  padding: 0.5em;
}
.doc-toc p {
  margin: 0.1em 0;
  line-height: 1.2;
}

.subtitle { display: block; }
.positioned-image { width: 100%; }
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
.page-break { break-after: page; }
.column-break-inline { break-after: column; }
.footnotes {
  margin-top: 2em;
  font-size: 0.9em;
}
.footnotes > ol { padding-inline-start: 2em; }
.footnote-backref { margin-inline-start: 0.35em; text-decoration: none; }

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
  margin: 0;
  padding-left: 2em;
}
li {
  margin: 0;
}
.doc-content li > p {
  /* Docs list spacing belongs to the list line, not an extra block box. */
  margin-block: 0;
}
ul ul, ol ol, ul ol, ol ul {
  margin: 0;
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
  namedStylesMap,
  styleRegistry
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
        namedStylesMap,
        styleRegistry
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
  namedStylesMap,
  styleRegistry
) {
  if (!paragraph) {
    return { html: '', listChange: null };
  }

  const style = paragraph.paragraphStyle || {};
  const namedType = style.namedStyleType || 'NORMAL_TEXT';

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
  let listChange=null;
  if (paragraph.bullet) {
    // prevNestingLevel and prevListId are passed from the main loop
    const prevLevel = paragraph.___prevNestingLevel ?? -1;
    const prevListId = paragraph.___prevListId;
    listChange=detectListChange(paragraph,doc,listStack,isRTL,prevLevel,prevListId);
  } else {
    if(listStack.length>0){
      // Exiting all lists
      let actions = [];
      for(let i = listStack.length - 1; i >= 0; i--){
        actions.push(i === 0 ? 'endROOT' : 'endLIST');
      }
      listChange = actions.join('|');
    }
  }

  // Title => <h1 class="title">, Subtitle => <h2 class="subtitle">, heading => <hX>, else <p>
  let tag='p';
  const paragraphClasses=[];
  if(namedType==='TITLE'){
    tag='h1';
    paragraphClasses.push('title');
  } else if(namedType==='SUBTITLE'){
    tag='h2';
    paragraphClasses.push('subtitle');
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
  const metricTextStyle=deepCopy(mergedTextStyle);
  const firstTextRun=(paragraph.elements || []).find(element => element.textRun?.content);
  if(firstTextRun) deepMerge(metricTextStyle, firstTextRun.textRun.textStyle || {});
  const baseFontFamily=metricTextStyle.weightedFontFamily?.fontFamily;
  let lineHeightMultiplier=null;
  if(align && alignmentMapLTR[align]){
    inlineStyle += `text-align:${alignmentMapLTR[align]};`;
  }
  if(mergedParaStyle.lineSpacing){
    // Google Docs exposes lineSpacing as a percentage of the line box.
    lineHeightMultiplier=docsLineHeight(baseFontFamily, mergedParaStyle.lineSpacing);
    inlineStyle += `line-height:${lineHeightMultiplier};`;
  } else if(namedType === 'TITLE'){
    lineHeightMultiplier=docsTitleLineHeight(baseFontFamily);
    inlineStyle += `line-height:${lineHeightMultiplier};`;
  }
  if(mergedParaStyle.spaceAbove?.magnitude){
    inlineStyle += `margin-top:${ptToPx(mergedParaStyle.spaceAbove.magnitude)}px;`;
  }
  if(mergedParaStyle.spaceBelow?.magnitude){
    const spaceBelowPx=ptToPx(mergedParaStyle.spaceBelow.magnitude);
    // Docs' compact Title preset (10pt above rather than the usual 24pt)
    // retains four pixels of bottom leading outside the title line box.
    // A plain CSS margin loses that leading and shifts all following content.
    if(
      namedType === 'TITLE' &&
      mergedParaStyle.spaceAbove?.magnitude === 10 &&
      !(paragraph.positionedObjectIds || []).length
    ){
      inlineStyle += 'padding-bottom:4px;';
    }
    inlineStyle += `margin-bottom:${spaceBelowPx}px;`;
  }
  if(!paragraph.bullet){
    // Docs stores both values as absolute offsets from the page's start edge;
    // CSS text-indent is relative to the already-indented content box.
    const startIndent=mergedParaStyle.indentStart?.magnitude || 0;
    const firstLineIndent=mergedParaStyle.indentFirstLine?.magnitude;
    if(startIndent){
      inlineStyle += `margin-inline-start:${ptToPx(startIndent)}px;`;
    }
    if(firstLineIndent !== undefined){
      const firstLineDelta=ptToPx(firstLineIndent - startIndent);
      if(firstLineDelta) inlineStyle += `text-indent:${firstLineDelta}px;`;
    }
  }
  if(!paragraph.bullet && mergedParaStyle.indentEnd?.magnitude){
    inlineStyle += `margin-inline-end:${ptToPx(mergedParaStyle.indentEnd.magnitude)}px;`;
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
  // Only the named paragraph style is safe to inherit across the whole block.
  // The first run is useful for line metrics, but it may be a link or otherwise
  // specially formatted; promoting it would leak its color/decoration to siblings.
  inlineStyle += inheritedTextStyleCSS(mergedTextStyle, usedFonts);

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
  } else if(mergedParaStyle.direction==='LEFT_TO_RIGHT'){
    // This must be explicit inside an RTL document; otherwise the page direction
    // reorders neutral punctuation and makes LTR lines align from the wrong edge.
    dirAttr=' dir="ltr"';
  }

  // Merge text runs
  const mergedRuns=mergeTextRuns(paragraph.elements||[]);
  if(mergedRuns.length===1 && mergedRuns[0].horizontalRule){
    return { html:'<hr>', listChange };
  }
  let innerHtml='';
  for(const r of mergedRuns){
    if(r.inlineObjectElement){
      const objId=r.inlineObjectElement.inlineObjectId;
      innerHtml += await renderInlineObject(objId, doc, authClient, outputDir, imagesDir, styleRegistry);
    } else if(r.textRun){
      innerHtml += renderTextRun(
        r.textRun,
        usedFonts,
        mergedTextStyle,
        styleRegistry,
        mergedTextStyle,
        mergedParaStyle.lineSpacing
      );
    } else if(r.footnoteReference){
      innerHtml += renderFootnoteReference(r.footnoteReference, doc);
    } else if(r.equation){
      innerHtml += renderEquation(r.equation);
    } else if(r.autoText){
      innerHtml += renderAutoText(r.autoText);
    } else if(r.pageBreak){
      innerHtml += '<span class="page-break" aria-hidden="true"></span>';
    } else if(r.columnBreak){
      innerHtml += '<span class="column-break-inline" aria-hidden="true"></span>';
    } else if(r.horizontalRule){
      innerHtml += '<hr>';
    } else if(r.person){
      innerHtml += renderPerson(r.person);
    } else if(r.richLink){
      innerHtml += renderRichLink(r.richLink);
    }
  }
  if(!innerHtml && lineHeightMultiplier){
    const fontPoints=metricTextStyle.fontSize?.magnitude || 12;
    const lineBox=fontPoints * 1.3333 * lineHeightMultiplier;
    const positionedAnchorSpace=Object.keys(doc.positionedObjects || {}).length
      ? Math.max(
        ptToPx(mergedParaStyle.spaceAbove?.magnitude || 0),
        ptToPx(mergedParaStyle.spaceBelow?.magnitude || 0)
      )
      : 0;
    // A fractional line box can otherwise round down and make every intentional
    // blank paragraph slightly shorter than the corresponding Docs line. Paragraph
    // margins are separate spacing and must not be deducted from the blank line.
    // Positioned-object documents use the blank paragraph's surrounding space
    // as part of the object's anchor flow, so counting it twice shifts captions.
    inlineStyle += `min-height:${Math.max(0, Math.ceil(lineBox - positionedAnchorSpace))}px;`;
  }

  paragraphClasses.push(styleRegistry.add('p', inlineStyle));
  const className=joinClasses(paragraphClasses);
  const classAttr=className ? ` class="${className}"` : '';
  const paragraphHtml=`<${tag}${headingIdAttr}${dirAttr}${classAttr}>${innerHtml}</${tag}>`;

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
    } else if(e.pageBreak){
      merged.push({ pageBreak:e.pageBreak});
      last=null;
    } else if(e.columnBreak){
      merged.push({ columnBreak:e.columnBreak});
      last=null;
    } else if(e.horizontalRule){
      merged.push({ horizontalRule:e.horizontalRule});
      last=null;
    } else if(e.person){
      merged.push({ person:e.person});
      last=null;
    } else if(e.richLink){
      merged.push({ richLink:e.richLink});
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
function inheritedTextStyleCSS(style, usedFonts){
  let css='';
  if(style.bold) css+='font-weight:bold;';
  if(style.italic) css+='font-style:italic;';
  const decorations=[];
  if(style.underline) decorations.push('underline');
  if(style.strikethrough) decorations.push('line-through');
  if(decorations.length) css+=`text-decoration-line:${decorations.join(' ')};`;
  if(style.smallCaps) css+='font-variant:small-caps;';
  if(style.fontSize?.magnitude) css+=`font-size:${style.fontSize.magnitude}pt;`;
  if(style.weightedFontFamily?.fontFamily){
    const family=style.weightedFontFamily.fontFamily;
    const weight=style.weightedFontFamily.weight || 400;
    usedFonts.add(`${family}:${weight}`);
    css+=`font-family:'${family}',sans-serif;`;
    if(weight !== 400) css+=`font-weight:${weight};`;
  }
  if(style.foregroundColor?.color?.rgbColor){
    const rgb=style.foregroundColor.color.rgbColor;
    css+=`color:${rgbToHex(rgb.red||0,rgb.green||0,rgb.blue||0)};`;
  }
  if(style.backgroundColor?.color?.rgbColor){
    const rgb=style.backgroundColor.color.rgbColor;
    css+=`background-color:${rgbToHex(rgb.red||0,rgb.green||0,rgb.blue||0)};`;
  }
  return css;
}

function renderTextRun(textRun, usedFonts, baseStyle, styleRegistry, inheritedStyle, paragraphLineSpacing){
  const finalStyle=deepCopy(baseStyle||{});
  deepMerge(finalStyle, textRun.textStyle||{});
  const inherited=inheritedStyle || {};

  let content=textRun.content||'';
  // Remove trailing newline (marks end of paragraph)
  content=content.replace(/\n$/,'');
  // Convert vertical tabs (\u000b) to a placeholder
  // These represent soft line breaks within a paragraph (Shift+Enter in Google Docs)
  content=content.replace(/\u000b/g,'__LINEBREAK__');

  let inlineStyle='';

  // Small caps support
  if(finalStyle.smallCaps !== inherited.smallCaps){
    inlineStyle+=`font-variant:${finalStyle.smallCaps ? 'small-caps' : 'normal'};`;
  }

  if(JSON.stringify(finalStyle.fontSize||null)!==JSON.stringify(inherited.fontSize||null) && finalStyle.fontSize?.magnitude){
    inlineStyle+=`font-size:${finalStyle.fontSize.magnitude}pt;`;
  }
  if(JSON.stringify(finalStyle.weightedFontFamily||null)!==JSON.stringify(inherited.weightedFontFamily||null) && finalStyle.weightedFontFamily?.fontFamily){
    const fam=finalStyle.weightedFontFamily.fontFamily;
    const weight = finalStyle.weightedFontFamily.weight || 400;
    // Track font with its weight for better loading
    usedFonts.add(`${fam}:${weight}`);
    inlineStyle+=`font-family:'${fam}',sans-serif;`;
    // A font override can make the line box taller than the paragraph's base
    // font. Docs applies the paragraph's spacing percentage to that font's own
    // natural metrics; an inherited CSS line-height would keep the base font's
    // shorter box and pull every following paragraph upward.
    if(paragraphLineSpacing){
      inlineStyle+=`line-height:${docsLineHeight(fam, paragraphLineSpacing)};`;
    }
    // Font weight if specified
    if(weight && weight !== 400){
      inlineStyle+=`font-weight:${weight};`;
    }
  }
  if(JSON.stringify(finalStyle.foregroundColor||null)!==JSON.stringify(inherited.foregroundColor||null) && finalStyle.foregroundColor?.color?.rgbColor){
    const rgb=finalStyle.foregroundColor.color.rgbColor;
    const hex=rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle+=`color:${hex};`;
  }

  // Background color support
  if(JSON.stringify(finalStyle.backgroundColor||null)!==JSON.stringify(inherited.backgroundColor||null) && finalStyle.backgroundColor?.color?.rgbColor){
    const rgb=finalStyle.backgroundColor.color.rgbColor;
    const hex=rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle+=`background-color:${hex};`;
  }

  let linkHref='';
  if(finalStyle.link){
    if(finalStyle.link.headingId){
      linkHref=`#heading-${escapeHtml(finalStyle.link.headingId)}`;
    } else if(finalStyle.link.url){
      linkHref=finalStyle.link.url;
    }
  }

  // Replace line break placeholder with actual <br> tags after escaping
  let escapedContent = escapeHtml(content).replace(/__LINEBREAK__/g, '<br>');
  // A terminal <br> does not create a second line box in HTML, whereas a
  // trailing soft break does in Google Docs. The zero-width character keeps
  // that final blank line measurable without changing the visible content.
  if (escapedContent.endsWith('<br>')) escapedContent += '&#8203;';
  if (!escapedContent) return '';

  const semanticTags=[];
  if(finalStyle.bold && !inherited.bold) semanticTags.push('strong');
  if(finalStyle.italic && !inherited.italic) semanticTags.push('em');
  if(finalStyle.underline && !inherited.underline) semanticTags.push('u');
  if(finalStyle.strikethrough && !inherited.strikethrough) semanticTags.push('s');
  if(finalStyle.baselineOffset==='SUPERSCRIPT') semanticTags.push('sup');
  if(finalStyle.baselineOffset==='SUBSCRIPT') semanticTags.push('sub');

  if(inherited.bold && !finalStyle.bold) inlineStyle+='font-weight:normal;';
  if(inherited.italic && !finalStyle.italic) inlineStyle+='font-style:normal;';
  if((inherited.underline && !finalStyle.underline) || (inherited.strikethrough && !finalStyle.strikethrough)){
    const decorations=[];
    if(finalStyle.underline) decorations.push('underline');
    if(finalStyle.strikethrough) decorations.push('line-through');
    inlineStyle+=`text-decoration-line:${decorations.join(' ') || 'none'};`;
  }
  const textClass=styleRegistry.add('t', inlineStyle);

  let html=escapedContent;
  for(let i=semanticTags.length-1;i>=0;i--){
    const tag=semanticTags[i];
    const classAttr=(!linkHref && i===0 && textClass) ? ` class="${textClass}"` : '';
    html=`<${tag}${classAttr}>${html}</${tag}>`;
  }

  if(linkHref){
    const classAttr=textClass ? ` class="${textClass}"` : '';
    return `<a href="${escapeHtml(linkHref)}"${classAttr}>${html}</a>`;
  }
  if(semanticTags.length===0 && textClass){
    return `<span class="${textClass}">${html}</span>`;
  }
  return html;
}

// -----------------------------------------------------
// 7) Inline Objects (Images)
// -----------------------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir, styleRegistry){
  try {
    const inlineObj = doc.inlineObjects?.[objectId];
    if (!inlineObj) return '';

    const embedded = inlineObj.inlineObjectProperties?.embeddedObject;
    if (!embedded?.imageProperties) return '';

    const { imageProperties } = embedded;
    const { contentUri, cropProperties } = imageProperties;

    // Check both locations for size: prefer imageProperties.size, fallback to embedded.size
    const size = imageProperties.size || embedded.size;

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
    const { filePath } = await writeOptimizedImage(buffer, imagesDir, `image_${objectId}`);

    const imgSrc = path.relative(outputDir, filePath);

  // Always constrain images to container width and maintain aspect ratio
  let style='max-width:100%;height:auto;';
  if(size?.width?.magnitude && size?.height?.magnitude){
    const width=ptToPx(size.width.magnitude * scaleX);
    const height=ptToPx(size.height.magnitude * scaleY);
    style+=`width:${width}px;height:${height}px;`;
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
    style += `margin-top:${inlineImageVerticalMarginPx(embedded.marginTop.magnitude)}px;`;
  }
  if(embedded.marginBottom?.magnitude){
    style += `margin-bottom:${inlineImageVerticalMarginPx(embedded.marginBottom.magnitude)}px;`;
  }
  if(embedded.marginLeft?.magnitude){
    style += `margin-left:${ptToPx(embedded.marginLeft.magnitude)}px;`;
  }
  if(embedded.marginRight?.magnitude){
    style += `margin-right:${ptToPx(embedded.marginRight.magnitude)}px;`;
  }

    const alt = embedded.title || embedded.description || '';
    const imageClass=styleRegistry.add('i', style);
    return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}"${imageClass ? ` class="${imageClass}"` : ''}>`;
  } catch (error) {
    console.error(`Error rendering image ${objectId}:`, error.message);
    return `<!-- Image ${objectId} failed to render -->`;
  }
}

// -----------------------------------------------------
// 8) Lists
// -----------------------------------------------------
function detectListChange(paragraph, doc, listStack, isRTL, prevLevel, prevListId){
  const bullet = paragraph.bullet;
  if (!bullet) return null;

  const listId = bullet.listId;
  const nestingLevel = inferNestingLevel(bullet, paragraph.paragraphStyle, prevLevel);

  const listDef = doc.lists?.[listId];
  if (!listDef?.listProperties?.nestingLevels) return null;

  const glyph = listDef.listProperties.nestingLevels[nestingLevel];
  if (!glyph) return null;

  // Determine if this is a numbered or bullet list
  const isNumbered = isNumberedList(glyph, listId, nestingLevel, doc.___listItemCounts);

  // Detect the specific style
  const bulletStyle = isNumbered ? null : detectBulletStyle(glyph);
  const numberStyle = isNumbered ? detectNumberStyle(glyph) : null;

  // Build the list type identifier
  const startType = isNumbered ? 'OL' : 'UL';
  const rtlFlag = isRTL ? '_RTL' : '';
  const styleFlag = isNumbered
    ? (numberStyle !== 'decimal' ? `_${numberStyle.toUpperCase().replace(/-/g, '_')}` : '')
    : (bulletStyle !== 'disc' ? `_${bulletStyle.toUpperCase()}` : '');

  // Starting a list for the first time
  if(listStack.length === 0){
    return `start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
  }

  // Check if nesting level changed
  if(nestingLevel > prevLevel){
    // Going deeper - start nested list
    return `start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
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
      const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '') + styleFlag.toLowerCase();

      // Check if parent list type changed (not listId - Google Docs splits numbered lists)
      if(parentType !== wantType){
        actions.push(`end${parentType?.toUpperCase() || 'UL'}`);
        actions.push(`start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`);
      }
    }

    return actions.join('|');
  }

  // Same level - check if list type changed
  const currentType = listStack[listStack.length - 1]?.split(':')[0];
  const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '') + styleFlag.toLowerCase();

  // Only switch lists if the TYPE changed (OL vs UL)
  // Don't switch just because listId changed - Google Docs splits numbered lists across listIds
  if(currentType !== wantType){
    return `end${currentType?.toUpperCase() || 'UL'}|start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
  }

  return null;
}

/**
 * Handle list state transitions by processing list change actions.
 * Actions can be: startUL:0, startOL_RTL:1, startUL_DASH:0, endLIST, endUL, etc.
 *
 * @param {string} listChange - Pipe-separated list of actions
 * @param {Array} listStack - Stack tracking open lists
 * @param {Array} htmlLines - Array of HTML lines being built
 */
function handleListState(listChange, listStack, htmlLines, styleRegistry){
  if (!listChange) return;

  const actions = listChange.split('|');
  for (const action of actions) {
    if (action.startsWith('start')) {
      // Extract type and level (format: "startUL:0", "startOL_RTL:1", "startUL_DASH:0")
      const parts = action.split(':');
      const typeInfo = parts[0].replace('start', '');
      const level = parts[1] || '0';

      // Parse style flags (DASH, CIRCLE, SQUARE for UL; UPPER_ALPHA, LOWER_ROMAN, etc. for OL)
      let listStyle = '';
      if(typeInfo.includes('_DASH')){
        listStyle = "list-style-type:'− ';";
      } else if(typeInfo.includes('_CIRCLE')){
        listStyle = 'list-style-type:circle;';
      } else if(typeInfo.includes('_SQUARE')){
        listStyle = 'list-style-type:square;';
      } else if(typeInfo.includes('_UPPER_ALPHA')){
        listStyle = 'list-style-type:upper-alpha;';
      } else if(typeInfo.includes('_LOWER_ALPHA')){
        listStyle = 'list-style-type:lower-alpha;';
      } else if(typeInfo.includes('_UPPER_ROMAN')){
        listStyle = 'list-style-type:upper-roman;';
      } else if(typeInfo.includes('_LOWER_ROMAN')){
        listStyle = 'list-style-type:lower-roman;';
      }

      const listClass=styleRegistry.add('l', listStyle);
      const classAttr=listClass ? ` class="${listClass}"` : '';

      if(typeInfo.includes('UL_RTL') || (typeInfo.includes('UL') && typeInfo.includes('_RTL'))){
        htmlLines.push(`<ul dir="rtl"${classAttr}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else if(typeInfo.includes('OL_RTL')){
        htmlLines.push(`<ol dir="rtl"${classAttr}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else if(typeInfo.includes('UL')){
        htmlLines.push(`<ul${classAttr}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else {
        htmlLines.push(`<ol${classAttr}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
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
/**
 * Close all open lists in the stack.
 *
 * @param {Array} listStack - Stack of open lists
 * @param {Array} htmlLines - Array of HTML lines being built
 */
function closeAllLists(listStack, htmlLines){
  if (!listStack || !htmlLines) return;

  while (listStack.length > 0) {
    const top = listStack.pop();
    if (!top) continue;
    if (top.startsWith('u')) htmlLines.push('</ul>');
    else htmlLines.push('</ol>');
  }
}

// -----------------------------------------------------
// 9) Table
// -----------------------------------------------------
/**
 * Render a Google Docs table to HTML.
 *
 * @param {object} table - The table element from Google Docs
 * @param {object} doc - The full document object
 * @param {Set} usedFonts - Set to track used fonts
 * @param {object} authClient - Auth client for API calls
 * @param {string} outputDir - Output directory path
 * @param {string} imagesDir - Images directory path
 * @param {object} namedStylesMap - Map of named styles
 * @returns {Promise<string>} HTML string for the table
 */
async function renderTable(
  table,
  doc,
  usedFonts,
  authClient,
  outputDir,
  imagesDir,
  namedStylesMap,
  styleRegistry
){
  if (!table) return '';

  try {
    let html = '<table class="doc-table">';
    for (const row of table.tableRows || []) {
    // Row styling
    let rowStyle = '';
    if(row.tableCellStyle?.backgroundColor?.color?.rgbColor){
      const rgb = row.tableCellStyle.backgroundColor.color.rgbColor;
      const hex = rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
      rowStyle = `background-color:${hex};`;
    }
    if(row.tableRowStyle?.minRowHeight?.magnitude){
      rowStyle += `min-height:${ptToPx(row.tableRowStyle.minRowHeight.magnitude)}px;`;
    }
    const rowClass=styleRegistry.add('r', rowStyle);
    html+=`<tr${rowClass ? ` class="${rowClass}"` : ''}>`;

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
      const columnSpan=cellStyleObj.columnSpan || cell.colspan || 1;
      const rowSpan=cellStyleObj.rowSpan || cell.rowspan || 1;
      if(columnSpan > 1){
        colspan = ` colspan="${columnSpan}"`;
      }
      if(rowSpan > 1){
        rowspan = ` rowspan="${rowSpan}"`;
      }

      const cellClass=styleRegistry.add('c', cellStyle);
      html+=`<td${colspan}${rowspan}${cellClass ? ` class="${cellClass}"` : ''}>`;
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
            namedStylesMap,
            styleRegistry
          );
          html+=pHtml;
        }
      }
      html+='</td>';
    }
      html += '</tr>';
    }
    html += '</table>';
    return html;
  } catch (error) {
    console.error('Error rendering table:', error.message);
    return '<!-- Table rendering failed -->';
  }
}

// -----------------------------------------------------
// 10) Named Styles
// -----------------------------------------------------
function buildNamedStylesMap(doc){
  const map={};
  const named=doc.namedStyles?.styles||[];
  const normal=named.find(style => style.namedStyleType === 'NORMAL_TEXT');
  for(const s of named){
    // Headings are refinements of Normal text in Docs, but Title and Subtitle
    // have independent paragraph metrics. Inheriting a document's body line
    // spacing into its title can produce dramatically oversized title lines.
    const inheritsNormal=!['NORMAL_TEXT', 'TITLE', 'SUBTITLE'].includes(s.namedStyleType);
    const paragraphStyle=!inheritsNormal
      ? {}
      : deepCopy(normal?.paragraphStyle || {});
    const textStyle=!inheritsNormal
      ? {}
      : deepCopy(normal?.textStyle || {});
    deepMerge(paragraphStyle, s.paragraphStyle || {});
    deepMerge(textStyle, s.textStyle || {});
    map[s.namedStyleType]={
      paragraphStyle,
      textStyle
    };
  }
  return map;
}

function detectDocumentLanguage(doc){
  let text=doc.title || '';
  for(const element of doc.body?.content || []){
    for(const paragraphElement of element.paragraph?.elements || []){
      text += paragraphElement.textRun?.content || '';
    }
  }
  const rtlCount=(text.match(/[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff]/g) || []).length;
  const latinCount=(text.match(/[A-Za-z]/g) || []).length;
  return rtlCount > latinCount
    ? { lang:'fa', direction:'rtl' }
    : { lang:'en', direction:'ltr' };
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
  if (str === null || str === undefined) return '';
  const strValue = String(str);
  return strValue
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}
/**
 * Convert points to pixels (1pt ≈ 1.3333px).
 *
 * @param {number} pts - Points value
 * @returns {number} Pixels value
 */
function ptToPx(pts){
  if (typeof pts !== 'number' || isNaN(pts)) return 0;
  return Math.round(pts * 1.3333);
}

/**
 * Inline objects expose the same wrap-clearance margins as floating objects,
 * but Docs only uses one third of that clearance in the inline line box.
 */
function inlineImageVerticalMarginPx(pts){
  return Math.round(ptToPx(pts) / 3);
}

/**
 * Docs applies line-spacing percentages to each font's natural line metrics. CSS
 * unitless line-height instead multiplies only the font size, so script fonts with
 * tall ascenders/descenders need their metric ratio restored explicitly.
 */
function docsLineHeight(fontFamily, spacingPercent){
  if(spacingPercent === 100) return 1;
  if(fontFamily === 'Noto Naskh Arabic') return 1.6875 * spacingPercent / 100;
  // Display fonts have substantially taller natural line boxes than PT Sans.
  // Restore those metrics before applying Docs' percentage line spacing.
  if(fontFamily === 'Anton') return 1.49565 * spacingPercent / 100;
  if(fontFamily === 'Lalezar') return 1.41304 * spacingPercent / 100;
  // Google Docs' Latin line box is taller than CSS `font-size`, but additional
  // line spacing grows more slowly than a direct percentage multiplication.
  return 1.4185 + (spacingPercent - 100) * 0.00543;
}

function docsTitleLineHeight(fontFamily){
  if(fontFamily === 'Anton') return 1.72;
  if(fontFamily === 'Lalezar') return 1.625;
  return 1.5;
}

/**
 * Convert RGB values (0-1 range) to hex color.
 *
 * @param {number} r - Red (0-1)
 * @param {number} g - Green (0-1)
 * @param {number} b - Blue (0-1)
 * @returns {string} Hex color string
 */
function rgbToHex(r, g, b){
  const clamp = (val) => Math.max(0, Math.min(1, val || 0));
  const nr = Math.round(clamp(r) * 255);
  const ng = Math.round(clamp(g) * 255);
  const nb = Math.round(clamp(b) * 255);
  return '#' + [nr, ng, nb].map(x => x.toString(16).padStart(2, '0')).join('');
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

function buildFootnoteNumberMap(content){
  const numbers={};
  let next=1;
  const visit=value=>{
    if(!value || typeof value!=='object') return;
    if(value.footnoteReference?.footnoteId){
      const id=value.footnoteReference.footnoteId;
      if(!numbers[id]) numbers[id]=next++;
    }
    for(const child of Object.values(value)) visit(child);
  };
  visit(content);
  return numbers;
}

async function renderFootnotes(doc, usedFonts, authClient, outputDir, imagesDir, namedStylesMap, styleRegistry){
  const entries=Object.entries(doc.footnotes || {});
  if(entries.length===0) return '';

  entries.sort(([left],[right]) =>
    (doc.___footnoteNumbers?.[left] || Infinity) - (doc.___footnoteNumbers?.[right] || Infinity)
  );
  let html='<section class="footnotes" aria-label="Footnotes"><hr><ol>';
  for(const [footnoteId, footnote] of entries){
    html+=`<li id="footnote-${escapeHtml(footnoteId)}">`;
    for(const element of footnote.content || []){
      if(element.paragraph){
        const { html:paragraphHtml }=await renderParagraph(
          element.paragraph,
          doc,
          usedFonts,
          [],
          authClient,
          outputDir,
          imagesDir,
          namedStylesMap,
          styleRegistry
        );
        html+=paragraphHtml;
      } else if(element.table){
        html+=await renderTable(
          element.table,
          doc,
          usedFonts,
          authClient,
          outputDir,
          imagesDir,
          namedStylesMap,
          styleRegistry
        );
      }
    }
    html+=`<a class="footnote-backref" href="#footnote-ref-${escapeHtml(footnoteId)}" aria-label="Back to reference">↩</a></li>`;
  }
  return html+'</ol></section>';
}

function renderFootnoteReference(footnoteRef, doc){
  const footnoteId = footnoteRef.footnoteId;
  const footnoteNumber = footnoteRef.footnoteNumber || doc.___footnoteNumbers?.[footnoteId] || '?';
  return `<sup><a href="#footnote-${escapeHtml(footnoteId)}" id="footnote-ref-${escapeHtml(footnoteId)}">${footnoteNumber}</a></sup>`;
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

function renderPerson(person){
  const properties=person.personProperties || {};
  const label=properties.name || properties.email || '';
  if(!label) return '';
  return properties.email
    ? `<a href="mailto:${escapeHtml(properties.email)}">${escapeHtml(label)}</a>`
    : escapeHtml(label);
}

function renderRichLink(richLink){
  const properties=richLink.richLinkProperties || {};
  const url=properties.uri || properties.url || '';
  const label=properties.title || url;
  return url
    ? `<a href="${escapeHtml(url)}">${escapeHtml(label)}</a>`
    : escapeHtml(label);
}

module.exports = {
  exportDocToHTML,
  generateGlobalCSS,
  renderParagraph,
  inlineImageVerticalMarginPx,
  buildNamedStylesMap
};

// CLI
if(require.main===module){
  const docId=process.argv[2];
  const outDir=process.argv[3];

  // Show help if requested
  if (docId === '--help' || docId === '-h') {
    console.log(`
Google Docs High-Fidelity HTML Exporter

Usage:
  node gdocs-me-up.js <DOC_ID> <OUTPUT_DIR>

Arguments:
  DOC_ID      Google Docs document ID (from the URL)
  OUTPUT_DIR  Directory where HTML and images will be exported

Example:
  node gdocs-me-up.js 1AbCdEfgHIjKLMnOP ./output

The exported HTML will be saved as OUTPUT_DIR/index.html
Images will be saved in OUTPUT_DIR/images/

Requirements:
  - service_account.json file with Google Docs API access
  - Document must be accessible by the service account
    `);
    process.exit(0);
  }

  if(!docId||!outDir){
    console.error('Error: Missing required arguments\n');
    console.error('Usage: node gdocs-me-up.js <DOC_ID> <OUTPUT_DIR>');
    console.error('Run with --help for more information');
    process.exit(1);
  }

  exportDocToHTML(docId, outDir).catch(err=>{
    console.error('Export error:',err.message || err);
    process.exit(1);
  });
}
