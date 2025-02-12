/**
 * google-docs-high-fidelity-export.js
 *
 * High-Fidelity Exporter for Google Docs
 *
 * Features:
 *  - **Named Styles**: Title, Subtitle, Headings (H1..H6) mapped to real HTML headings,
 *    with doc-based font sizing and inline styling (bold, italic, color, etc.).
 *  - **Heading Size Override**: Resets default <h1>.. <h6> to neutral size in CSS,
 *    so doc’s inline size precisely controls final heading font.
 *  - **Line Spacing**: Honors doc’s paragraph lineSpacing (e.g., 1.15, 1.5),
 *    plus spaceAbove/spaceBelow, and text indentation.
 *  - **Right-to-Left Paragraphs**: Sets `dir="rtl"` for paragraphs whose style
 *    indicates RIGHT_TO_LEFT, flipping alignment START/END if needed.
 *  - **Alignment**: PRESERVES doc-based alignment (CENTER, JUSTIFIED, etc.).
 *  - **Lists**: Bullet / Numbered lists, including RTL bullets if direction=RIGHT_TO_LEFT.
 *  - **Images**: Exact doc-based sizes, with transform scaling if used in the doc.
 *    Exports images to an `images/` folder. Uses `max-width` / `max-height` so they
 *    never exceed doc’s reported size or the container width.
 *  - **Table of Contents**: Indents each TOC entry based on the heading level of its
 *    linked heading (Heading 1 => level 1, etc.).
 *  - **Tables**: Exports Google Docs tables using <table>, <tr>, <td>,
 *    preserving paragraph styles inside each cell.
 *  - **Column Width**: Infers container width from doc’s pageSize minus margins
 *    (with a small tweak). Then sets `.doc-content { max-width: ... }`.
 *  - **Google Fonts**: Gathers all distinct fonts used, generating a <link> to
 *    https://fonts.googleapis.com for a near-pixel match to the doc’s text families.
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
 * justification, bullet-lists, images, and more are also handled for a truly “high-fidelity”
 * offline representation of your Google Doc.
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
  htmlLines.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlLines.push('  <style>');
  htmlLines.push(globalCSS);
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
  margin: 0;
  font-family: sans-serif;
}
.doc-content {
  margin: 1em auto;
  max-width: ${containerPx}px;
  padding: 2em 1em;
}
p, li {
  margin: 0.5em 0;
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
}

.doc-subtitle {
  display: block;
  white-space: pre-wrap;
}
.doc-table {
  border-collapse: collapse;
  margin: 0.5em 0;
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
  let listChange=null;
  if (paragraph.bullet) {
    listChange=detectListChange(paragraph.bullet,doc,listStack,isRTL);
  } else {
    if(listStack.length>0){
      const top=listStack[listStack.length-1];
      listChange=`end${top.toUpperCase()}`;
    }
  }

  // Title => <h1>, Subtitle => <h2 class="doc-subtitle">, heading => <hX>, else <p>
  let tag='p';
  if(namedType==='TITLE'){
    tag='h1';
  } else if(namedType==='SUBTITLE'){
    tag='h2 class="doc-subtitle"';
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
    // browsers default to 1.2; empirically 1.25 matches docs
    const ls=mergedParaStyle.lineSpacing*1.25/100;
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
    'foregroundColor','link'
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
  content=content.replace(/\n$/,'');

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
  if(finalStyle.fontSize?.magnitude){
    inlineStyle+=`font-size:${finalStyle.fontSize.magnitude}pt;`;
  }
  if(finalStyle.weightedFontFamily?.fontFamily){
    const fam=finalStyle.weightedFontFamily.fontFamily;
    usedFonts.add(fam);
    inlineStyle+=`font-family:'${fam}',sans-serif;`;
  }
  if(finalStyle.foregroundColor?.color?.rgbColor){
    const rgb=finalStyle.foregroundColor.color.rgbColor;
    const hex=rgbToHex(rgb.red||0, rgb.green||0, rgb.blue||0);
    inlineStyle+=`color:${hex};`;
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

  return openTag+escapeHtml(content)+closeTag;
}

// -----------------------------------------------------
// 7) Inline Objects (Images)
// -----------------------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir){
  const inlineObj=doc.inlineObjects?.[objectId];
  if(!inlineObj) return'';

  const embedded=inlineObj.inlineObjectProperties?.embeddedObject;
  if(!embedded?.imageProperties) return'';

  const { imageProperties }=embedded;
  const { contentUri, size }=imageProperties;

  let scaleX=1, scaleY=1;
  if(embedded.transform){
    if(embedded.transform.scaleX) scaleX=embedded.transform.scaleX;
    if(embedded.transform.scaleY) scaleY=embedded.transform.scaleY;
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
  const alt=embedded.title||embedded.description||'';
  return `<img src="${escapeHtml(imgSrc)}" alt="${escapeHtml(alt)}" style="${style}" />`;
}

// -----------------------------------------------------
// 8) Lists
// -----------------------------------------------------
function detectListChange(bullet, doc, listStack, isRTL){
  const listId=bullet.listId;
  const nestingLevel=bullet.nestingLevel||0;
  const listDef=doc.lists?.[listId];
  if(!listDef?.listProperties?.nestingLevels)return null;

  const glyph=listDef.listProperties.nestingLevels[nestingLevel];
  const isNumbered=glyph?.glyphType?.toLowerCase().includes('number');
  const top=listStack[listStack.length-1];

  const startType=isNumbered?'OL':'UL';
  const rtlFlag=isRTL?'_RTL':'';

  if(!top||!top.startsWith(startType.toLowerCase())){
    if(top){
      return `end${top.toUpperCase()}|start${startType}${rtlFlag}`;
    } else {
      return `start${startType}${rtlFlag}`;
    }
  }
  return null;
}

function handleListState(listChange, listStack, htmlLines){
  const actions=listChange.split('|');
  for(const action of actions){
    if(action.startsWith('start')){
      if(action.includes('UL_RTL')){
        htmlLines.push('<ul dir="rtl">');
        listStack.push('ul_rtl');
      } else if(action.includes('OL_RTL')){
        htmlLines.push('<ol dir="rtl">');
        listStack.push('ol_rtl');
      } else if(action.includes('UL')){
        htmlLines.push('<ul>');
        listStack.push('ul');
      } else {
        htmlLines.push('<ol>');
        listStack.push('ol');
      }
    } else if(action.startsWith('end')){
      const top=listStack.pop();
      if(top.startsWith('u')) htmlLines.push('</ul>');
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
  let html='<table class="doc-table" style="border-collapse:collapse; border:1px solid #ccc;">';
  for(const row of table.tableRows||[]){
    html+='<tr>';
    for(const cell of row.tableCells||[]){
      html+='<td style="border:1px solid #ccc; padding:0.5em;">';
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
  const familiesParam=unique.map(f=>f.trim().replace(/\s+/g,'+')).join('&family=');
  return `https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
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

