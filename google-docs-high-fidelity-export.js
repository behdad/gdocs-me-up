/**
 * google-docs-high-fidelity-export.js
 *
 * Usage:
 *   node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
 *
 * Features:
 *   - Hierarchical TOC indentation (based on heading level)
 *   - Doc-based column width (slightly tweaked)
 *   - Doc-based lineSpacing if present
 *   - Right-to-left paragraphs & bullet lists
 *   - Center, justify, start/end alignment preserved
 *   - Title, Subtitle, Headings from named styles
 *   - Subtitle supports multi-line if SHIFT+ENTER in the doc
 *   - Images with max-size logic (includes transform scale if present)
 *   - Table rendering
 *   - Merging consecutive text runs
 *   - Minimal .htaccess
 *
 * Fixes fetchAsBase64() not defined error.
 */

const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// ----------------- CONFIG -----------------
const SERVICE_ACCOUNT_KEY_FILE = 'service_account.json';
const EMBED_IMAGES_AS_BASE64 = false; // store images in "images/" folder

// Basic alignment map for LTR
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

  // Auth & fetch doc
  const authClient = await getAuthClient();
  const docs = google.docs({ version: 'v1', auth: authClient });
  const { data: doc } = await docs.documents.get({ documentId: docId });
  console.log(`Exporting: ${doc.title}`);

  // Build named styles (Title, Subtitle, headings)
  const namedStylesMap = buildNamedStylesMap(doc);

  // Determine container width from doc
  const colInfo = computeDocContainerWidth(doc);

  // Build global CSS
  const globalCSS = generateGlobalCSS(doc, colInfo);

  const usedFonts = new Set();
  let htmlOutput = [];

  // Basic HTML skeleton
  htmlOutput.push('<!DOCTYPE html>');
  htmlOutput.push('<html lang="en">');
  htmlOutput.push('<head>');
  htmlOutput.push('  <meta charset="UTF-8">');
  htmlOutput.push(`  <title>${escapeHtml(doc.title)}</title>`);
  htmlOutput.push('  <style>');
  htmlOutput.push(globalCSS);
  htmlOutput.push('  </style>');
  htmlOutput.push('</head>');
  htmlOutput.push('<body>');
  htmlOutput.push('<div class="doc-content">');

  let listStack = [];
  const bodyContent = doc.body?.content || [];

  for (const element of bodyContent) {
    if (element.sectionBreak) {
      htmlOutput.push('<div class="section-break"></div>');
      continue;
    }
    if (element.tableOfContents) {
      closeAllLists(listStack, htmlOutput);
      const tocHtml = await renderTableOfContents(
        element.tableOfContents,
        doc,
        usedFonts,
        authClient,
        outputDir,
        namedStylesMap
      );
      htmlOutput.push(tocHtml);
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
        path.join(outputDir, 'images'),
        namedStylesMap
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
    if (element.table) {
      closeAllLists(listStack, htmlOutput);
      const tableHtml = await renderTable(
        element.table,
        doc,
        usedFonts,
        authClient,
        outputDir,
        path.join(outputDir, 'images'),
        namedStylesMap
      );
      htmlOutput.push(tableHtml);
      continue;
    }
  }

  closeAllLists(listStack, htmlOutput);

  htmlOutput.push('</div>');
  htmlOutput.push('</body>');
  htmlOutput.push('</html>');

  // Insert Google Fonts link if needed
  const fontLink = buildGoogleFontsLink(Array.from(usedFonts));
  if (fontLink) {
    const idx = htmlOutput.findIndex(l => l.includes('</title>'));
    if (idx >= 0) {
      htmlOutput.splice(idx + 1, 0, `  <link rel="stylesheet" href="${fontLink}">`);
    }
  }

  // Write index.html
  const indexPath = path.join(outputDir, 'index.html');
  fs.writeFileSync(indexPath, htmlOutput.join('\n'), 'utf8');
  console.log(`HTML exported to: ${indexPath}`);

  // Write .htaccess
  const htaccessPath = path.join(outputDir, '.htaccess');
  fs.writeFileSync(htaccessPath, 'DirectoryIndex index.html\n', 'utf8');
  console.log(`.htaccess written to: ${htaccessPath}`);
}

// -------------------------------------
// 1) Column width from doc
// -------------------------------------
function computeDocContainerWidth(doc) {
  let containerPx = 800; // fallback
  const ds = doc.documentStyle;
  if (ds && ds.pageSize && ds.pageSize.width?.magnitude) {
    const pageW = ds.pageSize.width.magnitude;
    const leftM = ds.marginLeft?.magnitude || 72;
    const rightM = ds.marginRight?.magnitude || 72;
    const usablePts = pageW - (leftM + rightM);
    if (usablePts>0) containerPx = ptToPx(usablePts);
  }
  // Add a small tweak if you want
  containerPx += 50; // or remove
  return containerPx;
}

// -------------------------------------
// 2) Global CSS
// -------------------------------------
function generateGlobalCSS(doc, containerPx) {
  const lines = [];
  lines.push(`
body {
  margin: 0;
  font-family: sans-serif;
}
.doc-content {
  margin: 1em auto;
  max-width: ${containerPx}px;
  padding: 0 1em;
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
  border: 1px dashed #666;
  background: #f9f9f9;
}
.doc-toc h2 {
  margin: 0 0 0.3em 0;
  font-size: 1.2em;
  font-weight: bold;
}
/* TOC indentation */
.doc-toc .toc-level-1 { margin-left: 0; }
.doc-toc .toc-level-2 { margin-left: 1em; }
.doc-toc .toc-level-3 { margin-left: 2em; }
.doc-toc .toc-level-4 { margin-left: 3em; }

.doc-subtitle {
  display: block;
  white-space: pre-wrap; /* multi-line subtitle */
}

.doc-table {
  border-collapse: collapse;
  margin: 0.5em 0;
}
`);

  // If doc has pageSize => add @page
  if (doc.documentStyle?.pageSize?.width?.magnitude && doc.documentStyle?.pageSize?.height?.magnitude){
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

// -------------------------------------
// 3) Table of Contents with indentation
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
      if (!c.paragraph) continue;
      // detect the deepest heading level referenced by text runs in this paragraph
      let headingLevel = 1;
      for (const elem of c.paragraph.elements||[]) {
        const s = elem.textRun?.textStyle;
        if (s?.link?.headingId) {
          const lvl = findHeadingLevelById(doc, s.link.headingId);
          if(lvl>headingLevel) headingLevel=lvl;
        }
      }
      if(headingLevel<1) headingLevel=1;
      if(headingLevel>4) headingLevel=4;

      // Render the paragraph normally
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
      html += `<div class="toc-level-${headingLevel}">${pHtml}</div>\n`;
    }
  }

  html += '</div>\n';
  return html;
}

/** findHeadingLevelById => integer 1..6 (or 1 fallback) */
function findHeadingLevelById(doc, headingId) {
  const content = doc.body?.content||[];
  for(const e of content){
    if(e.paragraph) {
      const ps=e.paragraph.paragraphStyle;
      if(ps?.headingId===headingId) {
        const namedType=ps.namedStyleType||'NORMAL_TEXT';
        if(namedType.startsWith('HEADING_')){
          const lvl=parseInt(namedType.replace('HEADING_',''),10);
          if(lvl>=1 && lvl<=6) return lvl;
        }
      }
    }
  }
  return 1;
}

// -------------------------------------
// 4) Paragraph
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
  const style = paragraph.paragraphStyle||{};
  const namedType = style.namedStyleType||'NORMAL_TEXT';

  // Merge doc-based named style
  let mergedParaStyle={};
  let mergedTextStyle={};
  if (namedStylesMap[namedType]) {
    mergedParaStyle = deepCopy(namedStylesMap[namedType].paragraphStyle);
    mergedTextStyle = deepCopy(namedStylesMap[namedType].textStyle);
  }
  deepMerge(mergedParaStyle, style);

  // bullet logic
  let listChange = null;
  const isRTL = (mergedParaStyle.direction==='RIGHT_TO_LEFT');
  if (paragraph.bullet) {
    listChange = detectListChange(paragraph.bullet, doc, listStack, isRTL);
  } else {
    if(listStack.length>0){
      const top=listStack[listStack.length-1];
      listChange=`end${top.toUpperCase()}`;
    }
  }

  // Title => h1, Subtitle => h2 class="doc-subtitle", etc.
  let tag='p';
  if (namedType==='TITLE'){
    tag='h1';
  } else if (namedType==='SUBTITLE'){
    tag='h2 class="doc-subtitle"';
  } else if (namedType.startsWith('HEADING_')){
    const lv=parseInt(namedType.replace('HEADING_',''),10);
    if(lv>=1 && lv<=6) tag=`h${lv}`;
  }

  let headingIdAttr = '';
  if (style.headingId) {
    headingIdAttr = ` id="heading-${escapeHtml(style.headingId)}"`;
  }

  // alignment flipping
  let align = mergedParaStyle.alignment;
  if(isRTL && (align==='START' || align==='END')){
    align=(align==='START')?'END':'START';
  }

  // Build doc-based inline style: lineSpacing => line-height
  let inlineStyle='';
  if(align && alignmentMapLTR[align]){
    inlineStyle += `text-align: ${alignmentMapLTR[align]};`;
  }
  if(mergedParaStyle.lineSpacing){
    const ls=mergedParaStyle.lineSpacing/100;
    inlineStyle += `line-height: ${ls};`;
  }
  if(mergedParaStyle.spaceAbove?.magnitude){
    inlineStyle += `margin-top: ${ptToPx(mergedParaStyle.spaceAbove.magnitude)}px;`;
  }
  if(mergedParaStyle.spaceBelow?.magnitude){
    inlineStyle += `margin-bottom: ${ptToPx(mergedParaStyle.spaceBelow.magnitude)}px;`;
  }
  if(mergedParaStyle.indentFirstLine?.magnitude){
    inlineStyle += `text-indent: ${ptToPx(mergedParaStyle.indentFirstLine.magnitude)}px;`;
  } else if(mergedParaStyle.indentStart?.magnitude){
    inlineStyle += `margin-left: ${ptToPx(mergedParaStyle.indentStart.magnitude)}px;`;
  }

  let dirAttr = '';
  if(mergedParaStyle.direction==='RIGHT_TO_LEFT') {
    dirAttr=' dir="rtl"';
  }

  // Merge text runs
  const mergedRuns=mergeTextRuns(paragraph.elements||[]);
  let innerHtml='';
  for(const r of mergedRuns){
    if(r.inlineObjectElement){
      const objId=r.inlineObjectElement.inlineObjectId;
      innerHtml+= await renderInlineObject(objId, doc, authClient, outputDir, imagesDir);
    } else if(r.textRun){
      innerHtml+= renderTextRun(r.textRun, usedFonts, mergedTextStyle);
    }
  }

  let paragraphHtml = `<${tag}${headingIdAttr}${dirAttr}>${innerHtml}</${tag.split(' ')[0]}>`;
  if(inlineStyle){
    // insert style= inline
    const closeBracket=paragraphHtml.indexOf('>');
    if(closeBracket>0){
      paragraphHtml=paragraphHtml.slice(0, closeBracket)
        + ` style="${inlineStyle}"`
        + paragraphHtml.slice(closeBracket);
    }
  }

  return { html: paragraphHtml, listChange };
}

// -------------------------------------
// 5) Merge text runs
// -------------------------------------
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
        merged.push({ textRun:{ content, textStyle: deepCopy(style) }});
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

// -------------------------------------
// 6) Render text run
// -------------------------------------
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

// -------------------------------------
// 7) Inline Objects (Images)
// -------------------------------------
async function renderInlineObject(objectId, doc, authClient, outputDir, imagesDir){
  const inlineObj=doc.inlineObjects?.[objectId];
  if(!inlineObj) return'';

  const embedded=inlineObj.inlineObjectProperties?.embeddedObject;
  if(!embedded?.imageProperties) return'';

  const { imageProperties }=embedded;
  const { contentUri, size }=imageProperties;

  // Some docs store scale in embedded.transform => scaleX, scaleY
  let scaleX=1.0, scaleY=1.0;
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

// -------------------------------------
// 8) Lists
// -------------------------------------
function detectListChange(bullet,doc,listStack,isRTL){
  const listId=bullet.listId;
  const nestingLevel=bullet.nestingLevel||0;
  const listDef=doc.lists?.[listId];
  if(!listDef?.listProperties?.nestingLevels) return null;

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

function handleListState(listChange,listStack,htmlLines){
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

// -------------------------------------
// 9) Table
// -------------------------------------
async function renderTable(
  table, doc, usedFonts, authClient, outputDir, imagesDir, namedStylesMap
){
  let html='<table class="doc-table" style="border-collapse:collapse; border:1px solid #ccc;">';
  for(const row of table.tableRows||[]){
    html+='<tr>';
    for(const cell of row.tableCells||[]){
      html+='<td style="border:1px solid #ccc; padding:0.5em;">';
      for(const c of cell.content||[]){
        if(c.paragraph){
          const { html:pHtml } = await renderParagraph(
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

// -------------------------------------
// 10) Named Styles
// -------------------------------------
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

// -------------------------------------
// UTILS
// -------------------------------------
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

// The missing function that was causing ReferenceError
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
  const nr=Math.round(r*255);
  const ng=Math.round(g*255);
  const nb=Math.round(b*255);
  return '#'+[nr,ng,nb].map(x=>x.toString(16).padStart(2,'0')).join('');
}
function buildGoogleFontsLink(fontFamilies){
  if(!fontFamilies||fontFamilies.length===0)return'';
  const unique=Array.from(new Set(fontFamilies));
  const familiesParam=unique.map(f=>f.trim().replace(/\s+/g,'+')).join('&family=');
  return`https://fonts.googleapis.com/css2?family=${familiesParam}&display=swap`;
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

