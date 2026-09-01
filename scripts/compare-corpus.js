#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const { promisify } = require('util');
const { chromium } = require('playwright');
const sharp = require('sharp');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'tests', 'visual', 'corpus');
const VIEWPORT = { width: 1280, height: 720 };
const execFileAsync = promisify(execFile);

function discoverDocuments(sourceDir) {
  const byId = new Map();
  for (const name of fs.readdirSync(sourceDir).sort()) {
    const filePath = path.join(sourceDir, name);
    let source;
    try {
      if (!fs.statSync(filePath).isFile()) continue;
      source = fs.readFileSync(filePath, 'utf8');
    } catch {
      continue;
    }
    const match = source.match(/https:\/\/docs\.google\.com\/document\/d\/([\w-]+)\/preview/);
    if (!match) continue;
    const docId = match[1];
    const current = byId.get(docId) || { docId, names: [] };
    current.names.push(name);
    byId.set(docId, current);
  }
  return [...byId.values()].map(item => ({ ...item, name: chooseName(item.names) }));
}

function chooseName(names) {
  return names.find(name => !/^\d+$/.test(name)) || names[0];
}

async function pixelDifference(left, right) {
  const [a, b] = await Promise.all([
    sharp(left).removeAlpha().grayscale().blur(2).raw().toBuffer(),
    sharp(right).removeAlpha().grayscale().blur(2).raw().toBuffer()
  ]);
  let difference = 0;
  for (let i = 0; i < a.length; i++) difference += Math.abs(a[i] - b[i]);
  return difference / a.length / 255;
}

async function analyzeExport(page, htmlPath) {
  await page.goto(`file://${htmlPath}`, { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  return page.evaluate(() => ({
    height: document.documentElement.scrollHeight,
    inlineStyles: document.querySelectorAll('[style]').length,
    paragraphs: document.querySelectorAll('p').length,
    headings: document.querySelectorAll('h1,h2,h3,h4,h5,h6').length,
    lists: document.querySelectorAll('ul,ol').length,
    listItems: document.querySelectorAll('li').length,
    tables: document.querySelectorAll('table').length,
    images: document.querySelectorAll('.doc-content img').length,
    topLevelListItems: document.querySelectorAll('.doc-content > li').length
  }));
}

async function main() {
  const sourceDir = process.env.GDOCS_CORPUS_DIR;
  if (!sourceDir) {
    throw new Error('Set GDOCS_CORPUS_DIR to a directory of Google Docs preview-link fixtures.');
  }
  const limitArg = process.argv.find(arg => arg.startsWith('--limit='));
  const limit = limitArg ? Number(limitArg.split('=')[1]) : Infinity;
  const namesArg = process.argv.find(arg => arg.startsWith('--names='));
  const selectedNames = namesArg ? new Set(namesArg.split('=')[1].split(',')) : null;
  const concurrencyArg = process.argv.find(arg => arg.startsWith('--concurrency='));
  const concurrency = Math.max(1, Number(concurrencyArg?.split('=')[1] || process.env.CORPUS_CONCURRENCY || 16));
  const documents = discoverDocuments(sourceDir)
    .filter(document => !selectedNames || selectedNames.has(document.name) || document.names.some(name => selectedNames.has(name)))
    .slice(0, limit);
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const browser = await chromium.launch({ headless: true });
  const report = new Array(documents.length);
  const contactRows = new Array(documents.length);
  const workerCount = Math.min(concurrency, documents.length);
  process.stdout.write(`Running ${documents.length} documents with ${workerCount} workers\n`);

  async function processDocument(document, index, googlePage, exportPage) {
    const outputDir = path.join(OUTPUT_DIR, 'exports', document.name);
    const screenshotDir = path.join(OUTPUT_DIR, 'screenshots', document.name);
    fs.mkdirSync(screenshotDir, { recursive: true });
    process.stdout.write(`[${index + 1}/${documents.length}] ${document.name}\n`);

    await execFileAsync(process.execPath, [path.join(ROOT, 'gdocs-me-up.js'), document.docId, outputDir], {
      cwd: ROOT,
      maxBuffer: 10 * 1024 * 1024
    });

    const googleUrl = `https://docs.google.com/document/d/${document.docId}/preview`;
    await googlePage.goto(googleUrl, { waitUntil: 'networkidle', timeout: 60000 });
    await googlePage.waitForTimeout(750);
    const googleScreenshot = await googlePage.screenshot();

    const htmlPath = path.join(outputDir, 'index.html');
    const metrics = await analyzeExport(exportPage, htmlPath);
    const exportScreenshot = await exportPage.screenshot();
    const difference = await pixelDifference(googleScreenshot, exportScreenshot);

    fs.writeFileSync(path.join(screenshotDir, 'google.png'), googleScreenshot);
    fs.writeFileSync(path.join(screenshotDir, 'export.png'), exportScreenshot);
    const comparison = await sharp({
      create: { width: VIEWPORT.width * 2, height: VIEWPORT.height, channels: 3, background: 'white' }
    }).composite([
      { input: googleScreenshot, left: 0, top: 0 },
      { input: exportScreenshot, left: VIEWPORT.width, top: 0 }
    ]).jpeg({ quality: 85 }).toBuffer();
    fs.writeFileSync(path.join(screenshotDir, 'comparison.jpg'), comparison);

    const thumbnail = await sharp(comparison).resize(640, 180, { fit: 'fill' }).toBuffer();
    const label = Buffer.from(
      `<svg width="640" height="24"><rect width="640" height="24" fill="white"/>` +
      `<text x="8" y="17" font-family="sans-serif" font-size="14">${escapeXml(document.name)} — ${(difference * 100).toFixed(1)}% pixel difference</text></svg>`
    );
    contactRows[index] = await sharp({
      create: { width: 640, height: 204, channels: 3, background: 'white' }
    }).composite([{ input: label, top: 0, left: 0 }, { input: thumbnail, top: 24, left: 0 }]).png().toBuffer();

    const imageFiles = fs.existsSync(path.join(outputDir, 'images'))
      ? fs.readdirSync(path.join(outputDir, 'images'))
      : [];
    report[index] = {
      ...document,
      difference,
      htmlBytes: fs.statSync(htmlPath).size,
      imageBytes: imageFiles.reduce((total, file) => total + fs.statSync(path.join(outputDir, 'images', file)).size, 0),
      imageFormats: imageFiles.reduce((formats, file) => {
        const extension = path.extname(file).slice(1);
        formats[extension] = (formats[extension] || 0) + 1;
        return formats;
      }, {}),
      ...metrics
    };
  }

  await Promise.all(Array.from({ length: workerCount }, async (_, workerIndex) => {
    const googlePage = await browser.newPage({ viewport: VIEWPORT });
    const exportPage = await browser.newPage({ viewport: VIEWPORT });
    try {
      for (let index = workerIndex; index < documents.length; index += workerCount) {
        await processDocument(documents[index], index, googlePage, exportPage);
      }
    } finally {
      await Promise.all([googlePage.close(), exportPage.close()]);
    }
  }));

  await browser.close();
  fs.writeFileSync(path.join(OUTPUT_DIR, 'report.json'), JSON.stringify(report, null, 2));
  if (contactRows.length) {
    await sharp({
      create: { width: 640, height: contactRows.length * 204, channels: 3, background: 'white' }
    }).composite(contactRows.map((input, index) => ({ input, left: 0, top: index * 204 })))
      .jpeg({ quality: 82 })
      .toFile(path.join(OUTPUT_DIR, 'contact-sheet.jpg'));
  }

  const average = report.reduce((sum, item) => sum + item.difference, 0) / report.length;
  process.stdout.write(`Compared ${report.length} unique documents; average top-viewport difference ${(average * 100).toFixed(2)}%\n`);
}

function escapeXml(value) {
  return value.replace(/[<>&"']/g, character => ({
    '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;'
  })[character]);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
