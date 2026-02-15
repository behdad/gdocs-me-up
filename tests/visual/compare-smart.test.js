const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

/**
 * Smart Visual Comparison
 *
 * Instead of pixel-by-pixel comparison, this compares:
 * 1. Document structure (elements, hierarchy)
 * 2. Content (text, links, images)
 * 3. Layout metrics (sizes, positions, spacing)
 * 4. Side-by-side screenshots for human review
 *
 * This gives actionable insights for improvement.
 */

const TEST_DOCS = {
  'text-rendering': '1UnR2zKf3Z_DDRS6vLgBkSHUeqI3IGOEhWYh7rAIvsb8',
  'behdad-story': '1MVNNjtoejIqvJrVruFo20qfW36ydLEebj0dRN7-bZrA'
};

test.describe('Smart Document Comparison', () => {
  test.beforeAll(async () => {
    execSync('npx playwright install chromium', { stdio: 'inherit' });
  });

  for (const [docName, docId] of Object.entries(TEST_DOCS)) {
    test(`analyze: ${docName}`, async ({ browser }) => {
      const outputDir = path.join(__dirname, 'output', docName);
      const analysisDir = path.join(__dirname, 'analysis', docName);
      fs.mkdirSync(analysisDir, { recursive: true });

      // Export the document
      console.log(`\n📦 Exporting: ${docName}`);
      try {
        execSync(`node gdocs-me-up.js ${docId} ${outputDir}`, {
          cwd: path.join(__dirname, '../..'),
          stdio: 'pipe'
        });
      } catch (error) {
        console.error('Export failed:', error.message);
        throw error;
      }

      // Create two browser contexts for parallel analysis
      const googleContext = await browser.newContext();
      const exportContext = await browser.newContext();

      const googlePage = await googleContext.newPage();
      const exportPage = await exportContext.newPage();

      // Load both versions
      const googleDocUrl = `https://docs.google.com/document/d/${docId}/preview`;
      await googlePage.goto(googleDocUrl, { waitUntil: 'networkidle' });
      await googlePage.waitForTimeout(2000);

      const htmlPath = `file://${path.join(outputDir, 'index.html')}`;
      await exportPage.goto(htmlPath, { waitUntil: 'networkidle' });
      await exportPage.waitForTimeout(1000);

      // ====================
      // 1. STRUCTURAL ANALYSIS
      // ====================
      console.log('\n📊 Analyzing structure...');
      const exportStructure = await analyzeStructure(exportPage);
      console.log('Exported document structure:', JSON.stringify(exportStructure, null, 2));

      // ====================
      // 2. CONTENT EXTRACTION
      // ====================
      console.log('\n📝 Extracting content...');
      const googleContent = await extractContent(googlePage);
      const exportContent = await extractContent(exportPage);

      // Compare text content
      const contentComparison = compareContent(googleContent, exportContent);
      console.log('Content comparison:', contentComparison);

      // ====================
      // 3. LAYOUT METRICS
      // ====================
      console.log('\n📐 Analyzing layout...');
      const exportLayout = await analyzeLayout(exportPage);
      console.log('Layout metrics:', JSON.stringify(exportLayout, null, 2));

      // ====================
      // 4. VISUAL ELEMENTS
      // ====================
      console.log('\n🖼️  Analyzing images...');
      const googleImages = await analyzeImages(googlePage);
      const exportImages = await analyzeImages(exportPage);

      console.log(`Google Doc images: ${googleImages.length}`);
      console.log(`Exported images: ${exportImages.length}`);

      if (exportImages.length > 0) {
        console.log('Sample exported image:', exportImages[0]);
      }

      // ====================
      // 5. SIDE-BY-SIDE SCREENSHOTS
      // ====================
      console.log('\n📸 Creating side-by-side comparison...');

      await googlePage.screenshot({
        path: path.join(analysisDir, 'google-doc.png'),
        fullPage: true
      });

      await exportPage.screenshot({
        path: path.join(analysisDir, 'exported.png'),
        fullPage: true
      });

      // ====================
      // 6. GENERATE REPORT
      // ====================
      const report = {
        docName,
        docId,
        timestamp: new Date().toISOString(),
        structure: exportStructure,
        content: {
          googleWords: googleContent.text.split(/\s+/).length,
          exportedWords: exportContent.text.split(/\s+/).length,
          textMatch: contentComparison.textSimilarity,
          missingHeadings: contentComparison.missingHeadings,
        },
        layout: exportLayout,
        images: {
          google: googleImages.length,
          exported: exportImages.length,
          sizingIssues: exportImages.filter(img =>
            img.width > 1000 || img.height > 1000
          ).length
        },
        recommendations: generateRecommendations(exportStructure, exportLayout, exportImages)
      };

      fs.writeFileSync(
        path.join(analysisDir, 'report.json'),
        JSON.stringify(report, null, 2)
      );

      // Generate HTML report for easy viewing
      const htmlReport = generateHTMLReport(report, analysisDir);
      fs.writeFileSync(
        path.join(analysisDir, 'report.html'),
        htmlReport
      );

      console.log(`\n✅ Analysis complete!`);
      console.log(`   Report: ${path.join(analysisDir, 'report.html')}`);
      console.log(`\n🎯 Recommendations:`);
      report.recommendations.forEach((rec, i) => {
        console.log(`   ${i + 1}. ${rec}`);
      });

      await googleContext.close();
      await exportContext.close();

      // Test assertions
      expect(exportStructure.paragraphs).toBeGreaterThan(0);
    });
  }
});

/**
 * Analyze document structure
 */
async function analyzeStructure(page) {
  return await page.evaluate(() => {
    return {
      headings: {
        h1: document.querySelectorAll('h1').length,
        h2: document.querySelectorAll('h2').length,
        h3: document.querySelectorAll('h3').length,
        h4: document.querySelectorAll('h4').length,
        h5: document.querySelectorAll('h5').length,
        h6: document.querySelectorAll('h6').length,
      },
      paragraphs: document.querySelectorAll('p').length,
      lists: {
        unordered: document.querySelectorAll('ul').length,
        ordered: document.querySelectorAll('ol').length,
        items: document.querySelectorAll('li').length,
      },
      tables: document.querySelectorAll('table').length,
      images: document.querySelectorAll('img').length,
      links: document.querySelectorAll('a').length,
      formatting: {
        bold: document.querySelectorAll('.bold, strong, b').length,
        italic: document.querySelectorAll('.italic, em, i').length,
        underline: document.querySelectorAll('.underline, u').length,
      }
    };
  });
}

/**
 * Extract text content
 */
async function extractContent(page) {
  return await page.evaluate(() => {
    const getVisibleText = (el) => {
      const style = window.getComputedStyle(el);
      if (style.display === 'none' || style.visibility === 'hidden') return '';
      return el.innerText || '';
    };

    return {
      text: getVisibleText(document.body),
      headings: Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, h6')).map(h => ({
        level: h.tagName.toLowerCase(),
        text: h.innerText.trim()
      })),
      links: Array.from(document.querySelectorAll('a')).map(a => ({
        text: a.innerText.trim(),
        href: a.getAttribute('href')
      }))
    };
  });
}

/**
 * Compare content between two documents
 */
function compareContent(google, exported) {
  const googleHeadings = new Set(google.headings.map(h => h.text));
  const exportedHeadings = new Set(exported.headings.map(h => h.text));

  const missingHeadings = [...googleHeadings].filter(h => !exportedHeadings.has(h));

  // Simple text similarity (could be improved with proper diff algorithm)
  const googleWords = google.text.toLowerCase().split(/\s+/).filter(w => w.length > 3);
  const exportedWords = exported.text.toLowerCase().split(/\s+/).filter(w => w.length > 3);
  const commonWords = googleWords.filter(w => exportedWords.includes(w));
  const textSimilarity = (commonWords.length / Math.max(googleWords.length, exportedWords.length)) * 100;

  return {
    textSimilarity: textSimilarity.toFixed(2) + '%',
    missingHeadings,
    googleHeadingCount: googleHeadings.size,
    exportedHeadingCount: exportedHeadings.size
  };
}

/**
 * Analyze layout metrics
 */
async function analyzeLayout(page) {
  return await page.evaluate(() => {
    const docContent = document.querySelector('.doc-content, body');
    const bbox = docContent?.getBoundingClientRect();

    const paragraphs = Array.from(document.querySelectorAll('p')).slice(0, 10);
    const avgParagraphSpacing = paragraphs.length > 1
      ? paragraphs.slice(1).reduce((sum, p, i) => {
          const prev = paragraphs[i].getBoundingClientRect();
          const curr = p.getBoundingClientRect();
          return sum + (curr.top - prev.bottom);
        }, 0) / (paragraphs.length - 1)
      : 0;

    return {
      documentWidth: bbox?.width || 0,
      documentHeight: bbox?.height || 0,
      averageParagraphSpacing: Math.round(avgParagraphSpacing),
      firstParagraphFontSize: paragraphs[0]
        ? window.getComputedStyle(paragraphs[0]).fontSize
        : 'N/A'
    };
  });
}

/**
 * Analyze images
 */
async function analyzeImages(page) {
  return await page.evaluate(() => {
    return Array.from(document.querySelectorAll('img')).map(img => {
      const bbox = img.getBoundingClientRect();
      const style = window.getComputedStyle(img);
      return {
        src: img.src.substring(0, 50) + '...',
        width: Math.round(bbox.width),
        height: Math.round(bbox.height),
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        maxWidth: style.maxWidth,
        objectFit: style.objectFit,
        display: style.display
      };
    });
  });
}

/**
 * Generate improvement recommendations
 */
function generateRecommendations(structure, layout, images) {
  const recommendations = [];

  if (images.some(img => img.width > 1000)) {
    recommendations.push('Some images are very wide (>1000px). Consider constraining to container width.');
  }

  if (layout.averageParagraphSpacing < 5) {
    recommendations.push('Paragraph spacing seems tight. Consider increasing line-height or margins.');
  }

  if (structure.images === 0 && images.length === 0) {
    recommendations.push('No images detected. If original has images, check positioned objects support.');
  }

  if (structure.lists.items === 0) {
    recommendations.push('No list items detected. If original has lists, check list rendering.');
  }

  if (recommendations.length === 0) {
    recommendations.push('Document structure looks good! ✨');
  }

  return recommendations;
}

/**
 * Generate HTML report for easy viewing
 */
function generateHTMLReport(report, analysisDir) {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Document Analysis: ${report.docName}</title>
  <style>
    body { font-family: system-ui, sans-serif; max-width: 1400px; margin: 40px auto; padding: 0 20px; }
    h1 { color: #333; }
    h2 { color: #666; margin-top: 2em; border-bottom: 2px solid #eee; padding-bottom: 0.5em; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0; }
    .card { background: #f9f9f9; padding: 20px; border-radius: 8px; }
    .card h3 { margin-top: 0; color: #555; }
    .metric { font-size: 2em; font-weight: bold; color: #007bff; }
    .screenshots { margin: 20px 0; }
    .screenshots img { max-width: 100%; border: 1px solid #ddd; border-radius: 4px; }
    .recommendation { background: #fff3cd; padding: 10px 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #ffc107; }
    pre { background: #f5f5f5; padding: 15px; border-radius: 4px; overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
    td, th { padding: 10px; text-align: left; border-bottom: 1px solid #eee; }
    th { background: #f5f5f5; font-weight: 600; }
  </style>
</head>
<body>
  <h1>📊 Document Analysis Report</h1>
  <p><strong>Document:</strong> ${report.docName} | <strong>Date:</strong> ${new Date(report.timestamp).toLocaleString()}</p>

  <h2>🎯 Recommendations</h2>
  ${report.recommendations.map(rec => `<div class="recommendation">${rec}</div>`).join('')}

  <h2>📐 Key Metrics</h2>
  <div class="grid">
    <div class="card">
      <h3>Structure</h3>
      <table>
        <tr><td>Paragraphs</td><td class="metric">${report.structure.paragraphs}</td></tr>
        <tr><td>Headings</td><td class="metric">${Object.values(report.structure.headings).reduce((a,b)=>a+b,0)}</td></tr>
        <tr><td>Images</td><td class="metric">${report.structure.images}</td></tr>
        <tr><td>Lists</td><td class="metric">${report.structure.lists.unordered + report.structure.lists.ordered}</td></tr>
      </table>
    </div>

    <div class="card">
      <h3>Content</h3>
      <table>
        <tr><td>Word Count</td><td class="metric">${report.content.exportedWords}</td></tr>
        <tr><td>Text Similarity</td><td class="metric">${report.content.textSimilarity}</td></tr>
        <tr><td>Images</td><td class="metric">${report.images.exported}</td></tr>
      </table>
    </div>
  </div>

  <h2>📸 Side-by-Side Comparison</h2>
  <div class="grid screenshots">
    <div>
      <h3>Google Docs (Original)</h3>
      <img src="google-doc.png" alt="Google Doc">
    </div>
    <div>
      <h3>Exported HTML</h3>
      <img src="exported.png" alt="Exported">
    </div>
  </div>

  <h2>📋 Detailed Analysis</h2>
  <h3>Structure</h3>
  <pre>${JSON.stringify(report.structure, null, 2)}</pre>

  <h3>Layout</h3>
  <pre>${JSON.stringify(report.layout, null, 2)}</pre>

  <h3>Images</h3>
  <pre>${JSON.stringify(report.images, null, 2)}</pre>
</body>
</html>`;
}
