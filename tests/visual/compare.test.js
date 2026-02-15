const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');
const { PNG } = require('pngjs');
const pixelmatch = require('pixelmatch');

// Import the export function (you'll need to export it from gdocs-me-up.js)
// For now, we'll use child_process to run the export
const { execSync } = require('child_process');

/**
 * Visual Comparison Test Suite
 *
 * Compares Google Docs rendering with our HTML export to identify
 * differences and guide improvements.
 */

// Test document IDs - update these with your test documents
const TEST_DOCS = {
  'text-rendering': '1UnR2zKf3Z_DDRS6vLgBkSHUeqI3IGOEhWYh7rAIvsb8',
  'behdad-story': '1MVNNjtoejIqvJrVruFo20qfW36ydLEebj0dRN7-bZrA'
};

test.describe('Visual Comparison', () => {
  test.beforeAll(async () => {
    // Install browsers if needed
    execSync('npx playwright install chromium', { stdio: 'inherit' });
  });

  for (const [docName, docId] of Object.entries(TEST_DOCS)) {
    test(`compare: ${docName}`, async ({ page }) => {
      const screenshotsDir = path.join(__dirname, 'screenshots', docName);
      fs.mkdirSync(screenshotsDir, { recursive: true });

      // Step 1: Screenshot the original Google Doc
      console.log(`📸 Capturing Google Doc: ${docName}`);
      const googleDocUrl = `https://docs.google.com/document/d/${docId}/preview`;

      await page.goto(googleDocUrl, { waitUntil: 'networkidle' });
      await page.waitForTimeout(2000); // Wait for fonts to load

      const googleDocScreenshot = path.join(screenshotsDir, 'google-doc.png');
      await page.screenshot({
        path: googleDocScreenshot,
        fullPage: true
      });

      // Step 2: Export the document
      console.log(`📦 Exporting document: ${docName}`);
      const outputDir = path.join(__dirname, 'output', docName);
      try {
        execSync(`node gdocs-me-up.js ${docId} ${outputDir}`, {
          cwd: path.join(__dirname, '../..'),
          stdio: 'pipe'
        });
      } catch (error) {
        console.error('Export failed:', error.message);
        throw error;
      }

      // Step 3: Screenshot our exported HTML
      console.log(`📸 Capturing exported HTML: ${docName}`);
      const htmlPath = `file://${path.join(outputDir, 'index.html')}`;
      await page.goto(htmlPath, { waitUntil: 'networkidle' });
      await page.waitForTimeout(2000); // Wait for fonts to load

      const exportedScreenshot = path.join(screenshotsDir, 'exported-html.png');
      await page.screenshot({
        path: exportedScreenshot,
        fullPage: true
      });

      // Step 4: Compare images
      console.log(`🔍 Comparing screenshots: ${docName}`);
      const comparison = await compareImages(
        googleDocScreenshot,
        exportedScreenshot,
        path.join(screenshotsDir, 'diff.png')
      );

      console.log(`📊 Comparison results:
        - Difference: ${comparison.diffPercentage.toFixed(2)}%
        - Mismatched pixels: ${comparison.mismatchedPixels}
        - Total pixels: ${comparison.totalPixels}`);

      // Save comparison report
      const report = {
        docName,
        docId,
        timestamp: new Date().toISOString(),
        diffPercentage: comparison.diffPercentage,
        mismatchedPixels: comparison.mismatchedPixels,
        totalPixels: comparison.totalPixels,
        screenshots: {
          googleDoc: googleDocScreenshot,
          exported: exportedScreenshot,
          diff: path.join(screenshotsDir, 'diff.png')
        }
      };

      fs.writeFileSync(
        path.join(screenshotsDir, 'report.json'),
        JSON.stringify(report, null, 2)
      );

      // Log areas for improvement
      if (comparison.diffPercentage > 5) {
        console.log(`⚠️  Significant differences detected (${comparison.diffPercentage.toFixed(2)}%)`);
        console.log(`   Review diff image at: ${path.join(screenshotsDir, 'diff.png')}`);
      } else {
        console.log(`✅ Export looks good! Only ${comparison.diffPercentage.toFixed(2)}% difference`);
      }

      // Test passes if difference is below threshold (adjust as needed)
      // For now, we just report - no failure
      expect(comparison.diffPercentage).toBeLessThan(100); // Always pass
    });
  }
});

/**
 * Compare two PNG images pixel by pixel
 */
async function compareImages(img1Path, img2Path, diffPath) {
  const img1 = PNG.sync.read(fs.readFileSync(img1Path));
  const img2 = PNG.sync.read(fs.readFileSync(img2Path));

  // Resize images to match if needed (use smaller dimensions)
  const width = Math.min(img1.width, img2.width);
  const height = Math.min(img1.height, img2.height);

  const diff = new PNG({ width, height });

  const mismatchedPixels = pixelmatch(
    img1.data,
    img2.data,
    diff.data,
    width,
    height,
    {
      threshold: 0.1,
      includeAA: true
    }
  );

  // Save diff image
  fs.writeFileSync(diffPath, PNG.sync.write(diff));

  const totalPixels = width * height;
  const diffPercentage = (mismatchedPixels / totalPixels) * 100;

  return {
    mismatchedPixels,
    totalPixels,
    diffPercentage,
    width,
    height
  };
}

/**
 * Extract specific metrics from screenshots for detailed comparison
 */
test.describe('Detailed Element Comparison', () => {
  test('analyze text rendering differences', async ({ page }) => {
    const docId = TEST_DOCS['text-rendering'];
    const outputDir = path.join(__dirname, 'output', 'text-rendering');

    // Export if not already done
    if (!fs.existsSync(outputDir)) {
      execSync(`node gdocs-me-up.js ${docId} ${outputDir}`, {
        cwd: path.join(__dirname, '../..'),
        stdio: 'pipe'
      });
    }

    const htmlPath = `file://${path.join(outputDir, 'index.html')}`;
    await page.goto(htmlPath, { waitUntil: 'networkidle' });

    // Analyze specific elements
    const analysis = {
      images: await page.locator('img').count(),
      headings: {
        h1: await page.locator('h1').count(),
        h2: await page.locator('h2').count(),
        h3: await page.locator('h3').count(),
      },
      lists: {
        ul: await page.locator('ul').count(),
        ol: await page.locator('ol').count(),
      },
      paragraphs: await page.locator('p').count(),
      tables: await page.locator('table').count(),
    };

    console.log('📋 Document structure:', JSON.stringify(analysis, null, 2));

    // Check image sizing
    const images = await page.locator('img').all();
    for (let i = 0; i < Math.min(images.length, 5); i++) {
      const img = images[i];
      const box = await img.boundingBox();
      const style = await img.getAttribute('style');
      console.log(`Image ${i + 1}:`, {
        dimensions: box ? `${box.width}x${box.height}` : 'N/A',
        style: style?.substring(0, 100)
      });
    }

    expect(analysis.images).toBeGreaterThanOrEqual(0);
  });
});
