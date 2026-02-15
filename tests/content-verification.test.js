/**
 * Content Verification Tests
 *
 * Uses stable reference documents as golden standards:
 * - Behdad's Story (1MVNNjtoejIqvJrVruFo20qfW36ydLEebj0dRN7-bZrA)
 * - State of Text Rendering 2024 (1UnR2zKf3Z_DDRS6vLgBkSHUeqI3IGOEhWYh7rAIvsb8)
 *
 * These documents won't be modified, so we can use them to verify:
 * - Content accuracy (text, links, images)
 * - Structure preservation (headings, lists, paragraphs)
 * - No regressions over time
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const cheerio = require('cheerio');
const { google } = require('googleapis');

const REFERENCE_DOCS = {
  'behdad-story': {
    id: '1MVNNjtoejIqvJrVruFo20qfW36ydLEebj0dRN7-bZrA',
    name: "Behdad's Story"
  },
  'text-rendering-2024': {
    id: '1UnR2zKf3Z_DDRS6vLgBkSHUeqI3IGOEhWYh7rAIvsb8',
    name: 'State of Text Rendering 2024'
  }
};

// Increase timeout for API calls
jest.setTimeout(30000);

describe('Content Verification - Reference Documents', () => {
  let authClient;

  beforeAll(async () => {
    // Get auth client (same path as main script)
    const keyFile = path.join(__dirname, '..', 'service_account.json');
    const auth = new google.auth.GoogleAuth({
      keyFile,
      scopes: [
        'https://www.googleapis.com/auth/documents.readonly',
        'https://www.googleapis.com/auth/drive.readonly'
      ]
    });
    authClient = await auth.getClient();
  });

  for (const [docKey, docInfo] of Object.entries(REFERENCE_DOCS)) {
    describe(docInfo.name, () => {
      let googleDoc;
      let exportedHTML;
      let $; // Cheerio instance

      beforeAll(async () => {
        // Fetch from Google Docs API
        const docs = google.docs({ version: 'v1', auth: authClient });
        const response = await docs.documents.get({ documentId: docInfo.id });
        googleDoc = response.data;

        // Export to HTML
        const outputDir = path.join(__dirname, 'fixtures', docKey);
        try {
          execSync(`node gdocs-me-up.js ${docInfo.id} ${outputDir}`, {
            cwd: path.join(__dirname, '..'),
            stdio: 'pipe'
          });
        } catch (error) {
          throw new Error(`Export failed: ${error.message}`);
        }

        // Load exported HTML
        const htmlPath = path.join(outputDir, 'index.html');
        exportedHTML = fs.readFileSync(htmlPath, 'utf8');
        $ = cheerio.load(exportedHTML);
      });

      test('has correct document title', () => {
        const title = $('title').text();
        expect(title).toBe(googleDoc.title);
      });

      test('preserves heading hierarchy', () => {
        const apiHeadings = extractHeadingsFromAPI(googleDoc);
        const htmlHeadings = extractHeadingsFromHTML($);

        // Count by level
        const apiCounts = countByLevel(apiHeadings);
        const htmlCounts = countByLevel(htmlHeadings);

        console.log('API headings by level:', apiCounts);
        console.log('HTML headings by level:', htmlCounts);

        // Check that we have headings
        expect(htmlHeadings.length).toBeGreaterThan(0);

        // Main headings (h1/h2) should match or be close
        if (apiCounts.h1) {
          expect(htmlCounts.h1).toBeGreaterThanOrEqual(apiCounts.h1 * 0.8);
        }
      });

      test('exports all images', () => {
        const apiImageCount = countImagesFromAPI(googleDoc);
        const htmlImageCount = $('img').length;

        console.log(`Images - API: ${apiImageCount}, HTML: ${htmlImageCount}`);

        // Allow some tolerance for positioned objects that may render differently
        // Most documents should have all images, but positioned objects are tricky
        if (apiImageCount > 0) {
          expect(htmlImageCount).toBeGreaterThanOrEqual(Math.max(1, apiImageCount * 0.8));
        } else {
          expect(htmlImageCount).toBe(0);
        }
      });

      test('preserves all links', () => {
        const apiLinks = extractLinksFromAPI(googleDoc);
        const htmlLinks = extractLinksFromHTML($);

        console.log(`Links - API: ${apiLinks.length}, HTML: ${htmlLinks.length}`);

        // Check that most links are preserved (some tolerance for anchors, etc.)
        expect(htmlLinks.length).toBeGreaterThanOrEqual(apiLinks.length * 0.9);

        // Sample check: verify first few links match
        const sampleSize = Math.min(5, apiLinks.length);
        for (let i = 0; i < sampleSize; i++) {
          const apiLink = apiLinks[i];
          const matchingHtmlLink = htmlLinks.find(hl => hl.url === apiLink.url);
          expect(matchingHtmlLink).toBeDefined();
        }
      });

      test('preserves text content', () => {
        const apiText = extractTextFromAPI(googleDoc);
        const htmlText = extractTextFromHTML($);

        // Compare word counts (allow some variance for whitespace)
        const apiWords = apiText.split(/\s+/).filter(w => w.length > 0);
        const htmlWords = htmlText.split(/\s+/).filter(w => w.length > 0);

        console.log(`Words - API: ${apiWords.length}, HTML: ${htmlWords.length}`);

        // HTML should have at least 95% of the words from API
        expect(htmlWords.length).toBeGreaterThanOrEqual(apiWords.length * 0.95);

        // Sample check: first 100 words should appear in HTML
        const sampleWords = apiWords.slice(0, 100);
        const matchCount = sampleWords.filter(word =>
          htmlText.toLowerCase().includes(word.toLowerCase())
        ).length;

        expect(matchCount / sampleWords.length).toBeGreaterThan(0.9);
      });

      test('preserves list structures', () => {
        const apiLists = countListsFromAPI(googleDoc);
        const htmlLists = {
          ul: $('ul').length,
          ol: $('ol').length,
          li: $('li').length
        };

        console.log('Lists - API:', apiLists);
        console.log('Lists - HTML:', htmlLists);

        // Check list items are preserved
        if (apiLists.items > 0) {
          expect(htmlLists.li).toBeGreaterThanOrEqual(apiLists.items * 0.9);
        }
      });

      test('generates valid HTML structure', () => {
        expect(exportedHTML).toMatch(/<!DOCTYPE html>/i);
        expect(exportedHTML).toContain('<html');
        expect(exportedHTML).toContain('</html>');
        expect(exportedHTML).toContain('<head>');
        expect(exportedHTML).toContain('</head>');
        expect(exportedHTML).toContain('<body>');
        expect(exportedHTML).toContain('</body>');
      });

      test('applies formatting correctly', () => {
        const boldCount = $('.bold, strong, b').length;
        const italicCount = $('.italic, em, i').length;
        const underlineCount = $('.underline, u').length;

        console.log(`Formatting - Bold: ${boldCount}, Italic: ${italicCount}, Underline: ${underlineCount}`);

        // Just verify we have some formatting (exact counts are hard to match)
        const totalFormatting = boldCount + italicCount + underlineCount;
        expect(totalFormatting).toBeGreaterThan(0);
      });

      test('HTML output matches snapshot (regression detection)', () => {
        // This catches ANY changes to the HTML output
        // If this test fails, review the diff:
        // - If change is intentional: npm test -- -u (update snapshot)
        // - If change is a bug: fix the export code

        // Normalize dynamic content that changes between runs
        const normalizedHTML = exportedHTML
          // Remove absolute file paths
          .replace(/file:\/\/[^\s"]+/g, 'file://PATH')
          // Remove timestamps if any
          .replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/g, 'TIMESTAMP')
          // Normalize whitespace for consistent comparison
          .replace(/\s+/g, ' ')
          .trim();

        expect(normalizedHTML).toMatchSnapshot();
      });
    });
  }
});

// ============================================================================
// Helper Functions - Extract data from Google Docs API
// ============================================================================

function extractHeadingsFromAPI(doc) {
  const headings = [];
  const content = doc.body?.content || [];

  for (const element of content) {
    if (element.paragraph) {
      const namedType = element.paragraph.paragraphStyle?.namedStyleType;
      if (namedType && (namedType === 'TITLE' || namedType === 'SUBTITLE' || namedType.startsWith('HEADING_'))) {
        const text = element.paragraph.elements
          ?.map(e => e.textRun?.content || '')
          .join('')
          .trim();

        if (text) {
          headings.push({
            type: namedType,
            text
          });
        }
      }
    }
  }

  return headings;
}

function countImagesFromAPI(doc) {
  let count = 0;
  const content = doc.body?.content || [];

  for (const element of content) {
    if (element.paragraph?.elements) {
      for (const e of element.paragraph.elements) {
        if (e.inlineObjectElement) {
          count++;
        }
      }
    }
  }

  // Also count positioned objects
  if (doc.positionedObjects) {
    count += Object.keys(doc.positionedObjects).length;
  }

  return count;
}

function extractLinksFromAPI(doc) {
  const links = [];
  const content = doc.body?.content || [];

  for (const element of content) {
    if (element.paragraph?.elements) {
      for (const e of element.paragraph.elements) {
        if (e.textRun?.textStyle?.link?.url) {
          links.push({
            url: e.textRun.textStyle.link.url,
            text: e.textRun.content?.trim() || ''
          });
        }
      }
    }
  }

  return links;
}

function extractTextFromAPI(doc) {
  let text = '';
  const content = doc.body?.content || [];

  for (const element of content) {
    if (element.paragraph?.elements) {
      for (const e of element.paragraph.elements) {
        if (e.textRun?.content) {
          text += e.textRun.content;
        }
      }
    }
  }

  return text;
}

function countListsFromAPI(doc) {
  const listIds = new Set();
  let itemCount = 0;
  const content = doc.body?.content || [];

  for (const element of content) {
    if (element.paragraph?.bullet) {
      listIds.add(element.paragraph.bullet.listId);
      itemCount++;
    }
  }

  return {
    lists: listIds.size,
    items: itemCount
  };
}

// ============================================================================
// Helper Functions - Extract data from HTML
// ============================================================================

function extractHeadingsFromHTML($) {
  const headings = [];
  $('h1, h2, h3, h4, h5, h6').each((i, el) => {
    headings.push({
      level: el.name,
      text: $(el).text().trim()
    });
  });
  return headings;
}

function extractLinksFromHTML($) {
  const links = [];
  $('a').each((i, el) => {
    const url = $(el).attr('href');
    const text = $(el).text().trim();
    if (url) {
      links.push({ url, text });
    }
  });
  return links;
}

function extractTextFromHTML($) {
  // Get all text, removing extra whitespace
  return $('body').text().replace(/\s+/g, ' ').trim();
}

function countByLevel(headings) {
  const counts = {};
  for (const h of headings) {
    const level = h.type || h.level;
    counts[level] = (counts[level] || 0) + 1;
  }
  return counts;
}
