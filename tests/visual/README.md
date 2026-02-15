# Visual Testing & Comparison

This directory contains visual testing tools that compare Google Docs output with our HTML export to identify differences and guide improvements.

## Approach

Instead of pixel-perfect comparison (which fails due to font rendering, anti-aliasing, etc.), we use **smart semantic comparison**:

### What We Compare

1. **Structure** - Element counts (headings, paragraphs, lists, images)
2. **Content** - Text content, headings, links
3. **Layout** - Spacing, sizing, positioning metrics
4. **Visual Elements** - Image sizing, rendering, positioning
5. **Side-by-Side Screenshots** - For human review

### What We Generate

- ✅ Actionable recommendations for improvements
- 📊 Detailed comparison metrics
- 📸 Side-by-side screenshots
- 📋 HTML report for easy review
- 📈 JSON data for programmatic analysis

## Usage

```bash
# Run visual comparison tests
npm run test:visual

# Run with Playwright UI (interactive)
npm run test:visual:ui

# Run specific test
npx playwright test tests/visual/compare-smart.test.js
```

## Output

After running tests, check:

```
tests/visual/analysis/<doc-name>/
├── report.html          # Human-readable report (open in browser!)
├── report.json          # Machine-readable data
├── google-doc.png       # Screenshot of Google Doc
└── exported.png         # Screenshot of our export
```

## Configuration

Edit test document IDs in `compare-smart.test.js`:

```javascript
const TEST_DOCS = {
  'my-doc': 'YOUR_GOOGLE_DOC_ID',
  // Add more test documents here
};
```

## Iterative Improvement Workflow

1. **Run comparison**: `npm run test:visual`
2. **Review report.html**: Open in browser to see differences
3. **Check recommendations**: Listed at top of report
4. **Make improvements**: Update export code based on findings
5. **Re-run comparison**: Verify improvements
6. **Repeat** until satisfied with output quality

## Example Recommendations

The system automatically detects issues like:

- ✅ Images too wide - need container constraints
- ✅ Paragraph spacing too tight - adjust line-height
- ✅ Missing images - check positioned objects support
- ✅ Missing lists - check list rendering
- ✅ Text content mismatch - check extraction logic

## Adding New Checks

To add new comparison checks, edit `compare-smart.test.js`:

```javascript
// Add to analyzeStructure()
async function analyzeStructure(page) {
  return await page.evaluate(() => {
    return {
      // ... existing checks
      yourNewCheck: document.querySelectorAll('your-selector').length
    };
  });
}

// Add to recommendations
function generateRecommendations(structure, layout, images) {
  if (structure.yourNewCheck === 0) {
    recommendations.push('Your recommendation message');
  }
}
```

## Tips

- **Start with one document**: Get it looking good before testing many
- **Focus on structure first**: Element counts, hierarchy
- **Then content**: Text accuracy, links
- **Finally layout**: Spacing, sizing - most subjective
- **Use screenshots**: For final visual verification
- **Iterate quickly**: Run tests frequently during development

## Future Enhancements

- [ ] Automated regression testing in CI/CD
- [ ] Perceptual diff algorithms for layout comparison
- [ ] Text diff visualization (like git diff)
- [ ] Performance metrics (export time, file size)
- [ ] Cross-browser testing (Chrome, Firefox, Safari)
- [ ] Mobile viewport testing
- [ ] Accessibility analysis (ARIA, contrast, etc.)
