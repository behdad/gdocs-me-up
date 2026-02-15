# Testing Guide

## Running Tests

```bash
# Run all tests
npm test

# Run tests in watch mode (auto-rerun on file changes)
npm run test:watch

# Run tests with coverage report
npm run test:coverage
```

## Test Structure

### Unit Tests
- `lib/utils.test.js` - Tests for utility functions (escapeHtml, ptToPx, rgbToHex, etc.)
- `lib/lists.test.js` - Tests for list handling functions

Coverage for utility modules is excellent (>96% for utils.js).

### Integration Tests
Integration tests would export full documents and verify the output. These require:
- Google Docs API credentials
- Test documents
- More complex setup

Example integration test structure (not yet implemented):

```javascript
describe('Document Export', () => {
  test('exports document with images', async () => {
    const docId = 'TEST_DOC_ID';
    const outputDir = path.join(__dirname, 'fixtures/output');

    await exportDocToHTML(docId, outputDir);

    const html = fs.readFileSync(path.join(outputDir, 'index.html'), 'utf8');
    expect(html).toContain('<img');
    expect(html).toContain('<!DOCTYPE html>');
  });
});
```

## Test Coverage

Current coverage (as of last run):
- **lib/utils.js**: ~97% - excellent coverage of all utility functions
- **lib/lists.js**: ~28% - covers core functions, integration tests would improve this
- **lib/constants.js**: 100% - just exports
- **Main file**: 0% - would require integration tests

## Adding New Tests

### For new utility functions:
1. Add test file in `lib/*.test.js`
2. Test edge cases (null, undefined, empty values)
3. Test normal operation
4. Test error conditions

### For integration tests:
1. Create test documents with known content
2. Export and verify key elements exist
3. Use snapshot testing for HTML output
4. Consider visual regression testing with headless browser

## Future Improvements

1. **Integration tests**: Export known documents and verify output structure
2. **Snapshot tests**: Capture HTML output and detect regressions
3. **Visual tests**: Use Playwright/Puppeteer to render and compare screenshots
4. **E2E tests**: Test full workflow including auth and document fetching
5. **CI/CD**: Run tests automatically on every commit
