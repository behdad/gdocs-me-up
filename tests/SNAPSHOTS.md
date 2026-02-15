# Snapshot Testing

The content verification tests include **snapshot testing** - saving the complete HTML output to detect ANY changes.

## How It Works

1. **First run**: Exports HTML and saves as a "golden snapshot"
2. **Future runs**: Compares new export against the snapshot
3. **If different**: Test fails and shows the diff
4. **Review**: You decide if the change is good or bad

## Snapshot Files

Located in `tests/__snapshots__/content-verification.test.js.snap`:
- Behdad's Story (full HTML ~60KB)
- State of Text Rendering 2024 (full HTML ~350KB)

**These files ARE committed to git** - they're the reference for detecting regressions.

## When Tests Fail

If snapshot tests fail, you'll see a diff showing what changed:

```
- Expected
+ Received

- <p style="margin-top:10px;">
+ <p style="margin-top:12px;">
```

### If Change Is a Bug (Unintended)
Fix your code and re-run tests:
```bash
npm test
```

### If Change Is Intentional (Improvement)
Update the snapshots:
```bash
npm test -- -u
# or
npm test -- --updateSnapshot
```

Then commit the updated snapshots:
```bash
git add tests/__snapshots__
git commit -m "Update snapshots for [reason]"
```

## What Gets Detected

Snapshot tests catch **everything**:
- ✅ Changed CSS styles
- ✅ Different HTML structure
- ✅ Added/removed elements
- ✅ Changed attributes
- ✅ Text content changes
- ✅ Link URLs
- ✅ Image paths

This is more comprehensive than the other content tests, which only check structure/content.

## Best Practices

### DO Update Snapshots When:
- ✅ You improved the HTML output
- ✅ You fixed a formatting bug
- ✅ You added a new feature
- ✅ You changed CSS generation

### DON'T Update Snapshots If:
- ❌ Tests are failing for unknown reasons
- ❌ You didn't intend to change output
- ❌ You haven't reviewed the diff

### Always:
1. **Review the diff** - understand what changed
2. **Verify in browser** - check the actual HTML looks right
3. **Update snapshots** - only if changes are correct
4. **Commit snapshots** - so others have the updated baseline

## Example Workflow

```bash
# Make changes to export code
vim gdocs-me-up.js

# Run tests
npm test

# See snapshot test fail with diff
# Review the diff carefully

# If change looks good:
npm test -- -u

# Verify the updated export
open tests/fixtures/behdad-story/index.html

# If it looks right, commit
git add tests/__snapshots__
git commit -m "Improve paragraph spacing"
```

## Debugging Snapshot Failures

If snapshot tests are failing but content tests pass:
1. The content is correct, but formatting/structure changed
2. Review the diff in the test output
3. Check if CSS or HTML structure was modified
4. Decide if it's an improvement or regression

## When to Skip Snapshot Updates

Sometimes you're experimenting and don't want to update snapshots:
```bash
# Skip snapshot tests temporarily
npm test -- --testPathIgnorePatterns=content-verification
```

## Snapshot Size

Current snapshots are ~400KB total. This is fine for:
- Git (text compresses well)
- CI/CD (fast to load)
- Reviews (can diff in GitHub)

If snapshots get too large (>1MB), consider:
- Testing fewer documents
- Normalizing more dynamic content
- Using content tests instead of snapshots
