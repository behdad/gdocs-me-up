Note: The code in this repository, as well as this README, was generated iteratively with ChatGPT o1 and o1pro for my own needs. I have started making manual changes now, and PRs are welcome. Sharing because caring. -behdad

# GDocs-Me-Up: A Google Docs High-Fidelity Exporter

A Node.js script that exports Google Docs to HTML+CSS with high fidelity, preserving essential formatting like headings, line spacing, alignment, bullet lists, images, and more. Perfect for creating an **offline** or **self-hosted** version of your docs that closely matches the original.

---

## Overview

**Why this script?** Because simpler exports often lose alignment, line spacing, or right-to-left details. This script pulls detailed styling info (like heading levels, inline font sizes, doc-based image sizes, and bullet indentation) directly from the Google Docs API. It then **merges** that styling into HTML and **inserts** a corresponding CSS that replicates Google Docs formatting while also **neutralizing** default browser quirks.

### What It Does

1. **Named Styles**: Detects **Title**, **Subtitle**, **HEADING_1..6**, and applies them to HTML headings (H1..H6) or custom classes.
2. **Line Spacing & Margins**: Honors `paragraphStyle.lineSpacing`, `spaceAbove`, `spaceBelow`, indentation, alignment.
3. **Right-to-Left**: If the doc says a paragraph is RTL, we add `dir="rtl"` and flip alignment (START → right).
4. **Tables**: GDocs tables become `<table>` with `<tr>` and `<td>`, keeping paragraph formatting in each cell.
5. **Images**: Exports inline and positioned images at their document dimensions. Original JPEG, GIF, WebP, SVG, and other recognized formats keep their real extension; opaque photographic PNGs become high-quality JPEGs when that saves at least 20%. Transparent PNGs remain PNG.
6. **TOC**: If your doc has a table of contents, we export it in a `<div class="doc-toc">`, indenting each line by its heading level.
7. **Bullet/Numbered Lists**: Detects all GDocs list styles (disc, circle, square, dash bullets; decimal, roman, alphabetic numbering) with proper nesting. RTL lists use `<ul dir="rtl">` so bullets align on the right.
8. **Google Fonts**: Gathers unique fonts used in the doc. Inserts a `<link>` to [fonts.googleapis.com](https://fonts.googleapis.com/) so text families match.
9. **Neutralized Headings**: Browsers normally inflate `<h3>`. We override heading tags (`h1..h6 { font-size: 1em }`) so Google Docs' inline style alone sets the final size.  
10. **Semantic Main Content**: Wraps the exported document in `<main class="doc-content">` so injected navigation can remain outside the page's primary-content landmark.

---

## Installation

1. **Prerequisites**:
   - **Node.js** (v20.9 or later).
   - Run `npm install` to install the declared dependencies.
   - A **Google Cloud** service account JSON file with read permissions on the doc.

2. **Get the Script**:
   - Download or clone this repository.
   - Put `service_account.json` beside `gdocs-me-up.js`, or set
     `SERVICE_ACCOUNT_KEY_FILE` to another path. The default works regardless of
     the directory from which the exporter is invoked.

3. **Authenticate**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/), enable **Docs API** + **Drive API**.
   - Create a service account with suitable permissions.
   - Download the JSON key file as `service_account.json`.
   - Make sure your doc is accessible by that service account (share it if needed).

---

## Usage

```bash
node gdocs-me-up.js <DOC_ID> <OUTPUT_DIR> [options]
```

For help:
```bash
node gdocs-me-up.js --help
```

- **`<DOC_ID>`**: The unique ID from your doc's URL. For example:
  ```
  https://docs.google.com/document/d/1AbCdE-FgHiJKlMnOpQRs7TuVMue/edit
                ^^^^^^^^^^^^^^^^^^^^^
  ```
- **`<OUTPUT_DIR>`**: The folder where the script will write `index.html` and an `images/` subfolder.
- **`--stylesheet <HREF>`**: Inserts an external stylesheet after the generated styles. Repeat the option to include multiple stylesheets in argument order. Hrefs are resolved relative to the generated HTML file.
- **`--script <SRC>`**: Inserts an external script immediately before `</body>`. Repeat the option to include multiple scripts in argument order. Sources are resolved relative to the generated HTML file.
- **`--html-name <NAME>`**: Replaces the default `index.html` filename.
- **`--images-dir <NAME>`**: Replaces the default `images` directory name.

**Example**:

```bash
node gdocs-me-up.js 1AbCdE-FgHiJK docs_export
```

To apply stylesheets and scripts stored next to the output directory:

```bash
node gdocs-me-up.js 1AbCdE-FgHiJK docs_export \
  --stylesheet ../style.css \
  --stylesheet ../theme.css \
  --script ../script.js \
  --html-name article.html \
  --images-dir assets
```

On completion:
- **`docs_export/index.html`**: Your doc in near-pixel HTML+CSS fidelity.  
- **`docs_export/images/`**: Downloaded images.  

Open `docs_export/index.html` in your browser. You'll see headings, bullet-lists, alignment, images, and more, closely mirroring the original doc.

---

## Testing

The project includes comprehensive testing to ensure export quality:

### Unit Tests
```bash
npm test
```
Tests core utility functions (escapeHtml, color conversion, list detection, etc.) with 97% coverage.

### Content Verification Tests
```bash
npm test tests/content-verification.test.js
```
Compares Google Docs API data with exported HTML to verify:
- Text content accuracy
- Link preservation
- Image export
- Heading hierarchy
- List structures

Uses two stable reference documents as golden standards.

### Snapshot Tests
Full HTML regression detection - catches **any** change to output:
```bash
npm test  # Runs automatically with other tests
npm test -- -u  # Update snapshots after intentional changes
```

### Visual Comparison Tests
```bash
npm run test:visual
```
Generates side-by-side screenshots and analysis reports comparing Google Docs with exported HTML. Reports include:
- Structure analysis (element counts)
- Layout metrics (spacing, sizing)
- Actionable recommendations

See `tests/visual/README.md` for details.

To run the local visual-comparison corpus:

```bash
GDOCS_CORPUS_DIR=/path/to/fixtures npm run compare:corpus
```

The corpus runner deduplicates document IDs, captures the Google preview and local export at the same viewport, and writes screenshots, a contact sheet, and a JSON report under `tests/visual/corpus/`. Use `-- --limit=5` for a short run or `-- --names=story,butterflies` to select fixture names. Fixture files and their location are not part of the repository.

---

## Key Details

1. **Line Spacing**: The script reads `paragraphStyle.lineSpacing` and maps Google Docs' font-dependent metrics to browser line boxes. It also applies `spaceAbove` + `spaceBelow` as `margin-top` + `margin-bottom`.

2. **Right-to-Left Paragraphs**: If `paragraphStyle.direction = RIGHT_TO_LEFT`, we add `dir="rtl"`. If alignment=START, it becomes `right`; alignment=END => `left`. Lists also carry `dir="rtl"` so bullets go on the right side.  

3. **Images**: Supports both inline images and positioned objects (header photos, wrapped images). Images are constrained to container width while retaining their explicit aspect ratio. We read size info from both `imageProperties.size` and `embedded.size`, converting points to pixels (~1.333 ratio) and respecting transforms. Positioned objects render at their anchor paragraph. Opaque PNG photographs are tested against a quality-92 JPEG candidate and converted only when the candidate is at least 20% smaller.

4. **TOC Indentation**: For each line in the doc’s table of contents, the script checks the heading level of the link target. It then adds a `<div class="toc-level-3">` (for example) with a margin-left rule in the CSS.  

5. **Compact Semantic Markup**: Google Docs often splits text into many runs. Consecutive compatible runs are merged, common paragraph formatting is inherited, emphasis uses semantic tags such as `<strong>` and `<em>`, and repeated declarations are deduplicated into generated CSS classes. Exported content does not repeat inline `style` attributes.

6. **Heading Size**: We reset browser heading defaults, then a generated class supplies the document's exact size and weight without the browser multiplying them.

7. **Fonts**: If your doc uses “Roboto” and “Lato,” we add a single `<link>` to `https://fonts.googleapis.com/css2?family=Roboto&family=Lato&display=swap`, letting the final HTML use those fonts.

---

## Customizing

- **Force a Different Column Width**: Edit `computeDocContainerWidth()` to remove the `+ 50`, or set a fixed width.  
- **Line Spacing**: If you want a global `line-height:1.2`, remove or comment out the lines in `renderParagraph` referencing `paragraphStyle.lineSpacing`.  
- **Heading Tags**: If you’d rather not use `<h1>.. <h6>`, replace them with `<p class="doc-heading-level-X">` in the code. Then style them in CSS as you like.  

---

## Troubleshooting

1. **Invalid Grant / 401**: Check your service account JSON, or ensure the doc is shared with your service account email.  
2. **Images All Full-Width**: Possibly the doc’s stored size is as wide as the page. Shrink them in GDocs or scale them down.  
3. **TOC Not Indented**: Make sure your doc has headings labeled `HEADING_1..6`. If your doc uses custom styles, the script may not see them as headings.  
4. **H3 Still Big**: Confirm the code’s `<h3>` CSS override is present, or remove any conflicting styles from your own stylesheet.  
5. **Using a Different Auth**: If you want user-based OAuth, adapt `getAuthClient()` to your flow.

---

## Example

**Doc**: “My Example Document” with:
- Heading 3 at 14pt
- Right-to-left paragraphs
- A table of contents
- Several images scaled to 50%

**Command**:

```bash
node gdocs-me-up.js 1XYZabc docs_export
```

**Result**:
- `docs_export/index.html`: Headings, bullet-lists, alignment, images at half-size, lines spaced as in doc, etc.  
- `docs_export/images/`: Images in their source format, with space-saving JPEG conversion for suitable opaque PNG photographs.
- The TOC lines are indented by heading level.

Open the HTML in your browser or upload to a simple web server. Should be extremely close to the Google Doc’s layout, including RTL paragraphs and scaled images.

---

## Contributing

1. **Fork** or clone this repository.  
2. Modify the script (e.g., add footnote support or custom style merges).  
3. **Submit a Pull Request** describing your changes, or open an issue with suggestions.  

We welcome improvements or bug fixes. This script is licensed under **MIT**, so feel free to adapt or include it in your projects, with attribution appreciated.

---

**Thanks** for checking out **GDocs-Me-Up**! We hope it helps you create accurate offline or self-hosted versions of your docs. If you have suggestions, issues, or ideas, please open an issue or PR. Happy exporting!
