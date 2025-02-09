# Google Docs High-Fidelity Export

A Node.js script that exports Google Docs to HTML+CSS with **near-pixel fidelity**, preserving essential formatting like headings, line spacing, alignment, bullet lists, images, and more. Perfect for creating an **offline** or **self-hosted** version of your docs that closely matches the original.

---

## Overview

**Why this script?** Because simpler exports often lose alignment, line spacing, or right-to-left details. This script pulls detailed styling info (like heading levels, inline font sizes, doc-based image sizes, and bullet indentation) directly from the Google Docs API. It then **merges** that styling into HTML and **inserts** a corresponding CSS that replicates Google Docs formatting while also **neutralizing** default browser quirks.

### What It Does

1. **Named Styles**: Detects **Title**, **Subtitle**, **HEADING_1..6**, and applies them to HTML headings (H1..H6) or custom classes.  
2. **Line Spacing & Margins**: Honors `paragraphStyle.lineSpacing`, `spaceAbove`, `spaceBelow`, indentation, alignment.  
3. **Right-to-Left**: If the doc says a paragraph is RTL, we add `dir="rtl"` and flip alignment (START → right).  
4. **Tables**: GDocs tables become `<table>` with `<tr>` and `<td>`, keeping paragraph formatting in each cell.  
5. **Images**: Exports each embedded image at the doc’s reported width/height (in pt → px), respecting scaling. Saves images in an `images/` folder.  
6. **TOC**: If your doc has a table of contents, we export it in a `<div class="doc-toc">`, indenting each line by its heading level.  
7. **Bullet/Numbered Lists**: Detects GDocs bullet styles, outputting `<ul>` / `<ol>`. If the doc is RTL, we do `<ul dir="rtl">` so bullets align on the right.  
8. **Google Fonts**: Gathers unique fonts used in the doc. Inserts a `<link>` to [fonts.googleapis.com](https://fonts.googleapis.com/) so text families match.  
9. **Neutralized Headings**: Browsers normally inflate `<h3>`. We override heading tags (`h1..h6 { font-size: 1em }`) so Google Docs’ inline style alone sets the final size.  
10. **Minimal .htaccess**: Writes a `DirectoryIndex index.html`, so your exported folder works out-of-the-box if served by Apache.

---

## Installation

1. **Prerequisites**:
   - **Node.js** (v14 or later).
   - `npm install googleapis`.
   - A **Google Cloud** service account JSON file with read permissions on the doc.

2. **Get the Script**:
   - Download or clone this repository.
   - Ensure `google-docs-high-fidelity-export.js` and your `service_account.json` are in the same folder (or update the path in the script).

3. **Authenticate**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/), enable **Docs API** + **Drive API**.
   - Create a service account with suitable permissions.
   - Download the JSON key file as `service_account.json`.
   - Make sure your doc is accessible by that service account (share it if needed).

---

## Usage

```bash
node google-docs-high-fidelity-export.js <DOC_ID> <OUTPUT_DIR>
```

- **`<DOC_ID>`**: The unique ID from your doc’s URL. For example:
  ```
  https://docs.google.com/document/d/1AbCdE-FgHiJKlMnOpQRs7TuVMue/edit
                ^^^^^^^^^^^^^^^^^^^^^
  ```
- **`<OUTPUT_DIR>`**: The folder where the script will write `index.html`, `.htaccess`, and an `images/` subfolder.

**Example**:

```bash
node google-docs-high-fidelity-export.js 1AbCdE-FgHiJK docs_export
```

On completion:
- **`docs_export/index.html`**: Your doc in near-pixel HTML+CSS fidelity.  
- **`docs_export/images/`**: Downloaded images.  
- **`docs_export/.htaccess`**: Minimal config to make `index.html` default.

Open `docs_export/index.html` in your browser. You’ll see headings, bullet-lists, alignment, images, and more, closely mirroring the original doc.

---

## Key Details

1. **Line Spacing**: The script reads `paragraphStyle.lineSpacing` (e.g., 100 = 1.0, 115 = 1.15, 200 = 2.0) and sets `line-height`. It also applies `spaceAbove` + `spaceBelow` as `margin-top` + `margin-bottom`.  

2. **Right-to-Left Paragraphs**: If `paragraphStyle.direction = RIGHT_TO_LEFT`, we add `dir="rtl"`. If alignment=START, it becomes `right`; alignment=END => `left`. Lists also carry `dir="rtl"` so bullets go on the right side.  

3. **Images**: We read `imageProperties` to get `width.magnitude` + `height.magnitude` (in points), multiply by ~1.333 to convert to px, and store them in `<img style="max-width: Xpx; max-height: Ypx;">`. If the doc scaled an image, we read `transform.scaleX/scaleY`.  

4. **TOC Indentation**: For each line in the doc’s table of contents, the script checks the heading level of the link target. It then adds a `<div class="toc-level-3">` (for example) with a margin-left rule in the CSS.  

5. **Merging Identical Runs**: Google Docs often splits text into multiple runs. If two consecutive runs share the same style (bold, color, font-size, etc.), we merge them to keep the final HTML lean.  

6. **Heading Size**: We override heading tags in CSS to `font-size: 1em; font-weight: normal;`. The doc sets an inline `font-size: 18pt;` (for example), so you get exactly 18pt, not 18pt multiplied by the browser’s default heading scale.  

7. **Fonts**: If your doc uses “Roboto” and “Lato,” we add a single `<link>` to `https://fonts.googleapis.com/css2?family=Roboto&family=Lato&display=swap`, letting the final HTML use those fonts.

---

## Customizing

- **Force a Different Column Width**: Edit `computeDocContainerWidth()` to remove the `+ 50`, or set a fixed width.  
- **Line Spacing**: If you want a global `line-height:1.2`, remove or comment out the lines in `renderParagraph` referencing `paragraphStyle.lineSpacing`.  
- **Images as Base64**: Set `EMBED_IMAGES_AS_BASE64 = true;`, so images are embedded inline instead of written to `images/`.  
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
node google-docs-high-fidelity-export.js 1XYZabc docs_export
```

**Result**:
- `docs_export/index.html`: Headings, bullet-lists, alignment, images at half-size, lines spaced as in doc, etc.  
- `docs_export/images/`: The images as `png`.  
- The TOC lines are indented by heading level.

Open the HTML in your browser or upload to a simple web server. Should be extremely close to the Google Doc’s layout, including RTL paragraphs and scaled images.

---

## Contributing

1. **Fork** or clone this repository.  
2. Modify the script (e.g., add footnote support or custom style merges).  
3. **Submit a Pull Request** describing your changes, or open an issue with suggestions.  

We welcome improvements or bug fixes. This script is licensed under **MIT**, so feel free to adapt or include it in your projects, with attribution appreciated.

---

**Thanks** for checking out **Google Docs High-Fidelity Export**! We hope it helps you create accurate offline or self-hosted versions of your docs. If you have suggestions, issues, or ideas, please open an issue or PR. Happy exporting!
