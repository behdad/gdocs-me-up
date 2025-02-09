# Google Docs High-Fidelity Export

A Node.js script that exports a Google Doc to **HTML + CSS** with near-pixel fidelity, preserving headings, images, line spacing, bullet lists, alignment, right-to-left paragraphs, and more.

## Features

- **Title, Subtitle, Headings** (H1..H6) recognized from Google Docs named styles.
- **Doc-based alignment** (CENTER, JUSTIFIED, flipping START/END in RTL).
- **Line Spacing** (paragraphStyle.lineSpacing), plus margin spaceAbove / spaceBelow.
- **Right-to-left** text with `dir="rtl"`, bullet lists in RTL if needed.
- **Images** at their true doc-based dimensions, including transform scaling.
- **Table of Contents** with hierarchical indentation (if doc has a TOC).
- **Tables** using `<table>`/`<tr>`/`<td>`.
- **Merging consecutive text runs** to avoid overly-fragmented HTML.
- **Google Fonts** link for any distinct fonts used.
- Minimal `.htaccess` setting `DirectoryIndex index.html`.

## Requirements

- **Node.js** (v14 or later recommended)
- **NPM package**: `googleapis` (`npm install googleapis`)
- **Service Account JSON** for accessing the doc.  
  - Place credentials in a file named `service_account.json` (or update the script config).

## Usage

1. **Download** or clone this repo.  
2. Install dependencies:

   ```bash
   npm install googleapis

