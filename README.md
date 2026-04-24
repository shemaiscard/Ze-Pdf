# ZePdf

Universal document converter. Convert between 20+ formats and split or merge PDFs — fast, clean, and browser-based.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Overview

ZePdf is a client-side web application deployed on Netlify. Document conversions are handled server-side by CloudConvert. PDF split and merge run entirely in the browser using pdf-lib with no external API needed.

---

## Features

### Document Conversion

- Supported formats: PDF, DOCX, DOC, PPTX, XLSX, RTF, ODT, ODP, ODS, EPUB, MOBI, JPG, PNG, WEBP, GIF, TXT
- Batch conversion — upload multiple files and convert them all at once
- Per-file format selector — assign a different target format to each file
- Real-time progress bar per file
- Auto-download of converted output

### PDF Tools

- Merge — combine multiple PDFs into one, drag rows to reorder before merging
- Split by every page — extract each page as a separate PDF file
- Split by page range — extract specific pages using a range like `1-3, 5, 7-9`
- Split every N pages — divide the PDF into equal-sized chunks

### Interface

- Drag and drop file upload
- Dark and light theme toggle
- Conversion history panel stored in localStorage, up to 50 entries
- Daily conversion counter
- Toast notifications
- Responsive layout

---

## Deployment

### Deploy to Netlify from GitHub

1. Fork or clone this repository.
2. Go to [app.netlify.com](https://app.netlify.com), create a new site from GitHub, and select this repository.
3. In your Netlify site settings, go to **Environment variables** and add:

   ```
   CLOUDCONVERT_API_KEY = your_cloudconvert_api_key
   ```

4. Trigger a deploy. Netlify will inject the API key at build time via the build command in `netlify.toml`.

The CloudConvert API key is never exposed in the repository or in the browser. It is written into a generated `_config.js` file during the Netlify build and served as a static asset.

### Get a CloudConvert API key

Create a free account at [cloudconvert.com](https://cloudconvert.com/api/v2).
The free tier provides 25 conversion minutes per day.

---

## Project Structure

```
Ze-Pdf/
├── index.html       Main application
├── style.css        Design system and themes
├── app.js           Conversion logic (CloudConvert)
├── pdf-tools.js     PDF split and merge (pdf-lib, client-side)
├── _config.js       Generated at build by Netlify (do not edit)
├── netlify.toml     Build command and security headers
├── favicon/         App icons and web manifest
└── README.md
```

---

## Tech Stack

- HTML, CSS, JavaScript — no framework, no build step
- [CloudConvert API v2](https://cloudconvert.com/api/v2) — document conversion
- [pdf-lib](https://pdf-lib.js.org/) — client-side PDF manipulation
- [Netlify](https://netlify.com) — hosting, build, environment variables

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

Developed by [Giscard](https://github.com/shemaiscard)
