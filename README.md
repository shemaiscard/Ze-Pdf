# ZePdf

Universal document converter and PDF toolkit. Convert between 20+ formats and split or merge PDFs — fast, clean, and browser-based.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## What it does

ZePdf lets you convert documents between popular formats and work with PDFs directly in your browser. No software to install, no account required.

Document conversions are processed server-side for speed and reliability. PDF split and merge run entirely in the browser — nothing is uploaded, nothing leaves your device.

---

## Features

### Document Conversion

- Convert between 20+ formats: PDF, DOCX, DOC, PPTX, XLSX, RTF, ODT, ODP, ODS, EPUB, MOBI, JPG, PNG, WEBP, GIF, TXT
- Batch convert — upload multiple files at once
- Per-file format selector — assign different target formats to each file
- Real-time progress bar per file with status updates
- Auto-download of converted output

### PDF Tools

- Merge — combine multiple PDFs into one; drag rows to reorder before merging
- Split by every page — extract each page as a separate file
- Split by page range — extract specific pages using a range like `1-3, 5, 7-9`
- Split every N pages — divide a PDF into equal-sized chunks

### Interface

- PDF Tools shortcut button in the navigation bar for quick access
- Drag and drop file upload
- Dark and light theme toggle
- Conversion history panel with up to 50 entries
- Daily conversion counter
- Toast notifications for all events
- Responsive layout for mobile and desktop

---

## Deployment

### Deploy to Netlify from GitHub

1. Fork or clone this repository.
2. Go to [app.netlify.com](https://app.netlify.com), create a new site from GitHub, and select this repository.
3. In your Netlify site settings, go to **Environment variables** and add:

   ```
   CLOUDCONVERT_API_KEY = your_api_key
   ```

4. Trigger a deploy. The build command in `netlify.toml` injects the key into `_config.js` at build time. Users never see or interact with the key.

---

## Project Structure

```
Ze-Pdf/
├── index.html       Main application
├── style.css        Design system and themes
├── app.js           Conversion logic and UI
├── pdf-tools.js     PDF split and merge (runs locally in browser)
├── _config.js       Generated at build by Netlify — do not edit
├── netlify.toml     Build command and security headers
├── favicon/         App icons and web manifest
└── README.md
```

---

## Tech Stack

- HTML, CSS, JavaScript — no framework, no build step
- [pdf-lib](https://pdf-lib.js.org/) — client-side PDF manipulation
- [Netlify](https://netlify.com) — hosting, build, environment variables

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

Developed by [Giscard](https://github.com/shemaiscard)
