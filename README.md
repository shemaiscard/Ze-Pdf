# ZePdf

**Live Demo:** [zepdf.netlify.app](https://zepdf.netlify.app)

Universal document converter and PDF toolkit. Convert between 20+ formats and split or merge PDFs in the browser.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## Privacy

PDF split and merge run **entirely in your browser**, so those files never leave your device. Document conversions are processed server-side for performance.

## Features

### Document Conversion
- Convert between 20+ formats: PDF, DOCX, DOC, PPTX, XLSX, RTF, ODT, ODP, ODS, EPUB, MOBI, JPG, PNG, WEBP, GIF, TXT
- Batch convert: upload multiple files at once
- Per-file format selector
- Real-time progress bar with auto-download

### PDF Tools
- **Merge**: combine multiple PDFs; drag to reorder before merging
- **Split by page**: extract every page as a separate file
- **Split by range**: e.g. `1-3, 5, 7-9`
- **Split every N pages**: divide into equal chunks

### Interface
- Dark/light theme toggle
- Drag and drop upload
- Conversion history (up to 50 entries)
- Responsive, works on mobile and desktop

## Deploy on Netlify

1. Fork this repo and connect to [Netlify](https://app.netlify.com).
2. Add environment variable:
   ```
   CLOUDCONVERT_API_KEY = your_api_key
   ```

## Project Structure

```
Ze-Pdf/
├── index.html       Main application
├── style.css        Design system and themes
├── app.js           Conversion logic and UI
├── pdf-tools.js     Client-side PDF split/merge
├── netlify/         Serverless functions
├── netlify.toml     Netlify config and security headers
└── favicon/
```

## Tech Stack
- HTML, CSS, JavaScript (no framework)
- [pdf-lib](https://pdf-lib.js.org/): client-side PDF manipulation
- [Netlify](https://netlify.com): hosting and serverless functions

## License
MIT. Developed by [Giscard](https://github.com/shemaiscard)
