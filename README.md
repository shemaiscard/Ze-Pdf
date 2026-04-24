# ZePdf

Universal document converter built with vanilla JavaScript and the CloudConvert API.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Overview

ZePdf is a client-side web application that converts documents between 20+ formats and provides PDF split and merge tools — all without a backend server. Conversions are handled by the CloudConvert API. PDF split and merge run entirely in the browser using pdf-lib, requiring no API key.

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
- Toast notifications for all events
- Responsive layout for mobile and desktop

---

## Getting Started

### 1. Get a CloudConvert API key

Create a free account at [cloudconvert.com](https://cloudconvert.com/api/v2).  
The free tier provides 25 conversion minutes per day, which covers approximately 300 small file conversions.

### 2. Deploy to Netlify

Drag the project folder to [app.netlify.com/drop](https://app.netlify.com/drop).  
No build step is required. The site is live immediately.

### 3. Run locally

Open `index.html` in any modern browser. No installation or build step needed.

---

## API Key Security

The CloudConvert API key is stored only in the browser's localStorage. It is never sent to any server other than CloudConvert directly and is not stored in the repository.

---

## Project Structure

```
Ze-Pdf/
├── index.html       Main application structure
├── style.css        Design system, dark and light themes, all components
├── app.js           Conversion logic using the CloudConvert API
├── pdf-tools.js     PDF split and merge using pdf-lib, fully client-side
├── netlify.toml     Netlify deployment config and security headers
└── README.md
```

---

## Tech Stack

- HTML, CSS, JavaScript — no framework, no build step required
- [CloudConvert API v2](https://cloudconvert.com/api/v2) — document conversion engine
- [pdf-lib](https://pdf-lib.js.org/) — client-side PDF manipulation
- [Netlify](https://netlify.com) — hosting and deployment

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

Developed by Giscard
