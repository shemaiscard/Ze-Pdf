# ZePdf — Universal Document Converter

> Convert any document instantly. PDF, Word, Excel, PowerPoint, Images, eBooks and more — powered by CloudConvert.

[![Netlify Status](https://img.shields.io/badge/deploy-netlify-00C7B7?logo=netlify)](https://netlify.com)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## ✨ Features

### 📄 Document Conversion (CloudConvert)
- **20+ formats supported:** PDF, DOCX, DOC, PPTX, XLSX, RTF, ODT, ODP, ODS, EPUB, MOBI, JPG, PNG, WEBP, GIF, TXT
- **Batch conversion** — upload multiple files and convert them all at once
- **Per-file format selector** — assign different target formats to each file
- **Real-time progress bars** per file
- **Auto-download** of converted files

### ✂️ PDF Tools (100% Client-Side, No API Key)
- **Merge PDFs** — combine multiple PDFs into one, drag to reorder
- **Split PDF** — by every page, custom page range (e.g. `1-3, 5, 7-9`), or every N pages

### 🖥️ UI & UX
- **Drag & Drop** upload zone
- **Dark / Light theme** toggle (teal/emerald accent)
- **Conversion history** panel (up to 50 entries, stored in localStorage)
- **Daily conversion counter**
- **Toast notifications** for all events
- **Fully responsive** — works on mobile

---

## 🚀 Getting Started

### 1. Get a free CloudConvert API key
Visit [cloudconvert.com/api/v2](https://cloudconvert.com/api/v2) and create a free account.  
The free tier gives you **25 minutes/day** of conversion time — enough for ~300 small file conversions.

### 2. Deploy to Netlify (Recommended)
1. Fork or clone this repository.
2. Drag the project folder to [app.netlify.com/drop](https://app.netlify.com/drop).
3. Your site is live instantly!

### 3. Run locally
Simply open `index.html` in any modern browser — no build step needed.

---

## 🔑 API Key Security
Your CloudConvert API key is stored **only in your browser's localStorage**. It is never sent to any server other than CloudConvert directly. The key is not included in the repository.

---

## 🗂️ Project Structure

```
Ze-Pdf-main/
├── index.html       # Main app structure
├── style.css        # Design system (dark/light theme, all components)
├── app.js           # Conversion logic (CloudConvert API)
├── pdf-tools.js     # PDF split/merge (pdf-lib, client-side)
├── netlify.toml     # Netlify deploy config + security headers
└── README.md
```

---

## 🛠️ Tech Stack
- **Vanilla HTML / CSS / JavaScript** — no framework, no build step
- **[CloudConvert API v2](https://cloudconvert.com/api/v2)** — document conversion engine
- **[pdf-lib](https://pdf-lib.js.org/)** — client-side PDF manipulation (split/merge)
- **Netlify** — hosting + deployment

---

## 📜 Credits
Developed by **Giscard**
