ZEPDF - DOCUMENT CONVERTER & VIEWER
===================================

OVERVIEW
--------
ZePdf is a comprehensive desktop application built with Python and PySide6 (Qt) 
designed for advanced PDF manipulation and document conversion. It provides 
users with a high-performance graphical interface to view, split, merge, and 
convert documents between a wide variety of formats.

TECHNICAL CONSTRUCTION
----------------------
This project demonstrates advanced integration of several document processing 
libraries and multi-threaded GUI development.

1. ARCHITECTURE (Python & PySide6):
   - Multi-Threaded Engine: Utilizes a dedicated ConversionThread to handle 
     heavy processing tasks. This prevents UI freezing and allows users to 
     cancel long-running conversions in real-time.
   - Component-Based Design: Separates logic into specialized managers, 
     including a PreviewManager for document rendering and a DocumentConverter 
     for file transformations.
   - Zoomable Graphics Engine: Implements a custom ZoomableGraphicsView 
     using Qt’s Graphics View Framework, allowing for smooth, antialiased 
     rendering and mouse-wheel zooming.

2. DOCUMENT PROCESSING INTEGRATION:
   - High-Fidelity PDF Handling: Uses PyMuPDF (fitz) for superior PDF rendering 
     and text extraction compared to standard libraries.
   - Office Integration: Interfaces with LibreOffice/soffice headless mode 
     via subprocess to handle complex DOCX to PDF/Office conversions.
   - DOCX Extraction: Leverages Mammoth to extract clean, raw text from Word 
     documents for the text preview mode.

3. USER INTERFACE (UI/UX):
   - Adaptive Theming: Automatically detects system palette lightness to 
     switch between custom Light and Dark modes via CSS/QSS.
   - Interactive Preview: Supports drag-and-drop file loading and provides 
     synchronized image and text previews for multi-page documents.
   - Visual Feedback: Features an AnimatedProgressBar for indeterminate 
     background tasks and a status bar for real-time file save notifications.

CORE FEATURES
-------------
* Universal Converter: Convert between PDF, DOCX, PPTX, XLSX, RTF, ODT, EPUB, 
  and image formats (JPG/PNG).
* PDF Toolset: Merge multiple PDFs into a single file or split specific 
  page ranges from an existing document.
* Professional Viewer: High-resolution PDF rendering with page-by-page 
  navigation and text extraction views.
* Cross-Platform Logic: Includes path and command handling for Windows, 
  Linux, and macOS environments.

INSTALLATION & REQUIREMENTS
---------------------------
1. Prerequisites:
   - Python 3.8+
   - LibreOffice (required for DOCX/Office format conversions)
   - FFmpeg (recommended for certain image processing)

2. Python Dependencies:
   pip install PySide6 PyPDF2 pdf2docx mammoth pymupdf

3. Running the App:
   python ZePdf.py

USAGE
-----
1. Select File: Drag and drop a file or use the "Select File" button.
2. Preview: Use the "Image/PDF" or "Text" tabs to inspect the document.
3. Convert: Choose a target format from the dropdown and select an output 
   directory.
4. PDF Tools: If a PDF is loaded, use the "Merge" or "Split" buttons to 
   perform page-specific operations.

CREDITS
-------
Developed by Giscard
