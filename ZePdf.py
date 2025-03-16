#!/usr/bin/env python3
import sys
import os
import tempfile
import time
import subprocess
import platform
import math
import re
from pathlib import Path
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from PySide6.QtCore import Qt, QSize, QThread, Signal, QTimer, QPointF, QEvent
from PySide6.QtGui import (QPixmap, QIcon, QColor, QPalette, QPainter, 
                          QLinearGradient, QTextCursor, QTextFormat, QFont, 
                          QImage, QTransform)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QFileDialog, QTextEdit, QMessageBox,
    QProgressBar, QTabWidget, QStyle, QSizePolicy, QScrollArea, 
    QGraphicsScene, QGraphicsView, QSlider, QSpacerItem, QInputDialog,
    QStyleOptionProgressBar, QStatusBar
)
from pdf2docx import Converter as PDFToDocx
import mammoth
import fitz  # PyMuPDF for better PDF handling

STYLE_SHEET = """
/* Light Theme */
QMainWindow {
    background-color: #f0f0f0;
    color: #333333;
}

QWidget {
    font-family: 'Segoe UI', Arial;
    font-size: 12px;
}

QPushButton {
    background-color: #4CAF50;
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    min-width: 80px;
}

QPushButton:hover {
    background-color: #45a049;
}

QPushButton:disabled {
    background-color: #cccccc;
    color: #666666;
}

QTextEdit {
    font-family: 'Courier New', monospace;
    font-size: 12px;
    line-height: 1.5;
}

/* Dark Theme */
[data-theme="dark"] QMainWindow {
    background-color: #2d2d2d;
    color: #ffffff;
}

[data-theme="dark"] QTextEdit {
    background-color: #404040;
    color: #ffffff;
    font-family: 'Courier New', monospace;
    font-size: 12px;
    line-height: 1.5;
}

[data-theme="dark"] QComboBox,
[data-theme="dark"] QProgressBar {
    background-color: #404040;
    color: #ffffff;
}

ZoomableGraphicsView {
    background-color: palette(window);
}
"""

class ConversionThread(QThread):
    progress_updated = Signal(int)
    conversion_finished = Signal(bool, str, list)
    time_updated = Signal(str)

    def __init__(self, conversion_func, args):
        super().__init__()
        self.conversion_func = conversion_func
        self.args = args
        self._is_running = True

    def run(self):
        start_time = time.time()
        try:
            result = self.conversion_func(*self.args)
            elapsed = time.time() - start_time
            self.time_updated.emit(f"Time: {elapsed:.1f}s")
            self.conversion_finished.emit(*result)
        except Exception as e:
            self.conversion_finished.emit(False, f"Error: {str(e)}", None)

    def cancel(self):
        self._is_running = False

class DocumentConverter:
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
        self.active_thread = None

    def split_pdf(self, input_path, output_dir, split_range):
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
            
            pages = set()
            for part in split_range.split(','):
                part = part.strip()
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    start = max(1, start)
                    end = min(total_pages, end)
                    pages.update(range(start-1, end))
                else:
                    page = int(part)
                    if 1 <= page <= total_pages:
                        pages.add(page-1)

            if not pages:
                return False, "Invalid page numbers", None

            writer = PdfWriter()
            for page_num in sorted(pages):
                writer.add_page(reader.pages[page_num])

            output_path = os.path.join(output_dir, f"split_{int(time.time())}.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)

            return True, "PDF split successful", [output_path]
        except Exception as e:
            return False, f"PDF split failed: {str(e)}", None

    def merge_pdfs(self, file_paths, output_path):
        try:
            writer = PdfWriter()
            for path in file_paths:
                reader = PdfReader(path)
                for page in reader.pages:
                    writer.add_page(page)

            with open(output_path, "wb") as f:
                writer.write(f)

            return True, "PDF merge successful", [output_path]
        except Exception as e:
            return False, f"PDF merge failed: {str(e)}", None

    def _check_libreoffice_available(self):
        """Check if LibreOffice/unoconv is available"""
        try:
            # Check for soffice command (LibreOffice)
            if platform.system() == "Windows":
                result = subprocess.run(["where", "soffice"], 
                                      stdout=subprocess.PIPE, 
                                      stderr=subprocess.PIPE)
            else:  # Linux/Mac
                result = subprocess.run(["which", "soffice"], 
                                      stdout=subprocess.PIPE, 
                                      stderr=subprocess.PIPE)
            
            return result.returncode == 0
        except:
            return False

    def convert_docx_to_pdf(self, input_path, output_path):
        try:
            if self._check_libreoffice_available():
                # Try LibreOffice conversion
                if platform.system() == "Windows":
                    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", 
                                  "--outdir", os.path.dirname(output_path), input_path], 
                                  check=True)
                    # Rename the output file if needed
                    base_name = os.path.basename(input_path).rsplit('.', 1)[0] + ".pdf"
                    tmp_output = os.path.join(os.path.dirname(output_path), base_name)
                    if tmp_output != output_path and os.path.exists(tmp_output):
                        os.rename(tmp_output, output_path)
                else:
                    # On Linux/Mac, we can try unoconv first, then fallback to soffice
                    try:
                        subprocess.run(["unoconv", "-f", "pdf", "-o", output_path, input_path], check=True)
                    except:
                        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", 
                                      "--outdir", os.path.dirname(output_path), input_path], 
                                      check=True)
                        # Rename if needed
                        base_name = os.path.basename(input_path).rsplit('.', 1)[0] + ".pdf"
                        tmp_output = os.path.join(os.path.dirname(output_path), base_name)
                        if tmp_output != output_path and os.path.exists(tmp_output):
                            os.rename(tmp_output, output_path)
                
                if os.path.exists(output_path):
                    return True, "DOCX to PDF conversion successful", [output_path]
                else:
                    return False, "Conversion failed: Output file not found", None
            else:
                return False, "LibreOffice not installed. Please install LibreOffice for document conversion.", None
        except Exception as e:
            return False, f"Conversion failed: {str(e)}", None

    def convert_pdf_to_docx(self, input_path, output_path):
        try:
            converter = PDFToDocx(input_path)
            converter.convert(output_path)
            converter.close()
            return True, "PDF to DOCX conversion successful", [output_path]
        except Exception as e:
            return False, f"Conversion failed: {str(e)}", None

    def convert_docx_to_other(self, input_path, output_format):
        try:
            output_path = input_path.rsplit('.', 1)[0] + f".{output_format}"
            
            if self._check_libreoffice_available():
                # Try LibreOffice conversion
                if platform.system() == "Windows":
                    subprocess.run(["soffice", "--headless", "--convert-to", output_format, 
                                  "--outdir", os.path.dirname(output_path), input_path], 
                                  check=True)
                    # Rename the output file if needed
                    base_name = os.path.basename(input_path).rsplit('.', 1)[0] + f".{output_format}"
                    tmp_output = os.path.join(os.path.dirname(output_path), base_name)
                    if tmp_output != output_path and os.path.exists(tmp_output):
                        os.rename(tmp_output, output_path)
                else:
                    # On Linux/Mac, we can try unoconv first, then fallback to soffice
                    try:
                        subprocess.run(["unoconv", "-f", output_format, "-o", output_path, input_path], check=True)
                    except:
                        subprocess.run(["soffice", "--headless", "--convert-to", output_format, 
                                      "--outdir", os.path.dirname(output_path), input_path], 
                                      check=True)
                        # Rename if needed
                        base_name = os.path.basename(input_path).rsplit('.', 1)[0] + f".{output_format}"
                        tmp_output = os.path.join(os.path.dirname(output_path), base_name)
                        if tmp_output != output_path and os.path.exists(tmp_output):
                            os.rename(tmp_output, output_path)
                
                if os.path.exists(output_path):
                    return True, f"DOCX to {output_format.upper()} conversion successful", [output_path]
                else:
                    return False, "Conversion failed: Output file not found", None
            else:
                return False, "LibreOffice not installed. Please install LibreOffice for document conversion.", None
        except Exception as e:
            return False, f"Conversion failed: {str(e)}", None

    def convert_pdf_to_images(self, input_path, output_dir, fmt="jpg", dpi=150):
        try:
            # Make sure output_dir exists
            os.makedirs(output_dir, exist_ok=True)
            
            # Use PyMuPDF (fitz) for better PDF image extraction
            pdf_document = fitz.open(input_path)
            image_paths = []
            
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better quality
                
                img_path = os.path.join(output_dir, f"page_{page_num+1}.{fmt}")
                pix.save(img_path)
                image_paths.append(img_path)
                
            pdf_document.close()
            return True, "PDF to images conversion successful", image_paths
        except Exception as e:
            # Fallback to pdf2image if available
            try:
                from pdf2image import convert_from_path
                images = convert_from_path(input_path, dpi=dpi)
                image_paths = []
                for i, image in enumerate(images):
                    img_path = os.path.join(output_dir, f"page_{i+1}.{fmt}")
                    image.save(img_path, "JPEG" if fmt == "jpg" else fmt.upper())
                    image_paths.append(img_path)
                return True, "PDF to images conversion successful", image_paths
            except Exception as e2:
                return False, f"Conversion failed: {str(e2)}", None

    def convert_generic(self, input_path, output_format):
        try:
            output_dir = os.path.dirname(input_path)
            output_path = os.path.join(output_dir, f"{Path(input_path).stem}.{output_format}")
            
            if self._check_libreoffice_available():
                # Try LibreOffice conversion
                if platform.system() == "Windows":
                    subprocess.run(["soffice", "--headless", "--convert-to", output_format, 
                                "--outdir", output_dir, input_path], 
                                check=True)
                    # Rename the output file if needed
                    base_name = os.path.basename(input_path).rsplit('.', 1)[0] + f".{output_format}"
                    tmp_output = os.path.join(output_dir, base_name)
                    if tmp_output != output_path and os.path.exists(tmp_output):
                        os.rename(tmp_output, output_path)
                else:
                    # On Linux/Mac, we can try unoconv first, then fallback to soffice
                    try:
                        subprocess.run(["unoconv", "-f", output_format, "-o", output_path, input_path], check=True)
                    except:
                        subprocess.run(["soffice", "--headless", "--convert-to", output_format, 
                                    "--outdir", output_dir, input_path], 
                                    check=True)
                        # Rename if needed
                        base_name = os.path.basename(input_path).rsplit('.', 1)[0] + f".{output_format}"
                        tmp_output = os.path.join(output_dir, base_name)
                        if tmp_output != output_path and os.path.exists(tmp_output):
                            os.rename(tmp_output, output_path)
                
                # Check if file exists and return it
                if os.path.isfile(output_path):
                    return True, f"Conversion to {output_format.upper()} successful", [output_path]
                else:
                    # Try to find the converted file in the output directory
                    expected_filename = f"{Path(input_path).stem}.{output_format}"
                    for file in os.listdir(output_dir):
                        if file.lower() == expected_filename.lower():
                            return True, f"Conversion to {output_format.upper()} successful", [os.path.join(output_dir, file)]
                            
                    return False, "Conversion failed: Output file not found", None
            else:
                return False, "LibreOffice not installed. Please install LibreOffice for document conversion.", None
        except Exception as e:
            return False, f"Conversion failed: {str(e)}", None

class PreviewManager:
    def __init__(self):
        self.current_page = 0
        self.total_pages = 0
        self.pdf_document = None
        self.preview_content = []
        self.preview_type = None
        self.zoom_level = 1.0
        self.current_file_path = None

    def load_file(self, file_path):
        self.current_page = 0
        self.preview_content = []
        self.current_file_path = file_path
        file_extension = Path(file_path).suffix.lower()

        if file_extension in ['.jpg', '.jpeg', '.png']:
            self.preview_type = 'image'
            self.preview_content = [file_path]
            self.total_pages = 1
            if self.pdf_document:
                self.pdf_document.close()
                self.pdf_document = None

        elif file_extension == '.pdf':
            try:
                # Close previous document if any
                if self.pdf_document:
                    self.pdf_document.close()
                
                # Use PyMuPDF for better PDF handling
                self.pdf_document = fitz.open(file_path)
                self.total_pages = len(self.pdf_document)
                self.preview_type = 'pdf'
                self.preview_content = []  # Will be loaded on demand
                
            except Exception as e:
                if self.pdf_document:
                    self.pdf_document.close()
                    self.pdf_document = None
                self.preview_content = [f"Error reading PDF: {str(e)}"]
                self.total_pages = 1
                self.preview_type = 'text'

        elif file_extension in ['.docx', '.doc']:
            try:
                with open(file_path, "rb") as docx_file:
                    result = mammoth.extract_raw_text(docx_file)
                    self.preview_type = 'text'
                    self.preview_content = [result.value]
                    self.total_pages = 1
                    
                    if self.pdf_document:
                        self.pdf_document.close()
                        self.pdf_document = None

            except Exception as e:
                self.preview_content = [f"Error reading DOCX: {str(e)}"]
                self.total_pages = 1
                self.preview_type = 'text'
                if self.pdf_document:
                    self.pdf_document.close()
                    self.pdf_document = None

        else:
            self.preview_type = 'text'
            self.preview_content = ["Preview not available for this file format"]
            self.total_pages = 1
            if self.pdf_document:
                self.pdf_document.close()
                self.pdf_document = None

        return self.total_pages
    
    def get_current_page_content(self):
        if self.preview_type == 'pdf' and self.pdf_document:
            try:
                if 0 <= self.current_page < self.total_pages:
                    page = self.pdf_document.load_page(self.current_page)
                    return page.get_text()
                return "Page out of range"
            except:
                return "Error extracting text from PDF"
        elif 0 <= self.current_page < len(self.preview_content):
            return self.preview_content[self.current_page]
        return ""
    
    def get_current_page_image(self):
        if self.preview_type == 'pdf' and self.pdf_document:
            try:
                if 0 <= self.current_page < self.total_pages:
                    page = self.pdf_document.load_page(self.current_page)
                    # Get pixmap with good resolution
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    img_data = pix.samples
                    
                    # Convert to QImage
                    qimg = QImage(img_data, pix.width, pix.height, 
                                  pix.stride, QImage.Format_RGB888)
                    return QPixmap.fromImage(qimg)
                return None
            except Exception as e:
                print(f"Error rendering PDF page: {e}")
                return None
        elif self.preview_type == 'image' and len(self.preview_content) > 0:
            return QPixmap(self.preview_content[0])
        return None

    def cleanup(self):
        if self.pdf_document:
            self.pdf_document.close()
            self.pdf_document = None

class ZoomableGraphicsView(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.graphics_scene = QGraphicsScene()
        self.setScene(self.graphics_scene)
        self.pixmap_item = None

    def wheelEvent(self, event):
        zoom_factor = 1.25
        if event.angleDelta().y() > 0:
            self.scale(zoom_factor, zoom_factor)
        else:
            self.scale(1/zoom_factor, 1/zoom_factor)

    def fitInViewWithoutDistortion(self):
        if self.pixmap_item and not self.pixmap_item.pixmap().isNull():
            self.resetTransform()
            self.fitInView(self.pixmap_item, Qt.KeepAspectRatio)

class ZePdfWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ZePdf - Document Converter")
        self.resize(1200, 800)
        self.setAcceptDrops(True)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.palette = QApplication.palette()
        self.input_file_path = None
        self.set_theme()
        self.converter = DocumentConverter()
        self.preview_manager = PreviewManager()
        self.init_ui()
        self.setup_connections()
        # Status bar for messages
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

    def set_theme(self):
        bg_color = self.palette.color(QPalette.Window)
        if bg_color.lightness() < 128:
            self.setProperty("data-theme", "dark")
        else:
            self.setProperty("data-theme", "light")
        self.setStyleSheet(STYLE_SHEET)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # File selection
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.select_button = QPushButton(" Select File ")
        self.select_button.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.select_button)
        main_layout.addLayout(file_layout)

        # Preview area
        self.init_preview_area()
        main_layout.addWidget(self.preview_tabs, 1)  # Give preview more space

        # PDF Tools
        self.pdf_tools = QHBoxLayout()
        self.merge_button = QPushButton("Merge PDFs")
        self.split_button = QPushButton("Split PDF")
        self.pdf_tools.addWidget(self.merge_button)
        self.pdf_tools.addWidget(self.split_button)
        
        # Hide PDF tools initially
        for i in range(self.pdf_tools.count()):
            self.pdf_tools.itemAt(i).widget().setVisible(False)
        
        main_layout.addLayout(self.pdf_tools)

        # Conversion controls
        control_layout = QHBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems(["PDF", "DOCX", "PPTX", "XLSX", "RTF", "ODT", "EPUB", "JPG", "PNG"])
        
        self.convert_button = QPushButton(" Convert ")
        self.convert_button.setIcon(QIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton)))
        
        self.progress_bar = AnimatedProgressBar()
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        
        control_layout.addWidget(QLabel("Convert to:"))
        control_layout.addWidget(self.format_combo)
        control_layout.addWidget(self.convert_button)
        control_layout.addWidget(self.progress_bar)
        control_layout.addWidget(self.cancel_button)
        main_layout.addLayout(control_layout)

        # Time estimation
        self.time_label = QLabel()
        main_layout.addWidget(self.time_label)

    def init_preview_area(self):
        self.preview_tabs = QTabWidget()
        
        # Image preview with zoom
        self.image_view = ZoomableGraphicsView()
        
        # Text preview
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        font = QFont("Courier New", 12)
        self.text_preview.setFont(font)
        
        # Navigation controls
        nav_widget = QWidget()
        nav_layout = QHBoxLayout(nav_widget)
        self.prev_button = QPushButton()
        self.prev_button.setIcon(self.style().standardIcon(QStyle.SP_ArrowLeft))
        self.next_button = QPushButton()
        self.next_button.setIcon(self.style().standardIcon(QStyle.SP_ArrowRight))
        self.page_label = QLabel("Page 1 of 1")
        nav_layout.addWidget(self.prev_button)
        nav_layout.addWidget(self.page_label)
        nav_layout.addWidget(self.next_button)
        
        # Combine preview and controls
        preview_container = QVBoxLayout()
        preview_container.addWidget(self.text_preview, 1)
        preview_container.addWidget(nav_widget)
        text_widget = QWidget()
        text_widget.setLayout(preview_container)
        
        self.preview_tabs.addTab(self.image_view, "Image/PDF Preview")
        self.preview_tabs.addTab(text_widget, "Text Preview")
        self.preview_tabs.currentChanged.connect(self.on_tab_changed)

    def on_tab_changed(self, index):
        # When switching to image tab, make sure to display the image
        if index == 0 and self.preview_manager.preview_type == 'pdf':
            self.update_image_preview()

    def setup_connections(self):
        self.select_button.clicked.connect(self.select_file)
        self.merge_button.clicked.connect(self.merge_pdfs)
        self.split_button.clicked.connect(self.split_pdf)
        self.convert_button.clicked.connect(self.convert_file)
        self.cancel_button.clicked.connect(self.cancel_conversion)
        self.prev_button.clicked.connect(self.prev_page)
        self.next_button.clicked.connect(self.next_page)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            file_path = url.toLocalFile()
            self.handle_file(file_path)
            event.acceptProposedAction()

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", 
            "All Supported Files (*.pdf *.docx *.doc *.pptx *.ppt *.xlsx *.xls *.rtf *.odt *.odp *.ods *.epub *.mobi *.jpg *.jpeg *.png)"
        )
        if file_path:
            self.handle_file(file_path)

    def handle_file(self, file_path):
        self.input_file_path = file_path
        self.file_label.setText(f"Selected: {os.path.basename(file_path)}")
        
        is_pdf = Path(file_path).suffix.lower() == '.pdf'
        for i in range(self.pdf_tools.count()):
            self.pdf_tools.itemAt(i).widget().setVisible(is_pdf)
        
        self.update_preview()

    def update_preview(self):
        if self.input_file_path is None:
            return
            
        total_pages = self.preview_manager.load_file(self.input_file_path)
        
        if self.preview_manager.preview_type in ['image', 'pdf']:
            self.update_image_preview()
            if self.preview_manager.preview_type == 'pdf':
                # Also update text preview
                self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
        else:
            self.preview_tabs.setCurrentIndex(1)  # Switch to text tab
            self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
        
        self.update_page_label()
        
        # Update navigation buttons state
        self.prev_button.setEnabled(self.preview_manager.current_page > 0)
        self.next_button.setEnabled(
            self.preview_manager.current_page < self.preview_manager.total_pages - 1
        )

    def update_image_preview(self):
        pixmap = self.preview_manager.get_current_page_image()
        self.image_view.graphics_scene.clear()
        
        if pixmap and not pixmap.isNull():
            self.image_view.pixmap_item = self.image_view.graphics_scene.addPixmap(pixmap)
            self.image_view.fitInViewWithoutDistortion()
            self.preview_tabs.setCurrentIndex(0)  # Switch to image tab

    def update_page_label(self):
        self.page_label.setText(
            f"Page {self.preview_manager.current_page + 1} of {self.preview_manager.total_pages}"
        )

    def next_page(self):
        if self.preview_manager.current_page < self.preview_manager.total_pages - 1:
            self.preview_manager.current_page += 1
            
            if self.preview_manager.preview_type == 'pdf':
                self.update_image_preview()
                self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
            else:
                self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
                
            self.update_page_label()
            
            # Update navigation buttons state
            self.prev_button.setEnabled(True)
            self.next_button.setEnabled(
                self.preview_manager.current_page < self.preview_manager.total_pages - 1
            )
    def prev_page(self):
        if self.preview_manager.current_page > 0:
            self.preview_manager.current_page -= 1
            
            if self.preview_manager.preview_type == 'pdf':
                self.update_image_preview()
                self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
            else:
                self.text_preview.setPlainText(self.preview_manager.get_current_page_content())
                
            self.update_page_label()
            
            # Update navigation buttons state
            self.prev_button.setEnabled(
                self.preview_manager.current_page > 0
            )
            self.next_button.setEnabled(
                self.preview_manager.current_page < self.preview_manager.total_pages - 1
            )

    def merge_pdfs(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF files to merge", "", "PDF Files (*.pdf)"
        )
        if file_paths:
            output_path, _ = QFileDialog.getSaveFileName(
                self, "Save Merged PDF", "", "PDF Files (*.pdf)"
            )
            if output_path:
                self.start_conversion(self.converter.merge_pdfs, (file_paths, output_path)) 
    def convert_file(self):
        if not self.input_file_path:
            QMessageBox.warning(self, "No file selected", "Please select a file to convert.")
            return
        
        output_format = self.format_combo.currentText().lower()
        
        # Different handling for image output
        if output_format in ["jpg", "png"]:
            output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
            if not output_dir:
                return
            
            self.progress_bar.setVisible(True)
            self.progress_bar.startAnimation()  # Use updated method
            self.cancel_button.setEnabled(True)
            self.convert_button.setEnabled(False)
            
            self.active_thread = ConversionThread(
                self.converter.convert_pdf_to_images, 
                (self.input_file_path, output_dir, output_format)
            )
        else:
            # For document formats
            output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
            if not output_dir:
                return
            
            output_path = os.path.join(output_dir, f"{Path(self.input_file_path).stem}.{output_format}")
            
            self.progress_bar.setVisible(True)
            self.progress_bar.startAnimation()  # Use updated method
            self.cancel_button.setEnabled(True)
            self.convert_button.setEnabled(False)
            
            conversion_func = None
            input_ext = Path(self.input_file_path).suffix.lower()
            
            if output_format == "pdf":
                if input_ext in [".docx", ".doc"]:
                    conversion_func = self.converter.convert_docx_to_pdf
                else:
                    conversion_func = self.converter.convert_generic
            elif output_format == "docx":
                if input_ext == ".pdf":
                    conversion_func = self.converter.convert_pdf_to_docx
                else:
                    conversion_func = self.converter.convert_generic
            else:
                # For other formats like xlsx, pptx, etc.
                conversion_func = self.converter.convert_generic
            
            args = (self.input_file_path, output_path)
            if output_format in ["xlsx", "pptx", "rtf", "odt", "epub"]:
                # Add output format as parameter for generic conversion
                args = (self.input_file_path, output_format)
                
            self.active_thread = ConversionThread(conversion_func, args)
        
        self.active_thread.progress_updated.connect(self.progress_bar.setValue)
        self.active_thread.conversion_finished.connect(self.on_conversion_finished)
        self.active_thread.time_updated.connect(self.time_label.setText)
        self.active_thread.start()

    def on_conversion_finished(self, success, message, output_files):
        self.progress_bar.stopAnimation()  # Stop animation
        self.progress_bar.setVisible(False)
        self.cancel_button.setEnabled(False)
        self.convert_button.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "Conversion Successful", message)
            # Handle directory of images
            if isinstance(output_files, list) and len(output_files) > 0:
                if os.path.isdir(os.path.dirname(output_files[0])):
                    self.statusBar.showMessage(f"Files saved to: {os.path.dirname(output_files[0])}")
                if Path(output_files[0]).suffix.lower() not in ['.jpg', '.jpeg', '.png']:
                    self.handle_file(output_files[0])
        else:
            QMessageBox.critical(self, "Conversion Failed", message)

    def cancel_conversion(self):
        if self.active_thread:
            self.active_thread.cancel()
            self.active_thread = None
            self.progress_bar.stopAnimation()  # Stop animation
            self.progress_bar.setVisible(False)
            self.cancel_button.setEnabled(False)
            self.convert_button.setEnabled(True)
            self.time_label.clear()
            QMessageBox.information(self, "Conversion Cancelled", "The conversion has been cancelled.")
    
    def split_pdf(self):
        if not self.input_file_path or Path(self.input_file_path).suffix.lower() != '.pdf':
            QMessageBox.warning(self, "Invalid file", "Please select a PDF file to split.")
            return
            
        # Get total pages from the current PDF document
        total_pages = self.preview_manager.total_pages
        
        split_range, ok = QInputDialog.getText(self, "Split PDF", 
            f"Available pages: 1-{total_pages}\nEnter page numbers to split (e.g. 1-3, 5, 7-10):")
        if ok:
            output_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
            if output_dir:
                self.start_conversion(self.converter.split_pdf, 
                                    (self.input_file_path, output_dir, split_range))
    
    def start_conversion(self, conversion_func, args):
        self.progress_bar.setVisible(True)
        self.progress_bar.startAnimation()  # Use updated method
        self.cancel_button.setEnabled(True)
        self.convert_button.setEnabled(False)
        
        self.active_thread = ConversionThread(conversion_func, args)
        self.active_thread.progress_updated.connect(self.progress_bar.setValue)
        self.active_thread.conversion_finished.connect(self.on_conversion_finished)
        self.active_thread.time_updated.connect(self.time_label.setText)
        self.active_thread.start()

    def closeEvent(self, event):
        self.preview_manager.cleanup()
        event.accept()  # Allow the window to close     

class AnimatedProgressBar(QProgressBar):        
    def __init__(self):
        super().__init__()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress)
        self.setVisible(False)  # Hide initially
        self.animated_value = 0
        self.direction = 1
        self.is_animating = False

    def update_progress(self):
        if not self.isVisible() or not self.is_animating:
            return
            
        self.animated_value += self.direction
        if self.animated_value >= 100:
            self.direction = -1
        elif self.animated_value <= 0:
            self.direction = 1
        super().setValue(self.animated_value)
    
    def setValue(self, value):
        if value < 100:
            super().setValue(value)
        else:
            super().setValue(100)
            self.is_animating = False
            
    def startAnimation(self):
        self.is_animating = True
        self.animated_value = 0
        self.direction = 1
        if not self.timer.isActive():
            self.timer.start(50)
            
    def stopAnimation(self):
        self.is_animating = False
        if self.timer.isActive():
            self.timer.stop()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ZePdfWindow()
    window.show()           
    sys.exit(app.exec())
