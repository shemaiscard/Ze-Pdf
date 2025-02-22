import streamlit as st
import os
import tempfile
import time
import subprocess
import platform
from pathlib import Path
from datetime import datetime
from pdf2docx import Converter as PDFToDocx
from PyPDF2 import PdfReader, PdfWriter
import mammoth
from PIL import Image
from pdf2image import convert_from_path

class DocumentConverter:
    def __init__(self):
        self.supported_input_formats = {
            'PDF': ['.pdf'],
            'Word': ['.docx', '.doc'],
            'PowerPoint': ['.pptx', '.ppt'],
            'Excel': ['.xlsx', '.xls'],
            'Rich Text': ['.rtf'],
            'OpenDocument': ['.odt', '.odp', '.ods'],
            'E-book': ['.epub', '.mobi'],
            'Image': ['.jpg', '.jpeg', '.png']
        }
        self.temp_dir = tempfile.mkdtemp()

    def validate_file(self, uploaded_file):
        file_extension = Path(uploaded_file.name).suffix.lower()
        valid_formats = [ext for formats in self.supported_input_formats.values() for ext in formats]

        if file_extension not in valid_formats:
            return False, "Unsupported file format"

        if uploaded_file.size > 200 * 1024 * 1024:  # 200MB limit
            return False, "File size too large (max 200MB)"

        return True, "File is valid"

    def convert_pdf_to_docx(self, input_path, output_path):
        try:
            converter = PDFToDocx(input_path)
            converter.convert(output_path)
            converter.close()
            return True, "Conversion successful"
        except Exception as e:
            return False, f"Conversion failed: {str(e)}"

    def convert_docx_to_pdf(self, input_path, output_path):
        try:
            if platform.system() == "Linux":
                subprocess.run(["unoconv", "-f", "pdf", "-o", output_path, input_path], check=True)
                return True, "Conversion successful"
            else:
                return False, "Unoconv-based conversion is for Linux only."
        except Exception as e:
            return False, f"Conversion failed: {str(e)}"

    def convert_pdf_to_image(self, input_path, output_folder, output_format="jpg"):
        try:
            images = convert_from_path(input_path, dpi=300)
            image_paths = []
            save_format = "JPEG" if output_format.lower() == "jpg" else output_format.upper()
            for i, image in enumerate(images):
                img_path = os.path.join(output_folder, f"page_{i+1}.{output_format.lower()}")
                image.save(img_path, save_format)
                image_paths.append(img_path)
            return True, "Conversion successful", image_paths
        except Exception as e:
            return False, f"Conversion failed: {str(e)}", None

    def convert_docx_to_image_chain(self, input_path, output_format="jpg"):
        temp_pdf = os.path.join(self.temp_dir, f"temp_{int(time.time())}.pdf")
        success, msg = self.convert_docx_to_pdf(input_path, temp_pdf)
        if not success:
            return False, msg, None
        return self.convert_pdf_to_image(temp_pdf, self.temp_dir, output_format)

    def convert_pdf_to_other(self, input_path, output_path, output_format):
        temp_docx = os.path.join(self.temp_dir, f"temp_{int(time.time())}.docx")
        success, msg = self.convert_pdf_to_docx(input_path, temp_docx)
        if not success:
            return False, "Intermediate PDF to DOCX conversion failed: " + msg
        try:
            subprocess.run(["unoconv", "-f", output_format, "-o", output_path, temp_docx], check=True)
            return True, "Conversion successful"
        except Exception as e:
            return False, f"Conversion failed: {str(e)}"

    def convert_other_formats(self, input_path, output_path, output_format):
        try:
            subprocess.run(["unoconv", "-f", output_format, "-o", output_path, input_path], check=True)
            return True, "Conversion successful"
        except Exception as e:
            return False, f"Conversion failed: {str(e)}"

def preview_file(uploaded_file):
    file_extension = Path(uploaded_file.name).suffix.lower()
    if file_extension in ['.jpg', '.jpeg', '.png']:
        st.image(uploaded_file, caption="Image Preview", use_column_width=True)
    elif file_extension == '.pdf':
        temp_pdf_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        try:
            pdf_reader = PdfReader(temp_pdf_path)
            first_page = pdf_reader.pages[0]
            text = first_page.extract_text()
        except Exception as e:
            text = f"Error reading PDF: {str(e)}"
        st.subheader("PDF Preview (First Page Text)")
        st.text(text if text else "No text detected (possibly an image-based PDF).")
    elif file_extension in ['.docx', '.doc']:
        temp_docx_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_docx_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        with open(temp_docx_path, "rb") as docx_file:
            text = mammoth.extract_raw_text(docx_file).value
        st.subheader("Word Document Preview (Text Content)")
        st.text(text[:500] + "..." if len(text) > 500 else text)

def main():
    st.set_page_config(page_title="ZePdf", page_icon="📄", layout="wide")

    # Inject custom CSS for a vibrant green theme that adapts to dark/light mode.
    st.markdown(
        """
        <style>
        :root {
          --primary-color: #28a745;
          --primary-color-light: #34d058;
          --background-color: #ffffff;
          --text-color: #333333;
        }
        @media (prefers-color-scheme: dark) {
          :root {
             --background-color: #121212;
             --text-color: #eaeaea;
          }
        }
        body {
          background-color: var(--background-color) !important;
          color: var(--text-color) !important;
        }
        h1, h2, h3, h4, h5, h6 {
          color: var(--primary-color) !important;
        }
        .stButton>button {
          background-color: var(--primary-color) !important;
          color: #fff !important;
          border: none !important;
          border-radius: 8px !important;
        }
        .stButton>button:hover {
          background-color: var(--primary-color-light) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    converter = DocumentConverter()
    st.title("📄 ZePdf")
    st.write("Convert your documents to various formats with a vibrant green touch.")

    uploaded_file = st.file_uploader("Choose a file", type=[
        "pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls", "rtf", 
        "odt", "odp", "ods", "epub", "mobi", "jpg", "jpeg", "png"
    ])

    if uploaded_file:
        # Display file details for extra clarity.
        st.markdown("### Uploaded File Details")
        st.write(f"**Name:** {uploaded_file.name}")
        st.write(f"**Size:** {round(uploaded_file.size / 1024, 2)} KB")
        st.write(f"**Type:** {Path(uploaded_file.name).suffix.lower()}")

        is_valid, message = converter.validate_file(uploaded_file)
        if is_valid:
            with st.expander("🔍 File Preview"):
                preview_file(uploaded_file)

            output_format = st.selectbox(
                "Convert to:",
                ["PDF", "DOCX", "PPTX", "XLSX", "RTF", "ODT", "EPUB", "JPG", "PNG"]
            )

            if st.button("Convert"):
                start_time = time.time()
                with st.spinner("Converting..."):
                    input_path = os.path.join(converter.temp_dir, uploaded_file.name)
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    output_filename = f"converted_{int(time.time())}.{output_format.lower()}"
                    output_path = os.path.join(converter.temp_dir, output_filename)
                    success = False
                    message = "Conversion not attempted"

                    file_ext = Path(uploaded_file.name).suffix.lower()

                    if file_ext == '.pdf':
                        if output_format.lower() == 'docx':
                            success, message = converter.convert_pdf_to_docx(input_path, output_path)
                        elif output_format.lower() in ['jpg', 'png']:
                            success, message, image_paths = converter.convert_pdf_to_image(input_path, converter.temp_dir, output_format.lower())
                            if success and image_paths:
                                output_path = image_paths[0]
                        else:
                            success, message = converter.convert_pdf_to_other(input_path, output_path, output_format.lower())
                    elif file_ext in ['.docx', '.doc']:
                        if output_format.lower() == 'pdf':
                            success, message = converter.convert_docx_to_pdf(input_path, output_path)
                        elif output_format.lower() in ['jpg', 'png']:
                            success, message, image_paths = converter.convert_docx_to_image_chain(input_path, output_format.lower())
                            if success and image_paths:
                                output_path = image_paths[0]
                        else:
                            success, message = converter.convert_other_formats(input_path, output_path, output_format.lower())
                    else:
                        success, message = converter.convert_other_formats(input_path, output_path, output_format.lower())

                    elapsed_time = time.time() - start_time

                    if success:
                        with open(output_path, "rb") as file:
                            st.download_button("Download converted file", file, output_filename, "application/octet-stream")
                        st.success("✅ Conversion complete!")
                        st.info(f"Conversion took {elapsed_time:.2f} seconds.")
                    else:
                        st.error(message)
        else:
            st.error(message)

if __name__ == "__main__":
    main()
