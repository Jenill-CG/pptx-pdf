import streamlit as st
import os
import tempfile
import zipfile
import subprocess

st.title("Batch PPTX to PDF Converter")

uploaded_files = st.file_uploader("Upload all PPTX files from your folder", type=["pptx"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:
        st.info("Uploading files...")

        # Save all uploaded files to a temp input folder
        for file in uploaded_files:
            pptx_path = os.path.join(input_dir, file.name)
            with open(pptx_path, "wb") as f:
                f.write(file.read())

        st.info("Converting PPTX to PDFs...")

        # Convert each PPTX using LibreOffice
        for file_name in os.listdir(input_dir):
            if file_name.endswith(".pptx"):
                full_path = os.path.join(input_dir, file_name)
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, full_path
                ], check=True)

        st.info("Creating ZIP archive...")

        # Create a ZIP of all PDFs
        zip_path = os.path.join(output_dir, "converted_pdfs.zip")
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for file in os.listdir(output_dir):
                if file.endswith(".pdf"):
                    zipf.write(os.path.join(output_dir, file), arcname=file)

        with open(zip_path, "rb") as zip_file:
            st.success("Conversion complete! Download your PDFs below.")
            st.download_button("Download All PDFs (ZIP)", zip_file, "converted_pdfs.zip", "application/zip")
