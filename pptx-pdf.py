import streamlit as st
import os
import tempfile
import zipfile
import subprocess

st.title("PPTX to PDF Converter")

uploaded_files = st.file_uploader("Upload PPTX files", type=["pptx"], accept_multiple_files=True)

if uploaded_files:
    with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:
        st.info("Converting...")

        # Save PPTX files
        for file in uploaded_files:
            input_path = os.path.join(input_dir, file.name)
            with open(input_path, "wb") as f:
                f.write(file.read())

        # Convert with LibreOffice
        for filename in os.listdir(input_dir):
            if filename.endswith(".pptx"):
                input_path = os.path.join(input_dir, filename)
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_path
                ], check=True)

        # Create ZIP of PDFs
        zip_path = os.path.join(output_dir, "converted_pdfs.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for pdf_file in os.listdir(output_dir):
                if pdf_file.endswith(".pdf"):
                    pdf_path = os.path.join(output_dir, pdf_file)
                    zipf.write(pdf_path, arcname=pdf_file)

        # Serve ZIP
        with open(zip_path, "rb") as f:
            st.success("Conversion completed!")
            st.download_button("Download All PDFs (ZIP)", f, file_name="converted_pdfs.zip", mime="application/zip")
