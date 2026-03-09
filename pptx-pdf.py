# import streamlit as st
# import os
# import tempfile
# import zipfile
# import subprocess

# st.title("Batch PPTX to PDF Converter")

# uploaded_files = st.file_uploader("Upload all PPTX files from your folder", type=["pptx"], accept_multiple_files=True)

# if uploaded_files:
#     with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:
#         st.info("Uploading files...")

#         # Save all uploaded files to a temp input folder
#         for file in uploaded_files:
#             pptx_path = os.path.join(input_dir, file.name)
#             with open(pptx_path, "wb") as f:
#                 f.write(file.read())

#         st.info("Converting PPTX to PDFs...")

#         # Convert each PPTX using LibreOffice
#         for file_name in os.listdir(input_dir):
#             if file_name.endswith(".pptx"):
#                 full_path = os.path.join(input_dir, file_name)
#                 subprocess.run([
#                     "libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, full_path
#                 ], check=True)

#         st.info("Creating ZIP archive...")

#         # Create a ZIP of all PDFs
#         zip_path = os.path.join(output_dir, "converted_pdfs.zip")
#         with zipfile.ZipFile(zip_path, "w") as zipf:
#             for file in os.listdir(output_dir):
#                 if file.endswith(".pdf"):
#                     zipf.write(os.path.join(output_dir, file), arcname=file)

#         with open(zip_path, "rb") as zip_file:
#             st.success("Conversion complete! Download your PDFs below.")
#             st.download_button("Download All PDFs (ZIP)", zip_file, "converted_pdfs.zip", "application/zip")


import streamlit as st
import os
import tempfile
import zipfile
from comtypes.client import CreateObject

st.title("PPT to PDF Converter & Segregator")

uploaded_files = st.file_uploader(
    "Upload PPT files",
    type=["ppt","pptx"],
    accept_multiple_files=True
)

if st.button("Convert to PDF"):

    temp_dir = tempfile.mkdtemp()
    output_dir = os.path.join(temp_dir,"pdfs")
    os.makedirs(output_dir,exist_ok=True)

    powerpoint = CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    for file in uploaded_files:

        ppt_path = os.path.join(temp_dir,file.name)

        with open(ppt_path,"wb") as f:
            f.write(file.read())

        pdf_path = os.path.join(output_dir,file.name.replace(".pptx",".pdf"))

        presentation = powerpoint.Presentations.Open(ppt_path)
        presentation.SaveAs(pdf_path,32)
        presentation.Close()

    powerpoint.Quit()

    zip_path = os.path.join(temp_dir,"output.zip")

    with zipfile.ZipFile(zip_path,"w") as z:
        for root,dirs,files in os.walk(output_dir):
            for file in files:
                z.write(os.path.join(root,file),file)

    with open(zip_path,"rb") as f:
        st.download_button("Download PDFs",f,"pdf_files.zip")
