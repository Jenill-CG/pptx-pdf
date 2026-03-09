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


# import streamlit as st
# import os
# import tempfile
# import zipfile
# from comtypes.client import CreateObject

# st.title("PPT to PDF Converter & Segregator")

# uploaded_files = st.file_uploader(
#     "Upload PPT files",
#     type=["ppt","pptx"],
#     accept_multiple_files=True
# )

# if st.button("Convert to PDF"):

#     temp_dir = tempfile.mkdtemp()
#     output_dir = os.path.join(temp_dir,"pdfs")
#     os.makedirs(output_dir,exist_ok=True)

#     powerpoint = CreateObject("Powerpoint.Application")
#     powerpoint.Visible = 1

#     for file in uploaded_files:

#         ppt_path = os.path.join(temp_dir,file.name)

#         with open(ppt_path,"wb") as f:
#             f.write(file.read())

#         pdf_path = os.path.join(output_dir,file.name.replace(".pptx",".pdf"))

#         presentation = powerpoint.Presentations.Open(ppt_path)
#         presentation.SaveAs(pdf_path,32)
#         presentation.Close()

#     powerpoint.Quit()

#     zip_path = os.path.join(temp_dir,"output.zip")

#     with zipfile.ZipFile(zip_path,"w") as z:
#         for root,dirs,files in os.walk(output_dir):
#             for file in files:
#                 z.write(os.path.join(root,file),file)

#     with open(zip_path,"rb") as f:
#         st.download_button("Download PDFs",f,"pdf_files.zip")

import os
import shutil
import pandas as pd
import streamlit as st

st.title("📂 File Segregation Tool")

uploaded_file = st.file_uploader(
    "Upload segregation_sequence (CSV or Excel)",
    type=["csv", "xlsx"]
)

parent_folder = st.text_input(
    "Enter Parent Folder Path (where files exist)"
)

if uploaded_file and parent_folder:

    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)

    st.write("Preview of Uploaded File")
    st.dataframe(df.head())

    if st.button("Start Segregation"):

        file_column = df.columns[-1]

        progress_bar = st.progress(0)
        status = st.empty()

        rows = df.to_dict("records")
        total = len(rows)

        moved_files = 0
        missing_files = []

        for i, row in enumerate(rows):

            file_name = str(row[file_column]).strip()

            if file_name == "" or file_name == "nan":
                continue

            source_path = os.path.join(parent_folder, file_name)

            if not os.path.exists(source_path):
                missing_files.append(file_name)
                continue

            folder_path = parent_folder

            for col in df.columns[:-1]:

                value = str(row[col]).strip()

                if value == "" or value == "nan":
                    continue

                folder_path = os.path.join(folder_path, value)

            os.makedirs(folder_path, exist_ok=True)

            destination_path = os.path.join(folder_path, file_name)

            shutil.move(source_path, destination_path)

            moved_files += 1

            progress_bar.progress((i + 1) / total)
            status.text(f"Processing {i+1} of {total}")

        st.success(f"Segregation Complete! Files moved: {moved_files}")

        if missing_files:

            st.warning(f"{len(missing_files)} files missing")

            missing_df = pd.DataFrame(missing_files, columns=["Missing Files"])

            st.download_button(
                "Download Missing File List",
                missing_df.to_csv(index=False),
                file_name="missing_files.csv"
            )
