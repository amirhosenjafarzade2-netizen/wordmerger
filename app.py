import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from tempfile import NamedTemporaryFile
import io
import os

st.set_page_config(
    page_title="DOCX Merger",
    page_icon="📄",
    layout="wide"
)

st.title("📄 DOCX Merger")
st.write("Upload Word documents, arrange their order, and merge them into a single DOCX.")

uploaded_files = st.file_uploader(
    "Upload DOCX files",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.subheader("Merge Order")

    file_names = [f.name for f in uploaded_files]

    order = []

    available = file_names.copy()

    for i in range(len(file_names)):
        selected = st.selectbox(
            f"Position {i + 1}",
            available,
            key=f"pos_{i}"
        )
        order.append(selected)
        available.remove(selected)

    if st.button("Merge Documents", type="primary"):

        try:

            ordered_files = []

            for filename in order:
                for uploaded in uploaded_files:
                    if uploaded.name == filename:
                        ordered_files.append(uploaded)
                        break

            temp_paths = []

            try:

                for uploaded in ordered_files:
                    tmp = NamedTemporaryFile(
                        delete=False,
                        suffix=".docx"
                    )

                    tmp.write(uploaded.getbuffer())
                    tmp.close()

                    temp_paths.append(tmp.name)

                master_doc = Document(temp_paths[0])
                composer = Composer(master_doc)

                for path in temp_paths[1:]:
                    composer.append(Document(path))

                output_path = NamedTemporaryFile(
                    delete=False,
                    suffix=".docx"
                ).name

                composer.save(output_path)

                with open(output_path, "rb") as f:
                    merged_bytes = f.read()

                st.success("Merge completed successfully!")

                st.download_button(
                    label="⬇ Download Merged DOCX",
                    data=merged_bytes,
                    file_name="merged_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            finally:

                for path in temp_paths:
                    if os.path.exists(path):
                        os.remove(path)

                if 'output_path' in locals():
                    if os.path.exists(output_path):
                        os.remove(output_path)

        except Exception as e:
            st.error(f"Merge failed: {e}")
