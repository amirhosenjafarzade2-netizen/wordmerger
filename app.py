import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from docx.enum.text import WD_BREAK
from tempfile import NamedTemporaryFile
import os

st.set_page_config(
    page_title="DOCX Merger",
    page_icon="📄",
    layout="wide"
)

st.title("📄 DOCX Merger")
st.write(
    "Upload Word documents, choose their order, and merge them into a single DOCX."
)

uploaded_files = st.file_uploader(
    "Upload DOCX files",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.subheader("Document Order")

    filenames = [f.name for f in uploaded_files]

    order = []
    remaining = filenames.copy()

    for i in range(len(filenames)):
        selected = st.selectbox(
            f"Position {i + 1}",
            remaining,
            key=f"order_{i}"
        )
        order.append(selected)
        remaining.remove(selected)

    st.divider()

    if st.button("Merge Documents", type="primary"):

        temp_paths = []
        output_path = None

        try:

            ordered_files = []

            for filename in order:
                for uploaded in uploaded_files:
                    if uploaded.name == filename:
                        ordered_files.append(uploaded)
                        break

            # Save uploaded files temporarily
            for uploaded in ordered_files:

                tmp = NamedTemporaryFile(
                    delete=False,
                    suffix=".docx"
                )

                tmp.write(uploaded.getbuffer())
                tmp.close()

                temp_paths.append(tmp.name)

            if len(temp_paths) == 0:
                st.error("No files selected.")
                st.stop()

            # First document becomes master
            master_doc = Document(temp_paths[0])
            composer = Composer(master_doc)

            # Append remaining docs with page breaks
            for path in temp_paths[1:]:

                # Force next document onto a fresh page
                paragraph = master_doc.add_paragraph()
                run = paragraph.add_run()
                run.add_break(WD_BREAK.PAGE)

                composer.append(Document(path))

            # Save merged file
            output_file = NamedTemporaryFile(
                delete=False,
                suffix=".docx"
            )

            output_path = output_file.name
            output_file.close()

            composer.save(output_path)

            with open(output_path, "rb") as f:
                merged_bytes = f.read()

            st.success("Documents merged successfully!")

            st.download_button(
                label="⬇ Download Merged DOCX",
                data=merged_bytes,
                file_name="merged_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:

            st.error(f"Merge failed: {str(e)}")

        finally:

            for path in temp_paths:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass

            if output_path:
                try:
                    if os.path.exists(output_path):
                        os.remove(output_path)
                except Exception:
                    pass
