"""
Streamlit front end for the SampleSearch reformatter.

Run with: streamlit run app.py
"""

import tempfile
from pathlib import Path

import streamlit as st

from sample_search_reformat import reformat_sample_search

st.title("Sample Search Reformatter")

uploaded_file = st.file_uploader("Upload Sample Search export (.xlsx)", type="xlsx")

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = Path(tmp_dir) / uploaded_file.name
        input_path.write_bytes(uploaded_file.getvalue())

        try:
            output_path = reformat_sample_search(str(input_path))
        except ValueError as e:
            st.error(str(e))
        else:
            output_bytes = Path(output_path).read_bytes()
            st.success("File reformatted successfully.")
            st.download_button(
                label="Download formatted file",
                data=output_bytes,
                file_name=Path(output_path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
