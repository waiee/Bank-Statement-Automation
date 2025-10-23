# app_streamlit.py
import os
import tempfile
import streamlit as st
import pandas as pd
from pathlib import Path
import shutil

# import your functions from main.py
# make sure main.py is in same folder or available in PYTHONPATH
import main as bank_parser

st.set_page_config(page_title="Bank Statement Automation", layout="wide")

st.title("Bank Statement Automation")
st.markdown("Upload one or more `bank statements` and get a combined `mastersheet.xlsx`.")

# sidebar options
st.sidebar.header("Options")
year = st.sidebar.text_input("Default year (used when date has no year)", value=bank_parser.YEAR)
prefix_or = st.sidebar.text_input("Cash In ID", value=bank_parser.PREFIX_OR)
prefix_pv = st.sidebar.text_input("Cash Out ID", value=bank_parser.PREFIX_PV)

# override the constants in memory (won't rewrite file)
bank_parser.YEAR = year
bank_parser.PREFIX_OR = prefix_or
bank_parser.PREFIX_PV = prefix_pv

uploaded = st.file_uploader("Upload files", type=["xlsx"], accept_multiple_files=True)

if uploaded:
    st.info(f"{len(uploaded)} file(s) uploaded. Processing...")
    # create temp data dir
    tmpdir = tempfile.mkdtemp(prefix="bank_app_")
    try:
        # save uploaded files to temp dir
        for f in uploaded:
            dest = Path(tmpdir) / f.name
            with open(dest, "wb") as out:
                out.write(f.getbuffer())

        # point the script to temp data dir and output dir
        bank_parser.DATA_DIR = tmpdir
        out_dir = Path(tmpdir) / "output"
        bank_parser.OUTPUT_DIR = str(out_dir)
        bank_parser.OUTPUT_FILE = str(out_dir / "DB_statements.xlsx")

        # run the process and capture printed lines
        with st.spinner("Parsing files..."):
            bank_parser.process_excel_files()

        # show preview of generated file (if exists)
        out_file = Path(bank_parser.OUTPUT_FILE)
        if out_file.exists():
            st.success("Finished processing ✅")
            df_preview = pd.read_excel(out_file)
            st.subheader("Mastersheet Preview")
            st.dataframe(df_preview.head(20))

            # provide download
            with open(out_file, "rb") as f:
                st.download_button("Download DB_statements.xlsx", data=f, file_name="DB_statements.xlsx")
        else:
            st.error("Processing finished but no output file found. Check logs in console.")
    finally:
        # cleanup tempdir on demand (comment if you want to keep)
        # shutil.rmtree(tmpdir)
        pass
else:
    st.info("Upload files from the sidebar or drag-and-drop here.")

st.sidebar.markdown("---")
# st.sidebar.write("Config reference: extracted from your `main.py`.")
