import streamlit as st
import pandas as pd
import tempfile
import os
from pdf_extractor import PrimeConduitExtractor

st.set_page_config(page_title="PDF Extractor", page_icon="📊", layout="wide")
st.title("📊 PDF Extractor")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract"):
        for uploaded_file in uploaded_files:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(uploaded_file.getbuffer())
                    tmp_path = tmp.name
                
                extractor = PrimeConduitExtractor()
                df, _ = extractor.extract_from_pdf(tmp_path)
                os.unlink(tmp_path)
                
                st.success(f"✅ {uploaded_file.name}: {len(df)} rows")
                
                # Create Excel
                excel_path = f"extracted_{uploaded_file.name.split('.')[0]}.xlsx"
                extractor.rows = [tuple(row) for row in df.values]
                extractor.to_excel(excel_path)
                
                # Download button
                with open(excel_path, 'rb') as f:
                    st.download_button(
                        label=f"⬇️ Download {uploaded_file.name.split('.')[0]}.xlsx",
                        data=f.read(),
                        file_name=excel_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                os.unlink(excel_path)
                
            except Exception as e:
                st.error(f"❌ {uploaded_file.name}: {str(e)}")
