import streamlit as st
import tempfile
import os
from openpyxl import Workbook

st.set_page_config(page_title="PDF Extractor", page_icon="📊")
st.title("📊 PDF Extractor")

uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Extract"):
        for file in uploaded_files:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(file.getbuffer())
                    tmp_path = tmp.name
                
                import pdfplumber
                rows = []
                with pdfplumber.open(tmp_path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            rows.append(text)
                
                os.unlink(tmp_path)
                
                st.success(f"✅ {file.name}: Processed")
                
                wb = Workbook()
                ws = wb.active
                ws.append(['Invoice Data'])
                for i, row in enumerate(rows[:100], 1):
                    ws.append([row[:100]])
                
                excel_file = f"export_{file.name.split('.')[0]}.xlsx"
                wb.save(excel_file)
                
                with open(excel_file, 'rb') as f:
                    st.download_button(
                        label=f"⬇️ {file.name}",
                        data=f.read(),
                        file_name=excel_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                os.unlink(excel_file)
            except Exception as e:
                st.error(f"Error: {str(e)}")
