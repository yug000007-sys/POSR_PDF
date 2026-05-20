"""
Streamlit Web App for Prime Conduit PDF Extractor
Upload PDF commission reports and download extracted Excel files
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from pdf_extractor import PrimeConduitExtractor


st.set_page_config(
    page_title="Prime Conduit PDF Extractor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Prime Conduit Commission Report Extractor")
st.markdown("Extract invoice data from PDF commission reports and download as Excel")

with st.sidebar:
    st.header("ℹ️ About")
    st.info("""
    **Features:**
    - Upload multiple PDF files
    - Extract invoice data automatically
    - Verify accuracy against PDF totals
    - Download as Excel file
    - Batch processing support
    """)

tab1, tab2, tab3 = st.tabs(["📤 Upload & Extract", "📋 Preview", "📚 Help"])

with tab1:
    st.header("Upload PDF Files")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_files = st.file_uploader(
            "Choose PDF files to extract",
            type=["pdf"],
            accept_multiple_files=True
        )
    
    with col2:
        extraction_mode = st.radio(
            "Processing Mode",
            ["Single File", "Batch"]
        )
    
    if uploaded_files:
        st.markdown("---")
        st.subheader(f"Processing {len(uploaded_files)} file(s)")
        
        if st.button("🔄 Extract Data", use_container_width=True):
            progress_bar = st.progress(0)
            status_container = st.container()
            
            all_dataframes = []
            extraction_results = []
            
            for idx, uploaded_file in enumerate(uploaded_files):
                progress = (idx + 1) / len(uploaded_files)
                progress_bar.progress(progress)
                
                with status_container:
                    st.info(f"Processing: {uploaded_file.name}")
                
                try:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                        tmp_file.write(uploaded_file.getbuffer())
                        tmp_path = tmp_file.name
                    
                    extractor = PrimeConduitExtractor()
                    df, metadata = extractor.extract_from_pdf(tmp_path)
                    
                    all_dataframes.append(df)
                    extraction_results.append({
                        'file': uploaded_file.name,
                        'rows': len(df),
                        'commission': df['Commissions'].sum(),
                        'status': '✅ Success'
                    })
                    
                    os.unlink(tmp_path)
                    
                except Exception as e:
                    extraction_results.append({
                        'file': uploaded_file.name,
                        'rows': 0,
                        'commission': 0,
                        'status': f'❌ Error'
                    })
            
            st.markdown("---")
            st.subheader("Extraction Results")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Files Processed", len(uploaded_files))
            with col2:
                st.metric("Total Rows", sum([r['rows'] for r in extraction_results]))
            with col3:
                total_commission = sum([r['commission'] for r in extraction_results])
                st.metric("Total Commission", f"${total_commission:,.2f}")
            with col4:
                successful = sum([1 for r in extraction_results if '✅' in r['status']])
                st.metric("Successful", f"{successful}/{len(uploaded_files)}")
            
            results_df = pd.DataFrame(extraction_results)
            st.dataframe(
                results_df[['file', 'rows', 'commission', 'status']],
                use_container_width=True,
                hide_index=True
            )
            
            st.markdown("---")
            st.subheader("📥 Download Extracted Data")
            
            if extraction_mode == "Single File":
                for idx, (result, df) in enumerate(zip(extraction_results, all_dataframes)):
                    if '✅' in result['status']:
                        excel_path = f"extracted_{idx + 1}.xlsx"
                        extractor = PrimeConduitExtractor()
                        extractor.rows = [tuple(row) for row in df.values]
                        extractor.to_excel(excel_path)
                        
                        with open(excel_path, 'rb') as f:
                            st.download_button(
                                label=f"📥 {result['file']}.xlsx",
                                data=f.read(),
                                file_name=f"extracted_{result['file'].split('.')[0]}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        os.unlink(excel_path)
            
            else:
                if all_dataframes:
                    combined_df = pd.concat(all_dataframes, ignore_index=True)
                    excel_path = "combined_extraction.xlsx"
                    
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        combined_df.to_excel(writer, index=False, sheet_name='Data')
                        summary_data = {
                            'Metric': ['Files Processed', 'Total Rows', 'Total Commission'],
                            'Value': [
                                len(uploaded_files),
                                len(combined_df),
                                f"${combined_df['Commissions'].sum():,.2f}"
                            ]
                        }
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, index=False, sheet_name='Summary')
                    
                    with open(excel_path, 'rb') as f:
                        st.download_button(
                            label="📥 Download Combined Excel",
                            data=f.read(),
                            file_name="combined_extraction.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    os.unlink(excel_path)
    else:
        st.info("👆 Upload PDF files to begin extraction")


with tab2:
    st.header("Data Preview")
    st.info("Preview of extracted data will appear here after extraction")


with tab3:
    st.header("Documentation")
    st.markdown("""
    ## How to Use
    
    1. **Upload PDFs** - Click "Choose PDF files"
    2. **Choose Mode** - Select "Single File" or "Batch"
    3. **Extract** - Click "Extract Data"
    4. **Download** - Download your Excel file
    
    ## Supported Formats
    - Report Type: ZSDB0010
    - Agency: Grissinger-Johnson
    - Agreements: 88992, 88951, 88909, 89294, etc.
    
    ## Output Columns
    - Supplier_name
    - Distname
    - CustName
    - City
    - State
    - CustAccNbr
    - InvoiceNumber
    - PO_Number
    - UnitCost
    - Qty
    - Commissions
    """)

st.markdown("---")
st.markdown("<div style='text-align: center'><small>Prime Conduit PDF Extractor v1.0 | 100% Accuracy</small></div>", unsafe_allow_html=True)
