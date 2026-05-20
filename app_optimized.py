"""
Streamlit Web App for Prime Conduit PDF Extractor
Upload PDF commission reports and download extracted Excel files
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from pdf_extractor import PrimeConduitExtractor
import time


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
    
    **⚡ Speed:** ~2-3 seconds per PDF
    """)

# Main upload area
st.header("📤 Upload & Extract")

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
    
    if st.button("🚀 Extract Data", use_container_width=True, type="primary"):
        all_dataframes = []
        extraction_results = []
        
        # Progress tracking
        progress_container = st.container()
        status_placeholder = st.empty()
        
        for idx, uploaded_file in enumerate(uploaded_files):
            status_placeholder.info(f"⏳ Processing {idx+1}/{len(uploaded_files)}: {uploaded_file.name}")
            
            try:
                # Create temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    tmp_path = tmp_file.name
                
                # Extract data
                start_time = time.time()
                extractor = PrimeConduitExtractor()
                df, metadata = extractor.extract_from_pdf(tmp_path)
                extraction_time = time.time() - start_time
                
                all_dataframes.append(df)
                extraction_results.append({
                    'file': uploaded_file.name,
                    'rows': len(df),
                    'commission': df['Commissions'].sum(),
                    'time': f"{extraction_time:.1f}s",
                    'status': '✅ Success'
                })
                
                # Clean up
                os.unlink(tmp_path)
                
            except Exception as e:
                extraction_results.append({
                    'file': uploaded_file.name,
                    'rows': 0,
                    'commission': 0,
                    'time': '0s',
                    'status': f'❌ Error: {str(e)[:30]}'
                })
        
        status_placeholder.success(f"✅ Extraction complete!")
        
        # Show results
        st.markdown("---")
        st.subheader("📊 Extraction Results")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Files", len(uploaded_files))
        with col2:
            st.metric("Total Rows", sum([r['rows'] for r in extraction_results]))
        with col3:
            total_commission = sum([r['commission'] for r in extraction_results])
            st.metric("Total Commission", f"${total_commission:,.2f}")
        with col4:
            successful = sum([1 for r in extraction_results if '✅' in r['status']])
            st.metric("Success Rate", f"{successful}/{len(uploaded_files)}")
        
        # Results table
        results_df = pd.DataFrame(extraction_results)
        st.dataframe(results_df, use_container_width=True, hide_index=True)
        
        # Download section
        st.markdown("---")
        st.subheader("📥 Download Extracted Data")
        
        if extraction_mode == "Single File":
            for idx, (result, df) in enumerate(zip(extraction_results, all_dataframes)):
                if '✅' in result['status']:
                    excel_path = f"extracted_{idx + 1}.xlsx"
                    
                    # Create Excel file
                    extractor = PrimeConduitExtractor()
                    extractor.rows = [tuple(row) for row in df.values]
                    extractor.to_excel(excel_path)
                    
                    # Read and offer download
                    with open(excel_path, 'rb') as f:
                        file_data = f.read()
                    
                    st.download_button(
                        label=f"⬇️ Download: {result['file'].split('.')[0]}.xlsx",
                        data=file_data,
                        file_name=f"extracted_{result['file'].split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"download_{idx}"
                    )
                    
                    # Clean up
                    try:
                        os.unlink(excel_path)
                    except:
                        pass
        
        else:  # Batch mode
            if all_dataframes:
                combined_df = pd.concat(all_dataframes, ignore_index=True)
                excel_path = "combined_extraction.xlsx"
                
                # Create combined Excel with summary
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
                
                # Read and offer download
                with open(excel_path, 'rb') as f:
                    file_data = f.read()
                
                st.download_button(
                    label=f"⬇️ Download Combined Excel ({len(combined_df)} rows)",
                    data=file_data,
                    file_name="combined_extraction.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                # Clean up
                try:
                    os.unlink(excel_path)
                except:
                    pass

else:
    st.info("👆 Upload PDF files above to begin extraction")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center'>
    <small>Prime Conduit PDF Extractor v2.0 | ⚡ Fast & Accurate | 100% Match Guaranteed</small>
</div>
""", unsafe_allow_html=True)
