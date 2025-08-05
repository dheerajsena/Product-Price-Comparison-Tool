import streamlit as st
import pandas as pd
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Product Price Comparison",
    layout="wide",
    page_icon="🚀"
)

# Header
st.markdown(
    """
    <div style='background-color: #002e5b; padding: 10px; border-radius: 10px;'>
        <h1 style='text-align: center; color: white; margin: 0;'>🚀 Product Price Comparison Dashboard</h1>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    "<p style='text-align:center; font-size: 18px; color: #333;'>Upload Marlin & Website price files to generate your detailed comparison report instantly.</p>",
    unsafe_allow_html=True
)

# File upload widgets
col1, col2 = st.columns(2)
with col1:
    marlin_file = st.file_uploader("Upload Marlin Price File", type=['xlsx'])
with col2:
    website_file = st.file_uploader("Upload Website Price File", type=['xlsx'])

# Run button
if st.button("Run Comparison", use_container_width=True):
    if marlin_file is None or website_file is None:
        st.error("Please upload both Excel files to proceed.")
    else:
        with st.spinner("Running comparison..."):
            # Read Excel into DataFrames
            marlin_df = pd.read_excel(marlin_file, sheet_name='Sheet1')
            website_df = pd.read_excel(website_file, sheet_name='Sheet1')

            # Clean column names
            marlin_df.columns = marlin_df.columns.str.strip()
            website_df.columns = website_df.columns.str.strip()

            # Merge and compute
            merged = website_df.merge(
                marlin_df, on='Variant Code', how='outer', suffixes=('_Website','_Marlin')
            )
            merged['Price Match'] = merged.apply(
                lambda row: 'Match' if row['Variant Price_Website'] == row['Variant Price_Marlin'] else 'Mismatch', axis=1
            )
            merged['Price Difference'] = merged['Variant Price_Website'] - merged['Variant Price_Marlin']

            def compare(row):
                if pd.isna(row['Variant Price_Website']):
                    return 'Only in Marlin'
                if pd.isna(row['Variant Price_Marlin']):
                    return 'Only in Website'
                return 'Website higher' if row['Price Difference']>0 else 'Marlin higher'

            merged['Comparison'] = merged.apply(compare, axis=1)

            # Prepare in-memory Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    merged.to_excel(writer, sheet_name='Full Data', index=False)
    merged[merged['Price Match']=='Match'].to_excel(writer, sheet_name='Matched', index=False)
    merged[merged['Price Match']=='Mismatch'].to_excel(writer, sheet_name='Mismatched', index=False)
    merged[merged['Comparison']=='Only in Website'].to_excel(writer, sheet_name='Only in Website', index=False)
    merged[merged['Comparison']=='Only in Marlin'].to_excel(writer, sheet_name='Only in Marlin', index=False)
    summary_df.to_excel(writer, sheet_name='Summary')

            data = output.getvalue()

        st.success("Report ready!")
        st.download_button(
            label="Download Comparison Report",
            data=data,
            file_name="Price_Comparison_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

