
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

# Page config
st.set_page_config(page_title="Terminal Summary Generator", layout="wide")

# Terminal to Bank mapping
TERMINAL_BANK_MAP = {
    '63188996': 'Bank Al Bilad', '63189100': 'Bank Al Bilad', '63189101': 'Bank Al Bilad',
    '63189103': 'Bank Al Bilad', '63189105': 'Bank Al Bilad', '63189106': 'Bank Al Bilad',
    '63189107': 'Bank Al Bilad', '63189108': 'Bank Al Bilad', '63189110': 'Bank Al Bilad',
    '63189112': 'Bank Al Bilad', '63189113': 'Bank Al Bilad', '63189116': 'Bank Al Bilad',
    '63189117': 'Bank Al Bilad', '63189119': 'Bank Al Bilad', '63189120': 'Bank Al Bilad',
    '63189121': 'Bank Al Bilad', '63189122': 'Bank Al Bilad', '63189123': 'Bank Al Bilad',
    '63189124': 'Bank Al Bilad', '63189167': 'Bank Al Bilad', '63189168': 'Bank Al Bilad',
    '63189169': 'Bank Al Bilad', '63189490': 'Bank Al Bilad', '63189491': 'Bank Al Bilad',
    '63189492': 'Bank Al Bilad', '63189493': 'Bank Al Bilad', '63189494': 'Bank Al Bilad',
    '63189496': 'Bank Al Bilad', '63189497': 'Bank Al Bilad', '63189498': 'Bank Al Bilad',
    '63189499': 'Bank Al Bilad', '63189502': 'Bank Al Bilad', '63189503': 'Bank Al Bilad',
    '63189504': 'Bank Al Bilad', '63189506': 'Bank Al Bilad', '63189508': 'Bank Al Bilad',
    '63189510': 'Bank Al Bilad', '63189512': 'Bank Al Bilad', '63933955': 'Bank Al Bilad',
    '63933956': 'Bank Al Bilad', '63933957': 'Bank Al Bilad', '63933958': 'Bank Al Bilad',
    '63933959': 'Bank Al Bilad', '63934016': 'Bank Al Bilad', '63934017': 'Bank Al Bilad',
    '63934018': 'Bank Al Bilad', '63934019': 'Bank Al Bilad', '63934020': 'Bank Al Bilad',
    '63934021': 'Bank Al Bilad', '63934022': 'Bank Al Bilad', '63934023': 'Bank Al Bilad',
    '63934024': 'Bank Al Bilad', '63934025': 'Bank Al Bilad', '64729693': 'Bank Al Bilad',
    '64729694': 'Bank Al Bilad', '64729695': 'Bank Al Bilad', '64729696': 'Bank Al Bilad'
}

def process_data(df):
    """Process uploaded data"""
    # Clean terminal column
    df['Terminal'] = df['Terminal'].astype(str).str.strip().str.replace('.0', '', regex=False)
    
    # Add Bank Name
    df['Bank Name'] = df['Terminal'].map(TERMINAL_BANK_MAP).fillna('Unknown Bank')
    
    # Calculate Total
    df['Total'] = df['Ter. Total Debit'].fillna(0) + df['Ter. Total Credit'].fillna(0)
    
    # Aggregate by Bank + Card
    summary = df.groupby(['Bank Name', 'Card Name']).agg({
        'Total': 'sum'
    }).reset_index()
    
    # Sort
    summary['Sort'] = summary['Bank Name'].apply(lambda x: 1 if x == 'Unknown Bank' else 0)
    summary = summary.sort_values(['Sort', 'Bank Name', 'Card Name']).drop('Sort', axis=1)
    
    return summary, df['Bank Name'].unique()

def create_excel(summary):
    """Create formatted Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    # Styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    # Headers
    for col, header in enumerate(['Bank Name', 'Card Scheme', 'Total'], 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Data
    row_idx = 2
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data['Bank Name'])
        ws.cell(row=row_idx, column=2, value=data['Card Name'])
        ws.cell(row=row_idx, column=3, value=data['Total'])
        ws.cell(row=row_idx, column=3).number_format = '#,##0.00'
        ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="right")
        
        if data['Bank Name'] == 'Unknown Bank':
            for col in range(1, 4):
                ws.cell(row=row_idx, column=col).fill = unknown_fill
                ws.cell(row=row_idx, column=col).font = Font(bold=True, color="FFFFFF")
        row_idx += 1
    
    # Grand Total
    grand_total = summary['Total'].sum()
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=2, value="ALL")
    ws.cell(row=row_idx, column=3, value=grand_total)
    ws.cell(row=row_idx, column=3).number_format = '#,##0.00'
    for col in range(1, 4):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
    
    # Widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 15
    
    # Save to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, grand_total

# UI
st.title("🏦 Terminal Summary Generator")
st.markdown("Upload your terminal reconciliation Excel file to generate summary by Bank and Card Scheme")

with st.expander("📋 View Terminal Mapping"):
    st.write("**Mapped Terminals:**")
    st.json({k: v for k, v in list(TERMINAL_BANK_MAP.items())[:10]})
    st.caption(f"... and {len(TERMINAL_BANK_MAP)-10} more terminals all mapped to Bank Al Bilad")

# File upload
uploaded_file = st.file_uploader("📁 Upload Excel file", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ Loaded {len(df)} rows")
        
        # Show preview
        with st.expander("🔍 Preview Raw Data"):
            st.dataframe(df.head(20), use_container_width=True)
        
        # Process
        with st.spinner("Processing..."):
            summary, banks = process_data(df)
        
        # Show results
        st.subheader("📊 Summary Results")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Banks", len(banks))
        col2.metric("Card Schemes", summary['Card Name'].nunique())
        col3.metric("Total Rows", len(summary))
        
        # Unknown warning
        if 'Unknown Bank' in banks:
            unknown_count = len(summary[summary['Bank Name'] == 'Unknown Bank'])
            st.warning(f"⚠️ {unknown_count} rows with Unknown Bank (terminals not in mapping)")
        
        # Display table
        st.dataframe(
            summary.style.format({'Total': '{:,.2f}'})
            .apply(lambda x: ['background-color: #FF6B6B; color: white' if x['Bank Name'] == 'Unknown Bank' else '' for _ in x], axis=1),
            use_container_width=True,
            height=400
        )
        
        # Download
        excel_buffer, grand_total = create_excel(summary)
        
        st.subheader(f"💰 Grand Total: {grand_total:,.2f}")
        
        st.download_button(
            label="📥 Download Excel Summary",
            data=excel_buffer,
            file_name="terminal_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Please check your file format and try again.")

st.markdown("---")
st.caption("Made with Streamlit | Terminal Summary Generator v1.0")
