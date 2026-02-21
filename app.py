
# Update the file reading section to handle different Excel formats and engines

app_code = r'''
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="Geidea & Foodics Summary Generator", layout="wide")

# Terminal to Bank mapping
TERMINAL_BANK_MAP = {
    "63188996": "Bank Al Bilad", "63189100": "Bank Al Bilad", "63189101": "Bank Al Bilad",
    "63189103": "Bank Al Bilad", "63189105": "Bank Al Bilad", "63189106": "Bank Al Bilad",
    "63189107": "Bank Al Bilad", "63189108": "Bank Al Bilad", "63189110": "Bank Al Bilad",
    "63189112": "Bank Al Bilad", "63189113": "Bank Al Bilad", "63189116": "Bank Al Bilad",
    "63189117": "Bank Al Bilad", "63189119": "Bank Al Bilad", "63189120": "Bank Al Bilad",
    "63189121": "Bank Al Bilad", "63189122": "Bank Al Bilad", "63189123": "Bank Al Bilad",
    "63189124": "Bank Al Bilad", "63189167": "Bank Al Bilad", "63189168": "Bank Al Bilad",
    "63189169": "Bank Al Bilad", "63189490": "Bank Al Bilad", "63189491": "Bank Al Bilad",
    "63189492": "Bank Al Bilad", "63189493": "Bank Al Bilad", "63189494": "Bank Al Bilad",
    "63189496": "Bank Al Bilad", "63189497": "Bank Al Bilad", "63189498": "Bank Al Bilad",
    "63189499": "Bank Al Bilad", "63189502": "Bank Al Bilad", "63189503": "Bank Al Bilad",
    "63189504": "Bank Al Bilad", "63189506": "Bank Al Bilad", "63189508": "Bank Al Bilad",
    "63189510": "Bank Al Bilad", "63189512": "Bank Al Bilad", "63933955": "Bank Al Bilad",
    "63933956": "Bank Al Bilad", "63933957": "Bank Al Bilad", "63933958": "Bank Al Bilad",
    "63933959": "Bank Al Bilad", "63934016": "Bank Al Bilad", "63934017": "Bank Al Bilad",
    "63934018": "Bank Al Bilad", "63934019": "Bank Al Bilad", "63934020": "Bank Al Bilad",
    "63934021": "Bank Al Bilad", "63934022": "Bank Al Bilad", "63934023": "Bank Al Bilad",
    "63934024": "Bank Al Bilad", "63934025": "Bank Al Bilad", "64729693": "Bank Al Bilad",
    "64729694": "Bank Al Bilad", "64729695": "Bank Al Bilad", "64729696": "Bank Al Bilad"
}

def read_excel_file(uploaded_file):
    """Read Excel file with multiple engine attempts"""
    try:
        # Try openpyxl first (for .xlsx)
        return pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e1:
        try:
            # Try xlrd for older .xls files
            return pd.read_excel(uploaded_file, engine='xlrd')
        except Exception as e2:
            try:
                # Try without specifying engine
                return pd.read_excel(uploaded_file)
            except Exception as e3:
                raise Exception(f"Could not read Excel file. Tried openpyxl, xlrd, and default engine. Errors: {e1}; {e2}; {e3}")

def detect_file_type(df):
    """Detect if file is Geidea or Foodics format"""
    columns = [str(col).lower() for col in df.columns]
    
    # Check for Foodics indicators
    if any("payment method" in col for col in columns) and any("branch" in col for col in columns):
        return "foodics"
    
    # Check for Geidea indicators  
    if "terminal" in columns and "card name" in columns:
        return "geidea"
    
    return "unknown"

def parse_foodics_date_range(date_range_str):
    """Extract dates from Foodics date range string"""
    try:
        # Pattern: 2026-02-01 - 2026-02-03
        dates = re.findall(r"\d{4}-\d{2}-\d{2}", str(date_range_str))
        if len(dates) >= 2:
            start_date = pd.to_datetime(dates[0]).date()
            end_date = pd.to_datetime(dates[1]).date()
            # Generate all dates in range
            date_list = pd.date_range(start=start_date, end=end_date, freq="D").date.tolist()
            return date_list
    except:
        pass
    return []

def process_geidea_data(df):
    """Process Geidea uploaded data"""
    df["Terminal"] = df["Terminal"].astype(str).str.strip().str.replace(".0", "", regex=False)
    df["Bank Name"] = df["Terminal"].map(TERMINAL_BANK_MAP).fillna("Unknown Bank")
    df["Total"] = df["Ter. Total Debit"].fillna(0) + df["Ter. Total Credit"].fillna(0)
    
    df["Total Debit"] = df["Ter. Total Debit"]
    df["Total Credit"] = df["Ter. Total Credit"]
    df["Total Debit Credit"] = df["Ter.Total Debit Credit"]
    
    # Check if Reconciliation Date column exists
    date_col = None
    for col in df.columns:
        if "date" in col.lower() and "recon" in col.lower():
            date_col = col
            break
    
    if date_col:
        df["Reconciliation Date"] = pd.to_datetime(df[date_col]).dt.date
    else:
        df["Reconciliation Date"] = None
    
    return df

def process_foodics_data(df):
    """Process Foodics Payments Report data"""
    # Find the actual data start (skip title rows)
    data_start = 0
    for idx, row in df.iterrows():
        if "Payment Method" in str(row.values):
            data_start = idx
            break
    
    # Re-read with correct header
    df_clean = df.iloc[data_start:].reset_index(drop=True)
    df_clean.columns = df_clean.iloc[0]
    df_clean = df_clean[1:].reset_index(drop=True)
    
    # Clean column names
    df_clean.columns = [str(col).strip() for col in df_clean.columns]
    
    # Extract date range from metadata
    date_range = ""
    for idx, row in df.iterrows():
        if "Date Range" in str(row.values):
            date_range = row.iloc[1] if len(row) > 1 else ""
            break
    
    dates = parse_foodics_date_range(date_range)
    
    # Process data
    df_clean["Net Amount"] = pd.to_numeric(df_clean["Net Amount"], errors="coerce").fillna(0)
    df_clean["Amount"] = pd.to_numeric(df_clean["Amount"], errors="coerce").fillna(0)
    df_clean["Return Amount"] = pd.to_numeric(df_clean["Return Amount"], errors="coerce").fillna(0)
    df_clean["Count"] = pd.to_numeric(df_clean["Count"], errors="coerce").fillna(0).astype(int)
    
    # Add date info
    df_clean["Report Date Range"] = date_range
    df_clean["Dates"] = [dates] * len(df_clean)
    
    return df_clean, dates

# ==================== GEIDEA FUNCTIONS ====================

def create_geidea_summary_file(df):
    """Create simple summary by Bank + Card (Totals only)"""
    summary = df.groupby(["Bank Name", "Card Name"]).agg({
        "Total": "sum"
    }).reset_index()
    
    summary["Sort"] = summary["Bank Name"].apply(lambda x: 1 if x == "Unknown Bank" else 0)
    summary = summary.sort_values(["Sort", "Bank Name", "Card Name"]).drop("Sort", axis=1)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    headers = ["Bank Name", "Card Scheme", "Total"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    row_idx = 2
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data["Bank Name"])
        ws.cell(row=row_idx, column=2, value=data["Card Name"])
        ws.cell(row=row_idx, column=3, value=data["Total"])
        ws.cell(row=row_idx, column=3).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=3).alignment = Alignment(horizontal="right")
        
        if data["Bank Name"] == "Unknown Bank":
            for col in range(1, 4):
                ws.cell(row=row_idx, column=col).fill = unknown_fill
                ws.cell(row=row_idx, column=col).font = Font(bold=True, color="FFFFFF")
        row_idx += 1
    
    grand_total = summary["Total"].sum()
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=2, value="ALL")
    ws.cell(row=row_idx, column=3, value=grand_total)
    ws.cell(row=row_idx, column=3).number_format = "#,##0.00"
    for col in range(1, 4):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
    
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 15
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, summary, grand_total

def create_geidea_summary_by_date_file(df):
    """Create summary by Date + Bank + Card"""
    if df["Reconciliation Date"].isna().all():
        return None, None, 0
    
    summary = df.groupby(["Reconciliation Date", "Bank Name", "Card Name"]).agg({
        "Total": "sum"
    }).reset_index()
    
    summary["Sort"] = summary["Bank Name"].apply(lambda x: 1 if x == "Unknown Bank" else 0)
    summary = summary.sort_values(["Reconciliation Date", "Sort", "Bank Name", "Card Name"]).drop("Sort", axis=1)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary_by_Date"
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    date_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    date_font = Font(color="FFFFFF", bold=True, size=11)
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    subtotal_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    
    headers = ["Date", "Bank Name", "Card Scheme", "Total"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    row_idx = 2
    current_date = None
    date_totals = {}
    
    for _, data in summary.iterrows():
        date_val = data["Reconciliation Date"]
        date_str = date_val.strftime("%A/%d/%b/%Y") if hasattr(date_val, "strftime") else str(date_val)
        
        if date_val != current_date:
            if current_date is not None:
                row_idx += 1
                ws.cell(row=row_idx, column=1, value="")
                ws.cell(row=row_idx, column=2, value="DATE SUBTOTAL")
                ws.cell(row=row_idx, column=3, value="")
                ws.cell(row=row_idx, column=4, value=date_totals[current_date])
                ws.cell(row=row_idx, column=4).number_format = "#,##0.00"
                ws.cell(row=row_idx, column=4).font = Font(bold=True)
                for col in range(1, 5):
                    ws.cell(row=row_idx, column=col).fill = subtotal_fill
                row_idx += 1
            
            ws.cell(row=row_idx, column=1, value=date_str)
            ws.cell(row=row_idx, column=1).fill = date_fill
            ws.cell(row=row_idx, column=1).font = date_font
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            row_idx += 1
            current_date = date_val
            date_totals[current_date] = 0
        
        ws.cell(row=row_idx, column=1, value="")
        ws.cell(row=row_idx, column=2, value=data["Bank Name"])
        ws.cell(row=row_idx, column=3, value=data["Card Name"])
        ws.cell(row=row_idx, column=4, value=data["Total"])
        ws.cell(row=row_idx, column=4).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=4).alignment = Alignment(horizontal="right")
        
        if data["Bank Name"] == "Unknown Bank":
            for col in range(2, 5):
                ws.cell(row=row_idx, column=col).fill = unknown_fill
                ws.cell(row=row_idx, column=col).font = Font(bold=True, color="FFFFFF")
        
        date_totals[current_date] += data["Total"]
        row_idx += 1
    
    if current_date is not None:
        row_idx += 1
        ws.cell(row=row_idx, column=1, value="")
        ws.cell(row=row_idx, column=2, value="DATE SUBTOTAL")
        ws.cell(row=row_idx, column=3, value="")
        ws.cell(row=row_idx, column=4, value=date_totals[current_date])
        ws.cell(row=row_idx, column=4).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=4).font = Font(bold=True)
        for col in range(1, 5):
            ws.cell(row=row_idx, column=col).fill = subtotal_fill
        row_idx += 1
    
    row_idx += 1
    grand_total = summary["Total"].sum()
    ws.cell(row=row_idx, column=1, value="")
    ws.cell(row=row_idx, column=2, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=3, value="ALL DATES")
    ws.cell(row=row_idx, column=4, value=grand_total)
    ws.cell(row=row_idx, column=4).number_format = "#,##0.00"
    for col in range(1, 5):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
    
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, summary, len(summary["Reconciliation Date"].unique())

def create_geidea_detailed_file(df):
    """Create detailed file with all terminals as columns"""
    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit": "sum",
        "Total Credit": "sum",
        "Total Debit Credit": "sum"
    }).reset_index()
    
    terminals = sorted(summary["Terminal"].unique())
    banks = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())
    
    rows = []
    for bank in banks:
        for card in card_schemes:
            bank_card_data = summary[(summary["Bank Name"] == bank) & (summary["Card Name"] == card)]
            if len(bank_card_data) == 0:
                continue
            
            row = {"Bank Name": bank, "Card Scheme": card}
            for term in terminals:
                term_data = bank_card_data[bank_card_data["Terminal"] == term]
                if len(term_data) > 0:
                    row[f"{term}_Debit"] = term_data["Total Debit"].values[0]
                    row[f"{term}_Credit"] = term_data["Total Credit"].values[0]
                    row[f"{term}_Total"] = term_data["Total Debit Credit"].values[0]
                else:
                    row[f"{term}_Debit"] = 0
                    row[f"{term}_Credit"] = 0
                    row[f"{term}_Total"] = 0
            rows.append(row)
    
    total_row = {"Bank Name": "TOTAL", "Card Scheme": "ALL"}
    for term in terminals:
        term_data = summary[summary["Terminal"] == term]
        total_row[f"{term}_Debit"] = term_data["Total Debit"].sum()
        total_row[f"{term}_Credit"] = term_data["Total Credit"].sum()
        total_row[f"{term}_Total"] = term_data["Total Debit Credit"].sum()
    rows.append(total_row)
    
    avg_row = {"Bank Name": "AVG", "Card Scheme": "ALL"}
    for term in terminals:
        term_data = summary[summary["Terminal"] == term]
        avg_row[f"{term}_Debit"] = round(term_data["Total Debit"].mean(), 2)
        avg_row[f"{term}_Credit"] = round(term_data["Total Credit"].mean(), 2)
        avg_row[f"{term}_Total"] = round(term_data["Total Debit Credit"].mean(), 2)
    rows.append(avg_row)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Detailed"
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=9)
    sub_header_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    sub_header_font = Font(bold=True, size=8)
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    ws.cell(row=1, column=1, value="Bank Name")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = header_font
    ws.cell(row=1, column=1).alignment = center_align
    
    ws.cell(row=1, column=2, value="Card Scheme")
    ws.cell(row=1, column=2).fill = header_fill
    ws.cell(row=1, column=2).font = header_font
    ws.cell(row=1, column=2).alignment = center_align
    
    col_idx = 3
    for term in terminals:
        ws.cell(row=1, column=col_idx, value=f"#{term}")
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).font = header_font
        ws.cell(row=1, column=col_idx).alignment = center_align
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx+2)
        
        ws.cell(row=2, column=col_idx, value="Debit").fill = sub_header_fill
        ws.cell(row=2, column=col_idx).font = sub_header_font
        ws.cell(row=2, column=col_idx).alignment = center_align
        
        ws.cell(row=2, column=col_idx+1, value="Credit").fill = sub_header_fill
        ws.cell(row=2, column=col_idx+1).font = sub_header_font
        ws.cell(row=2, column=col_idx+1).alignment = center_align
        
        ws.cell(row=2, column=col_idx+2, value="Total").fill = sub_header_fill
        ws.cell(row=2, column=col_idx+2).font = sub_header_font
        ws.cell(row=2, column=col_idx+2).alignment = center_align
        
        col_idx += 3
    
    for r_idx, row_data in enumerate(rows, 3):
        bank_val = row_data["Bank Name"]
        card_val = row_data["Card Scheme"]
        
        cell = ws.cell(row=r_idx, column=1, value=bank_val)
        if bank_val == "Unknown Bank":
            cell.fill = unknown_fill
            cell.font = Font(bold=True, color="FFFFFF")
        elif bank_val in ["TOTAL", "AVG"]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")
        
        cell = ws.cell(row=r_idx, column=2, value=card_val)
        if bank_val == "Unknown Bank":
            cell.fill = unknown_fill
            cell.font = Font(bold=True, color="FFFFFF")
        elif bank_val in ["TOTAL", "AVG"]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")
        
        col_idx = 3
        for term in terminals:
            debit = row_data[f"{term}_Debit"]
            credit = row_data[f"{term}_Credit"]
            total = row_data[f"{term}_Total"]
            
            ws.cell(row=r_idx, column=col_idx, value=debit)
            ws.cell(row=r_idx, column=col_idx+1, value=credit)
            ws.cell(row=r_idx, column=col_idx+2, value=total)
            
            ws.cell(row=r_idx, column=col_idx).number_format = "#,##0.00"
            ws.cell(row=r_idx, column=col_idx+1).number_format = "#,##0.00"
            ws.cell(row=r_idx, column=col_idx+2).number_format = "#,##0.00"
            
            col_idx += 3
    
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15
    for i in range(3, col_idx):
        ws.column_dimensions[get_column_letter(i)].width = 11
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, len(terminals)

def create_geidea_detailed_by_date_file(df):
    """Create detailed file grouped by Reconciliation Date"""
    if df["Reconciliation Date"].isna().all():
        return None, 0, 0
    
    summary = df.groupby(["Reconciliation Date", "Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit": "sum",
        "Total Credit": "sum",
        "Total Debit Credit": "sum"
    }).reset_index()
    
    dates = sorted(summary["Reconciliation Date"].unique())
    terminals = sorted(summary["Terminal"].unique())
    banks = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Detailed_by_Date"
    
    date_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    date_font = Font(color="FFFFFF", bold=True, size=11)
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=9)
    sub_header_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    sub_header_font = Font(bold=True, size=8)
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    rows = []
    for bank in banks:
        for card in card_schemes:
            bank_card_data = summary[(summary["Bank Name"] == bank) & (summary["Card Name"] == card)]
            if len(bank_card_data) == 0:
                continue
            
            row = {"Bank Name": bank, "Card Scheme": card}
            
            for date in dates:
                date_data = bank_card_data[bank_card_data["Reconciliation Date"] == date]
                for term in terminals:
                    term_data = date_data[date_data["Terminal"] == term]
                    if len(term_data) > 0:
                        row[f"{date}_{term}_Debit"] = term_data["Total Debit"].values[0]
                        row[f"{date}_{term}_Credit"] = term_data["Total Credit"].values[0]
                        row[f"{date}_{term}_Total"] = term_data["Total Debit Credit"].values[0]
                    else:
                        row[f"{date}_{term}_Debit"] = 0
                        row[f"{date}_{term}_Credit"] = 0
                        row[f"{date}_{term}_Total"] = 0
            rows.append(row)
    
    total_row = {"Bank Name": "TOTAL", "Card Scheme": "ALL"}
    for date in dates:
        date_data = summary[summary["Reconciliation Date"] == date]
        for term in terminals:
            term_data = date_data[date_data["Terminal"] == term]
            total_row[f"{date}_{term}_Debit"] = term_data["Total Debit"].sum()
            total_row[f"{date}_{term}_Credit"] = term_data["Total Credit"].sum()
            total_row[f"{date}_{term}_Total"] = term_data["Total Debit Credit"].sum()
    rows.append(total_row)
    
    avg_row = {"Bank Name": "AVG", "Card Scheme": "ALL"}
    for date in dates:
        date_data = summary[summary["Reconciliation Date"] == date]
        for term in terminals:
            term_data = date_data[date_data["Terminal"] == term]
            avg_row[f"{date}_{term}_Debit"] = round(term_data["Total Debit"].mean(), 2)
            avg_row[f"{date}_{term}_Credit"] = round(term_data["Total Credit"].mean(), 2)
            avg_row[f"{date}_{term}_Total"] = round(term_data["Total Debit Credit"].mean(), 2)
    rows.append(avg_row)
    
    ws.cell(row=1, column=1, value="Bank Name")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = header_font
    ws.cell(row=1, column=1).alignment = center_align
    
    ws.cell(row=1, column=2, value="Card Scheme")
    ws.cell(row=1, column=2).fill = header_fill
    ws.cell(row=1, column=2).font = header_font
    ws.cell(row=1, column=2).alignment = center_align
    
    col_idx = 3
    for date in dates:
        date_str = date.strftime("%A/%d/%b/%Y")
        ws.cell(row=1, column=col_idx, value=date_str)
        ws.cell(row=1, column=col_idx).fill = date_fill
        ws.cell(row=1, column=col_idx).font = date_font
        ws.cell(row=1, column=col_idx).alignment = center_align
        
        end_col = col_idx + (len(terminals) * 3) - 1
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=end_col)
        
        term_col = col_idx
        for term in terminals:
            ws.cell(row=2, column=term_col, value=f"#{term}")
            ws.cell(row=2, column=term_col).fill = header_fill
            ws.cell(row=2, column=term_col).font = header_font
            ws.cell(row=2, column=term_col).alignment = center_align
            ws.merge_cells(start_row=2, start_column=term_col, end_row=2, end_column=term_col+2)
            
            ws.cell(row=3, column=term_col, value="Debit").fill = sub_header_fill
            ws.cell(row=3, column=term_col).font = sub_header_font
            ws.cell(row=3, column=term_col).alignment = center_align
            
            ws.cell(row=3, column=term_col+1, value="Credit").fill = sub_header_fill
            ws.cell(row=3, column=term_col+1).font = sub_header_font
            ws.cell(row=3, column=term_col+1).alignment = center_align
            
            ws.cell(row=3, column=term_col+2, value="Total").fill = sub_header_fill
            ws.cell(row=3, column=term_col+2).font = sub_header_font
            ws.cell(row=3, column=term_col+2).alignment = center_align
            
            term_col += 3
        
        col_idx = end_col + 1
    
    for r_idx, row_data in enumerate(rows, 4):
        bank_val = row_data["Bank Name"]
        card_val = row_data["Card Scheme"]
        
        cell = ws.cell(row=r_idx, column=1, value=bank_val)
        if bank_val == "Unknown Bank":
            cell.fill = unknown_fill
            cell.font = Font(bold=True, color="FFFFFF")
        elif bank_val in ["TOTAL", "AVG"]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")
        
        cell = ws.cell(row=r_idx, column=2, value=card_val)
        if bank_val == "Unknown Bank":
            cell.fill = unknown_fill
            cell.font = Font(bold=True, color="FFFFFF")
        elif bank_val in ["TOTAL", "AVG"]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid")
        
        col_idx = 3
        for date in dates:
            for term in terminals:
                debit = row_data[f"{date}_{term}_Debit"]
                credit = row_data[f"{date}_{term}_Credit"]
                total = row_data[f"{date}_{term}_Total"]
                
                ws.cell(row=r_idx, column=col_idx, value=debit)
                ws.cell(row=r_idx, column=col_idx+1, value=credit)
                ws.cell(row=r_idx, column=col_idx+2, value=total)
                
                ws.cell(row=r_idx, column=col_idx).number_format = "#,##0.00"
                ws.cell(row=r_idx, column=col_idx+1).number_format = "#,##0.00"
                ws.cell(row=r_idx, column=col_idx+2).number_format = "#,##0.00"
                
                col_idx += 3
    
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15
    for i in range(3, col_idx):
        ws.column_dimensions[get_column_letter(i)].width = 11
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, len(dates), len(terminals)

# ==================== FOODICS FUNCTIONS ====================

def create_foodics_summary_by_branch(df):
    """Create Foodics summary grouped by Branch"""
    summary = df.groupby(["Branch", "Payment Method"]).agg({
        "Net Amount": "sum",
        "Amount": "sum",
        "Return Amount": "sum",
        "Count": "sum"
    }).reset_index()
    
    summary = summary.sort_values(["Branch", "Net Amount"], ascending=[True, False])
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary_by_Branch"
    
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    branch_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    branch_font = Font(color="FFFFFF", bold=True, size=10)
    subtotal_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    headers = ["Branch", "Payment Method", "Net Amount", "Amount", "Returns", "Count"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    
    row_idx = 2
    current_branch = None
    branch_totals = {}
    
    for _, data in summary.iterrows():
        branch = data["Branch"]
        
        if branch != current_branch:
            if current_branch is not None:
                row_idx += 1
                ws.cell(row=row_idx, column=1, value="")
                ws.cell(row=row_idx, column=2, value="BRANCH SUBTOTAL")
                ws.cell(row=row_idx, column=3, value=branch_totals[current_branch]["net"])
                ws.cell(row=row_idx, column=4, value=branch_totals[current_branch]["amount"])
                ws.cell(row=row_idx, column=5, value=branch_totals[current_branch]["returns"])
                ws.cell(row=row_idx, column=6, value=branch_totals[current_branch]["count"])
                
                for col in range(1, 7):
                    ws.cell(row=row_idx, column=col).fill = subtotal_fill
                    ws.cell(row=row_idx, column=col).font = Font(bold=True)
                    if col >= 3:
                        ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
                row_idx += 1
            
            ws.cell(row=row_idx, column=1, value=branch)
            ws.cell(row=row_idx, column=1).fill = branch_fill
            ws.cell(row=row_idx, column=1).font = branch_font
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
            row_idx += 1
            current_branch = branch
            branch_totals[current_branch] = {"net": 0, "amount": 0, "returns": 0, "count": 0}
        
        ws.cell(row=row_idx, column=1, value="")
        ws.cell(row=row_idx, column=2, value=data["Payment Method"])
        ws.cell(row=row_idx, column=3, value=data["Net Amount"])
        ws.cell(row=row_idx, column=4, value=data["Amount"])
        ws.cell(row=row_idx, column=5, value=data["Return Amount"])
        ws.cell(row=row_idx, column=6, value=data["Count"])
        
        for col in range(3, 7):
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col < 6 else "#,##0"
        
        branch_totals[current_branch]["net"] += data["Net Amount"]
        branch_totals[current_branch]["amount"] += data["Amount"]
        branch_totals[current_branch]["returns"] += data["Return Amount"]
        branch_totals[current_branch]["count"] += data["Count"]
        row_idx += 1
    
    if current_branch is not None:
        row_idx += 1
        ws.cell(row=row_idx, column=1, value="")
        ws.cell(row=row_idx, column=2, value="BRANCH SUBTOTAL")
        ws.cell(row=row_idx, column=3, value=branch_totals[current_branch]["net"])
        ws.cell(row=row_idx, column=4, value=branch_totals[current_branch]["amount"])
        ws.cell(row=row_idx, column=5, value=branch_totals[current_branch]["returns"])
        ws.cell(row=row_idx, column=6, value=branch_totals[current_branch]["count"])
        
        for col in range(1, 7):
            ws.cell(row=row_idx, column=col).fill = subtotal_fill
            ws.cell(row=row_idx, column=col).font = Font(bold=True)
            if col >= 3:
                ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
        row_idx += 1
    
    row_idx += 1
    grand_net = summary["Net Amount"].sum()
    grand_amount = summary["Amount"].sum()
    grand_returns = summary["Return Amount"].sum()
    grand_count = summary["Count"].sum()
    
    ws.cell(row=row_idx, column=1, value="")
    ws.cell(row=row_idx, column=2, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=3, value=grand_net)
    ws.cell(row=row_idx, column=4, value=grand_amount)
    ws.cell(row=row_idx, column=5, value=grand_returns)
    ws.cell(row=row_idx, column=6, value=grand_count)
    
    for col in range(1, 7):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
        if col >= 3:
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
    
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 12
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, summary, len(summary["Branch"].unique())

def create_foodics_summary_by_payment_method(df):
    """Create Foodics summary grouped by Payment Method across all branches"""
    summary = df.groupby(["Payment Method"]).agg({
        "Net Amount": "sum",
        "Amount": "sum",
        "Return Amount": "sum",
        "Count": "sum"
    }).reset_index()
    
    summary = summary.sort_values("Net Amount", ascending=False)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary_by_Payment"
    
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    headers = ["Payment Method", "Net Amount", "Amount", "Returns", "Count"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    
    row_idx = 2
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data["Payment Method"])
        ws.cell(row=row_idx, column=2, value=data["Net Amount"])
        ws.cell(row=row_idx, column=3, value=data["Amount"])
        ws.cell(row=row_idx, column=4, value=data["Return Amount"])
        ws.cell(row=row_idx, column=5, value=data["Count"])
        
        for col in range(2, 6):
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col < 5 else "#,##0"
        row_idx += 1
    
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=2, value=summary["Net Amount"].sum())
    ws.cell(row=row_idx, column=3, value=summary["Amount"].sum())
    ws.cell(row=row_idx, column=4, value=summary["Return Amount"].sum())
    ws.cell(row=row_idx, column=5, value=summary["Count"].sum())
    
    for col in range(1, 6):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
        if col >= 2:
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
    
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 12
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, summary

def create_foodics_daily_avg_report(df, dates):
    """Create Foodics daily average report by Payment Method"""
    if not dates:
        return None
    
    num_days = len(dates)
    
    summary = df.groupby(["Payment Method"]).agg({
        "Net Amount": "sum",
        "Amount": "sum",
        "Return Amount": "sum",
        "Count": "sum"
    }).reset_index()
    
    summary["Avg Net Amount/Day"] = summary["Net Amount"] / num_days
    summary["Avg Count/Day"] = summary["Count"] / num_days
    summary = summary.sort_values("Avg Net Amount/Day", ascending=False)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily_Averages"
    
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    avg_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    
    ws.cell(row=1, column=1, value=f"Report Period: {dates[0]} to {dates[-1]} ({num_days} days)")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    
    headers = ["Payment Method", "Total Net Amount", "Daily Avg Net", "Total Count", "Daily Avg Count", "Total Returns", "Days"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    
    row_idx = 3
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data["Payment Method"])
        ws.cell(row=row_idx, column=2, value=data["Net Amount"])
        ws.cell(row=row_idx, column=3, value=data["Avg Net Amount/Day"])
        ws.cell(row=row_idx, column=4, value=data["Count"])
        ws.cell(row=row_idx, column=5, value=data["Avg Count/Day"])
        ws.cell(row=row_idx, column=6, value=data["Return Amount"])
        ws.cell(row=row_idx, column=7, value=num_days)
        
        ws.cell(row=row_idx, column=3).fill = avg_fill
        ws.cell(row=row_idx, column=5).fill = avg_fill
        
        for col in range(2, 7):
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col != 4 else "#,##0"
        row_idx += 1
    
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="GRAND TOTAL")
    ws.cell(row=row_idx, column=2, value=summary["Net Amount"].sum())
    ws.cell(row=row_idx, column=3, value=summary["Net Amount"].sum() / num_days)
    ws.cell(row=row_idx, column=4, value=summary["Count"].sum())
    ws.cell(row=row_idx, column=5, value=summary["Count"].sum() / num_days)
    ws.cell(row=row_idx, column=6, value=summary["Return Amount"].sum())
    ws.cell(row=row_idx, column=7, value=num_days)
    
    for col in range(1, 8):
        ws.cell(row=row_idx, column=col).fill = total_fill
        ws.cell(row=row_idx, column=col).font = Font(bold=True, size=12)
        if col >= 2 and col != 4:
            ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
    
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["G"].width = 10
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, summary, num_days

# ==================== UI ====================

st.title("🏦 Geidea & Foodics Summary Generator")
st.markdown("Upload your reconciliation file to generate summary reports")

uploaded_file = st.file_uploader("📁 Upload Excel file (Geidea or Foodics)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Use the new read function with multiple engine attempts
        df_raw = read_excel_file(uploaded_file)
        file_type = detect_file_type(df_raw)
        
        if file_type == "geidea":
            st.success(f"✅ Detected **Geidea** file: {uploaded_file.name} ({len(df_raw)} rows)")
            
            with st.expander("🔍 Preview Raw Data"):
                st.dataframe(df_raw.head(10), use_container_width=True)
            
            with st.spinner("Processing Geidea reports..."):
                df_processed = process_geidea_data(df_raw)
                summary_buffer, summary_df, grand_total = create_geidea_summary_file(df_processed)
                detailed_buffer, num_terminals = create_geidea_detailed_file(df_processed)
                
                unique_dates = df_processed["Reconciliation Date"].dropna().unique()
                has_multiple_dates = len(unique_dates) > 1
                
                if has_multiple_dates:
                    summary_date_buffer, summary_date_df, num_dates = create_geidea_summary_by_date_file(df_processed)
                    date_buffer, num_dates_detailed, num_terminals_date = create_geidea_detailed_by_date_file(df_processed)
                else:
                    summary_date_buffer = None
                    date_buffer = None
                    num_dates = 0
            
            st.subheader("📊 Geidea Summary Preview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Banks", summary_df["Bank Name"].nunique())
            col2.metric("Card Schemes", summary_df["Card Name"].nunique())
            col3.metric("Grand Total", f"{grand_total:,.0f}")
            
            if has_multiple_dates:
                st.info(f"📅 Detected {num_dates} reconciliation dates: {', '.join([d.strftime('%Y-%m-%d') for d in unique_dates])}")
            
            if "Unknown Bank" in summary_df["Bank Name"].values:
                st.warning("⚠️ Some terminals not found in mapping (shown in red)")
            
            st.dataframe(
                summary_df.style.format({"Total": "{:,.2f}"})
                .apply(lambda x: ["background-color: #FF6B6B; color: white"]*3 if x["Bank Name"]=="Unknown Bank" else [""]*3, axis=1),
                use_container_width=True,
                height=300
            )
            
            st.subheader("⬇️ Geidea: Download Reports")
            
            if has_multiple_dates:
                st.markdown("**Four Geidea reports generated with multi-date support:**")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📊 Geidea: Summary Totals Only\n\nConsolidated totals by Bank and Card Scheme across all dates",
                        data=summary_buffer,
                        file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.download_button(
                        label="📋 Geidea: Detailed All Terminals\n\nAll terminals as columns with Debit/Credit/Total breakdown",
                        data=detailed_buffer,
                        file_name=f"Geidea_02_DETAILED_All_Terminals_{num_terminals}_columns.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    st.download_button(
                        label="📅 Geidea: Summary by Date\n\nTotals grouped by Reconciliation Date with daily subtotals",
                        data=summary_date_buffer,
                        file_name=f"Geidea_03_SUMMARY_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.download_button(
                        label="📆 Geidea: Detailed by Date\n\nTerminal breakdown grouped by each Reconciliation Date",
                        data=date_buffer,
                        file_name=f"Geidea_04_DETAILED_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.success(f"✅ All 4 Geidea reports ready! ({num_dates} dates × {num_terminals} terminals)")
                
            else:
                st.markdown("**Two Geidea reports generated:**")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📊 Geidea: Summary Totals Only\n\nConsolidated totals by Bank and Card Scheme",
                        data=summary_buffer,
                        file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    st.download_button(
                        label="📋 Geidea: Detailed All Terminals\n\nAll terminals as columns with Debit/Credit/Total breakdown",
                        data=detailed_buffer,
                        file_name=f"Geidea_02_DETAILED_All_Terminals_{num_terminals}_columns.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.success(f"✅ Both Geidea reports ready! ({num_terminals} terminal columns)")
        
        elif file_type == "foodics":
            st.success(f"✅ Detected **Foodics** Payments Report: {uploaded_file.name}")
            
            with st.expander("🔍 Preview Raw Data"):
                st.dataframe(df_raw.head(15), use_container_width=True)
            
            with st.spinner("Processing Foodics reports..."):
                df_processed, dates = process_foodics_data(df_raw)
                
                branch_buffer, branch_summary, num_branches = create_foodics_summary_by_branch(df_processed)
                payment_buffer, payment_summary = create_foodics_summary_by_payment_method(df_processed)
                
                if dates:
                    avg_buffer, avg_summary, num_days = create_foodics_daily_avg_report(df_processed, dates)
                else:
                    avg_buffer = None
                    num_days = 0
            
            st.subheader("📊 Foodics Summary Preview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Branches", num_branches)
            col2.metric("Payment Methods", payment_summary["Payment Method"].nunique())
            col3.metric("Total Net Amount", f"{payment_summary['Net Amount'].sum():,.0f}")
            
            if dates:
                st.info(f"📅 Report period: {dates[0]} to {dates[-1]} ({num_days} days)")
            
            st.dataframe(
                payment_summary.style.format({
                    "Net Amount": "{:,.2f}",
                    "Amount": "{:,.2f}",
                    "Return Amount": "{:,.2f}",
                    "Count": "{:,.0f}"
                }),
                use_container_width=True,
                height=300
            )
            
            st.subheader("⬇️ Foodics: Download Reports")
            st.markdown("**Three Foodics reports generated:**")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="🏪 Foodics: Summary by Branch\n\nNet Amount grouped by Branch and Payment Method with subtotals",
                    data=branch_buffer,
                    file_name=f"Foodics_01_SUMMARY_by_Branch_{num_branches}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                if dates:
                    st.download_button(
                        label="📈 Foodics: Daily Averages\n\nAverage Net Amount and Count per day by Payment Method",
                        data=avg_buffer,
                        file_name=f"Foodics_03_Daily_Averages_{num_days}_days.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            with col2:
                st.download_button(
                    label="💳 Foodics: Summary by Payment Method\n\nConsolidated totals across all branches by Payment Type",
                    data=payment_buffer,
                    file_name="Foodics_02_SUMMARY_by_Payment_Method.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            if dates:
                st.success(f"✅ All 3 Foodics reports ready! ({num_branches} branches, {num_days} days)")
            else:
                st.success(f"✅ 2 Foodics reports ready! ({num_branches} branches)")
        
        else:
            st.error("❌ Could not detect file type. Please ensure your file matches Geidea or Foodics format.")
            st.info("Geidea: Should have 'Terminal', 'Card Name' columns")
            st.info("Foodics: Should have 'Payment Method', 'Branch' columns")
        
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("Please check your file format and try again.")

st.markdown("---")
st.caption("Geidea & Foodics Summary Generator v5.1 | Multi-Platform Support with Enhanced File Reading")
'''

with open('app.py', 'w') as f:
    f.write(app_code)

print("✅ Fixed file reading issue!")
print("\n🔧 Changes made:")
print("   • Added read_excel_file() function with multiple engine attempts")
print("   • Tries openpyxl first (for .xlsx files)")
print("   • Falls back to xlrd (for older .xls files)")
print("   • Finally tries default engine as last resort")
print("   • Provides detailed error messages if all engines fail")
print("\n📋 Supported formats:")
print("   • .xlsx (Excel 2007+, uses openpyxl)")
print("   • .xls (Excel 97-2003, uses xlrd)")
print("   • Any Excel format pandas can detect")
