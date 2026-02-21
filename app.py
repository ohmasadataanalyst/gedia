
app_code = '''
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Geidea Terminal Summary Generator", layout="wide")

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

def process_data(df):
    """Process uploaded data"""
    df["Terminal"] = df["Terminal"].astype(str).str.strip().str.replace(".0", "", regex=False)
    df["Bank Name"] = df["Terminal"].map(TERMINAL_BANK_MAP).fillna("Unknown Bank")
    df["Total"] = df["Ter. Total Debit"].fillna(0) + df["Ter. Total Credit"].fillna(0)
    
    # For detailed view
    df["Total Debit"] = df["Ter. Total Debit"]
    df["Total Credit"] = df["Ter. Total Credit"]
    df["Total Debit Credit"] = df["Ter.Total Debit Credit"]
    
    # Check if Reconciliation Date column exists
    date_col = None
    for col in df.columns:
        if 'date' in col.lower() and 'recon' in col.lower():
            date_col = col
            break
    
    if date_col:
        df["Reconciliation Date"] = pd.to_datetime(df[date_col]).dt.date
    else:
        df["Reconciliation Date"] = None
    
    return df

def create_summary_file(df):
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

def create_summary_by_date_file(df):
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
        date_str = date_val.strftime("%A/%d/%b/%Y") if hasattr(date_val, 'strftime') else str(date_val)
        
        # Add date header row when date changes
        if date_val != current_date:
            if current_date is not None:
                # Add subtotal for previous date
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
            
            # New date header
            ws.cell(row=row_idx, column=1, value=date_str)
            ws.cell(row=row_idx, column=1).fill = date_fill
            ws.cell(row=row_idx, column=1).font = date_font
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            row_idx += 1
            current_date = date_val
            date_totals[current_date] = 0
        
        # Write data row
        ws.cell(row=row_idx, column=1, value="")  # Empty date cell (merged above)
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
    
    # Add final date subtotal
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
    
    # Grand Total
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

def create_detailed_file(df):
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
    
    # TOTAL row
    total_row = {"Bank Name": "TOTAL", "Card Scheme": "ALL"}
    for term in terminals:
        term_data = summary[summary["Terminal"] == term]
        total_row[f"{term}_Debit"] = term_data["Total Debit"].sum()
        total_row[f"{term}_Credit"] = term_data["Total Credit"].sum()
        total_row[f"{term}_Total"] = term_data["Total Debit Credit"].sum()
    rows.append(total_row)
    
    # AVG row
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

def create_detailed_by_date_file(df):
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
    
    # TOTAL row
    total_row = {"Bank Name": "TOTAL", "Card Scheme": "ALL"}
    for date in dates:
        date_data = summary[summary["Reconciliation Date"] == date]
        for term in terminals:
            term_data = date_data[date_data["Terminal"] == term]
            total_row[f"{date}_{term}_Debit"] = term_data["Total Debit"].sum()
            total_row[f"{date}_{term}_Credit"] = term_data["Total Credit"].sum()
            total_row[f"{date}_{term}_Total"] = term_data["Total Debit Credit"].sum()
    rows.append(total_row)
    
    # AVG row
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

# UI
st.title("🏦 Geidea Terminal Summary Generator")
st.markdown("Upload your terminal reconciliation file to generate Geidea reports")

uploaded_file = st.file_uploader("📁 Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ Loaded {len(df)} rows from {uploaded_file.name}")
        
        with st.expander("🔍 Preview Raw Data"):
            st.dataframe(df.head(10), use_container_width=True)
        
        with st.spinner("Processing Geidea reports..."):
            df_processed = process_data(df)
            summary_buffer, summary_df, grand_total = create_summary_file(df_processed)
            detailed_buffer, num_terminals = create_detailed_file(df_processed)
            
            unique_dates = df_processed["Reconciliation Date"].dropna().unique()
            has_multiple_dates = len(unique_dates) > 1
            
            if has_multiple_dates:
                summary_date_buffer, summary_date_df, num_dates = create_summary_by_date_file(df_processed)
                date_buffer, num_dates_detailed, num_terminals_date = create_detailed_by_date_file(df_processed)
            else:
                summary_date_buffer = None
                date_buffer = None
                num_dates = 0
        
        st.subheader("📊 Summary Preview")
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
                    label="📊 Geidea: Summary Totals Only\\n\\nConsolidated totals by Bank and Card Scheme across all dates",
                    data=summary_buffer,
                    file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.download_button(
                    label="📋 Geidea: Detailed All Terminals\\n\\nAll terminals as columns with Debit/Credit/Total breakdown",
                    data=detailed_buffer,
                    file_name=f"Geidea_02_DETAILED_All_Terminals_{num_terminals}_columns.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                st.download_button(
                    label="📅 Geidea: Summary by Date\\n\\nTotals grouped by Reconciliation Date with daily subtotals",
                    data=summary_date_buffer,
                    file_name=f"Geidea_03_SUMMARY_by_Date_{num_dates}_dates.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.download_button(
                    label="📆 Geidea: Detailed by Date\\n\\nTerminal breakdown grouped by each Reconciliation Date",
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
                    label="📊 Geidea: Summary Totals Only\\n\\nConsolidated totals by Bank and Card Scheme",
                    data=summary_buffer,
                    file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                st.download_button(
                    label="📋 Geidea: Detailed All Terminals\\n\\nAll terminals as columns with Debit/Credit/Total breakdown",
                    data=detailed_buffer,
                    file_name=f"Geidea_02_DETAILED_All_Terminals_{num_terminals}_columns.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.success(f"✅ Both Geidea reports ready! ({num_terminals} terminal columns)")
        
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        st.info("Please check your file has columns: Terminal, Card Name, Ter. Total Debit, Ter. Total Credit")

st.markdown("---")
st.caption("Geidea Terminal Summary Generator v4.0 | Multi-Date Reports with Zero Values")
'''

with open('app.py', 'w') as f:
    f.write(app_code)

print("✅ Updated app.py with Geidea branding and Summary by Date!")
print("\n🆕 New Features:")
print("   • Geidea: prefix on all button labels")
print("   • Descriptive subtitles under each button (using \\n\\n)")
print("   • NEW: Summary by Date report (03) when multiple dates detected")
print("   • Renumbered: 01-Summary, 02-Detailed, 03-Summary by Date, 04-Detailed by Date")
print("   • All filenames prefixed with 'Geidea_'")
print("\n📋 Button format:")
print("   📊 Geidea: Summary Totals Only")
print("   Consolidated totals by Bank and Card Scheme across all dates")
