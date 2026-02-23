import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="Geidea & Foodics Summary Generator", layout="wide")

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


# ==================== FILE READING & DETECTION ====================

def read_uploaded_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith('.xlsx'):
            return pd.read_excel(uploaded_file, engine='openpyxl')
        elif file_name.endswith('.xls'):
            try:
                return pd.read_excel(uploaded_file, engine='xlrd')
            except ImportError:
                st.error("Missing 'xlrd' package. Please install: pip install xlrd>=2.0.1")
                raise
        elif file_name.endswith('.csv'):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep='\t')
                if df.shape[1] > 1:
                    return df
            except Exception:
                pass
            uploaded_file.seek(0)
            try:
                return pd.read_csv(uploaded_file)
            except Exception:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, sep=';')
        else:
            try:
                return pd.read_excel(uploaded_file, engine='openpyxl')
            except Exception:
                uploaded_file.seek(0)
                try:
                    return pd.read_excel(uploaded_file, engine='xlrd')
                except Exception:
                    uploaded_file.seek(0)
                    return pd.read_csv(uploaded_file)
    except Exception as e:
        raise Exception(f"Could not read file '{uploaded_file.name}'. Error: {str(e)}")


def detect_file_type(df):
    columns = [str(col).lower().strip() for col in df.columns]
    if any("payment method" in col for col in columns) and any("branch" in col for col in columns):
        return "foodics"
    if "terminal" in columns and "card name" in columns:
        return "geidea"
    try:
        all_values = df.astype(str).values.flatten()
        if any("Payment Method" in v for v in all_values) and any("Branch" in v for v in all_values):
            return "foodics"
    except Exception:
        pass
    return "unknown"


def parse_foodics_date_range(date_range_str):
    try:
        dates = re.findall(r"\d{4}-\d{2}-\d{2}", str(date_range_str))
        if len(dates) >= 2:
            start_date = pd.to_datetime(dates[0]).date()
            end_date = pd.to_datetime(dates[1]).date()
            return pd.date_range(start=start_date, end=end_date, freq="D").date.tolist()
    except Exception:
        pass
    return []


# ==================== DATA PROCESSING ====================

def process_geidea_data(df):
    df["Terminal"] = df["Terminal"].astype(str).str.strip().str.replace(".0", "", regex=False)
    df["Bank Name"] = df["Terminal"].map(TERMINAL_BANK_MAP).fillna("Unknown Bank")
    df["Total"] = df["Ter. Total Debit"].fillna(0) + df["Ter. Total Credit"].fillna(0)
    df["Total Debit"] = df["Ter. Total Debit"]
    df["Total Credit"] = df["Ter. Total Credit"]
    df["Total Debit Credit"] = df["Ter.Total Debit Credit"]
    date_col = next((col for col in df.columns if "date" in col.lower() and "recon" in col.lower()), None)
    df["Reconciliation Date"] = pd.to_datetime(df[date_col]).dt.date if date_col else None
    return df


def process_foodics_data(df):
    columns_lower = [str(col).lower().strip() for col in df.columns]
    if "payment method" in columns_lower and "branch" in columns_lower:
        df_clean = df.copy()
        df_clean.columns = [str(col).strip() for col in df_clean.columns]
        date_range = ""
        dates = []
    else:
        data_start = 0
        for idx, row in df.iterrows():
            if "Payment Method" in str(row.values):
                data_start = idx
                break
        df_clean = df.iloc[data_start:].reset_index(drop=True)
        df_clean.columns = df_clean.iloc[0]
        df_clean = df_clean[1:].reset_index(drop=True)
        df_clean.columns = [str(col).strip() for col in df_clean.columns]
        date_range = ""
        for idx, row in df.iterrows():
            if "Date Range" in str(row.values):
                date_range = row.iloc[1] if len(row) > 1 else ""
                break
        dates = parse_foodics_date_range(date_range)

    df_clean["Net Amount"] = pd.to_numeric(df_clean["Net Amount"], errors="coerce").fillna(0)
    df_clean["Amount"] = pd.to_numeric(df_clean["Amount"], errors="coerce").fillna(0)
    df_clean["Return Amount"] = pd.to_numeric(df_clean["Return Amount"], errors="coerce").fillna(0)
    df_clean["Count"] = pd.to_numeric(df_clean["Count"], errors="coerce").fillna(0).astype(int)
    df_clean["Report Date Range"] = date_range
    df_clean["Dates"] = [dates] * len(df_clean)
    return df_clean, dates


# ==================== GEIDEA FUNCTIONS ====================

def create_geidea_summary_file(df):
    summary = df.groupby(["Bank Name", "Card Name"]).agg({"Total": "sum"}).reset_index()
    summary["Sort"] = summary["Bank Name"].apply(lambda x: 1 if x == "Unknown Bank" else 0)
    summary = summary.sort_values(["Sort", "Bank Name", "Card Name"]).drop("Sort", axis=1)

    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    header_fill  = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill   = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

    for col, header in enumerate(["Bank Name", "Card Scheme", "Total"], 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 2
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data["Bank Name"])
        ws.cell(row=row_idx, column=2, value=data["Card Name"])
        c = ws.cell(row=row_idx, column=3, value=data["Total"])
        c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right")
        if data["Bank Name"] == "Unknown Bank":
            for col in range(1, 4):
                ws.cell(row=row_idx, column=col).fill = unknown_fill
                ws.cell(row=row_idx, column=col).font = Font(bold=True, color="FFFFFF")
        row_idx += 1

    grand_total = summary["Total"].sum()
    row_idx += 1
    for col, val in {1: "GRAND TOTAL", 2: "ALL", 3: grand_total}.items():
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True, size=12)
    ws.cell(row=row_idx, column=3).number_format = "#,##0.00"

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 15
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, summary, grand_total


def create_geidea_summary_by_date_file(df):
    if df["Reconciliation Date"].isna().all():
        return None, None, 0

    summary = df.groupby(["Reconciliation Date", "Bank Name", "Card Name"]).agg({"Total": "sum"}).reset_index()
    summary["Sort"] = summary["Bank Name"].apply(lambda x: 1 if x == "Unknown Bank" else 0)
    summary = summary.sort_values(["Reconciliation Date", "Sort", "Bank Name", "Card Name"]).drop("Sort", axis=1)

    wb = Workbook(); ws = wb.active; ws.title = "Summary_by_Date"
    header_fill   = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    date_fill     = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    unknown_fill  = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill    = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    subtotal_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")

    for col, header in enumerate(["Date", "Bank Name", "Card Scheme", "Total"], 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 2; current_date = None; date_totals = {}
    for _, data in summary.iterrows():
        date_val = data["Reconciliation Date"]
        date_str = date_val.strftime("%A/%d/%b/%Y") if hasattr(date_val, "strftime") else str(date_val)
        if date_val != current_date:
            if current_date is not None:
                row_idx += 1
                ws.cell(row=row_idx, column=2, value="DATE SUBTOTAL")
                c = ws.cell(row=row_idx, column=4, value=date_totals[current_date])
                c.number_format = "#,##0.00"; c.font = Font(bold=True)
                for col in range(1, 5): ws.cell(row=row_idx, column=col).fill = subtotal_fill
                row_idx += 1
            ws.cell(row=row_idx, column=1, value=date_str)
            ws.cell(row=row_idx, column=1).fill = date_fill
            ws.cell(row=row_idx, column=1).font = Font(color="FFFFFF", bold=True, size=11)
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            row_idx += 1; current_date = date_val; date_totals[current_date] = 0

        ws.cell(row=row_idx, column=2, value=data["Bank Name"])
        ws.cell(row=row_idx, column=3, value=data["Card Name"])
        c = ws.cell(row=row_idx, column=4, value=data["Total"])
        c.number_format = "#,##0.00"; c.alignment = Alignment(horizontal="right")
        if data["Bank Name"] == "Unknown Bank":
            for col in range(2, 5):
                ws.cell(row=row_idx, column=col).fill = unknown_fill
                ws.cell(row=row_idx, column=col).font = Font(bold=True, color="FFFFFF")
        date_totals[current_date] += data["Total"]; row_idx += 1

    if current_date is not None:
        row_idx += 1
        ws.cell(row=row_idx, column=2, value="DATE SUBTOTAL")
        c = ws.cell(row=row_idx, column=4, value=date_totals[current_date])
        c.number_format = "#,##0.00"; c.font = Font(bold=True)
        for col in range(1, 5): ws.cell(row=row_idx, column=col).fill = subtotal_fill
        row_idx += 1

    row_idx += 1; grand_total = summary["Total"].sum()
    for col, val in {2: "GRAND TOTAL", 3: "ALL DATES", 4: grand_total}.items():
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True, size=12)
    ws.cell(row=row_idx, column=4).number_format = "#,##0.00"
    for ltr, w in zip("ABCD", [20, 20, 18, 15]): ws.column_dimensions[ltr].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, summary, len(summary["Reconciliation Date"].unique())


def create_geidea_detailed_file(df):
    """Full detailed: Bank+Card rows × Terminal columns with Debit / Credit / Total sub-cols."""
    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit": "sum", "Total Credit": "sum", "Total Debit Credit": "sum"
    }).reset_index()

    terminals    = sorted(summary["Terminal"].unique())
    banks        = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())

    rows = []
    for bank in banks:
        for card in card_schemes:
            bc = summary[(summary["Bank Name"] == bank) & (summary["Card Name"] == card)]
            if bc.empty: continue
            row = {"Bank Name": bank, "Card Scheme": card}
            for term in terminals:
                td = bc[bc["Terminal"] == term]
                row[f"{term}_Debit"]  = td["Total Debit"].values[0]         if not td.empty else 0
                row[f"{term}_Credit"] = td["Total Credit"].values[0]        if not td.empty else 0
                row[f"{term}_Total"]  = td["Total Debit Credit"].values[0]  if not td.empty else 0
            rows.append(row)

    for label in ["TOTAL", "AVG"]:
        row = {"Bank Name": label, "Card Scheme": "ALL"}
        for term in terminals:
            td = summary[summary["Terminal"] == term]
            row[f"{term}_Debit"]  = round(td["Total Debit"].sum()         if label=="TOTAL" else td["Total Debit"].mean(),        2)
            row[f"{term}_Credit"] = round(td["Total Credit"].sum()        if label=="TOTAL" else td["Total Credit"].mean(),       2)
            row[f"{term}_Total"]  = round(td["Total Debit Credit"].sum()  if label=="TOTAL" else td["Total Debit Credit"].mean(), 2)
        rows.append(row)

    wb = Workbook(); ws = wb.active; ws.title = "Detailed"
    header_fill  = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_fill     = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    for c, v in [(1, "Bank Name"), (2, "Card Scheme")]:
        cell = ws.cell(row=1, column=c, value=v)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center

    col_idx = 3
    for term in terminals:
        ws.cell(row=1, column=col_idx, value=f"#{term}").fill = header_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=9)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 2)
        for lbl, off in [("Debit", 0), ("Credit", 1), ("Total", 2)]:
            c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
            c2.fill = sub_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
        col_idx += 3

    for r_idx, row_data in enumerate(rows, 3):
        bv, cv = row_data["Bank Name"], row_data["Card Scheme"]
        for c, val in [(1, bv), (2, cv)]:
            cell = ws.cell(row=r_idx, column=c, value=val)
            if bv == "Unknown Bank":
                cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")
            elif bv in ["TOTAL", "AVG"]:
                cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid"); cell.font = Font(bold=True)
        col_idx = 3
        for term in terminals:
            for off, key in enumerate(["Debit", "Credit", "Total"]):
                cell = ws.cell(row=r_idx, column=col_idx + off, value=row_data[f"{term}_{key}"])
                cell.number_format = "#,##0.00"
            col_idx += 3

    ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 15
    for i in range(3, col_idx): ws.column_dimensions[get_column_letter(i)].width = 11
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(terminals)


def create_geidea_detailed_totals_only(df):
    """
    Simplified Geidea detailed — matches screenshot layout exactly:
    Row 1: Bank Name | Card Scheme | #TERM1 (merged 2 cols) | #TERM2 (merged 2 cols) | ... | GRAND TOTAL (merged 2 cols)
    Row 2: (blank)  | (blank)     |  Total  |  Avg.         |  Total  |  Avg.        | ... |  Total      |  Avg.
    Data rows show the terminal's value in Total, and the column average in Avg.
    Bottom: TOTAL row + AVG row.
    """
    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit Credit": "sum"
    }).reset_index()

    terminals    = sorted(summary["Terminal"].unique())
    banks        = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())

    pivot = {}
    for _, row in summary.iterrows():
        pivot[(row["Bank Name"], row["Card Name"], row["Terminal"])] = row["Total Debit Credit"]

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_TotalAvg_Only"
    header_fill  = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_fill     = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill   = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill     = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # ── Row 1 & 2: headers ───────────────────────────────────────────────────
    # Fixed label columns span both rows
    for c, v in [(1, "Bank Name"), (2, "Card Scheme")]:
        ws.cell(row=1, column=c, value=v)
        ws.cell(row=1, column=c).fill = header_fill
        ws.cell(row=1, column=c).font = Font(color="FFFFFF", bold=True, size=10)
        ws.cell(row=1, column=c).alignment = center
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)

    col_idx = 3
    for term in terminals:
        # Row 1: terminal name merged over 2 cols
        ws.cell(row=1, column=col_idx, value=f"#{term}")
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=9)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 1)
        # Row 2: Total | Avg. sub-headers
        for lbl, off in [("Total", 0), ("Avg.", 1)]:
            c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
            c2.fill = sub_fill; c2.font = Font(bold=True, size=9); c2.alignment = center
        col_idx += 2

    # Grand-total group
    grand_col = col_idx
    ws.cell(row=1, column=grand_col, value="GRAND TOTAL")
    ws.cell(row=1, column=grand_col).fill = total_fill
    ws.cell(row=1, column=grand_col).font = Font(bold=True, size=10)
    ws.cell(row=1, column=grand_col).alignment = center
    ws.merge_cells(start_row=1, start_column=grand_col, end_row=1, end_column=grand_col + 1)
    for lbl, off in [("Total", 0), ("Avg.", 1)]:
        c2 = ws.cell(row=2, column=grand_col + off, value=lbl)
        c2.fill = total_fill; c2.font = Font(bold=True, size=9); c2.alignment = center
    last_col = grand_col + 1

    # pre-compute column totals (sum across all bank/card rows) for Avg. calculation
    col_totals = {t: sum(pivot.get((b, c, t), 0) for b in banks for c in card_schemes) for t in terminals}
    n_data_rows = sum(1 for b in banks for c in card_schemes if any((b, c, t) in pivot for t in terminals))

    # ── Data rows ─────────────────────────────────────────────────────────────
    row_idx = 3
    for bank in banks:
        for card in card_schemes:
            if not any((bank, card, t) in pivot for t in terminals):
                continue
            is_unknown = bank == "Unknown Bank"

            for c, val in [(1, bank), (2, card)]:
                cell = ws.cell(row=row_idx, column=c, value=val)
                if is_unknown:
                    cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")

            grand_row_total = 0.0
            col = 3
            for term in terminals:
                val = pivot.get((bank, card, term), 0)
                # Avg. = column total / number of data rows
                avg_val = round(col_totals[term] / n_data_rows, 2) if n_data_rows else 0

                cell_t = ws.cell(row=row_idx, column=col, value=val)
                cell_t.number_format = "#,##0.00"; cell_t.alignment = right
                if is_unknown: cell_t.fill = unknown_fill

                cell_a = ws.cell(row=row_idx, column=col + 1, value=avg_val)
                cell_a.number_format = "#,##0.00"; cell_a.alignment = right
                if is_unknown: cell_a.fill = unknown_fill

                grand_row_total += val; col += 2

            n_terms = len(terminals)
            cell_gt = ws.cell(row=row_idx, column=grand_col, value=grand_row_total)
            cell_gt.number_format = "#,##0.00"; cell_gt.font = Font(bold=True); cell_gt.alignment = right
            if is_unknown: cell_gt.fill = unknown_fill

            cell_ga = ws.cell(row=row_idx, column=grand_col + 1,
                              value=round(grand_row_total / n_terms, 2) if n_terms else 0)
            cell_ga.number_format = "#,##0.00"; cell_ga.font = Font(bold=True); cell_ga.alignment = right
            if is_unknown: cell_ga.fill = unknown_fill
            row_idx += 1

    # ── TOTAL row ─────────────────────────────────────────────────────────────
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=2, value="").fill = total_fill

    grand_total = 0.0; col = 3
    for term in terminals:
        val = col_totals[term]
        avg_val = round(val / n_data_rows, 2) if n_data_rows else 0
        for off, v2 in enumerate([val, avg_val]):
            cell = ws.cell(row=row_idx, column=col + off, value=v2)
            cell.fill = total_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00"; cell.alignment = right
        grand_total += val; col += 2

    n_terms = len(terminals)
    ws.cell(row=row_idx, column=grand_col, value=grand_total).fill = total_fill
    ws.cell(row=row_idx, column=grand_col).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=grand_col).number_format = "#,##0.00"
    ws.cell(row=row_idx, column=grand_col).alignment = right
    ws.cell(row=row_idx, column=grand_col + 1, value=round(grand_total / n_terms, 2) if n_terms else 0).fill = total_fill
    ws.cell(row=row_idx, column=grand_col + 1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=grand_col + 1).number_format = "#,##0.00"
    ws.cell(row=row_idx, column=grand_col + 1).alignment = right

    # ── AVG row ───────────────────────────────────────────────────────────────
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=2, value="").fill = avg_fill

    col = 3
    for term in terminals:
        avg_val = round(col_totals[term] / n_data_rows, 2) if n_data_rows else 0
        for off in [0, 1]:
            cell = ws.cell(row=row_idx, column=col + off, value=avg_val)
            cell.fill = avg_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00"; cell.alignment = right
        col += 2

    overall_avg = round(grand_total / n_data_rows, 2) if n_data_rows else 0
    for off in [0, 1]:
        cell = ws.cell(row=row_idx, column=grand_col + off, value=overall_avg)
        cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
        cell.number_format = "#,##0.00"; cell.alignment = right

    # Column widths
    ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 16
    for i in range(3, last_col + 1): ws.column_dimensions[get_column_letter(i)].width = 13

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(terminals)


def create_geidea_detailed_by_date_file(df):
    if df["Reconciliation Date"].isna().all():
        return None, 0, 0

    summary = df.groupby(["Reconciliation Date", "Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit": "sum", "Total Credit": "sum", "Total Debit Credit": "sum"
    }).reset_index()

    dates = sorted(summary["Reconciliation Date"].unique())
    terminals = sorted(summary["Terminal"].unique())
    banks = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())

    rows = []
    for bank in banks:
        for card in card_schemes:
            bc = summary[(summary["Bank Name"] == bank) & (summary["Card Name"] == card)]
            if bc.empty: continue
            row = {"Bank Name": bank, "Card Scheme": card}
            for date in dates:
                dd = bc[bc["Reconciliation Date"] == date]
                for term in terminals:
                    td = dd[dd["Terminal"] == term]
                    row[f"{date}_{term}_Debit"]  = td["Total Debit"].values[0]         if not td.empty else 0
                    row[f"{date}_{term}_Credit"] = td["Total Credit"].values[0]        if not td.empty else 0
                    row[f"{date}_{term}_Total"]  = td["Total Debit Credit"].values[0]  if not td.empty else 0
            rows.append(row)

    for label in ["TOTAL", "AVG"]:
        row = {"Bank Name": label, "Card Scheme": "ALL"}
        for date in dates:
            dd = summary[summary["Reconciliation Date"] == date]
            for term in terminals:
                td = dd[dd["Terminal"] == term]
                row[f"{date}_{term}_Debit"]  = round(td["Total Debit"].sum()        if label=="TOTAL" else td["Total Debit"].mean(),        2)
                row[f"{date}_{term}_Credit"] = round(td["Total Credit"].sum()       if label=="TOTAL" else td["Total Credit"].mean(),       2)
                row[f"{date}_{term}_Total"]  = round(td["Total Debit Credit"].sum() if label=="TOTAL" else td["Total Debit Credit"].mean(), 2)
        rows.append(row)

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_by_Date"
    date_fill   = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_fill    = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    for c, v in [(1, "Bank Name"), (2, "Card Scheme")]:
        cell = ws.cell(row=1, column=c, value=v)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center

    col_idx = 3
    for date in dates:
        end_col = col_idx + (len(terminals) * 3) - 1
        ws.cell(row=1, column=col_idx, value=date.strftime("%A/%d/%b/%Y"))
        ws.cell(row=1, column=col_idx).fill = date_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=11)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=end_col)
        term_col = col_idx
        for term in terminals:
            ws.cell(row=2, column=term_col, value=f"#{term}").fill = header_fill
            ws.cell(row=2, column=term_col).font = Font(color="FFFFFF", bold=True, size=9)
            ws.cell(row=2, column=term_col).alignment = center
            ws.merge_cells(start_row=2, start_column=term_col, end_row=2, end_column=term_col + 2)
            for lbl, off in [("Debit", 0), ("Credit", 1), ("Total", 2)]:
                c2 = ws.cell(row=3, column=term_col + off, value=lbl)
                c2.fill = sub_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
            term_col += 3
        col_idx = end_col + 1

    for r_idx, row_data in enumerate(rows, 4):
        bv, cv = row_data["Bank Name"], row_data["Card Scheme"]
        for c, val in [(1, bv), (2, cv)]:
            cell = ws.cell(row=r_idx, column=c, value=val)
            if bv == "Unknown Bank":
                cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")
            elif bv in ["TOTAL", "AVG"]:
                cell.fill = PatternFill(start_color="E0E0E0", fill_type="solid"); cell.font = Font(bold=True)
        col_idx = 3
        for date in dates:
            for term in terminals:
                for off, key in enumerate(["Debit", "Credit", "Total"]):
                    cell = ws.cell(row=r_idx, column=col_idx + off, value=row_data[f"{date}_{term}_{key}"])
                    cell.number_format = "#,##0.00"
                col_idx += 3

    ws.column_dimensions["A"].width = 18; ws.column_dimensions["B"].width = 15
    for i in range(3, col_idx): ws.column_dimensions[get_column_letter(i)].width = 11
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(dates), len(terminals)


# ==================== FOODICS FUNCTIONS ====================

def create_foodics_summary_by_branch(df):
    summary = df.groupby(["Branch", "Payment Method"]).agg({
        "Net Amount": "sum", "Amount": "sum", "Return Amount": "sum", "Count": "sum"
    }).reset_index().sort_values(["Branch", "Net Amount"], ascending=[True, False])

    wb = Workbook(); ws = wb.active; ws.title = "Summary_by_Branch"
    header_fill   = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    branch_fill   = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    subtotal_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    total_fill    = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    for col, header in enumerate(["Branch", "Payment Method", "Net Amount", "Amount", "Returns", "Count"], 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=11); cell.alignment = center

    row_idx = 2; current_branch = None; branch_totals = {}
    for _, data in summary.iterrows():
        branch = data["Branch"]
        if branch != current_branch:
            if current_branch is not None:
                row_idx += 1; bt = branch_totals[current_branch]
                for col, val in [(2, "BRANCH SUBTOTAL"), (3, bt["net"]), (4, bt["amount"]), (5, bt["returns"]), (6, bt["count"])]:
                    ws.cell(row=row_idx, column=col, value=val)
                for col in range(1, 7):
                    ws.cell(row=row_idx, column=col).fill = subtotal_fill
                    ws.cell(row=row_idx, column=col).font = Font(bold=True)
                    if col >= 3: ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
                row_idx += 1
            ws.cell(row=row_idx, column=1, value=branch)
            ws.cell(row=row_idx, column=1).fill = branch_fill
            ws.cell(row=row_idx, column=1).font = Font(color="FFFFFF", bold=True, size=10)
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
            row_idx += 1; current_branch = branch
            branch_totals[current_branch] = {"net": 0, "amount": 0, "returns": 0, "count": 0}

        ws.cell(row=row_idx, column=2, value=data["Payment Method"])
        ws.cell(row=row_idx, column=3, value=data["Net Amount"])
        ws.cell(row=row_idx, column=4, value=data["Amount"])
        ws.cell(row=row_idx, column=5, value=data["Return Amount"])
        ws.cell(row=row_idx, column=6, value=data["Count"])
        for col in range(3, 7): ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col < 6 else "#,##0"
        branch_totals[current_branch]["net"]     += data["Net Amount"]
        branch_totals[current_branch]["amount"]  += data["Amount"]
        branch_totals[current_branch]["returns"] += data["Return Amount"]
        branch_totals[current_branch]["count"]   += data["Count"]
        row_idx += 1

    if current_branch is not None:
        row_idx += 1; bt = branch_totals[current_branch]
        for col, val in [(2, "BRANCH SUBTOTAL"), (3, bt["net"]), (4, bt["amount"]), (5, bt["returns"]), (6, bt["count"])]:
            ws.cell(row=row_idx, column=col, value=val)
        for col in range(1, 7):
            ws.cell(row=row_idx, column=col).fill = subtotal_fill
            ws.cell(row=row_idx, column=col).font = Font(bold=True)
            if col >= 3: ws.cell(row=row_idx, column=col).number_format = "#,##0.00"
        row_idx += 1

    row_idx += 1
    for col, val in [(2, "GRAND TOTAL"), (3, summary["Net Amount"].sum()), (4, summary["Amount"].sum()),
                     (5, summary["Return Amount"].sum()), (6, summary["Count"].sum())]:
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True, size=12)
        if col >= 3: cell.number_format = "#,##0.00"

    for ltr, w in zip("ABCDEF", [15, 25, 15, 15, 15, 12]): ws.column_dimensions[ltr].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, summary, len(summary["Branch"].unique())


def create_foodics_summary_by_payment_method(df):
    summary = df.groupby(["Payment Method"]).agg({
        "Net Amount": "sum", "Amount": "sum", "Return Amount": "sum", "Count": "sum"
    }).reset_index().sort_values("Net Amount", ascending=False)

    wb = Workbook(); ws = wb.active; ws.title = "Summary_by_Payment"
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    for col, header in enumerate(["Payment Method", "Net Amount", "Amount", "Returns", "Count"], 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=11); cell.alignment = center

    row_idx = 2
    for _, data in summary.iterrows():
        ws.cell(row=row_idx, column=1, value=data["Payment Method"])
        ws.cell(row=row_idx, column=2, value=data["Net Amount"])
        ws.cell(row=row_idx, column=3, value=data["Amount"])
        ws.cell(row=row_idx, column=4, value=data["Return Amount"])
        ws.cell(row=row_idx, column=5, value=data["Count"])
        for col in range(2, 6): ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col < 5 else "#,##0"
        row_idx += 1

    row_idx += 1
    for col, val in [(1, "GRAND TOTAL"), (2, summary["Net Amount"].sum()), (3, summary["Amount"].sum()),
                     (4, summary["Return Amount"].sum()), (5, summary["Count"].sum())]:
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True, size=12)
        if col >= 2: cell.number_format = "#,##0.00"

    for ltr, w in zip("ABCDE", [30, 15, 15, 15, 12]): ws.column_dimensions[ltr].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, summary


def _foodics_net_only_pivot(df, row_field, col_field, row_label, col_label,
                             header_hex, sub_hex):
    """
    Simplified Foodics pivot — Net Amount only.
    Header layout matches screenshot exactly:
      Row 1: row_label (merged 2 rows) | ColKey1 (merged 2 cols) | ColKey2 (merged 2 cols) | ... | GRAND TOTAL (merged 2 cols)
      Row 2: (blank)                   |  Total  |  Avg.          |  Total  |  Avg.         | ... |  Total      |  Avg.
    Data: each cell shows the row's actual value (Total) and the column average (Avg.).
    Bottom: TOTAL row + AVG row.
    """
    summary = df.groupby([row_field, col_field]).agg({"Net Amount": "sum"}).reset_index()
    row_keys = sorted(summary[row_field].unique())
    col_keys = sorted(summary[col_field].unique())
    pivot = {(r[row_field], r[col_field]): r["Net Amount"] for _, r in summary.iterrows()}

    # pre-compute column totals and averages
    col_totals = {ck: sum(pivot.get((rk, ck), 0) for rk in row_keys) for ck in col_keys}
    n_rows = len(row_keys); n_cols = len(col_keys)

    wb = Workbook(); ws = wb.active
    ws.title = f"Det_{col_field[:10]}_NetOnly"

    header_fill = PatternFill(start_color=header_hex, end_color=header_hex, fill_type="solid")
    sub_fill    = PatternFill(start_color=sub_hex,    end_color=sub_hex,    fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000",   end_color="FFC000",   fill_type="solid")
    avg_fill    = PatternFill(start_color="E0E0E0",   end_color="E0E0E0",   fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # ── Row 1 & 2: headers ───────────────────────────────────────────────────
    # Row-label header spans both header rows (rows 1 & 2)
    ws.cell(row=1, column=1, value=row_label)
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=10)
    ws.cell(row=1, column=1).alignment = center
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col_idx = 2
    for ck in col_keys:
        # Row 1: col name merged over 2 sub-cols
        ws.cell(row=1, column=col_idx, value=ck)
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=9)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 1)
        # Row 2: Total | Avg. sub-headers
        for lbl, off in [("Total", 0), ("Avg.", 1)]:
            c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
            c2.fill = sub_fill; c2.font = Font(bold=True, size=9); c2.alignment = center
        col_idx += 2

    # GRAND TOTAL group
    grand_col = col_idx
    ws.cell(row=1, column=grand_col, value="GRAND TOTAL")
    ws.cell(row=1, column=grand_col).fill = total_fill
    ws.cell(row=1, column=grand_col).font = Font(bold=True, size=10)
    ws.cell(row=1, column=grand_col).alignment = center
    ws.merge_cells(start_row=1, start_column=grand_col, end_row=1, end_column=grand_col + 1)
    for lbl, off in [("Total", 0), ("Avg.", 1)]:
        c2 = ws.cell(row=2, column=grand_col + off, value=lbl)
        c2.fill = total_fill; c2.font = Font(bold=True, size=9); c2.alignment = center
    last_col = grand_col + 1

    # ── Data rows ─────────────────────────────────────────────────────────────
    row_idx = 3
    for rk in row_keys:
        ws.cell(row=row_idx, column=1, value=rk)
        grand_row = 0.0; col = 2
        for ck in col_keys:
            val = pivot.get((rk, ck), 0)
            col_avg = round(col_totals[ck] / n_rows, 2) if n_rows else 0

            ws.cell(row=row_idx, column=col, value=val).number_format = "#,##0.00"
            ws.cell(row=row_idx, column=col).alignment = right
            ws.cell(row=row_idx, column=col + 1, value=col_avg).number_format = "#,##0.00"
            ws.cell(row=row_idx, column=col + 1).alignment = right

            grand_row += val; col += 2

        ws.cell(row=row_idx, column=grand_col, value=grand_row).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=grand_col).font = Font(bold=True)
        ws.cell(row=row_idx, column=grand_col).alignment = right
        ws.cell(row=row_idx, column=grand_col + 1,
                value=round(grand_row / n_cols, 2) if n_cols else 0).number_format = "#,##0.00"
        ws.cell(row=row_idx, column=grand_col + 1).font = Font(bold=True)
        ws.cell(row=row_idx, column=grand_col + 1).alignment = right
        row_idx += 1

    # ── TOTAL row ─────────────────────────────────────────────────────────────
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    grand_total = 0.0; col = 2
    for ck in col_keys:
        val = col_totals[ck]
        avg_val = round(val / n_rows, 2) if n_rows else 0
        for off, v in enumerate([val, avg_val]):
            cell = ws.cell(row=row_idx, column=col + off, value=v)
            cell.fill = total_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00"; cell.alignment = right
        grand_total += val; col += 2

    ws.cell(row=row_idx, column=grand_col, value=grand_total).fill = total_fill
    ws.cell(row=row_idx, column=grand_col).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=grand_col).number_format = "#,##0.00"
    ws.cell(row=row_idx, column=grand_col).alignment = right
    ws.cell(row=row_idx, column=grand_col + 1,
            value=round(grand_total / n_cols, 2) if n_cols else 0).fill = total_fill
    ws.cell(row=row_idx, column=grand_col + 1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=grand_col + 1).number_format = "#,##0.00"
    ws.cell(row=row_idx, column=grand_col + 1).alignment = right

    # ── AVG row ───────────────────────────────────────────────────────────────
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    col = 2
    for ck in col_keys:
        val = round(col_totals[ck] / n_rows, 2) if n_rows else 0
        for off in [0, 1]:
            cell = ws.cell(row=row_idx, column=col + off, value=val)
            cell.fill = avg_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00"; cell.alignment = right
        col += 2

    overall_avg = round(grand_total / n_rows, 2) if n_rows else 0
    for off in [0, 1]:
        cell = ws.cell(row=row_idx, column=grand_col + off, value=overall_avg)
        cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
        cell.number_format = "#,##0.00"; cell.alignment = right

    # Column widths
    ws.column_dimensions["A"].width = 30
    for i in range(2, last_col + 1): ws.column_dimensions[get_column_letter(i)].width = 14

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(col_keys), len(row_keys)


def create_foodics_detailed_by_branch(df):
    """Full: Payment Methods × Branches with Net/Amount/Returns/Count sub-cols."""
    summary = df.groupby(["Payment Method", "Branch"]).agg({
        "Net Amount": "sum", "Amount": "sum", "Return Amount": "sum", "Count": "sum"
    }).reset_index()
    payment_methods = sorted(summary["Payment Method"].unique())
    branches = sorted(summary["Branch"].unique())

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_by_Branch"
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    sub_fill    = PatternFill(start_color="A5D6A7", end_color="A5D6A7", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value="Payment Method")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=10)
    ws.cell(row=1, column=1).alignment = center

    col_idx = 2
    for branch in branches:
        ws.cell(row=1, column=col_idx, value=branch)
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=9)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)
        for lbl, off in [("Net Amount", 0), ("Amount", 1), ("Returns", 2), ("Count", 3)]:
            c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
            c2.fill = sub_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
        col_idx += 4

    ws.cell(row=1, column=col_idx, value="TOTAL").fill = total_fill
    ws.cell(row=1, column=col_idx).font = Font(bold=True, size=10)
    ws.cell(row=1, column=col_idx).alignment = center
    ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)
    for lbl, off in [("Net Amount", 0), ("Amount", 1), ("Returns", 2), ("Count", 3)]:
        c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
        c2.fill = total_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
    total_col_start = col_idx; col_idx += 4

    row_idx = 3
    for pm in payment_methods:
        ws.cell(row=row_idx, column=1, value=pm)
        pm_data = summary[summary["Payment Method"] == pm]
        c = 2; row_net = row_amt = row_ret = row_cnt = 0
        for branch in branches:
            bd = pm_data[pm_data["Branch"] == branch]
            net = bd["Net Amount"].values[0] if not bd.empty else 0
            amt = bd["Amount"].values[0]      if not bd.empty else 0
            ret = bd["Return Amount"].values[0] if not bd.empty else 0
            cnt = int(bd["Count"].values[0])  if not bd.empty else 0
            for off, val in enumerate([net, amt, ret, cnt]):
                cell = ws.cell(row=row_idx, column=c + off, value=val)
                cell.number_format = "#,##0.00" if off < 3 else "#,##0"
            row_net += net; row_amt += amt; row_ret += ret; row_cnt += cnt; c += 4
        for off, val in enumerate([row_net, row_amt, row_ret, row_cnt]):
            cell = ws.cell(row=row_idx, column=total_col_start + off, value=val)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"; cell.font = Font(bold=True)
        row_idx += 1

    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    c = 2
    for branch in branches:
        bd = summary[summary["Branch"] == branch]
        for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
            cell = ws.cell(row=row_idx, column=c + off, value=bd[cn].sum())
            cell.fill = total_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"
        c += 4
    for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
        cell = ws.cell(row=row_idx, column=total_col_start + off, value=summary[cn].sum())
        cell.fill = total_fill; cell.font = Font(bold=True, size=11)
        cell.number_format = "#,##0.00" if off < 3 else "#,##0"

    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    c = 2
    for branch in branches:
        bd = summary[summary["Branch"] == branch]
        for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
            val = round(bd[cn].mean(), 2) if not bd.empty else 0
            cell = ws.cell(row=row_idx, column=c + off, value=val)
            cell.fill = avg_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"
        c += 4
    for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
        cell = ws.cell(row=row_idx, column=total_col_start + off, value=round(summary[cn].mean(), 2))
        cell.fill = avg_fill; cell.font = Font(bold=True); cell.number_format = "#,##0.00" if off < 3 else "#,##0"

    ws.column_dimensions["A"].width = 30
    for i in range(2, col_idx): ws.column_dimensions[get_column_letter(i)].width = 13
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(branches), len(payment_methods)


def create_foodics_detailed_by_payment_method(df):
    """Full: Branches × Payment Methods with Net/Amount/Returns/Count sub-cols."""
    summary = df.groupby(["Branch", "Payment Method"]).agg({
        "Net Amount": "sum", "Amount": "sum", "Return Amount": "sum", "Count": "sum"
    }).reset_index()
    branches = sorted(summary["Branch"].unique())
    payment_methods = sorted(summary["Payment Method"].unique())

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_by_PayMethod"
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    sub_fill    = PatternFill(start_color="90CAF9", end_color="90CAF9", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value="Branch")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=10)
    ws.cell(row=1, column=1).alignment = center

    col_idx = 2
    for pm in payment_methods:
        ws.cell(row=1, column=col_idx, value=pm)
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).font = Font(color="FFFFFF", bold=True, size=9)
        ws.cell(row=1, column=col_idx).alignment = center
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)
        for lbl, off in [("Net Amount", 0), ("Amount", 1), ("Returns", 2), ("Count", 3)]:
            c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
            c2.fill = sub_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
        col_idx += 4

    ws.cell(row=1, column=col_idx, value="TOTAL").fill = total_fill
    ws.cell(row=1, column=col_idx).font = Font(bold=True, size=10)
    ws.cell(row=1, column=col_idx).alignment = center
    ws.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)
    for lbl, off in [("Net Amount", 0), ("Amount", 1), ("Returns", 2), ("Count", 3)]:
        c2 = ws.cell(row=2, column=col_idx + off, value=lbl)
        c2.fill = total_fill; c2.font = Font(bold=True, size=8); c2.alignment = center
    total_col_start = col_idx; col_idx += 4

    row_idx = 3
    for branch in branches:
        ws.cell(row=row_idx, column=1, value=branch)
        branch_data = summary[summary["Branch"] == branch]
        c = 2; row_net = row_amt = row_ret = row_cnt = 0
        for pm in payment_methods:
            pd_ = branch_data[branch_data["Payment Method"] == pm]
            net = pd_["Net Amount"].values[0]   if not pd_.empty else 0
            amt = pd_["Amount"].values[0]        if not pd_.empty else 0
            ret = pd_["Return Amount"].values[0] if not pd_.empty else 0
            cnt = int(pd_["Count"].values[0])    if not pd_.empty else 0
            for off, val in enumerate([net, amt, ret, cnt]):
                cell = ws.cell(row=row_idx, column=c + off, value=val)
                cell.number_format = "#,##0.00" if off < 3 else "#,##0"
            row_net += net; row_amt += amt; row_ret += ret; row_cnt += cnt; c += 4
        for off, val in enumerate([row_net, row_amt, row_ret, row_cnt]):
            cell = ws.cell(row=row_idx, column=total_col_start + off, value=val)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"; cell.font = Font(bold=True)
        row_idx += 1

    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    c = 2
    for pm in payment_methods:
        pd_ = summary[summary["Payment Method"] == pm]
        for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
            cell = ws.cell(row=row_idx, column=c + off, value=pd_[cn].sum())
            cell.fill = total_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"
        c += 4
    for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
        cell = ws.cell(row=row_idx, column=total_col_start + off, value=summary[cn].sum())
        cell.fill = total_fill; cell.font = Font(bold=True, size=11)
        cell.number_format = "#,##0.00" if off < 3 else "#,##0"

    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    c = 2
    for pm in payment_methods:
        pd_ = summary[summary["Payment Method"] == pm]
        for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
            val = round(pd_[cn].mean(), 2) if not pd_.empty else 0
            cell = ws.cell(row=row_idx, column=c + off, value=val)
            cell.fill = avg_fill; cell.font = Font(bold=True)
            cell.number_format = "#,##0.00" if off < 3 else "#,##0"
        c += 4
    for off, cn in enumerate(["Net Amount", "Amount", "Return Amount", "Count"]):
        cell = ws.cell(row=row_idx, column=total_col_start + off, value=round(summary[cn].mean(), 2))
        cell.fill = avg_fill; cell.font = Font(bold=True); cell.number_format = "#,##0.00" if off < 3 else "#,##0"

    ws.column_dimensions["A"].width = 15
    for i in range(2, col_idx): ws.column_dimensions[get_column_letter(i)].width = 13
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(payment_methods), len(branches)


def create_foodics_daily_avg_report(df, dates):
    if not dates: return None, None, 0
    num_days = len(dates)
    summary = df.groupby(["Payment Method"]).agg({
        "Net Amount": "sum", "Amount": "sum", "Return Amount": "sum", "Count": "sum"
    }).reset_index()
    summary["Avg Net Amount/Day"] = summary["Net Amount"] / num_days
    summary["Avg Count/Day"] = summary["Count"] / num_days
    summary = summary.sort_values("Avg Net Amount/Day", ascending=False)

    wb = Workbook(); ws = wb.active; ws.title = "Daily_Averages"
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    avg_fill    = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value=f"Report Period: {dates[0]} to {dates[-1]} ({num_days} days)")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)

    for col, header in enumerate(["Payment Method", "Total Net Amount", "Daily Avg Net",
                                   "Total Count", "Daily Avg Count", "Total Returns", "Days"], 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=11); cell.alignment = center

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
        for col in range(2, 7): ws.cell(row=row_idx, column=col).number_format = "#,##0.00" if col != 4 else "#,##0"
        row_idx += 1

    row_idx += 1
    for col, val in [(1, "GRAND TOTAL"), (2, summary["Net Amount"].sum()),
                     (3, summary["Net Amount"].sum() / num_days), (4, summary["Count"].sum()),
                     (5, summary["Count"].sum() / num_days), (6, summary["Return Amount"].sum()), (7, num_days)]:
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True, size=12)
        if col >= 2 and col != 4: cell.number_format = "#,##0.00"

    for ltr, w in zip("ABCDEFG", [30, 18, 18, 15, 18, 15, 10]): ws.column_dimensions[ltr].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, summary, num_days


# ==================== GOOGLE SHEETS PUSH ====================

SHEET_ID = "1m8beFoYLJAudtKp-aTAwM-spxekgQZTGJwUJjoJE1gU"


SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "python-scripts-463823",
    "private_key_id": "6b7ae7beeaae6a31e655233a34f89e2a81293549",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDLtHiNxJ53k36V\nUiOPQbAp6/TA4bLnuRfypMty/H8JZwznTZvtC8S4EaRrLApyt89azFpPwRcSN7gp\nXKsoOkeHry9vpgKDALtSn7BhWUteLx2WA3ajKzI2b/JouhMAB0emA1f7ZdeI5Yc+\neygmvgkOBLN575okgMv0MD6/de6dUDsDWlXJgPudT/3KPTDM3ht2CRj9rs6lbBbw\nWn4rYXI7kW5/3VMqW08JBnM/uRZ3tgOIbN0nMpwoZfWg3y3GBc2OLMjnsUtzBANZ\nZPPK1jDHbjYq1BffkLt3FfSV9pSxvhVmkCCBrPT/M/xEgZN1pku25Eyd59Cjbhi/\nobfQtv6ZAgMBAAECggEAWafeMWNa7b0shvMGdJxQOTtBV41ezQ9Ro3l1k+/ex9gj\nvUASwzOdSvh02biiBpiw+kEb9KNDEMEWXJoNOODhr63inmy+CUOOrtBa9JW1DsiE\n6IwwsKMn7/64ffB7wVTy63XoSN0rjnSbYFwbMWYNnS5jgeT7flpzqc98Jo90zKaE\nqPw2s4RkJLX8ymxFPVXiK9QBb/xxRjVbJnMY6xfOmnBm9Jx3pBb2KNewLGs/Cc26\nJtA2m6Hl7iOvMvTDTCKZCscVlR2pS/dyQX4+hrwQ/0dD/j6bZMKYQXoBiyGu0grI\nf1cZXARCvv9dxg1kVKa75EANRnSm1yNh3BL2TMNQYwKBgQDpmem9dfDC5J30ldpg\nkre710xCTKaXIEX2hmigk8od94XW0dLIL/ge19wuDwi6bqc7Y+AsOUdDQSic/oPM\ntj/xKnCkuo8q5HONq2GDLLJfYerxwOea7xyG9uCzN/9ftelun4TteDEnzn/RoElT\nWwS3EmSdzoUV6xN0lIDaPJtAIwKBgQDfPLXGfjd1oqTXzmMysRyVihDYIR9H+S1o\n8qxtaaWTlobUZjTJVsgUO4hdX8SaqsP5LYF9yZxlyQcLfo8sI8z35RXF4DN84BXG\nsoNd1ETPMmX2o/iSA7n6JVm8j+igH3Ih2litiR7xUr95vyQfXzs0gb57vgWjtEhn\nHvdu0r+UEwKBgQC3sTdTq73Ck+H95iTOEjF2/YtTC1Fov5EklXcK5oxmWjEdxut4\nTfhP0LCsa1gSula45gXu4K/AHCniomVkAeBwNU5Uyvsv4GtZeO36J5iwVqBYsLev\nZt3I57O0WpFvYu4H9lqiHgSRZ9mtLtzaNlWT3FvQmAihPrSS1QAqHMR8fwKBgFVT\nxT87m0MxicSbNLt5iy11en7CGkzOZ5cHuvSPPySskpi5AFA9BXkGUFcwdduQjhu+\nUxKbb1ZQgorYMy1x+bR/MdVSnxuKI4ixTxkcO7je0K53elmFZx7ADA7RCt+5ZUyf\nQuoB0Xv4XwvQDaSYJ+8n8IEn3sv16v7PjVAk6elVAoGBAKwOlTQc9PoDPhR0skOs\nTbr6GlSzAAgOZq+6Yf9A/ho6Oao0qwBn/chX7U4WFe1j0N7DHTJcrH5MgqSv5tGI\nQYY0IzIu2N/1uQf1cfotVRNV7hVkh1mKefUC9vrpA9SEbPPpdaXHFIWDVhOBEtJA\n606BGk8LXEsCROoU100I5xQ2\n-----END PRIVATE KEY-----\n",
    "client_email": "invoices-writer@python-scripts-463823.iam.gserviceaccount.com",
    "client_id": "114382450136588356580",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/invoices-writer%40python-scripts-463823.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}


def get_gspread_client():
    import gspread
    from google.oauth2.service_account import Credentials
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    return gspread.authorize(creds)



def _get_sheet_ws(tab_name, n_rows, n_cols):
    """Get or create worksheet, always cleared."""
    import gspread
    gc = get_gspread_client()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(tab_name)
        ws.clear()
        # resize if needed
        if ws.row_count < n_rows + 10 or ws.col_count < n_cols + 5:
            ws.resize(rows=n_rows + 10, cols=n_cols + 5)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=n_rows + 10, cols=n_cols + 5)
    return sh, ws


def _col_letter(n):
    """Convert 1-based column index to A1 letter(s)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _hex_to_rgb(hex_color):
    h = hex_color.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


def _cell_fmt(bg_hex=None, fg_hex="000000", bold=False, font_size=10,
              h_align="LEFT", number_fmt=None):
    """Build a gspread-compatible cell format dict."""
    fmt = {
        "textFormat": {
            "bold": bold,
            "fontSize": font_size,
            "foregroundColor": _hex_to_rgb(fg_hex),
        },
        "horizontalAlignment": h_align,
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "CLIP",
    }
    if bg_hex:
        fmt["backgroundColor"] = _hex_to_rgb(bg_hex)
    if number_fmt:
        fmt["numberFormat"] = {"type": "NUMBER", "pattern": number_fmt}
    return fmt


def _requests_for_pivot(ws, rows_data, col_keys, row_label,
                         header_hex, sub_hex,
                         total_hex="FFC000", avg_hex="E0E0E0",
                         label_col_count=1):
    """
    Build Sheets API batchUpdate requests to format a simplified pivot sheet.
    Layout (0-indexed rows):
      Row 0 : col-group headers (merged 2 cols each) + GRAND TOTAL group
      Row 1 : Total | Avg. sub-headers
      Row 2+ : data rows
      Last-2 : TOTAL row
      Last-1 : AVG row

    label_col_count: how many fixed label columns before data columns (1 for Foodics, 2 for Geidea + 1 date = 3 total but date is col 0)
    """
    requests = []
    sheet_id = ws.id

    n_data_cols = len(col_keys)           # number of groups
    n_pivot_cols = n_data_cols * 2 + 2   # each group=2 cols + GRAND TOTAL group=2 cols
    total_cols = label_col_count + 1 + n_pivot_cols   # +1 for Date col (col 0)
    # Actually: col 0 = Date, col 1..label_col_count = labels, then data
    # We treat first_data_col as label_col_count + 1 (after Date)
    first_data_col = label_col_count + 1   # 0-indexed

    n_data_rows = len(rows_data) - 2      # exclude TOTAL and AVG rows
    total_row_idx = 2 + n_data_rows + 1   # +1 blank gap  ... actually no blank, rows: 0=hdr1, 1=hdr2, 2..n+1=data, n+2=total, n+3=avg
    total_row_idx = 2 + n_data_rows
    avg_row_idx   = total_row_idx + 1
    last_row      = avg_row_idx + 1

    def fmt_req(r1, c1, r2, c2, fmt):
        return {"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat,wrapStrategy)"
        }}

    def merge_req(r1, c1, r2, c2):
        return {"mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL"
        }}

    def col_width_req(c1, c2, px):
        return {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": c1, "endIndex": c2},
            "properties": {"pixelSize": px},
            "fields": "pixelSize"
        }}

    def row_height_req(r1, r2, px):
        return {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": r1, "endIndex": r2},
            "properties": {"pixelSize": px},
            "fields": "pixelSize"
        }}

    # ── Column widths ──────────────────────────────────────────────────────
    requests.append(col_width_req(0, 1, 110))                          # Date
    for i in range(1, label_col_count + 1):
        requests.append(col_width_req(i, i + 1, 160))                 # label cols
    for i in range(first_data_col, first_data_col + n_pivot_cols):
        requests.append(col_width_req(i, i + 1, 110))                 # data cols

    # ── Row heights ────────────────────────────────────────────────────────
    requests.append(row_height_req(0, 1, 32))   # group header row
    requests.append(row_height_req(1, 2, 24))   # sub-header row
    requests.append(row_height_req(2, last_row, 22))  # data rows

    # ── Header row 0: label cells (merge both header rows) ────────────────
    # Date col spans rows 0-1
    requests.append(merge_req(0, 0, 2, 1))
    requests.append(fmt_req(0, 0, 2, 1, _cell_fmt(header_hex, "FFFFFF", True, 10, "CENTER")))
    # Label cols span rows 0-1
    for i in range(1, label_col_count + 1):
        requests.append(merge_req(0, i, 2, i + 1))
        requests.append(fmt_req(0, i, 2, i + 1, _cell_fmt(header_hex, "FFFFFF", True, 10, "CENTER")))

    # ── Group headers: col_key merged over 2 cols ─────────────────────────
    col = first_data_col
    for ck in col_keys:
        requests.append(merge_req(0, col, 1, col + 2))
        requests.append(fmt_req(0, col, 1, col + 2, _cell_fmt(header_hex, "FFFFFF", True, 9, "CENTER")))
        # Sub-headers row 1
        requests.append(fmt_req(1, col,     2, col + 1, _cell_fmt(sub_hex, "000000", True, 9, "CENTER")))
        requests.append(fmt_req(1, col + 1, 2, col + 2, _cell_fmt(sub_hex, "000000", True, 9, "CENTER")))
        col += 2

    # GRAND TOTAL group header
    requests.append(merge_req(0, col, 1, col + 2))
    requests.append(fmt_req(0, col, 1, col + 2, _cell_fmt(total_hex, "000000", True, 10, "CENTER")))
    requests.append(fmt_req(1, col,     2, col + 1, _cell_fmt(total_hex, "000000", True, 9, "CENTER")))
    requests.append(fmt_req(1, col + 1, 2, col + 2, _cell_fmt(total_hex, "000000", True, 9, "CENTER")))

    # ── Data rows formatting ───────────────────────────────────────────────
    num_fmt = "#,##0.00"
    # Date col
    requests.append(fmt_req(2, 0, total_row_idx, 1, _cell_fmt(None, "000000", False, 9, "CENTER")))
    # Label cols
    for i in range(1, label_col_count + 1):
        requests.append(fmt_req(2, i, total_row_idx, i + 1, _cell_fmt(None, "000000", False, 9, "LEFT")))
    # Numeric data cols
    requests.append(fmt_req(2, first_data_col, total_row_idx,
                            first_data_col + n_pivot_cols,
                            _cell_fmt(None, "000000", False, 9, "RIGHT", num_fmt)))

    # ── TOTAL row ─────────────────────────────────────────────────────────
    requests.append(fmt_req(total_row_idx, 0, total_row_idx + 1, total_cols,
                            _cell_fmt(total_hex, "000000", True, 11, "CENTER")))
    requests.append(fmt_req(total_row_idx, first_data_col, total_row_idx + 1,
                            first_data_col + n_pivot_cols,
                            _cell_fmt(total_hex, "000000", True, 11, "RIGHT", num_fmt)))

    # ── AVG row ───────────────────────────────────────────────────────────
    requests.append(fmt_req(avg_row_idx, 0, avg_row_idx + 1, total_cols,
                            _cell_fmt(avg_hex, "000000", True, 11, "CENTER")))
    requests.append(fmt_req(avg_row_idx, first_data_col, avg_row_idx + 1,
                            first_data_col + n_pivot_cols,
                            _cell_fmt(avg_hex, "000000", True, 11, "RIGHT", num_fmt)))

    # ── Freeze header rows ─────────────────────────────────────────────────
    requests.append({"updateSheetProperties": {
        "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 2, "frozenColumnCount": first_data_col}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
    }})

    # ── Outer border ──────────────────────────────────────────────────────
    requests.append({"updateBorders": {
        "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": last_row,
                  "startColumnIndex": 0, "endColumnIndex": total_cols},
        "outerBorder": {"style": "SOLID_MEDIUM", "color": _hex_to_rgb("000000")}
    }})

    return requests


def _flatten_and_format_geidea(df, tab_name):
    """Write Geidea simplified pivot to Sheets with full formatting."""
    import datetime, gspread
    today = datetime.date.today().strftime("%Y-%m-%d")

    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg(
        {"Total Debit Credit": "sum"}
    ).reset_index()
    terminals    = sorted(summary["Terminal"].unique())
    banks        = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())
    pivot = {(r["Bank Name"], r["Card Name"], r["Terminal"]): r["Total Debit Credit"]
             for _, r in summary.iterrows()}
    col_totals = {t: sum(pivot.get((b, c, t), 0) for b in banks for c in card_schemes)
                  for t in terminals}
    n_data_rows = sum(1 for b in banks for c in card_schemes
                      if any((b, c, t) in pivot for t in terminals))
    n_terms = len(terminals)

    # Build rows: [Date, Bank, Card, T1_total, T1_avg, ..., Grand_total, Grand_avg]
    header1 = ["Date", "Bank Name", "Card Scheme"]
    for t in terminals: header1 += [f"#{t}", ""]
    header1 += ["GRAND TOTAL", ""]

    header2 = ["", "", ""]
    for t in terminals: header2 += ["Total", "Avg."]
    header2 += ["Total", "Avg."]

    data_rows = []
    for bank in banks:
        for card in card_schemes:
            if not any((bank, card, t) in pivot for t in terminals): continue
            row = [today, bank, card]
            grand = 0.0
            for t in terminals:
                val = pivot.get((bank, card, t), 0)
                avg = round(col_totals[t] / n_data_rows, 2) if n_data_rows else 0
                row += [val, avg]; grand += val
            row += [grand, round(grand / n_terms, 2) if n_terms else 0]
            data_rows.append(row)

    grand_total = sum(col_totals.values())
    total_row = [today, "TOTAL", ""]
    for t in terminals:
        val = col_totals[t]
        total_row += [val, round(val / n_data_rows, 2) if n_data_rows else 0]
    total_row += [grand_total, round(grand_total / n_terms, 2) if n_terms else 0]

    avg_row = [today, "AVG", ""]
    for t in terminals:
        val = round(col_totals[t] / n_data_rows, 2) if n_data_rows else 0
        avg_row += [val, val]
    avg_row += [round(grand_total / n_data_rows, 2) if n_data_rows else 0,
                round(grand_total / (n_data_rows * n_terms), 2) if n_data_rows and n_terms else 0]

    all_rows = [header1, header2] + data_rows + [total_row, avg_row]
    n_rows = len(all_rows); n_cols = len(header1)

    sh, ws = _get_sheet_ws(tab_name, n_rows, n_cols)
    ws.update(all_rows, value_input_option="USER_ENTERED")

    reqs = _requests_for_pivot(
        ws, data_rows, terminals, "Bank Name",
        header_hex="366092", sub_hex="B8CCE4",
        label_col_count=2   # Bank Name + Card Scheme
    )
    sh.batch_update({"requests": reqs})
    return len(data_rows)


def _flatten_and_format_foodics(df, row_field, col_field, tab_name,
                                  header_hex, sub_hex):
    """Write Foodics Net-only pivot to Sheets with full formatting."""
    import datetime
    today = datetime.date.today().strftime("%Y-%m-%d")

    summary = df.groupby([row_field, col_field]).agg({"Net Amount": "sum"}).reset_index()
    row_keys = sorted(summary[row_field].unique())
    col_keys = sorted(summary[col_field].unique())
    pivot = {(r[row_field], r[col_field]): r["Net Amount"] for _, r in summary.iterrows()}
    col_totals = {ck: sum(pivot.get((rk, ck), 0) for rk in row_keys) for ck in col_keys}
    n_rows_count = len(row_keys); n_cols_count = len(col_keys)

    header1 = ["Date", row_field]
    for ck in col_keys: header1 += [ck, ""]
    header1 += ["GRAND TOTAL", ""]

    header2 = ["", ""]
    for ck in col_keys: header2 += ["Total", "Avg."]
    header2 += ["Total", "Avg."]

    data_rows = []
    for rk in row_keys:
        row = [today, rk]; grand = 0.0
        for ck in col_keys:
            val = pivot.get((rk, ck), 0)
            col_avg = round(col_totals[ck] / n_rows_count, 2) if n_rows_count else 0
            row += [val, col_avg]; grand += val
        row += [grand, round(grand / n_cols_count, 2) if n_cols_count else 0]
        data_rows.append(row)

    grand_total = sum(col_totals.values())
    total_row = [today, "TOTAL"]
    for ck in col_keys:
        val = col_totals[ck]
        total_row += [val, round(val / n_rows_count, 2) if n_rows_count else 0]
    total_row += [grand_total, round(grand_total / n_cols_count, 2) if n_cols_count else 0]

    avg_row = [today, "AVG"]
    for ck in col_keys:
        val = round(col_totals[ck] / n_rows_count, 2) if n_rows_count else 0
        avg_row += [val, val]
    avg_row += [round(grand_total / n_rows_count, 2) if n_rows_count else 0,
                round(grand_total / (n_rows_count * n_cols_count), 2) if n_rows_count and n_cols_count else 0]

    all_rows = [header1, header2] + data_rows + [total_row, avg_row]
    n_rows = len(all_rows); n_cols = len(header1)

    sh, ws = _get_sheet_ws(tab_name, n_rows, n_cols)
    ws.update(all_rows, value_input_option="USER_ENTERED")

    reqs = _requests_for_pivot(
        ws, data_rows, col_keys, row_field,
        header_hex=header_hex, sub_hex=sub_hex,
        label_col_count=1   # just the row_field label
    )
    sh.batch_update({"requests": reqs})
    return len(data_rows)


def push_rows_to_sheet(kind, df, tab_name, row_field=None, col_field=None,
                        header_hex=None, sub_hex=None):
    """Unified push: writes data + formatting to Google Sheets."""
    try:
        gc = get_gspread_client()
    except Exception as e:
        return False, f"Auth failed: {str(e)}"
    try:
        gc.open_by_key(SHEET_ID)
    except Exception as e:
        return False, (
            "Could not open sheet. Share it with:\n"
            "invoices-writer@python-scripts-463823.iam.gserviceaccount.com (Editor)\n\n"
            f"Error: {str(e)}"
        )
    try:
        if kind == "geidea":
            n = _flatten_and_format_geidea(df, tab_name)
        else:
            n = _flatten_and_format_foodics(df, row_field, col_field, tab_name, header_hex, sub_hex)
        return True, f"Pushed {n} rows to tab '{tab_name}'"
    except Exception as e:
        import traceback
        return False, f"Write error: {str(e)}\n\n{traceback.format_exc()}"



# ==================== UI ====================

st.title("🏦 Geidea & Foodics Summary Generator")
st.markdown("Upload your reconciliation file to generate summary reports")

uploaded_file = st.file_uploader(
    "📁 Upload file (Geidea or Foodics) — supports .xlsx, .xls, .csv",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file:
    try:
        df_raw = read_uploaded_file(uploaded_file)
        file_type = detect_file_type(df_raw)

        # ── GEIDEA ──────────────────────────────────────────────────────────────
        if file_type == "geidea":
            st.success(f"✅ Detected **Geidea** file: {uploaded_file.name} ({len(df_raw)} rows)")
            with st.expander("🔍 Preview Raw Data"):
                st.dataframe(df_raw.head(10), use_container_width=True)

            with st.spinner("Processing Geidea reports..."):
                df_processed = process_geidea_data(df_raw)
                summary_buffer, summary_df, grand_total = create_geidea_summary_file(df_processed)
                detailed_buffer, num_terminals          = create_geidea_detailed_file(df_processed)
                detailed_tot_buffer, _                  = create_geidea_detailed_totals_only(df_processed)

                unique_dates = df_processed["Reconciliation Date"].dropna().unique()
                has_multiple_dates = len(unique_dates) > 1
                if has_multiple_dates:
                    summary_date_buffer, _, num_dates = create_geidea_summary_by_date_file(df_processed)
                    date_buffer, _, _                 = create_geidea_detailed_by_date_file(df_processed)
                else:
                    summary_date_buffer = date_buffer = None; num_dates = 0

            st.subheader("📊 Geidea Summary Preview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Banks", summary_df["Bank Name"].nunique())
            col2.metric("Card Schemes", summary_df["Card Name"].nunique())
            col3.metric("Grand Total", f"{grand_total:,.0f}")

            if has_multiple_dates:
                st.info(f"📅 Detected {num_dates} dates: {', '.join([d.strftime('%Y-%m-%d') for d in unique_dates])}")
            if "Unknown Bank" in summary_df["Bank Name"].values:
                st.warning("⚠️ Some terminals not found in mapping (shown in red)")

            st.dataframe(
                summary_df.style.format({"Total": "{:,.2f}"})
                .apply(lambda x: ["background-color: #FF6B6B; color: white"] * 3
                       if x["Bank Name"] == "Unknown Bank" else [""] * 3, axis=1),
                use_container_width=True, height=300
            )

            st.subheader("⬇️ Geidea: Download Reports")
            n_reports = 5 if has_multiple_dates else 3
            st.markdown(f"**{n_reports} Geidea reports generated:**")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📊 Summary Totals Only", data=summary_buffer,
                    file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📋 Detailed — Full (Debit / Credit / Total)", data=detailed_buffer,
                    file_name=f"Geidea_02_DETAILED_Full_{num_terminals}_terminals.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📋 Detailed — Total & Avg. per Terminal", data=detailed_tot_buffer,
                    file_name=f"Geidea_03_DETAILED_TotalAvg_{num_terminals}_terminals.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c2:
                if has_multiple_dates:
                    st.download_button("📅 Summary by Date", data=summary_date_buffer,
                        file_name=f"Geidea_04_SUMMARY_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📆 Detailed by Date (Full)", data=date_buffer,
                        file_name=f"Geidea_05_DETAILED_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.success(f"✅ All {n_reports} Geidea reports ready! ({num_terminals} terminals)")

            # ── Add to Google Sheet ────────────────────────────────────────────
            st.markdown("---")
            st.subheader("📤 Add to Google Sheet")
            st.markdown(
                f"Push the **Simplified Detailed (Total & Avg. per Terminal)** data to "
                f"[the tracking sheet](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit).",
                unsafe_allow_html=False
            )
            import datetime as _dt
            default_tab_g = f"Geidea_{_dt.date.today().strftime('%Y-%m-%d')}"
            tab_name_g = st.text_input("Sheet tab name", value=default_tab_g, key="geidea_tab")
            if st.button("📤 Push Geidea Simplified to Sheet", key="geidea_push", type="primary", use_container_width=True):
                with st.spinner("Connecting to Google Sheets..."):
                    ok, msg = push_rows_to_sheet("geidea", df_processed, tab_name_g)
                if ok:
                    st.success(f"✅ {msg}")
                    st.markdown(f"[🔗 Open Google Sheet](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)")
                else:
                    st.error(msg)

        # ── FOODICS ─────────────────────────────────────────────────────────────
        elif file_type == "foodics":
            st.success(f"✅ Detected **Foodics** Payments Report: {uploaded_file.name}")
            with st.expander("🔍 Preview Raw Data"):
                st.dataframe(df_raw.head(15), use_container_width=True)

            with st.spinner("Processing Foodics reports..."):
                df_processed, dates = process_foodics_data(df_raw)

                branch_buffer,   branch_summary, num_branches = create_foodics_summary_by_branch(df_processed)
                payment_buffer,  payment_summary              = create_foodics_summary_by_payment_method(df_processed)
                det_br_full_buf, num_br_det, _                = create_foodics_detailed_by_branch(df_processed)
                det_pm_full_buf, num_pm_cols, _               = create_foodics_detailed_by_payment_method(df_processed)

                # Simplified: Net Amount only, each col = [Total | Avg.] sub-cols
                det_br_net_buf, _, _ = _foodics_net_only_pivot(
                    df_processed, row_field="Payment Method", col_field="Branch",
                    row_label="Payment Method", col_label="Branch",
                    header_hex="2E7D32", sub_hex="A5D6A7"
                )
                det_pm_net_buf, _, _ = _foodics_net_only_pivot(
                    df_processed, row_field="Branch", col_field="Payment Method",
                    row_label="Branch", col_label="Payment Method",
                    header_hex="1565C0", sub_hex="90CAF9"
                )

                if dates:
                    avg_buffer, _, num_days = create_foodics_daily_avg_report(df_processed, dates)
                else:
                    avg_buffer = None; num_days = 0

            st.subheader("📊 Foodics Summary Preview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Branches", num_branches)
            col2.metric("Payment Methods", payment_summary["Payment Method"].nunique())
            col3.metric("Total Net Amount", f"{payment_summary['Net Amount'].sum():,.0f}")

            if dates:
                st.info(f"📅 Report period: {dates[0]} to {dates[-1]} ({num_days} days)")
            else:
                st.info("ℹ️ No date range metadata — daily averages not available for plain CSV.")

            st.dataframe(
                payment_summary.style.format({
                    "Net Amount": "{:,.2f}", "Amount": "{:,.2f}",
                    "Return Amount": "{:,.2f}", "Count": "{:,.0f}"
                }),
                use_container_width=True, height=300
            )

            st.subheader("⬇️ Foodics: Download Reports")
            total_reports = 7 if dates else 6
            st.markdown(f"**{total_reports} Foodics reports generated:**")

            c1, c2 = st.columns(2)
            with c1:
                st.download_button("🏪 Summary by Branch", data=branch_buffer,
                    file_name=f"Foodics_01_SUMMARY_by_Branch_{num_branches}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📋 Detailed by Branch — Full (Net/Amount/Returns/Count)", data=det_br_full_buf,
                    file_name=f"Foodics_03_DETAILED_Branch_Full_{num_br_det}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📋 Detailed by Branch — Net Amount [Total | Avg.] per Branch", data=det_br_net_buf,
                    file_name=f"Foodics_05_DETAILED_Branch_NetAvg_{num_br_det}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                if dates:
                    st.download_button("📈 Daily Averages", data=avg_buffer,
                        file_name=f"Foodics_07_Daily_Averages_{num_days}_days.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c2:
                st.download_button("💳 Summary by Payment Method", data=payment_buffer,
                    file_name="Foodics_02_SUMMARY_by_Payment_Method.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📊 Detailed by Payment Method — Full (Net/Amount/Returns/Count)", data=det_pm_full_buf,
                    file_name=f"Foodics_04_DETAILED_PayMethod_Full_{num_pm_cols}_methods.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("📊 Detailed by Payment Method — Net Amount [Total | Avg.] per Method", data=det_pm_net_buf,
                    file_name=f"Foodics_06_DETAILED_PayMethod_NetAvg_{num_pm_cols}_methods.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            st.success(f"✅ {total_reports} Foodics reports ready! ({num_branches} branches · {payment_summary['Payment Method'].nunique()} payment methods)")

            # ── Add to Google Sheet ────────────────────────────────────────────
            st.markdown("---")
            st.subheader("📤 Add to Google Sheet")
            st.markdown(
                f"Push the **Simplified Detailed (Net Amount [Total | Avg.])** data to "
                f"[the tracking sheet](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit). "
                "Choose which pivot to push.",
            )
            import datetime as _dt2
            default_date_str = _dt2.date.today().strftime("%Y-%m-%d")
            push_col1, push_col2 = st.columns(2)
            with push_col1:
                tab_br = st.text_input("Tab: by Branch", value=f"Foodics_Branch_{default_date_str}", key="f_tab_br")
                if st.button("📤 Push: Net by Branch → Sheet", key="f_push_br", type="primary", use_container_width=True):
                    with st.spinner("Pushing Branch pivot..."):
                        ok, msg = push_rows_to_sheet("foodics", df_processed, tab_br,
                                                     row_field="Payment Method", col_field="Branch",
                                                     header_hex="2E7D32", sub_hex="A5D6A7")
                    if ok:
                        st.success(f"✅ {msg}")
                        st.markdown(f"[🔗 Open Sheet](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)")
                    else:
                        st.error(msg)
            with push_col2:
                tab_pm = st.text_input("Tab: by Payment Method", value=f"Foodics_PayMethod_{default_date_str}", key="f_tab_pm")
                if st.button("📤 Push: Net by Pay Method → Sheet", key="f_push_pm", type="primary", use_container_width=True):
                    with st.spinner("Pushing Payment Method pivot..."):
                        ok, msg = push_rows_to_sheet("foodics", df_processed, tab_pm,
                                                     row_field="Branch", col_field="Payment Method",
                                                     header_hex="1565C0", sub_hex="90CAF9")
                    if ok:
                        st.success(f"✅ {msg}")
                        st.markdown(f"[🔗 Open Sheet](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)")
                    else:
                        st.error(msg)

        else:
            st.error("❌ Could not detect file type.")
            st.info("**Geidea:** File must have 'Terminal' and 'Card Name' columns")
            st.info("**Foodics:** File must have 'Payment Method' and 'Branch' columns")

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("Please check your file format and try again.")

st.markdown("---")
st.caption("Geidea & Foodics Summary Generator v5.7 | Google Sheets integration · Simplified pivot [Total | Avg.] columns")
