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


# ==================== SHARED PIVOT BUILDER ====================

def _build_pivot_sheet(
    ws, pivot_data, row_keys, col_keys,
    row_label_header, col_label_header,
    value_col, value_label,
    header_fill_hex, sub_fill_hex,
    total_fill_hex="FFC000", avg_fill_hex="E0E0E0",
    unknown_key=None
):
    """
    Generic pivot builder: rows × cols with a single value column.
    Produces: header row, data rows, blank row, TOTAL row, AVG row.
    Returns the next unused col_idx (for column width sizing).
    """
    header_fill = PatternFill(start_color=header_fill_hex, end_color=header_fill_hex, fill_type="solid")
    sub_fill    = PatternFill(start_color=sub_fill_hex,    end_color=sub_fill_hex,    fill_type="solid")
    total_fill  = PatternFill(start_color=total_fill_hex,  end_color=total_fill_hex,  fill_type="solid")
    avg_fill    = PatternFill(start_color=avg_fill_hex,    end_color=avg_fill_hex,    fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # ── Row 1: row-label header + col group headers ──────────────────────────
    cell = ws.cell(row=1, column=1, value=row_label_header)
    cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=10); cell.alignment = center

    for c_idx, col_key in enumerate(col_keys):
        col = 2 + c_idx
        cell = ws.cell(row=1, column=col, value=col_key)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center

    # TOTAL column header
    total_col = 2 + len(col_keys)
    cell = ws.cell(row=1, column=total_col, value="TOTAL")
    cell.fill = total_fill; cell.font = Font(bold=True, size=10); cell.alignment = center

    # AVG column header
    avg_col = total_col + 1
    cell = ws.cell(row=1, column=avg_col, value="AVG")
    cell.fill = avg_fill; cell.font = Font(bold=True, size=10); cell.alignment = center

    # ── Data rows ─────────────────────────────────────────────────────────────
    row_idx = 2
    col_totals = {ck: 0.0 for ck in col_keys}

    for row_key in row_keys:
        is_unknown = unknown_key and row_key == unknown_key

        cell = ws.cell(row=row_idx, column=1, value=row_key)
        if is_unknown:
            cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")

        row_sum = 0.0
        for c_idx, col_key in enumerate(col_keys):
            col = 2 + c_idx
            val = pivot_data.get((row_key, col_key), 0)
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.number_format = "#,##0.00"; cell.alignment = right
            if is_unknown:
                cell.fill = unknown_fill; cell.font = Font(color="FFFFFF")
            row_sum += val
            col_totals[col_key] += val

        # Row TOTAL
        cell = ws.cell(row=row_idx, column=total_col, value=row_sum)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        if is_unknown:
            cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")

        # Row AVG (average across columns that have data)
        n_cols = len(col_keys)
        cell = ws.cell(row=row_idx, column=avg_col, value=round(row_sum / n_cols, 2) if n_cols else 0)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        if is_unknown:
            cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")

        row_idx += 1

    # ── TOTAL row ─────────────────────────────────────────────────────────────
    row_idx += 1
    cell = ws.cell(row=row_idx, column=1, value="TOTAL")
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)

    grand_total = 0.0
    for c_idx, col_key in enumerate(col_keys):
        col = 2 + c_idx
        val = col_totals[col_key]
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right
        grand_total += val

    cell = ws.cell(row=row_idx, column=total_col, value=grand_total)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    n_cols = len(col_keys)
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand_total / n_cols, 2) if n_cols else 0)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    # ── AVG row ───────────────────────────────────────────────────────────────
    row_idx += 1
    cell = ws.cell(row=row_idx, column=1, value="AVG")
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)

    n_rows = len(row_keys)
    col_avg_sum = 0.0
    for c_idx, col_key in enumerate(col_keys):
        col = 2 + c_idx
        val = round(col_totals[col_key] / n_rows, 2) if n_rows else 0
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = avg_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right
        col_avg_sum += val

    cell = ws.cell(row=row_idx, column=total_col, value=round(grand_total / n_rows, 2) if n_rows else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    cell = ws.cell(row=row_idx, column=avg_col, value=round(col_avg_sum / n_cols, 2) if n_cols else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    return avg_col + 1  # next col after last used


# ==================== GEIDEA FUNCTIONS ====================

def create_geidea_summary_file(df):
    summary = df.groupby(["Bank Name", "Card Name"]).agg({"Total": "sum"}).reset_index()
    summary["Sort"] = summary["Bank Name"].apply(lambda x: 1 if x == "Unknown Bank" else 0)
    summary = summary.sort_values(["Sort", "Bank Name", "Card Name"]).drop("Sort", axis=1)

    wb = Workbook(); ws = wb.active; ws.title = "Summary"
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

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
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    date_fill   = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
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
    """Full detailed: Bank+Card as rows, Terminals as columns with Debit/Credit/Total sub-cols."""
    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit": "sum", "Total Credit": "sum", "Total Debit Credit": "sum"
    }).reset_index()

    terminals   = sorted(summary["Terminal"].unique())
    banks       = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())

    rows = []
    for bank in banks:
        for card in card_schemes:
            bc = summary[(summary["Bank Name"] == bank) & (summary["Card Name"] == card)]
            if bc.empty: continue
            row = {"Bank Name": bank, "Card Scheme": card}
            for term in terminals:
                td = bc[bc["Terminal"] == term]
                row[f"{term}_Debit"]  = td["Total Debit"].values[0]  if not td.empty else 0
                row[f"{term}_Credit"] = td["Total Credit"].values[0] if not td.empty else 0
                row[f"{term}_Total"]  = td["Total Debit Credit"].values[0] if not td.empty else 0
            rows.append(row)

    for label in ["TOTAL", "AVG"]:
        row = {"Bank Name": label, "Card Scheme": "ALL"}
        for term in terminals:
            td = summary[summary["Terminal"] == term]
            row[f"{term}_Debit"]  = round(td["Total Debit"].sum()  if label == "TOTAL" else td["Total Debit"].mean(),  2)
            row[f"{term}_Credit"] = round(td["Total Credit"].sum() if label == "TOTAL" else td["Total Credit"].mean(), 2)
            row[f"{term}_Total"]  = round(td["Total Debit Credit"].sum() if label == "TOTAL" else td["Total Debit Credit"].mean(), 2)
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
    Simplified detailed: Bank+Card as rows, Terminals as columns — TOTAL only (no Debit/Credit).
    Includes TOTAL column + AVG column on the right.
    """
    summary = df.groupby(["Terminal", "Bank Name", "Card Name"]).agg({
        "Total Debit Credit": "sum"
    }).reset_index()

    terminals    = sorted(summary["Terminal"].unique())
    banks        = sorted(summary["Bank Name"].unique(), key=lambda x: (x == "Unknown Bank", x))
    card_schemes = sorted(summary["Card Name"].unique())

    # Build pivot dict
    pivot = {}
    for _, row in summary.iterrows():
        pivot[(row["Bank Name"], row["Card Name"], row["Terminal"])] = row["Total Debit Credit"]

    row_keys = [(bank, card) for bank in banks for card in card_schemes
                if any((bank, card, t) in pivot for t in terminals)]

    pivot_data = {(f"{b}||{c}", t): pivot.get((b, c, t), 0) for (b, c) in row_keys for t in terminals}
    row_label_keys = [f"{b}||{c}" for b, c in row_keys]

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_Totals_Only"
    header_fill  = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_fill     = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    unknown_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    total_fill   = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill     = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # Headers — 2 fixed cols then one col per terminal, then TOTAL, AVG
    for c, v in [(1, "Bank Name"), (2, "Card Scheme")]:
        cell = ws.cell(row=1, column=c, value=v)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=10); cell.alignment = center

    col_idx = 3
    for term in terminals:
        cell = ws.cell(row=1, column=col_idx, value=f"#{term}")
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center
        col_idx += 1

    total_col = col_idx
    ws.cell(row=1, column=total_col, value="TOTAL").fill = total_fill
    ws.cell(row=1, column=total_col).font = Font(bold=True, size=10)
    ws.cell(row=1, column=total_col).alignment = center
    avg_col = total_col + 1
    ws.cell(row=1, column=avg_col, value="AVG").fill = avg_fill
    ws.cell(row=1, column=avg_col).font = Font(bold=True, size=10)
    ws.cell(row=1, column=avg_col).alignment = center

    col_totals = {t: 0.0 for t in terminals}
    row_idx = 2

    for bank, card in row_keys:
        is_unknown = bank == "Unknown Bank"
        ws.cell(row=row_idx, column=1, value=bank)
        ws.cell(row=row_idx, column=2, value=card)
        if is_unknown:
            for c in [1, 2]:
                ws.cell(row=row_idx, column=c).fill = unknown_fill
                ws.cell(row=row_idx, column=c).font = Font(bold=True, color="FFFFFF")

        row_sum = 0.0
        col = 3
        for term in terminals:
            val = pivot.get((bank, card, term), 0)
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.number_format = "#,##0.00"; cell.alignment = right
            if is_unknown: cell.fill = unknown_fill; cell.font = Font(color="FFFFFF")
            row_sum += val; col_totals[term] += val; col += 1

        cell = ws.cell(row=row_idx, column=total_col, value=row_sum)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        if is_unknown: cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")

        n = len(terminals)
        cell = ws.cell(row=row_idx, column=avg_col, value=round(row_sum / n, 2) if n else 0)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        if is_unknown: cell.fill = unknown_fill; cell.font = Font(bold=True, color="FFFFFF")
        row_idx += 1

    # TOTAL row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=2, value="").fill = total_fill
    grand = 0.0; col = 3
    for term in terminals:
        val = col_totals[term]
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right
        grand += val; col += 1
    cell = ws.cell(row=row_idx, column=total_col, value=grand)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    n = len(terminals)
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / n, 2) if n else 0)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    # AVG row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    ws.cell(row=row_idx, column=2, value="").fill = avg_fill
    n_rows = len(row_keys); col = 3
    for term in terminals:
        val = round(col_totals[term] / n_rows, 2) if n_rows else 0
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = avg_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right; col += 1
    cell = ws.cell(row=row_idx, column=total_col, value=round(grand / n_rows, 2) if n_rows else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / (n_rows * n), 2) if n_rows and n else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 16
    for i in range(3, avg_col + 1): ws.column_dimensions[get_column_letter(i)].width = 13

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
                    row[f"{date}_{term}_Debit"]  = td["Total Debit"].values[0]  if not td.empty else 0
                    row[f"{date}_{term}_Credit"] = td["Total Credit"].values[0] if not td.empty else 0
                    row[f"{date}_{term}_Total"]  = td["Total Debit Credit"].values[0] if not td.empty else 0
            rows.append(row)

    for label in ["TOTAL", "AVG"]:
        row = {"Bank Name": label, "Card Scheme": "ALL"}
        for date in dates:
            dd = summary[summary["Reconciliation Date"] == date]
            for term in terminals:
                td = dd[dd["Terminal"] == term]
                row[f"{date}_{term}_Debit"]  = round(td["Total Debit"].sum()  if label=="TOTAL" else td["Total Debit"].mean(),  2)
                row[f"{date}_{term}_Credit"] = round(td["Total Credit"].sum() if label=="TOTAL" else td["Total Credit"].mean(), 2)
                row[f"{date}_{term}_Total"]  = round(td["Total Debit Credit"].sum() if label=="TOTAL" else td["Total Debit Credit"].mean(), 2)
        rows.append(row)

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_by_Date"
    date_fill   = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    sub_fill    = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    unknown_fill= PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
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
    header_fill  = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    branch_fill  = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    subtotal_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    total_fill   = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
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


def create_foodics_detailed_by_branch(df):
    """
    Full detailed: Payment Methods as rows, Branches as column groups
    with 4 sub-cols each (Net Amount / Amount / Returns / Count).
    """
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
    ws.cell(row=1, column=col_idx).font = Font(bold=True, size=10); ws.cell(row=1, column=col_idx).alignment = center
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
    """
    Full detailed: Branches as rows, Payment Methods as column groups
    with 4 sub-cols each (Net Amount / Amount / Returns / Count).
    """
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
    ws.cell(row=1, column=col_idx).font = Font(bold=True, size=10); ws.cell(row=1, column=col_idx).alignment = center
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
            net = pd_["Net Amount"].values[0]    if not pd_.empty else 0
            amt = pd_["Amount"].values[0]         if not pd_.empty else 0
            ret = pd_["Return Amount"].values[0]  if not pd_.empty else 0
            cnt = int(pd_["Count"].values[0])     if not pd_.empty else 0
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


def create_foodics_detailed_branch_net_only(df):
    """
    Simplified: Payment Methods as rows, Branches as columns — Net Amount only.
    Includes TOTAL column + AVG column on the right, TOTAL row + AVG row at bottom.
    """
    summary = df.groupby(["Payment Method", "Branch"]).agg({"Net Amount": "sum"}).reset_index()
    payment_methods = sorted(summary["Payment Method"].unique())
    branches = sorted(summary["Branch"].unique())
    pivot = {(row["Payment Method"], row["Branch"]): row["Net Amount"] for _, row in summary.iterrows()}

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_Branch_NetOnly"
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # Row 1: headers
    ws.cell(row=1, column=1, value="Payment Method")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=10)
    ws.cell(row=1, column=1).alignment = center

    col_idx = 2
    for branch in branches:
        cell = ws.cell(row=1, column=col_idx, value=branch)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center
        col_idx += 1

    total_col = col_idx
    ws.cell(row=1, column=total_col, value="TOTAL").fill = total_fill
    ws.cell(row=1, column=total_col).font = Font(bold=True, size=10); ws.cell(row=1, column=total_col).alignment = center
    avg_col = total_col + 1
    ws.cell(row=1, column=avg_col, value="AVG").fill = avg_fill
    ws.cell(row=1, column=avg_col).font = Font(bold=True, size=10); ws.cell(row=1, column=avg_col).alignment = center

    col_totals = {b: 0.0 for b in branches}
    row_idx = 2

    for pm in payment_methods:
        ws.cell(row=row_idx, column=1, value=pm)
        row_sum = 0.0; col = 2
        for branch in branches:
            val = pivot.get((pm, branch), 0)
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.number_format = "#,##0.00"; cell.alignment = right
            row_sum += val; col_totals[branch] += val; col += 1
        n = len(branches)
        cell = ws.cell(row=row_idx, column=total_col, value=row_sum)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        cell = ws.cell(row=row_idx, column=avg_col, value=round(row_sum / n, 2) if n else 0)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        row_idx += 1

    # TOTAL row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    grand = 0.0; col = 2
    for branch in branches:
        val = col_totals[branch]
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right
        grand += val; col += 1
    n = len(branches)
    cell = ws.cell(row=row_idx, column=total_col, value=grand)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / n, 2) if n else 0)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    # AVG row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    n_rows = len(payment_methods); col = 2
    for branch in branches:
        val = round(col_totals[branch] / n_rows, 2) if n_rows else 0
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = avg_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right; col += 1
    cell = ws.cell(row=row_idx, column=total_col, value=round(grand / n_rows, 2) if n_rows else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / (n_rows * n), 2) if n_rows and n else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    ws.column_dimensions["A"].width = 30
    for i in range(2, avg_col + 1): ws.column_dimensions[get_column_letter(i)].width = 13
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, len(branches), len(payment_methods)


def create_foodics_detailed_pm_net_only(df):
    """
    Simplified: Branches as rows, Payment Methods as columns — Net Amount only.
    Includes TOTAL column + AVG column on the right, TOTAL row + AVG row at bottom.
    """
    summary = df.groupby(["Branch", "Payment Method"]).agg({"Net Amount": "sum"}).reset_index()
    branches = sorted(summary["Branch"].unique())
    payment_methods = sorted(summary["Payment Method"].unique())
    pivot = {(row["Branch"], row["Payment Method"]): row["Net Amount"] for _, row in summary.iterrows()}

    wb = Workbook(); ws = wb.active; ws.title = "Detailed_PayMethod_NetOnly"
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    total_fill  = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    avg_fill    = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center")
    right  = Alignment(horizontal="right")

    # Row 1: headers
    ws.cell(row=1, column=1, value="Branch")
    ws.cell(row=1, column=1).fill = header_fill
    ws.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=10)
    ws.cell(row=1, column=1).alignment = center

    col_idx = 2
    for pm in payment_methods:
        cell = ws.cell(row=1, column=col_idx, value=pm)
        cell.fill = header_fill; cell.font = Font(color="FFFFFF", bold=True, size=9); cell.alignment = center
        col_idx += 1

    total_col = col_idx
    ws.cell(row=1, column=total_col, value="TOTAL").fill = total_fill
    ws.cell(row=1, column=total_col).font = Font(bold=True, size=10); ws.cell(row=1, column=total_col).alignment = center
    avg_col = total_col + 1
    ws.cell(row=1, column=avg_col, value="AVG").fill = avg_fill
    ws.cell(row=1, column=avg_col).font = Font(bold=True, size=10); ws.cell(row=1, column=avg_col).alignment = center

    col_totals = {pm: 0.0 for pm in payment_methods}
    row_idx = 2

    for branch in branches:
        ws.cell(row=row_idx, column=1, value=branch)
        row_sum = 0.0; col = 2
        for pm in payment_methods:
            val = pivot.get((branch, pm), 0)
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.number_format = "#,##0.00"; cell.alignment = right
            row_sum += val; col_totals[pm] += val; col += 1
        n = len(payment_methods)
        cell = ws.cell(row=row_idx, column=total_col, value=row_sum)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        cell = ws.cell(row=row_idx, column=avg_col, value=round(row_sum / n, 2) if n else 0)
        cell.number_format = "#,##0.00"; cell.font = Font(bold=True); cell.alignment = right
        row_idx += 1

    # TOTAL row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="TOTAL").fill = total_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    grand = 0.0; col = 2
    for pm in payment_methods:
        val = col_totals[pm]
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = total_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right
        grand += val; col += 1
    n = len(payment_methods)
    cell = ws.cell(row=row_idx, column=total_col, value=grand)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / n, 2) if n else 0)
    cell.fill = total_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    # AVG row
    row_idx += 1
    ws.cell(row=row_idx, column=1, value="AVG").fill = avg_fill
    ws.cell(row=row_idx, column=1).font = Font(bold=True, size=11)
    n_rows = len(branches); col = 2
    for pm in payment_methods:
        val = round(col_totals[pm] / n_rows, 2) if n_rows else 0
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.fill = avg_fill; cell.font = Font(bold=True)
        cell.number_format = "#,##0.00"; cell.alignment = right; col += 1
    cell = ws.cell(row=row_idx, column=total_col, value=round(grand / n_rows, 2) if n_rows else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right
    cell = ws.cell(row=row_idx, column=avg_col, value=round(grand / (n_rows * n), 2) if n_rows and n else 0)
    cell.fill = avg_fill; cell.font = Font(bold=True, size=11)
    cell.number_format = "#,##0.00"; cell.alignment = right

    ws.column_dimensions["A"].width = 15
    for i in range(2, avg_col + 1): ws.column_dimensions[get_column_letter(i)].width = 16
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
                summary_buffer, summary_df, grand_total    = create_geidea_summary_file(df_processed)
                detailed_buffer, num_terminals             = create_geidea_detailed_file(df_processed)
                detailed_tot_buffer, _                     = create_geidea_detailed_totals_only(df_processed)

                unique_dates = df_processed["Reconciliation Date"].dropna().unique()
                has_multiple_dates = len(unique_dates) > 1
                if has_multiple_dates:
                    summary_date_buffer, _, num_dates = create_geidea_summary_by_date_file(df_processed)
                    date_buffer, _, _                 = create_geidea_detailed_by_date_file(df_processed)
                else:
                    summary_date_buffer = date_buffer = None
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
                .apply(lambda x: ["background-color: #FF6B6B; color: white"] * 3
                       if x["Bank Name"] == "Unknown Bank" else [""] * 3, axis=1),
                use_container_width=True, height=300
            )

            st.subheader("⬇️ Geidea: Download Reports")
            if has_multiple_dates:
                st.markdown(f"**5 reports — multi-date support:**")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("📊 Summary Totals Only", data=summary_buffer,
                        file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📋 Detailed — All Columns (Debit/Credit/Total)", data=detailed_buffer,
                        file_name=f"Geidea_02_DETAILED_Full_{num_terminals}_terminals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📋 Detailed — Total & AVG Only", data=detailed_tot_buffer,
                        file_name=f"Geidea_03_DETAILED_TotalAVG_Only_{num_terminals}_terminals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with c2:
                    st.download_button("📅 Summary by Date", data=summary_date_buffer,
                        file_name=f"Geidea_04_SUMMARY_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📆 Detailed by Date (Full)", data=date_buffer,
                        file_name=f"Geidea_05_DETAILED_by_Date_{num_dates}_dates.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.success(f"✅ All 5 Geidea reports ready! ({num_dates} dates × {num_terminals} terminals)")
            else:
                st.markdown("**3 Geidea reports generated:**")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("📊 Summary Totals Only", data=summary_buffer,
                        file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    st.download_button("📋 Detailed — All Columns (Debit/Credit/Total)", data=detailed_buffer,
                        file_name=f"Geidea_02_DETAILED_Full_{num_terminals}_terminals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with c2:
                    st.download_button("📋 Detailed — Total & AVG Only", data=detailed_tot_buffer,
                        file_name=f"Geidea_03_DETAILED_TotalAVG_Only_{num_terminals}_terminals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.success(f"✅ All 3 Geidea reports ready! ({num_terminals} terminal columns)")

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
                det_br_net_buf,  _, _                         = create_foodics_detailed_branch_net_only(df_processed)
                det_pm_net_buf,  _, _                         = create_foodics_detailed_pm_net_only(df_processed)

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
                st.info("ℹ️ No date range metadata — daily averages report not available for plain CSV uploads.")

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
                st.download_button(
                    "🏪 Summary by Branch\n\nPayment methods grouped under each branch with subtotals",
                    data=branch_buffer,
                    file_name=f"Foodics_01_SUMMARY_by_Branch_{num_branches}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button(
                    "📋 Detailed by Branch — Full (Net/Amount/Returns/Count)",
                    data=det_br_full_buf,
                    file_name=f"Foodics_03_DETAILED_Branch_Full_{num_br_det}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button(
                    "📋 Detailed by Branch — Net Amount Only",
                    data=det_br_net_buf,
                    file_name=f"Foodics_05_DETAILED_Branch_NetOnly_{num_br_det}_branches.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                if dates:
                    st.download_button(
                        "📈 Daily Averages",
                        data=avg_buffer,
                        file_name=f"Foodics_07_Daily_Averages_{num_days}_days.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c2:
                st.download_button(
                    "💳 Summary by Payment Method\n\nConsolidated totals across all branches",
                    data=payment_buffer,
                    file_name="Foodics_02_SUMMARY_by_Payment_Method.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button(
                    "📊 Detailed by Payment Method — Full (Net/Amount/Returns/Count)",
                    data=det_pm_full_buf,
                    file_name=f"Foodics_04_DETAILED_PayMethod_Full_{num_pm_cols}_methods.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button(
                    "📊 Detailed by Payment Method — Net Amount Only",
                    data=det_pm_net_buf,
                    file_name=f"Foodics_06_DETAILED_PayMethod_NetOnly_{num_pm_cols}_methods.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            if dates:
                st.success(f"✅ All 7 Foodics reports ready! ({num_branches} branches · {payment_summary['Payment Method'].nunique()} payment methods)")
            else:
                st.success(f"✅ 6 Foodics reports ready! ({num_branches} branches · {payment_summary['Payment Method'].nunique()} payment methods)")

        # ── UNKNOWN ──────────────────────────────────────────────────────────────
        else:
            st.error("❌ Could not detect file type.")
            st.info("**Geidea:** File must have 'Terminal' and 'Card Name' columns")
            st.info("**Foodics:** File must have 'Payment Method' and 'Branch' columns")

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("Please check your file format and try again.")

st.markdown("---")
st.caption("Geidea & Foodics Summary Generator v5.5 | 5 Geidea reports + 7 Foodics reports")
