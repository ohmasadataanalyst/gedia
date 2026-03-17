import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io, re, datetime as _dt_ui

st.set_page_config(page_title="Geidea & Foodics Summary Generator", layout="wide")

TERMINAL_BRANCH_CODE_MAP = {
    "8189573201099250":"B01","8189572201099250":"B01","8189568401099240":"B01",
    "1551813501455580":"B02","8189580801099270":"B02","8189569601099240":"B02","18680360":"B02",
    "1551824001455540":"B03","8109111101030290":"B03","63189123":"B03",
    "1551835401455550":"B04","8189570001099240":"B04","8109110701030290":"B04",
    "1551842301455540":"B05","1551848901455540":"B05",
    "63189498":"B06","63189502":"B06",
    "63189106":"B07","63189110":"B07",
    "1554733501483710":"B08","8189582401099270":"B08",
    "1551822701455590":"B09","63189490":"B09","63933957":"B09",
    "1551838501455590":"B10","63189504":"B10",
    "1554861501483710":"B11","8189567201099240":"B11","63189494":"B11",
    "1551865401455580":"B12","63189100":"B12","63189101":"B12","63189103":"B12",
    "8189582001099270":"B13","63189105":"B13",
    "8189568001099240":"B14","63188996":"B14",
    "63189121":"B15","63189122":"B15","63189124":"B15",
    "8189583201099280":"B16","63189107":"B16",
    "63189493":"B17","63189496":"B17",
    "63189508":"B18","63189510":"B18","63189512":"B18",
    "1551874401455580":"B19","8189566401099230":"B19","63189113":"B19",
    "8189578001099260":"B22","8189574801099250":"B22",
    "63934017":"B21",
    "8189579201099270":"B26","8189577201099260":"B26",
    "8189565201099230":"B27","8189564801099230":"B27",
    "8189579601099270":"B28",
    "8189568801099240":"B30","8189566801099230":"B30","8189566001099230":"B30",
    "8189564001099230":"B31","8189563601099230":"B31",
    "8189578401099260":"B33","8189571201099250":"B33","8189570801099240":"B33","8189570401099240":"B33",
    "1551816501455580":"B34","8189581201099270":"B34","8189578801099260":"B34",
    "1551829301455570":"B35","1551843501455570":"B35","1551892301455570":"B35",
    "64729693":"B36","64729694":"B36","64729695":"B36","64729696":"B36",
    "8144089701539140":"B37","8199376901539140":"B37","8134184401539140":"B37","8145052001539140":"B37",
    "5567115001585440":"B39","5567068001585440":"B39","5566976601585440":"B39",
}

BRANCH_CODE_NAME_MAP = {
    "B01":"NURUH","B02":"KHRUH","B03":"GHRUH","B04":"NSRUH","B05":"RWRUH",
    "B06":"DARUH","B07":"LBRUH","B08":"SWRUH","B09":"AZRUH","B10":"SHRUH",
    "B11":"NRRUH","B12":"TWRUH","B13":"AQRUH","B14":"RBRUH","B15":"NDRUH",
    "B16":"BDRUH","B17":"QRRUH","B18":"TKRUH","B19":"MURUH","B21":"KRRUH",
    "B22":"OBJED","B24":"SFJED","B25":"RWAHS","B26":"HAJED","B27":"SARUH",
    "B28":"MAJED","B30":"QARUH","B31":"ANRUH","B32":"FYJED","B33":"HIRJED",
    "B34":"URRUH","B35":"IRRUH","B36":"PSJED","B37":"SHMAK","B38":"UHDMM",
    "B39":"HSRUH","B23":"SLAHS",
}

BRANCH_CODE_ORDER = [
    "B01","B02","B03","B04","B05","B06","B07","B08","B09","B10",
    "B11","B12","B13","B14","B15","B16","B17","B18","B19","B20",
    "B21","B22","B23","B24","B25","B26","B27","B28","B29","B30",
    "B31","B32","B33","B34","B35","B36","B37","B38","B39",
]

def branch_code_sort_key(name):
    for code, bname in BRANCH_CODE_NAME_MAP.items():
        if bname == name:
            try: return (BRANCH_CODE_ORDER.index(code), name)
            except ValueError: pass
    try: return (BRANCH_CODE_ORDER.index(name), name)
    except ValueError: return (9999, name)

TERMINAL_BANK_MAP = {tid: "Bank Al Bilad" for tid in TERMINAL_BRANCH_CODE_MAP}
for _t in ["63189108","63189112","63189116","63189117","63189119","63189120",
           "63189167","63189168","63189169","63189491","63189492","63189497",
           "63189499","63189503","63189506","63933955","63933956","63933958",
           "63933959","63934016","63934018","63934019","63934020","63934021",
           "63934022","63934023","63934024","63934025"]:
    TERMINAL_BANK_MAP[_t] = "Bank Al Bilad"

TERMINAL_BRANCH_LABEL_MAP = {
    tid: BRANCH_CODE_NAME_MAP.get(code, f"#{tid}")
    for tid, code in TERMINAL_BRANCH_CODE_MAP.items()
}

# ==================== FILE READING & DETECTION ====================

def read_uploaded_file(uploaded_file):
    fn = uploaded_file.name.lower()
    try:
        if fn.endswith('.xlsx'): return pd.read_excel(uploaded_file, engine='openpyxl')
        elif fn.endswith('.xls'): return pd.read_excel(uploaded_file, engine='xlrd')
        elif fn.endswith('.csv'):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep='\t')
                if df.shape[1] > 1: return df
            except Exception: pass
            uploaded_file.seek(0)
            try: return pd.read_csv(uploaded_file)
            except Exception:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, sep=';')
        else:
            try: return pd.read_excel(uploaded_file, engine='openpyxl')
            except Exception:
                uploaded_file.seek(0)
                try: return pd.read_excel(uploaded_file, engine='xlrd')
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
    except Exception: pass
    return "unknown"

def parse_foodics_date_range(date_range_str):
    try:
        dates = re.findall(r"\d{4}-\d{2}-\d{2}", str(date_range_str))
        if len(dates) >= 2:
            return pd.date_range(start=pd.to_datetime(dates[0]).date(),
                                 end=pd.to_datetime(dates[1]).date(), freq="D").date.tolist()
    except Exception: pass
    return []

# ==================== DATA PROCESSING ====================

def process_geidea_data(df):
    df = df.copy()
    df["Terminal"]  = df["Terminal"].astype(str).str.strip().str.replace(".0","",regex=False)
    df["Bank Name"] = df["Terminal"].map(TERMINAL_BANK_MAP).fillna("Unknown Bank")
    df["Total"]     = df["Ter. Total Debit"].fillna(0) + df["Ter. Total Credit"].fillna(0)
    df["Total Debit"]        = df["Ter. Total Debit"]
    df["Total Credit"]       = df["Ter. Total Credit"]
    df["Total Debit Credit"] = df["Ter.Total Debit Credit"]
    date_col = next((col for col in df.columns if "date" in col.lower() and "recon" in col.lower()), None)
    df["Reconciliation Date"] = pd.to_datetime(df[date_col]).dt.date if date_col else None
    return df

def process_foodics_data(df):
    columns_lower = [str(col).lower().strip() for col in df.columns]
    if "payment method" in columns_lower and "branch" in columns_lower:
        df_clean = df.copy()
        df_clean.columns = [str(col).strip() for col in df_clean.columns]
        date_range, dates = "", []
    else:
        data_start = 0
        for idx, row in df.iterrows():
            if "Payment Method" in str(row.values):
                data_start = idx; break
        df_clean = df.iloc[data_start:].reset_index(drop=True)
        df_clean.columns = df_clean.iloc[0]
        df_clean = df_clean[1:].reset_index(drop=True)
        df_clean.columns = [str(col).strip() for col in df_clean.columns]
        date_range = ""
        for idx, row in df.iterrows():
            if "Date Range" in str(row.values):
                date_range = row.iloc[1] if len(row) > 1 else ""; break
        dates = parse_foodics_date_range(date_range)
    df_clean["Net Amount"]    = pd.to_numeric(df_clean["Net Amount"],    errors="coerce").fillna(0)
    df_clean["Amount"]        = pd.to_numeric(df_clean["Amount"],        errors="coerce").fillna(0)
    df_clean["Return Amount"] = pd.to_numeric(df_clean["Return Amount"], errors="coerce").fillna(0)
    df_clean["Count"]         = pd.to_numeric(df_clean["Count"],         errors="coerce").fillna(0).astype(int)
    df_clean["Report Date Range"] = date_range
    df_clean["Dates"] = [dates] * len(df_clean)
    return df_clean, dates

# ==================== SHARED EXCEL HELPERS ====================

def _apply_simplified_pivot_sheet(ws, df_subset, label, header_hex, sub_hex, tab_label="Detailed_Total_Only"):
    df_work = df_subset.copy()
    df_work["Branch Label"] = df_work["Terminal"].apply(lambda t: TERMINAL_BRANCH_LABEL_MAP.get(str(t).strip(), f"#{t}"))
    summary = df_work.groupby(["Branch Label","Bank Name","Card Name"]).agg({"Total Debit Credit":"sum"}).reset_index()
    branches     = sorted(summary["Branch Label"].unique(), key=branch_code_sort_key)
    banks        = sorted(summary["Bank Name"].unique(), key=lambda x:(x=="Unknown Bank",branch_code_sort_key(x)))
    card_schemes = sorted(summary["Card Name"].unique())
    pivot = {(r["Bank Name"],r["Card Name"],r["Branch Label"]):r["Total Debit Credit"] for _,r in summary.iterrows()}
    col_totals = {b:sum(pivot.get((bk,c,b),0) for bk in banks for c in card_schemes) for b in branches}
    hf = PatternFill(start_color=header_hex,end_color=header_hex,fill_type="solid")
    sf = PatternFill(start_color=sub_hex,   end_color=sub_hex,   fill_type="solid")
    uf = PatternFill(start_color="FF6B6B",  end_color="FF6B6B",  fill_type="solid")
    tf = PatternFill(start_color="FFC000",  end_color="FFC000",  fill_type="solid")
    ct = Alignment(horizontal="center",vertical="center")
    rt = Alignment(horizontal="right")
    if label:
        lf = PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
        ws.cell(row=1,column=1,value=label); ws.cell(row=1,column=1).fill=lf
        ws.cell(row=1,column=1).font=Font(color="FFFFFF",bold=True,size=12); ws.cell(row=1,column=1).alignment=ct
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=2+len(branches)+1)
        hdr=2
    else: hdr=1
    for c,v in [(1,"Bank Name"),(2,"Card Scheme")]:
        ws.cell(row=hdr,column=c,value=v); ws.cell(row=hdr,column=c).fill=hf
        ws.cell(row=hdr,column=c).font=Font(color="FFFFFF",bold=True,size=10); ws.cell(row=hdr,column=c).alignment=ct
        ws.merge_cells(start_row=hdr,start_column=c,end_row=hdr+1,end_column=c)
    ci=3
    for b in branches:
        ws.cell(row=hdr,column=ci,value=b); ws.cell(row=hdr,column=ci).fill=hf
        ws.cell(row=hdr,column=ci).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=hdr,column=ci).alignment=ct
        c2=ws.cell(row=hdr+1,column=ci,value="Total"); c2.fill=sf; c2.font=Font(bold=True,size=9); c2.alignment=ct
        ci+=1
    gc=ci
    ws.cell(row=hdr,column=gc,value="GRAND TOTAL"); ws.cell(row=hdr,column=gc).fill=tf
    ws.cell(row=hdr,column=gc).font=Font(bold=True,size=10); ws.cell(row=hdr,column=gc).alignment=ct
    c2=ws.cell(row=hdr+1,column=gc,value="Total"); c2.fill=tf; c2.font=Font(bold=True,size=9); c2.alignment=ct
    ri=hdr+2
    for bk in banks:
        for card in card_schemes:
            if not any((bk,card,br) in pivot for br in branches): continue
            iu=bk=="Unknown Bank"
            for c,val in [(1,bk),(2,card)]:
                cell=ws.cell(row=ri,column=c,value=val)
                if iu: cell.fill=uf; cell.font=Font(bold=True,color="FFFFFF")
            gr=0.0; col=3
            for b in branches:
                val=pivot.get((bk,card,b),0)
                ct2=ws.cell(row=ri,column=col,value=val); ct2.number_format="#,##0.00"; ct2.alignment=rt
                if iu: ct2.fill=uf
                gr+=val; col+=1
            cgt=ws.cell(row=ri,column=gc,value=gr); cgt.number_format="#,##0.00"; cgt.font=Font(bold=True); cgt.alignment=rt
            if iu: cgt.fill=uf
            ri+=1
    ri+=1
    ws.cell(row=ri,column=1,value="TOTAL").fill=tf; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    ws.cell(row=ri,column=2,value="").fill=tf
    gt=0.0; col=3
    for b in branches:
        val=col_totals[b]; cell=ws.cell(row=ri,column=col,value=val)
        cell.fill=tf; cell.font=Font(bold=True); cell.number_format="#,##0.00"; cell.alignment=rt
        gt+=val; col+=1
    cgt=ws.cell(row=ri,column=gc,value=gt); cgt.fill=tf; cgt.font=Font(bold=True,size=11); cgt.number_format="#,##0.00"; cgt.alignment=rt
    ws.column_dimensions["A"].width=20; ws.column_dimensions["B"].width=16
    for i in range(3,gc+1): ws.column_dimensions[get_column_letter(i)].width=14
    return len(branches)

# ==================== GEIDEA EXCEL FUNCTIONS ====================

def create_geidea_summary_file(df):
    summary=df.groupby(["Bank Name","Card Name"]).agg({"Total":"sum"}).reset_index()
    summary["Sort"]=summary["Bank Name"].apply(lambda x:1 if x=="Unknown Bank" else 0)
    summary["BranchOrd"]=summary["Bank Name"].apply(branch_code_sort_key)
    summary=summary.sort_values(["Sort","BranchOrd","Card Name"]).drop(["Sort","BranchOrd"],axis=1)
    wb=Workbook(); ws=wb.active; ws.title="Summary"
    hf=PatternFill(start_color="366092",end_color="366092",fill_type="solid")
    uf=PatternFill(start_color="FF6B6B",end_color="FF6B6B",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    for col,header in enumerate(["Bank Name","Card Scheme","Total"],1):
        cell=ws.cell(row=1,column=col,value=header); cell.fill=hf
        cell.font=Font(color="FFFFFF",bold=True,size=12)
        cell.alignment=Alignment(horizontal="center",vertical="center")
    ri=2
    for _,data in summary.iterrows():
        ws.cell(row=ri,column=1,value=data["Bank Name"]); ws.cell(row=ri,column=2,value=data["Card Name"])
        c=ws.cell(row=ri,column=3,value=data["Total"]); c.number_format="#,##0.00"; c.alignment=Alignment(horizontal="right")
        if data["Bank Name"]=="Unknown Bank":
            for col in range(1,4): ws.cell(row=ri,column=col).fill=uf; ws.cell(row=ri,column=col).font=Font(bold=True,color="FFFFFF")
        ri+=1
    gt=summary["Total"].sum(); ri+=1
    for col,val in {1:"GRAND TOTAL",2:"ALL",3:gt}.items():
        cell=ws.cell(row=ri,column=col,value=val); cell.fill=tf; cell.font=Font(bold=True,size=12)
    ws.cell(row=ri,column=3).number_format="#,##0.00"
    ws.column_dimensions["A"].width=20; ws.column_dimensions["B"].width=18; ws.column_dimensions["C"].width=15
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,summary,gt

def create_geidea_summary_by_date_file(df):
    if df["Reconciliation Date"].isna().all(): return None,None,0
    summary=df.groupby(["Reconciliation Date","Bank Name","Card Name"]).agg({"Total":"sum"}).reset_index()
    summary["Sort"]=summary["Bank Name"].apply(lambda x:1 if x=="Unknown Bank" else 0)
    summary["BranchOrd"]=summary["Bank Name"].apply(branch_code_sort_key)
    summary=summary.sort_values(["Reconciliation Date","Sort","BranchOrd","Card Name"]).drop(["Sort","BranchOrd"],axis=1)
    wb=Workbook(); ws=wb.active; ws.title="Summary_by_Date"
    hf=PatternFill(start_color="366092",end_color="366092",fill_type="solid")
    df2=PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    uf=PatternFill(start_color="FF6B6B",end_color="FF6B6B",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    stf=PatternFill(start_color="E0E0E0",end_color="E0E0E0",fill_type="solid")
    for col,header in enumerate(["Date","Bank Name","Card Scheme","Total"],1):
        cell=ws.cell(row=1,column=col,value=header); cell.fill=hf
        cell.font=Font(color="FFFFFF",bold=True,size=12)
        cell.alignment=Alignment(horizontal="center",vertical="center")
    ri=2; cd=None; dt={}
    for _,data in summary.iterrows():
        dv=data["Reconciliation Date"]
        ds=dv.strftime("%A/%d/%b/%Y") if hasattr(dv,"strftime") else str(dv)
        if dv!=cd:
            if cd is not None:
                ri+=1; ws.cell(row=ri,column=2,value="DATE SUBTOTAL")
                c=ws.cell(row=ri,column=4,value=dt[cd]); c.number_format="#,##0.00"; c.font=Font(bold=True)
                for col in range(1,5): ws.cell(row=ri,column=col).fill=stf
                ri+=1
            ws.cell(row=ri,column=1,value=ds); ws.cell(row=ri,column=1).fill=df2
            ws.cell(row=ri,column=1).font=Font(color="FFFFFF",bold=True,size=11)
            ws.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=4)
            ri+=1; cd=dv; dt[cd]=0
        ws.cell(row=ri,column=2,value=data["Bank Name"]); ws.cell(row=ri,column=3,value=data["Card Name"])
        c=ws.cell(row=ri,column=4,value=data["Total"]); c.number_format="#,##0.00"; c.alignment=Alignment(horizontal="right")
        if data["Bank Name"]=="Unknown Bank":
            for col in range(2,5): ws.cell(row=ri,column=col).fill=uf; ws.cell(row=ri,column=col).font=Font(bold=True,color="FFFFFF")
        dt[cd]+=data["Total"]; ri+=1
    if cd is not None:
        ri+=1; ws.cell(row=ri,column=2,value="DATE SUBTOTAL")
        c=ws.cell(row=ri,column=4,value=dt[cd]); c.number_format="#,##0.00"; c.font=Font(bold=True)
        for col in range(1,5): ws.cell(row=ri,column=col).fill=stf; ri+=1
    ri+=1; gt=summary["Total"].sum()
    for col,val in {2:"GRAND TOTAL",3:"ALL DATES",4:gt}.items():
        cell=ws.cell(row=ri,column=col,value=val); cell.fill=tf; cell.font=Font(bold=True,size=12)
    ws.cell(row=ri,column=4).number_format="#,##0.00"
    for ltr,w in zip("ABCD",[20,20,18,15]): ws.column_dimensions[ltr].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,summary,len(summary["Reconciliation Date"].unique())

def create_geidea_detailed_file(df):
    summary=df.groupby(["Terminal","Bank Name","Card Name"]).agg({"Total Debit":"sum","Total Credit":"sum","Total Debit Credit":"sum"}).reset_index()
    terminals=sorted(summary["Terminal"].unique(),key=lambda t:branch_code_sort_key(TERMINAL_BRANCH_LABEL_MAP.get(t,t)))
    banks=sorted(summary["Bank Name"].unique(),key=lambda x:(x=="Unknown Bank",branch_code_sort_key(x)))
    card_schemes=sorted(summary["Card Name"].unique())
    rows=[]
    for bk in banks:
        for card in card_schemes:
            bc=summary[(summary["Bank Name"]==bk)&(summary["Card Name"]==card)]
            if bc.empty: continue
            row={"Bank Name":bk,"Card Scheme":card}
            for term in terminals:
                td=bc[bc["Terminal"]==term]
                row[f"{term}_Debit"]=td["Total Debit"].values[0] if not td.empty else 0
                row[f"{term}_Credit"]=td["Total Credit"].values[0] if not td.empty else 0
                row[f"{term}_Total"]=td["Total Debit Credit"].values[0] if not td.empty else 0
            rows.append(row)
    for label in ["TOTAL","AVG"]:
        row={"Bank Name":label,"Card Scheme":"ALL"}
        for term in terminals:
            td=summary[summary["Terminal"]==term]
            row[f"{term}_Debit"]=round(td["Total Debit"].sum() if label=="TOTAL" else td["Total Debit"].mean(),2)
            row[f"{term}_Credit"]=round(td["Total Credit"].sum() if label=="TOTAL" else td["Total Credit"].mean(),2)
            row[f"{term}_Total"]=round(td["Total Debit Credit"].sum() if label=="TOTAL" else td["Total Debit Credit"].mean(),2)
        rows.append(row)
    wb=Workbook(); ws=wb.active; ws.title="Detailed"
    hf=PatternFill(start_color="366092",end_color="366092",fill_type="solid")
    sf=PatternFill(start_color="B8CCE4",end_color="B8CCE4",fill_type="solid")
    uf=PatternFill(start_color="FF6B6B",end_color="FF6B6B",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    for c,v in [(1,"Bank Name"),(2,"Card Scheme")]:
        cell=ws.cell(row=1,column=c,value=v); cell.fill=hf; cell.font=Font(color="FFFFFF",bold=True,size=9); cell.alignment=ct
    ci=3
    for term in terminals:
        ws.cell(row=1,column=ci,value=TERMINAL_BRANCH_LABEL_MAP.get(term,f"#{term}")).fill=hf
        ws.cell(row=1,column=ci).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=1,column=ci).alignment=ct
        ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ci+2)
        for lbl,off in [("Debit",0),("Credit",1),("Total",2)]:
            c2=ws.cell(row=2,column=ci+off,value=lbl); c2.fill=sf; c2.font=Font(bold=True,size=8); c2.alignment=ct
        ci+=3
    for ri,row_data in enumerate(rows,3):
        bv,cv=row_data["Bank Name"],row_data["Card Scheme"]
        for c,val in [(1,bv),(2,cv)]:
            cell=ws.cell(row=ri,column=c,value=val)
            if bv=="Unknown Bank": cell.fill=uf; cell.font=Font(bold=True,color="FFFFFF")
            elif bv in ["TOTAL","AVG"]: cell.fill=PatternFill(start_color="E0E0E0",fill_type="solid"); cell.font=Font(bold=True)
        ci=3
        for term in terminals:
            for off,key in enumerate(["Debit","Credit","Total"]):
                cell=ws.cell(row=ri,column=ci+off,value=row_data[f"{term}_{key}"]); cell.number_format="#,##0.00"
            ci+=3
    ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=15
    for i in range(3,ci): ws.column_dimensions[get_column_letter(i)].width=11
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,len(terminals)

def create_geidea_detailed_totals_only(date_slices):
    colour_pairs=[("366092","B8CCE4"),("1F4E78","9DC3E6"),("4A235A","D2B4DE"),("1B4F2E","A9DFBF"),("784212","F5CBA7")]
    wb=Workbook(); first=True; nt=0
    for i,(label,df_slice) in enumerate(date_slices):
        if df_slice is None or len(df_slice)==0: continue
        hx,sx=colour_pairs[min(i,len(colour_pairs)-1)]
        if first: ws=wb.active; first=False
        else: ws=wb.create_sheet()
        ws.title=label
        _apply_simplified_pivot_sheet(ws,df_slice,label=label,header_hex=hx,sub_hex=sx)
        if i==0: nt=len(df_slice["Terminal"].unique()) if "Terminal" in df_slice.columns else 0
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,nt

def create_geidea_detailed_by_date_file(df):
    if df["Reconciliation Date"].isna().all(): return None,0,0
    summary=df.groupby(["Reconciliation Date","Terminal","Bank Name","Card Name"]).agg({"Total Debit":"sum","Total Credit":"sum","Total Debit Credit":"sum"}).reset_index()
    dates=sorted(summary["Reconciliation Date"].unique())
    terminals=sorted(summary["Terminal"].unique(),key=lambda t:branch_code_sort_key(TERMINAL_BRANCH_LABEL_MAP.get(t,t)))
    banks=sorted(summary["Bank Name"].unique(),key=lambda x:(x=="Unknown Bank",branch_code_sort_key(x)))
    card_schemes=sorted(summary["Card Name"].unique())
    rows=[]
    for bk in banks:
        for card in card_schemes:
            bc=summary[(summary["Bank Name"]==bk)&(summary["Card Name"]==card)]
            if bc.empty: continue
            row={"Bank Name":bk,"Card Scheme":card}
            for date in dates:
                dd=bc[bc["Reconciliation Date"]==date]
                for term in terminals:
                    td=dd[dd["Terminal"]==term]
                    row[f"{date}_{term}_Debit"]=td["Total Debit"].values[0] if not td.empty else 0
                    row[f"{date}_{term}_Credit"]=td["Total Credit"].values[0] if not td.empty else 0
                    row[f"{date}_{term}_Total"]=td["Total Debit Credit"].values[0] if not td.empty else 0
            rows.append(row)
    for label in ["TOTAL","AVG"]:
        row={"Bank Name":label,"Card Scheme":"ALL"}
        for date in dates:
            dd=summary[summary["Reconciliation Date"]==date]
            for term in terminals:
                td=dd[dd["Terminal"]==term]
                row[f"{date}_{term}_Debit"]=round(td["Total Debit"].sum() if label=="TOTAL" else td["Total Debit"].mean(),2)
                row[f"{date}_{term}_Credit"]=round(td["Total Credit"].sum() if label=="TOTAL" else td["Total Credit"].mean(),2)
                row[f"{date}_{term}_Total"]=round(td["Total Debit Credit"].sum() if label=="TOTAL" else td["Total Debit Credit"].mean(),2)
        rows.append(row)
    wb=Workbook(); ws=wb.active; ws.title="Detailed_by_Date"
    dfill=PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    hf=PatternFill(start_color="366092",end_color="366092",fill_type="solid")
    sf=PatternFill(start_color="B8CCE4",end_color="B8CCE4",fill_type="solid")
    uf=PatternFill(start_color="FF6B6B",end_color="FF6B6B",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    for c,v in [(1,"Bank Name"),(2,"Card Scheme")]:
        cell=ws.cell(row=1,column=c,value=v); cell.fill=hf; cell.font=Font(color="FFFFFF",bold=True,size=9); cell.alignment=ct
    ci=3
    for date in dates:
        ec=ci+(len(terminals)*3)-1
        ws.cell(row=1,column=ci,value=date.strftime("%A/%d/%b/%Y")); ws.cell(row=1,column=ci).fill=dfill
        ws.cell(row=1,column=ci).font=Font(color="FFFFFF",bold=True,size=11); ws.cell(row=1,column=ci).alignment=ct
        ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ec)
        tc=ci
        for term in terminals:
            ws.cell(row=2,column=tc,value=TERMINAL_BRANCH_LABEL_MAP.get(term,f"#{term}")).fill=hf
            ws.cell(row=2,column=tc).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=2,column=tc).alignment=ct
            ws.merge_cells(start_row=2,start_column=tc,end_row=2,end_column=tc+2)
            for lbl,off in [("Debit",0),("Credit",1),("Total",2)]:
                c2=ws.cell(row=3,column=tc+off,value=lbl); c2.fill=sf; c2.font=Font(bold=True,size=8); c2.alignment=ct
            tc+=3
        ci=ec+1
    for ri,row_data in enumerate(rows,4):
        bv,cv=row_data["Bank Name"],row_data["Card Scheme"]
        for c,val in [(1,bv),(2,cv)]:
            cell=ws.cell(row=ri,column=c,value=val)
            if bv=="Unknown Bank": cell.fill=uf; cell.font=Font(bold=True,color="FFFFFF")
            elif bv in ["TOTAL","AVG"]: cell.fill=PatternFill(start_color="E0E0E0",fill_type="solid"); cell.font=Font(bold=True)
        ci=3
        for date in dates:
            for term in terminals:
                for off,key in enumerate(["Debit","Credit","Total"]):
                    cell=ws.cell(row=ri,column=ci+off,value=row_data[f"{date}_{term}_{key}"]); cell.number_format="#,##0.00"
                ci+=3
    ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=15
    for i in range(3,ci): ws.column_dimensions[get_column_letter(i)].width=11
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,len(dates),len(terminals)

# ==================== FOODICS FUNCTIONS ====================

def create_foodics_summary_by_branch(df):
    summary=df.groupby(["Branch","Payment Method"]).agg({"Net Amount":"sum","Amount":"sum","Return Amount":"sum","Count":"sum"}).reset_index()
    summary["BranchOrd"]=summary["Branch"].apply(branch_code_sort_key)
    summary=summary.sort_values(["BranchOrd","Net Amount"],ascending=[True,False]).drop("BranchOrd",axis=1)
    wb=Workbook(); ws=wb.active; ws.title="Summary_by_Branch"
    hf=PatternFill(start_color="2E7D32",end_color="2E7D32",fill_type="solid")
    bf=PatternFill(start_color="4CAF50",end_color="4CAF50",fill_type="solid")
    stf=PatternFill(start_color="C8E6C9",end_color="C8E6C9",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    for col,header in enumerate(["Branch","Payment Method","Net Amount","Amount","Returns","Count"],1):
        cell=ws.cell(row=1,column=col,value=header); cell.fill=hf; cell.font=Font(color="FFFFFF",bold=True,size=11); cell.alignment=ct
    ri=2; cb=None; bt={}
    for _,data in summary.iterrows():
        branch=data["Branch"]
        if branch!=cb:
            if cb is not None:
                ri+=1; b=bt[cb]
                for col,val in [(2,"BRANCH SUBTOTAL"),(3,b["net"]),(4,b["amount"]),(5,b["returns"]),(6,b["count"])]:
                    ws.cell(row=ri,column=col,value=val)
                for col in range(1,7):
                    ws.cell(row=ri,column=col).fill=stf; ws.cell(row=ri,column=col).font=Font(bold=True)
                    if col>=3: ws.cell(row=ri,column=col).number_format="#,##0.00"
                ri+=1
            ws.cell(row=ri,column=1,value=branch); ws.cell(row=ri,column=1).fill=bf
            ws.cell(row=ri,column=1).font=Font(color="FFFFFF",bold=True,size=10)
            ws.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=6)
            ri+=1; cb=branch; bt[cb]={"net":0,"amount":0,"returns":0,"count":0}
        ws.cell(row=ri,column=2,value=data["Payment Method"])
        ws.cell(row=ri,column=3,value=data["Net Amount"]); ws.cell(row=ri,column=4,value=data["Amount"])
        ws.cell(row=ri,column=5,value=data["Return Amount"]); ws.cell(row=ri,column=6,value=data["Count"])
        for col in range(3,7): ws.cell(row=ri,column=col).number_format="#,##0.00" if col<6 else "#,##0"
        bt[cb]["net"]+=data["Net Amount"]; bt[cb]["amount"]+=data["Amount"]
        bt[cb]["returns"]+=data["Return Amount"]; bt[cb]["count"]+=data["Count"]; ri+=1
    if cb is not None:
        ri+=1; b=bt[cb]
        for col,val in [(2,"BRANCH SUBTOTAL"),(3,b["net"]),(4,b["amount"]),(5,b["returns"]),(6,b["count"])]:
            ws.cell(row=ri,column=col,value=val)
        for col in range(1,7):
            ws.cell(row=ri,column=col).fill=stf; ws.cell(row=ri,column=col).font=Font(bold=True)
            if col>=3: ws.cell(row=ri,column=col).number_format="#,##0.00"
        ri+=1
    ri+=1
    for col,val in [(2,"GRAND TOTAL"),(3,summary["Net Amount"].sum()),(4,summary["Amount"].sum()),(5,summary["Return Amount"].sum()),(6,summary["Count"].sum())]:
        cell=ws.cell(row=ri,column=col,value=val); cell.fill=tf; cell.font=Font(bold=True,size=12)
        if col>=3: cell.number_format="#,##0.00"
    for ltr,w in zip("ABCDEF",[15,25,15,15,15,12]): ws.column_dimensions[ltr].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,summary,len(summary["Branch"].unique())

def create_foodics_summary_by_payment_method(df):
    summary=df.groupby(["Payment Method"]).agg({"Net Amount":"sum","Amount":"sum","Return Amount":"sum","Count":"sum"}).reset_index().sort_values("Net Amount",ascending=False)
    wb=Workbook(); ws=wb.active; ws.title="Summary_by_Payment"
    hf=PatternFill(start_color="2E7D32",end_color="2E7D32",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    for col,header in enumerate(["Payment Method","Net Amount","Amount","Returns","Count"],1):
        cell=ws.cell(row=1,column=col,value=header); cell.fill=hf; cell.font=Font(color="FFFFFF",bold=True,size=11); cell.alignment=ct
    ri=2
    for _,data in summary.iterrows():
        ws.cell(row=ri,column=1,value=data["Payment Method"]); ws.cell(row=ri,column=2,value=data["Net Amount"])
        ws.cell(row=ri,column=3,value=data["Amount"]); ws.cell(row=ri,column=4,value=data["Return Amount"]); ws.cell(row=ri,column=5,value=data["Count"])
        for col in range(2,6): ws.cell(row=ri,column=col).number_format="#,##0.00" if col<5 else "#,##0"
        ri+=1
    ri+=1
    for col,val in [(1,"GRAND TOTAL"),(2,summary["Net Amount"].sum()),(3,summary["Amount"].sum()),(4,summary["Return Amount"].sum()),(5,summary["Count"].sum())]:
        cell=ws.cell(row=ri,column=col,value=val); cell.fill=tf; cell.font=Font(bold=True,size=12)
        if col>=2: cell.number_format="#,##0.00"
    for ltr,w in zip("ABCDE",[30,15,15,15,12]): ws.column_dimensions[ltr].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,summary

def _foodics_net_only_pivot(df,row_field,col_field,row_label,col_label,header_hex,sub_hex):
    summary=df.groupby([row_field,col_field]).agg({"Net Amount":"sum"}).reset_index()
    def _bs(vals,fn): return sorted(vals,key=branch_code_sort_key) if fn=="Branch" else sorted(vals)
    row_keys=_bs(summary[row_field].unique(),row_field); col_keys=_bs(summary[col_field].unique(),col_field)
    pivot={(r[row_field],r[col_field]):r["Net Amount"] for _,r in summary.iterrows()}
    col_totals={ck:sum(pivot.get((rk,ck),0) for rk in row_keys) for ck in col_keys}; nr=len(row_keys)
    wb=Workbook(); ws=wb.active; ws.title=f"Det_{col_field[:10]}_NetOnly"
    hf=PatternFill(start_color=header_hex,end_color=header_hex,fill_type="solid")
    sf=PatternFill(start_color=sub_hex,end_color=sub_hex,fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    af=PatternFill(start_color="E0E0E0",end_color="E0E0E0",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center"); rt=Alignment(horizontal="right")
    ws.cell(row=1,column=1,value=row_label); ws.cell(row=1,column=1).fill=hf
    ws.cell(row=1,column=1).font=Font(color="FFFFFF",bold=True,size=10); ws.cell(row=1,column=1).alignment=ct
    ws.merge_cells(start_row=1,start_column=1,end_row=2,end_column=1)
    ci=2
    for ck in col_keys:
        ws.cell(row=1,column=ci,value=ck); ws.cell(row=1,column=ci).fill=hf
        ws.cell(row=1,column=ci).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=1,column=ci).alignment=ct
        c2=ws.cell(row=2,column=ci,value="Total"); c2.fill=sf; c2.font=Font(bold=True,size=9); c2.alignment=ct; ci+=1
    gc=ci
    ws.cell(row=1,column=gc,value="GRAND TOTAL"); ws.cell(row=1,column=gc).fill=tf
    ws.cell(row=1,column=gc).font=Font(bold=True,size=10); ws.cell(row=1,column=gc).alignment=ct
    c2=ws.cell(row=2,column=gc,value="Total"); c2.fill=tf; c2.font=Font(bold=True,size=9); c2.alignment=ct
    ri=3
    for rk in row_keys:
        ws.cell(row=ri,column=1,value=rk); gr=0.0; col=2
        for ck in col_keys:
            val=pivot.get((rk,ck),0); ws.cell(row=ri,column=col,value=val).number_format="#,##0.00"
            ws.cell(row=ri,column=col).alignment=rt; gr+=val; col+=1
        ws.cell(row=ri,column=gc,value=gr).number_format="#,##0.00"
        ws.cell(row=ri,column=gc).font=Font(bold=True); ws.cell(row=ri,column=gc).alignment=rt; ri+=1
    ri+=1; ws.cell(row=ri,column=1,value="TOTAL").fill=tf; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    gt=0.0; col=2
    for ck in col_keys:
        val=col_totals[ck]; cell=ws.cell(row=ri,column=col,value=val)
        cell.fill=tf; cell.font=Font(bold=True); cell.number_format="#,##0.00"; cell.alignment=rt; gt+=val; col+=1
    cgt=ws.cell(row=ri,column=gc,value=gt); cgt.fill=tf; cgt.font=Font(bold=True,size=11); cgt.number_format="#,##0.00"; cgt.alignment=rt
    ri+=1; ws.cell(row=ri,column=1,value="AVG").fill=af; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    col=2
    for ck in col_keys:
        val=round(col_totals[ck]/nr,2) if nr else 0; cell=ws.cell(row=ri,column=col,value=val)
        cell.fill=af; cell.font=Font(bold=True); cell.number_format="#,##0.00"; cell.alignment=rt; col+=1
    oa=round(gt/nr,2) if nr else 0; cga=ws.cell(row=ri,column=gc,value=oa)
    cga.fill=af; cga.font=Font(bold=True,size=11); cga.number_format="#,##0.00"; cga.alignment=rt
    ws.column_dimensions["A"].width=30
    for i in range(2,gc+1): ws.column_dimensions[get_column_letter(i)].width=14
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,len(col_keys),len(row_keys)

def create_foodics_detailed_by_branch(df):
    summary=df.groupby(["Payment Method","Branch"]).agg({"Net Amount":"sum","Amount":"sum","Return Amount":"sum","Count":"sum"}).reset_index()
    pms=sorted(summary["Payment Method"].unique()); brs=sorted(summary["Branch"].unique(),key=branch_code_sort_key)
    wb=Workbook(); ws=wb.active; ws.title="Detailed_by_Branch"
    hf=PatternFill(start_color="2E7D32",end_color="2E7D32",fill_type="solid")
    sf=PatternFill(start_color="A5D6A7",end_color="A5D6A7",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    af=PatternFill(start_color="E0E0E0",end_color="E0E0E0",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    ws.cell(row=1,column=1,value="Payment Method"); ws.cell(row=1,column=1).fill=hf
    ws.cell(row=1,column=1).font=Font(color="FFFFFF",bold=True,size=10); ws.cell(row=1,column=1).alignment=ct
    ci=2
    for branch in brs:
        ws.cell(row=1,column=ci,value=branch); ws.cell(row=1,column=ci).fill=hf
        ws.cell(row=1,column=ci).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=1,column=ci).alignment=ct
        ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ci+3)
        for lbl,off in [("Net Amount",0),("Amount",1),("Returns",2),("Count",3)]:
            c2=ws.cell(row=2,column=ci+off,value=lbl); c2.fill=sf; c2.font=Font(bold=True,size=8); c2.alignment=ct
        ci+=4
    ws.cell(row=1,column=ci,value="TOTAL").fill=tf; ws.cell(row=1,column=ci).font=Font(bold=True,size=10); ws.cell(row=1,column=ci).alignment=ct
    ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ci+3)
    for lbl,off in [("Net Amount",0),("Amount",1),("Returns",2),("Count",3)]:
        c2=ws.cell(row=2,column=ci+off,value=lbl); c2.fill=tf; c2.font=Font(bold=True,size=8); c2.alignment=ct
    tcs=ci; ci+=4
    ri=3
    for pm in pms:
        ws.cell(row=ri,column=1,value=pm); pmd=summary[summary["Payment Method"]==pm]
        c=2; rn=ra=rr=rc=0
        for branch in brs:
            bd=pmd[pmd["Branch"]==branch]
            net=bd["Net Amount"].values[0] if not bd.empty else 0; amt=bd["Amount"].values[0] if not bd.empty else 0
            ret=bd["Return Amount"].values[0] if not bd.empty else 0; cnt=int(bd["Count"].values[0]) if not bd.empty else 0
            for off,val in enumerate([net,amt,ret,cnt]):
                cell=ws.cell(row=ri,column=c+off,value=val); cell.number_format="#,##0.00" if off<3 else "#,##0"
            rn+=net; ra+=amt; rr+=ret; rc+=cnt; c+=4
        for off,val in enumerate([rn,ra,rr,rc]):
            cell=ws.cell(row=ri,column=tcs+off,value=val); cell.number_format="#,##0.00" if off<3 else "#,##0"; cell.font=Font(bold=True)
        ri+=1
    ri+=1; ws.cell(row=ri,column=1,value="TOTAL").fill=tf; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    c=2
    for branch in brs:
        bd=summary[summary["Branch"]==branch]
        for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
            cell=ws.cell(row=ri,column=c+off,value=bd[cn].sum()); cell.fill=tf; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
        c+=4
    for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
        cell=ws.cell(row=ri,column=tcs+off,value=summary[cn].sum()); cell.fill=tf; cell.font=Font(bold=True,size=11); cell.number_format="#,##0.00" if off<3 else "#,##0"
    ri+=1; ws.cell(row=ri,column=1,value="AVG").fill=af; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    c=2
    for branch in brs:
        bd=summary[summary["Branch"]==branch]
        for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
            val=round(bd[cn].mean(),2) if not bd.empty else 0; cell=ws.cell(row=ri,column=c+off,value=val)
            cell.fill=af; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
        c+=4
    for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
        cell=ws.cell(row=ri,column=tcs+off,value=round(summary[cn].mean(),2)); cell.fill=af; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
    ws.column_dimensions["A"].width=30
    for i in range(2,ci): ws.column_dimensions[get_column_letter(i)].width=13
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,len(brs),len(pms)

def create_foodics_detailed_by_payment_method(df):
    summary=df.groupby(["Branch","Payment Method"]).agg({"Net Amount":"sum","Amount":"sum","Return Amount":"sum","Count":"sum"}).reset_index()
    brs=sorted(summary["Branch"].unique(),key=branch_code_sort_key); pms=sorted(summary["Payment Method"].unique())
    wb=Workbook(); ws=wb.active; ws.title="Detailed_by_PayMethod"
    hf=PatternFill(start_color="1565C0",end_color="1565C0",fill_type="solid")
    sf=PatternFill(start_color="90CAF9",end_color="90CAF9",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    af=PatternFill(start_color="E0E0E0",end_color="E0E0E0",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    ws.cell(row=1,column=1,value="Branch"); ws.cell(row=1,column=1).fill=hf
    ws.cell(row=1,column=1).font=Font(color="FFFFFF",bold=True,size=10); ws.cell(row=1,column=1).alignment=ct
    ci=2
    for pm in pms:
        ws.cell(row=1,column=ci,value=pm); ws.cell(row=1,column=ci).fill=hf
        ws.cell(row=1,column=ci).font=Font(color="FFFFFF",bold=True,size=9); ws.cell(row=1,column=ci).alignment=ct
        ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ci+3)
        for lbl,off in [("Net Amount",0),("Amount",1),("Returns",2),("Count",3)]:
            c2=ws.cell(row=2,column=ci+off,value=lbl); c2.fill=sf; c2.font=Font(bold=True,size=8); c2.alignment=ct
        ci+=4
    ws.cell(row=1,column=ci,value="TOTAL").fill=tf; ws.cell(row=1,column=ci).font=Font(bold=True,size=10); ws.cell(row=1,column=ci).alignment=ct
    ws.merge_cells(start_row=1,start_column=ci,end_row=1,end_column=ci+3)
    for lbl,off in [("Net Amount",0),("Amount",1),("Returns",2),("Count",3)]:
        c2=ws.cell(row=2,column=ci+off,value=lbl); c2.fill=tf; c2.font=Font(bold=True,size=8); c2.alignment=ct
    tcs=ci; ci+=4
    ri=3
    for branch in brs:
        ws.cell(row=ri,column=1,value=branch); bd=summary[summary["Branch"]==branch]
        c=2; rn=ra=rr=rc=0
        for pm in pms:
            pd_=bd[bd["Payment Method"]==pm]
            net=pd_["Net Amount"].values[0] if not pd_.empty else 0; amt=pd_["Amount"].values[0] if not pd_.empty else 0
            ret=pd_["Return Amount"].values[0] if not pd_.empty else 0; cnt=int(pd_["Count"].values[0]) if not pd_.empty else 0
            for off,val in enumerate([net,amt,ret,cnt]):
                cell=ws.cell(row=ri,column=c+off,value=val); cell.number_format="#,##0.00" if off<3 else "#,##0"
            rn+=net; ra+=amt; rr+=ret; rc+=cnt; c+=4
        for off,val in enumerate([rn,ra,rr,rc]):
            cell=ws.cell(row=ri,column=tcs+off,value=val); cell.number_format="#,##0.00" if off<3 else "#,##0"; cell.font=Font(bold=True)
        ri+=1
    ri+=1; ws.cell(row=ri,column=1,value="TOTAL").fill=tf; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    c=2
    for pm in pms:
        pd_=summary[summary["Payment Method"]==pm]
        for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
            cell=ws.cell(row=ri,column=c+off,value=pd_[cn].sum()); cell.fill=tf; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
        c+=4
    for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
        cell=ws.cell(row=ri,column=tcs+off,value=summary[cn].sum()); cell.fill=tf; cell.font=Font(bold=True,size=11); cell.number_format="#,##0.00" if off<3 else "#,##0"
    ri+=1; ws.cell(row=ri,column=1,value="AVG").fill=af; ws.cell(row=ri,column=1).font=Font(bold=True,size=11)
    c=2
    for pm in pms:
        pd_=summary[summary["Payment Method"]==pm]
        for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
            val=round(pd_[cn].mean(),2) if not pd_.empty else 0; cell=ws.cell(row=ri,column=c+off,value=val)
            cell.fill=af; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
        c+=4
    for off,cn in enumerate(["Net Amount","Amount","Return Amount","Count"]):
        cell=ws.cell(row=ri,column=tcs+off,value=round(summary[cn].mean(),2)); cell.fill=af; cell.font=Font(bold=True); cell.number_format="#,##0.00" if off<3 else "#,##0"
    ws.column_dimensions["A"].width=15
    for i in range(2,ci): ws.column_dimensions[get_column_letter(i)].width=13
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,len(pms),len(brs)

def create_foodics_daily_avg_report(df,dates):
    if not dates: return None,None,0
    nd=len(dates)
    summary=df.groupby(["Payment Method"]).agg({"Net Amount":"sum","Amount":"sum","Return Amount":"sum","Count":"sum"}).reset_index()
    summary["Avg Net Amount/Day"]=summary["Net Amount"]/nd; summary["Avg Count/Day"]=summary["Count"]/nd
    summary=summary.sort_values("Avg Net Amount/Day",ascending=False)
    wb=Workbook(); ws=wb.active; ws.title="Daily_Averages"
    hf=PatternFill(start_color="2E7D32",end_color="2E7D32",fill_type="solid")
    avf=PatternFill(start_color="E8F5E9",end_color="E8F5E9",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center")
    ws.cell(row=1,column=1,value=f"Report Period: {dates[0]} to {dates[-1]} ({nd} days)"); ws.cell(row=1,column=1).font=Font(bold=True,size=12)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=7)
    for col,header in enumerate(["Payment Method","Total Net Amount","Daily Avg Net","Total Count","Daily Avg Count","Total Returns","Days"],1):
        cell=ws.cell(row=2,column=col,value=header); cell.fill=hf; cell.font=Font(color="FFFFFF",bold=True,size=11); cell.alignment=ct
    ri=3
    for _,data in summary.iterrows():
        ws.cell(row=ri,column=1,value=data["Payment Method"]); ws.cell(row=ri,column=2,value=data["Net Amount"])
        ws.cell(row=ri,column=3,value=data["Avg Net Amount/Day"]); ws.cell(row=ri,column=4,value=data["Count"])
        ws.cell(row=ri,column=5,value=data["Avg Count/Day"]); ws.cell(row=ri,column=6,value=data["Return Amount"]); ws.cell(row=ri,column=7,value=nd)
        ws.cell(row=ri,column=3).fill=avf; ws.cell(row=ri,column=5).fill=avf
        for col in range(2,7): ws.cell(row=ri,column=col).number_format="#,##0.00" if col!=4 else "#,##0"
        ri+=1
    ri+=1
    for col,val in [(1,"GRAND TOTAL"),(2,summary["Net Amount"].sum()),(3,summary["Net Amount"].sum()/nd),(4,summary["Count"].sum()),(5,summary["Count"].sum()/nd),(6,summary["Return Amount"].sum()),(7,nd)]:
        cell=ws.cell(row=ri,column=col,value=val); cell.fill=tf; cell.font=Font(bold=True,size=12)
        if col>=2 and col!=4: cell.number_format="#,##0.00"
    for ltr,w in zip("ABCDEFG",[30,18,18,15,18,15,10]): ws.column_dimensions[ltr].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf,summary,nd

# ==================== ZENPUT ====================

ZENPUT_API_KEY="8051c8104fd221694d9aeb305f7f4abb"
ZENPUT_FINANCIAL_TID=1556454
ZENPUT_SHEET_ID="1QmcxxyEyDJTyaWp9zg5shcVHOLZPH7ibgr5STFT1ZhY"

ZENPUT_BRANCH_MAP={
    "2197299":"LBRUH B07","2239240":"FYJED B32","2235670":"ANRUH B31","2190657":"SLAHS B23",
    "2164026":"NDRUH B15","2164019":"SWRUH B08","2203271":"SARUH B27","2164017":"DARUH B06",
    "2164032":"KRRUH B21","2164031":"SFJED B24","2164025":"RBRUH B14","2164016":"RWRUH B05",
    "2197297":"NSRUH B04","2164021":"SHRUH B10","2164013":"KHRUH B02","2155652":"NURUH B01",
    "2164023":"TWRUH B12","2164020":"AZRUH B09","2199002":"RWAHS B25","2242934":"HIRJED B33",
    "2164022":"NRRUH B11","2164030":"MURUH B19","2164014":"GHRUH B03","2211854":"QARUH B30",
    "2258220":"PSJED B36","2185452":"OBJED B22","2243963":"URRUH B34","2199835":"HAJED B26",
    "2210205":"MAJED B28","2250799":"IRRUH B35","2164027":"BDRUH B16","2155654":"AQRUH B13",
    "2197298":"TKRUH B18","2164028":"QRRUH B17","2257790":"SHMAK B37","2260889":"UHDMM B38","2263062":"HSRUH B39",
}
ZENPUT_CHANNELS=["Noon - نون","To you - تو يو","Barakah - بركه","Mr. Manddob - مستر مندوب","Ninja - نينجا","The chefz - ذا شيفز","Marsool - مرسول","Solo loyalty - سولو لوياليتي","Jahez - جاهز","Hungerstation - هنقرستيشن","Ketta - كيتا","Mada - مدى","Cash - كاش","Cash used without invoice - الكاش المستخدم من غير فاتورة","Cash purchase invoice - فواتير الشراء النقدية"]
ZENPUT_CHANNEL_SHORT={"Noon - نون":"Noon","To you - تو يو":"To You","Barakah - بركه":"Barakah","Mr. Manddob - مستر مندوب":"Mr. Manddob","Ninja - نينجا":"Ninja","The chefz - ذا شيفز":"The Chefz","Marsool - مرسول":"Marsool","Solo loyalty - سولو لوياليتي":"Solo Loyalty","Jahez - جاهز":"Jahez","Hungerstation - هنقرستيشن":"Hungerstation","Ketta - كيتا":"Ketta","Mada - مدى":"Mada","Cash - كاش":"Cash","Cash used without invoice - الكاش المستخدم من غير فاتورة":"Cash With No Invoices","Cash purchase invoice - فواتير الشراء النقدية":"Cash Purchase Inv."}

def zenput_fetch_financial(start_date_str,end_date_str):
    import requests,pytz
    from datetime import datetime
    TZ=pytz.timezone("Asia/Baghdad")
    start_dt=TZ.localize(datetime.strptime(start_date_str,"%Y-%m-%d").replace(hour=0,minute=0,second=0))
    end_dt=TZ.localize(datetime.strptime(end_date_str,"%Y-%m-%d").replace(hour=23,minute=59,second=59))
    headers={"X-API-TOKEN":ZENPUT_API_KEY,"Content-Type":"application/json"}
    all_subs=[]; offset,limit=0,100
    while True:
        resp=requests.get("https://www.zenput.com/api/v3/submissions/",headers=headers,params={"form_template_id":ZENPUT_FINANCIAL_TID,"limit":limit,"start":offset,"date_submitted_start":start_date_str},timeout=30)
        if resp.status_code!=200: raise Exception(f"Zenput API error {resp.status_code}: {resp.text[:300]}")
        batch=resp.json().get("data",[])
        if not batch: break
        for s in batch:
            raw_date=s.get("smetadata",{}).get("date_submitted_local","")
            if not raw_date: continue
            try:
                sub_dt=datetime.fromisoformat(raw_date)
                if sub_dt.tzinfo is None: sub_dt=TZ.localize(sub_dt)
                if start_dt<=sub_dt<=end_dt: all_subs.append(s)
            except Exception: pass
        offset+=limit
        last_raw=batch[-1].get("smetadata",{}).get("date_submitted_local","")
        if last_raw:
            try:
                if datetime.fromisoformat(last_raw).replace(tzinfo=None)<datetime.strptime(start_date_str,"%Y-%m-%d"): break
            except Exception: pass
    if not all_subs: return pd.DataFrame()
    rows=[]
    for s in all_subs:
        raw_date=s.get("smetadata",{}).get("date_submitted_local",""); date_str=raw_date[:10] if raw_date else ""
        answers=s.get("answers",[]); row={"Date":date_str,"Branch":""}
        for ans in answers:
            t=ans.get("title","")
            if "Store" in t or "الفرع" in t:
                code=str(ans.get("value","")).strip(); row["Branch"]=ZENPUT_BRANCH_MAP.get(code,code); break
        i=0
        while i<len(answers):
            ans=answers[i]; title=ans.get("title","").strip(); matched=None
            for ch in ZENPUT_CHANNELS:
                if title==ch or title.startswith(ch[:20]): matched=ch; break
            if matched:
                short=ZENPUT_CHANNEL_SHORT[matched]; amt=ans.get("value",0) or 0; inv=0
                if i+1<len(answers):
                    nxt=answers[i+1]
                    if "Invoices" in nxt.get("title","") or "فاتور" in nxt.get("title",""): inv=nxt.get("value",0) or 0; i+=1
                row[f"{short} - Amount"]=amt; row[f"{short} - Invoices"]=inv
            elif any(k in title for k in ["Total invoices","مجموع عدد الفواتير"]): row["Total Invoices"]=ans.get("value",0) or 0
            elif any(k in title for k in ["Total cash & credit","اجمالي المبيعات"]): row["Total Sales"]=ans.get("value",0) or 0
            elif any(k in title for k in ["Sales by Foodics","المبيعات في فوديكس"]): row["Foodics Sales"]=ans.get("value",0) or 0
            elif any(k in title for k in ["Difference","الفرق"]): row["Difference"]=ans.get("value",0) or 0
            elif any(k in title for k in ["Notes","ملاحظات"]): row["Notes"]=ans.get("value","") or ""
            i+=1
        rows.append(row)
    df=pd.DataFrame(rows); df["_sort"]=df["Branch"].apply(branch_code_sort_key)
    df=df.sort_values(["Date","_sort"]).drop("_sort",axis=1).reset_index(drop=True); return df

def _zenput_group_df(df):
    if df.empty: return df
    text_cols=["Date","Branch","Notes"]; numeric_cols=[c for c in df.columns if c not in text_cols]
    df_work=df.copy()
    for c in numeric_cols: df_work[c]=pd.to_numeric(df_work[c],errors="coerce").fillna(0)
    grp=df_work.groupby(["Date","Branch"],sort=False)
    agg_dict={c:"sum" for c in numeric_cols}
    if "Notes" in df_work.columns: agg_dict["Notes"]=lambda x:" | ".join(v for v in x if str(v).strip())
    grouped=grp.agg(agg_dict).reset_index()
    grouped["_sort"]=grouped["Branch"].apply(branch_code_sort_key)
    grouped=grouped.sort_values(["Date","_sort"]).drop("_sort",axis=1).reset_index(drop=True)
    drop_cols=[c for c in grouped.columns if (c in ("Notes","Total Invoices")) or ("Invoices" in c and c!="Cash Purchase Inv. - Invoices")]
    grouped=grouped.drop(columns=drop_cols,errors="ignore")
    priority_end=["Total Sales","Foodics Sales","Difference"]
    middle=[c for c in grouped.columns if c not in ["Date","Branch"]+priority_end]
    return grouped[["Date","Branch"]+middle+[c for c in priority_end if c in grouped.columns]]

def zenput_build_excel(df):
    df_grouped=_zenput_group_df(df)
    if df_grouped.empty:
        buf=io.BytesIO(); Workbook().save(buf); buf.seek(0); return buf
    wb=Workbook(); wb.remove(wb.active)
    hf=PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    brf=PatternFill(start_color="2E4057",end_color="2E4057",fill_type="solid")
    amf=PatternFill(start_color="DEEAF1",end_color="DEEAF1",fill_type="solid")
    inf=PatternFill(start_color="E2EFDA",end_color="E2EFDA",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    smf=PatternFill(start_color="FCE4D6",end_color="FCE4D6",fill_type="solid")
    gf=PatternFill(start_color="FF0000",end_color="FF0000",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center",wrap_text=True); rt=Alignment(horizontal="right"); lt=Alignment(horizontal="left")
    dates=sorted(df_grouped["Date"].unique()); cols=[c for c in df_grouped.columns if c!="Date"]
    acols=[c for c in cols if "Amount" in c or "Sales" in c or "Difference" in c]
    icols=[c for c in cols if "Invoices" in c and c!="Total Invoices"]
    ticols=[c for c in cols if c=="Total Invoices"]; tscols=[c for c in cols if c in ("Total Sales","Foodics Sales")]
    for date_str in dates:
        try: tab_name=_dt_ui.datetime.strptime(date_str,"%Y-%m-%d").strftime("%d-%b")
        except Exception: tab_name=str(date_str)[:10]
        ws=wb.create_sheet(title=tab_name); df_date=df_grouped[df_grouped["Date"]==date_str].reset_index(drop=True)
        ws.cell(row=1,column=1,value=f"Financial Summary — {tab_name}").fill=hf
        ws.cell(row=1,column=1).font=Font(color="FFFFFF",bold=True,size=12); ws.cell(row=1,column=1).alignment=ct
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols))
        for ci,cn in enumerate(cols,1):
            cell=ws.cell(row=2,column=ci,value=cn); cell.font=Font(color="FFFFFF",bold=True,size=9); cell.alignment=ct
            if cn=="Branch": cell.fill=brf
            elif cn in acols: cell.fill=PatternFill(start_color="1A5276",end_color="1A5276",fill_type="solid")
            elif cn in icols: cell.fill=PatternFill(start_color="1E8449",end_color="1E8449",fill_type="solid")
            elif cn in ticols+tscols: cell.fill=PatternFill(start_color="B7770D",end_color="B7770D",fill_type="solid")
            elif cn=="Difference": cell.fill=PatternFill(start_color="922B21",end_color="922B21",fill_type="solid")
            elif cn=="Notes": cell.fill=PatternFill(start_color="566573",end_color="566573",fill_type="solid")
            else: cell.fill=hf
        for ri,row_data in enumerate(df_date.itertuples(index=False),3):
            rd=dict(zip(df_date.columns,row_data))
            for ci,cn in enumerate(cols,1):
                val=rd.get(cn,""); cell=ws.cell(row=ri,column=ci,value=val)
                if cn=="Branch": cell.font=Font(bold=True,size=9); cell.alignment=lt
                elif cn in acols: cell.fill=amf; cell.number_format="#,##0.00"; cell.alignment=rt; cell.font=Font(size=9)
                elif cn in icols: cell.fill=inf; cell.number_format="#,##0"; cell.alignment=rt; cell.font=Font(size=9)
                elif cn in ticols: cell.fill=tf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0"; cell.alignment=rt
                elif cn in tscols: cell.fill=tf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0.00"; cell.alignment=rt
                elif cn=="Difference": cell.fill=smf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0.00"; cell.alignment=rt
                elif cn=="Notes": cell.alignment=lt; cell.font=Font(size=9)
        trix=3+len(df_date); nsc=[c for c in cols if c not in ("Branch","Notes")]
        for ci,cn in enumerate(cols,1):
            cell=ws.cell(row=trix,column=ci); cell.fill=gf; cell.font=Font(bold=True,color="FFFFFF",size=10)
            if cn=="Branch": cell.value="TOTAL"; cell.alignment=ct
            elif cn in nsc:
                cv=pd.to_numeric(df_date[cn],errors="coerce").fillna(0); cell.value=cv.sum()
                cell.number_format="#,##0" if cn in icols+ticols else "#,##0.00"; cell.alignment=rt
            else: cell.value=""
        ws.column_dimensions["A"].width=16
        for i in range(2,len(cols)+1):
            cn=cols[i-1]
            if cn=="Notes": ws.column_dimensions[get_column_letter(i)].width=25
            elif "Invoices" in cn: ws.column_dimensions[get_column_letter(i)].width=9
            else: ws.column_dimensions[get_column_letter(i)].width=14
        ws.freeze_panes="B3"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def zenput_build_excel_by_branch(df):
    df_grouped=_zenput_group_df(df)
    if df_grouped.empty:
        buf=io.BytesIO(); Workbook().save(buf); buf.seek(0); return buf
    cols_nd=[c for c in df_grouped.columns if c!="Date"]; nc=[c for c in cols_nd if c not in ("Branch","Notes")]
    df_work=df_grouped[cols_nd].copy()
    for c in nc: df_work[c]=pd.to_numeric(df_work[c],errors="coerce").fillna(0)
    agg={c:"sum" for c in nc}
    if "Notes" in df_work.columns: agg["Notes"]=lambda x:" | ".join(v for v in x if str(v).strip())
    df_branch=df_work.groupby("Branch",sort=False).agg(agg).reset_index()
    df_branch["_sort"]=df_branch["Branch"].apply(branch_code_sort_key)
    df_branch=df_branch.sort_values("_sort").drop("_sort",axis=1).reset_index(drop=True)
    pe=["Total Invoices","Total Sales","Foodics Sales","Difference","Notes"]
    mid=[c for c in df_branch.columns if c not in ["Branch"]+pe]
    df_branch=df_branch[["Branch"]+mid+[c for c in pe if c in df_branch.columns]]
    cols=df_branch.columns.tolist()
    acols=[c for c in cols if "Amount" in c or "Sales" in c or "Difference" in c]
    icols=[c for c in cols if "Invoices" in c and c!="Total Invoices"]
    ticols=[c for c in cols if c=="Total Invoices"]; tscols=[c for c in cols if c in ("Total Sales","Foodics Sales")]
    wb=Workbook(); ws=wb.active; ws.title="By Branch"
    hf=PatternFill(start_color="1F4E78",end_color="1F4E78",fill_type="solid")
    amf=PatternFill(start_color="DEEAF1",end_color="DEEAF1",fill_type="solid")
    inf=PatternFill(start_color="E2EFDA",end_color="E2EFDA",fill_type="solid")
    tf=PatternFill(start_color="FFC000",end_color="FFC000",fill_type="solid")
    smf=PatternFill(start_color="FCE4D6",end_color="FCE4D6",fill_type="solid")
    gf=PatternFill(start_color="FF0000",end_color="FF0000",fill_type="solid")
    ct=Alignment(horizontal="center",vertical="center",wrap_text=True); rt=Alignment(horizontal="right"); lt=Alignment(horizontal="left")
    for ci,cn in enumerate(cols,1):
        cell=ws.cell(row=1,column=ci,value=cn); cell.font=Font(color="FFFFFF",bold=True,size=9); cell.alignment=ct
        if cn=="Branch": cell.fill=PatternFill(start_color="2E4057",end_color="2E4057",fill_type="solid")
        elif cn in acols: cell.fill=PatternFill(start_color="1A5276",end_color="1A5276",fill_type="solid")
        elif cn in icols: cell.fill=PatternFill(start_color="1E8449",end_color="1E8449",fill_type="solid")
        elif cn in ticols+tscols: cell.fill=PatternFill(start_color="B7770D",end_color="B7770D",fill_type="solid")
        elif cn=="Difference": cell.fill=PatternFill(start_color="922B21",end_color="922B21",fill_type="solid")
        elif cn=="Notes": cell.fill=PatternFill(start_color="566573",end_color="566573",fill_type="solid")
        else: cell.fill=hf
    for ri,row_data in enumerate(df_branch.itertuples(index=False),2):
        rd=dict(zip(cols,row_data))
        for ci,cn in enumerate(cols,1):
            val=rd.get(cn,""); cell=ws.cell(row=ri,column=ci,value=val)
            if cn=="Branch": cell.font=Font(bold=True,size=9); cell.alignment=lt
            elif cn in acols: cell.fill=amf; cell.number_format="#,##0.00"; cell.alignment=rt; cell.font=Font(size=9)
            elif cn in icols: cell.fill=inf; cell.number_format="#,##0"; cell.alignment=rt; cell.font=Font(size=9)
            elif cn in ticols: cell.fill=tf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0"; cell.alignment=rt
            elif cn in tscols: cell.fill=tf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0.00"; cell.alignment=rt
            elif cn=="Difference": cell.fill=smf; cell.font=Font(bold=True,size=9); cell.number_format="#,##0.00"; cell.alignment=rt
            elif cn=="Notes": cell.alignment=lt; cell.font=Font(size=9)
    trix=2+len(df_branch); nsc=[c for c in cols if c not in ("Branch","Notes")]
    for ci,cn in enumerate(cols,1):
        cell=ws.cell(row=trix,column=ci); cell.fill=gf; cell.font=Font(bold=True,color="FFFFFF",size=10)
        if cn=="Branch": cell.value="TOTAL"; cell.alignment=ct
        elif cn in nsc:
            cell.value=pd.to_numeric(df_branch[cn],errors="coerce").fillna(0).sum()
            cell.number_format="#,##0" if cn in icols+ticols else "#,##0.00"; cell.alignment=rt
        else: cell.value=""
    ws.column_dimensions["A"].width=16
    for i in range(2,len(cols)+1):
        cn=cols[i-1]
        if cn=="Notes": ws.column_dimensions[get_column_letter(i)].width=25
        elif "Invoices" in cn: ws.column_dimensions[get_column_letter(i)].width=9
        else: ws.column_dimensions[get_column_letter(i)].width=14
    ws.freeze_panes="B2"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def zenput_push_to_sheet(df,tab_name):
    try: gc=get_gspread_client()
    except Exception as e: return False,f"Auth failed: {str(e)}"
    try: sh=gc.open_by_key(ZENPUT_SHEET_ID)
    except Exception as e: return False,f"Cannot open sheet — share it with the service account first.\nError: {str(e)}"
    try:
        try:
            ws=sh.worksheet(tab_name); ws.clear()
            if ws.row_count<len(df)+10 or ws.col_count<len(df.columns)+5: ws.resize(rows=len(df)+10,cols=len(df.columns)+5)
        except Exception: ws=sh.add_worksheet(title=tab_name,rows=len(df)+10,cols=len(df.columns)+5)
        rows=[df.columns.tolist()]+df.fillna("").astype(str).values.tolist()
        ws.update(rows,value_input_option="USER_ENTERED")
        def _rgb(r,g,b): return {"red":r/255,"green":g/255,"blue":b/255}
        n=len(df.columns)
        sh.batch_update({"requests":[
            {"repeatCell":{"range":{"sheetId":ws.id,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":n},"cell":{"userEnteredFormat":{"backgroundColor":_rgb(31,78,120),"textFormat":{"bold":True,"fontSize":9,"foregroundColor":{"red":1,"green":1,"blue":1}},"horizontalAlignment":"CENTER","wrapStrategy":"WRAP"}},"fields":"userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)"}},
            {"updateSheetProperties":{"properties":{"sheetId":ws.id,"gridProperties":{"frozenRowCount":1}},"fields":"gridProperties.frozenRowCount"}},
            {"autoResizeDimensions":{"dimensions":{"sheetId":ws.id,"dimension":"COLUMNS","startIndex":0,"endIndex":n}}}
        ]})
        return True,f"Pushed {len(df)} rows to tab '{tab_name}'"
    except Exception as e:
        import traceback; return False,f"Write error: {str(e)}\n{traceback.format_exc()}"

# ==================== UI ====================

st.title("🏦 Geidea & Foodics Summary Generator")
st.markdown("Upload your reconciliation file to generate summary reports")

uploaded_file=st.file_uploader("📁 Upload file (Geidea or Foodics) — supports .xlsx, .xls, .csv",type=["xlsx","xls","csv"])

if uploaded_file:
    try:
        df_raw=read_uploaded_file(uploaded_file); file_type=detect_file_type(df_raw)

        if file_type=="geidea":
            st.success(f"✅ Detected **Geidea** file: {uploaded_file.name} ({len(df_raw)} rows)")
            with st.expander("🔍 Preview Raw Data (with row numbers)"):
                preview_df=df_raw.copy(); preview_df.index=range(1,len(preview_df)+1)
                st.dataframe(preview_df,use_container_width=True,height=300)

            st.markdown("---")
            st.subheader("📅 Split by Date (Optional)")

            _time_col=next((c for c in df_raw.columns if "time" in c.lower() and "recon" in c.lower()),None)
            _recon_col=next((c for c in df_raw.columns if "date" in c.lower() and "recon" in c.lower()),None)

            use_split=st.checkbox("✂️ Split rows into separate date sheets",value=False)
            date_slice_dfs=[]; valid=True

            if use_split and _time_col and _recon_col:
                st.markdown(
                    "Each row's **Reconciliation Date** is read from the file. "
                    "If the row's time is **before the cutoff set for that date**, it shifts back to the previous day. "
                    "You can set a **different cutoff per date**."
                )
                raw_times=pd.to_datetime(df_raw[_time_col],format="%H:%M:%S",errors="coerce")
                raw_dates=pd.to_datetime(df_raw[_recon_col],errors="coerce").dt.date
                min_ts=raw_times.dt.strftime("%H:%M").min() if raw_times.notna().any() else "00:00"
                max_ts=raw_times.dt.strftime("%H:%M").max() if raw_times.notna().any() else "23:59"
                unique_dates_in_file=sorted(raw_dates.dropna().unique())

                st.caption(f"Times in file range from **{min_ts}** to **{max_ts}**")
                st.caption(f"Reconciliation dates found: **{', '.join(d.strftime('%d-%b-%Y') for d in unique_dates_in_file)}**")

                if not unique_dates_in_file:
                    st.warning("⚠️ No valid dates found in the Reconciliation Date column."); valid=False
                else:
                    st.markdown("**Set a cutoff time for each date** — rows before the cutoff shift to the previous day:")
                    cutoff_map={}
                    date_cols=st.columns(min(len(unique_dates_in_file),4))
                    for i,recon_date in enumerate(unique_dates_in_file):
                        with date_cols[i%4]:
                            prev_day=(recon_date-_dt_ui.timedelta(days=1)).strftime("%d-%b")
                            cutoff_str=st.text_input(
                                f"{recon_date.strftime('%d-%b')} cutoff",
                                value="04:22",
                                key=f"cutoff_date_{recon_date}",
                                help=f"Rows on {recon_date.strftime('%d-%b')} before this time → go to {prev_day}"
                            )
                            try:
                                h,m=cutoff_str.strip().split(":")
                                cutoff_map[recon_date]=_dt_ui.time(int(h),int(m))
                            except Exception:
                                st.error(f"Invalid format for {recon_date.strftime('%d-%b')}: use HH:MM"); valid=False

                    if valid:
                        row_times=raw_times.dt.time

                        def _compute_real_date(row_time,row_date):
                            if pd.isna(row_time) or row_date is None: return row_date
                            cutoff=cutoff_map.get(row_date)
                            if cutoff is None: return row_date
                            return row_date-_dt_ui.timedelta(days=1) if row_time<cutoff else row_date

                        real_dates=[_compute_real_date(t,d) for t,d in zip(row_times,raw_dates)]
                        real_dates_series=pd.Series(real_dates,index=df_raw.index)
                        unique_real_dates=sorted(set(d for d in real_dates if d is not None),reverse=True)

                        for real_date in unique_real_dates:
                            mask=real_dates_series==real_date; df_slice=df_raw[mask].reset_index(drop=True)
                            if len(df_slice)>0: date_slice_dfs.append((real_date.strftime("%d-%b"),df_slice))

                        if date_slice_dfs:
                            st.info("📌 "+"  |  ".join([f"**`{lbl}`** {len(df_s)} rows" for lbl,df_s in date_slice_dfs]))
                            with st.expander("👁 Preview each date slice"):
                                for lbl,df_s in date_slice_dfs:
                                    st.markdown(f"**{lbl}** — {len(df_s)} rows")
                                    st.dataframe(df_s[[_time_col,_recon_col]+[c for c in df_s.columns if c not in [_time_col,_recon_col]]].head(5),use_container_width=True)

            elif use_split and (not _time_col or not _recon_col):
                st.warning("⚠️ Could not find both 'Reconciliation Time' and 'Reconciliation Date' columns.")

            if not date_slice_dfs:
                if _recon_col:
                    try:
                        _d=pd.to_datetime(df_raw[_recon_col],errors="coerce").dt.date.dropna().unique()
                        if len(_d)==1: lbl=sorted(_d)[-1].strftime("%d-%b")
                        elif len(_d)>1: lbl=f"{sorted(_d)[0].strftime('%d-%b')}–{sorted(_d)[-1].strftime('%d-%b')}"
                        else: lbl="All"
                    except Exception: lbl="All"
                else: lbl="All"
                date_slice_dfs=[(lbl,df_raw.copy())]

            st.markdown("---")
            with st.spinner("Processing Geidea reports..."):
                processed_slices=[(label,process_geidea_data(df_s)) for label,df_s in date_slice_dfs]
                df_processed=processed_slices[0][1]
                summary_buffer,summary_df,grand_total=create_geidea_summary_file(df_processed)
                detailed_buffer,num_terminals=create_geidea_detailed_file(df_processed)
                detailed_tot_buffer,_=create_geidea_detailed_totals_only(processed_slices)
                unique_dates=df_processed["Reconciliation Date"].dropna().unique()
                has_multiple_dates=len(unique_dates)>1
                if has_multiple_dates:
                    summary_date_buffer,_,num_dates=create_geidea_summary_by_date_file(df_processed)
                    date_buffer,_,_=create_geidea_detailed_by_date_file(df_processed)
                else: summary_date_buffer=date_buffer=None; num_dates=0

            st.subheader("📊 Geidea Summary Preview")
            c1,c2,c3=st.columns(3); c1.metric("Banks",summary_df["Bank Name"].nunique())
            c2.metric("Card Schemes",summary_df["Card Name"].nunique()); c3.metric("Grand Total",f"{grand_total:,.0f}")
            if has_multiple_dates: st.info(f"📅 Detected {num_dates} dates: {', '.join([d.strftime('%Y-%m-%d') for d in unique_dates])}")
            if "Unknown Bank" in summary_df["Bank Name"].values: st.warning("⚠️ Some terminals not found in mapping (shown in red)")
            st.dataframe(summary_df.style.format({"Total":"{:,.2f}"}).apply(lambda x:["background-color: #FF6B6B; color: white"]*3 if x["Bank Name"]=="Unknown Bank" else [""]*3,axis=1),use_container_width=True,height=300)

            st.subheader("⬇️ Geidea: Download Reports")
            nr=5 if has_multiple_dates else 3
            pn=f" *({len(date_slice_dfs)} date sheets)*" if use_split and valid and len(date_slice_dfs)>1 else ""
            st.markdown(f"**{nr} Geidea reports generated:**")
            c1,c2=st.columns(2)
            with c1:
                st.download_button("📊 Summary Totals Only",data=summary_buffer,file_name="Geidea_01_SUMMARY_Totals_Only.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("📋 Detailed — Full (Debit / Credit / Total)",data=detailed_buffer,file_name=f"Geidea_02_DETAILED_Full_{num_terminals}_terminals.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button(f"📋 Detailed — Total per Terminal{pn}",data=detailed_tot_buffer,file_name=f"Geidea_03_DETAILED_Total_{num_terminals}_terminals.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            with c2:
                if has_multiple_dates:
                    st.download_button("📅 Summary by Date",data=summary_date_buffer,file_name=f"Geidea_04_SUMMARY_by_Date_{num_dates}_dates.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                    st.download_button("📆 Detailed by Date (Full)",data=date_buffer,file_name=f"Geidea_05_DETAILED_by_Date_{num_dates}_dates.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.success(f"✅ All {nr} Geidea reports ready! ({num_terminals} terminals)")

        elif file_type=="foodics":
            st.success(f"✅ Detected **Foodics** Payments Report: {uploaded_file.name}")
            with st.expander("🔍 Preview Raw Data"): st.dataframe(df_raw.head(15),use_container_width=True)
            with st.spinner("Processing Foodics reports..."):
                df_processed,dates=process_foodics_data(df_raw)
                branch_buffer,branch_summary,num_branches=create_foodics_summary_by_branch(df_processed)
                payment_buffer,payment_summary=create_foodics_summary_by_payment_method(df_processed)
                det_br_full_buf,num_br_det,_=create_foodics_detailed_by_branch(df_processed)
                det_pm_full_buf,num_pm_cols,_=create_foodics_detailed_by_payment_method(df_processed)
                det_br_net_buf,_,_=_foodics_net_only_pivot(df_processed,row_field="Payment Method",col_field="Branch",row_label="Payment Method",col_label="Branch",header_hex="2E7D32",sub_hex="A5D6A7")
                det_pm_net_buf,_,_=_foodics_net_only_pivot(df_processed,row_field="Branch",col_field="Payment Method",row_label="Branch",col_label="Payment Method",header_hex="1565C0",sub_hex="90CAF9")
                if dates: avg_buffer,_,num_days=create_foodics_daily_avg_report(df_processed,dates)
                else: avg_buffer=None; num_days=0
            st.subheader("📊 Foodics Summary Preview")
            c1,c2,c3=st.columns(3); c1.metric("Branches",num_branches)
            c2.metric("Payment Methods",payment_summary["Payment Method"].nunique()); c3.metric("Total Net Amount",f"{payment_summary['Net Amount'].sum():,.0f}")
            if dates: st.info(f"📅 Report period: {dates[0]} to {dates[-1]} ({num_days} days)")
            else: st.info("ℹ️ No date range metadata — daily averages not available for plain CSV.")
            st.dataframe(payment_summary.style.format({"Net Amount":"{:,.2f}","Amount":"{:,.2f}","Return Amount":"{:,.2f}","Count":"{:,.0f}"}),use_container_width=True,height=300)
            st.subheader("⬇️ Foodics: Download Reports")
            tr=7 if dates else 6; st.markdown(f"**{tr} Foodics reports generated:**")
            c1,c2=st.columns(2)
            with c1:
                st.download_button("🏪 Summary by Branch",data=branch_buffer,file_name=f"Foodics_01_SUMMARY_by_Branch_{num_branches}_branches.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("📋 Detailed by Branch — Full",data=det_br_full_buf,file_name=f"Foodics_03_DETAILED_Branch_Full_{num_br_det}_branches.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("📋 Detailed by Branch — Net Total",data=det_br_net_buf,file_name=f"Foodics_05_DETAILED_Branch_Net_{num_br_det}_branches.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                if dates: st.download_button("📈 Daily Averages",data=avg_buffer,file_name=f"Foodics_07_Daily_Averages_{num_days}_days.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            with c2:
                st.download_button("💳 Summary by Payment Method",data=payment_buffer,file_name="Foodics_02_SUMMARY_by_Payment_Method.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("📊 Detailed by Payment Method — Full",data=det_pm_full_buf,file_name=f"Foodics_04_DETAILED_PayMethod_Full_{num_pm_cols}_methods.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("📊 Detailed by Payment Method — Net",data=det_pm_net_buf,file_name=f"Foodics_06_DETAILED_PayMethod_Net_{num_pm_cols}_methods.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            st.success(f"✅ {tr} Foodics reports ready! ({num_branches} branches · {payment_summary['Payment Method'].nunique()} payment methods)")
        else:
            st.error("❌ Could not detect file type.")
            st.info("**Geidea:** File must have 'Terminal' and 'Card Name' columns")
            st.info("**Foodics:** File must have 'Payment Method' and 'Branch' columns")
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}"); st.info("Please check your file format and try again.")

st.markdown("---")
st.header("📋 Zenput Financial Form")
st.markdown("Fetch financial form submissions (template **1556454**) from Zenput, generate a formatted Excel summary, and optionally push to Google Sheets.")
_today=_dt_ui.date.today()
z_c1,z_c2=st.columns(2)
with z_c1: z_start=st.date_input("📅 From date",value=_today,key="z_start")
with z_c2: z_end=st.date_input("📅 To date",value=_today,key="z_end")
z_c3,z_c4=st.columns(2)
with z_c3: z_dl=st.checkbox("⬇️ Download as Excel",value=True,key="z_dl")
with z_c4: z_push=st.checkbox("📤 Push to Google Sheet",value=False,key="z_push")
if z_push:
    z_tab_name=st.text_input("Sheet tab name",value=f"Financial_{z_start.strftime('%d-%b')}"+( f"_to_{z_end.strftime('%d-%b')}" if z_end!=z_start else ""),key="z_tab")
if st.button("🚀 Fetch Zenput Submissions",type="primary",use_container_width=True,key="z_run"):
    if z_start>z_end: st.error("Start date must be ≤ end date.")
    else:
        with st.spinner(f"Fetching Zenput submissions {z_start} → {z_end} ..."):
            try: z_df=zenput_fetch_financial(str(z_start),str(z_end))
            except Exception as e: st.error(f"❌ Fetch error: {str(e)}"); z_df=pd.DataFrame()
        if z_df.empty: st.warning("⚠️ No submissions found for the selected date range.")
        else:
            st.success(f"✅ Fetched **{len(z_df)} submissions** ({z_start} → {z_end})")
            with st.expander("👁 Preview data"): st.dataframe(z_df,use_container_width=True,height=300)
            if z_dl:
                _fb=f"Zenput_Financial_{z_start}{'_to_'+str(z_end) if z_end!=z_start else ''}"
                zbd=zenput_build_excel(z_df); zbb=zenput_build_excel_by_branch(z_df)
                dl_c1,dl_c2=st.columns(2)
                with dl_c1: st.download_button("⬇️ By Date (one sheet per day)",data=zbd,file_name=f"{_fb}_ByDate.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="z_dl_date")
                with dl_c2: st.download_button("⬇️ By Branch (all dates combined)",data=zbb,file_name=f"{_fb}_ByBranch.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="z_dl_branch")
            if z_push:
                with st.spinner("Pushing to Google Sheets..."):
                    ok,msg=zenput_push_to_sheet(z_df,z_tab_name)
                if ok: st.success(f"✅ {msg}"); st.markdown(f"[🔗 Open Google Sheet](https://docs.google.com/spreadsheets/d/{ZENPUT_SHEET_ID}/edit)")
                else: st.error(msg)

st.markdown("---")
st.caption("Geidea & Foodics v6.0 | Per-date cutoff times · Zenput Financial")
