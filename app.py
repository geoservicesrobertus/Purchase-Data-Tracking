from datetime import datetime, timezone, timedelta
import glob
import os
import re
import pandas as pd
import streamlit as st

# Konfigurasi Zona Waktu Jakarta (WIB / UTC+7)
WIB = timezone(timedelta(hours=7))
LOCAL_ARCHIVE_DIR = "archives"
os.makedirs(LOCAL_ARCHIVE_DIR, exist_ok=True)


def list_local_archives():
    try:
        pattern = os.path.join(LOCAL_ARCHIVE_DIR, "Tracking_Final_*.xlsx")
        files = glob.glob(pattern)
        files.sort(key=os.path.getmtime, reverse=True)
        result = []
        for f in files:
            fname = os.path.basename(f)
            dt_utc = datetime.fromtimestamp(os.path.getmtime(f), timezone.utc)
            dt_wib = dt_utc.astimezone(WIB)
            mtime = dt_wib.strftime("%Y-%m-%d %H:%M:%S WIB")
            result.append((f"{mtime} | {fname}", f, fname))
        return result
    except Exception as e:
        st.warning(f"Gagal membaca folder arsip lokal: {e}")
        return []


def clean_po(df, tipe_label):
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    # Sesuaikan nama kolom mentah file PO lu jika perlu di sini
    d['Tipe_PO'] = tipe_label
    return d


def process_tracking_data(lo_df, imp_df, pr_df, inb_df):
    pr_data = pr_df.copy()
    pr_data['PR_Qty_Num'] = pd.to_numeric(pr_data['PR_Qty'], errors='coerce').fillna(0)
    # Filter PR Qty == 0 atau NaN dianggap dibatalkan / skip
    pr_data = pr_data[pr_data['PR_Qty_Num'] > 0].copy()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.strip().str.upper()
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip().str.upper()

    # Gabung & Bersihkan PO
    po_df = pd.concat([lo_df.copy(), imp_df.copy()], ignore_index=True)
    po_df['Item_Code'] = po_df['Item_Code'].astype(str).str.strip().str.upper()
    po_df['PO_No'] = po_df['PO_No'].astype(str).str.strip()

    expanded_rows = []
    for _, row in po_df.iterrows():
        pr_str = str(row['PR_Manual_No_Orig']).strip()
        if pr_str in ('', 'nan', 'None', '-'):
            continue
        if ',' in pr_str or '/' in pr_str:
            parts = re.split(r'[,/]', pr_str)
            base_prefix = ""
            for p in [x.strip() for x in parts if x.strip()]:
                if not p.isdigit():
                    m = re.match(r'([A-Za-z.\-]+)(\d+)', p)
                    if m:
                        base_prefix = m.group(1)
                new_row = row.to_dict()
                c_val = ((base_prefix + p).replace(" ", "").upper() if p.isdigit() and base_prefix else p.replace(" ", "").upper())
                new_row['PR_Manual_No_Clean'] = c_val
                expanded_rows.append(new_row)
        else:
            new_row = row.to_dict()
            new_row['PR_Manual_No_Clean'] = pr_str.replace(" ", "").upper()
            expanded_rows.append(new_row)

    po_exp = pd.DataFrame(expanded_rows)

    if po_exp.empty:
        po_agg = pd.DataFrame(columns=['PR_Manual_No_Clean', 'Item_Code', 'PO_No', 'PO_Date', 'PO_Qty', 'Vendor', 'Tipe_PO'])
    else:
        po_exp['PO_Qty_Num'] = pd.to_numeric(po_exp['PO_Qty'], errors='coerce').fillna(0)
        po_exp = po_exp.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code', 'PO_No'], keep='last')
        po_agg = po_exp.groupby(['PR_Manual_No_Clean', 'Item_Code', 'PO_No'], as_index=False).agg(
            PO_Date=('PO_Date', 'max'),
            PO_Qty=('PO_Qty_Num', 'sum'),
            Vendor=('Vendor', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique())),
            Tipe_PO=('Tipe_PO', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique())),
        )

    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    if 'RequestClosed' in merged.columns:
        merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

    # Proses Inbound
    inb_clean = inb_df.copy()
    inb_clean['ItemCode'] = inb_clean['ItemCode'].astype(str).str.strip().str.upper()
    inb_clean['POnumber'] = inb_clean['POnumber'].astype(str).str.strip()
    inb_clean['ReceivedDate'] = pd.to_datetime(inb_clean['ReceivedDate'], errors='coerce')
    inb_clean['RcvQty'] = pd.to_numeric(inb_clean['RcvQty'], errors='coerce').fillna(0)
    inb_agg = inb_clean.groupby(['POnumber', 'ItemCode']).agg(
        Rcv_Date=('ReceivedDate', 'max'), Rcv_Qty=('RcvQty', 'sum')
    ).reset_index()

    def get_inb(po_str, item_code):
        empty = pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0.0})
        if not po_str or str(po_str) == '':
            return empty
        po_list = [p.strip() for p in str(po_str).split(',') if p.strip()]
        c_item = str(item_code).strip().upper()
        subset = inb_agg[inb_agg['POnumber'].isin(po_list) & (inb_agg['ItemCode'] == c_item)]
        if subset.empty:
            return empty
        return pd.Series({'Rcv_Date': subset['Rcv_Date'].max(), 'Rcv_Qty': subset['Rcv_Qty'].sum()})

    merged[['Rcv_Date', 'Rcv_Qty']] = merged.apply(lambda row: get_inb(row['PO_No'], row['Item_Code']), axis=1)

    merged['PO_Qty'] = pd.to_numeric(merged['PO_Qty'], errors='coerce').fillna(0)
    merged['Rcv_Qty'] = pd.to_numeric(merged['Rcv_Qty'], errors='coerce').fillna(0)

    def get_status(row):
        if row['PO_No'] == '':
            return 'Routing Approval'
        if row['Rcv_Qty'] == 0:
            return 'Menunggu Pengiriman'
        if row['Rcv_Qty'] < row['PO_Qty']:
            return 'Diterima Sebagian'
        return 'Sudah Diterima'

    merged['Status'] = merged.apply(get_status, axis=1)

    # Qty Outstanding rule: Routing approval = full PR_Qty, selain itu max(0, PO_Qty - Rcv_Qty)
    def calc_outstanding(row):
        if row['Status'] == 'Routing Approval':
            return row['PR_Qty_Num']
        return max(0, row['PO_Qty'] - row['Rcv_Qty'])

    merged['Qty_Outstanding'] = merged.apply(calc_outstanding, axis=1)

    for col in ('PR_Date', 'PO_Date', 'Rcv_Date'):
        if col in merged.columns:
            dt_s = pd.to_datetime(merged[col], errors='coerce')
            merged[col] = dt_s.dt.strftime('%m/%d/%Y').fillna('-')

    for col in ['PO_No', 'Vendor', 'Tipe_PO']:
        if col not in merged.columns:
            merged[col] = '-'
        else:
            merged[col] = merged[col].fillna('-')

    cols_order = [
        'PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty_Num',
        'PO_Date', 'PO_No', 'Vendor', 'Tipe_PO', 'PO_Qty',
        'Rcv_Date', 'Rcv_Qty', 'Qty_Outstanding', 'Status',
    ]
    valid_cols = [c for c in cols_order if c in merged.columns]
    res_df = merged[valid_cols].copy()
    if 'PR_Qty_Num' in res_df.columns:
        res_df.rename(columns={'PR_Qty_Num': 'PR_Qty'}, inplace=True)
    return res_df


# Streamlit UI Main Flow Stub (sesuaikan bagian uploader/tab lu yang sudah ada)
st.set_page_config(page_title="Tracking Pengadaan", layout="wide")
st.title("Tracking System PR / PO / Inbound")

# Contoh layout uploader/tab sederhana pemanggil
uploaded_lo = st.sidebar.file_uploader("Upload PO Lokal", type=["xlsx"])
uploaded_imp = st.sidebar.file_uploader("Upload PO Impor", type=["xlsx"])
uploaded_pr = st.sidebar.file_uploader("Upload PR", type=["xlsx"])
uploaded_inb = st.sidebar.file_uploader("Upload Inbound", type=["xlsx"])

if uploaded_pr and uploaded_inb:
    try:
        lo_df = pd.read_excel(uploaded_lo) if uploaded_lo else pd.DataFrame()
        imp_df = pd.read_excel(uploaded_imp) if uploaded_imp else pd.DataFrame()
        pr_df = pd.read_excel(uploaded_pr)
        inb_df = pd.read_excel(uploaded_inb)

        final_res = process_tracking_data(lo_df, imp_df, pr_df, inb_df)
        st.dataframe(final_res, use_container_width=True)
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
