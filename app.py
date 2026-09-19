import streamlit as st
import pandas as pd
import re
import io
import os
import base64
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account
import plotly.express as px

st.set_page_config(
    page_title="Purchase Data Tracking - PT. Geoservices",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .main-title {font-size: 1.5rem; font-weight: 700; color: #1e3d59; margin: 0;}
        .sub-meta {font-size: 0.8rem; color: #5f6c7b;}
        .erp-panel {background: #ffffff; border: 1px solid #c0c0c0; border-radius: 3px; padding: 15px; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

GDRIVE_FOLDER_ID = "1Z9-pgpCBqJ3iEjUdU7URLJZSlnU_Hgyk"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------- Google Drive

@st.cache_resource
def get_gdrive_service():
    """Bangun Drive client sekali saja, lalu dipakai ulang."""
    creds_dict = dict(st.secrets["gcp_service_account"])

    # Kalau private_key tersimpan dengan \n literal, ubah jadi newline asli.
    pk = str(creds_dict.get("private_key", "")).strip()
    if "\\n" in pk:
        pk = pk.replace("\\n", "\n")
    creds_dict["private_key"] = pk

    creds = service_account.Credentials.from_service_account_info(
        creds_dict, scopes=DRIVE_SCOPES
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def upload_to_gdrive(file_buffer, filename):
    try:
        service = get_gdrive_service()
        media = MediaIoBaseUpload(file_buffer, mimetype=XLSX_MIME, resumable=True)
        file = service.files().create(
            body={"name": filename, "parents": [GDRIVE_FOLDER_ID]},
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
        return True, file.get("id")
    except Exception as e:
        return False, str(e)


def list_gdrive_archives():
    """Return list of (label, file_id, filename), terbaru dulu."""
    try:
        service = get_gdrive_service()
        query = (
            f"'{GDRIVE_FOLDER_ID}' in parents and trashed=false "
            "and name contains 'Tracking_Final_'"
        )
        results = service.files().list(
            q=query,
            pageSize=50,
            fields="files(id, name, createdTime)",
            orderBy="createdTime desc",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = results.get("files", [])
        return [
            (f"{f['createdTime'][:19].replace('T', ' ')} | {f['name']}", f["id"], f["name"])
            for f in files
        ]
    except Exception as e:
        st.warning(f"Gagal membaca folder GDrive: {e}")
        return []


def download_from_gdrive(file_id):
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception as e:
        st.warning(f"Gagal mengunduh file: {e}")
        return None


def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return None


# ---------------------------------------------------------------- Bootstrap

if "df_final" not in st.session_state:
    archives = list_gdrive_archives()
    if archives:
        label, file_id, fname = archives[0]
        buf = download_from_gdrive(file_id)
        if buf:
            st.session_state["df_final"] = pd.read_excel(buf)
            st.session_state["last_saved"] = fname


# ---------------------------------------------------------------- Header

c_left, c_right = st.columns(2)
logo_b64 = get_base64_image("Logo_PT_Geoservices_4K_Transparent.jpg")
with c_left:
    c_img, c_txt = st.columns([1, 6])
    with c_img:
        if logo_b64:
            st.markdown(
                f'<img src="data:image/jpeg;base64,{logo_b64}" '
                'style="width: 48px; border-radius: 4px; border: 1px solid #ccc;">',
                unsafe_allow_html=True,
            )
        else:
            st.markdown("🏢")
    with c_txt:
        st.markdown('<div class="main-title">Purchase Data Tracking</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="sub-meta">PT. Geoservices — Warehouse & Procurement Analytics</div>',
            unsafe_allow_html=True,
        )
with c_right:
    today_str = datetime.now().strftime("%A, %d %b %Y | %H:%M WIB")
    st.markdown(
        "<div style='text-align: right; background: #f8f9fa; padding: 8px 12px; "
        "border: 1px solid #e9ecef; border-radius: 4px; font-size: 0.78rem;'>"
        f"<b>Business Unit:</b> PT. Geoservices (7001)<br>📅 {today_str}</div>",
        unsafe_allow_html=True,
    )

st.write("")


# ---------------------------------------------------------------- ETL

@st.cache_data(show_spinner=False)
def process_tracking_data(pr_new, pr_old, po_lok, po_imp, inb):
    # Uploader bisa dibaca ulang pada rerun, jadi pointer-nya dikembalikan ke awal.
    for f in (pr_new, pr_old, po_lok, po_imp, inb):
        f.seek(0)

    pr_df = pd.read_excel(pr_new, header=1)
    pr_data = pr_df[['Date', 'PR Number \n(Manual)', 'Item \nCode', 'Item Description', 'Qty']].copy()
    pr_data.columns = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty']
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.replace(" ", "")
    pr_data['PR_Date'] = pd.to_datetime(pr_data['PR_Date'], errors='coerce')
    pr_data = pr_data[
        pr_data['PR_Manual_No'].notna()
        & (pr_data['PR_Manual_No'] != '')
        & (pr_data['PR_Manual_No'].astype(str) != 'nan')
    ]

    pr_2426_df = pd.read_excel(pr_old, sheet_name=0, header=None)
    closed_info = pr_2426_df.iloc[7:, [2, 6, 7]].copy()
    closed_info.columns = ['PR_Manual_No', 'RequestClosed', 'Item_Code']
    closed_info['Item_Code'] = closed_info['Item_Code'].astype(str).str.strip()
    closed_info['PR_Manual_No_Clean'] = closed_info['PR_Manual_No'].astype(str).str.replace(" ", "")
    closed_info = closed_info.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code'], keep='last')
    pr_data = pd.merge(
        pr_data,
        closed_info[['PR_Manual_No_Clean', 'Item_Code', 'RequestClosed']],
        on=['PR_Manual_No_Clean', 'Item_Code'],
        how='left',
    )
    pr_data['RequestClosed'] = pr_data['RequestClosed'].fillna('No')

    def clean_po(file, tipe):
        file.seek(0)
        po = pd.read_excel(file, header=13)
        po = po[['Purchase Order Number', 'PO Date', 'PR Manual No.', 'Item Code', 'Qty', 'Unnamed: 5']].copy()
        po.columns = ['PO_No', 'PO_Date', 'PR_Manual_No_Orig', 'Item_Code', 'PO_Qty', 'Vendor']
        po['PO_No'] = po['PO_No'].ffill()
        po['PO_Date'] = po['PO_Date'].ffill()
        po['Vendor'] = po['Vendor'].ffill()
        po = po.dropna(subset=['Item_Code'])
        po = po[po['Item_Code'].astype(str).str.strip() != 'nan']
        po['Tipe_PO'] = tipe
        return po

    po_df = pd.concat([clean_po(po_lok, "Lokal"), clean_po(po_imp, "Impor")], ignore_index=True)
    po_df['Item_Code'] = po_df['Item_Code'].astype(str).str.strip()

    expanded_rows = []
    for _, row in po_df.iterrows():
        pr_str = str(row['PR_Manual_No_Orig']).strip()
        if pr_str in ('', 'nan', 'None'):
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
                new_row['PR_Manual_No_Clean'] = (
                    (base_prefix + p).replace(" ", "")
                    if p.isdigit() and base_prefix
                    else p.replace(" ", "")
                )
                expanded_rows.append(new_row)
        else:
            new_row = row.to_dict()
            new_row['PR_Manual_No_Clean'] = pr_str.replace(" ", "")
            expanded_rows.append(new_row)

    po_exp = pd.DataFrame(expanded_rows)

    if po_exp.empty:
        po_agg = pd.DataFrame(
            columns=['PR_Manual_No_Clean', 'Item_Code', 'PO_No', 'PO_Date', 'PO_Qty', 'Vendor', 'Tipe_PO']
        )
    else:
        po_exp['PO_Date'] = pd.to_datetime(po_exp['PO_Date'], errors='coerce')
        po_exp['PO_Qty'] = pd.to_numeric(po_exp['PO_Qty'], errors='coerce').fillna(0)
        po_agg = po_exp.groupby(['PR_Manual_No_Clean', 'Item_Code']).agg(
            PO_No=('PO_No', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique())),
            PO_Date=('PO_Date', 'max'),
            PO_Qty=('PO_Qty', 'sum'),
            Vendor=('Vendor', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique())),
            Tipe_PO=('Tipe_PO', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique())),
        ).reset_index()

    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

    inb.seek(0)
    inb_xls = pd.ExcelFile(inb)
    inb_df = pd.concat(
        [pd.read_excel(inb_xls, sheet_name=s) for s in inb_xls.sheet_names],
        ignore_index=True,
    )
    inb_df['ItemCode'] = inb_df['ItemCode'].astype(str).str.strip()
    inb_df['POnumber'] = inb_df['POnumber'].astype(str).str.strip()
    inb_df['ReceivedDate'] = pd.to_datetime(inb_df['ReceivedDate'], errors='coerce')
    inb_df['RcvQty'] = pd.to_numeric(inb_df['RcvQty'], errors='coerce').fillna(0)
    inb_agg = inb_df.groupby(['POnumber', 'ItemCode']).agg(
        Rcv_Date=('ReceivedDate', 'max'), Rcv_Qty=('RcvQty', 'sum')
    ).reset_index()

    def get_inb(po_str, item_code):
        empty = pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})
        if not po_str:
            return empty
        po_list = [p.strip() for p in str(po_str).split(',') if p.strip()]
        subset = inb_agg[inb_agg['POnumber'].isin(po_list) & (inb_agg['ItemCode'] == item_code)]
        if subset.empty:
            return empty
        return pd.Series({'Rcv_Date': subset['Rcv_Date'].max(), 'Rcv_Qty': subset['Rcv_Qty'].sum()})

    merged[['Rcv_Date', 'Rcv_Qty']] = merged.apply(
        lambda row: get_inb(row['PO_No'], row['Item_Code']), axis=1
    )

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
    merged['Qty_Outstanding'] = merged['PO_Qty'] - merged['Rcv_Qty']

    for col in ('PR_Date', 'PO_Date', 'Rcv_Date'):
        merged[col] = pd.to_datetime(merged[col], errors='coerce').dt.strftime('%d/%m/%Y')
        merged[col] = merged[col].fillna('-')

    merged['Vendor'] = merged['Vendor'].fillna('-')
    if 'Tipe_PO' not in merged.columns:
        merged['Tipe_PO'] = '-'
    merged['Tipe_PO'] = merged['Tipe_PO'].fillna('-')

    return merged[[
        'PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty',
        'PO_Date', 'PO_No', 'Vendor', 'Tipe_PO', 'PO_Qty',
        'Rcv_Date', 'Rcv_Qty', 'Qty_Outstanding', 'Status',
    ]]


def to_excel_bytes(df):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------- UI

selected_tab = st.radio(
    "Navigation",
    ["📊 Dashboard", "⚙️ Proses Data", "📥 Download / Arsip"],
    horizontal=True,
    label_visibility="collapsed",
)
st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

if selected_tab == "📊 Dashboard":
    st.subheader("Executive Dashboard")
    if 'df_final' in st.session_state:
        df_final = st.session_state['df_final']
        if 'last_saved' in st.session_state:
            st.caption(f"📁 Active Dataset (GDrive): `{st.session_state['last_saved']}`")

        c1, c2, c3, c4 = st.columns(4)
        sc = df_final['Status'].value_counts()
        c1.metric("Routing Approval", int(sc.get("Routing Approval", 0)))
        c2.metric("Menunggu Pengiriman", int(sc.get("Menunggu Pengiriman", 0)))
        c3.metric("Diterima Sebagian", int(sc.get("Diterima Sebagian", 0)))
        c4.metric("Sudah Diterima", int(sc.get("Sudah Diterima", 0)))

        ca1, ca2 = st.columns(2)
        with ca1:
            st.plotly_chart(
                px.pie(df_final, names='Status', title='Status Proporsi'),
                use_container_width=True,
            )
        with ca2:
            tv = df_final[df_final['Vendor'] != '-']['Vendor'].value_counts().head(10).reset_index()
            tv.columns = ['Vendor', 'Jumlah']
            if tv.empty:
                st.info("Belum ada data vendor.")
            else:
                st.plotly_chart(
                    px.bar(tv, x='Jumlah', y='Vendor', orientation='h', title='Top Vendor'),
                    use_container_width=True,
                )
        st.dataframe(df_final.head(100), use_container_width=True)
    else:
        st.warning("Belum ada data.")

elif selected_tab == "⚙️ Proses Data":
    st.subheader("Processing Center (Full ETL)")
    u1, u2 = st.columns(2)
    with u1:
        pr_b = st.file_uploader("PR Data (Base)", type=['xlsx'])
        pr_c = st.file_uploader("PRN Data (Closed)", type=['xlsx'])
        po_l = st.file_uploader("PO Data - Lokal", type=['xlsx'])
    with u2:
        po_i = st.file_uploader("PO Data - Impor", type=['xlsx'])
        in_b = st.file_uploader("Inbound Data", type=['xlsx'])

    if st.button("🚀 Proses & Sync GDrive", type="primary"):
        if all([pr_b, pr_c, po_l, po_i, in_b]):
            with st.spinner("Memproses..."):
                try:
                    df_out = process_tracking_data(pr_b, pr_c, po_l, po_i, in_b)
                except Exception as e:
                    st.error(f"Gagal memproses data: {e}")
                    st.stop()

                buf = to_excel_bytes(df_out)
                fname = f"Tracking_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ok, res = upload_to_gdrive(buf, fname)

                st.session_state['df_final'] = df_out
                if ok:
                    st.session_state['last_saved'] = fname
                    st.success("Tersimpan permanen di GDrive!")
                else:
                    st.error(f"Gagal upload ke GDrive: {res}")
                    st.info("Data tetap tersimpan di sesi ini. Silakan unduh manual di bawah.")
                    st.download_button(
                        "📥 Download hasil (.xlsx)",
                        to_excel_bytes(df_out).getvalue(),
                        fname,
                        mime=XLSX_MIME,
                    )
        else:
            st.error("Upload 5 file lengkap dulu.")

    st.divider()
    st.subheader("Filter Outstanding by Item Code")
    item_file = st.file_uploader("Upload Excel Item Code", type=['xlsx'], key="item_code_file")
    if st.button("🔍 Generate Outstanding"):
        if not item_file:
            st.error("Upload file item code dulu.")
        elif 'df_final' not in st.session_state:
            st.error("Belum ada data hasil proses. Jalankan ETL dulu.")
        else:
            base_df = st.session_state['df_final'].copy()
            up_df = pd.read_excel(item_file)
            target_col = [c for c in up_df.columns if 'item' in str(c).lower() and 'code' in str(c).lower()]
            col_name = target_col[0] if target_col else up_df.columns[0]
            codes = up_df[col_name].dropna().astype(str).str.strip().unique()

            base_df['Code_Clean'] = base_df['Item_Code'].astype(str).str.strip()
            out_df = base_df[
                base_df['Code_Clean'].isin(codes) & (base_df['Status'] != 'Sudah Diterima')
            ].drop(columns=['Code_Clean'])

            if out_df.empty:
                st.info("Tidak ada item outstanding yang cocok.")
            else:
                st.dataframe(out_df, use_container_width=True)
                st.download_button(
                    "📥 Download Outstanding (.xlsx)",
                    to_excel_bytes(out_df).getvalue(),
                    f"Outstanding_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime=XLSX_MIME,
                )

elif selected_tab == "📥 Download / Arsip":
    st.subheader("Cloud Repository (GDrive)")
    archives = list_gdrive_archives()
    if archives:
        labels = [a[0] for a in archives]
        selected_lbl = st.selectbox("Select Cloud Archived Document:", labels)
        idx = labels.index(selected_lbl)
        sel_id, sel_name = archives[idx][1], archives[idx][2]

        buf = download_from_gdrive(sel_id)
        if buf:
            df_prev = pd.read_excel(buf)
            st.dataframe(df_prev.head(50), use_container_width=True)
            buf.seek(0)
            st.download_button("📥 Download File Ini", buf.getvalue(), sel_name, mime=XLSX_MIME)
    else:
        st.warning("Belum ada arsip di GDrive.")
