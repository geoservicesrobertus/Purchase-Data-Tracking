import streamlit as st
import pandas as pd
import numpy as np
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

CREDS_DICT = {
  "type": "service_account",
  "project_id": "purchase-data-tracking",
  "private_key_id": "1b9b383c0665eb93d863f9b786c561c82063fe13",
  "private_key": (
      "-----BEGIN PRIVATE KEY-----\n"
      "MIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQDU/wp2SztkqhXW\n"
      "dm0AD+jTvXCm4D3ZAYW9v6KHBXWS15naqFI1bww3LnF07VVeOnkKlFIqVqFcR5qN\n"
      "+zsxBsDgrm3t3PFMKCMrDpDk3hfS3Hw3cPn8APnYs/EMtWRNJr+Ky9ZcqOT1De4H\n"
      "PlvCG0bIGqWG3tIqx8DsYcSGw9f0ruRUaXrvUF9YbKOPr9uEgTszkr60iZtFA5Bu\n"
      "2VOqBseMuQDswbqh4ZSqRDaLvjux91ijKYhHPVuEpUCB3LaiHPZEs/5YM2EF8SyD\n"
      "GW6lyA5HTwZYBwQ6ScrtJj5cyGYKKbIs3g2oIXvo2oc/YIxQjSKr3/MB02KCN4M9\n"
      "1MwjvbDbAgMBAAECggEAHyqAZ+XDN4wRrPNgKKmrSkxGbwyR0C6CWMzxJaudVBK7\n"
      "Hv0RJmNs2KgxjsfpfWO18V/Zk/tgGgYaLdtMgnR4BYhQaoUxQ5D98F9twSdkOgTs\n"
      "IhqkYYTtChHuXAswtX9NeKwx3hYShm723NV5jLH8DeykOtNg0kSvTIdTv9ppP5KR\n"
      "Lev4Th5Nhoj20g+1dc6GGZlGSFDjmENwVM96FRswfWMPttOi7/Q3GADv07Ts8QWk\n"
      "Gx6tO9c+MgJWXOGbvC6XBWke8tpDI9BEvUTheFs2cE8557FFvmDCjN3WFrFVSf5w\n"
      "u74t1rN2OFed8icyx50zI1zVdT/H2jpWHip3GIx75QKBgQD7V/A6Q6MiFyNMfofP\n"
      "U2Ky0c33ZkAzYhFWH2XB0eKaRt/xK2EXNGdkghCSISuOzwL+HDst9OQPF7mEmJjD\n"
      "uwMcBd7NLI16RwA36I9H5HecVAUIDdpS0nk711swSGeRmJ4IPBTs38otjIJ17A8m\n"
      "1FU8DIre+m+UCEBZp1VJT+2X9QKBgQDY8TsO/LxfOtb3LL7unMi8FqnHYiurQ0b/\n"
      "erebYsduzySzmq+SLgD5aUx8vg66IQwAs63cCVKp5PHPjbMKYylN0z6hkvtr+DTQ\n"
      "oUbzlg/i+uSYwRK/DnM+1X/M8N/0wAnjoBzzbfREPCBfVkAqgXbzJB5KLeEB5+Fz\n"
      "EXT/zcAzjwKBgQD2K38B0dUpQng0J4lkqkr00UBlmyQuL1LDgyTq3GKQr/IOB2qk\n"
      "i5Logesw9IPw7xgDQitEK6Jild4B3GNi8PtuquE5GvXGWVwBZilPRJlR54i2Brta\n"
      "ewJ6dca+V2v40f2WGyJzjgw66G+uh3Gfmj+Q/MfW9HnsBtjf9mA12a7fMQKBgQC6\n"
      "0AhOYJ8J1k5UrSiBq2tEZLOw6T23jgiuaYuAeDBKoH/3VaYI2Cqom99sr/FYoKqI\n"
      "VDHMAA86E9eTJm9d64Qe62DMnBh7olJAshC6I6fsiqadT+2HrrbZDdqurWH9jf02\n"
      "EaO8kBu/QpOR5WD9+VxoBds7f4R6Mqa2gvrgaNowywKBgQCinsuBMjoq8DZ5f7eO\n"
      "VpUWBPMr25XiaDUDx5ru/UP95GqSyB36YLLufC1CyJlriAPJqK93dmv/2pVADHTz\n"
      "vpeIl8gKZKsKXOC7llzcv0gF7k9LG6VDr4tjuSuULIDlm5aes2qYvVzx1atVLqxj\n"
      "eLz0JSdLNyxaBkVYgARpioi21A==\n"
      "-----END PRIVATE KEY-----\n"
  ),
  "client_email": "robertus-yuseno@purchase-data-tracking.iam.gserviceaccount.com",
  "client_id": "109589672772004723099",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/robertus-yuseno%40purchase-data-tracking.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

def get_gdrive_service():
    creds = service_account.Credentials.from_service_account_info(CREDS_DICT, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

def upload_to_gdrive(file_buffer, filename):
    try:
        service = get_gdrive_service()
        media = MediaIoBaseUpload(file_buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        file = service.files().create(body={'name': filename, 'parents': [GDRIVE_FOLDER_ID]}, media_body=media, fields='id').execute()
        return True, file.get('id')
    except Exception as e:
        return False, str(e)

def list_gdrive_archives():
    try:
        service = get_gdrive_service()
        query = f"'{GDRIVE_FOLDER_ID}' in parents and trashed=false and name contains 'Tracking_Final_'"
        results = service.files().list(q=query, pageSize=50, fields="files(id, name, createdTime)", orderBy="createdTime desc").execute()
        files = results.get('files', [])
        return [(f['createdTime'][:19].replace('T', ' ') + ' | ' + f['name'], f['id'], f['name']) for f in files]
    except Exception:
        return []

def download_from_gdrive(file_id):
    try:
        service = get_gdrive_service()
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception:
        return None

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode('utf-8')
    return None

if 'df_final' not in st.session_state:
    archives = list_gdrive_archives()
    if archives:
        buf = download_from_gdrive(archives[0])
        if buf:
            st.session_state['df_final'] = pd.read_excel(buf)
            st.session_state['last_saved'] = archives[0]

c_left, c_right = st.columns()
logo_b64 = get_base64_image("Logo_PT_Geoservices_4K_Transparent.jpg")
with c_left:
    c_img, c_txt = st.columns()
    with c_img:
        if logo_b64:
            st.markdown(f'<img src="data:image/jpeg;base64,{logo_b64}" style="width: 48px; border-radius: 4px; border: 1px solid #ccc;">', unsafe_allow_html=True)
        else:
            st.markdown("🏢")
    with c_txt:
        st.markdown('<div class="main-title">Purchase Data Tracking</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-meta">PT. Geoservices — Warehouse & Procurement Analytics</div>', unsafe_allow_html=True)
with c_right:
    today_str = datetime.now().strftime("%A, %d %b %Y | %H:%M WIB")
    st.markdown(f"<div style='text-align: right; background: #f8f9fa; padding: 8px 12px; border: 1px solid #e9ecef; border-radius: 4px; font-size: 0.78rem;'><b>Business Unit:</b> PT. Geoservices (7001)<br>📅 {today_str}</div>", unsafe_allow_html=True)

st.write("")

@st.cache_data
def process_tracking_data(pr_new, pr_old, po_lok, po_imp, inb):
    pr_df = pd.read_excel(pr_new, header=1)
    pr_data = pr_df[['Date', 'PR Number \n(Manual)', 'Item \nCode', 'Item Description', 'Qty']].copy()
    pr_data.columns = ['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty']
    pr_data['Item_Code'] = pr_data['Item_Code'].astype(str).str.strip()
    pr_data['PR_Manual_No_Clean'] = pr_data['PR_Manual_No'].astype(str).str.replace(" ", "")
    pr_data['PR_Date'] = pd.to_datetime(pr_data['PR_Date'], errors='coerce')
    pr_data = pr_data[pr_data['PR_Manual_No'].notna() & (pr_data['PR_Manual_No'] != '') & (pr_data['PR_Manual_No'].astype(str) != 'nan')]

    pr_2426_df = pd.read_excel(pr_old, sheet_name=0, header=None)
    closed_info = [pr_2426_df.iloc].copy()
    closed_info.columns = ['PR_Manual_No', 'RequestClosed', 'Item_Code']
    closed_info['Item_Code'] = closed_info['Item_Code'].astype(str).str.strip()
    closed_info['PR_Manual_No_Clean'] = closed_info['PR_Manual_No'].astype(str).str.replace(" ", "")
    closed_info = closed_info.drop_duplicates(subset=['PR_Manual_No_Clean', 'Item_Code'], keep='last')
    pr_data = pd.merge(pr_data, closed_info[['PR_Manual_No_Clean', 'Item_Code', 'RequestClosed']], on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    pr_data['RequestClosed'] = pr_data['RequestClosed'].fillna('No')

    def clean_po(file, tipe):
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
        if pd.isna(pr_str) or pr_str == 'nan': continue
        if ',' in pr_str or '/' in pr_str:
            parts = re.split(r'[,/]', pr_str)
            base_prefix = ""
            for p in [x.strip() for x in parts if x.strip()]:
                if not p.isdigit():
                    m = re.match(r'([A-Za-z.\-]+)(\d+)', p)
                    if m: base_prefix = m.group(1)
                new_row = row.to_dict()
                new_row['PR_Manual_No_Clean'] = (base_prefix + p).replace(" ", "") if p.isdigit() and base_prefix else p.replace(" ", "")
                expanded_rows.append(new_row)
        else:
            new_row = row.to_dict()
            new_row['PR_Manual_No_Clean'] = pr_str.replace(" ", "")
            expanded_rows.append(new_row)

    po_exp = pd.DataFrame(expanded_rows)
    po_exp['PO_Date'] = pd.to_datetime(po_exp['PO_Date'], errors='coerce')
    po_exp['PO_Qty'] = pd.to_numeric(po_exp['PO_Qty'], errors='coerce').fillna(0)
    po_agg = po_exp.groupby(['PR_Manual_No_Clean', 'Item_Code']).agg(
        PO_No=('PO_No', lambda x: ', '.join(x.dropna().unique().astype(str))),
        PO_Date=('PO_Date', 'max'),
        PO_Qty=('PO_Qty', 'sum'),
        Vendor=('Vendor', lambda x: ', '.join(x.dropna().unique().astype(str))),
        Tipe_PO=('Tipe_PO', lambda x: ', '.join(x.dropna().unique().astype(str)))
    ).reset_index()

    merged = pd.merge(pr_data, po_agg, on=['PR_Manual_No_Clean', 'Item_Code'], how='left')
    merged['PO_No'] = merged['PO_No'].fillna('')
    merged = merged[~((merged['PO_No'] == '') & (merged['RequestClosed'] == 'Yes'))]

    inb_xls = pd.ExcelFile(inb)
    inb_df = pd.concat([pd.read_excel(inb_xls, sheet_name=s) for s in inb_xls.sheet_names])
    inb_df['ItemCode'] = inb_df['ItemCode'].astype(str).str.strip()
    inb_df['POnumber'] = inb_df['POnumber'].astype(str).str.strip()
    inb_agg = inb_df.groupby(['POnumber', 'ItemCode']).agg(Rcv_Date=('ReceivedDate', 'max'), Rcv_Qty=('RcvQty', 'sum')).reset_index()

    def get_inb(po_str, item_code):
        if pd.isna(po_str) or po_str == '': return pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})
        subset = inb_agg[(inb_agg['POnumber'].isin([p.strip() for p in po_str.split(',')])) & (inb_agg['ItemCode'] == item_code)]
        return pd.Series({'Rcv_Date': subset['Rcv_Date'].max(), 'Rcv_Qty': subset['Rcv_Qty'].sum()}) if not subset.empty else pd.Series({'Rcv_Date': pd.NaT, 'Rcv_Qty': 0})

    merged[['Rcv_Date', 'Rcv_Qty']] = merged.apply(lambda row: get_inb(row['PO_No'], row['Item_Code']), axis=1)

    def get_status(row):
        if row['PO_No'] == '': return 'Routing Approval'
        elif row['Rcv_Qty'] == 0: return 'Menunggu Pengiriman'
        elif row['Rcv_Qty'] < row['PO_Qty']: return 'Diterima Sebagian'
        return 'Sudah Diterima'

    merged['Status'] = merged.apply(get_status, axis=1)
    merged['PO_Qty'] = pd.to_numeric(merged['PO_Qty'], errors='coerce').fillna(0)
    merged['Rcv_Qty'] = pd.to_numeric(merged['Rcv_Qty'], errors='coerce').fillna(0)
    merged['Qty_Outstanding'] = merged['PO_Qty'] - merged['Rcv_Qty']
    
    merged['PR_Date'] = merged['PR_Date'].dt.strftime('%m/%d/%Y').fillna('-')
    merged['PO_Date'] = merged['PO_Date'].dt.strftime('%m/%d/%Y').fillna('-')
    merged['Rcv_Date'] = pd.to_datetime(merged['Rcv_Date'], errors='coerce').dt.strftime('%m/%d/%Y').fillna('-')
    merged['Vendor'] = merged['Vendor'].fillna('-')
    merged['Tipe_PO'] = merged['Tipe_PO'].fillna('-')
    
    return merged[['PR_Date', 'PR_Manual_No', 'Item_Code', 'Item_Name', 'PR_Qty', 'PO_Date', 'PO_No', 'Vendor', 'Tipe_PO', 'PO_Qty', 'Rcv_Date', 'Rcv_Qty', 'Qty_Outstanding', 'Status']]

selected_tab = st.radio("Navigation", ["📊 Dashboard", "⚙️ Proses Data", "📥 Download / Arsip"], horizontal=True, label_visibility="collapsed")
st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

if selected_tab == "📊 Dashboard":
    st.subheader("Executive Dashboard")
    if 'df_final' in st.session_state:
        df_final = st.session_state['df_final']
        if 'last_saved' in st.session_state:
            st.caption(f"📁 Active Dataset (GDrive): `{st.session_state['last_saved']}`")
        c1, c2, c3, c4 = st.columns(4)
        sc = df_final['Status'].value_counts()
        c1.metric("Routing Approval", sc.get("Routing Approval", 0))
        c2.metric("Menunggu Pengiriman", sc.get("Menunggu Pengiriman", 0))
        c3.metric("Diterima Sebagian", sc.get("Diterima Sebagian", 0))
        c4.metric("Sudah Diterima", sc.get("Sudah Diterima", 0))
        
        ca1, ca2 = st.columns(2)
        with ca1: st.plotly_chart(px.pie(df_final, names='Status', title='Status Proporsi'), use_container_width=True)
        with ca2: 
            tv = df_final[df_final['Vendor'] != '-']['Vendor'].value_counts().head(10).reset_index()
            tv.columns = ['Vendor', 'Jumlah']
            st.plotly_chart(px.bar(tv, x='Jumlah', y='Vendor', orientation='h', title='Top Vendor'), use_container_width=True)
        st.dataframe(df_final.head(100), use_container_width=True)
    else: st.warning("Belum ada data.")

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
                df_out = process_tracking_data(pr_b, pr_c, po_l, po_i, in_b)
                buf = io.BytesIO()
                df_out.to_excel(buf, index=False)
                buf.seek(0)
                fname = f"Tracking_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                ok, res = upload_to_gdrive(buf, fname)
                if ok:
                    st.success("Tersimpan permanen di GDrive!")
                    st.session_state['df_final'] = df_out
                    st.session_state['last_saved'] = fname
                else: st.error(f"Gagal GDrive: {res}")
        else: st.error("Upload 5 file lengkap dulu.")

    st.divider()
    st.subheader("Filter Outstanding by Item Code")
    item_file = st.file_uploader("Upload Excel Item Code", type=['xlsx'])
    if st.button("🔍 Generate Outstanding"):
        if item_file and 'df_final' in st.session_state:
            base_df = st.session_state['df_final'].copy()
            up_df = pd.read_excel(item_file)
            target_col = [c for c in up_df.columns if 'item' in c.lower() and 'code' in c.lower()]
            col_name = target_col[0] if target_col else up_df.columns[0]
            codes = up_df[col_name].dropna().astype(str).str.strip().unique()
            
            base_df['Code_Clean'] = base_df['Item_Code'].astype(str).str.strip()
            out_df = base_df[base_df['Code_Clean'].isin(codes) & (base_df['Status'] != 'Sudah Diterima')].drop(columns=['Code_Clean'])
            st.dataframe(out_df, use_container_width=True)
            
            out_buf = io.BytesIO()
            out_df.to_excel(out_buf, index=False)
            out_buf.seek(0)
            st.download_button("📥 Download Outstanding (.xlsx)", out_buf.getvalue(), f"Outstanding_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif selected_tab == "📥 Download / Arsip":
    st.subheader("Cloud Repository (GDrive)")
    archives = list_gdrive_archives()
    if archives:
        labels = [a[0] for a in archives]
        ids = [a[1] for a in archives]
        selected_lbl = st.selectbox("Select Cloud Archived Document:", labels)
        sel_id = ids[labels.index(selected_lbl)]
        buf = download_from_gdrive(sel_id)
        if buf:
            df_prev = pd.read_excel(buf)
            st.dataframe(df_prev.head(50), use_container_width=True)
            buf.seek(0)
            target_fname = selected_lbl.split('|').strip() if '|' in selected_lbl else "archive.xlsx"
            st.download_button("📥 Download File Ini", buf.getvalue(), target_fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else: st.warning("Belum ada arsip di GDrive.")
