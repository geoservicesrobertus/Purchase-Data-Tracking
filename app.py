import base64
from datetime import datetime, timezone
from io import BytesIO
import os
import re
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from zoneinfo import ZoneInfo

# ==========================================
# 0. CONFIG & ENTERPRISE THEME SETUP
# ==========================================
st.set_page_config(
    page_title="Purchase Data Tracking - PT. Geoservices",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        
        .main-title {font-size: 1.5rem; font-weight: 700; color: #1e3d59; margin: 0;}
        .sub-meta {font-size: 0.8rem; color: #5f6c7b;}
        
        .erp-panel {
            background: #ffffff;
            border: 1px solid #c0c0c0;
            border-radius: 3px;
            padding: 15px;
            margin-bottom: 15px;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Folder penyimpanan lokal permanen
OUTPUT_FOLDER = r"D:\User Data - WarehouseSPV\Documents\Purchase Data Tracking"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
WIB = ZoneInfo("Asia/Jakarta")


def get_base64_image(image_path):
  if os.path.exists(image_path):
    with open(image_path, "rb") as img_file:
      return base64.b64encode(img_file.read()).decode("utf-8")
  return None


def get_latest_archive_filepath():
  files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]
  if files:
    files.sort(
        key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)),
        reverse=True,
    )
    return os.path.join(OUTPUT_FOLDER, files[0]), files[0]
  return None, None


def auto_load_latest_dashboard():
  if "df_final" not in st.session_state:
    latest_path, latest_name = get_latest_archive_filepath()
    if latest_path:
      try:
        st.session_state["df_final"] = pd.read_excel(latest_path)
        st.session_state["last_saved"] = latest_name
      except Exception:
        pass


auto_load_latest_dashboard()

c_left, c_right = st.columns([3, 1])
logo_base64 = get_base64_image("Logo_PT_Geoservices_4K_Transparent.jpg")

with c_left:
  c_img, c_txt = st.columns([1, 6])
  with c_img:
    if logo_base64:
      st.markdown(
          f'<img src="data:image/jpeg;base64,{logo_base64}" style="width:'
          ' 48px; border-radius: 4px; border: 1px solid #ccc;">',
          unsafe_allow_html=True,
      )
    else:
      st.markdown("🏢", unsafe_allow_html=True)
  with c_txt:
    st.markdown(
        '<div class="main-title">Purchase Data Tracking</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="sub-meta">PT. Geoservices — Warehouse & Procurement'
        " Analytics</div>",
        unsafe_allow_html=True,
    )

with c_right:
  today_str = datetime.now(WIB).strftime("%A, %d %b %Y | %H:%M WIB")
  st.markdown(
      f"""
        <div style='text-align: right; background: #f8f9fa; padding: 8px 12px; border: 1px solid #e9ecef; border-radius: 4px; font-size: 0.78rem; color: #333;'>
            <b>Business Unit:</b> PT. Geoservices (7001)<br>
            <span style='color: #666;'>📅 {today_str}</span>
        </div>
    """,
      unsafe_allow_html=True,
  )

st.write("")


# ==========================================
# CORE PROCESSING ENGINE
# ==========================================
@st.cache_data
def process_tracking_data(pr_new, pr_old, po_lok, po_imp, inb):
  pr_df = pd.read_excel(pr_new, header=1)
  pr_data = pr_df[
      ["Date", "PR Number \n(Manual)", "Item \nCode", "Item Description", "Qty"]
  ].copy()
  pr_data.columns = [
      "PR_Date",
      "PR_Manual_No",
      "Item_Code",
      "Item_Name",
      "PR_Qty",
  ]
  
  # 🔥 Filter PR Qty = 0 di awal sebelum tracking dimulai
  pr_data["PR_Qty"] = pd.to_numeric(pr_data["PR_Qty"], errors="coerce").fillna(0)
  pr_data = pr_data[pr_data["PR_Qty"] > 0].copy()

  pr_data["Item_Code"] = pr_data["Item_Code"].astype(str).str.strip()
  pr_data["PR_Manual_No_Clean"] = (
      pr_data["PR_Manual_No"].astype(str).str.replace(" ", "")
  )
  pr_data["PR_Date"] = pd.to_datetime(pr_data["PR_Date"], errors="coerce")
  pr_data = pr_data[
      pr_data["PR_Manual_No"].notna()
      & (pr_data["PR_Manual_No"] != "")
      & (pr_data["PR_Manual_No"].astype(str) != "nan")
  ]

  pr_2426_df = pd.read_excel(pr_old, sheet_name=0, header=None)
  closed_info = pr_2426_df.iloc[7:, [2, 6, 7]].copy()
  closed_info.columns = ["PR_Manual_No", "RequestClosed", "Item_Code"]
  closed_info["Item_Code"] = closed_info["Item_Code"].astype(str).str.strip()
  closed_info["PR_Manual_No_Clean"] = (
      closed_info["PR_Manual_No"].astype(str).str.replace(" ", "")
  )
  closed_info = closed_info.drop_duplicates(
      subset=["PR_Manual_No_Clean", "Item_Code"], keep="last"
  )

  pr_data = pd.merge(
      pr_data,
      closed_info[["PR_Manual_No_Clean", "Item_Code", "RequestClosed"]],
      on=["PR_Manual_No_Clean", "Item_Code"],
      how="left",
  )
  pr_data["RequestClosed"] = pr_data["RequestClosed"].fillna("No")

  def clean_po(file, tipe):
    po_raw = pd.read_excel(file, header=13)
    
    # 1. Ffill KHUSUS nomor PO saja dulu, supaya ketahuan tiap baris item punya siapa
    if "Purchase Order Number" in po_raw.columns:
        po_raw["Purchase Order Number"] = po_raw["Purchase Order Number"].ffill()
    
    # 2. Cari nomor PO yang harus dibuang (REJECT / DECLINE / CLOSED)
    rejected_po_list = set()
    for col in po_raw.columns:
        is_status_col = "status" in str(col).lower()
        
        # Kalau kolomnya berkaitan dengan status (misal: PO_Status), buang yg CLOSED & REJECT
        if is_status_col:
            mask_rejected = po_raw[col].astype(str).str.upper().str.contains("REJECT|DECLINE|CLOSED", na=False)
        else:
            # Kolom biasa cuma buang yg REJECT/DECLINE aja biar gak salah sasaran
            mask_rejected = po_raw[col].astype(str).str.upper().str.contains("REJECT|DECLINE", na=False)
            
        if "Purchase Order Number" in po_raw.columns:
            bad_pos = po_raw.loc[mask_rejected, "Purchase Order Number"].dropna().unique()
            rejected_po_list.update(bad_pos)
            
    # 3. Buang keseluruhan blok baris dari PO yang bermasalah tersebut
    if "Purchase Order Number" in po_raw.columns:
        po_safe = po_raw[~po_raw["Purchase Order Number"].isin(rejected_po_list)].copy()
    else:
        po_safe = po_raw.copy()
        
    # 4. AMAN melakukan Ffill untuk Date, Vendor, dll
    po_filled = po_safe.ffill()

    # 5. Ekstrak kolom standar
    cols_to_pull = [
        "Purchase Order Number",
        "PO Date",
        "PR Manual No.",
        "Item Code",
        "Qty",
        "Unnamed: 5", 
    ]
    cols_available = [c for c in cols_to_pull if c in po_filled.columns]
    po = po_filled[cols_available].copy()

    rename_map = {}
    standard_names = [
        "PO_No",
        "PO_Date",
        "PR_Manual_No_Orig",
        "Item_Code",
        "PO_Qty",
        "Vendor",
    ]
    for i, col_name in enumerate(cols_available):
      if i < len(standard_names):
        rename_map[col_name] = standard_names[i]

    po.rename(columns=rename_map, inplace=True)
    
    if "PO_No" not in po.columns: po["PO_No"] = ""
    if "PO_Date" not in po.columns: po["PO_Date"] = pd.NaT
    if "Vendor" not in po.columns: po["Vendor"] = "-"
    if "PO_Qty" not in po.columns: po["PO_Qty"] = 0

    po = po.dropna(subset=["Item_Code"])
    po = po[po["Item_Code"].astype(str).str.strip() != "nan"]
    po["Tipe_PO"] = tipe
    return po

  po_df = pd.concat(
      [clean_po(po_lok, "Lokal"), clean_po(po_imp, "Impor")], ignore_index=True
  )
  po_df["Item_Code"] = po_df["Item_Code"].astype(str).str.strip()

  expanded_rows = []
  for _, row in po_df.iterrows():
    pr_str = str(row["PR_Manual_No_Orig"]).strip()
    if pd.isna(pr_str) or pr_str == "nan":
      continue
    if "," in pr_str or "/" in pr_str:
      parts = re.split(r"[,/]", pr_str)
      base_prefix = ""
      for p in [x.strip() for x in parts if x.strip()]:
        if not p.isdigit():
          match = re.match(r"([A-Za-z.\-]+)(\d+)", p)
          if match:
            base_prefix = match.group(1)
        new_row = row.to_dict()
        new_row["PR_Manual_No_Clean"] = (
            (base_prefix + p).replace(" ", "")
            if p.isdigit() and base_prefix
            else p.replace(" ", "")
        )
        expanded_rows.append(new_row)
    else:
      new_row = row.to_dict()
      new_row["PR_Manual_No_Clean"] = pr_str.replace(" ", "")
      expanded_rows.append(new_row)

  po_exp = pd.DataFrame(expanded_rows)
  po_exp["PO_Date"] = pd.to_datetime(po_exp["PO_Date"], errors="coerce")
  po_exp["PO_Qty"] = pd.to_numeric(po_exp["PO_Qty"], errors="coerce").fillna(0)

  po_agg = (
      po_exp.groupby(["PR_Manual_No_Clean", "Item_Code"])
      .agg(
          PO_No=("PO_No", lambda x: ", ".join(x.dropna().unique().astype(str))),
          PO_Date=("PO_Date", "max"),
          PO_Qty=("PO_Qty", "sum"),
          Vendor=(
              "Vendor",
              lambda x: ", ".join(x.dropna().unique().astype(str)),
          ),
          Tipe_PO=(
              "Tipe_PO",
              lambda x: ", ".join(x.dropna().unique().astype(str)),
          ),
      )
      .reset_index()
  )

  merged = pd.merge(
      pr_data, po_agg, on=["PR_Manual_No_Clean", "Item_Code"], how="left"
  )
  merged["PO_No"] = merged["PO_No"].fillna("")
  merged = merged[
      ~((merged["PO_No"] == "") & (merged["RequestClosed"] == "Yes"))
  ]

  inb_xls = pd.ExcelFile(inb)
  inb_df = pd.concat(
      [pd.read_excel(inb_xls, sheet_name=s) for s in inb_xls.sheet_names]
  )
  inb_df["ItemCode"] = inb_df["ItemCode"].astype(str).str.strip()
  inb_df["POnumber"] = inb_df["POnumber"].astype(str).str.strip()
  inb_agg = (
      inb_df.groupby(["POnumber", "ItemCode"])
      .agg(Rcv_Date=("ReceivedDate", "max"), Rcv_Qty=("RcvQty", "sum"))
      .reset_index()
  )

  def get_inb(po_str, item_code):
    if pd.isna(po_str) or po_str == "":
      return pd.Series({"Rcv_Date": pd.NaT, "Rcv_Qty": 0})
    pos = [p.strip() for p in po_str.split(",")]
    subset = inb_agg[
        (inb_agg["POnumber"].isin(pos)) & (inb_agg["ItemCode"] == item_code)
    ]
    if subset.empty:
      return pd.Series({"Rcv_Date": pd.NaT, "Rcv_Qty": 0})
    return pd.Series(
        {"Rcv_Date": subset["Rcv_Date"].max(), "Rcv_Qty": subset["Rcv_Qty"].sum()}
    )

  merged[["Rcv_Date", "Rcv_Qty"]] = merged.apply(
      lambda row: get_inb(row["PO_No"], row["Item_Code"]), axis=1
  )

  def get_status(row):
    if row["PO_No"] == "":
      return "Routing Approval"
    elif row["Rcv_Qty"] == 0:
      return "Menunggu Pengiriman"
    elif row["Rcv_Qty"] < row["PO_Qty"]:
      return "Diterima Sebagian"
    else:
      return "Sudah Diterima"

  merged["Status"] = merged.apply(get_status, axis=1)

  merged["PO_Qty"] = pd.to_numeric(merged["PO_Qty"], errors="coerce").fillna(0)
  merged["Rcv_Qty"] = pd.to_numeric(merged["Rcv_Qty"], errors="coerce").fillna(0)
  merged['PR_Qty'] = pd.to_numeric(merged['PR_Qty'], errors="coerce").fillna(0)
  
  def calc_outstanding(row):
      if row['Status'] == 'Routing Approval':
          return row['PR_Qty']
      return max(0, row['PO_Qty'] - row['Rcv_Qty'])
      
  merged["Qty_Outstanding"] = merged.apply(calc_outstanding, axis=1)

  merged["PR_Date"] = merged["PR_Date"].dt.strftime("%m/%d/%Y").fillna("-")
  merged["PO_Date"] = merged["PO_Date"].dt.strftime("%m/%d/%Y").fillna("-")
  merged["Rcv_Date"] = (
      pd.to_datetime(merged["Rcv_Date"], errors="coerce")
      .dt.strftime("%m/%d/%Y")
      .fillna("-")
  )
  merged["Vendor"] = merged["Vendor"].fillna("-")
  merged["Tipe_PO"] = merged["Tipe_PO"].fillna("-")

  final_cols = [
      "PR_Date",
      "PR_Manual_No",
      "Item_Code",
      "Item_Name",
      "PR_Qty",
      "PO_Date",
      "PO_No",
      "Vendor",
      "Tipe_PO",
      "PO_Qty",
      "Rcv_Date",
      "Rcv_Qty",
      "Qty_Outstanding",
      "Status",
  ]
  return merged[final_cols]


def get_formatted_archive_list():
  files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]
  if not files:
    return []
  files.sort(
      key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)),
      reverse=True,
  )
  formatted_list = []
  for idx, f in enumerate(files):
    mtime = os.path.getmtime(os.path.join(OUTPUT_FOLDER, f))
    dt_wib = datetime.fromtimestamp(mtime, timezone.utc).astimezone(WIB)
    dt_str = dt_wib.strftime("%d/%m/%Y %H:%M:%S WIB")
    label = f"[{dt_str}] {f}"
    if idx == 0:
      label += " ⭐ [TERBARU]"
    formatted_list.append((label, f))
  return formatted_list


# ==========================================
# 3 TABS MENU NAVIGATION
# ==========================================
selected_tab = st.radio(
    "Navigation Menu",
    ["📊 Dashboard", "⚙️ Proses Data", "📥 Download / Arsip"],
    horizontal=True,
    label_visibility="collapsed",
)

st.markdown(
    """<hr style="margin: 5px 0 15px 0; border: none; border-top: 1px solid #1e3d59;">""",
    unsafe_allow_html=True,
)

if selected_tab == "📊 Dashboard":
  st.subheader("Purchase Data Tracking | Executive Dashboard")

  if "df_final" in st.session_state:
    df_final = st.session_state["df_final"]
    if "last_saved" in st.session_state:
      st.caption(f"📁 Active Dataset: `{st.session_state['last_saved']}`")

    c1, c2, c3, c4 = st.columns(4)
    status_counts = df_final["Status"].value_counts()
    with c1:
      st.metric("Routing Approval", status_counts.get("Routing Approval", 0))
    with c2:
      st.metric(
          "Menunggu Pengiriman", status_counts.get("Menunggu Pengiriman", 0)
      )
    with c3:
      st.metric("Diterima Sebagian", status_counts.get("Diterima Sebagian", 0))
    with c4:
      st.metric("Sudah Diterima", status_counts.get("Sudah Diterima", 0))

    st.write("")
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
      fig_pie = px.pie(
          df_final,
          names="Status",
          title="Proporsi Status Supply Chain",
          color_discrete_sequence=px.colors.qualitative.Pastel,
      )
      st.plotly_chart(fig_pie, use_container_width=True)
    with col_chart2:
      top_v = (
          df_final[df_final["Vendor"] != "-"]["Vendor"]
          .value_counts()
          .head(10)
          .reset_index()
      )
      top_v.columns = ["Vendor", "Jumlah"]
      fig_b = px.bar(
          top_v,
          x="Jumlah",
          y="Vendor",
          orientation="h",
          title="Top 10 Active Vendors",
          color="Jumlah",
      )
      fig_b.update_layout(yaxis={"categoryorder": "total ascending"})
      st.plotly_chart(fig_b, use_container_width=True)

    st.subheader("Master Data Grid Preview")
    st.dataframe(df_final.head(100), use_container_width=True, height=380)
  else:
    st.warning(
        "⚠️ Belum ada file arsip ditemukan. Silakan proses data melalui menu **⚙️ Proses Data**."
    )

elif selected_tab == "⚙️ Proses Data":
  st.subheader("Purchase Data Tracking | Processing Center")
  sub_proc_option = st.radio(
      "Pilih Modul Proses:",
      [
          "1. Purchase Data Tracking (Full ETL)",
          "2. Outstanding Information (Filter by Item Code)",
      ],
      horizontal=True,
  )
  st.divider()

  if sub_proc_option == "1. Purchase Data Tracking (Full ETL)":
    st.markdown(
        f'<div class="erp-panel">Unggah dokumen sumber (.xlsx). Sistem otomatis membuang PR dengan Qty=0, serta mendrop semua baris PO yang berstatus <i>Rejected/Closed</i>. Hasil disimpan permanen di folder lokal <code>{OUTPUT_FOLDER}</code>.</div>',
        unsafe_allow_html=True,
    )
    col_u1, col_u2 = st.columns(2)
    with col_u1:
      file_pr_2026 = st.file_uploader(
          "PR Data (Base 2026)", type=["xlsx"], key="u_pr_base"
      )
      file_pr_lama = st.file_uploader(
          "PRN Data (Closed Status Audit)", type=["xlsx"], key="u_pr_closed"
      )
      file_po_lokal = st.file_uploader(
          "PO Data - Lokal", type=["xlsx"], key="u_po_lok"
      )
    with col_u2:
      file_po_impor = st.file_uploader(
          "PO Data - Impor", type=["xlsx"], key="u_po_imp"
      )
      file_inbound = st.file_uploader(
          "Inbound Data (Warehouse Receipts)", type=["xlsx"], key="u_inb"
      )

    st.write("")
    run_process = st.button("⚙️ Execute Full ETL & Save Local", type="primary")

    if run_process:
      if all([
          file_pr_2026,
          file_pr_lama,
          file_po_lokal,
          file_po_impor,
          file_inbound,
      ]):
        with st.spinner("Menjalankan Data Transformation Pipeline..."):
          df_final = process_tracking_data(
              file_pr_2026,
              file_pr_lama,
              file_po_lokal,
              file_po_impor,
              file_inbound,
          )
          now_dt = datetime.now(WIB)
          timestamp_file = now_dt.strftime("%Y%m%d_%H%M%S")
          saved_filename = f"Tracking_Final_{timestamp_file}.xlsx"
          saved_filepath = os.path.join(OUTPUT_FOLDER, saved_filename)
          df_final.to_excel(saved_filepath, index=False)

          st.session_state["df_final"] = df_final
          st.session_state["last_saved"] = saved_filename
          time_wib_str = now_dt.strftime("%d/%m/%Y pukul %H:%M:%S WIB")
          st.success(
              f"✅ Data berhasil diproses & disimpan permanen di lokal: `{saved_filepath}` pada **{time_wib_str}**!"
          )
      else:
        st.error("⚠️ Error: Kelima file sumber wajib diunggah lengkap!")

  elif sub_proc_option == "2. Outstanding Information (Filter by Item Code)":
    st.markdown(
        '<div class="erp-panel">Unggah file Excel berisi kolom <b>Item Code</b>. Sistem akan otomatis menduplikasi yang unik, mencocokkan ke arsip terbaru, mengecualikan status <b>Sudah Diterima</b>, dan menghitung <b>Qty_Outstanding</b>.</div>',
        unsafe_allow_html=True,
    )

    file_item_input = st.file_uploader(
        "Upload File Item Code (.xlsx)", type=["xlsx"], key="u_item_filter"
    )
    run_filter = st.button("🔍 Generate Outstanding Report", type="primary")

    if run_filter:
      if file_item_input is not None:
        latest_path, latest_name = get_latest_archive_filepath()
        if not latest_path:
          st.error(
              "⚠️ Belum ada file arsip dasar di server/lokal. Jalankan Full ETL terlebih dahulu!"
          )
        else:
          try:
            df_base_master = pd.read_excel(latest_path)
            df_item_upload = pd.read_excel(file_item_input)

            target_col = None
            for col in df_item_upload.columns:
              if "item" in str(col).lower() and "code" in str(col).lower():
                target_col = col
                break
            if not target_col:
              target_col = df_item_upload.columns[0]

            unique_item_codes = (
                df_item_upload[target_col]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
                .unique()
            )

            df_base_master["Item_Code_Clean"] = (
                df_base_master["Item_Code"]
                .astype(str)
                .str.strip()
                .str.upper()
            )
            filtered_df = df_base_master[
                df_base_master["Item_Code_Clean"].isin(unique_item_codes)
                & (df_base_master["Status"] != "Sudah Diterima")
            ].copy()

            if "Item_Code_Clean" in filtered_df.columns:
              filtered_df = filtered_df.drop(columns=["Item_Code_Clean"])

            st.success(
                f"✅ Berhasil memproses! Menemukan {len(filtered_df)} baris data outstanding dari {len(unique_item_codes)} unique item code unik."
            )
            st.dataframe(filtered_df, use_container_width=True, height=400)

            out_io = BytesIO()
            filtered_df.to_excel(
                out_io, index=False, sheet_name="Outstanding Report"
            )
            timestamp_file = datetime.now(WIB).strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 Download Outstanding Report (.xlsx)",
                data=out_io.getvalue(),
                file_name=f"Outstanding_Report_{timestamp_file}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
          except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat membaca file/memproses: {e}")
      else:
        st.warning("⚠️ Mohon unggah file Excel berisi Item Code terlebih dahulu.")

elif selected_tab == "📥 Download / Arsip":
  st.subheader("Purchase Data Tracking | Audit & Document Repository")
  st.markdown(
      f'<div class="erp-panel">Arsip historis tersimpan aman di direktori lokal <code>{OUTPUT_FOLDER}</code>. Waktu tercatat dalam WIB (Zona Waktu Jakarta). File terbaru diberi penanda khusus.</div>',
      unsafe_allow_html=True,
  )

  archive_items = get_formatted_archive_list()

  if archive_items:
    labels = [item[0] for item in archive_items]
    filenames = [item[1] for item in archive_items]

    selected_label = st.selectbox("Select Archived Document:", labels)
    selected_filename = dict(zip(labels, filenames))[selected_label]

    file_path_dl = os.path.join(OUTPUT_FOLDER, selected_filename)
    if os.path.exists(file_path_dl):
      df_preview = pd.read_excel(file_path_dl)
      st.markdown(
          f"**Selected File Metadata:** `📁 {selected_filename}` | Total Records: `{len(df_preview)} rows`"
      )
      st.dataframe(df_preview.head(50), use_container_width=True, height=380)

      with open(file_path_dl, "rb") as f:
        st.download_button(
            label=f"📥 Download Repository File ({selected_filename})",
            data=f.read(),
            file_name=selected_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
  else:
    st.warning(
        f"⚠️ Belum ada file arsip ditemukan pada direktori lokal `{OUTPUT_FOLDER}`."
    )
