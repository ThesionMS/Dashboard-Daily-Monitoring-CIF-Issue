import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Daily Monitoring Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F0F4F8; }

/* Sidebar */
[data-testid="stSidebar"] { background: #1E3A5F !important; }
[data-testid="stSidebar"] > div:first-child { background: #1E3A5F !important; }
[data-testid="stSidebar"] * { color: #E2EBF5 !important; }
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #7EBFFF !important; font-size:0.78rem; letter-spacing:0.12em; text-transform:uppercase; margin-bottom:0.4rem; }
[data-testid="stSidebar"] label { color: #A8C8F0 !important; font-size:0.8rem; }
[data-testid="stSidebar"] .stTextInput input,
[data-testid="stSidebar"] .stMultiSelect > div,
[data-testid="stSidebar"] .stDateInput input { background:#2A4F7A !important; color:#E2EBF5 !important; border:1px solid #3D6FA0 !important; border-radius:6px; }
[data-testid="stSidebar"] .stMultiSelect span { color:#E2EBF5 !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] { background:#3D6FA0 !important; }
[data-testid="stSidebar"] [data-baseweb="tag"] span { color:#fff !important; }
[data-testid="stSidebar"] .stButton button {
    background: #2A6496 !important; color:#fff !important; border:none !important;
    border-radius:6px; font-size:0.8rem; padding:0.4rem 0.8rem;
}
[data-testid="stSidebar"] .stButton button:hover { background:#1A4F7A !important; }
[data-testid="stSidebar"] .stFileUploader { background:#2A4F7A !important; border:2px dashed #5B9BD5 !important; border-radius:8px; padding:0.5rem; }
[data-testid="stSidebar"] .stFileUploader label { color:#7EBFFF !important; }
[data-testid="stSidebar"] .stFileUploader [data-testid="stFileUploaderDropzone"] { background:#2A4F7A !important; }
[data-testid="stSidebar"] .stCaption { color:#A8C8F0 !important; font-size:0.7rem; }
[data-testid="stSidebar"] hr { border-color:#3D6FA0 !important; }

/* Header */
.main-header {
    background: linear-gradient(135deg, #1E3A5F 0%, #2C5F8A 100%);
    padding: 1.4rem 2rem; border-radius: 12px; margin-bottom: 1.4rem; color: white;
}
.main-header h1 { font-size:1.5rem; font-weight:700; margin:0; color:white; }
.main-header p  { font-size:0.82rem; color:#A8C8F0; margin:0.3rem 0 0 0; }

/* Metrics */
.metric-row { display:flex; gap:0.8rem; margin-bottom:1.2rem; flex-wrap:wrap; }
.mc { background:white; border-radius:10px; padding:0.9rem 1.2rem; flex:1; min-width:110px;
      box-shadow:0 1px 4px rgba(0,0,0,0.09); border-left:4px solid; }
.mc.blue  { border-color:#3B82F6; }
.mc.green { border-color:#22C55E; }
.mc.amber { border-color:#F59E0B; }
.mc.red   { border-color:#EF4444; }
.mc.purple{ border-color:#8B5CF6; }
.mc .val  { font-size:1.7rem; font-weight:800; color:#1E3A5F; line-height:1; }
.mc .lbl  { font-size:0.68rem; color:#6B7280; text-transform:uppercase; letter-spacing:0.06em; margin-top:0.25rem; }

.sec-title { font-size:0.72rem; font-weight:700; color:#6B7280; text-transform:uppercase;
             letter-spacing:0.1em; margin:1rem 0 0.5rem 0; }
</style>
""", unsafe_allow_html=True)

# ── Column Definitions ────────────────────────────────────────────────────────
EXPECTED_COLS = [
    "TanggalIssue", "KodeIssue", "Keterangan Issue", "Branch",
    "Nama Nasabah", "NomorCIF", "NomorRekening", "StatusMonitoring",
    "TanggalConfirm", "DataInput", "NIK", "SLA", "Months",
    "Day  Issue", "Year", "Konfirmasi", "Kategori", "Temuan",
    "User ID", "PIC",
]
DATE_COLS = ["TanggalIssue", "TanggalConfirm", "DataInput"]

# ── Helpers ───────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def parse_excel(file_bytes: bytes, filename: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
    df.columns = [str(c).strip() for c in df.columns]
    df["_SourceFile"] = filename
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def combine_dataframes(dfs):
    if not dfs:
        return pd.DataFrame(columns=EXPECTED_COLS + ["_SourceFile"])
    combined = pd.concat(dfs, ignore_index=True)
    for col in EXPECTED_COLS:
        if col not in combined.columns:
            combined[col] = None
    return combined


def fmt_date(series):
    return series.apply(
        lambda v: v.strftime("%d-%m-%Y") if pd.notna(v) and hasattr(v, "strftime") else ("" if pd.isna(v) else v)
    )


# ── Session State ─────────────────────────────────────────────────────────────
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data: dict = {}

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📋 Daily Monitoring Dashboard</h1>
  <p>Branch QA · Gabungan seluruh laporan harian</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📁 Upload File Excel")

    uploaded_files = st.file_uploader(
        "Pilih file Excel (.xlsx / .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        for f in uploaded_files:
            if f.name not in st.session_state.uploaded_data:
                try:
                    df_tmp = parse_excel(f.read(), f.name)
                    st.session_state.uploaded_data[f.name] = df_tmp
                    st.success(f"✅ {f.name}")
                except Exception as e:
                    st.error(f"❌ {f.name}: {e}")

    if st.session_state.uploaded_data:
        st.markdown("---")
        st.markdown("## 📂 File Aktif")
        to_remove = []
        for fname in list(st.session_state.uploaded_data.keys()):
            c1, c2 = st.columns([5, 1])
            c1.caption(f"📄 {fname}")
            if c2.button("✕", key=f"rm_{fname}"):
                to_remove.append(fname)
        for fname in to_remove:
            del st.session_state.uploaded_data[fname]
            st.rerun()

        if st.button("🗑️ Hapus Semua", use_container_width=True):
            st.session_state.uploaded_data = {}
            st.rerun()

    st.markdown("---")
    st.markdown("## 🔍 Filter Data")

# ── Build DataFrame ───────────────────────────────────────────────────────────
df_all = combine_dataframes(list(st.session_state.uploaded_data.values()))

if df_all.empty:
    st.info("👆 Upload file Excel harian di sidebar untuk memulai.")
    st.stop()

# ── Sidebar Filters ───────────────────────────────────────────────────────────
with st.sidebar:
    def opts(col):
        return sorted(df_all[col].dropna().astype(str).unique().tolist()) if col in df_all.columns else []

    selected_pic    = st.multiselect("PIC",               options=opts("PIC"),               placeholder="Semua PIC")
    selected_status = st.multiselect("Status Monitoring", options=opts("StatusMonitoring"),   placeholder="Semua Status")
    selected_cat    = st.multiselect("Kategori",          options=opts("Kategori"),           placeholder="Semua Kategori")
    selected_branch = st.multiselect("Branch",            options=opts("Branch"),             placeholder="Semua Branch")
    selected_konfirm= st.multiselect("Konfirmasi",        options=opts("Konfirmasi"),         placeholder="Semua Konfirmasi")
    selected_temuan = st.multiselect("Temuan",            options=opts("Temuan"),             placeholder="Semua")
    selected_source = st.multiselect("📄 File Sumber",    options=opts("_SourceFile"),        placeholder="Semua File")

    st.markdown("**Tanggal Issue**")
    date_min = df_all["TanggalIssue"].dropna().min()
    date_max = df_all["TanggalIssue"].dropna().max()
    if pd.notna(date_min) and pd.notna(date_max):
        date_from = st.date_input("Dari",    value=date_min.date(), min_value=date_min.date(), max_value=date_max.date())
        date_to   = st.date_input("Sampai",  value=date_max.date(), min_value=date_min.date(), max_value=date_max.date())
    else:
        date_from = date_to = None

# ── Apply Filters ─────────────────────────────────────────────────────────────
df_f = df_all.copy()
if selected_pic:     df_f = df_f[df_f["PIC"].astype(str).isin(selected_pic)]
if selected_status:  df_f = df_f[df_f["StatusMonitoring"].astype(str).isin(selected_status)]
if selected_cat:     df_f = df_f[df_f["Kategori"].astype(str).isin(selected_cat)]
if selected_branch:  df_f = df_f[df_f["Branch"].astype(str).isin(selected_branch)]
if selected_konfirm: df_f = df_f[df_f["Konfirmasi"].astype(str).isin(selected_konfirm)]
if selected_temuan:  df_f = df_f[df_f["Temuan"].astype(str).isin(selected_temuan)]
if selected_source:  df_f = df_f[df_f["_SourceFile"].isin(selected_source)]
if date_from and date_to:
    df_f = df_f[
        (df_f["TanggalIssue"] >= pd.Timestamp(date_from)) &
        (df_f["TanggalIssue"] <= pd.Timestamp(date_to))
    ]

# ── Search ────────────────────────────────────────────────────────────────────
search = st.text_input("🔎  Cari di seluruh kolom…", placeholder="Nama nasabah, CIF, kode issue, branch, PIC…")
if search:
    mask = df_f.apply(lambda col: col.astype(str).str.contains(search, case=False, na=False)).any(axis=1)
    df_f = df_f[mask]

# ── Metrics ───────────────────────────────────────────────────────────────────
total     = len(df_f)
total_all = len(df_all)
open_cnt  = df_f["StatusMonitoring"].astype(str).str.lower().str.contains("open").sum()
close_cnt = df_f["StatusMonitoring"].astype(str).str.lower().str.contains("close").sum()
temuan_cnt= df_f["Temuan"].astype(str).str.lower().isin(["ya","yes","1","true","temuan"]).sum()

st.markdown(f"""
<div class="metric-row">
  <div class="mc blue">  <div class="val">{total:,}</div>    <div class="lbl">Terfilter</div></div>
  <div class="mc blue">  <div class="val">{total_all:,}</div><div class="lbl">Total Semua</div></div>
  <div class="mc red">   <div class="val">{open_cnt:,}</div> <div class="lbl">Open</div></div>
  <div class="mc green"> <div class="val">{close_cnt:,}</div><div class="lbl">Closed</div></div>
  <div class="mc amber"> <div class="val">{temuan_cnt:,}</div><div class="lbl">Temuan</div></div>
  <div class="mc purple"><div class="val">{len(st.session_state.uploaded_data)}</div><div class="lbl">File Aktif</div></div>
</div>
""", unsafe_allow_html=True)

# ── Table ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">📊 Data Monitoring</div>', unsafe_allow_html=True)

if df_f.empty:
    st.warning("Tidak ada data yang cocok dengan filter / pencarian saat ini.")
else:
    show_cols = [c for c in EXPECTED_COLS if c in df_f.columns]
    if len(st.session_state.uploaded_data) > 1 or selected_source:
        show_cols = ["_SourceFile"] + show_cols
    df_show = df_f[show_cols].copy()

    for col in DATE_COLS:
        if col in df_show.columns:
            df_show[col] = fmt_date(df_show[col])

    col_cfg = {
        "_SourceFile":      st.column_config.TextColumn("📄 File",            width="medium"),
        "TanggalIssue":     st.column_config.TextColumn("Tgl Issue",          width="small"),
        "KodeIssue":        st.column_config.TextColumn("Kode Issue",         width="small"),
        "Keterangan Issue": st.column_config.TextColumn("Keterangan Issue",   width="large"),
        "Branch":           st.column_config.TextColumn("Branch",             width="medium"),
        "Nama Nasabah":     st.column_config.TextColumn("Nama Nasabah",       width="medium"),
        "NomorCIF":         st.column_config.TextColumn("Nomor CIF",          width="small"),
        "NomorRekening":    st.column_config.TextColumn("Nomor Rekening",     width="medium"),
        "StatusMonitoring": st.column_config.TextColumn("Status",             width="small"),
        "TanggalConfirm":   st.column_config.TextColumn("Tgl Confirm",        width="small"),
        "DataInput":        st.column_config.TextColumn("Data Input",         width="small"),
        "NIK":              st.column_config.TextColumn("NIK",                width="small"),
        "SLA":              st.column_config.TextColumn("SLA",                width="small"),
        "Months":           st.column_config.TextColumn("Bulan",              width="small"),
        "Day  Issue":       st.column_config.TextColumn("Hari",               width="small"),
        "Year":             st.column_config.TextColumn("Tahun",              width="small"),
        "Konfirmasi":       st.column_config.TextColumn("Konfirmasi",         width="large"),
        "Kategori":         st.column_config.TextColumn("Kategori",           width="small"),
        "Temuan":           st.column_config.TextColumn("Temuan",             width="small"),
        "User ID":          st.column_config.TextColumn("User ID",            width="small"),
        "PIC":              st.column_config.TextColumn("PIC",                width="small"),
    }

    st.dataframe(df_show, use_container_width=True, height=530, column_config=col_cfg, hide_index=True)

    # ── Export ────────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">⬇️ Export Data</div>', unsafe_allow_html=True)
    c1, c2, _ = st.columns([1, 1, 4])

    @st.cache_data
    def to_excel_bytes(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Monitoring")
        return buf.getvalue()

    c1.download_button("📥 Export Excel", data=to_excel_bytes(df_show),
        file_name=f"monitoring_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

    c2.download_button("📥 Export CSV", data=df_show.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"monitoring_{datetime.today().strftime('%Y%m%d')}.csv",
        mime="text/csv", use_container_width=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"Daily Monitoring · {datetime.now().strftime('%d %b %Y %H:%M')} · {len(st.session_state.uploaded_data)} file · {total_all:,} baris")
