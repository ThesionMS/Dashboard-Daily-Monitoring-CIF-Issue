import streamlit as st
import pandas as pd
import io
from datetime import datetime, date

# ── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Daily Monitoring Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Base */
[data-testid="stAppViewContainer"] { background: #F0F4F8; }
[data-testid="stSidebar"] { background: #1A2B45; }
[data-testid="stSidebar"] * { color: #E8EFF7 !important; }
[data-testid="stSidebar"] .stMarkdown h2 { color: #7DB8E8 !important; font-size: 0.85rem; letter-spacing: 0.1em; text-transform: uppercase; }
[data-testid="stSidebar"] label { color: #A8C4E0 !important; font-size: 0.8rem; }

/* Header */
.main-header {
    background: linear-gradient(135deg, #1A2B45 0%, #2C4A72 100%);
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    color: white;
}
.main-header h1 { font-size: 1.6rem; font-weight: 700; margin: 0; color: white; }
.main-header p { font-size: 0.85rem; color: #A8C4E0; margin: 0.3rem 0 0 0; }

/* Metric cards */
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: white;
    border-radius: 10px;
    padding: 1rem 1.4rem;
    flex: 1;
    min-width: 130px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    border-left: 4px solid;
}
.metric-card.blue  { border-color: #3B82F6; }
.metric-card.green { border-color: #22C55E; }
.metric-card.amber { border-color: #F59E0B; }
.metric-card.red   { border-color: #EF4444; }
.metric-card .val  { font-size: 1.8rem; font-weight: 800; color: #1A2B45; line-height: 1; }
.metric-card .lbl  { font-size: 0.72rem; color: #6B7280; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 0.3rem; }

/* Section title */
.section-title {
    font-size: 0.75rem;
    font-weight: 700;
    color: #6B7280;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin: 1.2rem 0 0.6rem 0;
}

/* Upload zone */
.upload-hint {
    background: #EFF6FF;
    border: 2px dashed #93C5FD;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
    font-size: 0.82rem;
    color: #3B82F6;
    margin-bottom: 1rem;
}

/* Table wrapper */
.table-wrap {
    background: white;
    border-radius: 12px;
    padding: 1.2rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}

/* Status badges via pandas style */
.badge-open     { background:#FEE2E2; color:#B91C1C; padding:2px 8px; border-radius:99px; font-size:0.75rem; font-weight:600; }
.badge-closed   { background:#D1FAE5; color:#065F46; padding:2px 8px; border-radius:99px; font-size:0.75rem; font-weight:600; }
.badge-progress { background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:99px; font-size:0.75rem; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
EXPECTED_COLS = [
    "TanggalIssue", "Kode", "IssueName", "BranchCode", "BranchName",
    "CustomerName", "CIFNumber", "AccountNo", "Status",
    "TanggalConfirm", "TanggalRespond", "TanggalSelesai",
    "Remarks", "IssueCat", "isFinding", "CreatedBy", "PIC"
]

DATE_COLS = ["TanggalIssue", "TanggalConfirm", "TanggalRespond", "TanggalSelesai"]

STATUS_COLORS = {
    "Open":        "#EF4444",
    "Closed":      "#22C55E",
    "In Progress": "#F59E0B",
    "Follow Up":   "#3B82F6",
    "Pending":     "#8B5CF6",
}

# ── Helpers ───────────────────────────────────────────────────────────────────
@st.cache_data
def parse_excel(file_bytes: bytes, filename: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    # Normalize column names (strip whitespace)
    df.columns = [c.strip() for c in df.columns]
    # Add source file column
    df["_SourceFile"] = filename
    # Coerce date columns
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def combine_dataframes(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    if not dfs:
        return pd.DataFrame(columns=EXPECTED_COLS + ["_SourceFile"])
    combined = pd.concat(dfs, ignore_index=True)
    # Ensure all expected columns exist
    for col in EXPECTED_COLS:
        if col not in combined.columns:
            combined[col] = None
    return combined


def color_status(val):
    colors = {
        "open":        ("FEE2E2", "B91C1C"),
        "closed":      ("D1FAE5", "065F46"),
        "in progress": ("FEF3C7", "92400E"),
        "follow up":   ("DBEAFE", "1E40AF"),
        "pending":     ("EDE9FE", "5B21B6"),
    }
    key = str(val).lower()
    for k, (bg, fg) in colors.items():
        if k in key:
            return f"background-color:#{bg}; color:#{fg}; border-radius:4px; font-weight:600;"
    return ""


def fmt_date(series: pd.Series) -> pd.Series:
    return series.apply(lambda v: v.strftime("%d-%m-%Y") if pd.notna(v) and hasattr(v, "strftime") else ("" if pd.isna(v) else v))


# ── Session State ──────────────────────────────────────────────────────────────
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data: dict[str, pd.DataFrame] = {}

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
    st.markdown('<div class="upload-hint">Upload satu atau banyak file Excel<br>(.xlsx / .xls) sekaligus</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        label="Pilih file Excel",
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

    # List loaded files with remove button
    if st.session_state.uploaded_data:
        st.markdown("---")
        st.markdown("## 📂 File Aktif")
        to_remove = []
        for fname in list(st.session_state.uploaded_data.keys()):
            c1, c2 = st.columns([5, 1])
            c1.caption(fname)
            if c2.button("✕", key=f"rm_{fname}"):
                to_remove.append(fname)
        for fname in to_remove:
            del st.session_state.uploaded_data[fname]
            st.rerun()

        if st.button("🗑️ Hapus Semua File", use_container_width=True):
            st.session_state.uploaded_data = {}
            st.rerun()

    st.markdown("---")
    st.markdown("## 🔍 Filter Data")

# ── Build Combined DataFrame ──────────────────────────────────────────────────
df_all = combine_dataframes(list(st.session_state.uploaded_data.values()))

if df_all.empty:
    st.info("👆 Upload file Excel harian di sidebar untuk memulai.")
    st.stop()

# ── Sidebar Filters ───────────────────────────────────────────────────────────
with st.sidebar:
    # PIC filter
    pic_options = sorted(df_all["PIC"].dropna().unique().tolist())
    selected_pic = st.multiselect("PIC", options=pic_options, placeholder="Semua PIC")

    # Status filter
    status_options = sorted(df_all["Status"].dropna().unique().tolist())
    selected_status = st.multiselect("Status", options=status_options, placeholder="Semua Status")

    # IssueCat filter
    cat_options = sorted(df_all["IssueCat"].dropna().unique().tolist())
    selected_cat = st.multiselect("Kategori Issue", options=cat_options, placeholder="Semua Kategori")

    # Branch filter
    branch_options = sorted(df_all["BranchName"].dropna().unique().tolist())
    selected_branch = st.multiselect("Branch", options=branch_options, placeholder="Semua Branch")

    # Date range filter
    st.markdown("**Tanggal Issue**")
    date_min = df_all["TanggalIssue"].dropna().min()
    date_max = df_all["TanggalIssue"].dropna().max()

    if pd.notna(date_min) and pd.notna(date_max):
        date_from = st.date_input("Dari", value=date_min.date(), min_value=date_min.date(), max_value=date_max.date())
        date_to   = st.date_input("Sampai", value=date_max.date(), min_value=date_min.date(), max_value=date_max.date())
    else:
        date_from, date_to = None, None

    # isFinding filter
    finding_options = sorted(df_all["isFinding"].dropna().unique().tolist())
    selected_finding = st.multiselect("isFinding", options=[str(x) for x in finding_options], placeholder="Semua")

    # Source file filter
    source_options = sorted(df_all["_SourceFile"].dropna().unique().tolist())
    selected_source = st.multiselect("📄 File Sumber", options=source_options, placeholder="Semua File")

# ── Apply Filters ─────────────────────────────────────────────────────────────
df_filtered = df_all.copy()

if selected_pic:
    df_filtered = df_filtered[df_filtered["PIC"].isin(selected_pic)]
if selected_status:
    df_filtered = df_filtered[df_filtered["Status"].isin(selected_status)]
if selected_cat:
    df_filtered = df_filtered[df_filtered["IssueCat"].isin(selected_cat)]
if selected_branch:
    df_filtered = df_filtered[df_filtered["BranchName"].isin(selected_branch)]
if date_from and date_to and "TanggalIssue" in df_filtered.columns:
    df_filtered = df_filtered[
        (df_filtered["TanggalIssue"] >= pd.Timestamp(date_from)) &
        (df_filtered["TanggalIssue"] <= pd.Timestamp(date_to))
    ]
if selected_finding:
    df_filtered = df_filtered[df_filtered["isFinding"].astype(str).isin(selected_finding)]
if selected_source:
    df_filtered = df_filtered[df_filtered["_SourceFile"].isin(selected_source)]

# ── Search Bar ────────────────────────────────────────────────────────────────
search = st.text_input(
    "🔎  Cari di seluruh kolom…",
    placeholder="Ketik nama nasabah, CIF, kode issue, branch, dll…"
)
if search:
    mask = df_filtered.apply(
        lambda col: col.astype(str).str.contains(search, case=False, na=False)
    ).any(axis=1)
    df_filtered = df_filtered[mask]

# ── Metric Cards ──────────────────────────────────────────────────────────────
total       = len(df_filtered)
total_all   = len(df_all)
open_cnt    = df_filtered["Status"].astype(str).str.lower().str.contains("open").sum()
closed_cnt  = df_filtered["Status"].astype(str).str.lower().str.contains("closed").sum()
finding_cnt = df_filtered["isFinding"].astype(str).str.lower().isin(["1", "true", "yes", "ya"]).sum()

st.markdown(f"""
<div class="metric-row">
  <div class="metric-card blue">
    <div class="val">{total:,}</div>
    <div class="lbl">Total Terfilter</div>
  </div>
  <div class="metric-card blue">
    <div class="val">{total_all:,}</div>
    <div class="lbl">Total Semua Data</div>
  </div>
  <div class="metric-card red">
    <div class="val">{open_cnt:,}</div>
    <div class="lbl">Open</div>
  </div>
  <div class="metric-card green">
    <div class="val">{closed_cnt:,}</div>
    <div class="lbl">Closed</div>
  </div>
  <div class="metric-card amber">
    <div class="val">{finding_cnt:,}</div>
    <div class="lbl">Finding</div>
  </div>
  <div class="metric-card blue">
    <div class="val">{len(st.session_state.uploaded_data)}</div>
    <div class="lbl">File Aktif</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Table ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">📊 Data Monitoring</div>', unsafe_allow_html=True)

if df_filtered.empty:
    st.warning("Tidak ada data yang cocok dengan filter/pencarian saat ini.")
else:
    # Display columns (hide _SourceFile by default unless filtered by source)
    display_cols = EXPECTED_COLS.copy()
    display_cols = [c for c in display_cols if c in df_filtered.columns]
    if selected_source or len(st.session_state.uploaded_data) > 1:
        display_cols = ["_SourceFile"] + display_cols

    df_show = df_filtered[display_cols].copy()

    # Format date columns for display
    for col in DATE_COLS:
        if col in df_show.columns:
            df_show[col] = fmt_date(df_show[col])

    # Column config for st.dataframe
    col_cfg = {
        "_SourceFile": st.column_config.TextColumn("📄 File", width="medium"),
        "TanggalIssue": st.column_config.TextColumn("Tgl Issue", width="small"),
        "TanggalConfirm": st.column_config.TextColumn("Tgl Confirm", width="small"),
        "TanggalRespond": st.column_config.TextColumn("Tgl Respond", width="small"),
        "TanggalSelesai": st.column_config.TextColumn("Tgl Selesai", width="small"),
        "Kode": st.column_config.TextColumn("Kode", width="small"),
        "IssueName": st.column_config.TextColumn("Issue Name", width="large"),
        "BranchCode": st.column_config.TextColumn("Kode Cabang", width="small"),
        "BranchName": st.column_config.TextColumn("Nama Cabang", width="medium"),
        "CustomerName": st.column_config.TextColumn("Nama Nasabah", width="medium"),
        "CIFNumber": st.column_config.TextColumn("CIF", width="small"),
        "AccountNo": st.column_config.TextColumn("No. Rekening", width="medium"),
        "Status": st.column_config.TextColumn("Status", width="small"),
        "Remarks": st.column_config.TextColumn("Remarks / Catatan Konfirmasi", width="large"),
        "IssueCat": st.column_config.TextColumn("Kategori", width="small"),
        "isFinding": st.column_config.TextColumn("Finding?", width="small"),
        "CreatedBy": st.column_config.TextColumn("Dibuat Oleh", width="small"),
        "PIC": st.column_config.TextColumn("PIC", width="small"),
    }

    with st.container():
        st.dataframe(
            df_show,
            use_container_width=True,
            height=520,
            column_config=col_cfg,
            hide_index=True,
        )

    # ── Export ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Export Data</div>', unsafe_allow_html=True)
    col_dl1, col_dl2, _ = st.columns([1, 1, 4])

    # Excel export
    @st.cache_data
    def to_excel_bytes(df: pd.DataFrame) -> bytes:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Monitoring")
        return buf.getvalue()

    excel_bytes = to_excel_bytes(df_show)
    col_dl1.download_button(
        label="📥 Export Excel",
        data=excel_bytes,
        file_name=f"monitoring_export_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # CSV export
    csv_bytes = df_show.to_csv(index=False).encode("utf-8-sig")
    col_dl2.download_button(
        label="📥 Export CSV",
        data=csv_bytes,
        file_name=f"monitoring_export_{datetime.today().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(f"Dashboard Daily Monitoring · Diperbarui: {datetime.now().strftime('%d %b %Y %H:%M')} · {len(st.session_state.uploaded_data)} file aktif · {total_all:,} baris total")
