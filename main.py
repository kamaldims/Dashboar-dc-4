# ==================================================================
# BANK XYZ — CUSTOMER EXPERIENCE DASHBOARD (v3, tampilan dirapikan)
# ------------------------------------------------------------------
# Jalankan:   streamlit run app.py
#
# File yang harus ada di folder yang SAMA dengan file ini:
#   1. Deka_project_dataset_BankXYZ.csv
#   2. metadata_dashboard.xlsx  (atau .csv)  -> hasil buat_metadata.py
# ==================================================================

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==================================================================
# KONFIGURASI
# ==================================================================

BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "C:/Users/Kamal Aldimas/PycharmProjects/PythonProject2/Deka_project_dataset_BankXYZ.csv"
META_XLSX = BASE_DIR / "C:/UsersK/amal Aldimas/PycharmProjects/PythonProject2/metadata_dashboard.xlsx"
META_CSV = BASE_DIR / "C:/Users/Kamal Aldimas/PycharmProjects/PythonProject2/metadata_dashboard.csv"

# ---- Palet warna ----
C_DARK = "#0F4C81"      # biru tua (judul, aksen)
C_MID = "#2E77AE"       # biru sedang
C_LIGHT = "#7FB3D5"     # biru muda
C_PALE = "#D6E9F8"      # biru sangat muda
C_RED = "#E08A8F"       # merah lembut (detractor)
C_GREY = "#94A3B8"      # abu (kompetitor)
C_BG = "#F4F8FC"        # latar utama
BLUE_SEQ = ["#D6E9F8", "#A9CCE3", "#7FB3D5", "#5499C7", "#2E77AE", "#0F4C81"]

# ---- Sidebar gelap (kontras tinggi: teks terang di latar gelap) ----
SB_BG = "#15243B"
SB_TXT = "#DCE6F2"
SB_MUTED = "#9DB1C8"
SB_INPUT = "#21344F"
SB_BORDER = "#37506E"
NAV_ACTIVE = "#2F6FB0"

PAGE_RINGKASAN = "Ringkasan"
PAGE_BI = "Brand Image"
PAGE_BF = "Branch Facilities"
PAGE_SE = "Service Experience"
PAGE_ATM = "ATM Experience"
PAGES = {
    PAGE_RINGKASAN: "📊",
    PAGE_BI: "⭐",
    PAGE_BF: "🏢",
    PAGE_SE: "🤝",
    PAGE_ATM: "🏧",
}
# subjudul tiap halaman (pola seperti dashboard referensi)
PAGE_SUB = {
    PAGE_BI: "Persepsi nasabah terhadap citra & reputasi Bank XYZ. "
             "Skala 1–6; jawaban “tidak relevan” dikeluarkan dari perhitungan.",
    PAGE_BF: "Kepuasan terhadap fasilitas fisik kantor cabang — lokasi, "
             "parkir, banking hall, ruang tunggu, toilet, dan sarana pendukung.",
    PAGE_ATM: "Kepuasan terhadap layanan mesin ATM — keamanan, kenyamanan, "
              "fitur, ketersediaan, dan keandalan mesin.",
}

TP_SERVICE = ["Customer Service", "Teller", "Security",
              "Customer Advisor", "Service Electronics"]
TP_LABEL = {"Security": "Sekuriti",
            "Service Electronics": "Sarana Elektronik"}

ORDER_USIA = [
    "17 -19 tahun", "20 - 25 tahun", "26 - 30 tahun", "31 - 35 tahun",
    "36 - 40 tahun", "41 - 45 tahun", "46 - 50 tahun",
    "50 tahun dan ke atas",
]
ORDER_LAMA = [
    "1 bulan s/d 3 bulan", "3 bulan s/d 11 bulan",
    "1 tahun s/d 2 tahun 11 bulan", "3 tahun s/d 4 tahun 11 bulan",
    "5 tahun atau lebih",
]
ORDER_BY_VAR = {"S2_2": ORDER_USIA, "S4": ORDER_LAMA}

st.set_page_config(
    page_title="Bank XYZ Customer Experience Dashboard",
    page_icon="🏦",
    layout="wide",
)

# ==================================================================
# TEMA / CSS
#   Prinsip kontras: sidebar gelap -> SEMUA teks dibuat terang (aturan
#   menyeluruh) supaya tidak ada tulisan yang "nyaru" dengan latarnya;
#   area utama terang -> teks gelap. Menu dropdown selalu putih + teks
#   gelap agar selalu terbaca di mana pun pemicunya.
# ==================================================================

RAW_CSS = """
<style>
.stApp { background-color: __BG__; }
.block-container { padding-top: 1.6rem; padding-bottom: 2rem; max-width: 1500px; }

h1, h2, h3, h4 { color: __DARK__ !important; font-weight: 800; }

/* ====================== SIDEBAR (gelap) ====================== */
section[data-testid="stSidebar"] {
    background-color: __SB_BG__;
    border-right: 1px solid #0C1828;
}
/* aturan menyeluruh: apa pun di sidebar tampil terang -> selalu terbaca */
section[data-testid="stSidebar"] * { color: __SB_TXT__; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #FFFFFF !important; }
section[data-testid="stSidebar"] label { color: __SB_TXT__ !important;
    font-weight: 600 !important; }
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {
    color: #F2C66B !important;   /* peringatan responden < 30 */
}

/* kotak pilihan (selectbox/multiselect) — gelap, teks terang, tepi jelas */
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background-color: __SB_INPUT__ !important;
    border: 1px solid __SB_BORDER__ !important;
    border-radius: 9px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] *,
section[data-testid="stSidebar"] [data-baseweb="select"] input {
    color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] svg { fill: #BFD3EA !important; }
section[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: __MID__ !important;
}
section[data-testid="stSidebar"] [data-baseweb="tag"] * {
    color: #FFFFFF !important; fill: #FFFFFF !important;
}

/* navigasi halaman (radio dijadikan menu) */
section[data-testid="stSidebar"] div[role="radiogroup"] { gap: 4px; }
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    display: block;
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 10px;
    padding: 10px 13px;
    margin-bottom: 2px;
    transition: background .15s ease;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    background: rgba(255,255,255,0.09);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label * {
    color: __SB_TXT__ !important; font-weight: 600;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    background: __NAV_ACTIVE__;
    border-color: __NAV_ACTIVE__;
    box-shadow: 0 4px 12px rgba(47,111,176,0.45);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) * {
    color: #FFFFFF !important; font-weight: 700;
}
/* sembunyikan bulatan radio bawaan agar bersih seperti menu */
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
    display: none;
}

/* tombol reset */
section[data-testid="stSidebar"] [data-testid="stButton"] button {
    background-color: __MID__ !important;
    color: #FFFFFF !important;
    border: 1px solid __SB_BORDER__ !important;
    border-radius: 9px !important;
    width: 100%; font-weight: 700;
}
section[data-testid="stSidebar"] [data-testid="stButton"] button:hover {
    background-color: __DARK__ !important;
}
section[data-testid="stSidebar"] [data-testid="stButton"] button * { color:#FFFFFF !important; }

/* expander Filter Tambahan */
section[data-testid="stSidebar"] details {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 10px;
}
section[data-testid="stSidebar"] details summary,
section[data-testid="stSidebar"] details summary * { color: __SB_TXT__ !important; }

/* metrik di sidebar (Jumlah Responden Terfilter) */
section[data-testid="stSidebar"] div[data-testid="stMetric"] {
    background: rgba(47,111,176,0.18);
    border: 1px solid rgba(255,255,255,0.12);
    border-radius: 12px; padding: 10px 14px;
}
section[data-testid="stSidebar"] div[data-testid="stMetric"] label { color: __SB_MUTED__ !important; }
section[data-testid="stSidebar"] div[data-testid="stMetricValue"] { color: #FFFFFF !important; }
section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12); }

/* ====================== MENU DROPDOWN (selalu putih) ====================== */
div[data-baseweb="popover"] ul[role="listbox"] {
    background-color: #FFFFFF !important;
    border: 1px solid #9DBFE0 !important;
}
div[data-baseweb="popover"] li[role="option"],
div[data-baseweb="popover"] li[role="option"] * {
    color: #0B2540 !important; background-color: transparent;
}
div[data-baseweb="popover"] li[role="option"]:hover,
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: __PALE__ !important;
}

/* ====================== AREA UTAMA ====================== */
/* hero halaman Ringkasan */
.page-hero { margin-bottom: .5rem; }
.page-hero h1 { font-size: 2.05rem; line-height: 1.15; margin: 0 0 .25rem; }
.hero-sub { color:#46586F; font-size: .98rem; line-height: 1.5; margin: 0 0 .7rem;
    max-width: 1000px; }
.hero-meta { display:inline-block; background:#E7F0FA; color:__DARK__;
    font-weight:600; font-size:.85rem; padding:5px 13px; border-radius:999px; }

/* judul seksi: judul tebal + satu baris keterangan abu */
.sec-head { margin: 1.7rem 0 .7rem; }
.sec-head h2 { font-size: 1.42rem; margin: 0 0 .15rem; }
.sec-head p { color:#56697F; font-size: .9rem; margin: 0; line-height: 1.45; }

/* kartu KPI khusus (label, angka besar, badge pill) */
.kpi-row { display:flex; gap:16px; flex-wrap:wrap; margin:.3rem 0 .2rem; }
.kpi { flex:1 1 200px; background:#FFFFFF; border:1px solid #E1EAF3;
    border-radius:16px; padding:18px 20px;
    box-shadow:0 2px 12px rgba(15,76,129,0.06); }
.kpi-label { color:#5B6E86; font-size:.84rem; font-weight:600; margin-bottom:7px; }
.kpi-value { color:#0B1F33; font-size:2.05rem; font-weight:800; line-height:1; }
.kpi-unit { color:#90A2B8; font-size:1rem; font-weight:600; }
.kpi-pill { display:inline-block; margin-top:11px; padding:3px 11px;
    border-radius:999px; font-size:.78rem; font-weight:600; }
.kpi-pill.up { background:#E3F2E7; color:#1E7D4F; }
.kpi-pill.info { background:#E7F0FA; color:__DARK__; }
.kpi-pill.warn { background:#FDEEDC; color:#B5651D; }

/* setiap grafik Plotly tampil sebagai kartu putih rapi */
div[data-testid="stPlotlyChart"] {
    background:#FFFFFF; border:1px solid #E4EDF6; border-radius:14px;
    padding:10px 10px 4px; box-shadow:0 2px 10px rgba(15,76,129,0.05);
}

/* kartu metrik bawaan (dipakai di tab sub-kategori, area utama) */
div[data-testid="stMetric"] {
    background-color:#FFFFFF; border:1px solid #DCE7F2; border-radius:14px;
    padding:14px 16px; box-shadow:0 2px 10px rgba(15,76,129,0.05);
}
div[data-testid="stMetric"] label { color:__DARK__ !important; font-weight:700; }
div[data-testid="stMetricValue"] { color:#0B1F33 !important; font-weight:800; }
div[data-testid="stMetricDelta"] { color:#51647A !important; }

/* selectbox di area utama (Service Experience) — putih, teks gelap */
section[data-testid="stMain"] [data-baseweb="select"] > div {
    background:#FFFFFF !important; border:1px solid #9DBFE0 !important;
    border-radius:9px !important;
}
section[data-testid="stMain"] [data-baseweb="select"] * { color:#0B2540 !important; }
section[data-testid="stMain"] [data-baseweb="select"] svg { fill:__DARK__ !important; }

/* tab sub-kategori */
button[data-baseweb="tab"] { color:#36506B !important; font-weight:600; }
button[data-baseweb="tab"][aria-selected="true"] {
    color:__DARK__ !important; border-bottom:3px solid __DARK__;
}
div[data-baseweb="tab-highlight"] { background-color:__DARK__ !important; }

/* expander area utama + tabel */
section[data-testid="stMain"] div[data-testid="stExpander"] details {
    background:#FFFFFF; border:1px solid #DCE7F2; border-radius:12px;
}
section[data-testid="stMain"] div[data-testid="stExpander"] summary,
section[data-testid="stMain"] div[data-testid="stExpander"] summary * {
    color:__DARK__ !important; font-weight:700;
}

/* toggle heatmap */
section[data-testid="stMain"] [data-testid="stToggle"] label,
section[data-testid="stMain"] [data-testid="stToggle"] label * { color:#33506B !important; }

hr { border-color:#DCE7F2; }
footer, #MainMenu { visibility:hidden; }
</style>
"""


def build_css():
    theme = {
        "BG": C_BG, "DARK": C_DARK, "MID": C_MID, "PALE": C_PALE,
        "SB_BG": SB_BG, "SB_TXT": SB_TXT, "SB_MUTED": SB_MUTED,
        "SB_INPUT": SB_INPUT, "SB_BORDER": SB_BORDER, "NAV_ACTIVE": NAV_ACTIVE,
    }
    css = RAW_CSS
    for k, v in theme.items():
        css = css.replace("__%s__" % k, v)
    return css


st.markdown(build_css(), unsafe_allow_html=True)


# ---- helper tampilan (judul seksi & kartu KPI) -------------------
def page_hero(emoji, title, subtitle, meta=None):
    meta_html = f'<span class="hero-meta">{meta}</span>' if meta else ""
    st.markdown(
        f'<div class="page-hero"><h1>{emoji} {title}</h1>'
        f'<p class="hero-sub">{subtitle}</p>{meta_html}</div>',
        unsafe_allow_html=True)


def section_header(emoji, title, subtitle=""):
    sub = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(
        f'<div class="sec-head"><h2>{emoji} {title}</h2>{sub}</div>',
        unsafe_allow_html=True)


def kpi_row(cards):
    """cards: list of dict(label, value, unit?, pill?, tone?)."""
    html = ['<div class="kpi-row">']
    for c in cards:
        unit = f'<span class="kpi-unit">{c["unit"]}</span>' if c.get("unit") else ""
        pill = (f'<div class="kpi-pill {c.get("tone", "info")}">{c["pill"]}</div>'
                if c.get("pill") else "")
        html.append(
            f'<div class="kpi"><div class="kpi-label">{c["label"]}</div>'
            f'<div class="kpi-value">{c["value"]}{unit}</div>{pill}</div>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


# ==================================================================
# LOAD DATA & METADATA
# ==================================================================

def to_score(series: pd.Series, max_valid: int = 6) -> pd.Series:
    """'6  SANGAT PUAS' -> 6 ; '99 TIDAK RELEVAN' / kosong -> NaN."""
    s = series.astype(str).str.extract(r"^\s*(\d+)")[0]
    s = pd.to_numeric(s, errors="coerce")
    s = s.where(s <= max_valid)
    return s.astype("Float64")


@st.cache_data(show_spinner="Memuat data survei…")
def load_all():
    df = pd.read_csv(DATA_PATH, sep=";", header=1,
                     low_memory=False, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]

    if META_XLSX.exists():
        meta = pd.read_excel(META_XLSX, sheet_name="Metadata")
    else:
        meta = pd.read_csv(META_CSV, encoding="utf-8-sig")
    meta.columns = [str(c).strip() for c in meta.columns]
    for c in ("variable", "question", "label", "bank", "section",
              "touchpoint", "subgroup", "role"):
        meta[c] = meta[c].astype(str).str.strip()
    meta["include"] = pd.to_numeric(meta["include"],
                                    errors="coerce").fillna(0).astype(int)
    meta["scale_max"] = pd.to_numeric(meta["scale_max"],
                                      errors="coerce").fillna(6).astype(int)
    meta = meta[meta["variable"].isin(df.columns)].reset_index(drop=True)
    return df, meta


if not DATA_PATH.exists():
    st.error("File **Deka_project_dataset_BankXYZ.csv** tidak ditemukan. "
             "Letakkan di folder yang sama dengan file dashboard ini.")
    st.stop()
if not META_XLSX.exists() and not META_CSV.exists():
    st.error("Metadata dashboard belum ada. Jalankan dulu "
             "`python buat_metadata.py` untuk membuat "
             "**metadata_dashboard.xlsx / .csv**.")
    st.stop()

df, META = load_all()

ATTR = META[(META["role"] == "Atribut") & (META["include"] == 1)
            & (META["bank"] == "XYZ")]
OVERALL = META[(META["role"] == "Overall") & (META["include"] == 1)
               & (META["bank"] == "XYZ")]
EXTRA_FILTERS = META[(META["role"].str.startswith("Filter"))
                     & (META["subgroup"] == "Tambahan")]


def attrs_of(section=None, touchpoint=None) -> pd.DataFrame:
    sub = ATTR
    if section:
        sub = sub[sub["section"] == section]
    if touchpoint:
        sub = sub[sub["touchpoint"] == touchpoint]
    return sub


def overall_of(section=None, touchpoint=None) -> pd.DataFrame:
    sub = OVERALL
    if section:
        sub = sub[sub["section"] == section]
    if touchpoint:
        sub = sub[sub["touchpoint"] == touchpoint]
    return sub


# ==================================================================
# SIDEBAR — NAVIGASI + FILTER (persisten antar halaman)
# ==================================================================

MAIN_FILTERS = [
    ("f_prov", "PROV", "Provinsi"),
    ("f_kab", "KABKOTA", "Kabupaten/Kota"),
    ("f_cab", "CABANG", "Cabang"),
    ("f_usia", "S2_2", "Kelompok Usia"),
    ("f_lama", "S4", "Lama Menjadi Nasabah"),
]


def ordered_options(values, var):
    vals = pd.Series(values).dropna().unique().tolist()
    order = ORDER_BY_VAR.get(var)
    if order:
        return [v for v in order if v in vals] + sorted(
            v for v in vals if v not in order)
    return sorted(vals)


def reset_filters():
    for key, _, _ in MAIN_FILTERS:
        st.session_state[key] = "Semua"
    for var in EXTRA_FILTERS["variable"]:
        st.session_state[f"fx_{var}"] = []


def select_persist(key, label, options):
    opts = ["Semua"] + options
    if key not in st.session_state or st.session_state[key] not in opts:
        st.session_state[key] = "Semua"
    return st.sidebar.selectbox(label, opts, key=key)


with st.sidebar:
    st.title("🏦 Bank XYZ")
    st.caption("Customer Experience Dashboard")
    st.markdown("###### NAVIGASI")
    page = st.radio("Halaman", list(PAGES),
                    format_func=lambda p: f"{PAGES[p]}  {p}",
                    key="nav_page", label_visibility="collapsed")
    st.markdown("---")
    st.markdown("###### 🔎 FILTER DATA")

prov = select_persist("f_prov", "Provinsi",
                      ordered_options(df["PROV"], "PROV"))
_t = df if prov == "Semua" else df[df["PROV"] == prov]
kab = select_persist("f_kab", "Kabupaten/Kota",
                     ordered_options(_t["KABKOTA"], "KABKOTA"))
if kab != "Semua":
    _t = _t[_t["KABKOTA"] == kab]
cab = select_persist("f_cab", "Cabang",
                     ordered_options(_t["CABANG"], "CABANG"))
usia = select_persist("f_usia", "Kelompok Usia",
                      ordered_options(df["S2_2"], "S2_2"))
lama = select_persist("f_lama", "Lama Menjadi Nasabah",
                      ordered_options(df["S4"], "S4"))

with st.sidebar.expander("⚙️ Filter Tambahan"):
    for _, r in EXTRA_FILTERS.iterrows():
        var = r["variable"]
        st.multiselect(r["label"], ordered_options(df[var], var),
                       key=f"fx_{var}", placeholder="Semua")

st.sidebar.button("🔄 Reset Semua Filter", on_click=reset_filters)

# Terapkan filter
fdf = df
for key, col, _ in MAIN_FILTERS:
    val = st.session_state.get(key, "Semua")
    if val != "Semua":
        fdf = fdf[fdf[col] == val]
for _, r in EXTRA_FILTERS.iterrows():
    sel = st.session_state.get(f"fx_{r['variable']}", [])
    if sel:
        fdf = fdf[fdf[r["variable"]].isin(sel)]

st.sidebar.markdown("---")
st.sidebar.metric("Jumlah Responden Terfilter", f"{len(fdf):,}")
if len(fdf) == 0:
    st.warning("Tidak ada responden yang sesuai dengan kombinasi filter ini. "
               "Silakan longgarkan filter atau tekan Reset.")
    st.stop()
if len(fdf) < 30:
    st.sidebar.caption("⚠️ Responden < 30 — hasil agregat kurang representatif.")


# ==================================================================
# FUNGSI PERHITUNGAN
# ==================================================================

def mean_score(data, var, scale=6):
    if var not in data.columns:
        return None
    s = to_score(data[var], scale)
    return None if s.notna().sum() == 0 else float(s.mean())


def t2b(data, var, scale=6):
    s = to_score(data[var], scale).dropna()
    if len(s) == 0:
        return None
    return float(s.isin([scale - 1, scale]).mean() * 100)


def nps(data, var="G1A"):
    s = to_score(data[var], 10).dropna()
    if len(s) == 0:
        return None, None, None, None
    p = float((s >= 9).mean() * 100)
    d = float((s <= 6).mean() * 100)
    return p - d, p, 100 - p - d, d


def attribute_table(data, attr_rows: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in attr_rows.iterrows():
        m = mean_score(data, r["variable"])
        if m is None:
            continue
        rows.append({"variable": r["variable"], "Atribut": r["label"],
                     "Subkategori": r["subgroup"], "Skor": round(m, 2),
                     "T2B": t2b(data, r["variable"])})
    return pd.DataFrame(rows)


def fmt(v, suffix=""):
    return "-" if v is None else f"{v:.2f}{suffix}"


# ==================================================================
# FUNGSI GRAFIK
# ==================================================================

def style_fig(fig, height=420, title=None):
    fig.update_layout(
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color="#0B1F33", size=13, family="Arial"),
        legend=dict(font=dict(size=12, color="#0B1F33")),
        title=dict(text=title, font=dict(color=C_DARK, size=15)) if title else None,
        margin=dict(l=10, r=12, t=46 if title else 16, b=10),
    )
    fig.update_xaxes(tickfont=dict(color="#0B1F33", size=12),
                     gridcolor="#EEF3F8")
    fig.update_yaxes(tickfont=dict(color="#0B1F33", size=12),
                     gridcolor="#EEF3F8")
    return fig


PLOTLY_CFG = {"displaylogo": False,
              "modeBarButtonsToRemove": ["lasso2d", "select2d"]}


def bar_attributes(am: pd.DataFrame, title=None):
    am = am.sort_values("Skor")
    fig = px.bar(am, x="Skor", y="Atribut", orientation="h",
                 color="Skor", color_continuous_scale=BLUE_SEQ,
                 text="Skor")
    fig.update_traces(textposition="outside", cliponaxis=False,
                      texttemplate="%{text:.2f}",
                      textfont=dict(color="#0B1F33"))
    fig.update_coloraxes(showscale=False)
    fig.update_layout(xaxis_range=[1, 6.6], yaxis_title=None,
                      xaxis_title="Skor (1–6)")
    st.plotly_chart(style_fig(fig, max(230, 30 * len(am) + 110), title),
                    width="stretch", config=PLOTLY_CFG)


def bar_subgroups(am: pd.DataFrame, title=None):
    g = (am.groupby("Subkategori", as_index=False)
           .agg(Skor=("Skor", "mean"), Jumlah=("Atribut", "size"))
           .sort_values("Skor"))
    g["Skor"] = g["Skor"].round(2)
    fig = px.bar(g, x="Skor", y="Subkategori", orientation="h",
                 color="Skor", color_continuous_scale=BLUE_SEQ,
                 text="Skor", hover_data={"Jumlah": True})
    fig.update_traces(textposition="outside", cliponaxis=False,
                      texttemplate="%{text:.2f}",
                      textfont=dict(color="#0B1F33"))
    fig.update_coloraxes(showscale=False)
    fig.update_layout(xaxis_range=[1, 6.6], yaxis_title=None,
                      xaxis_title="Rata-rata skor sub-kategori")
    st.plotly_chart(style_fig(fig, max(200, 40 * len(g) + 96), title),
                    width="stretch", config=PLOTLY_CFG)


def heatmap_usia(data, attr_rows, title, max_attrs=8):
    base = attribute_table(data, attr_rows)
    if base.empty:
        return
    pick = base.sort_values("Skor").head(max_attrs)
    usia_groups = [u for u in ORDER_USIA if u in data["S2_2"].unique()]
    if not usia_groups:
        return
    matrix, labels = [], []
    for _, r in pick.iterrows():
        row = [mean_score(data[data["S2_2"] == u], r["variable"])
               for u in usia_groups]
        matrix.append([None if v is None else round(v, 2) for v in row])
        labels.append(r["Atribut"])
    fig = px.imshow(matrix,
                    x=[u.replace(" tahun", "") for u in usia_groups],
                    y=labels, color_continuous_scale=BLUE_SEQ,
                    text_auto=".2f", aspect="auto", labels=dict(color="Skor"))
    st.plotly_chart(style_fig(fig, max(330, 34 * len(labels) + 140), title),
                    width="stretch", config=PLOTLY_CFG)
    st.caption("Menampilkan atribut dengan skor terendah (prioritas "
               "perbaikan). Semakin gelap = semakin puas.")


# ==================================================================
# HALAMAN SEKSI (Brand Image / Branch Facilities / ATM / 1 touchpoint)
# ==================================================================

def render_section_page(title, subtitle, attr_rows, ov_rows, heat_key,
                        as_hero=False):
    if as_hero:
        page_hero(PAGES.get(title, "📈"), title, subtitle)
    else:
        section_header(PAGES.get(title, "📈"), title, subtitle)

    am = attribute_table(fdf, attr_rows)
    if am.empty:
        st.info("Tidak ada data atribut untuk bagian ini "
                "(periksa kolom include pada metadata).")
        return

    # ---- kartu ringkasan ----
    t2b_avg = am["T2B"].dropna().mean() if am["T2B"].notna().any() else None
    cards = [
        {"label": "Skor Keseluruhan", "value": f"{am['Skor'].mean():.2f}",
         "unit": " / 6", "pill": f"{len(am)} atribut", "tone": "info"},
        {"label": "Rata-rata % Puas (T2B)",
         "value": "-" if t2b_avg is None else f"{t2b_avg:.0f}",
         "unit": "" if t2b_avg is None else "%",
         "pill": "jawaban 5–6", "tone": "up"},
    ]
    ov_m = ov_lbl = None
    for _, r in ov_rows.iterrows():
        ov_m = mean_score(fdf, r["variable"])
        ov_lbl = r["label"]
        break
    if ov_m is not None:
        cards.append({"label": ov_lbl[:38], "value": f"{ov_m:.2f}",
                      "unit": " / 6", "pill": "skor overall", "tone": "info"})
    cards.append({"label": "Responden", "value": f"{len(fdf):,}",
                  "pill": "sesuai filter", "tone": "info"})
    kpi_row(cards)

    # ---- skor per sub-kategori ----
    n_sub = am["Subkategori"].nunique()
    if n_sub > 1:
        section_header("📊", "Skor per Sub-Kategori",
                       "Urut dari yang terendah — sub-kategori paling atas "
                       "adalah prioritas perbaikan.")
        bar_subgroups(am)

    # ---- detail per sub-kategori ----
    section_header("🔍", "Detail per Sub-Kategori",
                   "Klik tab untuk melihat skor tiap atribut di dalamnya.")
    order = (am.groupby("Subkategori")["Skor"].mean()
               .sort_values().index.tolist())
    tabs = st.tabs(order) if n_sub > 1 else [st.container()]
    for tab, sub in zip(tabs, order):
        with tab:
            sam = am[am["Subkategori"] == sub]
            st2b = (sam["T2B"].dropna().mean()
                    if sam["T2B"].notna().any() else None)
            kpi_row([
                {"label": f"Overall {sub}", "value": f"{sam['Skor'].mean():.2f}",
                 "unit": " / 6", "tone": "info"},
                {"label": "% Puas (T2B)",
                 "value": "-" if st2b is None else f"{st2b:.0f}",
                 "unit": "" if st2b is None else "%", "tone": "up"},
                {"label": "Jumlah Atribut", "value": f"{len(sam)}",
                 "tone": "info"},
            ])
            bar_attributes(sam, f"Kepuasan per Atribut — {sub}")

    # ---- prioritas & kekuatan ----
    section_header("🎯", "Prioritas Perbaikan & Kekuatan Utama",
                   "Lima atribut terlemah dan terkuat pada bagian ini.")
    cl, cr = st.columns(2)
    show_cols = {"Atribut": "Atribut", "Subkategori": "Sub-Kategori",
                 "Skor": "Skor", "T2B": "% T2B"}
    low = am.sort_values("Skor").head(5)[list(show_cols)]
    high = am.sort_values("Skor", ascending=False).head(5)[list(show_cols)]
    with cl:
        st.markdown("**🔻 5 skor terendah (prioritas)**")
        st.dataframe(low.rename(columns=show_cols), hide_index=True,
                     width="stretch")
    with cr:
        st.markdown("**🔺 5 skor tertinggi (kekuatan)**")
        st.dataframe(high.rename(columns=show_cols), hide_index=True,
                     width="stretch")

    # ---- heatmap usia (opsional) ----
    if st.toggle("Tampilkan perbandingan antar kelompok usia (heatmap)",
                 key=heat_key):
        heatmap_usia(fdf, attr_rows,
                     f"Atribut Prioritas {title} per Kelompok Usia")


# ==================================================================
# HALAMAN 1 — RINGKASAN
# ==================================================================

def page_ringkasan():
    meta = (f"📅 {len(df):,} responden · {df['CABANG'].nunique()} cabang · "
            f"{df['PROV'].nunique()} provinsi")
    page_hero(
        "🏦", "Bank XYZ Customer Experience Dashboard",
        "Pantau CSAT, Loyalty, NPS, dan Customer Experience Bank XYZ "
        "berdasarkan survei nasabah. Gunakan filter di sidebar — pilihan "
        "filter tetap tersimpan saat berpindah halaman.",
        meta)

    csat_m, csat_p = mean_score(fdf, "E1A"), t2b(fdf, "E1A")
    loy_m, loy_p = mean_score(fdf, "F1A"), t2b(fdf, "F1A")
    nps_val, prom, pas, det = nps(fdf)

    kpi_row([
        {"label": "CSAT (Kepuasan)", "value": fmt(csat_m).split()[0]
         if csat_m is not None else "-", "unit": " / 6",
         "pill": None if csat_p is None else f"▲ {csat_p:.0f}% puas (T2B)",
         "tone": "up"},
        {"label": "Loyalty", "value": fmt(loy_m).split()[0]
         if loy_m is not None else "-", "unit": " / 6",
         "pill": None if loy_p is None else f"▲ {loy_p:.0f}% setuju (T2B)",
         "tone": "up"},
        {"label": "NPS", "value": "-" if nps_val is None else f"{nps_val:.0f}",
         "pill": None if prom is None else f"▲ {prom:.0f}% promoter",
         "tone": "up"},
        {"label": "Responden", "value": f"{len(fdf):,}",
         "pill": "sesuai filter", "tone": "info"},
    ])

    # ---- peta kepuasan ----
    section_header("🧭", "Peta Kepuasan per Touchpoint",
                   "Bandingkan kekuatan tiap titik layanan dan posisi terhadap "
                   "kompetitor.")
    col_l, col_r = st.columns(2)
    with col_l:
        tp_scores = []
        for tp in ATTR["touchpoint"].unique():
            am = attribute_table(fdf, attrs_of(touchpoint=tp))
            if not am.empty:
                tp_scores.append((TP_LABEL.get(tp, tp),
                                  round(am["Skor"].mean(), 2)))
        if tp_scores:
            cats = [t[0] for t in tp_scores]
            vals = [t[1] for t in tp_scores]
            rmin = min(4.0, min(vals) - 0.3)
            fig = go.Figure(go.Scatterpolar(
                r=vals + vals[:1], theta=cats + cats[:1],
                fill="toself", line_color=C_MID,
                fillcolor="rgba(46,119,174,0.25)"))
            fig.update_layout(
                polar=dict(radialaxis=dict(range=[rmin, 6], visible=True)),
                showlegend=False)
            st.plotly_chart(style_fig(fig, 420,
                            "Skor Customer Experience per Touchpoint"),
                            width="stretch", config=PLOTLY_CFG)
    with col_r:
        comp = []
        for label, vx, vk, sc in [("CSAT", "E1A", "E1B", 6),
                                  ("Loyalty", "F1A", "F1B", 6),
                                  ("NPS (rata-rata 0-10)", "G1A", "G1C", 10)]:
            mx, mk = mean_score(fdf, vx, sc), mean_score(fdf, vk, sc)
            if mx is not None:
                comp.append({"Indikator": label, "Bank": "XYZ", "Skor": mx})
            if mk is not None:
                comp.append({"Indikator": label, "Bank": "Kompetitor",
                             "Skor": mk})
        if comp:
            fig = px.bar(pd.DataFrame(comp), x="Indikator", y="Skor",
                         color="Bank", barmode="group", text_auto=".2f",
                         color_discrete_map={"XYZ": C_DARK,
                                             "Kompetitor": C_GREY})
            fig.update_traces(textfont=dict(color="#0B1F33"))
            fig.update_layout(legend_title=None, xaxis_title=None,
                              legend=dict(orientation="h", y=-0.2))
            st.plotly_chart(style_fig(fig, 420, "Bank XYZ vs Kompetitor"),
                            width="stretch", config=PLOTLY_CFG)

    # ---- komposisi NPS ----
    if nps_val is not None:
        section_header("📣", "Komposisi NPS",
                       "Proporsi Promoter, Passive, dan Detractor dari total "
                       "responden terfilter.")
        nps_df = pd.DataFrame({
            "Kategori": ["Promoter (9-10)", "Passive (7-8)", "Detractor (0-6)"],
            "Persen": [prom, pas, det]})
        fig = px.bar(nps_df, x="Persen", y=["NPS"] * 3, color="Kategori",
                     orientation="h", text="Persen",
                     color_discrete_sequence=[C_DARK, C_LIGHT, C_RED])
        fig.update_traces(texttemplate="%{x:.0f}%", textposition="inside",
                          textfont=dict(color="#FFFFFF"))
        fig.update_layout(barmode="stack", yaxis_title=None, xaxis_title="%",
                          yaxis_visible=False, legend_title=None,
                          legend=dict(orientation="h", y=-0.5))
        st.plotly_chart(style_fig(fig, 170), width="stretch", config=PLOTLY_CFG)

    # ---- prioritas lintas touchpoint ----
    all_am = attribute_table(fdf, ATTR)
    if not all_am.empty:
        section_header("🎯", "Prioritas Perbaikan Lintas Touchpoint",
                       "Delapan atribut dengan skor terendah di seluruh "
                       "layanan.")
        tp_of = dict(zip(ATTR["variable"], ATTR["touchpoint"]))
        low = all_am.sort_values("Skor").head(8).copy()
        low["Atribut"] = [
            f"{TP_LABEL.get(tp_of.get(v, ''), tp_of.get(v, ''))} — {a}"
            for v, a in zip(low["variable"], low["Atribut"])]
        bar_attributes(low)

    # ---- profil responden ----
    section_header("👤", "Profil Responden",
                   "Karakteristik nasabah yang mengisi survei.")
    p1, p2, p3 = st.columns(3)
    if "S1" in fdf.columns:
        with p1:
            g = fdf["S1"].value_counts().reset_index()
            g.columns = ["Gender", "Jumlah"]
            fig = px.pie(g, names="Gender", values="Jumlah", hole=0.5,
                         color_discrete_sequence=[C_DARK, C_LIGHT, C_PALE])
            fig.update_traces(textfont=dict(color="#FFFFFF", size=13))
            st.plotly_chart(style_fig(fig, 320, "Gender"),
                            width="stretch", config=PLOTLY_CFG)
    if "S2_2" in fdf.columns:
        with p2:
            u = (fdf["S2_2"].value_counts()
                 .reindex(ORDER_USIA).dropna().reset_index())
            u.columns = ["Usia", "Jumlah"]
            fig = px.bar(u, x="Usia", y="Jumlah", text="Jumlah",
                         color_discrete_sequence=[C_MID])
            fig.update_traces(textfont=dict(color="#0B1F33"))
            fig.update_layout(xaxis_title=None)
            st.plotly_chart(style_fig(fig, 320, "Kelompok Usia"),
                            width="stretch", config=PLOTLY_CFG)
    if "S4" in fdf.columns:
        with p3:
            lm = (fdf["S4"].value_counts()
                  .reindex(ORDER_LAMA).dropna().reset_index())
            lm.columns = ["Lama Menjadi Nasabah", "Jumlah"]
            fig = px.bar(lm, x="Jumlah", y="Lama Menjadi Nasabah",
                         orientation="h", text="Jumlah",
                         color_discrete_sequence=[C_MID])
            fig.update_traces(textfont=dict(color="#0B1F33"))
            fig.update_layout(yaxis_title=None)
            st.plotly_chart(style_fig(fig, 320, "Lama Menjadi Nasabah"),
                            width="stretch", config=PLOTLY_CFG)


# ==================================================================
# HALAMAN SERVICE EXPERIENCE (5 touchpoint petugas)
# ==================================================================

def page_service():
    section_header(
        PAGES[PAGE_SE], "Service Experience",
        "Pengalaman layanan dari petugas cabang: Customer Service, Teller, "
        "Sekuriti, Customer Advisor, dan Sarana Elektronik.")

    avail = [tp for tp in TP_SERVICE if not attrs_of(touchpoint=tp).empty]
    if not avail:
        st.info("Tidak ada atribut Service Experience pada metadata.")
        return

    cards = []
    for tp in avail:
        am = attribute_table(fdf, attrs_of(touchpoint=tp))
        cards.append({"label": TP_LABEL.get(tp, tp),
                      "value": "-" if am.empty else f"{am['Skor'].mean():.2f}",
                      "unit": "" if am.empty else " / 6",
                      "pill": f"{len(am)} atribut", "tone": "info"})
    kpi_row(cards)

    section_header("🔍", "Detail per Area Layanan",
                   "Pilih satu area untuk melihat sub-kategori dan atributnya. "
                   "Pilihan tersimpan saat berpindah halaman.")
    cur = st.session_state.get("se_tp_keep", avail[0])
    if cur not in avail:
        cur = avail[0]
    sel = st.selectbox("Pilih area layanan untuk detail:", avail,
                       index=avail.index(cur),
                       format_func=lambda t: TP_LABEL.get(t, t))
    st.session_state["se_tp_keep"] = sel

    render_section_page(
        TP_LABEL.get(sel, sel),
        f"Detail kepuasan layanan {TP_LABEL.get(sel, sel)} per sub-kategori "
        f"dan atribut.",
        attrs_of(touchpoint=sel), overall_of(touchpoint=sel),
        heat_key=f"hm_se_{sel}")


# ==================================================================
# ROUTING HALAMAN
# ==================================================================

if page == PAGE_RINGKASAN:
    page_ringkasan()
elif page == PAGE_SE:
    page_service()
else:
    render_section_page(page, PAGE_SUB.get(page, ""),
                        attrs_of(section=page), overall_of(section=page),
                        heat_key=f"hm_{page[:3].lower()}", as_hero=True)

st.markdown("---")
st.caption("Skala kepuasan 1–6 · **T2B** (Top-2-Box) = % responden menjawab 5 "
           "atau 6 · **NPS** = % Promoter (9–10) − % Detractor (0–6) · "
           "jawaban “tidak relevan” dikeluarkan dari perhitungan.")