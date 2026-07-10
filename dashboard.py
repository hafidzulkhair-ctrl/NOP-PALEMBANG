"""
=====================================================================
 AVA ANALYSIS DASHBOARD - PT Telkomsel NOP Palembang
=====================================================================
Dashboard interaktif berbasis Streamlit untuk analisa data Availability
(AVA) dari file Excel, dengan layout:

    [ TOP 10 (dropdown Kab / Kec / Cluster) ]
    [ COMPARE (bandingkan availability site) ]
    [ GRAFIK (tren availability per site)    ]

Cara menjalankan:
    pip install streamlit pandas openpyxl plotly
    streamlit run dashboard.py

File Excel sumber data (SCRIPT_AVA.xlsx) harus berada di folder yang
sama dengan dashboard.py, atau ubah EXCEL_FILE di bawah.
=====================================================================
"""

import os
import hmac
import hashlib
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

from logo_asset import LOGO_B64  # logo Telkomsel, disimpan sebagai base64
from nop_logo_asset import NOP_LOGO_B64  # logo NOP Palembang, disimpan sebagai base64

# ==================================================
# KONFIGURASI
# ==================================================
EXCEL_FILE = os.environ.get("EXCEL_FILE", "SCRIPT_AVA.xlsx")
EXCEL_SHEET_NAME = os.environ.get("EXCEL_SHEET_NAME", 0)

TSEL_RED = "#E4002B"
TSEL_RED_DARK = "#B0001F"
TSEL_ORANGE = "#FF7A1A"
TSEL_YELLOW = "#FFC63C"

PAGE_BG_1 = "#13294B"   # Corporate Blue
PAGE_BG_2 = "#1D4E89"
NAVY_CARD = "#111B36"
NAVY_CARD_2 = "#16223F"
NAVY_BORDER = "rgba(228,0,43,0.28)"
TEXT_LIGHT = "#F2F4F8"
TEXT_MUTED = "#9AA4BD"
TEXT_MUTED_ON_LIGHT = "#5B6478"

st.set_page_config(
    page_title="AVA Analysis Dashboard - TSEL NOP Palembang",
    page_icon="📶",
    layout="wide",
)

# ==================================================
# CUSTOM CSS - tema Telkomsel (merah - putih)
# ==================================================
st.markdown(
    f"""
    <style>
        html, body, [class*="css"] {{
            font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        }}

        /* Background halaman - terang */
        .stApp {{
            background: linear-gradient(180deg, {PAGE_BG_1} 0%, {PAGE_BG_2} 100%);
        }}

        /* Sembunyikan menu/footer bawaan Streamlit yang tidak perlu */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}

        /* Teks umum: gelap di atas background terang */
        html, body, [class*="css"] {{
            color: #1A2138;
        }}
        [data-testid="stCaptionContainer"], .stCaption {{
            color: {TEXT_MUTED_ON_LIGHT} !important;
        }}

        /* Kartu metric - putih bersih */
        div[data-testid="stMetric"] {{
            background: #FFFFFF;
            border: 1px solid #E7E5F0;
            border-radius: 14px;
            padding: 14px 16px 10px 16px;
            box-shadow: 0 4px 14px rgba(20,30,60,0.06);
        }}
        div[data-testid="stMetricValue"] {{
            color: {TSEL_RED_DARK};
            font-weight: 700;
        }}
        div[data-testid="stMetricLabel"] {{
            color: {TEXT_MUTED_ON_LIGHT} !important;
        }}

        /* Judul section (###) */
        h3 {{
            color: #1A2138 !important;
            border-left: 5px solid {TSEL_RED};
            padding-left: 10px;
            margin-top: 6px !important;
        }}
        h2 {{
            color: #FFFFFF !important;
        }}

        /* Baris judul section (badge angka + judul) - WRAP-SAFE
           supaya teks judul tidak pernah kepotong di layar sempit */
        .section-title-row {{
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 10px;
        }}
        .section-title-badge {{
            flex: 0 0 auto;
            color: #FFFFFF;
            font-size: 12px;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            letter-spacing: 0.5px;
        }}
        .section-title-row h3.section-title-text {{
            margin: 0 !important;
            border: none !important;
            padding: 0 !important;
            flex: 1 1 200px;
            min-width: 0;
            overflow-wrap: break-word;
            word-break: break-word;
            white-space: normal;
        }}

        /* Jaring pengaman umum: teks jangan pernah overflow / kepotong,
           di ukuran layar berapa pun */
        h1, h2, h3, h4, p, span, label, button, .stMarkdown, [data-testid="stCaptionContainer"] {{
            overflow-wrap: break-word;
            word-break: break-word;
        }}
        .stApp, .block-container {{
            overflow-x: hidden;
            max-width: 100%;
        }}
        img {{
            max-width: 100%;
            height: auto;
        }}

        /* Tabs - pill terang dengan tab aktif gradient merah-oranye-kuning */
        div[data-baseweb="tab-list"] {{
            gap: 4px;
            background: #F3F1F6;
            border: 1px solid #E7E5F0;
            border-radius: 12px;
            padding: 4px;
        }}
        button[data-baseweb="tab"] {{
            font-weight: 600;
            color: {TEXT_MUTED_ON_LIGHT} !important;
            border-radius: 8px !important;
        }}
        button[data-baseweb="tab"][aria-selected="true"] {{
            color: #FFFFFF !important;
            background: linear-gradient(120deg, {TSEL_RED} 0%, {TSEL_ORANGE} 100%) !important;
        }}
        div[data-baseweb="tab-highlight"] {{
            background-color: transparent !important;
        }}
        div[data-baseweb="tab-border"] {{
            background-color: transparent !important;
        }}

        /* Dataframe / tabel */
        div[data-testid="stDataFrame"] {{
            border-radius: 10px;
            overflow: hidden;
            border: 1px solid #E7E5F0;
        }}

        /* Divider tipis */
        hr {{
            border-top: 1px solid #E7E5F0 !important;
        }}

        /* Kartu section (container border=True) - putih bersih seperti kartu blog */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: #FFFFFF !important;
            border: 1px solid #E7E5F0 !important;
            border-radius: 16px !important;
            box-shadow: 0 6px 20px rgba(20,30,60,0.06);
            padding: 4px 6px;
        }}

        /* Header tabel */
        div[data-testid="stDataFrame"] thead tr th {{
            background-color: #FBEEEF !important;
            color: {TSEL_RED_DARK} !important;
            font-weight: 700 !important;
        }}

        /* Selectbox & input */
        div[data-baseweb="select"] > div {{
            border-radius: 8px !important;
            background-color: #FFFFFF !important;
            border-color: #E7E5F0 !important;
            color: #1A2138 !important;
        }}
        div[data-baseweb="popover"] li {{
            background-color: #FFFFFF !important;
            color: #1A2138 !important;
        }}

        /* Tombol submit form (misalnya tombol Login) - gradient merah-oranye-kuning */
        div[data-testid="stFormSubmitButton"] button {{
            background: linear-gradient(120deg, {TSEL_RED} 0%, {TSEL_ORANGE} 100%) !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            padding: 10px 0 !important;
            box-shadow: 0 4px 14px rgba(228,0,43,0.28);
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }}
        div[data-testid="stFormSubmitButton"] button:hover {{
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(228,0,43,0.38);
            color: #FFFFFF !important;
        }}

        /* Input teks (text_input) - konsisten dengan selectbox */
        div[data-testid="stTextInput"] input {{
            border-radius: 8px !important;
            background-color: #FFFFFF !important;
            border: 1px solid #E7E5F0 !important;
            color: #1A2138 !important;
        }}
        div[data-testid="stTextInput"] input:focus {{
            border-color: {TSEL_RED} !important;
            box-shadow: 0 0 0 1px {TSEL_RED} !important;
        }}

        /* Header: grid 3-kolom di desktop -> [logo Telkomsel] [teks judul] [logo NOP kanan] */
        .app-header {{
            display: grid;
            grid-template-columns: auto 1fr auto;
            grid-template-areas: "tsel text nop";
            align-items: center;
            gap: 20px;
            padding: 20px 28px;
            border-radius: 16px;
            margin-bottom: 22px;
        }}
        .app-header-logo-tsel {{
            grid-area: tsel;
            height: 38px;
        }}
        .app-header-logo-nop {{
            grid-area: nop;
            height: 56px;
            justify-self: end;
            filter: drop-shadow(0 2px 6px rgba(0,0,0,0.35));
        }}
        .app-header-text {{
            grid-area: text;
        }}
        .app-header-text h2 {{
            color: #FFFFFF !important;
            margin: 0;
            text-shadow: 0 1px 4px rgba(0,0,0,0.35);
        }}
        .app-header-text p {{
            color: #FFF3DC;
            margin: 2px 0 0 0;
            text-shadow: 0 1px 3px rgba(0,0,0,0.30);
        }}

        /* Footer - gradasi merah-oranye-kuning, senada dengan header */
        .app-footer {{
            background: linear-gradient(120deg, {TSEL_RED} 0%, {TSEL_ORANGE} 55%, {TSEL_YELLOW} 100%) !important;
            box-shadow: 0 6px 20px rgba(255,122,26,0.28);
            border-radius: 16px;
            margin-top: 34px;
            padding: 22px 24px;
        }}
        .app-footer p {{
            color: #FFF3DC !important;
            text-shadow: 0 1px 3px rgba(0,0,0,0.30);
        }}

        /* ================================================
           RESPONSIVE / MOBILE (tampilan HP)
           ================================================ */
        @media (max-width: 768px) {{
            /* Kurangi padding halaman biar tidak terlalu lebar-kosong */
            .block-container {{
                padding-top: 1rem !important;
                padding-left: 0.7rem !important;
                padding-right: 0.7rem !important;
                padding-bottom: 2rem !important;
            }}

            /* Header di HP: kolom kiri = logo Telkomsel (atas) + logo NOP (bawah),
               kolom kanan = judul teks, rata kanan, tinggi menyesuaikan 2 baris logo. */
            .app-header {{
                grid-template-columns: auto 1fr !important;
                grid-template-areas: "tsel text" "nop text" !important;
                row-gap: 8px !important;
                column-gap: 14px !important;
                padding: 16px 18px !important;
            }}
            .app-header-logo-tsel {{
                height: 30px !important;
            }}
            .app-header-logo-nop {{
                height: 50px !important;
                justify-self: start !important;
                align-self: start !important;
            }}
            .app-header-text {{
                align-self: center !important;
                text-align: left !important;
                width: 100% !important;
            }}
            .app-header-text h2 {{
                display: block !important;
                width: 100% !important;
                text-align: left !important;
                font-size: 1rem !important;
                line-height: 1.3 !important;
                white-space: normal !important;
            }}
            .app-header-text p {{
                display: block !important;
                width: 100% !important;
                text-align: left !important;
                font-size: 0.78rem !important;
            }}

            /* Kartu metric lebih ringkas & teks tidak kepotong */
            div[data-testid="stMetric"] {{
                padding: 10px 12px 8px 12px !important;
            }}
            div[data-testid="stMetricValue"] {{
                font-size: 1.1rem !important;
                white-space: normal !important;
                overflow-wrap: break-word !important;
                line-height: 1.25 !important;
            }}
            div[data-testid="stMetricLabel"] {{
                font-size: 0.72rem !important;
                white-space: normal !important;
            }}

            /* Judul section (###) lebih kecil */
            .section-title-text {{
                font-size: 1rem !important;
            }}
            .section-title-badge {{
                font-size: 11px !important;
                padding: 3px 9px !important;
            }}

            /* Tab lebih kecil & bisa discroll horizontal kalau perlu
               (bukan kepotong, tapi bisa di-swipe) */
            div[data-baseweb="tab-list"] {{
                overflow-x: auto !important;
                flex-wrap: nowrap !important;
                -webkit-overflow-scrolling: touch;
            }}
            button[data-baseweb="tab"] {{
                font-size: 0.78rem !important;
                padding: 8px 10px !important;
                white-space: nowrap !important;
            }}

            /* Tabel: aktifkan scroll horizontal biar kolom tidak terpotong */
            div[data-testid="stDataFrame"] {{
                font-size: 0.8rem !important;
                overflow-x: auto !important;
                -webkit-overflow-scrolling: touch;
            }}

            /* Kartu section (container border=True) lebih rapat */
            div[data-testid="stVerticalBlockBorderWrapper"] {{
                padding: 10px 10px !important;
            }}

            /* Selectbox teks lebih kecil biar muat, dan tidak kepotong */
            div[data-baseweb="select"] {{
                font-size: 0.85rem !important;
            }}
            div[data-baseweb="select"] > div {{
                min-height: 40px !important;
            }}
            div[data-baseweb="popover"] li {{
                white-space: normal !important;
                font-size: 0.85rem !important;
            }}

            /* Caption lebih kecil */
            .stCaption, [data-testid="stCaptionContainer"] {{
                font-size: 0.75rem !important;
            }}

            /* Judul grafik Plotly ikut mengecil biar tidak kepotong */
            .js-plotly-plot .gtitle {{
                font-size: 13px !important;
            }}
            .js-plotly-plot .xtitle, .js-plotly-plot .ytitle {{
                font-size: 11px !important;
            }}
        }}

        /* ================================================
           HP LAYAR SANGAT SEMPIT (<= 400px)
           ================================================ */
        @media (max-width: 400px) {{
            .app-header-text h2 {{
                font-size: 0.92rem !important;
            }}
            .app-header-logo-tsel {{
                height: 26px !important;
            }}
            .app-header-logo-nop {{
                height: 44px !important;
            }}
            .section-title-text {{
                font-size: 0.92rem !important;
            }}
            div[data-testid="stMetricValue"] {{
                font-size: 0.98rem !important;
            }}
            .block-container {{
                padding-left: 0.5rem !important;
                padding-right: 0.5rem !important;
            }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)
# ==================================================
# AUTENTIKASI (LOGIN) - supaya link public tidak bisa
# diakses sembarang orang, hanya user yang terdaftar
# di st.secrets["credentials"] yang bisa masuk.
# ==================================================
def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def _verify_login(username: str, password: str) -> bool:
    users = st.secrets.get("credentials", {})
    if not username or username not in users:
        return False
    expected_hash = users[username]
    return hmac.compare_digest(_hash_password(password), expected_hash)


def require_login():
    """Tampilkan halaman login kalau user belum login. Kalau belum, hentikan
    eksekusi script (st.stop()) supaya konten dashboard di bawahnya sama
    sekali tidak dirender/dikirim ke browser."""
    if st.session_state.get("authenticated"):
        return

    st.markdown("<div style='height:48px;'></div>", unsafe_allow_html=True)

    left, mid, right = st.columns([1, 1.15, 1])
    with mid:
        with st.container(border=True):
            st.markdown(
                f"""
                <div style="text-align:center;margin:-1px -1px 22px -1px;
                    padding:26px 20px 22px 20px;border-radius:15px 15px 0 0;
                    background:linear-gradient(120deg,{TSEL_RED} 0%,{TSEL_ORANGE} 55%,{TSEL_YELLOW} 100%);
                    box-shadow:0 6px 20px rgba(255,122,26,0.28);">
                    <div style="display:flex;justify-content:center;align-items:center;gap:16px;margin-bottom:14px;">
                        <img src="data:image/png;base64,{LOGO_B64}" style="height:32px;" />
                        <div style="width:1px;height:30px;background:rgba(255,255,255,0.55);"></div>
                        <img src="data:image/png;base64,{NOP_LOGO_B64}" style="height:44px;
                            filter:drop-shadow(0 2px 6px rgba(0,0,0,0.35));" />
                    </div>
                    <div style="font-size:20px;font-weight:800;color:#FFFFFF;
                        text-shadow:0 1px 4px rgba(0,0,0,0.35);letter-spacing:0.3px;">
                        AVA ANALYSIS DASHBOARD
                    </div>
                    <div style="font-size:12.5px;color:#FFF3DC;margin-top:3px;
                        text-shadow:0 1px 3px rgba(0,0,0,0.30);">
                        TSEL NOP Palembang
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            with st.form("login_form", clear_on_submit=False):
                username = st.text_input("Username", placeholder="Masukkan username")
                password = st.text_input("Password", type="password", placeholder="Masukkan password")
                st.markdown("<div style='height:4px;'></div>", unsafe_allow_html=True)
                submitted = st.form_submit_button("Login", use_container_width=True)

            if submitted:
                if _verify_login(username.strip(), password):
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username.strip()
                    st.rerun()
                else:
                    st.error("Username atau password salah. Silakan coba lagi.")

            st.markdown(
                f"""
                <p style="text-align:center;font-size:11px;color:{TEXT_MUTED_ON_LIGHT};
                    margin:16px 0 2px 0;">
                    Dashboard ini hanya untuk tim internal TSEL NOP Palembang.
                </p>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<div style='height:48px;'></div>", unsafe_allow_html=True)
    st.stop()


require_login()


# ==================================================
# LOAD & CLEAN DATA (logika sama seperti bot.py)
# ==================================================
@st.cache_data(show_spinner="Memuat data dari Excel...")
def load_data(path: str, sheet_name, mtime: float):
    """mtime dipakai sebagai cache-buster: kalau file excel di-update, cache
    otomatis refresh karena mtime berubah."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File Excel tidak ditemukan: {path}")

    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    if df.empty:
        raise ValueError("File Excel kosong atau tidak ada data yang terbaca.")

    df.columns = df.columns.str.strip()

    col_map = {}
    for col in df.columns:
        col_upper = str(col).upper().strip()
        if "SITE" in col_upper:
            col_map["site"] = col
        elif col_upper == "WEEK":
            col_map["week"] = col
        elif "AVA" in col_upper:
            col_map["ava"] = col
        elif col_upper in ["KAB", "KABUPATEN"]:
            col_map["kab"] = col
        elif col_upper in ["KEC", "KECAMATAN"]:
            col_map["kec"] = col
        elif "CLUSTER" in col_upper:
            col_map["cluster"] = col
        elif "GENSET" in col_upper:
            col_map["genset"] = col

    required = ["site", "week", "ava", "kab", "kec", "cluster"]
    missing = [x for x in required if x not in col_map]
    if missing:
        raise ValueError(
            f"Kolom berikut tidak ditemukan: {missing}\n"
            f"Kolom yang tersedia: {df.columns.tolist()}"
        )

    SITE, WEEK, AVA = col_map["site"], col_map["week"], col_map["ava"]
    KAB, KEC, CLUSTER = col_map["kab"], col_map["kec"], col_map["cluster"]

    df = df.dropna(subset=[SITE, WEEK, AVA])
    df[SITE] = df[SITE].astype(str).str.strip()
    df[KAB] = df[KAB].fillna("").astype(str).str.strip()
    df[KEC] = df[KEC].fillna("").astype(str).str.strip()
    df[CLUSTER] = df[CLUSTER].fillna("").astype(str).str.strip()
    if "genset" in col_map:
        GENSET = col_map["genset"]
        df[GENSET] = df[GENSET].fillna("").astype(str).str.strip()

    df[WEEK] = pd.to_numeric(df[WEEK], errors="coerce")
    df[AVA] = pd.to_numeric(df[AVA], errors="coerce")
    df = df.dropna(subset=[WEEK, AVA])
    df[WEEK] = df[WEEK].astype(int)

    return df, col_map


try:
    mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
    DF, COL_MAP = load_data(EXCEL_FILE, EXCEL_SHEET_NAME, mtime)
except Exception as e:
    st.error(f"Gagal memuat data dari file Excel: {e}")
    st.stop()

SITE, WEEK, AVA = COL_MAP["site"], COL_MAP["week"], COL_MAP["ava"]
KAB, KEC, CLUSTER = COL_MAP["kab"], COL_MAP["kec"], COL_MAP["cluster"]
HAS_GENSET = "genset" in COL_MAP
GENSET = COL_MAP.get("genset")

ALL_WEEKS = sorted(DF[WEEK].unique())

# ==================================================
# HELPER FUNCTIONS
# ==================================================
def compute_top_drop(filter_col_key: str, value: str, week_prev=None, week_last=None):
    filter_col = COL_MAP[filter_col_key]
    temp = DF[DF[filter_col].str.upper() == value.upper()]
    if temp.empty:
        return None, "Data tidak ditemukan!"

    weeks = sorted(temp[WEEK].dropna().unique())
    if len(weeks) < 2:
        return None, "Data week tidak cukup!"

    if week_prev is None or week_last is None:
        week_prev, week_last = weeks[-2], weeks[-1]

    prev = temp[temp[WEEK] == week_prev][[SITE, AVA]]
    last = temp[temp[WEEK] == week_last][[SITE, AVA]]

    compare = pd.merge(prev, last, on=SITE, suffixes=("_prev", "_last"))
    if compare.empty:
        return None, "Tidak ada site yang punya data di kedua week tersebut!"
    compare["Delta"] = compare[f"{AVA}_last"] - compare[f"{AVA}_prev"]

    top10 = compare.sort_values("Delta", ascending=True).head(10)
    top10 = top10.reset_index(drop=True)
    top10.insert(0, "Rank", range(1, len(top10) + 1))
    top10 = top10.rename(
        columns={
            SITE: "Site",
            f"{AVA}_prev": f"AVA W{week_prev}",
            f"{AVA}_last": f"AVA W{week_last}",
        }
    )
    return top10, None


def compute_consecutive_below_100(scope_df, n_weeks=5, threshold=100):
    weeks = sorted(scope_df[WEEK].dropna().unique())
    if len(weeks) < n_weeks:
        return None, f"Data minimal {n_weeks} week diperlukan untuk analisa ini!"

    lastn = weeks[-n_weeks:]
    dfn = scope_df[scope_df[WEEK].isin(lastn)]
    pivot = dfn.pivot_table(index=SITE, columns=WEEK, values=AVA, aggfunc="first")
    pivot = pivot.reindex(columns=lastn).dropna(how="any")

    if pivot.empty:
        return None, f"Tidak ada site dengan data lengkap di {n_weeks} week terakhir!"

    results = []
    for site, row in pivot.iterrows():
        values = row.tolist()
        streak = 0
        for v in reversed(values):
            if v < threshold:
                streak += 1
            else:
                break
        if streak == 0:
            continue
        results.append({"Site": site, "values": values, "Streak": streak, "avg": sum(values) / len(values)})

    if not results:
        return None, f"Tidak ada site dengan AVA < {threshold} berturut-turut di {n_weeks} week terakhir!"

    results.sort(key=lambda r: (-r["Streak"], r["avg"]))
    top10 = results[:10]

    rows = []
    for rank, r in enumerate(top10, start=1):
        row = {"Rank": rank, "Site": r["Site"]}
        for w, v in zip(lastn, r["values"]):
            row[f"W{int(w)}"] = round(v, 2)
        row["Streak"] = r["Streak"]
        rows.append(row)

    return pd.DataFrame(rows), None


def style_delta(df, col="Delta"):
    if col not in df.columns:
        return df

    def _color(v):
        if isinstance(v, (int, float)) and v < 0:
            return "color:#d62728;font-weight:600"
        if isinstance(v, (int, float)) and v > 0:
            return "color:#2ca02c;font-weight:600"
        return ""

    styler = df.style.format(precision=2)
    # pandas >= 2.1 renamed Styler.applymap to Styler.map; support both.
    if hasattr(styler, "map"):
        return styler.map(_color, subset=[col])
    return styler.applymap(_color, subset=[col])


def section_title(number: str, text: str):
    """Render judul section (badge angka + judul) dengan HTML yang wrap-safe,
    supaya tidak kepotong di layar HP yang sempit."""
    st.markdown(
        f"""
        <div class="section-title-row">
            <span class="section-title-badge" style="background:{TSEL_RED};">{number}</span>
            <h3 class="section-title-text">{text}</h3>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==================================================
# HEADER
# ==================================================
st.markdown(
    f"""
    <div class="app-header" style="background:linear-gradient(120deg,{TSEL_RED} 0%,{TSEL_ORANGE} 55%,{TSEL_YELLOW} 100%);
        box-shadow:0 6px 20px rgba(255,122,26,0.28);">
        <img class="app-header-logo-tsel" src="data:image/png;base64,{LOGO_B64}" />
        <img class="app-header-logo-nop" src="data:image/png;base64,{NOP_LOGO_B64}" />
        <div class="app-header-text">
            <h2>AVA ANALYSIS DASHBOARD</h2>
            <p>TSEL NOP Palembang</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown(f"👤**{st.session_state.get('username', '-')}**")
    if st.button("Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

c3, c4 = st.columns(2)
c3.metric("Week Ter-Update", f"{ALL_WEEKS[-1]}")
c4.metric("Rata-rata AVA (week terakhir)", f"{DF[DF[WEEK]==ALL_WEEKS[-1]][AVA].mean():.2f}%")

if mtime:
    import datetime
    updated_str = datetime.datetime.fromtimestamp(mtime).strftime("%d %b %Y, %H:%M")
    st.caption(f"🕒 Data terakhir diperbarui: {updated_str}")

st.divider()

# =====================================================================
# BAGIAN 1 — TOP 10 (dropdown Kab / Kec / Cluster)
# =====================================================================
with st.container(border=True):
    section_title("01", "🔻 TOP 10 DROP AVAILABILITY")

    tab_labels = ["Drop by Kab/Kec/Cluster", "Consecutive AVA < 100 (5 Weeks)"]
    if HAS_GENSET:
        tab_labels += ["Consecutive AVA - Genset STB", "Drop WoW - Genset STB"]
    tabs = st.tabs(tab_labels)

    # --- Tab 1: Top 10 Drop by Kab/Kec/Cluster (dropdown) ---
    with tabs[0]:
        colA, colB, colC = st.columns([1, 2, 1])
        with colA:
            level_label = st.selectbox("Level Wilayah", ["Kab", "Kec", "Cluster"], key="top10_level")
        level_key = {"Kab": "kab", "Kec": "kec", "Cluster": "cluster"}[level_label]
        options = sorted(DF[COL_MAP[level_key]].dropna().unique())
        options = [o for o in options if o != ""]
        with colB:
            value = st.selectbox(f"Pilih {level_label}", options, key="top10_value")
        with colC:
            week_options = ALL_WEEKS
            week_last_sel = st.selectbox("Week terakhir", week_options, index=len(week_options) - 1, key="top10_week_last")

        prev_idx = ALL_WEEKS.index(week_last_sel) - 1
        week_prev_sel = ALL_WEEKS[prev_idx] if prev_idx >= 0 else week_last_sel

        result, err = compute_top_drop(level_key, value, week_prev_sel, week_last_sel)
        if err:
            st.warning(err)
        else:
            st.caption(f"Week {week_prev_sel} → {week_last_sel} | {level_label}: **{value}**")
            left, right = st.columns([1.2, 1])
            with left:
                st.dataframe(style_delta(result), use_container_width=True, hide_index=True)
            with right:
                fig = px.bar(
                    result.sort_values("Delta"),
                    x="Delta", y="Site", orientation="h",
                    color="Delta", color_continuous_scale=["#d62728", "#FF7A1A"],
                    title="Delta AVA (Turun Terbesar)",
                )
                fig.update_layout(
                    height=380, margin=dict(l=10, r=10, t=40, b=10), showlegend=False,
                    template="plotly_white",
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font_color="#1A2138",
                    dragmode=False,
                    xaxis=dict(fixedrange=True),
                    yaxis=dict(fixedrange=True),
                )
                st.plotly_chart(
                    fig, use_container_width=True,
                    config={
                        "displayModeBar": True,
                        "scrollZoom": False,
                        "displaylogo": False,
                        "modeBarButtonsToRemove": [
                            "zoom2d", "pan2d", "select2d", "lasso2d",
                            "zoomIn2d", "zoomOut2d", "autoScale2d", "resetScale2d",
                        ],
                    },
                )

    # --- Tab 2: Consecutive AVA < 100 (5 weeks), scoped by Kab/Kec/Cluster ---
    with tabs[1]:
        colA, colB = st.columns([1, 2])
        with colA:
            level_label2 = st.selectbox("Level Wilayah", ["Semua Site", "Kab", "Kec", "Cluster"], key="cons_level")
        scope_df = DF
        if level_label2 != "Semua Site":
            level_key2 = {"Kab": "kab", "Kec": "kec", "Cluster": "cluster"}[level_label2]
            options2 = sorted([o for o in DF[COL_MAP[level_key2]].dropna().unique() if o != ""])
            with colB:
                value2 = st.selectbox(f"Pilih {level_label2}", options2, key="cons_value")
            scope_df = DF[DF[COL_MAP[level_key2]].str.upper() == value2.upper()]

        result2, err2 = compute_consecutive_below_100(scope_df)
        if err2:
            st.warning(err2)
        else:
            st.dataframe(result2, use_container_width=True, hide_index=True)

    # --- Tab 3 & 4: Genset STB specific views ---
    if HAS_GENSET:
        with tabs[2]:
            genset_scope = DF[DF[GENSET].str.upper() == "GENSET STB"]
            result3, err3 = compute_consecutive_below_100(genset_scope)
            if err3:
                st.warning(err3)
            else:
                st.dataframe(result3, use_container_width=True, hide_index=True)

        with tabs[3]:
            result4, err4 = compute_top_drop("site", "", None, None)  # placeholder, replaced below
            genset_scope4 = DF[DF[GENSET].str.upper() == "GENSET STB"]
            weeks4 = sorted(genset_scope4[WEEK].unique())
            if len(weeks4) < 2:
                st.warning("Data week tidak cukup untuk kategori Genset STB!")
            else:
                week_prev4, week_last4 = weeks4[-2], weeks4[-1]
                prev4 = genset_scope4[genset_scope4[WEEK] == week_prev4][[SITE, AVA]]
                last4 = genset_scope4[genset_scope4[WEEK] == week_last4][[SITE, AVA]]
                comp4 = pd.merge(prev4, last4, on=SITE, suffixes=("_prev", "_last"))
                comp4["Delta"] = comp4[f"{AVA}_last"] - comp4[f"{AVA}_prev"]
                top10_4 = comp4.sort_values("Delta").head(10).reset_index(drop=True)
                top10_4.insert(0, "Rank", range(1, len(top10_4) + 1))
                top10_4 = top10_4.rename(columns={SITE: "Site", f"{AVA}_prev": f"AVA W{week_prev4}", f"{AVA}_last": f"AVA W{week_last4}"})
                st.caption(f"Week {week_prev4} → {week_last4} | Genset: **STB**")
                st.dataframe(style_delta(top10_4), use_container_width=True, hide_index=True)

st.divider()

# =====================================================================
# BAGIAN 2 — COMPARE (bandingkan availability 1 site di 2 week)
# =====================================================================
with st.container(border=True):
    section_title("02", "⚖️ COMPARE AVAILABILITY SITE")

    site_options = sorted(DF[SITE].unique())
    col1, col2, col3 = st.columns([1.4, 1, 1])
    with col1:
        compare_site = st.selectbox("Site ID", site_options, key="cmp_site")
    site_weeks = sorted(DF[DF[SITE] == compare_site][WEEK].unique())
    with col2:
        cmp_week1 = st.selectbox("Week 1", site_weeks, index=max(0, len(site_weeks) - 2), key="cmp_w1")
    with col3:
        cmp_week2 = st.selectbox("Week 2", site_weeks, index=len(site_weeks) - 1, key="cmp_w2")

    row1 = DF[(DF[SITE] == compare_site) & (DF[WEEK] == cmp_week1)]
    row2 = DF[(DF[SITE] == compare_site) & (DF[WEEK] == cmp_week2)]

    if row1.empty or row2.empty:
        st.warning("Data tidak ditemukan untuk site/week yang dipilih.")
    else:
        ava1 = row1[AVA].iloc[0]
        ava2 = row2[AVA].iloc[0]
        delta = ava2 - ava1
        status = "NAIK" if delta > 0 else ("TURUN" if delta < 0 else "TETAP")
        status_color = "#2ca02c" if delta > 0 else ("#d62728" if delta < 0 else "#888")

        m1, m2, m3 = st.columns(3)
        m1.metric(f"AVA Week {cmp_week1}", f"{ava1:.2f}%")
        m2.metric(f"AVA Week {cmp_week2}", f"{ava2:.2f}%", delta=f"{delta:+.2f}%")
        m3.markdown(
            f"<div style='padding-top:10px'>Status: "
            f"<span style='color:{status_color};font-weight:700;font-size:18px'>{status}</span></div>",
            unsafe_allow_html=True,
        )

        info_cols = row1.iloc[0]
        st.caption(
            f"Kab: **{info_cols[KAB]}** | Kec: **{info_cols[KEC]}** | Cluster: **{info_cols[CLUSTER]}**"
            + (f" | Genset: **{row1.iloc[0][GENSET]}**" if HAS_GENSET else "")
        )

st.divider()

# =====================================================================
# BAGIAN 3 — GRAFIK (tren availability per site)
# =====================================================================
with st.container(border=True):
    section_title("03", "📈 GRAFIK PERBANDINGAN AVA")

    colg1, colg2 = st.columns([1.4, 2.6])
    with colg1:
        grafik_site = st.selectbox("Site ID", site_options, key="grafik_site")

    site_df = DF[DF[SITE] == grafik_site].sort_values(WEEK)
    grafik_weeks_all = site_df[WEEK].tolist()

    with colg2:
        if len(grafik_weeks_all) >= 2:
            n_default = 5
            default_start = grafik_weeks_all[-n_default] if len(grafik_weeks_all) >= n_default else grafik_weeks_all[0]
            default_end = grafik_weeks_all[-1]
            week_range = st.select_slider(
                "Rentang Week",
                options=grafik_weeks_all,
                value=(default_start, default_end),
                key="grafik_range",
                help="Default menampilkan 5 week terakhir. Geser slider untuk melihat rentang lain.",
            )
        else:
            week_range = (grafik_weeks_all[0], grafik_weeks_all[0]) if grafik_weeks_all else (None, None)

    if len(grafik_weeks_all) < 2:
        st.warning(f"Data site {grafik_site} tidak cukup untuk membuat grafik (minimal 2 week).")
    else:
        df_range = site_df[(site_df[WEEK] >= week_range[0]) & (site_df[WEEK] <= week_range[1])]
        weeks_plot = df_range[WEEK].tolist()
        ava_plot = df_range[AVA].tolist()

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=[str(w) for w in weeks_plot], y=ava_plot,
            mode="lines+markers+text",
            text=[f"{v:.1f}%" for v in ava_plot],
            textposition="top center",
            line=dict(color="#E4002B", width=3),
            marker=dict(size=7),
            name=grafik_site,
        ))
        fig2.add_hline(y=100, line_dash="dot", line_color="gray")
        fig2.update_layout(
            title=f"Tren Availability - Site {grafik_site}",
            xaxis_title="Week", yaxis_title="Availability (%)",
            height=420, margin=dict(l=10, r=10, t=50, b=10),
            template="plotly_white",
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font_color="#1A2138",
            dragmode=False,
            xaxis_fixedrange=True,
            yaxis_fixedrange=True,
        )
        st.plotly_chart(
            fig2, use_container_width=True,
            config={
                "displayModeBar": True,
                "scrollZoom": False,
                "displaylogo": False,
                "modeBarButtonsToRemove": [
                    "zoom2d", "pan2d", "select2d", "lasso2d",
                    "zoomIn2d", "zoomOut2d", "autoScale2d", "resetScale2d",
                ],
            },
        )

        delta_g = ava_plot[-1] - ava_plot[0]
        status_g = "NAIK" if delta_g > 0 else ("TURUN" if delta_g < 0 else "TETAP")
        st.caption(f"Week {weeks_plot[0]} → {weeks_plot[-1]} | Delta: **{delta_g:+.2f}%** ({status_g})")

st.markdown(
    f"""
    <div class="app-footer" style="text-align:center;">
        <img src="data:image/png;base64,{LOGO_B64}" style="height:20px;opacity:0.95;
            margin-bottom:8px;" />
        <p style="margin:4px 0 0 0;font-size:13px;">
            AVA Analysis Dashboard &middot; TSEL NOP Palembang
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)
