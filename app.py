import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import urllib.request
from datetime import datetime
import streamlit.components.v1 as components
import importlib
import pdf_generator
importlib.reload(pdf_generator)

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="VBSJA - जिलावार प्रविष्टि एवं रैंकिंग डैशबोर्ड",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for Simple, Clean, Modern UI/UX
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Mukta:wght@300;400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Mukta', 'Inter', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: white;
        padding: 22px 28px;
        border-radius: 14px;
        margin-bottom: 22px;
        box-shadow: 0 4px 14px rgba(30, 58, 138, 0.18);
    }
    .main-title {
        font-size: 2.1rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .subtitle {
        font-size: 1rem;
        color: #DBEAFE;
        margin-top: 6px;
        margin-bottom: 0;
    }
    
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(30, 58, 138, 0.08);
    }
    .metric-val {
        font-size: 1.75rem;
        font-weight: 700;
        margin: 3px 0;
    }
    .metric-lbl {
        font-size: 0.88rem;
        color: #475569;
        font-weight: 600;
    }
    
    .filter-banner {
        background: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-left: 5px solid #2563EB;
        border-radius: 8px;
        padding: 10px 16px;
        margin-bottom: 16px;
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 8px;
        font-size: 0.88rem;
    }
    .filter-badge {
        background: #FFFFFF;
        border: 1px solid #94A3B8;
        border-radius: 16px;
        padding: 3px 10px;
        color: #1E293B;
        font-size: 0.82rem;
        font-weight: 500;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    
    .rank-card {
        background: linear-gradient(135deg, #FEF3C7 0%, #FFFBEB 100%);
        border: 1.5px solid #F59E0B;
        border-radius: 14px;
        padding: 18px 24px;
        color: #92400E;
        box-shadow: 0 2px 8px rgba(245, 158, 11, 0.12);
        margin-bottom: 1.2rem;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 46px;
        border-radius: 8px 8px 0px 0px;
        padding: 8px 18px;
        font-weight: 600;
        font-size: 0.95rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# DATA LOADING & CACHING (ALWAYS LIVE FIRST)
# ==========================================
SHEET_ID = "1S7YsDPftknLLV_zib6L8lc039W0UTNHhCxNqp4cBBv0"
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
LOCAL_CACHE_FILE = os.path.join(os.path.dirname(__file__), "vbsja_data.csv")

NUMERIC_COLS = [
    "कुल की गई प्रविष्टि",
    "कुल प्रतिभागियों की संख्या",
    "कुल फोटो की संख्या",
    "कुल वीडियो की संख्या",
    "कुल photo",
    "कुल video"
]

STRING_COLS = [
    "जिला", "ब्लॉक", "ग्राम पंचायत", "निकाय", "वार्ड",
    "इवेंट स्तर", "इवेंट केटेगरी", "भागीदारी की स्थिति"
]

@st.cache_data(ttl=60, show_spinner=False)
def load_data(force_refresh=False):
    """
    Fetches the live Google Sheet first with timeout.
    Updates local backup cache on success.
    Falls back to local file only if offline/network fails.
    """
    df = None
    data_source = ""

    # Always try live Google Sheets first to ensure freshly entered data (like Salumber K, L) is loaded
    try:
        req = urllib.request.Request(GOOGLE_SHEET_URL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=7) as resp:
            content = resp.read()
        df = pd.read_csv(io.BytesIO(content), encoding='utf-8')
        with open(LOCAL_CACHE_FILE, 'wb') as f:
            f.write(content)
        data_source = "Google Sheets (लाइव डेटा - स्वतः सिंक)"
    except Exception as e:
        if os.path.exists(LOCAL_CACHE_FILE):
            try:
                df = pd.read_csv(LOCAL_CACHE_FILE, encoding='utf-8')
                data_source = f"स्थानीय बैकअप कैश (लाइव सिंक त्रुटि: {e})"
            except Exception as e2:
                raise RuntimeError(f"डेटा लोड करने में त्रुटि: {e} | {e2}")
        else:
            raise RuntimeError(f"Google Sheet से डेटा लोड नहीं हो सका: {e}")

    # Map photo and video columns flexibly (both Hindi and English names)
    photo_candidates = ["कुल फोटो की संख्या", "कुल photo", "कुल फोटो", "फोटो की संख्या"]
    found_photo_col = next((c for c in photo_candidates if c in df.columns), None)
    if found_photo_col:
        df["कुल फोटो की संख्या"] = pd.to_numeric(df[found_photo_col], errors='coerce').fillna(0).astype(int)
    else:
        df["कुल फोटो की संख्या"] = 0
    df["कुल photo"] = df["कुल फोटो की संख्या"]

    video_candidates = ["कुल वीडियो की संख्या", "कुल video", "कुल वीडियो", "वीडियो की संख्या"]
    found_video_col = next((c for c in video_candidates if c in df.columns), None)
    if found_video_col:
        df["कुल वीडियो की संख्या"] = pd.to_numeric(df[found_video_col], errors='coerce').fillna(0).astype(int)
    else:
        df["कुल वीडियो की संख्या"] = 0
    df["कुल video"] = df["कुल वीडियो की संख्या"]

    if "कुल की गई प्रविष्टि" in df.columns:
        df["कुल की गई प्रविष्टि"] = pd.to_numeric(df["कुल की गई प्रविष्टि"], errors='coerce').fillna(0).astype(int)
    else:
        df["कुल की गई प्रविष्टि"] = 0

    if "कुल प्रतिभागियों की संख्या" in df.columns:
        df["कुल प्रतिभागियों की संख्या"] = pd.to_numeric(df["कुल प्रतिभागियों की संख्या"], errors='coerce').fillna(0).astype(int)
    else:
        df["कुल प्रतिभागियों की संख्या"] = 0

    for col in STRING_COLS:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
        else:
            df[col] = ""

    return df, data_source

# ==========================================
# SIDEBAR CONTROLS & FILTERS
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3281/3281329.png", width=55)
    st.title("🎛️ फ़िल्टर पैनल")
    
    if st.button("🔄 डेटा रीफ़्रेश करें (Refresh)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")

# Load dataset
with st.spinner("डेटा लोड हो रहा है, कृपया प्रतीक्षा करें..."):
    try:
        raw_df, data_source_info = load_data()
    except Exception as e:
        st.error(f"डेटा लोड करने में असमर्थ: {e}")
        st.stop()

# ==========================================
@st.cache_data(show_spinner=False)
def compute_original_state_ranks(df):
    overall_district_agg = df.groupby("जिला", as_index=False).agg({
        "कुल की गई प्रविष्टि": "sum",
        "कुल प्रतिभागियों की संख्या": "sum",
        "कुल फोटो की संख्या": "sum",
        "कुल वीडियो की संख्या": "sum"
    })
    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        overall_district_agg[c] = overall_district_agg[c].astype(int)

    overall_district_agg["orig_ent_rank"] = overall_district_agg["कुल की गई प्रविष्टि"].rank(ascending=False, method='min').astype(int)
    overall_district_agg["orig_part_rank"] = overall_district_agg["कुल प्रतिभागियों की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_district_agg["orig_photo_rank"] = overall_district_agg["कुल फोटो की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_district_agg["orig_video_rank"] = overall_district_agg["कुल वीडियो की संख्या"].rank(ascending=False, method='min').astype(int)

    overall_district_agg["कुल photo"] = overall_district_agg["कुल फोटो की संख्या"]
    overall_district_agg["कुल video"] = overall_district_agg["कुल वीडियो की संख्या"]
    rank_map = overall_district_agg.set_index("जिला")[["orig_ent_rank", "orig_part_rank", "orig_photo_rank", "orig_video_rank"]].to_dict("index")
    return overall_district_agg, rank_map

overall_district_agg, rank_map = compute_original_state_ranks(raw_df)

# Build filter options: DISTRICT AND BLOCK FILTERS REMOVED AS REQUESTED
with st.sidebar:
    st.markdown("#### 🏷️ इवेंट केटेगरी फ़िल्टर")
    
    all_categories = sorted([cat for cat in raw_df["इवेंट केटेगरी"].unique() if cat])
    selected_categories = st.multiselect(
        "केटेगरी चुनें:",
        options=all_categories,
        default=all_categories
    )

    st.markdown("---")
    st.caption(f"📌 स्रोत: {data_source_info}")
    st.caption(f"🕒 लोड समय: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}")

# ==========================================
# APPLY FILTERS
# ==========================================
filtered_df = raw_df.copy()

if selected_categories:
    filtered_df = filtered_df[filtered_df["इवेंट केटेगरी"].isin(selected_categories)]
else:
    filtered_df = filtered_df.iloc[0:0]

# ==========================================
# ACTIVE FILTER SUMMARY FUNCTION
# ==========================================
ALL_DIST_OPTION = "🌟 सभी जिले (All Districts)"

def get_active_filter_summary(district_filter=None, count_label="प्रदर्शित रिकॉर्ड", count_value=None):
    chips = []
    text_summary = []
    
    # 1. District Filter
    if district_filter and district_filter != ALL_DIST_OPTION:
        chips.append(f"📍 <b>जिला:</b> {district_filter}")
        text_summary.append(f"जिला: {district_filter}")
    else:
        chips.append("📍 <b>जिला:</b> सभी जिले (All Districts)")
        text_summary.append("जिला: सभी जिले")
    
    # 2. Event Category Filter
    if len(selected_categories) == len(all_categories):
        cat_lbl = f"सभी {len(all_categories)} श्रेणियां"
    elif len(selected_categories) <= 2:
        cat_lbl = ", ".join(selected_categories)
    else:
        cat_lbl = f"{len(selected_categories)} श्रेणियां चयनित"
    chips.append(f"📁 <b>इवेंट केटेगरी:</b> {cat_lbl}")
    text_summary.append(f"इवेंट केटेगरी: {cat_lbl}")
    
    # 3. Records / Districts Count
    c_val = count_value if count_value is not None else len(filtered_df)
    chips.append(f"📊 <b>{count_label}:</b> {c_val:,}")
    text_summary.append(f"{count_label}: {c_val:,}")
    
    html = f"""
    <div class="filter-banner">
        <span style="font-weight: 700; color: #1E3A8A; display: inline-flex; align-items: center; gap: 4px;">
            📌 लागू फ़िल्टर (Active Filters):
        </span>
        {''.join([f"<span class='filter-badge'>{c}</span>" for c in chips])}
    </div>
    """
    plain_text = " | ".join(text_summary)
    return html, plain_text

# ==========================================
# DISTRICT LEVEL AGGREGATION & ORIGINAL RANKINGS
# ==========================================
@st.cache_data(show_spinner=False, max_entries=30)
def compute_district_aggregation(f_df, r_map):
    agg = f_df.groupby("जिला", as_index=False).agg({
        "कुल की गई प्रविष्टि": "sum",
        "कुल प्रतिभागियों की संख्या": "sum",
        "कुल फोटो की संख्या": "sum",
        "कुल वीडियो की संख्या": "sum"
    })

    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        agg[c] = agg[c].astype(int)

    agg["orig_ent_rank"] = agg["जिला"].map(lambda d: r_map.get(d, {}).get("orig_ent_rank", "-"))
    agg["orig_part_rank"] = agg["जिला"].map(lambda d: r_map.get(d, {}).get("orig_part_rank", "-"))
    agg["orig_photo_rank"] = agg["जिला"].map(lambda d: r_map.get(d, {}).get("orig_photo_rank", "-"))
    agg["orig_video_rank"] = agg["जिला"].map(lambda d: r_map.get(d, {}).get("orig_video_rank", "-"))

    agg["कुल photo"] = agg["कुल फोटो की संख्या"]
    agg["कुल video"] = agg["कुल वीडियो की संख्या"]

    agg = agg.sort_values(by="कुल की गई प्रविष्टि", ascending=False).reset_index(drop=True)
    return agg.rename(columns={"जिला": "District"})

district_agg = compute_district_aggregation(filtered_df, rank_map)

# 1. Merged Table (Single column with value and original rank in parentheses, e.g. "10,844 (1)")
# REMOVED "Rank:" word from all columns and headers
merged_district_table = pd.DataFrame()
merged_district_table["District"] = district_agg["District"]
merged_district_table["कुल की गई प्रविष्टि"] = [
    f"{v:,} ({r})" for v, r in zip(district_agg["कुल की गई प्रविष्टि"], district_agg["orig_ent_rank"])
]
merged_district_table["कुल प्रतिभागियों की संख्या"] = [
    f"{v:,} ({r})" for v, r in zip(district_agg["कुल प्रतिभागियों की संख्या"], district_agg["orig_part_rank"])
]
merged_district_table["कुल फोटो की संख्या"] = [
    f"{v:,} ({r})" for v, r in zip(district_agg["कुल फोटो की संख्या"], district_agg["orig_photo_rank"])
]
merged_district_table["कुल वीडियो की संख्या"] = [
    f"{v:,} ({r})" for v, r in zip(district_agg["कुल वीडियो की संख्या"], district_agg["orig_video_rank"])
]

# 2. Separate Columns Table (also without "Rank:" word)
separate_district_table = pd.DataFrame()
separate_district_table["District"] = district_agg["District"]
separate_district_table["कुल की गई प्रविष्टि"] = district_agg["कुल की गई प्रविष्टि"]
separate_district_table["प्रविष्टि मूल रैंक"] = district_agg["orig_ent_rank"]
separate_district_table["कुल प्रतिभागियों की संख्या"] = district_agg["कुल प्रतिभागियों की संख्या"]
separate_district_table["प्रतिभागी मूल रैंक"] = district_agg["orig_part_rank"]
separate_district_table["कुल फोटो की संख्या"] = district_agg["कुल फोटो की संख्या"]
separate_district_table["फोटो मूल रैंक"] = district_agg["orig_photo_rank"]
separate_district_table["कुल वीडियो की संख्या"] = district_agg["कुल वीडियो की संख्या"]
separate_district_table["वीडियो मूल रैंक"] = district_agg["orig_video_rank"]

# State totals
total_districts_count = len(district_agg)
total_entries = int(district_agg["कुल की गई प्रविष्टि"].sum())
total_participants = int(district_agg["कुल प्रतिभागियों की संख्या"].sum())
total_photos = int(district_agg["कुल फोटो की संख्या"].sum())
total_videos = int(district_agg["कुल वीडियो की संख्या"].sum())
top_district_name = district_agg.iloc[0]["District"] if not district_agg.empty else "N/A"
top_district_entries = district_agg.iloc[0]["कुल की गई प्रविष्टि"] if not district_agg.empty else 0

# ==========================================
# MAIN HEADER BANNER
# ==========================================
st.markdown("""
<div class="main-header">
    <div class="main-title">📊 VBSJA जिलावार प्रविष्टि एवं रैंकिंग पोर्टल</div>
    <div class="subtitle">Google Spreadsheet वास्तविक समय डेटा • प्रत्येक जिले की मूल राज्य रैंकिंग (Original State Rank) के साथ</div>
</div>
""", unsafe_allow_html=True)

# KPI Metric Cards
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
with kpi1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-lbl">कुल जिले (Districts)</div>
        <div class="metric-val" style="color: #1E3A8A;">{total_districts_count:,}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-lbl">कुल की गई प्रविष्टि</div>
        <div class="metric-val" style="color: #059669;">{total_entries:,}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-lbl">कुल प्रतिभागी संख्या</div>
        <div class="metric-val" style="color: #2563EB;">{total_participants:,}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-lbl">कुल फोटो / वीडियो की संख्या</div>
        <div class="metric-val" style="color: #D97706;">{total_photos:,} / {total_videos:,}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi5:
    st.markdown(f"""
    <div class="metric-card" style="border: 2px solid #F59E0B; background: #FFFBEB;">
        <div class="metric-lbl">🥇 #1 शीर्ष जिला (प्रविष्टि)</div>
        <div class="metric-val" style="color: #B45309; font-size: 1.45rem;">{top_district_name}</div>
        <div style="font-size: 0.8rem; color: #92400E; font-weight: 600;">{top_district_entries:,} प्रविष्टियां</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# DISTRICT SPOTLIGHT / ALL-RANKS CHECKER
# ==========================================
selected_spotlight_district = ALL_DIST_OPTION

if not district_agg.empty:
    pure_district_list = sorted([str(d).strip() for d in district_agg["District"].tolist() if str(d).strip()])
    
    # 1. Determine default district (Salumber) as requested
    default_district = "सलूम्बर"
    for d in pure_district_list:
        if "सलूम्बर" in d or "सलुम्बर" in d:
            default_district = d
            break

    # 2. Unified district state initialization (defaulting to Salumber)
    if "app_selected_district" not in st.session_state:
        st.session_state["app_selected_district"] = default_district
    elif st.session_state["app_selected_district"] not in pure_district_list:
        st.session_state["app_selected_district"] = default_district

    if "top_district_select" not in st.session_state:
        st.session_state["top_district_select"] = st.session_state["app_selected_district"]
    if "dist_summary_chosen_district" not in st.session_state:
        st.session_state["dist_summary_chosen_district"] = st.session_state["app_selected_district"]
    if "block_report_chosen_district" not in st.session_state:
        st.session_state["block_report_chosen_district"] = st.session_state["app_selected_district"]

    # 3. Synchronized callbacks: changing ANY dropdown synchronizes all others
    def on_top_district_change():
        chosen = st.session_state.get("top_district_select")
        if chosen and chosen != ALL_DIST_OPTION:
            st.session_state["app_selected_district"] = chosen
            st.session_state["dist_summary_chosen_district"] = chosen
            st.session_state["block_report_chosen_district"] = chosen

    def on_summary_district_change():
        chosen = st.session_state.get("dist_summary_chosen_district")
        if chosen:
            st.session_state["app_selected_district"] = chosen
            st.session_state["top_district_select"] = chosen
            st.session_state["block_report_chosen_district"] = chosen

    def on_block_district_change():
        chosen = st.session_state.get("block_report_chosen_district")
        if chosen:
            st.session_state["app_selected_district"] = chosen
            st.session_state["top_district_select"] = chosen
            st.session_state["dist_summary_chosen_district"] = chosen

    def set_spotlight_all():
        st.session_state["top_district_select"] = ALL_DIST_OPTION

    def set_spotlight_salumber():
        st.session_state["app_selected_district"] = default_district
        st.session_state["top_district_select"] = default_district
        st.session_state["dist_summary_chosen_district"] = default_district
        st.session_state["block_report_chosen_district"] = default_district

    with st.container():
        col_sel, col_det = st.columns([1, 2.5])
        
        with col_sel:
            st.markdown("#### 🎯 जिला विवरण व मूल रैंक (District Spotlight)")
            district_options = [ALL_DIST_OPTION] + pure_district_list
            
            selected_spotlight_district = st.selectbox(
                "जिला चुनें (Select District):",
                options=district_options,
                key="top_district_select",
                on_change=on_top_district_change,
                label_visibility="collapsed"
            )
            
            if selected_spotlight_district != ALL_DIST_OPTION:
                st.button("🌟 सभी 41 जिले देखें (Show All)", key="btn_show_all", on_click=set_spotlight_all, width="stretch")
            else:
                st.button(f"🎯 डिफ़ॉल्ट ({default_district}) देखें", key="btn_show_sal", on_click=set_spotlight_salumber, width="stretch")
        
        if selected_spotlight_district == ALL_DIST_OPTION:
            with col_det:
                st.markdown(f"""
                <div class="rank-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                        <div>
                            <span style="font-size: 1.55rem; font-weight: 800;">🌟 राज्य स्तरीय अवलोकन (All Districts)</span>
                            <span style="background: #1E3A8A; color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.9rem; margin-left: 10px; font-weight: 700;">
                                कुल जिले: {total_districts_count}
                            </span>
                        </div>
                        <div style="font-size: 0.95rem; font-weight: 600; color: #475569;">
                            🥇 शीर्ष जिला (प्रविष्टि): <b>{top_district_name}</b> ({top_district_entries:,})
                        </div>
                    </div>
                    <div style="margin-top: 12px; display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; font-size: 0.92rem;">
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            📝 <b>कुल प्रविष्टियां:</b> {total_entries:,}
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            👥 <b>कुल प्रतिभागी:</b> {total_participants:,}
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            📸 <b>कुल फोटो की संख्या:</b> {total_photos:,}
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            🎥 <b>कुल वीडियो की संख्या:</b> {total_videos:,}
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
        else:
            district_row = district_agg[district_agg["District"] == selected_spotlight_district].iloc[0]
            d_ent_rank = district_row["orig_ent_rank"]
            d_entries = int(district_row["कुल की गई प्रविष्टि"])
            d_part_rank = district_row["orig_part_rank"]
            d_participants = int(district_row["कुल प्रतिभागियों की संख्या"])
            d_photo_rank = district_row["orig_photo_rank"]
            d_photos = int(district_row["कुल फोटो की संख्या"])
            d_video_rank = district_row["orig_video_rank"]
            d_videos = int(district_row["कुल वीडियो की संख्या"])
            
            share_pct = (d_entries / total_entries * 100) if total_entries > 0 else 0
            badge_icon = "🥇" if d_ent_rank == 1 else ("🥈" if d_ent_rank == 2 else ("🥉" if d_ent_rank == 3 else f"#{d_ent_rank}"))
            
            with col_det:
                st.markdown(f"""
                <div class="rank-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                        <div>
                            <span style="font-size: 1.65rem; font-weight: 800;">{badge_icon} जिला: {selected_spotlight_district}</span>
                            <span style="background: #1E3A8A; color: white; padding: 4px 12px; border-radius: 20px; font-size: 0.9rem; margin-left: 10px; font-weight: 700;">
                                मूल राज्य प्रविष्टि रैंक: #{d_ent_rank} / {len(overall_district_agg)}
                            </span>
                        </div>
                        <div style="font-size: 0.95rem; font-weight: 600; color: #475569;">
                            प्रदर्शित हिस्सा: <b>{share_pct:.2f}%</b>
                        </div>
                    </div>
                    <div style="margin-top: 12px; display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; font-size: 0.92rem;">
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            📝 <b>प्रविष्टियां:</b> {d_entries:,}<br>
                            <span style="color: #059669; font-weight: 700;">मूल रैंक: #{d_ent_rank}</span>
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            👥 <b>प्रतिभागी:</b> {d_participants:,}<br>
                            <span style="color: #2563EB; font-weight: 700;">मूल रैंक: #{d_part_rank}</span>
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            📸 <b>फोटो संख्या:</b> {d_photos:,}<br>
                            <span style="color: #D97706; font-weight: 700;">मूल रैंक: #{d_photo_rank}</span>
                        </div>
                        <div style="background: white; padding: 8px 12px; border-radius: 8px; border: 1px solid #E2E8F0;">
                            🎥 <b>वीडियो संख्या:</b> {d_videos:,}<br>
                            <span style="color: #7C3AED; font-weight: 700;">मूल रैंक: #{d_video_rank}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

# ==========================================
# DISTRICT RANKING TABLE
# ==========================================
col_t_title, col_t_mode = st.columns([2.2, 1.8])
with col_t_title:
    st.subheader("🏆 जिलावार प्रविष्टि एवं रैंकिंग तालिका")
    st.caption("💡 नोट: प्रत्येक संख्या के साथ कोष्ठक `( )` में जिले की **मूल राज्य रैंकिंग (Overall State Rank)** दर्शायी गई है।")
with col_t_mode:
    table_view_mode = st.radio(
        "तालिका दृश्य:",
        options=["मर्ज प्रारूप (Value + Rank)", "अलग-अलग कॉलम (Separate Columns)"],
        horizontal=True,
        index=0,
        key="table_format_toggle"
    )

# Active table to display and export based on toggle & spotlight district filter
base_district_table = merged_district_table if "मर्ज" in table_view_mode else separate_district_table

if selected_spotlight_district != ALL_DIST_OPTION:
    active_district_table = base_district_table[base_district_table["District"] == selected_spotlight_district].reset_index(drop=True)
else:
    active_district_table = base_district_table.copy()

# DISPLAY ACTIVE FILTERS INFORMATION AT TOP OF TABLE
filter_html, filter_plain_text = get_active_filter_summary(
    district_filter=selected_spotlight_district,
    count_label="प्रदर्शित जिले",
    count_value=len(active_district_table)
)
st.markdown(filter_html, unsafe_allow_html=True)

if active_district_table.empty:
    st.warning("चयनित फ़िल्टर के अनुसार कोई डेटा उपलब्ध नहीं है। कृपया फ़िल्टर रीसेट करें।")
else:
    # Display Dataframe
    st.dataframe(
        active_district_table,
        use_container_width=True,
        hide_index=True,
        height=min(600, 38 * (len(active_district_table) + 1))
    )
    
    # Download and Print buttons for District Ranking Table
    report_title = "VBSJA - जिलावार प्रविष्टि एवं रैंकिंग रिपोर्ट" if selected_spotlight_district == ALL_DIST_OPTION else f"VBSJA - {selected_spotlight_district} प्रविष्टि एवं रैंकिंग रिपोर्ट"
    st.markdown("##### 📥 तालिका निर्यात एवं प्रिंट विकल्प (Export & Print):")
    c_pdf, c_event_pdf, c_print, c_excel, c_csv = st.columns([1, 1.25, 1, 1, 1])
    
    with c_pdf:
        district_pdf_bytes = pdf_generator.create_district_pdf(
            active_district_table,
            title=report_title,
            subtitle=filter_plain_text
        )
        st.download_button(
            label="📄 मुख्य तालिका PDF (समग्र)",
            data=district_pdf_bytes,
            file_name=f"District_Ranking_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="btn_dl_district_pdf"
        )

    with c_event_pdf:
        with st.popover("📑 इवेंटवार PDF", use_container_width=True):
            st.markdown("### 📑 विस्तृत इवेंटवार रिपोर्ट (Event-wise)")
            st.caption("प्रत्येक इवेंट का अलग-अलग डेटा एवं रैंकिंग सुरक्षित करें:")
            st.info("💡 **महत्वपूर्ण जानकारी:** इस विस्तृत रिपोर्ट के **पृष्ठ 1 पर समग्र मुख्य तालिका (सभी 10 इवेंट्स का कुल योग)** शामिल है, तथा आगे **पृष्ठ 2 से 11 तक प्रत्येक इवेंट का अलग पेज** दिया गया है।")
            
            current_tbl_fmt = "separate" if "अलग" in table_view_mode else "merged"
            dist_suffix = f"_{selected_spotlight_district}" if selected_spotlight_district != ALL_DIST_OPTION else "_All_Districts"

            st.markdown("##### 📄 1. सम्पूर्ण 11-पेज विस्तृत PDF (समग्र + 10 इवेंट)")
            col_b1, col_b2 = st.columns([1, 1.2])
            with col_b1:
                btn_prep_ev = st.button("⚡ विस्तृत PDF तैयार करें", key="btn_prep_ev_pdf", use_container_width=True)
            with col_b2:
                if btn_prep_ev or st.session_state.get("cached_event_pdf"):
                    if btn_prep_ev:
                        with st.spinner("विस्तृत 11-पेज PDF तैयार हो रही है..."):
                            st.session_state["cached_event_pdf"] = pdf_generator.create_event_wise_pdf(
                                raw_df=raw_df,
                                selected_district=selected_spotlight_district,
                                categories=selected_categories if selected_categories else all_categories,
                                all_dist_option=ALL_DIST_OPTION,
                                table_format=current_tbl_fmt,
                                rank_map=rank_map
                            )
                    st.download_button(
                        label="⬇️ सम्पूर्ण विस्तृत PDF डाउनलोड करें",
                        data=st.session_state["cached_event_pdf"],
                        file_name=f"VBSJA_Event_Wise_Report{dist_suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="btn_dl_event_full_pdf"
                    )
                else:
                    st.caption("👈 बटन दबाकर PDF तैयार करें")

            st.markdown("---")
            st.markdown("##### 📦 2. प्रत्येक इवेंट की अलग PDF (ZIP बंडल)")
            col_z1, col_z2 = st.columns([1, 1.2])
            with col_z1:
                btn_prep_z = st.button("⚡ ZIP बंडल तैयार करें", key="btn_prep_ev_zip", use_container_width=True)
            with col_z2:
                if btn_prep_z or st.session_state.get("cached_event_zip"):
                    if btn_prep_z:
                        with st.spinner("सभी 11 PDF फ़ाइलों का ZIP बंडल तैयार हो रहा है..."):
                            st.session_state["cached_event_zip"] = pdf_generator.create_event_wise_zip(
                                raw_df=raw_df,
                                selected_district=selected_spotlight_district,
                                categories=selected_categories if selected_categories else all_categories,
                                all_dist_option=ALL_DIST_OPTION,
                                table_format=current_tbl_fmt,
                                rank_map=rank_map
                            )
                    st.download_button(
                        label="⬇️ ZIP फ़ाइल डाउनलोड करें",
                        data=st.session_state["cached_event_zip"],
                        file_name=f"VBSJA_Event_Separate_PDFs{dist_suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                        key="btn_dl_event_zip"
                    )
                else:
                    st.caption("👈 बटन दबाकर ZIP तैयार करें")
        
    with c_print:
        with st.popover("🖨️ प्रिंट प्रिव्यू (Print)", use_container_width=True):
            st.markdown(f"### 🖨️ {report_title} प्रिंट (Portrait)")
            st.caption("नीचे दिए गए बटन से सीधे प्रिंट करें अथवा ब्राउज़र के 'Save as PDF' विकल्प का उपयोग करें (Portrait A4):")
            printable_html = pdf_generator.generate_printable_html(
                active_district_table,
                title=report_title,
                active_filters_text=filter_plain_text
            )
            components.html(printable_html, height=450, scrolling=True)
            
    with c_excel:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            active_district_table.to_excel(writer, index=False, sheet_name='District_Ranking')
        st.download_button(
            label="📊 Excel डाउनलोड करें",
            data=buffer.getvalue(),
            file_name=f"District_Ranking_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_dl_district_excel"
        )
        
    with c_csv:
        csv_data = active_district_table.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 CSV डाउनलोड करें",
            data=csv_data,
            file_name=f"District_Ranking_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="btn_dl_district_csv"
        )

    # Comparative Visualization: Top 15 if All, or District comparison if single district selected
    if selected_spotlight_district == ALL_DIST_OPTION:
        st.markdown("#### 📊 शीर्ष 15 जिलों की प्रविष्टि तुलना")
        top15 = district_agg.head(15).sort_values(by="कुल की गई प्रविष्टि", ascending=True)
        fig_bar = px.bar(
            top15,
            x="कुल की गई प्रविष्टि",
            y="District",
            orientation="h",
            text="कुल की गई प्रविष्टि",
            color="कुल की गई प्रविष्टि",
            color_continuous_scale="Blues",
            title="शीर्ष 15 जिले (प्रविष्टियों के अनुसार)"
        )
        fig_bar.update_layout(height=460, margin=dict(l=10, r=10, t=40, b=10))
        fig_bar.update_traces(texttemplate='%{text:,}', textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.markdown(f"#### 📊 {selected_spotlight_district} की राज्य औसत एवं शीर्ष जिले से तुलना")
        avg_entries = total_entries / total_districts_count if total_districts_count > 0 else 0
        comp_df = pd.DataFrame([
            {"श्रेणी": f"चयनित जिला ({selected_spotlight_district})", "प्रविष्टियां": d_entries},
            {"श्रेणी": "राज्य औसत (State Average)", "प्रविष्टियां": int(avg_entries)},
            {"श्रेणी": f"शीर्ष जिला ({top_district_name})", "प्रविष्टियां": top_district_entries}
        ])
        fig_comp = px.bar(
            comp_df,
            x="श्रेणी",
            y="प्रविष्टियां",
            text="प्रविष्टियां",
            color="श्रेणी",
            color_discrete_sequence=["#2563EB", "#94A3B8", "#F59E0B"],
            title=f"{selected_spotlight_district} तुलनात्मक विश्लेषण"
        )
        fig_comp.update_layout(height=360, margin=dict(l=10, r=10, t=40, b=10), showlegend=False)
        fig_comp.update_traces(texttemplate='%{text:,}', textposition='outside')
        st.plotly_chart(fig_comp, use_container_width=True)

# ==========================================
# 📊 जिला इवेंटवार सारांश रैंकिंग रिपोर्ट (DISTRICT EVENT-WISE SUMMARY REPORT - AS PER SCREENSHOT)
# ==========================================
st.markdown("---")
dsumm_col_t, dsumm_col_d = st.columns([2.2, 1.8])
with dsumm_col_t:
    st.subheader("📊 जिला इवेंटवार सारांश रैंकिंग रिपोर्ट")
    st.caption("💡 चयनित जिले का प्रत्येक इवेंट केटेगरी में राज्य स्तरीय प्रदर्शन एवं रैंक, तथा अंतिम पंक्ति में समग्र (सभी इवेंट जोड़कर) कुल योग एवं मूल राज्य रैंक")

with dsumm_col_d:
    if st.session_state.get("dist_summary_chosen_district") not in pure_district_list:
        st.session_state["dist_summary_chosen_district"] = st.session_state.get("app_selected_district", default_district)
    
    chosen_summ_dist = st.selectbox(
        "इवेंट सारांश रिपोर्ट हेतु जिला चुनें (Select District):",
        options=pure_district_list,
        key="dist_summary_chosen_district",
        on_change=on_summary_district_change
    )

# Compute the summary dataframe for chosen district
df_dist_summary = pdf_generator.compute_district_event_summary_table(
    raw_df=raw_df,
    target_district=chosen_summ_dist,
    categories=selected_categories if selected_categories else all_categories
)

# Export and Print action buttons
st.markdown(f"##### 📥 जिला {chosen_summ_dist} इवेंट सारांश निर्यात एवं प्रिंट विकल्प:")
s_exp1, s_exp2, s_exp3, s_exp4 = st.columns([1.1, 1.1, 1.1, 1.3])

with s_exp1:
    s_pdf_key = f"cached_dsumm_pdf_{chosen_summ_dist}"
    btn_s_pdf = st.button(f"⚡ {chosen_summ_dist} सारांश PDF बनाएं", width="stretch", key=f"btn_p_dsumm_pdf_{chosen_summ_dist}")
    if btn_s_pdf or st.session_state.get(s_pdf_key):
        if btn_s_pdf:
            with st.spinner("इवेंट सारांश PDF तैयार हो रही है..."):
                st.session_state[s_pdf_key] = pdf_generator.create_district_event_summary_pdf(
                    raw_df=raw_df,
                    target_district=chosen_summ_dist,
                    categories=selected_categories if selected_categories else all_categories
                )
        st.download_button(
            label=f"⬇️ {chosen_summ_dist} सारांश PDF",
            data=st.session_state[s_pdf_key],
            file_name=f"District_{chosen_summ_dist}_Event_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            width="stretch",
            key=f"btn_dl_dsumm_pdf_{chosen_summ_dist}"
        )
    else:
        st.caption("👈 बटन दबाकर PDF तैयार करें")

with s_exp2:
    s_xl_key = f"cached_dsumm_xl_{chosen_summ_dist}"
    btn_s_xl = st.button(f"⚡ {chosen_summ_dist} सारांश Excel बनाएं", width="stretch", key=f"btn_p_dsumm_xl_{chosen_summ_dist}")
    if btn_s_xl or st.session_state.get(s_xl_key):
        if btn_s_xl:
            with st.spinner("इवेंट सारांश Excel तैयार हो रहा है..."):
                st.session_state[s_xl_key] = pdf_generator.create_district_event_summary_excel(
                    raw_df=raw_df,
                    target_district=chosen_summ_dist,
                    categories=selected_categories if selected_categories else all_categories
                )
        st.download_button(
            label=f"⬇️ {chosen_summ_dist} सारांश Excel",
            data=st.session_state[s_xl_key],
            file_name=f"District_{chosen_summ_dist}_Event_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key=f"btn_dl_dsumm_xl_{chosen_summ_dist}"
        )
    else:
        st.caption("👈 बटन दबाकर Excel तैयार करें")

with s_exp3:
    with st.popover(f"🖨️ {chosen_summ_dist} प्रिंट प्रिव्यू", width="stretch"):
        st.markdown(f"### 🖨️ जिला {chosen_summ_dist} — इवेंट सारांश प्रिंट")
        st.caption("नीचे दिए गए बटन से सीधे प्रिंट करें अथवा ब्राउज़र के 'Save as PDF' विकल्प का उपयोग करें (Portrait A4):")
        sec_summ_print_html = pdf_generator.generate_district_event_summary_printable_html(
            raw_df=raw_df,
            target_district=chosen_summ_dist,
            categories=selected_categories if selected_categories else all_categories
        )
        components.html(sec_summ_print_html, height=520, scrolling=True)

with s_exp4:
    with st.popover("📦 सभी जिलों का ZIP बंडल", width="stretch"):
        st.markdown("### 📦 राज्य के सभी 41 जिलों का इवेंट सारांश PDF बंडल")
        st.caption("एक क्लिक में राजस्थान के सभी 41 जिलों की इवेंटवार सारांश रिपोर्ट PDF फ़ाइलों का ZIP डाउनलोड करें:")
        btn_prep_summ_zip = st.button("⚡ सभी 41 जिलों का ZIP तैयार करें", width="stretch", key="btn_prep_dsumm_all_zip")
        if btn_prep_summ_zip or st.session_state.get("cached_dsumm_all_zip"):
            if btn_prep_summ_zip:
                with st.spinner("सभी 41 जिलों की सारांश PDF फ़ाइलों का ZIP बंडल तैयार हो रहा है..."):
                    st.session_state["cached_dsumm_all_zip"] = pdf_generator.create_all_districts_event_summary_zip(
                        raw_df=raw_df,
                        categories=selected_categories if selected_categories else all_categories
                    )
            st.download_button(
                label="⬇️ सभी 41 जिलों का ZIP डाउनलोड करें",
                data=st.session_state["cached_dsumm_all_zip"],
                file_name=f"All_Districts_Event_Summary_PDFs_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                width="stretch",
                key="btn_dl_dsumm_all_zip"
            )
        else:
            st.caption("👈 बटन दबाकर ZIP तैयार करें")

# Render Interactive Screen Table (matching screenshot 1:1)
def render_interactive_summary_table(df_t, dist_name, total_dist_count):
    curr_t = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    t_title = f"जिला {dist_name} — सभी इवेंट केटेगरी में रैंक दिनांक {curr_t}"

    rows_html = []
    for _, r in df_t.iterrows():
        is_sum = bool(r['is_summary'])
        bg_main = '#FEF3C7' if is_sum else '#FFFFFF'
        f_weight = '700' if is_sum else '500'

        tds = []
        # Event category
        tds.append(f'<td style="text-align: left; padding: 9px 12px; font-weight: {f_weight}; background-color: {bg_main}; border: 1px solid #CBD5E1;">{r["इवेंट केटेगरी"]}</td>')
        
        # Entries
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_main}; border: 1px solid #CBD5E1;">{r["प्रविष्टि"]:,}</td>')
        
        # Rank Entries
        r_ent = int(r['रैंक-प्रविष्टि'])
        ratio_ent = (r_ent - 1) / (total_dist_count - 1) if total_dist_count > 1 else 0.0
        bg_ent = pdf_generator.get_rank_3color(ratio_ent)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_ent}; color: #000; border: 1px solid #CBD5E1;">{r_ent}</td>')
        
        # Participants
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_main}; border: 1px solid #CBD5E1;">{r["प्रतिभागी"]:,}</td>')
        
        # Rank Participants
        r_part = int(r['रैंक-प्रतिभागी'])
        ratio_part = (r_part - 1) / (total_dist_count - 1) if total_dist_count > 1 else 0.0
        bg_part = pdf_generator.get_rank_3color(ratio_part)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_part}; color: #000; border: 1px solid #CBD5E1;">{r_part}</td>')
        
        # Photos
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_main}; border: 1px solid #CBD5E1;">{r["फोटो"]:,}</td>')
        
        # Rank Photos
        r_photo = int(r['रैंक-फोटो'])
        ratio_photo = (r_photo - 1) / (total_dist_count - 1) if total_dist_count > 1 else 0.0
        bg_photo = pdf_generator.get_rank_3color(ratio_photo)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_photo}; color: #000; border: 1px solid #CBD5E1;">{r_photo}</td>')
        
        # Videos
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_main}; border: 1px solid #CBD5E1;">{r["वीडियो"]:,}</td>')
        
        # Rank Videos
        r_vid = int(r['रैंक-वीडियो'])
        ratio_vid = (r_vid - 1) / (total_dist_count - 1) if total_dist_count > 1 else 0.0
        bg_vid = pdf_generator.get_rank_3color(ratio_vid)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_vid}; color: #000; border: 1px solid #CBD5E1;">{r_vid}</td>')

        rows_html.append(f'<tr>{"".join(tds)}</tr>')

    return f'''
    <div style="background: white; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden; border: 1px solid #CBD5E1; margin-top: 14px; margin-bottom: 22px;">
        <div style="background: #F8FAFC; padding: 12px 18px; text-align: center; font-weight: 700; font-size: 1.15rem; color: #1F4D77; border-bottom: 2px solid #1F4D77;">
            {t_title}
        </div>
        <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; font-size: 0.94rem;">
                <thead>
                    <tr style="background-color: #1F4D77; color: white;">
                        <th style="padding: 10px 12px; text-align: left; border: 1px solid #1F4D77; width: 26%;">इवेंट केटेगरी</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 8%;">प्रविष्टि</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 10.5%;">रैंक-प्रविष्टि</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 10%;">प्रतिभागी</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 11%;">रैंक-प्रतिभागी</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 8%;">फोटो</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 9.5%;">रैंक-फोटो</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 8%;">वीडियो</th>
                        <th style="padding: 10px 6px; text-align: center; border: 1px solid #1F4D77; width: 9%;">रैंक-वीडियो</th>
                    </tr>
                </thead>
                <tbody>
                    {"".join(rows_html)}
                </tbody>
            </table>
        </div>
    </div>
    '''

all_d_count = max(1, len(raw_df['जिला'].dropna().unique()))
st.html(render_interactive_summary_table(df_dist_summary, chosen_summ_dist, all_d_count))

# ==========================================
# 🏘️ ब्लॉक स्तरीय इवेंट रैंकिंग रिपोर्ट (BLOCK-WISE REPORT)
# ==========================================
st.markdown("---")
b_col_t, b_col_d = st.columns([2.2, 1.8])
with b_col_t:
    st.subheader("🏘️ ब्लॉक स्तरीय इवेंट रैंकिंग रिपोर्ट")
    st.caption("💡 प्रत्येक ब्लॉक की पूरे राजस्थान राज्य के 450+ ब्लॉकों में मूल राज्य रैंकिंग (Statewide Rank) एवं इवेंटवार प्रदर्शन")

with b_col_d:
    if st.session_state.get("block_report_chosen_district") not in pure_district_list:
        st.session_state["block_report_chosen_district"] = st.session_state.get("app_selected_district", default_district)
    
    chosen_block_dist = st.selectbox(
        "ब्लॉक रिपोर्ट हेतु जिला चुनें (Select District):",
        options=pure_district_list,
        key="block_report_chosen_district",
        on_change=on_block_district_change
    )

# Compute tables for chosen district
block_tables = pdf_generator.compute_block_wise_tables(
    raw_df=raw_df,
    target_district=chosen_block_dist,
    categories=selected_categories if selected_categories else all_categories
)

# Block Report Action & Export Buttons
st.markdown(f"##### 📥 जिला {chosen_block_dist} ब्लॉक रिपोर्ट निर्यात एवं प्रिंट:")
b_exp1, b_exp2, b_exp3 = st.columns([1.2, 1.2, 1.2])
with b_exp1:
    b_pdf_key = f"cached_bpdf_{chosen_block_dist}"
    btn_bpdf = st.button(f"⚡ {chosen_block_dist} ब्लॉक PDF तैयार करें", use_container_width=True, key=f"btn_p_bpdf_{chosen_block_dist}")
    if btn_bpdf or st.session_state.get(b_pdf_key):
        if btn_bpdf:
            with st.spinner("ब्लॉक PDF तैयार हो रही है..."):
                st.session_state[b_pdf_key] = pdf_generator.create_block_wise_pdf(
                    raw_df=raw_df,
                    target_district=chosen_block_dist,
                    categories=selected_categories if selected_categories else all_categories
                )
        st.download_button(
            label=f"⬇️ {chosen_block_dist} ब्लॉक PDF डाउनलोड करें",
            data=st.session_state[b_pdf_key],
            file_name=f"VBSJA_Block_Ranking_{chosen_block_dist}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"btn_dl_sec_block_pdf_{chosen_block_dist}"
        )
    else:
        st.caption("👈 बटन दबाकर ब्लॉक PDF तैयार करें")

with b_exp2:
    b_xl_key = f"cached_bxl_{chosen_block_dist}"
    btn_bxl = st.button(f"⚡ {chosen_block_dist} ब्लॉक Excel तैयार करें", use_container_width=True, key=f"btn_p_bxl_{chosen_block_dist}")
    if btn_bxl or st.session_state.get(b_xl_key):
        if btn_bxl:
            with st.spinner("ब्लॉक Excel तैयार हो रहा है..."):
                st.session_state[b_xl_key] = pdf_generator.create_block_wise_excel(
                    raw_df=raw_df,
                    target_district=chosen_block_dist,
                    categories=selected_categories if selected_categories else all_categories
                )
        st.download_button(
            label=f"⬇️ {chosen_block_dist} ब्लॉक Excel डाउनलोड करें",
            data=st.session_state[b_xl_key],
            file_name=f"VBSJA_Block_Ranking_{chosen_block_dist}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"btn_dl_sec_block_excel_{chosen_block_dist}"
        )
    else:
        st.caption("👈 बटन दबाकर ब्लॉक Excel तैयार करें")

with b_exp3:
    with st.popover(f"🖨️ {chosen_block_dist} ब्लॉक प्रिंट प्रिव्यू", use_container_width=True):
        st.markdown(f"### 🖨️ जिला {chosen_block_dist} — ब्लॉक रिपोर्ट प्रिंट")
        st.caption("नीचे दिए गए बटन से सीधे प्रिंट करें अथवा ब्राउज़र के 'Save as PDF' विकल्प का उपयोग करें (Portrait A4):")
        sec_block_print_html = pdf_generator.generate_block_wise_printable_html(
            raw_df=raw_df,
            target_district=chosen_block_dist,
            categories=selected_categories if selected_categories else all_categories
        )
        components.html(sec_block_print_html, height=520, scrolling=True)

# Display options: All tables at once vs individual dropdown
b_view_col1, b_view_col2 = st.columns([1.5, 2.5])
with b_view_col1:
    b_view_mode = st.radio(
        "प्रदर्शित करने का प्रारूप:",
        options=["📋 सभी टेबल एक साथ (All in One View)", "🔍 इवेंट अनुसार अलग-अलग देखें"],
        horizontal=True,
        key="block_view_mode_toggle"
    )

def render_interactive_block_table(tbl_info):
    t_title = tbl_info["title"]
    df_t = tbl_info["df"]
    rank_cols = ['राज्य-रैंक-प्रविष्टि', 'राज्य-रैंक-प्रतिभागी', 'राज्य-रैंक-फोटो', 'राज्य-रैंक-वीडियो']
    col_min_max = {rc: (df_t[rc].min(), df_t[rc].max()) for rc in rank_cols}

    rows_html = []
    for _, r in df_t.iterrows():
        cells = []
        cells.append(f'<td style="text-align: center; font-weight: 600; width: 40px; border: 1px solid #CBD5E1; padding: 6px 4px;">{r["क्र.सं."]}</td>')
        cells.append(f'<td style="text-align: left; font-weight: 600; padding-left: 10px; width: 110px; border: 1px solid #CBD5E1;">{r["ब्लॉक"]}</td>')
        cells.append(f'<td style="text-align: center; width: 80px; border: 1px solid #CBD5E1;">{r["प्रविष्टि"]:,}</td>')

        min_v, max_v = col_min_max['राज्य-रैंक-प्रविष्टि']
        ratio = (r['राज्य-रैंक-प्रविष्टि'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
        bg = pdf_generator.get_rank_3color(ratio)
        cells.append(f'<td style="text-align: center; background-color: {bg}; color: #000; font-weight: 600; width: 90px; border: 1px solid #CBD5E1;">{r["राज्य-रैंक-प्रविष्टि"]}</td>')

        cells.append(f'<td style="text-align: center; width: 90px; border: 1px solid #CBD5E1;">{r["प्रतिभागी"]:,}</td>')

        min_v, max_v = col_min_max['राज्य-रैंक-प्रतिभागी']
        ratio = (r['राज्य-रैंक-प्रतिभागी'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
        bg = pdf_generator.get_rank_3color(ratio)
        cells.append(f'<td style="text-align: center; background-color: {bg}; color: #000; font-weight: 600; width: 90px; border: 1px solid #CBD5E1;">{r["राज्य-रैंक-प्रतिभागी"]}</td>')

        cells.append(f'<td style="text-align: center; width: 80px; border: 1px solid #CBD5E1;">{r["फोटो"]:,}</td>')

        min_v, max_v = col_min_max['राज्य-रैंक-फोटो']
        ratio = (r['राज्य-रैंक-फोटो'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
        bg = pdf_generator.get_rank_3color(ratio)
        cells.append(f'<td style="text-align: center; background-color: {bg}; color: #000; font-weight: 600; width: 90px; border: 1px solid #CBD5E1;">{r["राज्य-रैंक-फोटो"]}</td>')

        cells.append(f'<td style="text-align: center; width: 80px; border: 1px solid #CBD5E1;">{r["वीडियो"]:,}</td>')

        min_v, max_v = col_min_max['राज्य-रैंक-वीडियो']
        ratio = (r['राज्य-रैंक-वीडियो'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
        bg = pdf_generator.get_rank_3color(ratio)
        cells.append(f'<td style="text-align: center; background-color: {bg}; color: #000; font-weight: 600; width: 90px; border: 1px solid #CBD5E1;">{r["राज्य-रैंक-वीडियो"]}</td>')

        rows_html.append(f'<tr>{"".join(cells)}</tr>')

    tbl_markup = f'''
    <div style="margin-bottom: 22px; background: white; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); overflow: hidden; border: 1px solid #CBD5E1;">
        <div style="background: #F1F5F9; padding: 10px 14px; font-weight: 800; font-size: 1.05rem; color: #1F4D77; border-bottom: 2px solid #1F4D77;">
            📁 {t_title}
        </div>
        <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; font-size: 0.92rem;">
                <thead>
                    <tr style="background-color: #1F4D77; color: white;">
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 40px;">क्र.सं.</th>
                        <th style="padding: 9px 6px; text-align: left; border: 1px solid #1F4D77; padding-left: 10px; width: 110px;">ब्लॉक</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 80px;">प्रविष्टि</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 90px;">राज्य-रैंक-<br>प्रविष्टि</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 90px;">प्रतिभागी</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 90px;">राज्य-रैंक-<br>प्रतिभागी</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 80px;">फोटो</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 90px;">राज्य-रैंक-<br>फोटो</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 80px;">वीडियो</th>
                        <th style="padding: 9px 4px; text-align: center; border: 1px solid #1F4D77; width: 90px;">राज्य-रैंक-<br>वीडियो</th>
                    </tr>
                </thead>
                <tbody>
                    {"".join(rows_html)}
                </tbody>
            </table>
        </div>
    </div>
    '''
    return tbl_markup

if "सभी" in b_view_mode:
    for t_item in block_tables:
        st.html(render_interactive_block_table(t_item))
else:
    tbl_titles = [t["title"] for t in block_tables]
    selected_tbl_title = st.selectbox("इवेंट टेबल चुनें:", options=tbl_titles, key="blk_single_tbl_select")
    selected_tbl = next((t for t in block_tables if t["title"] == selected_tbl_title), block_tables[0])
    st.html(render_interactive_block_table(selected_tbl))


# ==========================================
# FOOTER
# ==========================================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #64748B; font-size: 0.88rem; padding: 10px;'>"
    "VBSJA Portal • डेटा स्रोत: Google Spreadsheet • विकसित: Streamlit Dashboard"
    "</div>",
    unsafe_allow_html=True
)
