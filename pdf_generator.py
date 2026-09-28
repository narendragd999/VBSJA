import io
import os
import re
import subprocess
import tempfile
import zipfile
import concurrent.futures
import pandas as pd
from datetime import datetime
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

FONT_NAME = "Helvetica"
HINDI_REGISTERED = False

def register_hindi_font():
    global FONT_NAME, HINDI_REGISTERED
    if HINDI_REGISTERED:
        return FONT_NAME
    
    font_candidates = [
        ("NirmalaHindi", "C:/Windows/Fonts/Nirmala.ttc", 0),
        ("MangalHindi", "C:/Windows/Fonts/mangal.ttf", None),
        ("AparajitaHindi", "C:/Windows/Fonts/aparaj.ttf", None),
        ("ArialUnicode", "C:/Windows/Fonts/ARIALUNI.TTF", None)
    ]
    
    for fname, path, subidx in font_candidates:
        if os.path.exists(path):
            try:
                if subidx is not None:
                    pdfmetrics.registerFont(TTFont(fname, path, subfontIndex=subidx))
                else:
                    pdfmetrics.registerFont(TTFont(fname, path))
                FONT_NAME = fname
                HINDI_REGISTERED = True
                return FONT_NAME
            except Exception:
                continue
                
    FONT_NAME = "Helvetica"
    HINDI_REGISTERED = True
    return FONT_NAME

def get_system_browser() -> str:
    """
    Finds Edge or Chrome executable on Windows for 100% accurate Devanagari/Hindi font rendering.
    """
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def render_html_to_pdf(html_content: str) -> bytes:
    """
    Uses headless Chromium (Edge or Chrome) to generate a PDF with flawless Hindi Unicode
    complex text shaping (matras, ligatures, half-consonants) and crisp print typography.
    """
    browser_path = get_system_browser()
    if not browser_path:
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            html_file = os.path.join(tmpdir, "report.html")
            pdf_file = os.path.join(tmpdir, "report.pdf")

            with open(html_file, "w", encoding="utf-8") as f:
                f.write(html_content)

            file_url = f"file:///{html_file.replace(os.sep, '/')}"
            is_chrome = "chrome.exe" in browser_path.lower()
            cmd = [
                browser_path,
                "--headless=new" if is_chrome else "--headless",
                "--disable-gpu",
                "--no-pdf-header-footer",
                "--run-all-compositor-stages-before-draw",
                f"--print-to-pdf={pdf_file}",
                file_url
            ]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=25, check=True)
            if os.path.exists(pdf_file) and os.path.getsize(pdf_file) > 0:
                with open(pdf_file, "rb") as pf:
                    return pf.read()
    except Exception:
        pass

    return None

# ==========================================
# 1. DISTRICT RANKING REPORT PDF
# ==========================================
def generate_district_pdf_html(df: pd.DataFrame, title: str, subtitle: str) -> str:
    curr_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
    is_landscape = len(df.columns) > 6
    orientation = "landscape" if is_landscape else "portrait"
    
    sub_text = f"रिपोर्ट दिनांक व समय: {curr_time} | कुल रिकॉर्ड: {len(df):,}"
    if subtitle:
        sub_text += f" | {subtitle}"
        
    th_cells = "".join([f"<th>{c}</th>" for c in df.columns])
    
    tr_rows = []
    for _, row in df.iterrows():
        td_cells = []
        for c in df.columns:
            val = row[c]
            if isinstance(val, (int, float)):
                formatted = f"{int(val):,}" if isinstance(val, int) or val.is_integer() else f"{val:.2f}"
            else:
                formatted = str(val) if pd.notna(val) else ""
            td_cells.append(f"<td>{formatted}</td>")
        tr_rows.append(f"<tr>{''.join(td_cells)}</tr>")
    
    tbody = "\n".join(tr_rows)
    
    font_size = "10.5px" if is_landscape else "12px"
    th_pad = "6px 4px" if is_landscape else "8px 6px"
    td_pad = "5px 4px" if is_landscape else "7px 6px"

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        @page {{
            size: A4 {orientation};
            margin: 12mm 15mm 15mm 15mm;
        }}
        * {{
            box-sizing: border-box;
            font-family: 'Nirmala UI', 'Mukta', 'Segoe UI', Arial, sans-serif;
        }}
        body {{
            margin: 0;
            padding: 0;
            color: #0f172a;
            background: white;
            font-size: {font_size};
        }}
        .header {{
            text-align: center;
            margin-bottom: 14px;
        }}
        .header h1 {{
            color: #1e3a8a;
            margin: 0 0 5px 0;
            font-size: 20px;
            font-weight: 700;
        }}
        .header p {{
            color: #475569;
            margin: 0;
            font-size: 11.5px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            font-size: {font_size};
        }}
        th {{
            background-color: #1e3a8a;
            color: white;
            padding: {th_pad};
            text-align: center;
            font-weight: 600;
            border: 1px solid #1e3a8a;
        }}
        td {{
            padding: {td_pad};
            border: 1px solid #cbd5e1;
            text-align: center;
        }}
        tr:nth-child(even) {{
            background-color: #f8fafc;
        }}
        .footer {{
            text-align: center;
            color: #94a3b8;
            font-size: 10.5px;
            margin-top: 18px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>{title}</h1>
        <p>{sub_text}</p>
    </div>
    <div style="font-size: 11px; color: #334155; margin: 4px 0 10px 0; background: #f1f5f9; padding: 6px 12px; border-radius: 4px; border-left: 3px solid #1e3a8a; text-align: left;">
        💡 नोट: प्रत्येक संख्या के साथ कोष्ठक ( ) में जिले की मूल राज्य रैंकिंग (Overall State Rank) दर्शायी गई है।
    </div>
    <table>
        <thead>
            <tr>{th_cells}</tr>
        </thead>
        <tbody>
            {tbody}
        </tbody>
    </table>
    <div class="footer">
        VBSJA Portal • Google Spreadsheet आधारित प्रमाणित रिपोर्ट
    </div>
</body>
</html>"""
    return html

@st.cache_data(show_spinner=False, max_entries=30)
def create_district_pdf(df: pd.DataFrame, title: str = "VBSJA - जिलावार प्रविष्टि एवं रैंकिंग रिपोर्ट", subtitle: str = "") -> bytes:
    """
    Generates a clean PDF for District Ranking Table with flawless Hindi typography.
    Uses Chromium headless for exact Hindi matras and ligatures, with ReportLab fallback.
    """
    # 1. Try high-fidelity Chromium rendering (flawless Hindi ligatures & matras)
    try:
        html = generate_district_pdf_html(df, title, subtitle)
        pdf_bytes = render_html_to_pdf(html)
        if pdf_bytes:
            return pdf_bytes
    except Exception:
        pass

    # 2. Fallback to ReportLab if browser is unavailable
    return _create_district_pdf_reportlab(df, title, subtitle)

def _create_district_pdf_reportlab(df: pd.DataFrame, title: str, subtitle: str) -> bytes:
    font = register_hindi_font()
    buf = io.BytesIO()
    is_landscape = len(df.columns) > 6
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4) if is_landscape else A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName=font,
        fontSize=15,
        leading=19,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1
    )
    sub_style = ParagraphStyle(
        'DocSub',
        parent=styles['Normal'],
        fontName=font,
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#475569'),
        alignment=1
    )
    rank_note_style = ParagraphStyle(
        'DocRankNote',
        parent=styles['Normal'],
        fontName=font,
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#334155'),
        alignment=0
    )
    header_cell_style = ParagraphStyle(
        'HeaderCell',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7.5,
        leading=10,
        textColor=colors.white,
        alignment=1
    )

    elements.append(Paragraph(title, title_style))
    curr_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
    sub_text = f"रिपोर्ट दिनांक व समय: {curr_time} | कुल जिले: {len(df)}"
    if subtitle:
        sub_text += f" | {subtitle}"
    elements.append(Paragraph(sub_text, sub_style))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("💡 नोट: प्रत्येक संख्या के साथ कोष्ठक ( ) में जिले की मूल राज्य रैंकिंग (Overall State Rank) दर्शायी गई है।", rank_note_style))
    elements.append(Spacer(1, 8))

    headers = [Paragraph(str(c), header_cell_style) for c in df.columns]
    table_data = [headers]

    for _, row in df.iterrows():
        r_vals = []
        for c in df.columns:
            val = row[c]
            if isinstance(val, (int, float)):
                r_vals.append(f"{int(val):,}" if isinstance(val, int) or val.is_integer() else f"{val:.2f}")
            else:
                r_vals.append(str(val) if pd.notna(val) else "")
        table_data.append(r_vals)

    col_widths = [130, 105, 130, 90, 90] if len(df.columns) == 5 else None

    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), font),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))
    elements.append(t)

    footer_style = ParagraphStyle(
        'DocFooter',
        parent=styles['Normal'],
        fontName=font,
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#94A3B8'),
        alignment=1
    )
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("VBSJA Portal • Google Spreadsheet आधारित प्रमाणित रिपोर्ट", footer_style))

    doc.build(elements)
    return buf.getvalue()

# ==========================================
# 2. EVENT-WISE REPORT PDF & ZIP
# ==========================================
def generate_event_wise_pdf_html(
    raw_df: pd.DataFrame, 
    selected_district: str = None, 
    categories: list = None, 
    all_dist_option: str = "🌟 सभी जिले (All Districts)",
    table_format: str = "merged",
    rank_map: dict = None,
    include_summary: bool = True,
    only_summary: bool = False
) -> str:
    curr_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
    if not categories:
        categories = sorted([c for c in raw_df["इवेंट केटेगरी"].dropna().unique() if str(c).strip()])

    # Ensure standardized photo and video columns exist
    if "कुल फोटो की संख्या" in raw_df.columns:
        raw_df["कुल फोटो की संख्या"] = pd.to_numeric(raw_df["कुल फोटो की संख्या"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल photo"] = raw_df["कुल फोटो की संख्या"]
    elif "कुल photo" in raw_df.columns:
        raw_df["कुल photo"] = pd.to_numeric(raw_df["कुल photo"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल फोटो की संख्या"] = raw_df["कुल photo"]
    else:
        raw_df["कुल फोटो की संख्या"] = 0
        raw_df["कुल photo"] = 0

    if "कुल वीडियो की संख्या" in raw_df.columns:
        raw_df["कुल वीडियो की संख्या"] = pd.to_numeric(raw_df["कुल वीडियो की संख्या"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल video"] = raw_df["कुल वीडियो की संख्या"]
    elif "कुल video" in raw_df.columns:
        raw_df["कुल video"] = pd.to_numeric(raw_df["कुल video"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल वीडियो की संख्या"] = raw_df["कुल video"]
    else:
        raw_df["कुल वीडियो की संख्या"] = 0
        raw_df["कुल video"] = 0

    # Compute overall original state summary unconditionally
    overall_agg = raw_df.groupby("जिला", as_index=False).agg({
        "कुल की गई प्रविष्टि": "sum",
        "कुल प्रतिभागियों की संख्या": "sum",
        "कुल फोटो की संख्या": "sum",
        "कुल वीडियो की संख्या": "sum"
    })
    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        overall_agg[c] = pd.to_numeric(overall_agg[c], errors='coerce').fillna(0).astype(int)

    overall_agg["orig_ent_rank"] = overall_agg["कुल की गई प्रविष्टि"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_part_rank"] = overall_agg["कुल प्रतिभागियों की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_photo_rank"] = overall_agg["कुल फोटो की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_video_rank"] = overall_agg["कुल वीडियो की संख्या"].rank(ascending=False, method='min').astype(int)

    overall_agg["कुल photo"] = overall_agg["कुल फोटो की संख्या"]
    overall_agg["कुल video"] = overall_agg["कुल वीडियो की संख्या"]

    if rank_map is None:
        rank_map = overall_agg.set_index("जिला")[["orig_ent_rank", "orig_part_rank", "orig_photo_rank", "orig_video_rank"]].to_dict("index")

    is_separate = "separate" in str(table_format).lower() or "अलग" in str(table_format)

    if is_separate:
        thead_html = """
                <tr>
                    <th style="width: 40px;">रैंक</th>
                    <th style="width: 115px; text-align: left; padding-left: 8px;">District</th>
                    <th>कुल की गई प्रविष्टि</th>
                    <th>प्रविष्टि मूल रैंक</th>
                    <th>कुल प्रतिभागी संख्या</th>
                    <th>प्रतिभागी मूल रैंक</th>
                    <th>कुल फोटो की संख्या</th>
                    <th>फोटो मूल रैंक</th>
                    <th>कुल वीडियो की संख्या</th>
                    <th>वीडियो मूल रैंक</th>
                </tr>
        """
    else:
        thead_html = """
                <tr>
                    <th style="width: 45px;">रैंक</th>
                    <th style="width: 140px; text-align: left; padding-left: 10px;">District</th>
                    <th>कुल की गई प्रविष्टि</th>
                    <th>कुल प्रतिभागियों की संख्या</th>
                    <th>कुल फोटो की संख्या</th>
                    <th>कुल वीडियो की संख्या</th>
                </tr>
        """

    pages_html = []
    should_include_summary = only_summary or (include_summary and len(categories) > 1)
    if only_summary:
        total_pages_count = 1
    elif should_include_summary:
        total_pages_count = len(categories) + 1
    else:
        total_pages_count = len(categories)

    # ----------------------------------------------------
    # PAGE 1: OVERALL STATE SUMMARY (समग्र - सभी 10 इवेंट श्रेणियां जोड़कर)
    # ----------------------------------------------------
    if should_include_summary:
        sorted_overall = overall_agg.sort_values(by="कुल की गई प्रविष्टि", ascending=False).reset_index(drop=True)
        sorted_overall["रैंक"] = range(1, len(sorted_overall) + 1)

        overall_tot_entries = int(sorted_overall["कुल की गई प्रविष्टि"].sum())
        overall_tot_part = int(sorted_overall["कुल प्रतिभागियों की संख्या"].sum())
        overall_tot_photos = int(sorted_overall["कुल फोटो की संख्या"].sum())
        overall_tot_videos = int(sorted_overall["कुल वीडियो की संख्या"].sum())
        overall_top_d = sorted_overall.iloc[0]["जिला"] if not sorted_overall.empty else "N/A"
        overall_top_v = sorted_overall.iloc[0]["कुल की गई प्रविष्टि"] if not sorted_overall.empty else 0

        overall_spotlight_box = ""
        if selected_district and selected_district != all_dist_option:
            s_matches = sorted_overall[sorted_overall["जिला"] == selected_district]
            if not s_matches.empty:
                s_row = s_matches.iloc[0]
                overall_spotlight_box = f"""
                <div style="background: #fef3c7; border: 1.5px solid #f59e0b; border-radius: 6px; padding: 6px 12px; margin-bottom: 8px; font-size: 11px; color: #92400e; line-height: 1.5;">
                    🎯 <strong>चयनित जिला ({selected_district}):</strong> 
                    कुल प्रविष्टियां: <strong>{s_row['कुल की गई प्रविष्टि']:,} (#{s_row['orig_ent_rank']})</strong> &nbsp;|&nbsp; 
                    कुल प्रतिभागी: <strong>{s_row['कुल प्रतिभागियों की संख्या']:,} (#{s_row['orig_part_rank']})</strong> &nbsp;|&nbsp; 
                    कुल फोटो: <strong>{s_row['कुल फोटो की संख्या']:,} (#{s_row['orig_photo_rank']})</strong> &nbsp;|&nbsp; 
                    कुल वीडियो: <strong>{s_row['कुल वीडियो की संख्या']:,} (#{s_row['orig_video_rank']})</strong>
                </div>
                """

        overall_rows = []
        for _, row in sorted_overall.iterrows():
            d_name = row["जिला"]
            ent_r = row["orig_ent_rank"]
            part_r = row["orig_part_rank"]
            photo_r = row["orig_photo_rank"]
            video_r = row["orig_video_rank"]

            is_spot = (selected_district and selected_district != all_dist_option and d_name == selected_district)
            row_style = ' style="background-color: #fef9c3; font-weight: bold; color: #1e3a8a;"' if is_spot else ""

            if is_separate:
                overall_rows.append(f"""
                <tr{row_style}>
                    <td>#{row['रैंक']}</td>
                    <td style="text-align: left; padding-left: 8px;">{d_name}</td>
                    <td>{row['कुल की गई प्रविष्टि']:,}</td>
                    <td>{ent_r}</td>
                    <td>{row['कुल प्रतिभागियों की संख्या']:,}</td>
                    <td>{part_r}</td>
                    <td>{row['कुल फोटो की संख्या']:,}</td>
                    <td>{photo_r}</td>
                    <td>{row['कुल वीडियो की संख्या']:,}</td>
                    <td>{video_r}</td>
                </tr>
                """)
            else:
                overall_rows.append(f"""
                <tr{row_style}>
                    <td>#{row['रैंक']}</td>
                    <td style="text-align: left; padding-left: 10px;">{d_name}</td>
                    <td>{row['कुल की गई प्रविष्टि']:,} ({ent_r})</td>
                    <td>{row['कुल प्रतिभागियों की संख्या']:,} ({part_r})</td>
                    <td>{row['कुल फोटो की संख्या']:,} ({photo_r})</td>
                    <td>{row['कुल वीडियो की संख्या']:,} ({video_r})</td>
                </tr>
                """)

        summary_note = f"💡 नोट: यह तालिका सभी {len(categories)} इवेंट श्रेणियों का कुल योग (Overall Summary) दर्शाती है। प्रत्येक संख्या के साथ कोष्ठक ( ) में मूल राज्य रैंकिंग दर्शायी गई है।"
        if not only_summary and len(categories) > 1:
            summary_note += " आगामी पृष्ठों पर प्रत्येक इवेंट श्रेणी का अलग-अलग डेटा दिया गया है।"

        page_1_html = f"""
        <div class="event-page">
            <div class="header">
                <h1>VBSJA - समग्र जिलावार प्रविष्टि एवं रैंकिंग रिपोर्ट</h1>
                <p>दिनांक व समय: {curr_time} | कुल जिले: {len(sorted_overall)} | कुल इवेंट श्रेणियां: {len(categories)}</p>
            </div>
            
            <div class="event-title-banner" style="background: #1e3a8a; color: white;">
                📊 <strong>समग्र जिला रैंकिंग (सभी {len(categories)} इवेंट श्रेणियां जोड़कर)</strong>
            </div>

            <div class="kpi-banner">
                <span>📝 <strong>कुल प्रविष्टियां:</strong> {overall_tot_entries:,}</span> &nbsp;|&nbsp;
                <span>👥 <strong>कुल प्रतिभागी:</strong> {overall_tot_part:,}</span> &nbsp;|&nbsp;
                <span>📸 <strong>Photos:</strong> {overall_tot_photos:,}</span> &nbsp;|&nbsp;
                <span>🎥 <strong>Videos:</strong> {overall_tot_videos:,}</span> &nbsp;|&nbsp;
                <span>🥇 <strong>शीर्ष जिला:</strong> {overall_top_d} ({overall_top_v:,})</span>
            </div>

            {overall_spotlight_box}

            <div class="rank-note">
                {summary_note}
            </div>

            <table>
                <thead>
                    {thead_html}
                </thead>
                <tbody>
                    {''.join(overall_rows)}
                </tbody>
            </table>
            
            <div class="page-footer">
                VBSJA Portal • समग्र रिपोर्ट • पृष्ठ 1 / {total_pages_count}
            </div>
        </div>
        """
        pages_html.append(page_1_html)

    if not only_summary:
        for idx, cat in enumerate(categories, 1):
            cat_df = raw_df[raw_df["इवेंट केटेगरी"] == cat]
            if cat_df.empty:
                continue

            cat_agg = cat_df.groupby("जिला", as_index=False).agg({
                "कुल की गई प्रविष्टि": "sum",
                "कुल प्रतिभागियों की संख्या": "sum",
                "कुल फोटो की संख्या": "sum",
                "कुल वीडियो की संख्या": "sum"
            })
            for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
                cat_agg[c] = pd.to_numeric(cat_agg[c], errors='coerce').fillna(0).astype(int)

            cat_agg["कुल photo"] = cat_agg["कुल फोटो की संख्या"]
            cat_agg["कुल video"] = cat_agg["कुल वीडियो की संख्या"]

            cat_agg = cat_agg.sort_values(by="कुल की गई प्रविष्टि", ascending=False).reset_index(drop=True)
            cat_agg["इवेंट रैंक"] = range(1, len(cat_agg) + 1)

            cat_tot_entries = int(cat_agg["कुल की गई प्रविष्टि"].sum())
            cat_tot_part = int(cat_agg["कुल प्रतिभागियों की संख्या"].sum())
            cat_tot_photos = int(cat_agg["कुल फोटो की संख्या"].sum())
            cat_tot_videos = int(cat_agg["कुल वीडियो की संख्या"].sum())
            top_d = cat_agg.iloc[0]["जिला"] if not cat_agg.empty else "N/A"
            top_v = cat_agg.iloc[0]["कुल की गई प्रविष्टि"] if not cat_agg.empty else 0

            # Spotlight callout box
            spotlight_box_html = ""
            if selected_district and selected_district != all_dist_option:
                dist_matches = cat_agg[cat_agg["जिला"] == selected_district]
                if not dist_matches.empty:
                    d_row = dist_matches.iloc[0]
                    d_r = rank_map.get(selected_district, {})
                    d_ent_r = d_r.get("orig_ent_rank", "-")
                    d_part_r = d_r.get("orig_part_rank", "-")
                    d_photo_r = d_r.get("orig_photo_rank", "-")
                    d_video_r = d_r.get("orig_video_rank", "-")

                    spotlight_box_html = f"""
                    <div style="background: #fef3c7; border: 1.5px solid #f59e0b; border-radius: 6px; padding: 6px 12px; margin-bottom: 8px; font-size: 11px; color: #92400e; line-height: 1.5;">
                        🎯 <strong>चयनित जिला ({selected_district}):</strong> 
                        प्रविष्टि: <strong>{d_row['कुल की गई प्रविष्टि']:,} ({d_ent_r})</strong> (इवेंट रैंक: <strong>#{d_row['इवेंट रैंक']}</strong>/{len(cat_agg)}) &nbsp;|&nbsp; 
                        प्रतिभागी: <strong>{d_row['कुल प्रतिभागियों की संख्या']:,} ({d_part_r})</strong> &nbsp;|&nbsp; 
                        फोटो: <strong>{d_row['कुल फोटो की संख्या']:,} ({d_photo_r})</strong> &nbsp;|&nbsp; 
                        वीडियो: <strong>{d_row['कुल वीडियो की संख्या']:,} ({d_video_r})</strong>
                    </div>
                    """

            tr_rows = []
            for _, row in cat_agg.iterrows():
                d_name = row["जिला"]
                d_ranks = rank_map.get(d_name, {})
                ent_r = d_ranks.get("orig_ent_rank", "-")
                part_r = d_ranks.get("orig_part_rank", "-")
                photo_r = d_ranks.get("orig_photo_rank", "-")
                video_r = d_ranks.get("orig_video_rank", "-")

                is_spot = (selected_district and selected_district != all_dist_option and d_name == selected_district)
                row_style = ' style="background-color: #fef9c3; font-weight: bold; color: #1e3a8a;"' if is_spot else ""

                if is_separate:
                    tr_rows.append(f"""
                    <tr{row_style}>
                        <td>#{row['इवेंट रैंक']}</td>
                        <td style="text-align: left; padding-left: 8px;">{d_name}</td>
                        <td>{row['कुल की गई प्रविष्टि']:,}</td>
                        <td>{ent_r}</td>
                        <td>{row['कुल प्रतिभागियों की संख्या']:,}</td>
                        <td>{part_r}</td>
                        <td>{row['कुल फोटो की संख्या']:,}</td>
                        <td>{photo_r}</td>
                        <td>{row['कुल वीडियो की संख्या']:,}</td>
                        <td>{video_r}</td>
                    </tr>
                    """)
                else:
                    tr_rows.append(f"""
                    <tr{row_style}>
                        <td>#{row['इवेंट रैंक']}</td>
                        <td style="text-align: left; padding-left: 10px;">{d_name}</td>
                        <td>{row['कुल की गई प्रविष्टि']:,} ({ent_r})</td>
                        <td>{row['कुल प्रतिभागियों की संख्या']:,} ({part_r})</td>
                        <td>{row['कुल फोटो की संख्या']:,} ({photo_r})</td>
                        <td>{row['कुल वीडियो की संख्या']:,} ({video_r})</td>
                    </tr>
                    """)

            table_rows_html = "".join(tr_rows)
            page_num = (idx + 1) if should_include_summary else idx

            page_div = f"""
            <div class="event-page">
                <div class="header">
                    <h1>VBSJA - इवेंटवार जिला रैंकिंग रिपोर्ट</h1>
                    <p>दिनांक व समय: {curr_time} | कुल इवेंट श्रेणियां: {len(categories)}</p>
                </div>
                
                <div class="event-title-banner">
                    📁 <strong>इवेंट श्रेणी #{idx}:</strong> {cat}
                </div>

                <div class="kpi-banner">
                    <span>📝 <strong>कुल प्रविष्टियां:</strong> {cat_tot_entries:,}</span> &nbsp;|&nbsp;
                    <span>👥 <strong>कुल प्रतिभागी:</strong> {cat_tot_part:,}</span> &nbsp;|&nbsp;
                    <span>📸 <strong>Photos:</strong> {cat_tot_photos:,}</span> &nbsp;|&nbsp;
                    <span>🎥 <strong>Videos:</strong> {cat_tot_videos:,}</span> &nbsp;|&nbsp;
                    <span>🥇 <strong>शीर्ष जिला:</strong> {top_d} ({top_v:,})</span>
                </div>

                {spotlight_box_html}

                <div class="rank-note">
                    💡 नोट: यह तालिका केवल इवेंट श्रेणी '<strong>{cat}</strong>' का डेटा दर्शाती है। प्रथम कॉलम 'रैंक' इस इवेंट की प्रविष्टि रैंकिंग है, तथा कोष्ठक ( ) में समग्र मूल राज्य रैंक दर्शायी गई है। सभी इवेंट्स का कुल योग पृष्ठ 1 पर देखें।
                </div>

                <table>
                    <thead>
                        {thead_html}
                    </thead>
                    <tbody>
                        {table_rows_html}
                    </tbody>
                </table>
                
                <div class="page-footer">
                    VBSJA Portal • इवेंटवार रिपोर्ट • पृष्ठ {page_num} / {total_pages_count}
                </div>
            </div>
            """
            pages_html.append(page_div)

    all_pages = "\n".join(pages_html)

    table_font_size = "8.5px" if is_separate else "9px"
    th_td_pad = "2px 2px" if is_separate else "2px 3px"

    return f"""<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <style>
        @page {{
            size: A4 portrait;
            margin: 6mm 10mm 6mm 10mm;
        }}
        * {{
            box-sizing: border-box;
            font-family: 'Nirmala UI', 'Mukta', 'Segoe UI', Arial, sans-serif;
        }}
        body {{
            margin: 0;
            padding: 0;
            color: #0f172a;
            background: white;
            font-size: 9.5px;
        }}
        .event-page {{
            break-after: page;
            page-break-after: always;
            padding-bottom: 2px;
        }}
        .header {{
            text-align: center;
            margin-bottom: 4px;
        }}
        .header h1 {{
            color: #1e3a8a;
            margin: 0 0 2px 0;
            font-size: 15px;
            font-weight: 700;
        }}
        .header p {{
            color: #64748b;
            margin: 0;
            font-size: 10px;
        }}
        .event-title-banner {{
            background: #1e3a8a;
            color: white;
            padding: 4px 10px;
            border-radius: 4px;
            font-size: 11px;
            margin-bottom: 4px;
        }}
        .kpi-banner {{
            background: #f1f5f9;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            padding: 4px 8px;
            margin-bottom: 4px;
            font-size: 9.5px;
            color: #1e293b;
        }}
        .rank-note {{
            font-size: 8.5px;
            color: #334155;
            margin: 2px 0 4px 0;
            background: #f1f5f9;
            padding: 3px 6px;
            border-radius: 4px;
            border-left: 3px solid #1e3a8a;
            line-height: 1.3;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: {table_font_size};
        }}
        th {{
            background-color: #1e3a8a;
            color: white;
            padding: {th_td_pad};
            text-align: center;
            font-weight: 600;
            border: 1px solid #1e3a8a;
            line-height: 1.15;
        }}
        td {{
            padding: {th_td_pad};
            border: 1px solid #cbd5e1;
            text-align: center;
            line-height: 1.15;
        }}
        tr:nth-child(even) {{
            background-color: #f8fafc;
        }}
        .page-footer {{
            text-align: center;
            color: #94a3b8;
            font-size: 8.5px;
            margin-top: 4px;
        }}
    </style>
</head>
<body>
    {all_pages}
</body>
</html>"""

@st.cache_data(show_spinner=False, max_entries=20)
def create_event_wise_pdf(
    raw_df: pd.DataFrame, 
    selected_district: str = None, 
    categories: list = None, 
    all_dist_option: str = "🌟 सभी जिले (All Districts)",
    table_format: str = "merged",
    rank_map: dict = None,
    include_summary: bool = True,
    only_summary: bool = False
) -> bytes:
    """
    Generates a multi-page A4 Portrait PDF report with each Event Category on a separate page.
    Uses Chromium for flawless Hindi typography with ReportLab fallback.
    """
    # 1. Try high-fidelity Chromium rendering (flawless Hindi matras and ligatures)
    try:
        html = generate_event_wise_pdf_html(raw_df, selected_district, categories, all_dist_option, table_format, rank_map, include_summary, only_summary)
        pdf_bytes = render_html_to_pdf(html)
        if pdf_bytes:
            return pdf_bytes
    except Exception:
        pass

    # 2. Fallback to ReportLab
    return _create_event_wise_pdf_reportlab(raw_df, selected_district, categories, all_dist_option, table_format, rank_map, include_summary, only_summary)

def _create_event_wise_pdf_reportlab(
    raw_df: pd.DataFrame, 
    selected_district: str = None, 
    categories: list = None, 
    all_dist_option: str = "🌟 सभी जिले (All Districts)",
    table_format: str = "merged",
    rank_map: dict = None,
    include_summary: bool = True,
    only_summary: bool = False
) -> bytes:
    font = register_hindi_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=25,
        rightMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'EvTitle',
        parent=styles['Heading1'],
        fontName=font,
        fontSize=15,
        leading=19,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1
    )
    cat_banner_style = ParagraphStyle(
        'EvBanner',
        parent=styles['Heading2'],
        fontName=font,
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#0F172A'),
        alignment=0
    )
    sub_style = ParagraphStyle(
        'EvSub',
        parent=styles['Normal'],
        fontName=font,
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor('#475569'),
        alignment=1
    )
    kpi_bar_style = ParagraphStyle(
        'EvKpi',
        parent=styles['Normal'],
        fontName=font,
        fontSize=8,
        leading=11,
        textColor=colors.HexColor('#1E293B'),
        alignment=0
    )
    rank_note_style = ParagraphStyle(
        'EvRankNote',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor('#334155'),
        alignment=0
    )
    table_header_style = ParagraphStyle(
        'EvTh',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7.5,
        leading=9.5,
        textColor=colors.white,
        alignment=1
    )
    table_cell_style = ParagraphStyle(
        'EvTd',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7,
        leading=9,
        textColor=colors.HexColor('#1E293B'),
        alignment=1
    )
    table_cell_bold = ParagraphStyle(
        'EvTdBold',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7,
        leading=9,
        textColor=colors.HexColor('#1E3A8A'),
        alignment=1
    )

    if not categories:
        categories = sorted([c for c in raw_df["इवेंट केटेगरी"].dropna().unique() if str(c).strip()])

    # Standardize photo and video columns
    if "कुल फोटो की संख्या" in raw_df.columns:
        raw_df["कुल फोटो की संख्या"] = pd.to_numeric(raw_df["कुल फोटो की संख्या"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल photo"] = raw_df["कुल फोटो की संख्या"]
    elif "कुल photo" in raw_df.columns:
        raw_df["कुल photo"] = pd.to_numeric(raw_df["कुल photo"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल फोटो की संख्या"] = raw_df["कुल photo"]
    else:
        raw_df["कुल फोटो की संख्या"] = 0
        raw_df["कुल photo"] = 0

    if "कुल वीडियो की संख्या" in raw_df.columns:
        raw_df["कुल वीडियो की संख्या"] = pd.to_numeric(raw_df["कुल वीडियो की संख्या"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल video"] = raw_df["कुल वीडियो की संख्या"]
    elif "कुल video" in raw_df.columns:
        raw_df["कुल video"] = pd.to_numeric(raw_df["कुल video"], errors='coerce').fillna(0).astype(int)
        raw_df["कुल वीडियो की संख्या"] = raw_df["कुल video"]
    else:
        raw_df["कुल वीडियो की संख्या"] = 0
        raw_df["कुल video"] = 0

    # Compute overall original state summary unconditionally
    overall_agg = raw_df.groupby("जिला", as_index=False).agg({
        "कुल की गई प्रविष्टि": "sum",
        "कुल प्रतिभागियों की संख्या": "sum",
        "कुल फोटो की संख्या": "sum",
        "कुल वीडियो की संख्या": "sum"
    })
    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        overall_agg[c] = pd.to_numeric(overall_agg[c], errors='coerce').fillna(0).astype(int)

    overall_agg["orig_ent_rank"] = overall_agg["कुल की गई प्रविष्टि"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_part_rank"] = overall_agg["कुल प्रतिभागियों की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_photo_rank"] = overall_agg["कुल फोटो की संख्या"].rank(ascending=False, method='min').astype(int)
    overall_agg["orig_video_rank"] = overall_agg["कुल वीडियो की संख्या"].rank(ascending=False, method='min').astype(int)

    overall_agg["कुल photo"] = overall_agg["कुल फोटो की संख्या"]
    overall_agg["कुल video"] = overall_agg["कुल वीडियो की संख्या"]

    if rank_map is None:
        rank_map = overall_agg.set_index("जिला")[["orig_ent_rank", "orig_part_rank", "orig_photo_rank", "orig_video_rank"]].to_dict("index")

    is_separate = "separate" in str(table_format).lower() or "अलग" in str(table_format)
    curr_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")

    # ----------------------------------------------------
    # PAGE 1: OVERALL STATE SUMMARY (समग्र - सभी 10 इवेंट श्रेणियां जोड़कर)
    # ----------------------------------------------------
    should_include_summary = only_summary or (include_summary and len(categories) > 1)
    if should_include_summary:
        sorted_overall = overall_agg.sort_values(by="कुल की गई प्रविष्टि", ascending=False).reset_index(drop=True)
        sorted_overall["रैंक"] = range(1, len(sorted_overall) + 1)

        overall_tot_entries = int(sorted_overall["कुल की गई प्रविष्टि"].sum())
        overall_tot_part = int(sorted_overall["कुल प्रतिभागियों की संख्या"].sum())
        overall_tot_photos = int(sorted_overall["कुल फोटो की संख्या"].sum())
        overall_tot_videos = int(sorted_overall["कुल वीडियो की संख्या"].sum())
        overall_top_d = sorted_overall.iloc[0]["जिला"] if not sorted_overall.empty else "N/A"
        overall_top_v = sorted_overall.iloc[0]["कुल की गई प्रविष्टि"] if not sorted_overall.empty else 0

        elements.append(Paragraph("VBSJA - समग्र जिलावार प्रविष्टि एवं रैंकिंग रिपोर्ट", title_style))
        elements.append(Paragraph(f"रिपोर्ट समय: {curr_time} | कुल जिले: {len(sorted_overall)} | कुल इवेंट श्रेणियां: {len(categories)}", sub_style))
        elements.append(Spacer(1, 6))

        elements.append(Paragraph(f"<b>📊 समग्र जिला रैंकिंग (सभी {len(categories)} इवेंट श्रेणियां जोड़कर)</b>", cat_banner_style))
        elements.append(Spacer(1, 4))

        kpi_all_text = (
            f"<b>कुल प्रविष्टियां:</b> {overall_tot_entries:,} &nbsp;|&nbsp; "
            f"<b>कुल प्रतिभागी:</b> {overall_tot_part:,} &nbsp;|&nbsp; "
            f"<b>Photos:</b> {overall_tot_photos:,} &nbsp;|&nbsp; "
            f"<b>Videos:</b> {overall_tot_videos:,} &nbsp;|&nbsp; "
            f"<b>🥇 शीर्ष जिला:</b> {overall_top_d} ({overall_top_v:,})"
        )
        kpi_all_table = Table([[Paragraph(kpi_all_text, kpi_bar_style)]], colWidths=[545])
        kpi_all_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F1F5F9')),
            ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
            ('PADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(kpi_all_table)
        elements.append(Spacer(1, 4))

        if selected_district and selected_district != all_dist_option:
            dist_matches = sorted_overall[sorted_overall["जिला"] == selected_district]
            if not dist_matches.empty:
                d_row = dist_matches.iloc[0]
                dist_info_text = (
                    f"🎯 <b>चयनित जिला ({selected_district}):</b> "
                    f"कुल प्रविष्टियां: <b>{d_row['कुल की गई प्रविष्टि']:,} (#{d_row['orig_ent_rank']})</b> | "
                    f"प्रतिभागी: {d_row['कुल प्रतिभागियों की संख्या']:,} (#{d_row['orig_part_rank']}) | "
                    f"फोटो: {d_row['कुल फोटो की संख्या']:,} (#{d_row['orig_photo_rank']}) | "
                    f"वीडियो: {d_row['कुल वीडियो की संख्या']:,} (#{d_row['orig_video_rank']})"
                )
                dist_box = Table([[Paragraph(dist_info_text, kpi_bar_style)]], colWidths=[545])
                dist_box.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FEF3C7')),
                    ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#F59E0B')),
                    ('PADDING', (0, 0), (-1, -1), 4),
                ]))
                elements.append(dist_box)
                elements.append(Spacer(1, 4))

        elements.append(Paragraph(f"💡 नोट: यह तालिका सभी {len(categories)} इवेंट श्रेणियों का कुल योग (Overall Summary) दर्शाती है। प्रत्येक संख्या के साथ कोष्ठक ( ) में मूल राज्य रैंकिंग दर्शायी गई है।", rank_note_style))
        elements.append(Spacer(1, 4))

        if is_separate:
            all_headers = [
                Paragraph("रैंक", table_header_style),
                Paragraph("District", table_header_style),
                Paragraph("कुल की गई प्रविष्टि", table_header_style),
                Paragraph("प्रविष्टि मूल रैंक", table_header_style),
                Paragraph("कुल प्रतिभागी संख्या", table_header_style),
                Paragraph("प्रतिभागी मूल रैंक", table_header_style),
                Paragraph("कुल फोटो की संख्या", table_header_style),
                Paragraph("फोटो मूल रैंक", table_header_style),
                Paragraph("कुल वीडियो की संख्या", table_header_style),
                Paragraph("वीडियो मूल रैंक", table_header_style)
            ]
            all_col_widths = [35, 90, 60, 45, 65, 45, 55, 45, 55, 45]
        else:
            all_headers = [
                Paragraph("रैंक", table_header_style),
                Paragraph("District", table_header_style),
                Paragraph("कुल की गई प्रविष्टि", table_header_style),
                Paragraph("कुल प्रतिभागियों की संख्या", table_header_style),
                Paragraph("कुल फोटो की संख्या", table_header_style),
                Paragraph("कुल वीडियो की संख्या", table_header_style)
            ]
            all_col_widths = [40, 110, 100, 115, 90, 90]

        all_table_rows = [all_headers]
        for _, row in sorted_overall.iterrows():
            d_name = row["जिला"]
            ent_r = row["orig_ent_rank"]
            part_r = row["orig_part_rank"]
            photo_r = row["orig_photo_rank"]
            video_r = row["orig_video_rank"]

            is_spotlight = (selected_district and selected_district != all_dist_option and d_name == selected_district)
            st_style = table_cell_bold if is_spotlight else table_cell_style

            if is_separate:
                all_table_rows.append([
                    Paragraph(f"#{row['रैंक']}", st_style),
                    Paragraph(str(d_name), st_style),
                    Paragraph(f"{row['कुल की गई प्रविष्टि']:,}", st_style),
                    Paragraph(str(ent_r), st_style),
                    Paragraph(f"{row['कुल प्रतिभागियों की संख्या']:,}", st_style),
                    Paragraph(str(part_r), st_style),
                    Paragraph(f"{row['कुल फोटो की संख्या']:,}", st_style),
                    Paragraph(str(photo_r), st_style),
                    Paragraph(f"{row['कुल वीडियो की संख्या']:,}", st_style),
                    Paragraph(str(video_r), st_style)
                ])
            else:
                all_table_rows.append([
                    Paragraph(f"#{row['रैंक']}", st_style),
                    Paragraph(str(d_name), st_style),
                    Paragraph(f"{row['कुल की गई प्रविष्टि']:,} ({ent_r})", st_style),
                    Paragraph(f"{row['कुल प्रतिभागियों की संख्या']:,} ({part_r})", st_style),
                    Paragraph(f"{row['कुल फोटो की संख्या']:,} ({photo_r})", st_style),
                    Paragraph(f"{row['कुल वीडियो की संख्या']:,} ({video_r})", st_style)
                ])

        all_event_table = Table(all_table_rows, colWidths=all_col_widths, repeatRows=1)
        t_all_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), font),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
        ]
        all_event_table.setStyle(TableStyle(t_all_style))
        elements.append(all_event_table)

    if only_summary:
        doc.build(elements)
        return buf.getvalue()

    if should_include_summary and len(categories) > 0:
        elements.append(PageBreak())

    for idx, cat in enumerate(categories):
        cat_df = raw_df[raw_df["इवेंट केटेगरी"] == cat]
        if cat_df.empty:
            continue

        cat_agg = cat_df.groupby("जिला", as_index=False).agg({
            "कुल की गई प्रविष्टि": "sum",
            "कुल प्रतिभागियों की संख्या": "sum",
            "कुल फोटो की संख्या": "sum",
            "कुल वीडियो की संख्या": "sum"
        })
        for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
            cat_agg[c] = pd.to_numeric(cat_agg[c], errors='coerce').fillna(0).astype(int)

        cat_agg["कुल photo"] = cat_agg["कुल फोटो की संख्या"]
        cat_agg["कुल video"] = cat_agg["कुल वीडियो की संख्या"]

        cat_agg = cat_agg.sort_values(by="कुल की गई प्रविष्टि", ascending=False).reset_index(drop=True)
        cat_agg["इवेंट रैंक"] = range(1, len(cat_agg) + 1)

        cat_tot_entries = int(cat_agg["कुल की गई प्रविष्टि"].sum())
        cat_tot_part = int(cat_agg["कुल प्रतिभागियों की संख्या"].sum())
        cat_tot_photos = int(cat_agg["कुल फोटो की संख्या"].sum())
        cat_tot_videos = int(cat_agg["कुल वीडियो की संख्या"].sum())
        top_d = cat_agg.iloc[0]["जिला"] if not cat_agg.empty else "N/A"
        top_v = cat_agg.iloc[0]["कुल की गई प्रविष्टि"] if not cat_agg.empty else 0

        elements.append(Paragraph("VBSJA - इवेंटवार जिला रैंकिंग रिपोर्ट", title_style))
        elements.append(Paragraph(f"रिपोर्ट समय: {curr_time} | कुल इवेंट श्रेणियां: {len(categories)}", sub_style))
        elements.append(Spacer(1, 6))

        event_num_label = f"📁 इवेंट श्रेणी #{idx + 1}: {cat}"
        elements.append(Paragraph(f"<b>{event_num_label}</b>", cat_banner_style))
        elements.append(Spacer(1, 4))

        kpi_text = (
            f"<b>कुल प्रविष्टियां:</b> {cat_tot_entries:,} &nbsp;|&nbsp; "
            f"<b>कुल प्रतिभागी:</b> {cat_tot_part:,} &nbsp;|&nbsp; "
            f"<b>कुल फोटो:</b> {cat_tot_photos:,} &nbsp;|&nbsp; "
            f"<b>कुल वीडियो:</b> {cat_tot_videos:,} &nbsp;|&nbsp; "
            f"<b>🥇 शीर्ष जिला:</b> {top_d} ({top_v:,})"
        )
        kpi_table = Table([[Paragraph(kpi_text, kpi_bar_style)]], colWidths=[545])
        kpi_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F1F5F9')),
            ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#CBD5E1')),
            ('PADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(kpi_table)
        elements.append(Spacer(1, 4))

        if selected_district and selected_district != all_dist_option:
            dist_matches = cat_agg[cat_agg["जिला"] == selected_district]
            if not dist_matches.empty:
                d_row = dist_matches.iloc[0]
                d_r = rank_map.get(selected_district, {})
                d_ent_r = d_r.get("orig_ent_rank", "-")
                d_part_r = d_r.get("orig_part_rank", "-")
                d_photo_r = d_r.get("orig_photo_rank", "-")
                d_video_r = d_r.get("orig_video_rank", "-")

                dist_info_text = (
                    f"🎯 <b>चयनित जिला ({selected_district}):</b> "
                    f"प्रविष्टि: <b>{d_row['कुल की गई प्रविष्टि']:,} ({d_ent_r})</b> (इवेंट रैंक: <b>#{d_row['इवेंट रैंक']}</b>/{len(cat_agg)}) | "
                    f"प्रतिभागी: {d_row['कुल प्रतिभागियों की संख्या']:,} ({d_part_r}) | "
                    f"फोटो: {d_row['कुल फोटो की संख्या']:,} ({d_photo_r}) | "
                    f"वीडियो: {d_row['कुल वीडियो की संख्या']:,} ({d_video_r})"
                )
                dist_box = Table([[Paragraph(dist_info_text, kpi_bar_style)]], colWidths=[545])
                dist_box.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FEF3C7')),
                    ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#F59E0B')),
                    ('PADDING', (0, 0), (-1, -1), 4),
                ]))
                elements.append(dist_box)
                elements.append(Spacer(1, 4))

        elements.append(Paragraph("💡 नोट: प्रथम कॉलम 'रैंक' इस इवेंट की प्रविष्टि रैंकिंग है। प्रत्येक संख्या के साथ कोष्ठक ( ) में जिले की मूल राज्य रैंकिंग (Overall State Rank) दर्शायी गई है।", rank_note_style))
        elements.append(Spacer(1, 4))

        if is_separate:
            headers = [
                Paragraph("रैंक", table_header_style),
                Paragraph("District", table_header_style),
                Paragraph("कुल की गई प्रविष्टि", table_header_style),
                Paragraph("प्रविष्टि मूल रैंक", table_header_style),
                Paragraph("कुल प्रतिभागी संख्या", table_header_style),
                Paragraph("प्रतिभागी मूल रैंक", table_header_style),
                Paragraph("कुल फोटो की संख्या", table_header_style),
                Paragraph("फोटो मूल रैंक", table_header_style),
                Paragraph("कुल वीडियो की संख्या", table_header_style),
                Paragraph("वीडियो मूल रैंक", table_header_style)
            ]
            col_widths = [35, 90, 60, 45, 65, 45, 55, 45, 55, 45]
        else:
            headers = [
                Paragraph("रैंक", table_header_style),
                Paragraph("District", table_header_style),
                Paragraph("कुल की गई प्रविष्टि", table_header_style),
                Paragraph("कुल प्रतिभागियों की संख्या", table_header_style),
                Paragraph("कुल फोटो की संख्या", table_header_style),
                Paragraph("कुल वीडियो की संख्या", table_header_style)
            ]
            col_widths = [40, 110, 100, 115, 90, 90]

        table_rows = [headers]

        for _, row in cat_agg.iterrows():
            d_name = row["जिला"]
            d_ranks = rank_map.get(d_name, {})
            ent_r = d_ranks.get("orig_ent_rank", "-")
            part_r = d_ranks.get("orig_part_rank", "-")
            photo_r = d_ranks.get("orig_photo_rank", "-")
            video_r = d_ranks.get("orig_video_rank", "-")

            is_spotlight = (selected_district and selected_district != all_dist_option and d_name == selected_district)
            st_style = table_cell_bold if is_spotlight else table_cell_style

            if is_separate:
                table_rows.append([
                    Paragraph(f"#{row['इवेंट रैंक']}", st_style),
                    Paragraph(str(d_name), st_style),
                    Paragraph(f"{row['कुल की गई प्रविष्टि']:,}", st_style),
                    Paragraph(str(ent_r), st_style),
                    Paragraph(f"{row['कुल प्रतिभागियों की संख्या']:,}", st_style),
                    Paragraph(str(part_r), st_style),
                    Paragraph(f"{row['कुल फोटो की संख्या']:,}", st_style),
                    Paragraph(str(photo_r), st_style),
                    Paragraph(f"{row['कुल वीडियो की संख्या']:,}", st_style),
                    Paragraph(str(video_r), st_style)
                ])
            else:
                table_rows.append([
                    Paragraph(f"#{row['इवेंट रैंक']}", st_style),
                    Paragraph(str(d_name), st_style),
                    Paragraph(f"{row['कुल की गई प्रविष्टि']:,} ({ent_r})", st_style),
                    Paragraph(f"{row['कुल प्रतिभागियों की संख्या']:,} ({part_r})", st_style),
                    Paragraph(f"{row['कुल फोटो की संख्या']:,} ({photo_r})", st_style),
                    Paragraph(f"{row['कुल वीडियो की संख्या']:,} ({video_r})", st_style)
                ])

        event_table = Table(table_rows, colWidths=col_widths, repeatRows=1)
        t_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), font),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
        ]

        if selected_district and selected_district != all_dist_option:
            for r_idx, r_data in enumerate(cat_agg.iterrows(), 1):
                if r_data[1]["जिला"] == selected_district:
                    t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor('#FEF9C3')))

        event_table.setStyle(TableStyle(t_style))
        elements.append(event_table)

        if idx < len(categories) - 1:
            elements.append(PageBreak())

    doc.build(elements)
    return buf.getvalue()

def sanitize_hindi_filename(name: str) -> str:
    r"""
    Sanitizes filename for Windows/filesystem while preserving all Hindi/Devanagari
    letters, matras, English letters, numbers, spaces, and brackets.
    Only strips characters that are illegal in Windows filenames: < > : " / \ | ? *
    """
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', str(name))
    clean = clean.replace('–', '-').replace('—', '-')
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean

@st.cache_data(show_spinner=False, max_entries=10)
def create_event_wise_zip(
    raw_df: pd.DataFrame, 
    selected_district: str = None, 
    categories: list = None, 
    all_dist_option: str = "🌟 सभी जिले (All Districts)",
    table_format: str = "merged",
    rank_map: dict = None
) -> bytes:
    """
    Generates a ZIP archive containing:
    1. 00_समग्र_जिलावार_रैंकिंग_रिपोर्ट.pdf (All 10 Categories Combined Overall Total - e.g. Bhilwara: 47,528)
    2. Separate individual PDF files for each Event Category (1 page each) with proper Hindi filenames.
    Accelerated with multi-threaded rendering and Streamlit caching.
    """
    if not categories:
        categories = sorted([c for c in raw_df["इवेंट केटेगरी"].dropna().unique() if str(c).strip()])
        
    # 1. Generate overall summary PDF
    summary_pdf = create_event_wise_pdf(
        raw_df=raw_df,
        selected_district=selected_district,
        categories=categories,
        all_dist_option=all_dist_option,
        table_format=table_format,
        rank_map=rank_map,
        include_summary=True,
        only_summary=True
    )

    # 2. Worker for generating individual event PDFs in parallel
    def _gen_single_cat(item):
        idx, cat = item
        single_cat_pdf = create_event_wise_pdf(
            raw_df=raw_df,
            selected_district=selected_district,
            categories=[cat],
            all_dist_option=all_dist_option,
            table_format=table_format,
            rank_map=rank_map,
            include_summary=False,
            only_summary=False
        )
        clean_cat_name = sanitize_hindi_filename(cat)
        if not clean_cat_name:
            clean_cat_name = f"Event_{idx}"
        file_name = f"{idx:02d}_{clean_cat_name}.pdf"
        return file_name, single_cat_pdf

    items = list(enumerate(categories, 1))
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        cat_files = list(executor.map(_gen_single_cat, items))

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("00_समग्र_जिलावार_रैंकिंग_रिपोर्ट.pdf", summary_pdf)
        for file_name, single_pdf in cat_files:
            zf.writestr(file_name, single_pdf)
            
    return zip_buf.getvalue()

# ==========================================
# 3. PRINTABLE HTML VIEW (PORTRAIT)
# ==========================================
def generate_printable_html(df: pd.DataFrame, title: str = "तालिका प्रिंट प्रिव्यू", max_rows: int = 500, active_filters_text: str = "", orientation: str = "portrait") -> str:
    curr_time = datetime.now().strftime("%d-%m-%Y %I:%M %p")
    df_to_render = df.head(max_rows)

    
    filter_banner_html = ""
    if active_filters_text:
        filter_banner_html = f"""
        <div style="background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 14px; margin-bottom: 14px; font-size: 12px; color: #334155; line-height: 1.5;">
            <strong>📌 लागू फ़िल्टर (Active Filters):</strong> {active_filters_text}
        </div>
        """

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        * {{
            box-sizing: border-box;
            font-family: 'Nirmala UI', 'Mukta', 'Segoe UI', Arial, sans-serif;
        }}
        body {{
            background: #f8fafc;
            color: #0f172a;
            padding: 20px;
            margin: 0;
        }}
        .print-toolbar {{
            position: sticky;
            top: 0;
            background: #1e3a8a;
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
        }}
        .print-btn {{
            background: #f59e0b;
            color: #111827;
            font-weight: 700;
            border: none;
            padding: 8px 20px;
            border-radius: 6px;
            font-size: 15px;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
        }}
        .print-btn:hover {{
            background: #d97706;
            transform: scale(1.03);
        }}
        .report-header {{
            text-align: center;
            margin-bottom: 16px;
        }}
        .report-header h1 {{
            color: #1e3a8a;
            margin: 0 0 6px 0;
            font-size: 22px;
        }}
        .report-header p {{
            color: #64748b;
            margin: 0;
            font-size: 13px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            border-radius: 6px;
            overflow: hidden;
            font-size: 12px;
        }}
        th {{
            background-color: #1e3a8a;
            color: white;
            text-align: center;
            padding: 9px 6px;
            font-weight: 600;
            border: 1px solid #1e3a8a;
        }}
        td {{
            padding: 7px 6px;
            border: 1px solid #cbd5e1;
            text-align: center;
        }}
        tr:nth-child(even) {{
            background-color: #f8fafc;
        }}
        @media print {{
            .print-toolbar {{
                display: none !important;
            }}
            body {{
                background: white;
                padding: 0;
            }}
            table {{
                box-shadow: none;
                font-size: 10.5px;
            }}
            th {{
                background-color: #1e3a8a !important;
                color: white !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
                padding: 6px 4px;
            }}
            td {{
                padding: 5px 4px;
            }}
            tr:nth-child(even) {{
                background-color: #f8fafc !important;
                -webkit-print-color-adjust: exact;
                print-color-adjust: exact;
            }}
            @page {{
                margin: 0.8cm;
                size: {orientation};
            }}
        }}
    </style>
</head>
<body>
    <div class="print-toolbar">
        <div>
            <strong>🖨️ प्रिंट प्रिव्यू ({orientation.capitalize()})</strong> &nbsp;|&nbsp; <span>{title}</span>
        </div>
        <button class="print-btn" onclick="window.print()">
            🖨️ अभी प्रिंट करें / Save as PDF
        </button>
    </div>

    <div class="report-header">
        <h1>{title}</h1>
        <p>दिनांक व समय: {curr_time} | कुल रिकॉर्ड: {len(df_to_render):,}</p>
    </div>

    {filter_banner_html}

    <table>
        <thead>
            <tr>
"""
    for col in df_to_render.columns:
        html += f"                <th>{col}</th>\n"
    html += "            </tr>\n        </thead>\n        <tbody>\n"

    for _, row in df_to_render.iterrows():
        html += "            <tr>\n"
        for col in df_to_render.columns:
            val = row[col]
            if isinstance(val, (int, float)):
                formatted = f"{int(val):,}" if isinstance(val, int) or val.is_integer() else f"{val:.2f}"
            else:
                formatted = str(val) if pd.notna(val) else ""
            html += f"                <td>{formatted}</td>\n"
        html += "            </tr>\n"

    html += """        </tbody>
    </table>
</body>
</html>"""
    return html

# ==========================================
# 4. EVENT-WISE, BLOCK-WISE REPORT WITH 3-COLOR CONDITIONAL FORMATTING
# ==========================================
PREFERRED_BLOCK_CAT_ORDER = [
    'सेवा की कहानी',
    'आशीर्वाद का दीया',
    'सेवा सेतु अभियान (ग्रामीण क्षेत्र)',
    'नमो सेवा–अनकही कहानियाँ (ग्रामीण क्षेत्र)',
    'सेवा के रंग, राष्ट्र के संग',
    'नमो सेवा–अनकही कहानियाँ (शहरी क्षेत्र)',
    'मेरे सपनों का भारत–चित्रसेवा',
    'रक्तदान शिविर',
    'आंगन मिलन सेवा',
    'Gramin seva shivir ILR Wise'
]

def get_rank_3color(ratio: float) -> str:
    """
    Excel 3-color scale interpolation:
    0.0 (best rank, lowest number) -> Soft Green (#63be7b)
    0.5 (middle rank)              -> Soft Yellow (#ffeb84)
    1.0 (worst rank, highest number)-> Soft Red (#f8696b)
    """
    ratio = max(0.0, min(1.0, float(ratio)))
    if ratio <= 0.5:
        r2 = ratio / 0.5
        r = int(99 + (255 - 99) * r2)
        g = int(190 + (235 - 190) * r2)
        b = int(123 + (132 - 123) * r2)
    else:
        r2 = (ratio - 0.5) / 0.5
        r = int(255 + (248 - 255) * r2)
        g = int(235 + (105 - 235) * r2)
        b = int(132 + (107 - 132) * r2)
    return f"#{r:02x}{g:02x}{b:02x}"

@st.cache_data(show_spinner=False, max_entries=50)
def compute_block_wise_tables(raw_df: pd.DataFrame, target_district: str, categories: list = None):
    """
    Computes Block-level aggregated tables with statewide ranks across ALL blocks in Rajasthan.
    Returns list of dicts:
    [{"title": str, "df": pd.DataFrame, "is_overall": bool}, ...]
    Each df has columns:
    ['क्र.सं.', 'ब्लॉक', 'प्रविष्टि', 'राज्य-रैंक-प्रविष्टि', 'प्रतिभागी', 'राज्य-रैंक-प्रतिभागी', 'फोटो', 'राज्य-रैंक-फोटो', 'वीडियो', 'राज्य-रैंक-वीडियो']
    """
    df_valid = raw_df.dropna(subset=['ब्लॉक']).copy()
    df_valid['ब्लॉक'] = df_valid['ब्लॉक'].astype(str).str.strip()
    df_valid = df_valid[df_valid['ब्लॉक'] != '']

    # Standardize photo and video columns
    if "कुल फोटो की संख्या" not in df_valid.columns and "कुल photo" in df_valid.columns:
        df_valid["कुल फोटो की संख्या"] = df_valid["कुल photo"]
    if "कुल वीडियो की संख्या" not in df_valid.columns and "कुल video" in df_valid.columns:
        df_valid["कुल वीडियो की संख्या"] = df_valid["कुल video"]

    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        if c in df_valid.columns:
            df_valid[c] = pd.to_numeric(df_valid[c], errors='coerce').fillna(0).astype(int)
        else:
            df_valid[c] = 0

    all_dist_blocks = df_valid[['जिला', 'ब्लॉक']].drop_duplicates().copy()

    # Determine categories
    if not categories:
        cats = [c for c in PREFERRED_BLOCK_CAT_ORDER if c in df_valid['इवेंट केटेगरी'].unique() and c != 'Gramin seva shivir ILR Wise']
        for c in sorted(df_valid['इवेंट केटेगरी'].dropna().unique()):
            if str(c).strip() and c not in cats and c != 'Gramin seva shivir ILR Wise':
                cats.append(c)
    else:
        cats = [c for c in PREFERRED_BLOCK_CAT_ORDER if c in categories]
        for c in categories:
            if c not in cats:
                cats.append(c)

    def _calc_table(sub_df, title, is_overall=False):
        agg = sub_df.groupby(['जिला', 'ब्लॉक'], as_index=False).agg({
            'कुल की गई प्रविष्टि': 'sum',
            'कुल प्रतिभागियों की संख्या': 'sum',
            'कुल फोटो की संख्या': 'sum',
            'कुल वीडियो की संख्या': 'sum'
        })
        merged = pd.merge(all_dist_blocks, agg, on=['जिला', 'ब्लॉक'], how='left').fillna(0)
        for num_c in ['कुल की गई प्रविष्टि', 'कुल प्रतिभागियों की संख्या', 'कुल फोटो की संख्या', 'कुल वीडियो की संख्या']:
            merged[num_c] = merged[num_c].astype(int)

        merged['राज्य-रैंक-प्रविष्टि'] = merged['कुल की गई प्रविष्टि'].rank(ascending=False, method='min').astype(int)
        merged['राज्य-रैंक-प्रतिभागी'] = merged['कुल प्रतिभागियों की संख्या'].rank(ascending=False, method='min').astype(int)
        merged['राज्य-रैंक-फोटो'] = merged['कुल फोटो की संख्या'].rank(ascending=False, method='min').astype(int)
        merged['राज्य-रैंक-वीडियो'] = merged['कुल वीडियो की संख्या'].rank(ascending=False, method='min').astype(int)

        res = merged[merged['जिला'] == target_district].sort_values('ब्लॉक').reset_index(drop=True)
        res.insert(0, 'क्र.सं.', range(1, len(res) + 1))
        
        cols_ordered = [
            'क्र.सं.',
            'ब्लॉक',
            'कुल की गई प्रविष्टि',
            'राज्य-रैंक-प्रविष्टि',
            'कुल प्रतिभागियों की संख्या',
            'राज्य-रैंक-प्रतिभागी',
            'कुल फोटो की संख्या',
            'राज्य-रैंक-फोटो',
            'कुल वीडियो की संख्या',
            'राज्य-रैंक-वीडियो'
        ]
        res = res[cols_ordered].rename(columns={
            'कुल की गई प्रविष्टि': 'प्रविष्टि',
            'कुल प्रतिभागियों की संख्या': 'प्रतिभागी',
            'कुल फोटो की संख्या': 'फोटो',
            'कुल वीडियो की संख्या': 'वीडियो'
        })
        return {"title": title, "df": res, "is_overall": is_overall}

    tables = []
    tables.append(_calc_table(df_valid, f"समग्र (सभी इवेंट केटेगरी जोड़कर) — {target_district}", is_overall=True))
    for cat in cats:
        cat_sub = df_valid[df_valid['इवेंट केटेगरी'] == cat]
        if not cat_sub.empty:
            tables.append(_calc_table(cat_sub, cat, is_overall=False))

    return tables

def generate_block_wise_pdf_html(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> str:
    curr_time = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    tables = compute_block_wise_tables(raw_df, target_district, categories)

    def _render_tbl_html(title, tbl_df):
        rank_cols = ['राज्य-रैंक-प्रविष्टि', 'राज्य-रैंक-प्रतिभागी', 'राज्य-रैंक-फोटो', 'राज्य-रैंक-वीडियो']
        col_min_max = {}
        for rc in rank_cols:
            col_min_max[rc] = (tbl_df[rc].min(), tbl_df[rc].max())

        tbody_rows = []
        for _, row in tbl_df.iterrows():
            tds = []
            tds.append(f'<td style="width: 35px;">{row["क्र.सं."]}</td>')
            tds.append(f'<td style="width: 80px; text-align: left; padding-left: 6px;">{row["ब्लॉक"]}</td>')
            tds.append(f'<td style="width: 55px;">{row["प्रविष्टि"]:,}</td>')

            # rank ent
            min_v, max_v = col_min_max['राज्य-रैंक-प्रविष्टि']
            ratio = (row['राज्य-रैंक-प्रविष्टि'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
            bg = get_rank_3color(ratio)
            tds.append(f'<td style="width: 70px; background-color: {bg}; color: #000; font-weight: 500;">{row["राज्य-रैंक-प्रविष्टि"]}</td>')

            tds.append(f'<td style="width: 60px;">{row["प्रतिभागी"]:,}</td>')

            # rank part
            min_v, max_v = col_min_max['राज्य-रैंक-प्रतिभागी']
            ratio = (row['राज्य-रैंक-प्रतिभागी'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
            bg = get_rank_3color(ratio)
            tds.append(f'<td style="width: 70px; background-color: {bg}; color: #000; font-weight: 500;">{row["राज्य-रैंक-प्रतिभागी"]}</td>')

            tds.append(f'<td style="width: 50px;">{row["फोटो"]:,}</td>')

            # rank photo
            min_v, max_v = col_min_max['राज्य-रैंक-फोटो']
            ratio = (row['राज्य-रैंक-फोटो'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
            bg = get_rank_3color(ratio)
            tds.append(f'<td style="width: 65px; background-color: {bg}; color: #000; font-weight: 500;">{row["राज्य-रैंक-फोटो"]}</td>')

            tds.append(f'<td style="width: 50px;">{row["वीडियो"]:,}</td>')

            # rank video
            min_v, max_v = col_min_max['राज्य-रैंक-वीडियो']
            ratio = (row['राज्य-रैंक-वीडियो'] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
            bg = get_rank_3color(ratio)
            tds.append(f'<td style="width: 65px; background-color: {bg}; color: #000; font-weight: 500;">{row["राज्य-रैंक-वीडियो"]}</td>')

            tbody_rows.append('<tr>' + ''.join(tds) + '</tr>')

        return f'''
        <div class="table-box">
            <div class="table-title">{title}</div>
            <table>
                <thead>
                    <tr>
                        <th style="width: 35px;">क्र.सं.</th>
                        <th style="width: 80px;">ब्लॉक</th>
                        <th style="width: 55px;">प्रविष्टि</th>
                        <th style="width: 70px;">राज्य-रैंक-<br>प्रविष्टि</th>
                        <th style="width: 60px;">प्रतिभागी</th>
                        <th style="width: 70px;">राज्य-रैंक-<br>प्रतिभागी</th>
                        <th style="width: 50px;">फोटो</th>
                        <th style="width: 65px;">राज्य-रैंक-<br>फोटो</th>
                        <th style="width: 50px;">वीडियो</th>
                        <th style="width: 65px;">राज्य-रैंक-<br>वीडियो</th>
                    </tr>
                </thead>
                <tbody>
                    {''.join(tbody_rows)}
                </tbody>
            </table>
        </div>
        '''

    pages_html = []
    chunk_size = 4
    for p_idx in range(0, len(tables), chunk_size):
        chunk = tables[p_idx:p_idx+chunk_size]
        tbl_htmls = [_render_tbl_html(t["title"], t["df"]) for t in chunk]
        p_html = f'''
        <div class="report-page">
            <div class="page-header">
                <h1>जिला {target_district} — ब्लॉक स्तरीय रैंकिंग (राज्य के सभी जिलों के ब्लॉक में से)</h1>
                <p>दिनांक {curr_time}</p>
            </div>
            {''.join(tbl_htmls)}
        </div>
        '''
        pages_html.append(p_html)

    return f'''<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>जिला {target_district} - ब्लॉक स्तरीय रैंकिंग</title>
    <style>
        @page {{
            size: A4 portrait;
            margin: 8mm 12mm 8mm 12mm;
        }}
        * {{
            box-sizing: border-box;
            font-family: 'Nirmala UI', 'Mukta', 'Segoe UI', Arial, sans-serif;
        }}
        body {{
            margin: 0;
            padding: 0;
            color: #000;
            background: white;
            font-size: 10px;
        }}
        .report-page {{
            break-after: page;
            page-break-after: always;
            padding-bottom: 2px;
        }}
        .report-page:last-child {{
            break-after: auto;
            page-break-after: auto;
        }}
        .page-header {{
            text-align: center;
            margin-bottom: 8px;
        }}
        .page-header h1 {{
            color: #1f4d77;
            font-size: 15px;
            font-weight: 700;
            margin: 0 0 3px 0;
        }}
        .page-header p {{
            color: #1f4d77;
            font-size: 12px;
            font-weight: 700;
            margin: 0;
        }}
        .table-box {{
            margin-bottom: 8px;
            break-inside: avoid;
            page-break-inside: avoid;
        }}
        .table-title {{
            font-size: 11px;
            font-weight: 700;
            color: #1f4d77;
            margin-bottom: 2px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 9.5px;
        }}
        th {{
            background-color: #1f4d77;
            color: white;
            border: 1px solid #1f4d77;
            padding: 3px 1px;
            text-align: center;
            font-weight: 600;
            line-height: 1.15;
            font-size: 9px;
        }}
        td {{
            border: 1px solid #777777;
            padding: 2.5px 1px;
            text-align: center;
            line-height: 1.15;
            font-size: 9.5px;
        }}
    </style>
</head>
<body>
    {''.join(pages_html)}
</body>
</html>'''

@st.cache_data(show_spinner=False, max_entries=30)
def create_block_wise_pdf(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> bytes:
    """
    Generates an exact 3-page A4 Portrait PDF for Block-level Event Rankings with 3-color conditional formatting.
    Uses Chromium headless for exact Hindi rendering, with ReportLab fallback.
    """
    try:
        html = generate_block_wise_pdf_html(raw_df, target_district, categories)
        pdf_bytes = render_html_to_pdf(html)
        if pdf_bytes:
            return pdf_bytes
    except Exception:
        pass

    return _create_block_wise_pdf_reportlab(raw_df, target_district, categories)

def _create_block_wise_pdf_reportlab(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> bytes:
    font = register_hindi_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    elements = []
    styles = getSampleStyleSheet()

    header_style = ParagraphStyle(
        'BlkTitle',
        parent=styles['Heading1'],
        fontName=font,
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#1F4D77'),
        alignment=1
    )
    sub_style = ParagraphStyle(
        'BlkSub',
        parent=styles['Normal'],
        fontName=font,
        fontSize=10,
        leading=13,
        textColor=colors.HexColor('#1F4D77'),
        alignment=1
    )
    tbl_title_style = ParagraphStyle(
        'BlkTblTitle',
        parent=styles['Heading2'],
        fontName=font,
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor('#1F4D77'),
        alignment=0
    )
    th_style = ParagraphStyle(
        'BlkTh',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7,
        leading=8.5,
        textColor=colors.white,
        alignment=1
    )
    td_style = ParagraphStyle(
        'BlkTd',
        parent=styles['Normal'],
        fontName=font,
        fontSize=7,
        leading=8.5,
        textColor=colors.black,
        alignment=1
    )

    curr_time = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    tables = compute_block_wise_tables(raw_df, target_district, categories)

    for idx, t_info in enumerate(tables):
        # Header on every 4th table
        if idx % 4 == 0:
            if idx > 0:
                elements.append(PageBreak())
            elements.append(Paragraph(f"जिला {target_district} — ब्लॉक स्तरीय रैंकिंग (राज्य के सभी जिलों के ब्लॉक में से)", header_style))
            elements.append(Paragraph(f"दिनांक {curr_time}", sub_style))
            elements.append(Spacer(1, 6))

        elements.append(Paragraph(f"<b>{t_info['title']}</b>", tbl_title_style))
        elements.append(Spacer(1, 2))

        df_t = t_info['df']
        rank_cols = ['राज्य-रैंक-प्रविष्टि', 'राज्य-रैंक-प्रतिभागी', 'राज्य-रैंक-फोटो', 'राज्य-रैंक-वीडियो']
        col_min_max = {rc: (df_t[rc].min(), df_t[rc].max()) for rc in rank_cols}

        headers = [
            Paragraph("क्र.सं.", th_style),
            Paragraph("ब्लॉक", th_style),
            Paragraph("प्रविष्टि", th_style),
            Paragraph("राज्य-रैंक-<br/>प्रविष्टि", th_style),
            Paragraph("प्रतिभागी", th_style),
            Paragraph("राज्य-रैंक-<br/>प्रतिभागी", th_style),
            Paragraph("फोटो", th_style),
            Paragraph("राज्य-रैंक-<br/>फोटो", th_style),
            Paragraph("वीडियो", th_style),
            Paragraph("राज्य-रैंक-<br/>वीडियो", th_style)
        ]
        t_data = [headers]
        t_styles = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4D77')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), font),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#777777')),
        ]

        for r_idx, (_, row) in enumerate(df_t.iterrows(), 1):
            row_cells = [
                Paragraph(str(row['क्र.सं.']), td_style),
                Paragraph(str(row['ब्लॉक']), td_style),
                Paragraph(f"{row['प्रविष्टि']:,}", td_style),
                Paragraph(str(row['राज्य-रैंक-प्रविष्टि']), td_style),
                Paragraph(f"{row['प्रतिभागी']:,}", td_style),
                Paragraph(str(row['राज्य-रैंक-प्रतिभागी']), td_style),
                Paragraph(f"{row['फोटो']:,}", td_style),
                Paragraph(str(row['फोटो']), td_style),
                Paragraph(f"{row['वीडियो']:,}", td_style),
                Paragraph(str(row['राज्य-रैंक-वीडियो']), td_style)
            ]
            t_data.append(row_cells)

            # Conditional colors on rank columns: col index 3, 5, 7, 9
            for col_idx, rc in [(3, 'राज्य-रैंक-प्रविष्टि'), (5, 'राज्य-रैंक-प्रतिभागी'), (7, 'राज्य-रैंक-फोटो'), (9, 'राज्य-रैंक-वीडियो')]:
                min_v, max_v = col_min_max[rc]
                ratio = (row[rc] - min_v) / (max_v - min_v) if max_v > min_v else 0.0
                bg = get_rank_3color(ratio)
                t_styles.append(('BACKGROUND', (col_idx, r_idx), (col_idx, r_idx), colors.HexColor(bg)))

        col_w = [28, 70, 52, 60, 55, 60, 48, 58, 48, 58]
        tbl_obj = Table(t_data, colWidths=col_w)
        tbl_obj.setStyle(TableStyle(t_styles))
        elements.append(tbl_obj)
        elements.append(Spacer(1, 6))

    doc.build(elements)
    return buf.getvalue()

def generate_block_wise_printable_html(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> str:
    """
    Generates a web-viewable printable HTML page with print toolbar.
    """
    body_html = generate_block_wise_pdf_html(raw_df, target_district, categories)
    # Inject print toolbar at top of body
    toolbar = f'''
    <div style="position: sticky; top: 0; background: #1f4d77; color: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; z-index: 1000; box-shadow: 0 3px 8px rgba(0,0,0,0.2); border-radius: 6px; margin: 10px 12px 14px 12px;">
        <span style="font-weight: 700; font-size: 15px;">🖨️ ब्लॉक स्तरीय रैंकिंग रिपोर्ट प्रिव्यू (जिला: {target_district})</span>
        <button onclick="window.print()" style="background: #f59e0b; color: #111827; font-weight: 700; border: none; padding: 7px 18px; border-radius: 5px; cursor: pointer; font-size: 14px;">
            🖨️ अभी प्रिंट करें / Save as PDF
        </button>
    </div>
    '''
    return body_html.replace('<body>', f'<body>\n{toolbar}')

@st.cache_data(show_spinner=False, max_entries=30)
def create_block_wise_excel(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> bytes:
    """
    Generates an Excel workbook with styled tables and 3-color conditional formatting on rank columns.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    tables = compute_block_wise_tables(raw_df, target_district, categories)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ब्लॉक_रैंकिंग"

    # Define styles
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4D77", end_color="1F4D77", fill_type="solid")
    title_font = Font(name="Calibri", size=11, bold=True, color="1F4D77")
    cell_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='777777'),
        right=Side(style='thin', color='777777'),
        top=Side(style='thin', color='777777'),
        bottom=Side(style='thin', color='777777')
    )

    curr_row = 1
    # Report Main Title
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=10)
    top_cell = ws.cell(row=curr_row, column=1, value=f"जिला {target_district} — ब्लॉक स्तरीय रैंकिंग (राज्य के सभी जिलों के ब्लॉक में से)")
    top_cell.font = Font(name="Calibri", size=13, bold=True, color="1F4D77")
    top_cell.alignment = center_align
    curr_row += 1

    date_cell = ws.cell(row=curr_row, column=1, value=f"दिनांक {datetime.now().strftime('%d/%m/%Y %I:%M %p')}")
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=10)
    date_cell.font = Font(name="Calibri", size=10, bold=True, color="1F4D77")
    date_cell.alignment = center_align
    curr_row += 2

    for t_info in tables:
        # Table title
        ws.cell(row=curr_row, column=1, value=t_info['title']).font = title_font
        curr_row += 1

        df_t = t_info['df']
        rank_cols = ['राज्य-रैंक-प्रविष्टि', 'राज्य-रैंक-प्रतिभागी', 'राज्य-रैंक-फोटो', 'राज्य-रैंक-वीडियो']
        col_min_max = {rc: (df_t[rc].min(), df_t[rc].max()) for rc in rank_cols}

        # Header row
        for c_idx, col_name in enumerate(df_t.columns, 1):
            cell = ws.cell(row=curr_row, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border
        curr_row += 1

        # Data rows
        for _, row in df_t.iterrows():
            for c_idx, col_name in enumerate(df_t.columns, 1):
                val = row[col_name]
                cell = ws.cell(row=curr_row, column=c_idx, value=val)
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = left_align if col_name == 'ब्लॉक' else center_align

                # Apply conditional formatting on rank columns
                if col_name in rank_cols:
                    min_v, max_v = col_min_max[col_name]
                    ratio = (val - min_v) / (max_v - min_v) if max_v > min_v else 0.0
                    hex_color = get_rank_3color(ratio).replace('#', '')
                    cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
            curr_row += 1

        curr_row += 1 # Space between tables

    # Auto fit column widths
    col_widths = {1: 8, 2: 14, 3: 12, 4: 18, 5: 12, 6: 18, 7: 10, 8: 16, 9: 10, 10: 16}
    for c_idx, w in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==========================================
# 5. DISTRICT EVENT-WISE SUMMARY REPORT (MATCHING USER SCREENSHOT)
# ==========================================
@st.cache_data(show_spinner=False, max_entries=60)
def compute_district_event_summary_table(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> pd.DataFrame:
    """
    Computes a summary table for a specific district across all event categories:
    - Each event category row displays: प्रविष्टि, रैंक-प्रविष्टि, प्रतिभागी, रैंक-प्रतिभागी, फोटो, रैंक-फोटो, वीडियो, रैंक-वीडियो
    - Statewide rank for each metric is calculated among all 41 districts in Rajasthan for that specific event.
    - Bottom row 'समग्र (सभी इवेंट जोड़कर)' shows the district's overall total and statewide overall ranks.
    """
    df_valid = raw_df.dropna(subset=['जिला']).copy()
    df_valid['जिला'] = df_valid['जिला'].astype(str).str.strip()
    df_valid = df_valid[df_valid['जिला'] != '']

    # Standardize photo and video columns
    if "कुल फोटो की संख्या" not in df_valid.columns and "कुल photo" in df_valid.columns:
        df_valid["कुल फोटो की संख्या"] = df_valid["कुल photo"]
    if "कुल वीडियो की संख्या" not in df_valid.columns and "कुल video" in df_valid.columns:
        df_valid["कुल वीडियो की संख्या"] = df_valid["कुल video"]

    for c in ["कुल की गई प्रविष्टि", "कुल प्रतिभागियों की संख्या", "कुल फोटो की संख्या", "कुल वीडियो की संख्या"]:
        if c in df_valid.columns:
            df_valid[c] = pd.to_numeric(df_valid[c], errors='coerce').fillna(0).astype(int)
        else:
            df_valid[c] = 0

    all_districts = sorted(df_valid['जिला'].unique())
    total_districts_count = max(1, len(all_districts))

    # Category ordering matching preferred order
    if not categories:
        cats = [c for c in PREFERRED_BLOCK_CAT_ORDER if c in df_valid['इवेंट केटेगरी'].unique() and c != 'Gramin seva shivir ILR Wise']
        for c in sorted(df_valid['इवेंट केटेगरी'].dropna().unique()):
            if str(c).strip() and c not in cats and c != 'Gramin seva shivir ILR Wise':
                cats.append(c)
    else:
        cats = [c for c in PREFERRED_BLOCK_CAT_ORDER if c in categories]
        for c in categories:
            if c not in cats:
                cats.append(c)

    rows = []
    for cat in cats:
        sub_df = df_valid[df_valid['इवेंट केटेगरी'] == cat]
        agg = sub_df.groupby('जिला', as_index=False).agg({
            'कुल की गई प्रविष्टि': 'sum',
            'कुल प्रतिभागियों की संख्या': 'sum',
            'कुल फोटो की संख्या': 'sum',
            'कुल वीडियो की संख्या': 'sum'
        })
        merged = pd.merge(pd.DataFrame({'जिला': all_districts}), agg, on='जिला', how='left').fillna(0)
        for num_c in ['कुल की गई प्रविष्टि', 'कुल प्रतिभागियों की संख्या', 'कुल फोटो की संख्या', 'कुल वीडियो की संख्या']:
            merged[num_c] = merged[num_c].astype(int)

        merged['रैंक-प्रविष्टि'] = merged['कुल की गई प्रविष्टि'].rank(ascending=False, method='min').astype(int)
        merged['रैंक-प्रतिभागी'] = merged['कुल प्रतिभागियों की संख्या'].rank(ascending=False, method='min').astype(int)
        merged['रैंक-फोटो'] = merged['कुल फोटो की संख्या'].rank(ascending=False, method='min').astype(int)
        merged['रैंक-वीडियो'] = merged['कुल वीडियो की संख्या'].rank(ascending=False, method='min').astype(int)

        d_row = merged[merged['जिला'] == target_district]
        if not d_row.empty:
            r = d_row.iloc[0]
            rows.append({
                'इवेंट केटेगरी': cat,
                'प्रविष्टि': int(r['कुल की गई प्रविष्टि']),
                'रैंक-प्रविष्टि': int(r['रैंक-प्रविष्टि']),
                'प्रतिभागी': int(r['कुल प्रतिभागियों की संख्या']),
                'रैंक-प्रतिभागी': int(r['रैंक-प्रतिभागी']),
                'फोटो': int(r['कुल फोटो की संख्या']),
                'रैंक-फोटो': int(r['रैंक-फोटो']),
                'वीडियो': int(r['कुल वीडियो की संख्या']),
                'रैंक-वीडियो': int(r['रैंक-वीडियो']),
                'is_summary': False
            })
        else:
            rows.append({
                'इवेंट केटेगरी': cat,
                'प्रविष्टि': 0,
                'रैंक-प्रविष्टि': total_districts_count,
                'प्रतिभागी': 0,
                'रैंक-प्रतिभागी': total_districts_count,
                'फोटो': 0,
                'रैंक-फोटो': total_districts_count,
                'वीडियो': 0,
                'रैंक-वीडियो': total_districts_count,
                'is_summary': False
            })

    # Summary Row across all events
    sub_all = df_valid[df_valid['इवेंट केटेगरी'].isin(cats)] if cats else df_valid
    agg_all = sub_all.groupby('जिला', as_index=False).agg({
        'कुल की गई प्रविष्टि': 'sum',
        'कुल प्रतिभागियों की संख्या': 'sum',
        'कुल फोटो की संख्या': 'sum',
        'कुल वीडियो की संख्या': 'sum'
    })
    merged_all = pd.merge(pd.DataFrame({'जिला': all_districts}), agg_all, on='जिला', how='left').fillna(0)
    for num_c in ['कुल की गई प्रविष्टि', 'कुल प्रतिभागियों की संख्या', 'कुल फोटो की संख्या', 'कुल वीडियो की संख्या']:
        merged_all[num_c] = merged_all[num_c].astype(int)

    merged_all['रैंक-प्रविष्टि'] = merged_all['कुल की गई प्रविष्टि'].rank(ascending=False, method='min').astype(int)
    merged_all['रैंक-प्रतिभागी'] = merged_all['कुल प्रतिभागियों की संख्या'].rank(ascending=False, method='min').astype(int)
    merged_all['रैंक-फोटो'] = merged_all['कुल फोटो की संख्या'].rank(ascending=False, method='min').astype(int)
    merged_all['रैंक-वीडियो'] = merged_all['कुल वीडियो की संख्या'].rank(ascending=False, method='min').astype(int)

    d_all_row = merged_all[merged_all['जिला'] == target_district]
    if not d_all_row.empty:
        r_all = d_all_row.iloc[0]
        rows.append({
            'इवेंट केटेगरी': 'समग्र (सभी इवेंट जोड़कर)',
            'प्रविष्टि': int(r_all['कुल की गई प्रविष्टि']),
            'रैंक-प्रविष्टि': int(r_all['रैंक-प्रविष्टि']),
            'प्रतिभागी': int(r_all['कुल प्रतिभागियों की संख्या']),
            'रैंक-प्रतिभागी': int(r_all['रैंक-प्रतिभागी']),
            'फोटो': int(r_all['कुल फोटो की संख्या']),
            'रैंक-फोटो': int(r_all['रैंक-फोटो']),
            'वीडियो': int(r_all['कुल वीडियो की संख्या']),
            'रैंक-वीडियो': int(r_all['रैंक-वीडियो']),
            'is_summary': True
        })
    else:
        rows.append({
            'इवेंट केटेगरी': 'समग्र (सभी इवेंट जोड़कर)',
            'प्रविष्टि': 0,
            'रैंक-प्रविष्टि': total_districts_count,
            'प्रतिभागी': 0,
            'रैंक-प्रतिभागी': total_districts_count,
            'फोटो': 0,
            'रैंक-फोटो': total_districts_count,
            'वीडियो': 0,
            'रैंक-वीडियो': total_districts_count,
            'is_summary': True
        })

    return pd.DataFrame(rows)

def generate_district_event_summary_pdf_html(raw_df: pd.DataFrame, target_district: str, categories: list = None, custom_title: str = None) -> str:
    """
    Generates exact HTML matching the user screenshot:
    - Title: जिला [जिला नाम] — सभी इवेंट केटेगरी में रैंक दिनांक [दिनांक समय]
    - Navy Blue header #1F4D77
    - 3-color conditional formatting on rank cells (Green -> Yellow -> Red)
    - Golden/Yellow highlight #FEF3C7 on summary row
    """
    df_res = compute_district_event_summary_table(raw_df, target_district, categories)
    curr_time = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    title = custom_title or f"जिला {target_district} — सभी इवेंट केटेगरी में रैंक दिनांक {curr_time}"

    # Total districts in state for rank normalization (41)
    all_dist_count = max(1, len(raw_df['जिला'].dropna().unique()))

    tbody_rows = []
    for _, row in df_res.iterrows():
        is_sum = bool(row['is_summary'])
        row_bg = '#FEF3C7' if is_sum else '#FFFFFF'
        f_weight = '700' if is_sum else '500'

        tds = []
        # 1. Event category
        tds.append(f'<td style="text-align: left; padding-left: 10px; font-weight: {f_weight}; background-color: {row_bg};">{row["इवेंट केटेगरी"]}</td>')

        # 2. Entries
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {row_bg};">{row["प्रविष्टि"]:,}</td>')

        # 3. Rank entries (with 3-color)
        r_ent = int(row['रैंक-प्रविष्टि'])
        ratio_ent = (r_ent - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
        bg_ent = get_rank_3color(ratio_ent)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_ent}; color: #000;">{r_ent}</td>')

        # 4. Participants
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {row_bg};">{row["प्रतिभागी"]:,}</td>')

        # 5. Rank participants
        r_part = int(row['रैंक-प्रतिभागी'])
        ratio_part = (r_part - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
        bg_part = get_rank_3color(ratio_part)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_part}; color: #000;">{r_part}</td>')

        # 6. Photos
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {row_bg};">{row["फोटो"]:,}</td>')

        # 7. Rank photos
        r_photo = int(row['रैंक-फोटो'])
        ratio_photo = (r_photo - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
        bg_photo = get_rank_3color(ratio_photo)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_photo}; color: #000;">{r_photo}</td>')

        # 8. Videos
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {row_bg};">{row["वीडियो"]:,}</td>')

        # 9. Rank videos
        r_vid = int(row['रैंक-वीडियो'])
        ratio_vid = (r_vid - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
        bg_vid = get_rank_3color(ratio_vid)
        tds.append(f'<td style="text-align: center; font-weight: {f_weight}; background-color: {bg_vid}; color: #000;">{r_vid}</td>')

        tbody_rows.append(f'<tr>{"".join(tds)}</tr>')

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        @page {{
            size: A4 portrait;
            margin: 22mm 15mm 20mm 15mm;
        }}
        * {{
            box-sizing: border-box;
            font-family: 'Nirmala UI', 'Mukta', 'Segoe UI', Arial, sans-serif;
        }}
        body {{
            margin: 0;
            padding: 0;
            background: #FFFFFF;
            color: #111827;
        }}
        .report-wrap {{
            max-width: 780px;
            margin: 0 auto;
            padding: 10px;
        }}
        .report-title {{
            text-align: center;
            font-size: 17px;
            font-weight: 700;
            color: #1F4D77;
            margin-top: 15px;
            margin-bottom: 22px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12.5px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        th {{
            background-color: #1F4D77;
            color: #FFFFFF;
            font-weight: 700;
            padding: 8px 6px;
            text-align: center;
            border: 1px solid #777777;
            line-height: 1.3;
        }}
        th:first-child {{
            text-align: center;
        }}
        td {{
            padding: 7px 6px;
            border: 1px solid #777777;
            line-height: 1.3;
        }}
    </style>
</head>
<body>
    <div class="report-wrap">
        <div class="report-title">{title}</div>
        <table>
            <thead>
                <tr>
                    <th style="width: 210px;">इवेंट केटेगरी</th>
                    <th style="width: 55px;">प्रविष्टि</th>
                    <th style="width: 65px;">रैंक-प्रविष्टि</th>
                    <th style="width: 65px;">प्रतिभागी</th>
                    <th style="width: 65px;">रैंक-प्रतिभागी</th>
                    <th style="width: 50px;">फोटो</th>
                    <th style="width: 60px;">रैंक-फोटो</th>
                    <th style="width: 50px;">वीडियो</th>
                    <th style="width: 60px;">रैंक-वीडियो</th>
                </tr>
            </thead>
            <tbody>
                {''.join(tbody_rows)}
            </tbody>
        </table>
    </div>
</body>
</html>"""
    return html

@st.cache_data(show_spinner=False, max_entries=50)
def create_district_event_summary_pdf(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> bytes:
    """
    Renders high-quality PDF matching the user's screenshot.
    Uses headless Chromium for perfect Hindi rendering, with ReportLab fallback.
    """
    html_content = generate_district_event_summary_pdf_html(raw_df, target_district, categories)
    pdf_bytes = render_html_to_pdf(html_content)
    if pdf_bytes:
        return pdf_bytes

    # ReportLab Fallback
    hindi_font = register_hindi_font()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=40,
        bottomMargin=30
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'SummTitle',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#1F4D77'),
        alignment=1,
        fontBold=True
    )
    th_style = ParagraphStyle(
        'SummTH',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=8.5,
        leading=11,
        textColor=colors.white,
        alignment=1,
        fontBold=True
    )
    td_style_left = ParagraphStyle(
        'SummTDL',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=8.5,
        leading=11,
        alignment=0
    )
    td_style_center = ParagraphStyle(
        'SummTDC',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=8.5,
        leading=11,
        alignment=1
    )
    td_style_bold_center = ParagraphStyle(
        'SummTDBoldC',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=8.5,
        leading=11,
        alignment=1,
        fontBold=True
    )
    td_style_bold_left = ParagraphStyle(
        'SummTDBoldL',
        parent=styles['Normal'],
        fontName=hindi_font,
        fontSize=8.5,
        leading=11,
        alignment=0,
        fontBold=True
    )

    df_res = compute_district_event_summary_table(raw_df, target_district, categories)
    curr_time = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    all_dist_count = max(1, len(raw_df['जिला'].dropna().unique()))

    elements = [
        Paragraph(f"जिला {target_district} — सभी इवेंट केटेगरी में रैंक दिनांक {curr_time}", title_style),
        Spacer(1, 14)
    ]

    t_data = [[
        Paragraph('इवेंट केटेगरी', th_style),
        Paragraph('प्रविष्टि', th_style),
        Paragraph('रैंक-प्रविष्टि', th_style),
        Paragraph('प्रतिभागी', th_style),
        Paragraph('रैंक-प्रतिभागी', th_style),
        Paragraph('फोटो', th_style),
        Paragraph('रैंक-फोटो', th_style),
        Paragraph('वीडियो', th_style),
        Paragraph('रैंक-वीडियो', th_style)
    ]]

    t_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4D77')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#777777')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]

    for r_idx, (_, r) in enumerate(df_res.iterrows(), start=1):
        is_sum = bool(r['is_summary'])
        l_st = td_style_bold_left if is_sum else td_style_left
        c_st = td_style_bold_center if is_sum else td_style_center

        t_data.append([
            Paragraph(str(r['इवेंट केटेगरी']), l_st),
            Paragraph(f"{r['प्रविष्टि']:,}", c_st),
            Paragraph(str(r['रैंक-प्रविष्टि']), c_st),
            Paragraph(f"{r['प्रतिभागी']:,}", c_st),
            Paragraph(str(r['रैंक-प्रतिभागी']), c_st),
            Paragraph(f"{r['फोटो']:,}", c_st),
            Paragraph(str(r['रैंक-फोटो']), c_st),
            Paragraph(f"{r['वीडियो']:,}", c_st),
            Paragraph(str(r['रैंक-वीडियो']), c_st)
        ])

        if is_sum:
            # Highlight non-rank cells in summary row with #FEF3C7
            for col_idx in [0, 1, 3, 5, 7]:
                t_styles.append(('BACKGROUND', (col_idx, r_idx), (col_idx, r_idx), colors.HexColor('#FEF3C7')))

        # Conditional rank coloring on columns 2, 4, 6, 8
        for col_idx, col_name in [(2, 'रैंक-प्रविष्टि'), (4, 'रैंक-प्रतिभागी'), (6, 'रैंक-फोटो'), (8, 'रैंक-वीडियो')]:
            r_val = int(r[col_name])
            ratio = (r_val - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
            hex_bg = get_rank_3color(ratio)
            t_styles.append(('BACKGROUND', (col_idx, r_idx), (col_idx, r_idx), colors.HexColor(hex_bg)))

    col_widths = [150, 46, 52, 50, 52, 44, 48, 44, 48]
    table_obj = Table(t_data, colWidths=col_widths)
    table_obj.setStyle(TableStyle(t_styles))
    elements.append(table_obj)

    doc.build(elements)
    return buf.getvalue()

def generate_district_event_summary_printable_html(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> str:
    """
    Generates a web-viewable printable HTML page with sticky print toolbar.
    """
    body_html = generate_district_event_summary_pdf_html(raw_df, target_district, categories)
    toolbar = f'''
    <div style="position: sticky; top: 0; background: #1f4d77; color: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; z-index: 1000; box-shadow: 0 3px 8px rgba(0,0,0,0.2); border-radius: 6px; margin: 10px 12px 14px 12px;">
        <span style="font-weight: 700; font-size: 15px;">🖨️ जिला {target_district} — इवेंटवार सारांश रैंकिंग रिपोर्ट प्रिव्यू</span>
        <button onclick="window.print()" style="background: #f59e0b; color: #111827; font-weight: 700; border: none; padding: 7px 18px; border-radius: 5px; cursor: pointer; font-size: 14px;">
            🖨️ अभी प्रिंट करें / Save as PDF
        </button>
    </div>
    '''
    return body_html.replace('<body>', f'<body>\n{toolbar}')

@st.cache_data(show_spinner=False, max_entries=50)
def create_district_event_summary_excel(raw_df: pd.DataFrame, target_district: str, categories: list = None) -> bytes:
    """
    Generates an Excel workbook with styled table and 3-color conditional formatting on rank columns.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    df_res = compute_district_event_summary_table(raw_df, target_district, categories)
    all_dist_count = max(1, len(raw_df['जिला'].dropna().unique()))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"इवेंट_रैंकिंग_{target_district}"

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4D77", end_color="1F4D77", fill_type="solid")
    cell_font = Font(name="Calibri", size=10)
    bold_cell_font = Font(name="Calibri", size=10, bold=True)
    summary_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='777777'),
        right=Side(style='thin', color='777777'),
        top=Side(style='thin', color='777777'),
        bottom=Side(style='thin', color='777777')
    )

    curr_row = 1
    # Title
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=9)
    top_cell = ws.cell(row=curr_row, column=1, value=f"जिला {target_district} — सभी इवेंट केटेगरी में रैंक")
    top_cell.font = Font(name="Calibri", size=13, bold=True, color="1F4D77")
    top_cell.alignment = center_align
    curr_row += 1

    date_cell = ws.cell(row=curr_row, column=1, value=f"दिनांक {datetime.now().strftime('%d/%m/%Y %I:%M %p')}")
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=9)
    date_cell.font = Font(name="Calibri", size=10, bold=True, color="1F4D77")
    date_cell.alignment = center_align
    curr_row += 2

    # Headers
    headers = ['इवेंट केटेगरी', 'प्रविष्टि', 'रैंक-प्रविष्टि', 'प्रतिभागी', 'रैंक-प्रतिभागी', 'फोटो', 'रैंक-फोटो', 'वीडियो', 'रैंक-वीडियो']
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=curr_row, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    curr_row += 1

    # Data Rows
    rank_cols = ['रैंक-प्रविष्टि', 'रैंक-प्रतिभागी', 'रैंक-फोटो', 'रैंक-वीडियो']
    for _, row in df_res.iterrows():
        is_sum = bool(row['is_summary'])
        c_font = bold_cell_font if is_sum else cell_font

        for c_idx, h in enumerate(headers, 1):
            val = row[h]
            cell = ws.cell(row=curr_row, column=c_idx, value=val)
            cell.font = c_font
            cell.border = thin_border
            cell.alignment = left_align if c_idx == 1 else center_align

            if is_sum:
                cell.fill = summary_fill

            if h in rank_cols:
                r_val = int(val)
                ratio = (r_val - 1) / (all_dist_count - 1) if all_dist_count > 1 else 0.0
                hex_color = get_rank_3color(ratio).replace('#', '')
                cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

        curr_row += 1

    # Column widths
    col_widths = {1: 32, 2: 12, 3: 15, 4: 15, 5: 15, 6: 12, 7: 14, 8: 12, 9: 14}
    for c_idx, w in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@st.cache_data(show_spinner=False, max_entries=10)
def create_all_districts_event_summary_zip(raw_df: pd.DataFrame, categories: list = None) -> bytes:
    """
    Generates a ZIP bundle containing the event-wise summary PDF for EVERY district in Rajasthan.
    Uses multi-threaded rendering for high performance.
    """
    all_districts = sorted([d for d in raw_df['जिला'].dropna().unique() if str(d).strip()])

    def _make_pdf(dist):
        try:
            pdf_bytes = create_district_event_summary_pdf(raw_df, dist, categories)
            clean_name = re.sub(r'[\\/*?:"<>|]', '_', dist)
            fname = f"जिला_{clean_name}_इवेंटवार_रैंकिंग_सारांश.pdf"
            return fname, pdf_bytes
        except Exception:
            return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            future_to_dist = {executor.submit(_make_pdf, dist): dist for dist in all_districts}
            for future in concurrent.futures.as_completed(future_to_dist):
                res = future.result()
                if res and res[1]:
                    zip_file.writestr(res[0], res[1])

    return zip_buffer.getvalue()

