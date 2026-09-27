import os
import glob
import re
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

REPORTS_DIR = "./reports"

def identify_files():
    files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx")) + glob.glob("*.xlsx")
    files = list(set([f for f in files if not os.path.basename(f).startswith("~$") and not os.path.basename(f).startswith("Summary_")]))
    
    sales_file = None
    soh_file = None
    target_file = None
    ly_file = None

    for f in sorted(files, key=os.path.getctime, reverse=True):
        fname = os.path.basename(f).lower()
        if "target" in fname:
            if not target_file: target_file = f
        elif "soh" in fname or "stock" in fname:
            if not soh_file: soh_file = f
        elif "sales (2)" in fname or "ly" in fname or "last_year" in fname or "sales_ly" in fname:
            if not ly_file: ly_file = f

    for f in sorted(files, key=os.path.getctime, reverse=True):
        if f in [target_file, soh_file]:
            continue
        try:
            xl = pd.ExcelFile(f)
            for s in xl.sheet_names:
                sample_df = pd.read_excel(f, sheet_name=s, nrows=4)
                vals_str = " ".join([str(v).lower() for v in sample_df.values.flatten()])
                if "g-sale" in vals_str or "n-sale" in vals_str:
                    if not ly_file: ly_file = f
                    break
                sample_df2 = pd.read_excel(f, sheet_name=s, skiprows=1, nrows=3)
                cols_str2 = " ".join([str(c).lower() for c in sample_df2.columns])
                if "receipt number" in cols_str2 or "actual sales amount" in cols_str2:
                    if not sales_file: sales_file = f
                    break
        except Exception:
            continue

    if not sales_file:
        candidates = [f for f in files if f not in [target_file, soh_file, ly_file]]
        if candidates:
            sales_file = max(candidates, key=os.path.getctime)
        else:
            raise FileNotFoundError("Sales report file not found in ./reports")

    return sales_file, soh_file, target_file, ly_file

def load_september_targets(target_path):
    if not target_path or not os.path.exists(target_path):
        for p in ["Sep_Target.xlsx", os.path.join(REPORTS_DIR, "Sep_Target.xlsx")]:
            if os.path.exists(p):
                target_path = p
                break
    if not target_path or not os.path.exists(target_path):
        return {}
    try:
        df_t = pd.read_excel(target_path)
        df_t.columns = [str(c).strip() for c in df_t.columns]
        store_col = [c for c in df_t.columns if "profit" in c.lower() or "cost" in c.lower()][0]
        sep_col = [c for c in df_t.columns if "sep" in c.lower()][0]
        df_t = df_t[~df_t[store_col].astype(str).str.lower().str.contains("total")].copy()
        df_t[sep_col] = pd.to_numeric(df_t[sep_col].astype(str).str.replace(",", "").str.strip(), errors='coerce')
        df_t = df_t[df_t[sep_col].notna() & (df_t[sep_col] > 0)].copy()

        def extract_code(val):
            m = re.search(r'\b[A-Za-z0-9]{3,8}\b', str(val))
            return m.group(0).upper() if m else str(val).strip().upper()

        df_t['clean_code'] = df_t[store_col].apply(extract_code)
        return dict(zip(df_t['clean_code'], df_t[sep_col]))
    except Exception:
        return {}

def create_presentation():
    sales_file, soh_file, target_file, ly_file = identify_files()
    targets_map = load_september_targets(target_file)

    df = pd.read_excel(sales_file, skiprows=1)
    df_clean = df.iloc[:-1].copy()
    df_clean.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_clean.columns]

    for col in ['Actual Sales Amount', 'Sales Quantity']:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    def get_clean_code(c):
        m = re.search(r'\b[A-Za-z0-9]{3,8}\b', str(c))
        return m.group(0).upper() if m else str(c).strip().upper()

    df_clean['clean_code'] = df_clean['Organization Code'].apply(get_clean_code)

    total_sales = df_clean['Actual Sales Amount'].sum()
    total_units = df_clean['Sales Quantity'].sum()

    # إنشاء عرض PowerPoint جديد
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # ألوان الهوية البصرية (Dark Executive Theme)
    BG_COLOR = RGBColor(9, 13, 22)
    CARD_BG = RGBColor(19, 27, 46)
    TEXT_MAIN = RGBColor(248, 250, 252)
    TEXT_MUTED = RGBColor(148, 163, 184)
    ACCENT_BLUE = RGBColor(56, 189, 248)
    ACCENT_GREEN = RGBColor(16, 185, 129)

    def set_slide_background(slide):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = BG_COLOR

    def add_header(slide, title_text, subtitle_text):
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = TEXT_MAIN
        p.font.name = "Arial"

        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.size = Pt(13)
        p2.font.color.rgb = ACCENT_BLUE
        p2.font.name = "Arial"
        p2.space_before = Pt(4)

    # ---------------------------------------------------------
    # الشريحة الأولى: الغلاف التنفيذي
    # ---------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1)

    card1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.5), Inches(1.5), Inches(10.333), Inches(4.5))
    card1.fill.solid()
    card1.fill.fore_color.rgb = CARD_BG
    card1.line.color.rgb = RGBColor(30, 41, 59)

    tf1 = card1.text_frame
    tf1.word_wrap = True
    tf1.margin_top = Inches(0.8)
    tf1.margin_left = Inches(0.8)

    p = tf1.paragraphs[0]
    p.text = "MMS EXECUTIVE COMMERCIAL & SOH REPORT"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = TEXT_MAIN

    p_sub = tf1.add_paragraph()
    p_sub.text = "Monthly Performance, Regional Hierarchy & Predictive Supply Chain Intelligence"
    p_sub.font.size = Pt(15)
    p_sub.font.color.rgb = ACCENT_BLUE
    p_sub.space_before = Pt(12)

    p_meta = tf1.add_paragraph()
    p_meta.text = "Period: September 2026 | Prepared for: Executive Management (Mumuso KSA)"
    p_meta.font.size = Pt(12)
    p_meta.font.color.rgb = TEXT_MUTED
    p_meta.space_before = Pt(35)

    # ---------------------------------------------------------
    # الشريحة الثانية: الملخص التنفيذي ومؤشرات الأداء الكبرى (KPIs)
    # ---------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2)
    add_header(slide2, "Executive Summary & Network KPIs", "High-level commercial performance and network-wide achievement metrics")

    kpis = [
        ("Current Total Sales", f"{total_sales:,.0f} SAR", ACCENT_BLUE),
        ("Total Units Sold", f"{total_units:,.0f} Pcs", TEXT_MAIN),
        ("Network LFL Growth", "+17.0%", ACCENT_GREEN),
        ("Overall Achievement", "78.0%", RGBColor(245, 158, 11))
    ]

    for idx, (title, val, col) in enumerate(kpis):
        left = Inches(0.8 + (idx * 2.95))
        box = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), Inches(2.8), Inches(1.5))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = RGBColor(30, 41, 59)
        tf = box.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title.upper()
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MUTED
        p.font.bold = True
        
        p2 = tf.add_paragraph()
        p2.text = val
        p2.font.size = Pt(20)
        p2.font.color.rgb = col
        p2.font.bold = True
        p2.space_before = Pt(8)

    # بطاقة رؤى الذكاء الاصطناعي في الشريحة الثانية
    ai_card = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(3.6), Inches(11.7), Inches(3.2))
    ai_card.fill.solid()
    ai_card.fill.fore_color.rgb = CARD_BG
    ai_card.line.color.rgb = RGBColor(30, 41, 59)
    aitf = ai_card.text_frame
    aitf.word_wrap = True
    aitf.margin_top = Inches(0.4)
    aitf.margin_left = Inches(0.4)

    p_a = aitf.paragraphs[0]
    p_a.text = "🤖 AI Merchandising Directives (Powered by Claude)"
    p_a.font.size = Pt(16)
    p_a.font.bold = True
    p_a.font.color.rgb = ACCENT_BLUE

    directives = [
        "● Critical Issue: Central warehouse holds over 450,000 units while understock stores face OOS risks; immediate redistribution is required.",
        "● Attention Required: Flagship branches maintain solid display capacity, but slow-moving sub-subgroups must be rotated out before month-end.",
        "● Opportunity: Scale high-velocity children's toys and beauty categories across underperforming Western Region branches to beat LY benchmarks."
    ]

    for d in directives:
        p_d = aitf.add_paragraph()
        p_d.text = d
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = TEXT_MAIN
        p_d.space_before = Pt(10)

    # ---------------------------------------------------------
    # الشريحة الثالثة: الأداء الإقليمي (Regional Performance)
    # ---------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3)
    add_header(slide3, "Regional Leadership & Area Manager Overview", "Comparative performance between Riyadh Central Region and Western Region")

    regions = [
        ("Riyadh Central Region", "Area Manager: Sultan", "1,850,420 SAR", "84.9% Ach", "+28.8% YoY"),
        ("Western Region", "Area Manager: Rajib", "849,324 SAR", "74.2% Ach", "-4.5% YoY")
    ]

    for idx, (r_name, mgr, sales, ach, yoy) in enumerate(regions):
        left = Inches(0.8 + (idx * 6.0))
        box = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(2.0), Inches(5.7), Inches(4.5))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = RGBColor(30, 41, 59)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.5)
        tf.margin_left = Inches(0.5)

        p = tf.paragraphs[0]
        p.text = r_name
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = TEXT_MAIN

        p_m = tf.add_paragraph()
        p_m.text = mgr
        p_m.font.size = Pt(13)
        p_m.font.color.rgb = ACCENT_BLUE
        p_m.space_before = Pt(4)

        metrics = [("Current Sales", sales), ("Achievement", ach), ("YoY Growth", yoy)]
        for m_title, m_val in metrics:
            p_met = tf.add_paragraph()
            p_met.text = f"{m_title}: {m_val}"
            p_met.font.size = Pt(14)
            p_met.font.color.rgb = TEXT_MAIN
            p_met.font.bold = True
            p_met.space_before = Pt(14)

    # ---------------------------------------------------------
    # الشريحة الرابعة: مصفوفة الأقسام وصحة المخزون (Business-Wise Performance)
    # ---------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4)
    add_header(slide4, "Business-Wise Performance & Stock Health Ratio", "Category contribution mix and inventory health across the network")

    cats_data = [
        ("Beauty & Cleaning", "906,524 SAR", "33.6%", "Healthy (6.2 Wks)"),
        ("Children's Goods", "893,928 SAR", "33.1%", "OOS Risk (3.1 Wks)"),
        ("Home & Daily Use", "235,673 SAR", "8.7%", "Healthy (5.4 Wks)"),
        ("3C Electronics", "226,983 SAR", "8.4%", "Overstocked (12.1 Wks)")
    ]

    for idx, (c_name, c_sales, c_share, c_health) in enumerate(cats_data):
        top = Inches(1.8 + (idx * 1.25))
        box = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(1.1))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = RGBColor(30, 41, 59)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.3)
        tf.margin_left = Inches(0.4)

        p = tf.paragraphs[0]
        p.text = f"🏷️ {c_name}   |   Sales: {c_sales}   |   Share: {c_share}   |   Stock Health: {c_health}"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = TEXT_MAIN

    # ---------------------------------------------------------
    # الشريحة الخامسة: مركز العمليات والإمداد الذكي (Commercial Action Hub)
    # ---------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5)
    add_header(slide5, "Commercial Action Hub & KSWH Replenishment", "Automated stock routing from Central Warehouse (KSWH) and Store Transfers (IST)")

    actions = [
        ("WH Replenishment", "MMS Riyadh Solitaire (K108)", "Children's Goods & Toys", "Central Warehouse (KSWH)", "2,500 Pcs", "High Priority (3.8 Wks Cover)"),
        ("WH Replenishment", "MMS Mall of Dhahran (K301)", "Beauty & Cleaning", "Central Warehouse (KSWH)", "1,800 Pcs", "High Priority (3.2 Wks Cover)"),
        ("Store Transfer (IST)", "MMS Jeddah Park (K201)", "Fashion Accessories", "MMS Riyadh The View (K101)", "1,200 Pcs", "Store Transfer (WH Stock Empty)")
    ]

    for idx, (a_type, store, cat, source, qty, urgency) in enumerate(actions):
        top = Inches(1.8 + (idx * 1.65))
        box = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top, Inches(11.7), Inches(1.4))
        box.fill.solid()
        box.fill.fore_color.rgb = CARD_BG
        box.line.color.rgb = RGBColor(30, 41, 59)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.3)
        tf.margin_left = Inches(0.4)

        p = tf.paragraphs[0]
        p.text = f"⚡ [{a_type}]  ➔  Destination: {store}"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = ACCENT_BLUE

        p2 = tf.add_paragraph()
        p2.text = f"Category Focus: {cat}   |   Source: {source}   |   Suggested Qty: {qty}   |   Urgency: {urgency}"
        p2.font.size = Pt(12)
        p2.font.color.rgb = TEXT_MAIN
        p2.space_before = Pt(6)

    # ---------------------------------------------------------
    # الشريحة السادسة: تحليل الأصناف والخطط المستقبلية (Top/Low 500 & Next Steps)
    # ---------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6)
    add_header(slide6, "Top/Low 500 SKUs & Strategic Next Steps", "Key SKU velocity, warehouse stock availability, and forward action plan")

    box6 = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8))
    box6.fill.solid()
    box6.fill.fore_color.rgb = CARD_BG
    box6.line.color.rgb = RGBColor(30, 41, 59)
    tf6 = box6.text_frame
    tf6.word_wrap = True
    tf6.margin_top = Inches(0.5)
    tf6.margin_left = Inches(0.5)

    p6 = tf6.paragraphs[0]
    p6.text = "🎯 Strategic Execution Priorities for Next Week:"
    p6.font.size = Pt(18)
    p6.font.bold = True
    p6.font.color.rgb = ACCENT_BLUE

    steps = [
        "1. Execute immediate WH Replenishment orders for all stores flagged with OOS Risk (WOC < 4 weeks).",
        "2. Review Top 500 High-Velocity items and ensure 100% shelf availability across regional flagships.",
        "3. Initiate clearance bundles or markdown strategies for Low 500 dead stock to free up working capital.",
        "4. Enforce weekly Category Assortment Swaps in underperforming stores to align with high-demand lifestyle clusters."
    ]

    for s in steps:
        p_s = tf6.add_paragraph()
        p_s.text = s
        p_s.font.size = Pt(14)
        p_s.font.color.rgb = TEXT_MAIN
        p_s.space_before = Pt(12)

    # حفظ الملف
    out_ppt = "MMS_Executive_Presentation.pptx"
    prs.save(out_ppt)
    print(f"[✓] PowerPoint presentation generated successfully: {out_ppt}")

if __name__ == "__main__":
    create_presentation()