import os
import glob
import re
import json
import pandas as pd
import numpy as np
import anthropic

REPORTS_DIR = "./reports"

STORE_MAPPING = {
    "K101": {"full_name": "MMS Riyadh The View Mall", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K102": {"full_name": "MMS Riyadh Tala Mall", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K108": {"full_name": "MMS Riyadh Solitaire", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K109": {"full_name": "MMS Riyadh Localizer", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K110": {"full_name": "MMS Riyadh U-Walk", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K130": {"full_name": "MMS Riyadh Al-Rabwa", "region": "Riyadh Central Region", "manager": "Sultan"},
    "K301": {"full_name": "MMS Mall of Dhahran", "region": "Riyadh Central Region", "manager": "Sultan"},
    
    "K201": {"full_name": "MMS Jeddah Park", "region": "Western Region", "manager": "Rajib"},
    "K202": {"full_name": "MMS Jeddah Yasmin Mall", "region": "Western Region", "manager": "Rajib"},
    "K205": {"full_name": "MMS Jeddah U-Walk", "region": "Western Region", "manager": "Rajib"},
    "K208": {"full_name": "MMS Jeddah Mall of Arabia", "region": "Western Region", "manager": "Rajib"},
    "K210": {"full_name": "MMS Makkah Salam Mall", "region": "Western Region", "manager": "Rajib"},
    "K211": {"full_name": "MMS Madinah", "region": "Western Region", "manager": "Rajib"},
    "K401": {"full_name": "MMS Najran Park", "region": "Western Region", "manager": "Rajib"},
    "K403": {"full_name": "MMS RMJ", "region": "Western Region", "manager": "Rajib"},
    "K404": {"full_name": "MMS Abha", "region": "Western Region", "manager": "Rajib"},
    "K501": {"full_name": "MMS Tabuk Park", "region": "Western Region", "manager": "Rajib"}
}

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

    # فحص محتوى الملف للتعرف على ملف المبيعات وملف العام الماضي بدقة
    for f in sorted(files, key=os.path.getctime, reverse=True):
        if f in [target_file, soh_file]:
            continue
        try:
            xl = pd.ExcelFile(f)
            for s in xl.sheet_names:
                sample_df = pd.read_excel(f, sheet_name=s, nrows=4)
                cols_str = " ".join([str(c).lower() for c in sample_df.columns])
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

def load_ly_sales_data(ly_path):
    if not ly_path or not os.path.exists(ly_path):
        print("[!] No Last Year (LY) sales file found.")
        return {}
    
    print(f"[*] Processing LY Sales file: {ly_path}")
    ly_totals = {}
    try:
        df_ly = pd.read_excel(ly_path, sheet_name="Sales" if "Sales" in pd.ExcelFile(ly_path).sheet_names else 0)
        
        # تحديد سطر G-Sale (Gross Sales)
        sale_type_col = next((c for c in df_ly.columns if any(df_ly[c].astype(str).str.strip().str.upper() == 'G-SALE')), None)
        if not sale_type_col:
            sale_type_col = df_ly.columns[2]

        df_gsale = df_ly[df_ly[sale_type_col].astype(str).str.strip().str.upper() == 'G-SALE'].copy()
        if df_gsale.empty:
            df_gsale = df_ly.copy()

        # أخذ أعمدة Mumuso (MMS) واستبعاد Dazzle (DZL)
        for col in df_ly.columns:
            col_str = str(col).strip()
            if "(MMS)" in col_str.upper() and "(DZL)" not in col_str.upper():
                m = re.search(r'^\d+', col_str)
                if m:
                    code = f"K{m.group(0)}"
                    tot_val = pd.to_numeric(df_gsale[col], errors='coerce').sum()
                    ly_totals[code] = round(float(tot_val), 2)
                    
        print(f"[✓] Extracted LY Gross Sales for {len(ly_totals)} Mumuso stores.")
        return ly_totals
    except Exception as e:
        print(f"[!] Error processing LY sales file: {e}")
        return {}

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
    except Exception as e:
        print(f"Target load error: {e}")
        return {}

def load_soh_data(soh_path):
    if not soh_path or not os.path.exists(soh_path):
        return {}, {}
    
    print(f"[*] Processing SOH file: {soh_path}")
    soh_store_summary = {}
    soh_hierarchy_map = {}
    try:
        xl = pd.ExcelFile(soh_path)
        sheet_to_use = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
        
        df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use, skiprows=1)
        df_soh.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]
        
        if "avail_stock" not in [c.lower() for c in df_soh.columns] and "current_stock" not in [c.lower() for c in df_soh.columns]:
            df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use)
            df_soh.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]

        code_col = next((c for c in df_soh.columns if c.lower() in ["org code", "organization code", "org_code", "store code"]), None)
        stock_col = next((c for c in df_soh.columns if c.lower() in ["avail_stock", "current_stock"]), None)
        price_col = next((c for c in df_soh.columns if "retail_price" in c.lower() or "price" in c.lower()), None)
        cat_col = next((c for c in df_soh.columns if c.lower() == "category"), None)
        pg_col = next((c for c in df_soh.columns if c.lower() in ["product_group", "product group"]), None)

        if not code_col or not stock_col:
            return {}, {}

        df_soh = df_soh[df_soh[code_col].notna()].copy()
        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)
        
        if cat_col and pg_col:
            pairs = df_soh[[cat_col, pg_col]].drop_duplicates().dropna()
            for _, r in pairs.iterrows():
                pg_val = str(r[pg_col]).strip()
                c_val = str(r[cat_col]).replace('_', ' ').replace('’', "'").strip()
                if pg_val:
                    soh_hierarchy_map[pg_val.lower()] = c_val

        if price_col:
            df_soh[price_col] = pd.to_numeric(df_soh[price_col], errors='coerce').fillna(0)
            df_soh['stock_val'] = df_soh[stock_col] * df_soh[price_col]
        else:
            df_soh['stock_val'] = 0

        grouped = df_soh.groupby(code_col).agg(
            soh_units=(stock_col, 'sum'),
            soh_val=('stock_val', 'sum')
        ).reset_index()

        def clean_c(v):
            m = re.search(r'\b[A-Za-z0-9]{3,8}\b', str(v))
            return m.group(0).upper() if m else str(v).strip().upper()

        grouped['clean_code'] = grouped[code_col].apply(clean_c)
        soh_store_summary = grouped.set_index('clean_code').to_dict(orient='index')
        return soh_store_summary, soh_hierarchy_map
    except Exception as e:
        print(f"[!] Error processing SOH file: {e}")
        return {}, {}

def generate_claude_insights(store_summary, total_sales, total_target, overall_ach, total_soh_units, lfl_growth_pct):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        return {
            "critical": f"Like-For-Like (LFL) sales growth stands at {lfl_growth_pct:+.1f}% vs LY Gross Sales; store assortment rotation is needed to sustain momentum in mature stores.",
            "attention": "Preserve 40,000-80,000 visual merchandise units in regional flagships while rotating out stagnant sub-categories.",
            "opportunity": "Scale high-velocity children's toys and beauty categories across underperforming Western Region branches to beat LY benchmarks."
        }

    top_stores = store_summary.head(3)[['full_name', 'sales', 'ach_pct', 'ly_sales', 'yoy_growth']].to_dict(orient="records")
    prompt = f"""
    You are a Senior Merchandising Director for Mumuso.
    - Total Sales: {total_sales:,.0f} SAR | Target: {total_target:,.0f} SAR | Ach: {overall_ach:.1f}%
    - Like-For-Like (LFL) YoY Growth: {lfl_growth_pct:+.1f}%
    - Top Stores: {top_stores}
    Provide 3 punchy commercial directives (1 sentence each):
    1. Critical Issues
    2. Attention Required
    3. Opportunities
    Respond ONLY in valid JSON: {{"critical": "...", "attention": "...", "opportunity": "..."}}
    """
    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}]
        )
        content = response.content[0].text.strip()
        if "```" in content:
            content = re.search(r'\{.*\}', content, re.DOTALL).group(0)
        return json.loads(content)
    except Exception:
        return {
            "critical": f"Like-For-Like (LFL) stores show {lfl_growth_pct:+.1f}% YoY trajectory against last year's gross sales baseline.",
            "attention": "Ensure balanced 40k-80k display capacity without clogging gondolas with slow-moving sub-subgroups.",
            "opportunity": "Drive cross-selling on high-margin accessory clusters to further expand positive YoY spread."
        }

def build_svg_bar_chart(chart_stores):
    svg_w, svg_h = 900, 320
    pad_left, pad_right, pad_top, pad_bottom = 60, 20, 30, 60
    plot_w = svg_w - pad_left - pad_right
    plot_h = svg_h - pad_top - pad_bottom

    max_val = max(chart_stores['sales'].max(), chart_stores['target'].fillna(0).max()) * 1.15
    if max_val == 0: max_val = 1

    n_stores = len(chart_stores)
    slot_w = plot_w / max(n_stores, 1)
    bar_w = min(slot_w * 0.35, 28)

    grid_lines = ""
    for i in range(5):
        val = (max_val / 4) * i
        y_pos = pad_top + plot_h - (i * (plot_h / 4))
        grid_lines += f"""
        <line x1="{pad_left}" y1="{y_pos}" x2="{svg_w - pad_right}" y2="{y_pos}" stroke="#1e293b" stroke-width="1" />
        <text x="{pad_left - 10}" y="{y_pos + 4}" fill="#94a3b8" font-size="11" text-anchor="end">{int(val):,}</text>
        """

    bars_svg = ""
    for idx, (_, r) in enumerate(chart_stores.iterrows()):
        slot_center = pad_left + (idx + 0.5) * slot_w
        s_val = r['sales']
        t_val = r['target'] if pd.notna(r['target']) else 0

        s_h = (s_val / max_val) * plot_h
        t_h = (t_val / max_val) * plot_h

        s_x = slot_center - bar_w - 2
        s_y = pad_top + plot_h - s_h
        t_x = slot_center + 2
        t_y = pad_top + plot_h - t_h

        name = str(r['full_name']).replace("MMS Riyadh ", "").replace("MMS ", "")
        if len(name) > 12: name = name[:11] + ".."

        bars_svg += f"""
        <rect x="{s_x:.1f}" y="{s_y:.1f}" width="{bar_w:.1f}" height="{s_h:.1f}" rx="3" fill="#38bdf8">
            <title>{r['full_name']} Sales: {s_val:,.2f} SAR</title>
        </rect>
        <rect x="{t_x:.1f}" y="{t_y:.1f}" width="{bar_w:.1f}" height="{t_h:.1f}" rx="3" fill="#334155">
            <title>{r['full_name']} Target: {t_val:,.2f} SAR</title>
        </rect>
        <text x="{slot_center:.1f}" y="{svg_h - 25}" fill="#cbd5e1" font-size="10" font-weight="600" text-anchor="middle">{name}</text>
        """

    return f"""<svg viewBox="0 0 {svg_w} {svg_h}" style="width:100%; height:auto; display:block;">{grid_lines}{bars_svg}</svg>"""

def process_and_build():
    sales_file, soh_file, target_file, ly_file = identify_files()
    print(f"[*] Sales Report: {sales_file}")
    print(f"[*] SOH File:     {soh_file}")
    print(f"[*] Target File:  {target_file}")
    print(f"[*] LY File:      {ly_file}")

    targets_map = load_september_targets(target_file)
    soh_map, soh_hier_map = load_soh_data(soh_file)
    ly_sales_map = load_ly_sales_data(ly_file)

    df = pd.read_excel(sales_file, skiprows=1)
    df_clean = df.iloc[:-1].copy()
    df_clean.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    raw_subsub_col = next((c for c in df_clean.columns if c.lower() in ['category name', 'product_category', 'category']), None)
    if not raw_subsub_col:
        df_clean['raw_subsub'] = "General Items"
        raw_subsub_col = 'raw_subsub'

    df_clean['sub_subgroup'] = df_clean[raw_subsub_col].fillna("Other").astype(str).str.strip()

    def map_to_main_category(subsub):
        sub_l = str(subsub).strip().lower()
        if sub_l in soh_hier_map:
            return soh_hier_map[sub_l]
        if any(x in sub_l for x in ['toy', 'doll', 'clay', 'puzzle', 'baby', 'block', 'gun', 'bubble']):
            return "Children's Goods"
        if any(x in sub_l for x in ['lip', 'mask', 'cream', 'perfume', 'makeup', 'eyebrow', 'clean', 'wipe', 'bath', 'nail', 'soap']):
            return "Beauty & Cleaning"
        if any(x in sub_l for x in ['pen', 'notebook', 'tape', 'sticker', 'stationery', 'pencil', 'eraser']):
            return "Stationery"
        if any(x in sub_l for x in ['cup', 'mat', 'storage', 'kitchen', 'umbrella', 'fragrance', 'hanger', 'mat']):
            return "Home & Daily Use"
        if any(x in sub_l for x in ['cable', 'headphone', 'fan', 'usb', 'charger', 'watch', 'phone']):
            return "3C Electronics"
        if any(x in sub_l for x in ['bag', 'backpack', 'wallet', 'purse']):
            return "Bags"
        if any(x in sub_l for x in ['sock', 'slipper', 'hat', 'sunglass', 'glove']):
            return "Apparel Accessories"
        if any(x in sub_l for x in ['hair', 'earring', 'clip', 'necklace', 'jewelry']):
            return "Fashion Accessories"
        if any(x in sub_l for x in ['pillow', 'towel', 'cushion', 'eyemask']):
            return "Home Textile"
        return "Variety Lifestyle"

    df_clean['main_category'] = df_clean['sub_subgroup'].apply(map_to_main_category)

    # 1. إجماليات المتاجر
    store_summary = df_clean.groupby(['Organization Code', 'Organization Name']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index()

    def get_clean_code(c):
        m = re.search(r'\b[A-Za-z0-9]{3,8}\b', str(c))
        return m.group(0).upper() if m else str(c).strip().upper()

    store_summary['clean_code'] = store_summary['Organization Code'].apply(get_clean_code)
    store_summary['full_name'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('full_name', str(r['Organization Name'])), axis=1
    )
    store_summary['region'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('region', 'Western Region' if 'JED' in str(r['Organization Name']).upper() or 'K2' in r['clean_code'] else 'Riyadh Central Region'), axis=1
    )
    store_summary['manager'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('manager', 'Sultan' if r['region'] == 'Riyadh Central Region' else 'Rajib'), axis=1
    )

    store_summary['atv'] = (store_summary['sales'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['upt'] = (store_summary['units'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['asp'] = (store_summary['sales'] / store_summary['units'].replace(0, np.nan)).fillna(0).round(2)

    total_sales = store_summary['sales'].sum()
    total_txns = store_summary['txns'].sum()
    total_units = store_summary['units'].sum()
    network_atv = (total_sales / total_txns) if total_txns > 0 else 0
    network_upt = (total_units / total_txns) if total_units > 0 else 0
    network_asp = (total_sales / total_units) if total_units > 0 else 0

    store_summary['share'] = ((store_summary['sales'] / total_sales) * 100).round(2)
    store_summary = store_summary.sort_values(by='sales', ascending=False).reset_index(drop=True)

    # دمج مبيعات العام الماضي (LY Gross Sales) بدقة
    store_summary['ly_sales'] = store_summary['clean_code'].map(ly_sales_map)
    store_summary['yoy_growth'] = store_summary.apply(
        lambda r: ((r['sales'] - r['ly_sales']) / r['ly_sales'] * 100) if pd.notna(r['ly_sales']) and r['ly_sales'] > 0 else None,
        axis=1
    )

    # حساب إجمالي LFL (Like-For-Like) على مستوى الشبكة
    lfl_stores = store_summary[store_summary['ly_sales'].notna()].copy()
    total_current_lfl_sales = lfl_stores['sales'].sum()
    total_ly_sales = lfl_stores['ly_sales'].sum()
    network_lfl_growth = ((total_current_lfl_sales - total_ly_sales) / total_ly_sales * 100) if total_ly_sales > 0 else 0

    # ربط SOH
    def match_soh(c_code):
        if c_code in soh_map: return soh_map[c_code]
        for k, v in soh_map.items():
            if k in c_code or c_code in k: return v
        return {"soh_units": 0, "soh_val": 0}

    soh_matched = store_summary['clean_code'].apply(match_soh)
    store_summary['soh_units'] = [x['soh_units'] for x in soh_matched]
    store_summary['soh_val'] = [x['soh_val'] for x in soh_matched]
    total_soh_units = store_summary['soh_units'].sum()

    # مطابقة الأهداف
    def match_target(row):
        c_code = row['clean_code']
        raw_name = str(row['Organization Name']).upper()
        if c_code in targets_map: return targets_map[c_code]
        for k, v in targets_map.items():
            if str(k).upper() in c_code or str(k).upper() in raw_name: return v
        return None

    store_summary['target'] = store_summary.apply(match_target, axis=1)
    store_summary['ach_pct'] = store_summary.apply(
        lambda r: (r['sales'] / r['target'] * 100) if pd.notna(r['target']) and r['target'] > 0 else None,
        axis=1
    )

    valid_targets = store_summary[store_summary['target'].notna()]
    total_target = valid_targets['target'].sum()
    sales_with_target = valid_targets['sales'].sum()
    overall_ach = (sales_with_target / total_target * 100) if total_target > 0 else 0

    # 2. الهيكل السلعي للمبيعات: المستوى الأول (Main Category)
    main_cat_summary = df_clean.groupby('main_category').agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='sales', ascending=False).reset_index(drop=True)
    main_cat_summary['contribution'] = ((main_cat_summary['sales'] / total_sales) * 100).round(2)
    main_cat_summary['asp'] = (main_cat_summary['sales'] / main_cat_summary['units'].replace(0, np.nan)).fillna(0).round(2)

    main_cat_store = df_clean.groupby(['main_category', 'Organization Code'])['Actual Sales Amount'].sum().reset_index()
    top_store_per_main_cat = {}
    for c_name, grp in main_cat_store.groupby('main_category'):
        best = grp.sort_values(by='Actual Sales Amount', ascending=False).iloc[0]
        c_code = get_clean_code(best['Organization Code'])
        st_name = STORE_MAPPING.get(c_code, {}).get('full_name', best['Organization Code'])
        top_store_per_main_cat[c_name] = f"{st_name} ({best['Actual Sales Amount']:,.0f} SAR)"
    main_cat_summary['leading_store'] = main_cat_summary['main_category'].map(top_store_per_main_cat).fillna("-")

    # 3. الهيكل السلعي: المستوى الرابع (Sub-Subgroups)
    subsub_summary = df_clean.groupby(['main_category', 'sub_subgroup']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='sales', ascending=False).reset_index(drop=True)
    subsub_summary['contribution'] = ((subsub_summary['sales'] / total_sales) * 100).round(2)
    subsub_summary['asp'] = (subsub_summary['sales'] / subsub_summary['units'].replace(0, np.nan)).fillna(0).round(2)

    # 4. تفاصيل الأصناف الدقيقة داخل كل متجر
    store_cat_details = {}
    grouped_st_cat = df_clean.groupby(['Organization Code', 'main_category', 'sub_subgroup']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index()

    for code, grp in grouped_st_cat.groupby('Organization Code'):
        c_code = get_clean_code(code)
        st_total = grp['sales'].sum()
        cats_list = []
        for _, r in grp.sort_values(by='sales', ascending=False).iterrows():
            c_sales = r['sales']
            share = (c_sales / st_total * 100) if st_total > 0 else 0
            asp_item = (c_sales / r['units']) if r['units'] > 0 else 0
            cats_list.append({
                "main_category": r['main_category'],
                "sub_subgroup": r['sub_subgroup'],
                "sales": f"{c_sales:,.2f}",
                "units": f"{int(r['units']):,}",
                "share": f"{share:.1f}%",
                "asp": f"{asp_item:,.2f}"
            })
        store_cat_details[c_code] = cats_list

    # محرك تشخيص Mumuso
    top_main_cats = set(main_cat_summary.head(3)['main_category'])
    def mumuso_commercial_engine(row):
        st_code = row['clean_code']
        soh = row['soh_units']
        ach = row['ach_pct'] if pd.notna(row['ach_pct']) else 0

        st_items = store_cat_details.get(st_code, [])
        st_top_cats = list(dict.fromkeys([c['main_category'] for c in st_items[:5]]))
        missing_cats = [c for c in top_main_cats if c not in st_top_cats]

        if soh >= 80000:
            cap_badge, cap_col = "Flagship Mega-Display", "#38bdf8"
        elif 40000 <= soh < 80000:
            cap_badge, cap_col = "Standard Full Display", "#10b981"
        elif 0 < soh < 40000:
            cap_badge, cap_col = "Lean Floor Stock", "#f59e0b"
        else:
            cap_badge, cap_col = "No SOH Synced", "#64748b"

        if ach >= 95:
            diag_title, diag_col = "Powerhouse Performer", "#10b981"
            prob = f"High commercial conversion ({ach:.1f}% Ach). Strong momentum in {', '.join(st_top_cats[:2])}."
            action = f"Maintain 100% shelf availability on leading sub-categories and introduce premium novelty SKUs."
            needs = f"Priority replenishment for core volume drivers in {st_top_cats[0] if st_top_cats else 'Toys'}."
        elif ach < 70 and soh >= 40000:
            diag_title, diag_col = "Assortment Mismatch", "#f59e0b"
            prob = f"Store holds solid display depth ({soh:,.0f} Pcs) but turnover is slow ({ach:.1f}% Ach). Gondolas tied to slow sub-subgroups."
            action = f"⚡ ACTION: Execute Category Assortment Swap. Reallocate front entrance to {missing_cats[0] if missing_cats else 'Children Toys & Beauty'} and bundle slow movers."
            needs = f"Inject high-velocity categories ({', '.join(missing_cats[:2]) if missing_cats else 'Toys & Novelties'})."
        elif ach < 70 and soh < 40000:
            diag_title, diag_col = "Under-Display Deficit", "#ef4444"
            prob = f"Target achievement is lagging ({ach:.1f}%) and visual density ({soh:,.0f} Pcs) is thin, depressing walk-in impulse purchases."
            action = f"⚡ ACTION: Increase store display depth to Mumuso visual benchmark (40k-80k Pcs) across core lifestyle categories."
            needs = "Floor-fill replenishment: +10,000 to +15,000 units of fast-moving impulse items."
        else:
            diag_title, diag_col = "Steady Flow", "#38bdf8"
            prob = f"Balanced run-rate ({ach:.1f}% Ach) with healthy display volume ({soh:,.0f} Pcs)."
            action = f"Focus cashier upselling to lift ATV (Current: {row['atv']:.1f} SAR) and rotate seasonal novelty end-caps."
            needs = "Routine weekly assortment replenishment and promotional feature rotation."

        return {
            "capacity_badge": cap_badge, "capacity_color": cap_col,
            "diag_title": diag_title, "diag_color": diag_col,
            "problem": prob, "action": action, "needs": needs,
            "top_categories_str": ", ".join(st_top_cats[:3]) if st_top_cats else "General"
        }

    engine_res = store_summary.apply(mumuso_commercial_engine, axis=1)
    store_summary['display_status'] = [e['capacity_badge'] for e in engine_res]
    store_summary['display_color'] = [e['capacity_color'] for e in engine_res]
    store_summary['diag_title'] = [e['diag_title'] for e in engine_res]
    store_summary['diag_color'] = [e['diag_color'] for e in engine_res]
    store_summary['problem'] = [e['problem'] for e in engine_res]
    store_summary['action'] = [e['action'] for e in engine_res]
    store_summary['needs'] = [e['needs'] for e in engine_res]
    store_summary['top_cats_str'] = [e['top_categories_str'] for e in engine_res]

    insights = generate_claude_insights(store_summary, total_sales, total_target, overall_ach, total_soh_units, network_lfl_growth)
    chart_stores = store_summary.head(8)
    chart_svg_markup = build_svg_bar_chart(chart_stores)

    # 5. بناء بطاقات الأقسام الرئيسية
    colors = ['#38bdf8', '#818cf8', '#a855f7', '#ec4899', '#f59e0b', '#10b981', '#06b6d4', '#e11d48', '#84cc16']
    main_cat_cards_html = ""
    for idx, r in main_cat_summary.iterrows():
        c_color = colors[idx % len(colors)]
        main_cat_cards_html += f"""
        <div onclick="filterByMainCategory('{r['main_category']}')" style="background:var(--card); border:1px solid var(--border); border-top:3px solid {c_color}; border-radius:10px; padding:16px; min-width:210px; max-width:240px; flex:1; cursor:pointer;" title="Click to filter sub-categories under {r['main_category']}">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                <span style="font-size:13px; font-weight:700; color:#fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{r['main_category']}</span>
                <span style="font-size:12px; font-weight:700; color:{c_color};">{r['contribution']:.1f}%</span>
            </div>
            <div style="font-size:17px; font-weight:700; color:#f8fafc; margin-bottom:6px;">{r['sales']:,.0f} <span style="font-size:11px; color:#94a3b8;">SAR</span></div>
            <div style="background:#090d16; border-radius:4px; height:5px; overflow:hidden;">
                <div style="background:{c_color}; width:{min(r['contribution'], 100):.1f}%; height:100%;"></div>
            </div>
            <div style="display:flex; justify-content:space-between; margin-top:8px; font-size:11px; color:#94a3b8;">
                <span>Units: {int(r['units']):,}</span>
                <span>ASP: {r['asp']:,.1f} SAR</span>
            </div>
        </div>
        """

    # 6. بناء جداول ومؤشرات المناطق (مع مقارنة مبيعات العام الماضي LY)
    region_kpi_cards = ""
    region_tables_html = ""

    for reg_name, grp in [("Riyadh Central Region", store_summary[store_summary['region'] == "Riyadh Central Region"]),
                          ("Western Region", store_summary[store_summary['region'] == "Western Region"])]:
        reg_mgr = "Sultan" if "Riyadh" in reg_name else "Rajib"
        r_sales = grp['sales'].sum()
        r_target = grp['target'].fillna(0).sum()
        r_ach = (r_sales / r_target * 100) if r_target > 0 else 0
        r_units = grp['units'].sum()
        r_txns = grp['txns'].sum()
        r_soh = grp['soh_units'].sum()
        r_atv = (r_sales / r_txns) if r_txns > 0 else 0
        r_upt = (r_units / r_txns) if r_txns > 0 else 0
        r_asp = (r_sales / r_units) if r_units > 0 else 0

        # حساب نمو مبيعات المنطقة مقارنة بالعام الماضي (LFL)
        reg_lfl = grp[grp['ly_sales'].notna()]
        reg_cur_lfl = reg_lfl['sales'].sum()
        reg_ly_tot = reg_lfl['ly_sales'].sum()
        reg_yoy = ((reg_cur_lfl - reg_ly_tot) / reg_ly_tot * 100) if reg_ly_tot > 0 else None

        ach_col = "#10b981" if r_ach >= 100 else ("#f59e0b" if r_ach >= 80 else "#ef4444")
        yoy_col = "#10b981" if (reg_yoy and reg_yoy >= 0) else "#ef4444"
        yoy_badge = f'<span style="color:{yoy_col}; font-weight:700;">{reg_yoy:+.1f}%</span>' if reg_yoy is not None else '<span style="color:#64748b;">N/A</span>'

        region_kpi_cards += f"""
        <div style="background:var(--card); border:1px solid var(--border); border-top:4px solid {'#38bdf8' if 'Riyadh' in reg_name else '#818cf8'}; border-radius:12px; padding:20px; flex:1; min-width:320px;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                <div>
                    <h3 style="margin:0; font-size:17px; color:#fff;">{reg_name}</h3>
                    <span style="font-size:12px; color:#38bdf8; font-weight:600;">Area Manager: {reg_mgr}</span>
                </div>
                <div style="text-align:right;">
                    <span class="badge" style="background:{ach_col}22; color:{ach_col}; border:1px solid {ach_col}55; font-size:12px; font-weight:700;">{r_ach:.1f}% Ach</span>
                    <div style="font-size:11px; margin-top:4px;">YoY: {yoy_badge}</div>
                </div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:10px; margin-top:14px; background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
                <div>
                    <div style="font-size:11px; color:#94a3b8;">CURRENT SALES</div>
                    <div style="font-size:15px; font-weight:700; color:#fff;">{r_sales:,.0f} <span style="font-size:10px;">SAR</span></div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">LY GROSS SALES</div>
                    <div style="font-size:15px; font-weight:700; color:#38bdf8;">{reg_ly_tot:,.0f} <span style="font-size:10px;">SAR</span></div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">SOH UNITS</div>
                    <div style="font-size:15px; font-weight:700; color:#fff;">{r_soh:,.0f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ATV</div>
                    <div style="font-size:13px; font-weight:700; color:#fff;">SAR {r_atv:.1f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">UPT</div>
                    <div style="font-size:13px; font-weight:700; color:#fff;">{r_upt:.2f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ASP</div>
                    <div style="font-size:13px; font-weight:700; color:#f59e0b;">SAR {r_asp:.1f}</div>
                </div>
            </div>
        </div>
        """

        reg_rows = ""
        for idx, r in grp.reset_index(drop=True).iterrows():
            t_str = f"{r['target']:,.0f}" if pd.notna(r['target']) else "-"
            ach_v = r['ach_pct'] if pd.notna(r['ach_pct']) else None
            if ach_v is not None:
                c_c = "#10b981" if ach_v >= 100 else ("#f59e0b" if ach_v >= 80 else "#ef4444")
                ach_cell = f"""
                <div style="display:flex;align-items:center;gap:6px;">
                    <span style="color:{c_c};font-weight:700;min-width:42px;">{ach_v:.1f}%</span>
                    <div style="flex:1;background:#1e293b;border-radius:4px;height:5px;overflow:hidden;">
                        <div style="width:{min(ach_v,100):.1f}%;background:{c_c};height:100%;"></div>
                    </div>
                </div>
                """
            else:
                ach_cell = '<span style="color:#64748b;">-</span>'

            # خلايا الـ LY
            if pd.notna(r['ly_sales']):
                ly_str = f"{r['ly_sales']:,.2f}"
                yoy_v = r['yoy_growth']
                y_col = "#10b981" if yoy_v >= 0 else "#ef4444"
                yoy_cell = f'<span style="color:{y_col}; font-weight:700;">{yoy_v:+.1f}%</span>'
            else:
                ly_str = '<span style="color:#64748b;">New Store</span>'
                yoy_cell = '<span style="color:#64748b;">-</span>'

            reg_rows += f"""
            <tr onclick="openStoreDetails('{r['clean_code']}')" style="cursor:pointer;">
                <td style="color:#64748b;">{idx+1}</td>
                <td style="color:#38bdf8;font-weight:600;">{r['clean_code']}</td>
                <td style="font-weight:600;color:#fff;">{r['full_name']} <span style="font-size:11px;color:#38bdf8;">🔍</span></td>
                <td style="font-weight:700;color:#f8fafc;">{r['sales']:,.2f}</td>
                <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
                <td>{yoy_cell}</td>
                <td style="color:#94a3b8;">{t_str}</td>
                <td style="min-width:120px;">{ach_cell}</td>
                <td style="font-weight:700;color:#fff;">{int(r['soh_units']):,}</td>
                <td>{r['atv']:,.2f}</td>
                <td>{r['upt']:,.2f}</td>
                <td style="color:#38bdf8;font-weight:600;">{r['asp']:,.2f}</td>
            </tr>
            """

        region_tables_html += f"""
        <div class="table-wrap" style="margin-bottom:30px;">
            <div class="table-header">
                <div>
                    <h3 style="color:#38bdf8; font-size:16px;">🏢 {reg_name.upper()}</h3>
                    <span style="color:var(--text-muted);font-size:12px;">Area Manager: <strong style="color:#fff;">{reg_mgr}</strong> | Stores: {len(grp)} Branches</span>
                </div>
            </div>
            <div style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Code</th>
                            <th>Full Store Name</th>
                            <th>Current Sales (SAR)</th>
                            <th>LY Gross Sales (SAR)</th>
                            <th>YoY Growth</th>
                            <th>Target (SAR)</th>
                            <th>% Ach</th>
                            <th>Floor SOH</th>
                            <th>ATV</th>
                            <th>UPT</th>
                            <th>ASP</th>
                        </tr>
                    </thead>
                    <tbody>
                        {reg_rows}
                        <tr style="background:#0c1220; font-weight:700; border-top:2px solid #38bdf8;">
                            <td colspan="3" style="color:#38bdf8; font-size:13px;">TOTAL {reg_name.upper()} ({reg_mgr})</td>
                            <td style="color:#fff; font-size:14px;">{r_sales:,.2f}</td>
                            <td style="color:#38bdf8; font-size:14px;">{reg_ly_tot:,.2f}</td>
                            <td>{yoy_badge}</td>
                            <td style="color:#94a3b8;">{r_target:,.0f}</td>
                            <td style="color:{ach_col};">{r_ach:.1f}%</td>
                            <td style="color:#fff;">{r_soh:,.0f}</td>
                            <td>{r_atv:,.2f}</td>
                            <td>{r_upt:,.2f}</td>
                            <td style="color:#f59e0b;">{r_asp:,.2f}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        """

    # سطر المجموع العام لكامل الشبكة (مع إجمالي LY و LFL Growth)
    net_yoy_col = "#10b981" if network_lfl_growth >= 0 else "#ef4444"
    grand_total_html = f"""
    <div style="background:#131b2e; border:2px solid #2563eb; border-radius:12px; padding:18px 24px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:30px;">
        <div>
            <div style="font-size:13px; color:#38bdf8; font-weight:700; text-transform:uppercase;">Network Grand Total (All Regions)</div>
            <div style="font-size:22px; font-weight:800; color:#fff; margin-top:2px;">{total_sales:,.2f} <span style="font-size:13px; font-weight:400; color:#94a3b8;">SAR</span></div>
        </div>
        <div style="display:flex; gap:20px; flex-wrap:wrap;">
            <div>
                <div style="font-size:11px; color:#94a3b8;">LY GROSS SALES (MMS)</div>
                <div style="font-size:16px; font-weight:700; color:#38bdf8;">{total_ly_sales:,.2f} SAR</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">LFL YoY GROWTH</div>
                <div style="font-size:16px; font-weight:800; color:{net_yoy_col};">{network_lfl_growth:+.1f}%</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">TOTAL TARGET</div>
                <div style="font-size:16px; font-weight:700; color:#fff;">{total_target:,.0f} SAR</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">ACHIEVEMENT</div>
                <div style="font-size:16px; font-weight:700; color:{'#10b981' if overall_ach>=100 else '#f59e0b'};">{overall_ach:.1f}%</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">TOTAL SOH</div>
                <div style="font-size:16px; font-weight:700; color:#fff;">{total_soh_units:,.0f} Pcs</div>
            </div>
        </div>
    </div>
    """

    # 7. جدول الفروع الرئيسي (Store Commercial Matrix) مع عمود مبيعات العام الماضي ونسبة النمو
    store_meta_map = {}
    store_table_rows = ""
    decision_cards_html = ""

    for idx, row in store_summary.iterrows():
        st_code = row['clean_code']
        st_name = row['full_name']
        
        if pd.notna(row['target']):
            target_str = f"{row['target']:,.0f}"
            ach_val = row['ach_pct']
            bar_w = min(ach_val, 100)
            color = "#10b981" if ach_val >= 100 else ("#f59e0b" if ach_val >= 80 else "#ef4444")
            ach_str = f"""
            <div style="display:flex;align-items:center;gap:8px;">
                <span style="color:{color};font-weight:700;min-width:45px;">{ach_val:.1f}%</span>
                <div style="flex:1;background:#1e293b;border-radius:4px;height:6px;overflow:hidden;">
                    <div style="width:{bar_w}%;background:{color};height:100%;"></div>
                </div>
            </div>
            """
        else:
            target_str = "-"
            ach_str = '<span style="color:#64748b;">-</span>'

        if pd.notna(row['ly_sales']):
            ly_str = f"{row['ly_sales']:,.2f}"
            yoy_val = row['yoy_growth']
            yoy_col = "#10b981" if yoy_val >= 0 else "#ef4444"
            yoy_cell = f'<span style="color:{yoy_col}; font-weight:700;">{yoy_val:+.1f}%</span>'
        else:
            ly_str = '<span style="color:#64748b;">New Store</span>'
            yoy_cell = '<span style="color:#64748b;">-</span>'

        diag_badge = f'<span class="badge" style="background:{row["diag_color"]}22; color:{row["diag_color"]}; border:1px solid {row["diag_color"]}66;">{row["diag_title"]}</span>'
        cap_badge = f'<span class="badge" style="background:{row["display_color"]}15; color:{row["display_color"]}; border:1px solid {row["display_color"]}44;">{row["display_status"]}</span>'

        store_meta_map[st_code] = {
            "name": st_name,
            "region": row['region'],
            "manager": row['manager'],
            "sales": f"{row['sales']:,.2f} SAR",
            "ly_sales": ly_str if "New" not in ly_str else "New Store (No LY)",
            "yoy": f"{row['yoy_growth']:+.1f}%" if pd.notna(row['yoy_growth']) else "-",
            "target": f"{target_str} SAR" if target_str != "-" else "No Target",
            "ach": f"{row['ach_pct']:.1f}%" if pd.notna(row['ach_pct']) else "-",
            "share": f"{row['share']:.2f}%",
            "txns": f"{int(row['txns']):,}",
            "atv": f"{row['atv']:,.2f} SAR",
            "upt": f"{row['upt']:.2f}",
            "asp": f"{row['asp']:,.2f} SAR",
            "soh_units": f"{int(row['soh_units']):,} Pcs",
            "capacity_badge": row['display_status'],
            "diag_title": row['diag_title'],
            "problem": row['problem'],
            "action": row['action'],
            "needs": row['needs'],
            "top_cats": row['top_cats_str']
        }

        if "Mismatch" in row['diag_title'] or "Deficit" in row['diag_title'] or "Performer" in row['diag_title']:
            decision_cards_html += f"""
            <div style="background:var(--card); border:1px solid var(--border); border-left:4px solid {row['diag_color']}; border-radius:10px; padding:18px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                    <div>
                        <span style="font-weight:700; color:#fff; font-size:15px;">{st_name} ({st_code})</span>
                        <div style="font-size:11px; color:#94a3b8; margin-top:2px;">{row['region']} | LY Gross: <strong style="color:#38bdf8;">{ly_str}</strong></div>
                    </div>
                    {diag_badge}
                </div>
                
                <div style="background:#090d16; padding:10px 12px; border-radius:6px; margin-bottom:10px; border:1px solid #1e293b; font-size:12px;">
                    <div style="color:#cbd5e1; margin-bottom:4px;"><strong>📦 Core Category Presence:</strong> Leading in {row['top_cats_str']}</div>
                    <div style="color:#f59e0b;"><strong>🎯 What It Needs:</strong> {row['needs']}</div>
                </div>

                <div style="font-size:12px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:8px 12px; border-radius:6px; border:1px solid rgba(56,189,248,0.2); font-weight:600; line-height:1.4;">
                    {row['action']}
                </div>
            </div>
            """

        store_table_rows += f"""
        <tr onclick="openStoreDetails('{st_code}')" style="cursor:pointer;" title="Click to view detailed store category mix & directives">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:600;">{st_code}</td>
            <td style="font-weight:600;color:#fff;">{st_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">🔍</span></td>
            <td style="color:#94a3b8;font-size:12px;">{row['region']}</td>
            <td style="font-weight:700;color:#f8fafc;">{row['sales']:,.2f}</td>
            <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
            <td>{yoy_cell}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:130px;">{ach_str}</td>
            <td style="font-weight:700;color:#fff;">{int(row['soh_units']):,}</td>
            <td>{cap_badge}</td>
            <td>{diag_badge}</td>
            <td>{row['atv']:,.2f}</td>
            <td>{row['upt']:,.2f}</td>
            <td style="color:#38bdf8;font-weight:600;">{row['asp']:,.2f}</td>
        </tr>
        """

    # 8. جدول الأقسام في شاشة Business-Wise
    main_cat_table_rows = ""
    for idx, r in main_cat_summary.iterrows():
        c_name = r['main_category']
        bar_w = min(r['contribution'], 100)
        main_cat_table_rows += f"""
        <tr onclick="filterByMainCategory('{c_name}')" style="cursor:pointer; background:rgba(56,189,248,0.03);" title="Click to filter sub-categories">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="font-weight:800;color:#fff;font-size:14px;">
                🏷️ {c_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">(Click to view items)</span>
            </td>
            <td style="font-weight:700;color:#38bdf8;">{r['sales']:,.2f}</td>
            <td>{int(r['units']):,}</td>
            <td style="min-width:140px;">
                <div style="display:flex;align-items:center;gap:8px;">
                    <span style="color:#f8fafc;font-weight:700;min-width:45px;">{r['contribution']:.1f}%</span>
                    <div style="flex:1;background:#1e293b;border-radius:4px;height:6px;overflow:hidden;">
                        <div style="width:{bar_w}%;background:#38bdf8;height:100%;"></div>
                    </div>
                </div>
            </td>
            <td style="color:#f59e0b;font-weight:700;">{r['asp']:,.2f}</td>
            <td style="color:#cbd5e1;font-weight:500;">{r['leading_store']}</td>
        </tr>
        """

    subsub_json_data = subsub_summary.to_dict(orient='records')
    main_cat_options = '<option value="ALL">-- All Main Categories (Overview) --</option>'
    for c_name in main_cat_summary['main_category']:
        main_cat_options += f'<option value="{c_name}">{c_name}</option>'

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS Executive Commercial & SOH Intelligence Dashboard</title>
    <style>
        :root {{
            --bg: #090d16;
            --card: #131b2e;
            --card-hover: #19233c;
            --border: #1e293b;
            --primary: #38bdf8;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
        }}
        * {{ box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
        body {{ background-color: var(--bg); color: var(--text-main); margin: 0; padding: 24px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 24px; }}
        .header h1 {{ margin: 0; font-size: 24px; font-weight: 700; }}
        .header p {{ margin: 4px 0 0 0; color: var(--text-muted); font-size: 14px; }}
        
        .view-toggle-bar {{ display: flex; background: #0c1220; padding: 4px; border-radius: 10px; border: 1px solid var(--border); margin-bottom: 24px; width: fit-content; gap: 4px; flex-wrap: wrap; }}
        .view-btn {{ background: transparent; border: none; color: var(--text-muted); padding: 10px 22px; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; transition: 0.2s; display: flex; align-items: center; gap: 8px; }}
        .view-btn.active {{ background: #2563eb; color: #fff; box-shadow: 0 4px 12px rgba(37,99,235,0.3); }}
        .view-btn:hover:not(.active) {{ color: #fff; background: rgba(255,255,255,0.05); }}

        .section-title {{ font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); margin-bottom: 14px; display: flex; align-items: center; justify-content:space-between; flex-wrap:wrap; gap:10px; }}
        .insights-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .insight-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; border-left: 4px solid var(--border); }}
        .insight-card.danger {{ border-left-color: #ef4444; }}
        .insight-card.warning {{ border-left-color: #f59e0b; }}
        .insight-card.success {{ border-left-color: #10b981; }}
        .insight-title {{ font-size: 13px; font-weight: 700; margin-bottom: 6px; }}
        .insight-body {{ font-size: 13px; color: var(--text-muted); line-height: 1.5; }}

        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 24px; }}
        .kpi-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }}
        .kpi-title {{ font-size: 11px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; margin-bottom: 6px; }}
        .kpi-value {{ font-size: 22px; font-weight: 700; color: #fff; }}
        .kpi-unit {{ font-size: 12px; color: var(--text-muted); font-weight: 400; }}

        .chart-container {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 22px; margin-bottom: 24px; }}

        .table-wrap {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 24px; }}
        .table-header {{ padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); flex-wrap: wrap; gap: 12px; }}
        .table-header h3 {{ margin: 0; font-size: 15px; font-weight: 700; }}
        .table-search {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #fff; outline: none; width: 260px; font-size: 13px; }}
        .table-select {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #38bdf8; outline: none; font-size: 13px; font-weight: 600; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }}
        th {{ background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-bottom: 1px solid var(--border); }}
        td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); }}
        tr:hover td {{ background: var(--card-hover); }}

        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
        .cards-scroll-container {{ display: flex; gap: 14px; overflow-x: auto; padding-bottom: 12px; margin-bottom: 24px; scroll-behavior: smooth; }}
        .cards-scroll-container::-webkit-scrollbar {{ height: 6px; }}
        .cards-scroll-container::-webkit-scrollbar-track {{ background: #090d16; }}
        .cards-scroll-container::-webkit-scrollbar-thumb {{ background: #1e293b; border-radius: 3px; }}

        .app-modal {{ position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(9, 13, 22, 0.85); backdrop-filter: blur(5px); z-index: 99999; display: none; align-items: center; justify-content: center; }}
        .modal-content {{ background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 92%; max-width: 1000px; max-height: 90vh; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.8); }}
        .modal-header {{ padding: 20px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; background: #0c1220; }}
        .modal-body {{ padding: 24px; overflow-y: auto; }}
        .close-btn {{ background: transparent; border: none; color: #94a3b8; font-size: 26px; cursor: pointer; line-height: 1; }}
        .close-btn:hover {{ color: #fff; }}
    </style>
</head>
<body>

<div id="auth-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:#090d16;z-index:999999;display:flex;align-items:center;justify-content:center;">
  <div style="background:#131b2e;padding:32px;border-radius:12px;box-shadow:0 15px 30px rgba(0,0,0,0.6);text-align:center;width:90%;max-width:380px;border:1px solid #1e293b;">
    <h3 style="color:#fff;margin:0 0 8px 0;font-size:20px;">🔒 MMS Executive Access</h3>
    <p style="color:#94a3b8;font-size:13px;margin:0 0 20px 0;">Enter authorization PIN to unlock dashboard</p>
    <input type="password" id="access-pass" placeholder="Password" style="width:100%;padding:12px;border-radius:6px;border:1px solid #334155;background:#090d16;color:#fff;font-size:16px;text-align:center;outline:none;box-sizing:border-box;margin-bottom:14px;">
    <button onclick="checkAccess()" style="width:100%;padding:12px;border-radius:6px;border:none;background:#2563eb;color:#fff;font-weight:700;font-size:15px;cursor:pointer;">Unlock Dashboard</button>
    <p id="error-msg" style="color:#ef4444;font-size:13px;margin:12px 0 0 0;display:none;">Invalid credentials</p>
  </div>
</div>

<!-- Modal 1: Commercial Deep-Dive Modal -->
<div id="store-modal" class="app-modal">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h2 id="modal-store-name" style="margin:0; font-size:20px; color:#fff;">Store Commercial Intelligence</h2>
        <span id="modal-store-code" style="color:#38bdf8; font-size:12px; font-weight:700;">CODE</span>
      </div>
      <button class="close-btn" onclick="closeModal('store-modal')">&times;</button>
    </div>
    <div class="modal-body">
      <div style="background:#090d16; border:1px solid #1e293b; border-radius:10px; padding:18px; margin-bottom:20px;">
        <div style="font-size:13px; color:#cbd5e1; margin-bottom:8px;">
          <strong style="color:#ef4444;">● Store Situation & Root Cause:</strong> <span id="modal-diag" style="color:#f8fafc;">-</span>
        </div>
        <div style="font-size:13px; color:#f59e0b; margin-bottom:10px;">
          <strong>🎯 What This Store Needs:</strong> <span id="modal-needs" style="color:#fff;">-</span>
        </div>
        <div style="font-size:13px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:10px 14px; border-radius:6px; border:1px solid rgba(56,189,248,0.25);">
          <strong style="color:#38bdf8;">⚡ Commercial Directive:</strong> <span id="modal-directive" style="color:#fff; font-weight:600;">-</span>
        </div>
      </div>

      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Store Commercial Metrics</div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:12px; margin-bottom:24px;">
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">CURRENT SALES</div>
          <div id="modal-sales" style="font-size:18px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">LY GROSS SALES</div>
          <div id="modal-ly" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div>
          <div id="modal-yoy" style="font-size:11px; margin-top:2px;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">TARGET & ACH</div>
          <div id="modal-ach" style="font-size:18px; font-weight:700; color:#10b981;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">FLOOR SOH (UNITS)</div>
          <div id="modal-soh" style="font-size:18px; font-weight:700; color:#38bdf8;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">DISPLAY DENSITY</div>
          <div id="modal-capacity" style="font-size:14px; font-weight:700; color:#cbd5e1;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">AVERAGE TICKET (ATV)</div>
          <div id="modal-atv" style="font-size:18px; font-weight:700; color:#fff;">-</div>
        </div>
      </div>

      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Store Multi-Tier Assortment Performance</div>
      <div style="border:1px solid #1e293b; border-radius:8px; overflow:hidden;">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Main Category</th>
              <th>Sub-Subgroup (Item Class)</th>
              <th>Sales (SAR)</th>
              <th>Units Sold</th>
              <th>Store Share</th>
              <th>ASP (SAR)</th>
            </tr>
          </thead>
          <tbody id="modal-cats-body"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="header">
    <div>
        <h1>MMS Executive Commercial & SOH Intelligence Dashboard</h1>
        <p>Operational Performance, Regional Hierarchy & Like-For-Like (LY) Benchmarks</p>
    </div>
</div>

<div class="section-title"><span>🤖 AI Merchandising Directives (Powered by Claude)</span></div>
<div class="insights-grid">
    <div class="insight-card danger">
        <div class="insight-title" style="color:#ef4444;">● Critical Issues</div>
        <div class="insight-body">{insights.get('critical', '')}</div>
    </div>
    <div class="insight-card warning">
        <div class="insight-title" style="color:#f59e0b;">● Attention Required</div>
        <div class="insight-body">{insights.get('attention', '')}</div>
    </div>
    <div class="insight-card success">
        <div class="insight-title" style="color:#10b981;">● Opportunities</div>
        <div class="insight-body">{insights.get('opportunity', '')}</div>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-title">Current Total Sales</div>
        <div class="kpi-value">{total_sales:,.0f} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">LY Gross Sales (MMS)</div>
        <div class="kpi-value" style="color:#38bdf8;">{total_ly_sales:,.0f} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network LFL YoY Growth</div>
        <div class="kpi-value" style="color:{net_yoy_col};">{network_lfl_growth:+.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Total Target</div>
        <div class="kpi-value">{total_target:,.0f} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Achievement (% Ach)</div>
        <div class="kpi-value" style="color: {'#10b981' if overall_ach >= 100 else ('#f59e0b' if overall_ach >= 80 else '#ef4444')};">{overall_ach:.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ATV</div>
        <div class="kpi-value">SAR {network_atv:.2f}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ASP</div>
        <div class="kpi-value">SAR {network_asp:.2f}</div>
    </div>
</div>

<div class="view-toggle-bar">
    <button class="view-btn active" id="btn-stores" onclick="switchView('stores')">🏢 Store Commercial Matrix</button>
    <button class="view-btn" id="btn-regions" onclick="switchView('regions')">🌍 Region-Wise Performance</button>
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business-Wise Performance (9 Categories)</button>
</div>

<!-- 1. Store Commercial Matrix View -->
<div id="view-stores">
    <div class="section-title"><span>⚡ Critical Action Directives (Category Swaps & Rebalancing Priorities)</span></div>
    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:14px; margin-bottom:24px;">
        {decision_cards_html}
    </div>

    <div class="chart-container">
        <div class="section-title">
            <span>📊 Top Stores: Actual Sales vs Target</span>
            <div style="font-size:12px; font-weight:500; display:flex; gap:16px;">
                <span style="display:flex; align-items:center; gap:6px;"><span style="display:inline-block; width:12px; height:12px; background:#38bdf8; border-radius:2px;"></span> Actual Sales (SAR)</span>
                <span style="display:flex; align-items:center; gap:6px;"><span style="display:inline-block; width:12px; height:12px; background:#334155; border-radius:2px;"></span> Target (SAR)</span>
            </div>
        </div>
        <div style="overflow-x:auto; width:100%;">
            {chart_svg_markup}
        </div>
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3>STORE COMMERCIAL & DISPLAY ASSORTMENT MATRIX</h3>
                <span style="color:var(--text-muted);font-size:12px;">Click any row to open store category performance, display density & tailored directives</span>
            </div>
            <input type="text" id="storeSearch" class="table-search" placeholder="Search full store name, code, or region..." onkeyup="filterStores()">
        </div>
        <div style="overflow-x:auto;">
            <table id="storesTable">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Store Code</th>
                        <th>Full Store Name</th>
                        <th>Region</th>
                        <th>Current Sales (SAR)</th>
                        <th>LY Gross Sales (SAR)</th>
                        <th>YoY Growth</th>
                        <th>Target (SAR)</th>
                        <th>% Ach</th>
                        <th>Floor SOH</th>
                        <th>Display Density</th>
                        <th>Commercial Diagnostic</th>
                        <th>ATV</th>
                        <th>UPT</th>
                        <th>ASP</th>
                    </tr>
                </thead>
                <tbody>
                    {store_table_rows}
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 2. Region-Wise Performance View -->
<div id="view-regions" style="display:none;">
    <div class="section-title"><span>🌍 REGIONAL LEADERSHIP & AREA MANAGER OVERVIEW</span></div>
    <div style="display:flex; flex-wrap:wrap; gap:16px; margin-bottom:24px;">
        {region_kpi_cards}
    </div>

    {grand_total_html}

    {region_tables_html}
</div>

<!-- 3. Business-Wise View -->
<div id="view-business" style="display:none;">
    <div class="section-title">
        <span>🏷️ MUMUSO MAIN PRODUCT CATEGORIES (LEVEL 1 HIERARCHY)</span>
        <span style="font-size:12px; color:var(--text-muted); font-weight:400;">Click any category card to drill down into its sub-subgroups &rarr;</span>
    </div>
    
    <div class="cards-scroll-container">
        {main_cat_cards_html}
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3 id="tableHierarchyTitle">PRODUCT HIERARCHY MATRIX (LEVEL 1: MAIN CATEGORIES)</h3>
                <span style="color:var(--text-muted);font-size:12px;">Select a Main Category from the dropdown or cards to view detailed Sub-Subgroups</span>
            </div>
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <select id="mainCatFilter" class="table-select" onchange="onCategoryFilterChange(this.value)">
                    {main_cat_options}
                </select>
                <input type="text" id="subsubSearch" class="table-search" placeholder="Search product / sub-subgroup..." onkeyup="filterSubSubTable()">
            </div>
        </div>
        <div style="overflow-x:auto;">
            <table id="hierarchyTable">
                <thead id="hierarchyTableHead">
                    <tr>
                        <th>#</th>
                        <th>Main Category</th>
                        <th>Sales Revenue (SAR)</th>
                        <th>Sales Units</th>
                        <th>Network Share (%)</th>
                        <th>ASP (SAR)</th>
                        <th>Leading Store Benchmark</th>
                    </tr>
                </thead>
                <tbody id="hierarchyTableBody">
                    {main_cat_table_rows}
                </tbody>
            </table>
        </div>
    </div>
</div>

<script>
  const PASS = "MMS2026";
  const STORE_DETAILS = {json.dumps(store_cat_details)};
  const STORE_META = {json.dumps(store_meta_map)};
  const SUBSUB_DATA = {json.dumps(subsub_json_data)};
  const MAIN_CAT_HTML = `{main_cat_table_rows}`;

  function openStoreDetails(storeCode) {{
    const meta = STORE_META[storeCode];
    const items = STORE_DETAILS[storeCode] || [];
    if (!meta) return;

    document.getElementById("modal-store-name").innerText = meta.name;
    document.getElementById("modal-store-code").innerText = "CODE: " + storeCode + " | " + meta.region + " (Manager: " + meta.manager + ")";
    document.getElementById("modal-sales").innerText = meta.sales;
    document.getElementById("modal-ly").innerText = meta.ly_sales;
    document.getElementById("modal-yoy").innerHTML = meta.yoy !== "-" ? "YoY: <strong>" + meta.yoy + "</strong>" : "New Location";
    document.getElementById("modal-ach").innerText = meta.ach;
    document.getElementById("modal-soh").innerText = meta.soh_units;
    document.getElementById("modal-capacity").innerText = meta.capacity_badge;
    document.getElementById("modal-atv").innerText = meta.atv;
    document.getElementById("modal-upt").innerText = meta.upt;
    document.getElementById("modal-diag").innerText = meta.problem;
    document.getElementById("modal-needs").innerText = meta.needs;
    document.getElementById("modal-directive").innerText = meta.action;

    let rowsHtml = "";
    items.forEach((c, idx) => {{
      rowsHtml += `
        <tr>
          <td style="color:#64748b;">${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{c.main_category}}</td>
          <td style="font-weight:700; color:#fff;">${{c.sub_subgroup}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{c.sales}}</td>
          <td>${{c.units}}</td>
          <td style="color:#f8fafc; font-weight:600;">${{c.share}}</td>
          <td style="color:#f59e0b;">${{c.asp}}</td>
        </tr>
      `;
    }});

    document.getElementById("modal-cats-body").innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No category data available</td></tr>";
    document.getElementById("store-modal").style.display = "flex";
  }}

  function filterByMainCategory(catName) {{
    document.getElementById("mainCatFilter").value = catName;
    onCategoryFilterChange(catName);
  }}

  function onCategoryFilterChange(catName) {{
    const thead = document.getElementById("hierarchyTableHead");
    const tbody = document.getElementById("hierarchyTableBody");
    const title = document.getElementById("tableHierarchyTitle");

    if (catName === "ALL") {{
      title.innerText = "PRODUCT HIERARCHY MATRIX (LEVEL 1: MAIN CATEGORIES)";
      thead.innerHTML = `
        <tr>
          <th>#</th>
          <th>Main Category</th>
          <th>Sales Revenue (SAR)</th>
          <th>Sales Units</th>
          <th>Network Share (%)</th>
          <th>ASP (SAR)</th>
          <th>Leading Store Benchmark</th>
        </tr>
      `;
      tbody.innerHTML = MAIN_CAT_HTML;
      return;
    }}

    title.innerText = "SUB-SUBGROUP BREAKDOWN: " + catName.toUpperCase();
    thead.innerHTML = `
      <tr>
        <th>#</th>
        <th>Main Category</th>
        <th>Sub-Subgroup (Product Group)</th>
        <th>Sales Revenue (SAR)</th>
        <th>Sales Units</th>
        <th>Contribution (%)</th>
        <th>ASP (SAR)</th>
      </tr>
    `;

    const filtered = SUBSUB_DATA.filter(x => x.main_category === catName);
    let rowsHtml = "";
    filtered.forEach((r, idx) => {{
      const bar_w = Math.min(r.contribution * 3, 100);
      rowsHtml += `
        <tr>
          <td style="color:#64748b;">${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{r.main_category}}</td>
          <td style="font-weight:700; color:#fff;">${{r.sub_subgroup}}</td>
          <td style="font-weight:700; color:#38bdf8;">${{Number(r.sales).toLocaleString(undefined, {{minimumFractionDigits:2, maximumFractionDigits:2}})}}</td>
          <td>${{Number(r.units).toLocaleString()}}</td>
          <td style="min-width:130px;">
            <div style="display:flex;align-items:center;gap:6px;">
              <span style="font-weight:700;color:#fff;min-width:40px;">${{r.contribution.toFixed(2)}}%</span>
              <div style="flex:1;background:#1e293b;border-radius:4px;height:5px;overflow:hidden;">
                <div style="width:${{bar_w}}%;background:#38bdf8;height:100%;"></div>
              </div>
            </div>
          </td>
          <td style="color:#f59e0b;font-weight:700;">${{r.asp.toFixed(2)}}</td>
        </tr>
      `;
    }});

    tbody.innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No sub-subgroups found for this category</td></tr>";
  }}

  function filterSubSubTable() {{
    const q = document.getElementById("subsubSearch").value.toLowerCase();
    const rows = document.querySelectorAll("#hierarchyTableBody tr");
    rows.forEach(r => {{
      r.style.display = r.innerText.toLowerCase().includes(q) ? "" : "none";
    }});
  }}

  function closeModal(modalId) {{
    document.getElementById(modalId).style.display = "none";
  }}

  window.onclick = function(event) {{
    if (event.target.classList.contains('app-modal')) {{
      event.target.style.display = "none";
    }}
  }};

  function switchView(viewName) {{
    const storesView = document.getElementById("view-stores");
    const regionsView = document.getElementById("view-regions");
    const businessView = document.getElementById("view-business");
    const btnStores = document.getElementById("btn-stores");
    const btnRegions = document.getElementById("btn-regions");
    const btnBusiness = document.getElementById("btn-business");

    storesView.style.display = "none";
    regionsView.style.display = "none";
    businessView.style.display = "none";
    btnStores.classList.remove("active");
    btnRegions.classList.remove("active");
    btnBusiness.classList.remove("active");

    if (viewName === 'stores') {{
        storesView.style.display = "block";
        btnStores.classList.add("active");
    }} else if (viewName === 'regions') {{
        regionsView.style.display = "block";
        btnRegions.classList.add("active");
    }} else {{
        businessView.style.display = "block";
        btnBusiness.classList.add("active");
    }}
  }}

  function checkAccess() {{
    const val = document.getElementById("access-pass").value;
    if (val === PASS) {{
      sessionStorage.setItem("mms_auth", "ok");
      document.getElementById("auth-overlay").style.display = "none";
    }} else {{
      document.getElementById("error-msg").style.display = "block";
    }}
  }}

  document.getElementById("access-pass").addEventListener("keypress", function(e) {{
    if (e.key === "Enter") checkAccess();
  }});

  if (sessionStorage.getItem("mms_auth") === "ok") {{
    document.getElementById("auth-overlay").style.display = "none";
  }}

  function filterStores() {{
      const query = document.getElementById("storeSearch").value.toLowerCase();
      const rows = document.querySelectorAll("#storesTable tbody tr");
      rows.forEach(r => {{
          const text = r.innerText.toLowerCase();
          r.style.display = text.includes(query) ? "" : "none";
      }});
  }}
</script>

</body>
</html>
    """

    out_file = os.path.join(REPORTS_DIR, "MMS_Executive_KPI_Dashboard.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"[✓] Dashboard generated successfully: {out_file}")

if __name__ == "__main__":
    process_and_build()