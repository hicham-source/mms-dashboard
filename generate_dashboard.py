import os
import glob
import re
import json
import pandas as pd
import numpy as np
import anthropic

REPORTS_DIR = "./reports"

def identify_files():
    files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~$") and not os.path.basename(f).startswith("Summary_")]
    
    sales_file = None
    soh_file = None
    target_file = None

    for f in sorted(files, key=os.path.getctime, reverse=True):
        fname = os.path.basename(f).lower()
        if "target" in fname or "sep_target" in fname:
            if not target_file:
                target_file = f
            continue
        if "soh" in fname or "stock" in fname:
            if not soh_file:
                soh_file = f
            continue

    for f in sorted(files, key=os.path.getctime, reverse=True):
        if f == target_file or f == soh_file:
            continue
        try:
            xl = pd.ExcelFile(f)
            for s in xl.sheet_names:
                sample_df = pd.read_excel(f, sheet_name=s, nrows=3)
                cols_str = " ".join([str(c).lower() for c in sample_df.columns])
                if "item stock report" in cols_str or "avail_stock" in cols_str or "current_stock" in cols_str:
                    if not soh_file:
                        soh_file = f
                    break
                sample_df2 = pd.read_excel(f, sheet_name=s, skiprows=1, nrows=3)
                cols_str2 = " ".join([str(c).lower() for c in sample_df2.columns])
                if "receipt number" in cols_str2 or "actual sales amount" in cols_str2:
                    if not sales_file:
                        sales_file = f
                    break
        except Exception:
            continue

    if not sales_file:
        candidates = [f for f in files if f != target_file and f != soh_file]
        if candidates:
            sales_file = max(candidates, key=os.path.getctime)
        else:
            raise FileNotFoundError("Sales report file not found in ./reports")

    return sales_file, soh_file, target_file

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
        print("[!] No SOH file provided or found.")
        return {}
    
    print(f"[*] Processing SOH file: {soh_path}")
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

        if not code_col or not stock_col:
            return {}

        df_soh = df_soh[df_soh[code_col].notna()].copy()
        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)
        
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
        return grouped.set_index('clean_code').to_dict(orient='index')
    except Exception as e:
        print(f"[!] Error processing SOH file: {e}")
        return {}

def generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp, total_soh_units, total_ideal_stock):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("[!] ANTHROPIC_API_KEY not found; using fallback executive insights.")
        return {
            "critical": f"Network holds {total_soh_units:,.0f} units vs an ideal target of {total_ideal_stock:,.0f} units. Overstocked branches are freezing cash while under-stocked locations risk running out of key sellers.",
            "attention": "Trigger direct Inter-Store Transfers (IST) from locations with WOS > 30 weeks directly to top footfall stores before placing new supplier orders.",
            "opportunity": "Re-align branch gondola capacities toward leading categories with sell-through > 20% to maximize cash return per square meter."
        }

    top_stores = store_summary.head(3)[['Organization Name', 'sales', 'ach_pct', 'wos', 'stock_gap_units']].to_dict(orient="records")
    bottom_stores = store_summary.tail(3)[['Organization Name', 'sales', 'ach_pct', 'wos', 'stock_gap_units']].to_dict(orient="records")
    top_categories = cat_summary.head(5)[['Category Name', 'sales', 'contribution']].to_dict(orient="records")

    prompt = f"""
    You are a Senior Retail Operations & Merchandising Director reviewing store sales vs Stock on Hand (SOH).
    - Total Sales: {total_sales:,.0f} SAR
    - Total Target: {total_target:,.0f} SAR
    - Network Achievement: {overall_ach:.1f}%
    - Total SOH: {total_soh_units:,.0f} units
    - Ideal Target SOH (5-week cover standard): {total_ideal_stock:,.0f} units
    - Top Performing Stores: {top_stores}
    - Low Performing Stores: {bottom_stores}
    - Key Categories: {top_categories}

    Generate 3 commercial, action-oriented directives (1 clear sentence each):
    1. Critical Issues: address stock imbalance, stockouts, or capital tie-up.
    2. Attention Required: specify Inter-Store Transfers (IST) and merchandising reallocation.
    3. Opportunities: commercial moves to lift basket value and sell-through.

    Respond ONLY in valid JSON:
    {{
        "critical": "...",
        "attention": "...",
        "opportunity": "..."
    }}
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
        insights = json.loads(content)
        print("[✓] Claude AI insights generated successfully.")
        return insights
    except Exception as e:
        print(f"[!] Claude API error: {e}")
        return {
            "critical": f"Stock depth ({total_soh_units:,.0f} units) heavily exceeds operational demand ({total_ideal_stock:,.0f} units). Fast-track stock rebalancing.",
            "attention": "Execute immediate Inter-Store Transfers (IST) from regional low-velocity branches to flagship locations.",
            "opportunity": "Bundle sluggish inventory with high-velocity toys and accessories to stimulate cash recovery."
        }

def build_svg_bar_chart(chart_stores):
    svg_w, svg_h = 900, 320
    pad_left, pad_right, pad_top, pad_bottom = 60, 20, 30, 60
    plot_w = svg_w - pad_left - pad_right
    plot_h = svg_h - pad_top - pad_bottom

    max_val = max(chart_stores['sales'].max(), chart_stores['target'].fillna(0).max()) * 1.15
    if max_val == 0:
        max_val = 1

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

        name = str(r['Organization Name']).replace("MMS ", "")

        bars_svg += f"""
        <rect x="{s_x:.1f}" y="{s_y:.1f}" width="{bar_w:.1f}" height="{s_h:.1f}" rx="3" fill="#38bdf8">
            <title>{r['Organization Name']} Sales: {s_val:,.2f} SAR</title>
        </rect>
        <rect x="{t_x:.1f}" y="{t_y:.1f}" width="{bar_w:.1f}" height="{t_h:.1f}" rx="3" fill="#334155">
            <title>{r['Organization Name']} Target: {t_val:,.2f} SAR</title>
        </rect>
        <text x="{slot_center:.1f}" y="{svg_h - 25}" fill="#cbd5e1" font-size="11" font-weight="600" text-anchor="middle">{name}</text>
        """

    return f"""<svg viewBox="0 0 {svg_w} {svg_h}" style="width:100%; height:auto; display:block;">{grid_lines}{bars_svg}</svg>"""

def process_and_build():
    sales_file, soh_file, target_file = identify_files()
    print(f"[*] Sales Report: {sales_file}")
    print(f"[*] SOH File:     {soh_file}")
    print(f"[*] Target File:  {target_file}")

    targets_map = load_september_targets(target_file)
    soh_map = load_soh_data(soh_file)

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

    cat_col = 'Category Name' if 'Category Name' in df_clean.columns else ('product_category' if 'product_category' in df_clean.columns else ('category' if 'category' in df_clean.columns else None))
    if cat_col:
        df_clean[cat_col] = df_clean[cat_col].fillna("Other").astype(str).str.strip()
    else:
        df_clean['Category Name'] = "General"
        cat_col = 'Category Name'

    # 1. إجماليات المتاجر
    store_summary = df_clean.groupby(['Organization Code', 'Organization Name']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index()

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

    # 2. ربط المخزون وحساب المقارنة التجارية (Actual SOH vs Ideal Stock)
    def match_soh(row):
        code_str = str(row['Organization Code']).strip()
        m = re.search(r'\b[A-Za-z0-9]{3,8}\b', code_str)
        c_code = m.group(0).upper() if m else code_str.upper()
        if c_code in soh_map:
            return soh_map[c_code]
        for k, v in soh_map.items():
            if k in c_code or c_code in k:
                return v
        return {"soh_units": 0, "soh_val": 0}

    soh_matched = store_summary.apply(match_soh, axis=1)
    store_summary['soh_units'] = [x['soh_units'] for x in soh_matched]
    store_summary['soh_val'] = [x['soh_val'] for x in soh_matched]
    total_soh_units = store_summary['soh_units'].sum()
    total_soh_val = store_summary['soh_val'].sum()

    # معدل البيع الأسبوعي ومتوسط المبيعات اليومية
    weekly_units = store_summary['units'] / 4.3
    daily_units = store_summary['units'] / 30.0
    daily_sales_val = store_summary['sales'] / 30.0

    store_summary['wos'] = (store_summary['soh_units'] / weekly_units.replace(0, np.nan)).fillna(0).round(1)

    # حساب المخزون المثالي (Ideal Stock = 5 Weeks of Sales Run-Rate)
    # المعيار التجاري الصحيح في متاجر التجزئة هو تغطية 5 أسابيع
    store_summary['ideal_stock_units'] = (weekly_units * 5.0).round(0)
    store_summary['ideal_stock_val'] = (store_summary['ideal_stock_units'] * store_summary['asp']).round(0)
    
    # فجوة المخزون (Stock Gap = Actual SOH - Ideal Stock)
    store_summary['stock_gap_units'] = (store_summary['soh_units'] - store_summary['ideal_stock_units']).round(0)
    store_summary['stock_gap_val'] = (store_summary['stock_gap_units'] * store_summary['asp']).round(0)

    # نسبة استهلاك المخزون (Sell-Through %)
    store_summary['sell_through'] = ((store_summary['units'] / (store_summary['units'] + store_summary['soh_units']).replace(0, np.nan)) * 100).fillna(0).round(1)

    total_ideal_stock = store_summary['ideal_stock_units'].sum()

    # 3. مطابقة الأهداف
    def match_target(row):
        code_str = str(row['Organization Code']).strip().upper()
        name_str = str(row['Organization Name']).strip().upper()
        for k, v in targets_map.items():
            if str(k).upper() in code_str or str(k).upper() in name_str:
                return v
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

    # 4. محرك التشخيص التجاري والقرار التشغيلي الحاسم (Commercial Decision Engine)
    def commercial_decision(row):
        ach = row['ach_pct'] if pd.notna(row['ach_pct']) else 0
        wos = row['wos']
        gap_u = row['stock_gap_units']
        gap_v = row['stock_gap_val']
        st = row['sell_through']
        soh = row['soh_units']
        ideal = row['ideal_stock_units']

        if soh == 0:
            return {
                "status": "No SOH Synced",
                "color": "#64748b",
                "diag": "Inventory figures missing from export.",
                "directive": "Verify warehouse stock mapping for this branch code.",
                "action_type": "DATA SYNC"
            }

        if wos < 3.0 and ach >= 80:
            return {
                "status": "Stockout Risk",
                "color": "#ef4444",
                "diag": f"Critically lean stock ({wos} Wks cover vs 5 Wks benchmark). Store run-rate is burning stock faster than intake.",
                "directive": f"⚡ DIRECTIVE: Emergency Warehouse Dispatch of +{int(abs(gap_u)):,} Units (~{abs(gap_v):,.0f} SAR) to protect target velocity.",
                "action_type": "EMERGENCY REPLENISH"
            }
        elif wos > 15.0 and ach < 70:
            return {
                "status": "Heavy Overstock",
                "color": "#f59e0b",
                "diag": f"Massive stock overload ({wos} Wks cover). Excess of +{int(gap_u):,} Units tying up ~{gap_v:,.0f} SAR in idle capital.",
                "directive": f"⚡ DIRECTIVE: Immediate Inter-Store Transfer (IST) of {int(gap_u * 0.6):,} Units to top performers + launch cashier multi-buy bundles.",
                "action_type": "TRANSFER OUT (IST)"
            }
        elif wos > 10.0 and ach >= 80:
            return {
                "status": "High Stock Cover",
                "color": "#38bdf8",
                "diag": f"Healthy sales with substantial buffer ({wos} Wks). Holds surplus of +{int(gap_u):,} Units.",
                "directive": f"⚡ DIRECTIVE: Freeze fresh warehouse purchase orders; fulfill demand from existing backroom capacity.",
                "action_type": "FREEZE ORDERS"
            }
        elif ach < 70 and 3.0 <= wos <= 10.0:
            return {
                "status": "Conversion Bottleneck",
                "color": "#eab308",
                "diag": f"Stock level is adequate ({wos} Wks), but commercial conversion is lagging ({ach:.1f}% Ach, {st}% Sell-through).",
                "directive": f"⚡ DIRECTIVE: Do NOT inject more inventory. Focus on sales floor coaching, gondola re-merchandising, and cashier ATV upselling.",
                "action_type": "FLOOR COACHING"
            }
        else:
            return {
                "status": "Balanced Flow",
                "color": "#10b981",
                "diag": f"Optimal inventory alignment ({wos} Wks cover, {st}% Sell-through). Stock closely matches sales velocity.",
                "directive": "⚡ DIRECTIVE: Maintain current replenishment cadence and display standards.",
                "action_type": "MAINTAIN"
            }

    decisions = store_summary.apply(commercial_decision, axis=1)
    store_summary['status'] = [d['status'] for d in decisions]
    store_summary['color'] = [d['color'] for d in decisions]
    store_summary['diag'] = [d['diag'] for d in decisions]
    store_summary['directive'] = [d['directive'] for d in decisions]
    store_summary['action_type'] = [d['action_type'] for d in decisions]

    # 5. إجماليات الأصناف
    cat_summary = df_clean.groupby(cat_col).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().rename(columns={cat_col: 'Category Name'})
    cat_summary['contribution'] = ((cat_summary['sales'] / total_sales) * 100).round(2)
    cat_summary['asp'] = (cat_summary['sales'] / cat_summary['units'].replace(0, np.nan)).fillna(0).round(2)

    cat_store_breakdown = {}
    grouped_cat_store = df_clean.groupby([cat_col, 'Organization Name', 'Organization Code']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index()

    for c_name, grp in grouped_cat_store.groupby(cat_col):
        c_total = grp['sales'].sum()
        grp_sorted = grp.sort_values(by='sales', ascending=False)
        st_list = []
        for _, s_row in grp_sorted.iterrows():
            st_sales = s_row['sales']
            st_units = s_row['units']
            st_share = (st_sales / c_total * 100) if c_total > 0 else 0
            st_asp = (st_sales / st_units) if st_units > 0 else 0
            st_list.append({
                "store": s_row['Organization Name'],
                "code": str(s_row['Organization Code']),
                "sales": f"{st_sales:,.2f}",
                "units": f"{int(st_units):,}",
                "share": f"{st_share:.1f}%",
                "asp": f"{st_asp:,.2f}"
            })
        cat_store_breakdown[c_name] = st_list

    top_store_per_cat = {}
    for c_name, st_list in cat_store_breakdown.items():
        if st_list:
            top_store_per_cat[c_name] = f"{st_list[0]['store']} ({st_list[0]['sales']} SAR - {st_list[0]['share']})"
    cat_summary['leading_store'] = cat_summary['Category Name'].map(top_store_per_cat).fillna("-")
    cat_summary = cat_summary.sort_values(by='sales', ascending=False).reset_index(drop=True)

    # تفاصيل المتجر لكل تصنيف
    store_category_details = {}
    grouped_store_cat = df_clean.groupby(['Organization Code', cat_col]).agg(
        cat_sales=('Actual Sales Amount', 'sum'),
        cat_units=('Sales Quantity', 'sum')
    ).reset_index()

    for code, grp in grouped_store_cat.groupby('Organization Code'):
        st_total = grp['cat_sales'].sum()
        grp_sorted = grp.sort_values(by='cat_sales', ascending=False)
        cats_list = []
        for _, c_row in grp_sorted.iterrows():
            c_sales = c_row['cat_sales']
            c_units = c_row['cat_units']
            share_st = (c_sales / st_total * 100) if st_total > 0 else 0
            asp_st = (c_sales / c_units) if c_units > 0 else 0
            cats_list.append({
                "category": c_row[cat_col],
                "sales": f"{c_sales:,.2f}",
                "units": f"{int(c_units):,}",
                "share": f"{share_st:.1f}%",
                "asp": f"{asp_st:,.2f}"
            })
        store_category_details[str(code)] = cats_list

    top_cat_per_store = {}
    for code, cats in store_category_details.items():
        if cats:
            top_cat_per_store[code] = f"{cats[0]['category']} ({cats[0]['share']})"
    store_summary['top_category'] = store_summary['Organization Code'].astype(str).map(top_cat_per_store).fillna("-")

    insights = generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp, total_soh_units, total_ideal_stock)

    chart_stores = store_summary.head(8)
    chart_svg_markup = build_svg_bar_chart(chart_stores)

    # بطاقات الأصناف
    colors = ['#38bdf8', '#818cf8', '#a855f7', '#ec4899', '#f59e0b', '#10b981', '#06b6d4', '#e11d48', '#84cc16']
    cat_cards_html = ""
    for idx, r in cat_summary.iterrows():
        c_color = colors[idx % len(colors)]
        cat_cards_html += f"""
        <div onclick="openCategoryStores('{r['Category Name']}')" style="background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; min-width:210px; max-width:250px; flex:0 0 auto; cursor:pointer;" title="Click to see all stores selling {r['Category Name']}">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-size:13px; font-weight:700; color:#f8fafc; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{r['Category Name']}</span>
                <span style="font-size:12px; font-weight:700; color:{c_color};">{r['contribution']:.1f}%</span>
            </div>
            <div style="font-size:18px; font-weight:700; color:#fff; margin-bottom:6px;">{r['sales']:,.0f} <span style="font-size:11px; color:#94a3b8;">SAR</span></div>
            <div style="background:#090d16; border-radius:4px; height:6px; overflow:hidden;">
                <div style="background:{c_color}; width:{min(r['contribution'], 100):.1f}%; height:100%;"></div>
            </div>
            <div style="display:flex; justify-content:space-between; margin-top:8px; font-size:11px; color:#94a3b8;">
                <span>Units: {int(r['units']):,}</span>
                <span>ASP: {r['asp']:,.1f} SAR</span>
            </div>
        </div>
        """

    # بناء بيانات المتاجر مع تفاصيل المقارنة الدقيقة
    store_meta_map = {}
    store_table_rows = ""
    decision_cards_html = ""

    for idx, row in store_summary.iterrows():
        st_code = str(row['Organization Code'])
        st_name = str(row['Organization Name'])
        
        if pd.notna(row['target']):
            target_str = f"{row['target']:,.0f}"
            ach_val = row['ach_pct']
            bar_w = min(ach_val, 100)
            if ach_val >= 100:
                color = "#10b981"
            elif ach_val >= 80:
                color = "#f59e0b"
            else:
                color = "#ef4444"

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

        diag_badge = f'<span class="badge" style="background:{row["color"]}22; color:{row["color"]}; border:1px solid {row["color"]}66;">{row["status"]}</span>'

        gap_u = row['stock_gap_units']
        gap_sign = "+" if gap_u > 0 else ""
        gap_color = "#f59e0b" if gap_u > 0 else ("#ef4444" if gap_u < 0 else "#10b981")

        store_meta_map[st_code] = {
            "name": st_name,
            "sales": f"{row['sales']:,.2f} SAR",
            "target": f"{target_str} SAR" if target_str != "-" else "No Target",
            "ach": f"{row['ach_pct']:.1f}%" if pd.notna(row['ach_pct']) else "-",
            "share": f"{row['share']:.2f}%",
            "txns": f"{int(row['txns']):,}",
            "atv": f"{row['atv']:,.2f} SAR",
            "upt": f"{row['upt']:.2f}",
            "asp": f"{row['asp']:,.2f} SAR",
            "soh_units": f"{int(row['soh_units']):,} Pcs",
            "ideal_units": f"{int(row['ideal_stock_units']):,} Pcs",
            "gap_units": f"{gap_sign}{int(gap_u):,} Pcs",
            "gap_val": f"{gap_sign}{row['stock_gap_val']:,.0f} SAR",
            "wos": f"{row['wos']} Wks",
            "sell_through": f"{row['sell_through']:.1f}%",
            "diag": row['diag'],
            "directive": row['directive']
        }

        # كروت التوجيهات التجارية الاستراتيجية (Executive Commercial Directives)
        if row['status'] in ["Stockout Risk", "Heavy Overstock", "Conversion Bottleneck"]:
            decision_cards_html += f"""
            <div style="background:var(--card); border:1px solid var(--border); border-left:4px solid {row['color']}; border-radius:10px; padding:18px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <span style="font-weight:700; color:#fff; font-size:15px;">{st_name} ({st_code})</span>
                    {diag_badge}
                </div>
                
                <!-- مقارنة المخزون الفعلي بالمستهدف -->
                <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:8px; background:#090d16; padding:10px; border-radius:6px; margin-bottom:10px; border:1px solid #1e293b;">
                    <div>
                        <div style="font-size:10px; color:#94a3b8;">CURRENT SOH</div>
                        <div style="font-size:13px; font-weight:700; color:#fff;">{int(row['soh_units']):,}</div>
                    </div>
                    <div>
                        <div style="font-size:10px; color:#94a3b8;">IDEAL (5 WKS)</div>
                        <div style="font-size:13px; font-weight:700; color:#38bdf8;">{int(row['ideal_stock_units']):,}</div>
                    </div>
                    <div>
                        <div style="font-size:10px; color:#94a3b8;">STOCK GAP</div>
                        <div style="font-size:13px; font-weight:700; color:{gap_color};">{gap_sign}{int(gap_u):,}</div>
                    </div>
                </div>

                <div style="font-size:12px; color:#cbd5e1; margin-bottom:8px; line-height:1.4;">
                    <strong>🔍 Commercial Issue:</strong> {row['diag']}
                </div>
                <div style="font-size:12px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:8px 12px; border-radius:6px; border:1px solid rgba(56,189,248,0.2); font-weight:600;">
                    {row['directive']}
                </div>
            </div>
            """

        store_table_rows += f"""
        <tr onclick="openStoreDetails('{st_code}')" style="cursor:pointer;" title="Click to view full stock vs target comparison & action plan">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:600;">{st_code}</td>
            <td style="font-weight:600;color:#fff;">
                {st_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">🔍</span>
            </td>
            <td style="font-weight:700;color:#f8fafc;">{row['sales']:,.2f}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:130px;">{ach_str}</td>
            <td style="font-weight:700;color:#fff;">{int(row['soh_units']):,}</td>
            <td style="font-weight:600;color:#38bdf8;">{int(row['ideal_stock_units']):,}</td>
            <td style="font-weight:700;color:{gap_color};">{gap_sign}{int(gap_u):,}</td>
            <td style="font-weight:700;color:{row['color']};">{row['wos']} Wks</td>
            <td>{diag_badge}</td>
            <td style="color:#38bdf8;font-weight:600;">{row['asp']:,.2f}</td>
        </tr>
        """

    cat_table_rows = ""
    for idx, r in cat_summary.iterrows():
        c_name = r['Category Name']
        bar_w = min(r['contribution'], 100)
        cat_table_rows += f"""
        <tr onclick="openCategoryStores('{c_name}')" style="cursor:pointer;" title="Click to see each store's contribution in {c_name}">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="font-weight:700;color:#fff;font-size:14px;">
                {c_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">🔍</span>
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

    store_options_html = '<option value="ALL">-- Select Store to filter categories --</option>'
    for _, s in store_summary.iterrows():
        store_options_html += f'<option value="{s["Organization Code"]}">{s["Organization Name"]} ({s["Organization Code"]})</option>'

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
        
        .view-toggle-bar {{ display: flex; background: #0c1220; padding: 4px; border-radius: 10px; border: 1px solid var(--border); margin-bottom: 24px; width: fit-content; gap: 4px; }}
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
        .table-search {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #fff; outline: none; width: 240px; font-size: 13px; }}
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

<!-- Modal 1: Commercial Deep-Dive & Decision Engine Modal -->
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
      
      <!-- الصندوق التنفيذي للقرار والتشخيص -->
      <div style="background:#090d16; border:1px solid #1e293b; border-radius:10px; padding:18px; margin-bottom:20px;">
        <div style="font-size:13px; color:#cbd5e1; margin-bottom:8px;">
          <strong style="color:#ef4444;">● Root-Cause Diagnostic:</strong> <span id="modal-diag" style="color:#f8fafc;">-</span>
        </div>
        <div style="font-size:13px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:10px 14px; border-radius:6px; border:1px solid rgba(56,189,248,0.25);">
          <strong style="color:#38bdf8;">⚡ Commercial Directive:</strong> <span id="modal-directive" style="color:#fff; font-weight:600;">-</span>
        </div>
      </div>

      <!-- شبكة مقارنة المخزون الفعلي بالمستهدف والفجوة التجارية -->
      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Inventory Benchmark vs Sales Demand</div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:12px; margin-bottom:24px;">
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">ACTUAL SOH</div>
          <div id="modal-soh" style="font-size:18px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">IDEAL TARGET (5 WKS)</div>
          <div id="modal-ideal" style="font-size:18px; font-weight:700; color:#38bdf8;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">STOCK GAP (SURPLUS/DEFICIT)</div>
          <div id="modal-gap-u" style="font-size:18px; font-weight:700; color:#f59e0b;">-</div>
          <div id="modal-gap-v" style="font-size:11px; color:#94a3b8; margin-top:2px;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">WEEKS OF SUPPLY (WOS)</div>
          <div id="modal-wos" style="font-size:18px; font-weight:700; color:#f59e0b;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">SELL-THROUGH RATE</div>
          <div id="modal-st" style="font-size:18px; font-weight:700; color:#10b981;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">STORE ASP</div>
          <div id="modal-asp" style="font-size:18px; font-weight:700; color:#fff;">-</div>
        </div>
      </div>

      <!-- جدول تفاصيل الأصناف للمتجر -->
      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Store Category Breakdown</div>
      <div style="border:1px solid #1e293b; border-radius:8px; overflow:hidden;">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Category</th>
              <th>Sales (SAR)</th>
              <th>Units Sold</th>
              <th>Contribution in Store</th>
              <th>ASP (SAR)</th>
            </tr>
          </thead>
          <tbody id="modal-cats-body"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<!-- Modal 2: Category Modal -->
<div id="cat-modal" class="app-modal">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h2 id="modal-cat-name" style="margin:0; font-size:18px; color:#38bdf8;">Category Breakdown Across Stores</h2>
        <span style="color:#94a3b8; font-size:12px;">Store-by-Store Sales Volume & Contribution</span>
      </div>
      <button class="close-btn" onclick="closeModal('cat-modal')">&times;</button>
    </div>
    <div class="modal-body">
      <div style="border:1px solid #1e293b; border-radius:8px; overflow:hidden;">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Store Code</th>
              <th>Store Name</th>
              <th>Sales (SAR)</th>
              <th>Units Sold</th>
              <th>Store Share in Category</th>
              <th>ASP (SAR)</th>
            </tr>
          </thead>
          <tbody id="modal-cat-stores-body"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<div class="header">
    <div>
        <h1>MMS Executive Commercial & SOH Intelligence Dashboard</h1>
        <p>Target Alignment, Actual vs Ideal Stock Gap & Actionable Directives</p>
    </div>
</div>

<div class="section-title"><span>🤖 AI Executive Directives (Powered by Claude)</span></div>
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
        <div class="kpi-title">Total Sales</div>
        <div class="kpi-value">{total_sales:,.0f} <span class="kpi-unit">SAR</span></div>
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
        <div class="kpi-title">Total Network SOH</div>
        <div class="kpi-value" style="color:#fff;">{total_soh_units:,.0f} <span class="kpi-unit">Units</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Ideal Network Target SOH</div>
        <div class="kpi-value" style="color:#38bdf8;">{total_ideal_stock:,.0f} <span class="kpi-unit">Units</span></div>
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
    <button class="view-btn active" id="btn-stores" onclick="switchView('stores')">🏢 Store Commercial Matrix & Stock Gap</button>
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business-Wise Performance ({len(cat_summary)} Categories)</button>
</div>

<!-- 1. Store Commercial Matrix View -->
<div id="view-stores">
    
    <div class="section-title"><span>⚡ Critical Action Directives (Stockout & Overstock Priorities)</span></div>
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
                <h3>STORE COMMERCIAL & INVENTORY GAP MATRIX</h3>
                <span style="color:var(--text-muted);font-size:12px;">Click any row to open in-depth SOH comparison, Ideal Target gap & commercial directives</span>
            </div>
            <input type="text" id="storeSearch" class="table-search" placeholder="Search store name or code..." onkeyup="filterStores()">
        </div>
        <div style="overflow-x:auto;">
            <table id="storesTable">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Store Code</th>
                        <th>Store Name</th>
                        <th>Sales (SAR)</th>
                        <th>Target (SAR)</th>
                        <th>% Ach</th>
                        <th>Actual SOH</th>
                        <th>Ideal SOH (5W)</th>
                        <th>Stock Gap</th>
                        <th>WOS</th>
                        <th>Status</th>
                        <th>ASP (SAR)</th>
                    </tr>
                </thead>
                <tbody>
                    {store_table_rows}
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 2. Business-Wise View -->
<div id="view-business" style="display:none;">
    <div class="section-title">
        <span>📦 ALL CATEGORIES CONTRIBUTION MIX ({len(cat_summary)} CATEGORIES)</span>
        <span style="font-size:12px; color:var(--text-muted); font-weight:400;">Click any category to see store-by-store sales details &rarr;</span>
    </div>
    
    <div class="cards-scroll-container">
        {cat_cards_html}
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3>CATEGORY PERFORMANCE MATRIX</h3>
                <span style="color:var(--text-muted);font-size:12px;">Click any category row to see all stores' sales, or filter by specific store:</span>
            </div>
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <select id="storeCategoryFilter" class="table-select" onchange="filterCategoryByStore(this.value)">
                    {store_options_html}
                </select>
                <input type="text" id="catSearch" class="table-search" placeholder="Search any category..." onkeyup="filterCategories()">
            </div>
        </div>
        <div style="overflow-x:auto;">
            <table id="categoriesTable">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Category Name</th>
                        <th>Sales Revenue (SAR)</th>
                        <th>Sales Units</th>
                        <th>Contribution (%)</th>
                        <th>ASP (SAR)</th>
                        <th>Leading Store Benchmark</th>
                    </tr>
                </thead>
                <tbody id="categoriesTableBody">
                    {cat_table_rows}
                </tbody>
            </table>
        </div>
    </div>
</div>

<script>
  const PASS = "MMS2026";
  const STORE_DETAILS = {json.dumps(store_category_details)};
  const STORE_META = {json.dumps(store_meta_map)};
  const CAT_STORE_BREAKDOWN = {json.dumps(cat_store_breakdown)};
  const DEFAULT_CAT_TABLE_HTML = `{cat_table_rows}`;

  function openStoreDetails(storeCode) {{
    const meta = STORE_META[storeCode];
    const cats = STORE_DETAILS[storeCode] || [];
    if (!meta) return;

    document.getElementById("modal-store-name").innerText = meta.name;
    document.getElementById("modal-store-code").innerText = "BRANCH CODE: " + storeCode + " | TARGET: " + meta.target;
    document.getElementById("modal-soh").innerText = meta.soh_units;
    document.getElementById("modal-ideal").innerText = meta.ideal_units;
    document.getElementById("modal-gap-u").innerText = meta.gap_units;
    document.getElementById("modal-gap-v").innerText = "Value Gap: " + meta.gap_val;
    document.getElementById("modal-wos").innerText = meta.wos;
    document.getElementById("modal-st").innerText = meta.sell_through;
    document.getElementById("modal-asp").innerText = meta.asp;
    document.getElementById("modal-diag").innerText = meta.diag;
    document.getElementById("modal-directive").innerText = meta.directive;

    let rowsHtml = "";
    cats.forEach((c, idx) => {{
      rowsHtml += `
        <tr>
          <td style="color:#64748b;">${{idx+1}}</td>
          <td style="font-weight:700; color:#fff;">${{c.category}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{c.sales}}</td>
          <td>${{c.units}}</td>
          <td style="color:#f8fafc; font-weight:600;">${{c.share}}</td>
          <td style="color:#f59e0b;">${{c.asp}}</td>
        </tr>
      `;
    }});

    document.getElementById("modal-cats-body").innerHTML = rowsHtml || "<tr><td colspan='6' style='text-align:center;'>No category data available</td></tr>";
    document.getElementById("store-modal").style.display = "flex";
  }}

  function openCategoryStores(catName) {{
    const stores = CAT_STORE_BREAKDOWN[catName] || [];
    document.getElementById("modal-cat-name").innerText = catName + " - Store Breakdown";

    let rowsHtml = "";
    stores.forEach((s, idx) => {{
      rowsHtml += `
        <tr>
          <td style="color:#64748b;">${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{s.code}}</td>
          <td style="font-weight:700; color:#fff;">${{s.store}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{s.sales}}</td>
          <td>${{s.units}}</td>
          <td style="color:#10b981; font-weight:700;">${{s.share}}</td>
          <td style="color:#f59e0b;">${{s.asp}}</td>
        </tr>
      `;
    }});

    document.getElementById("modal-cat-stores-body").innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No store data available</td></tr>";
    document.getElementById("cat-modal").style.display = "flex";
  }}

  function filterCategoryByStore(selectedCode) {{
    const tbody = document.getElementById("categoriesTableBody");
    if (selectedCode === "ALL") {{
      tbody.innerHTML = DEFAULT_CAT_TABLE_HTML;
      return;
    }}

    const storeCats = STORE_DETAILS[selectedCode] || [];
    const storeName = (STORE_META[selectedCode] && STORE_META[selectedCode].name) || selectedCode;

    let rowsHtml = "";
    storeCats.forEach((c, idx) => {{
      rowsHtml += `
        <tr onclick="openCategoryStores('${{c.category}}')" style="cursor:pointer;">
          <td style="color:#64748b;font-weight:600;">${{idx+1}}</td>
          <td style="font-weight:700;color:#fff;font-size:14px;">${{c.category}} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">🔍</span></td>
          <td style="font-weight:700;color:#38bdf8;">${{c.sales}}</td>
          <td>${{c.units}}</td>
          <td style="color:#f8fafc;font-weight:700;">${{c.share}}</td>
          <td style="color:#f59e0b;font-weight:700;">${{c.asp}}</td>
          <td style="color:#38bdf8;font-weight:500;">${{storeName}} (Filtered)</td>
        </tr>
      `;
    }});

    tbody.innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No categories found for this store.</td></tr>";
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
    const businessView = document.getElementById("view-business");
    const btnStores = document.getElementById("btn-stores");
    const btnBusiness = document.getElementById("btn-business");

    if (viewName === 'stores') {{
        storesView.style.display = "block";
        businessView.style.display = "none";
        btnStores.classList.add("active");
        btnBusiness.classList.remove("active");
    }} else {{
        storesView.style.display = "none";
        businessView.style.display = "block";
        btnStores.classList.remove("active");
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

  function filterCategories() {{
      const query = document.getElementById("catSearch").value.toLowerCase();
      const rows = document.querySelectorAll("#categoriesTableBody tr");
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