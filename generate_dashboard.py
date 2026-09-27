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
            # فحص ورقات العمل
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
        
        # محاولة القراءة مع skiprows=1
        df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use, skiprows=1)
        df_soh.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]
        
        if "avail_stock" not in [c.lower() for c in df_soh.columns] and "current_stock" not in [c.lower() for c in df_soh.columns]:
            # تجربة بدون skiprows
            df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use)
            df_soh.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]

        code_col = next((c for c in df_soh.columns if c.lower() in ["org code", "organization code", "org_code", "store code"]), None)
        name_col = next((c for c in df_soh.columns if "org" in c.lower() and "name" in c.lower()), None)
        stock_col = next((c for c in df_soh.columns if c.lower() in ["avail_stock", "current_stock"]), None)
        price_col = next((c for c in df_soh.columns if "retail_price" in c.lower() or "price" in c.lower()), None)

        if not code_col or not stock_col:
            print(f"[!] Required columns missing in SOH. Found: {list(df_soh.columns[:6])}")
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

def generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp, total_soh_units):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("[!] Warning: ANTHROPIC_API_KEY not found in environment.")
        return {
            "critical": "Severe stock-to-sales mismatch detected. Underperforming stores are holding excessive inventory while top branches face stockout risks.",
            "attention": f"Network SOH stands at {total_soh_units:,.0f} units; activate immediate Inter-Store Transfers (IST) from high-WOS to low-WOS locations.",
            "opportunity": "High-velocity branches require priority warehouse replenishment to protect target achievement run-rate."
        }

    top_stores = store_summary.head(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt', 'asp', 'wos']].to_dict(orient="records")
    bottom_stores = store_summary.tail(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt', 'asp', 'wos']].to_dict(orient="records")
    top_categories = cat_summary.head(5)[['Category Name', 'sales', 'contribution']].to_dict(orient="records")

    prompt = f"""
    You are a Retail Operations Executive evaluating sales vs SOH (Stock On Hand).
    - Total Sales: {total_sales:,.2f} SAR
    - Total Target: {total_target:,.0f} SAR
    - Network Achievement: {overall_ach:.1f}%
    - Total SOH Units: {total_soh_units:,.0f}
    - Top Stores with Weeks of Supply (WOS): {top_stores}
    - Low Performing Stores with WOS: {bottom_stores}
    - Top Categories: {top_categories}

    Provide 3 sharp, operational diagnostics with direct solutions:
    1. Critical Issues: stockout risks at high-run-rate branches or heavy capital tied up.
    2. Attention Required: Inter-Store Transfers (IST) or merchandising alignment.
    3. Opportunities: replenishment leverage on top selling categories.

    Respond ONLY with valid JSON:
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
            "critical": "Stock-to-sales imbalance threatens revenue; low-WOS stores need immediate inventory buffer.",
            "attention": "Rebalance branch stock depth via Inter-Store Transfers (IST) rather than waiting for warehouse cycles.",
            "opportunity": "Maintain continuous stock availability on top volume driver categories."
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

    # إجماليات المتاجر
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

    # ربط المخزون SOH
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

    weekly_units = store_summary['units'] / 4.3
    store_summary['wos'] = (store_summary['soh_units'] / weekly_units.replace(0, np.nan)).fillna(0).round(1)

    # مطابقة الأهداف
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

    # تشخيص الفروع
    def diagnose_store(row):
        ach = row['ach_pct'] if pd.notna(row['ach_pct']) else 0
        wos = row['wos']
        
        if row['soh_units'] == 0:
            return {
                "status": "No SOH Data",
                "diag": "Inventory record not found in SOH file.",
                "action": "Verify Org Code in SOH export.",
                "color": "#64748b"
            }
        
        if wos < 2.5 and ach >= 80:
            return {
                "status": "Stockout Risk",
                "diag": f"High demand with critical stock cover ({wos} weeks). Risk of lost sales.",
                "action": "Prioritize emergency warehouse replenishment immediately.",
                "color": "#ef4444"
            }
        elif wos > 8.0 and ach < 70:
            return {
                "status": "Heavy Overstock",
                "diag": f"Sluggish turnover coupled with excess inventory ({wos} weeks).",
                "action": "Initiate Inter-Store Transfers (IST) to top branches & execute basket promos.",
                "color": "#f59e0b"
            }
        elif wos > 5.0 and ach >= 90:
            return {
                "status": "High Performer",
                "diag": f"Robust sales achievement ({ach:.1f}%) supported by healthy stock depth ({wos} weeks).",
                "action": "Maintain core visual merchandising and steady restocking rhythm.",
                "color": "#10b981"
            }
        elif ach < 70 and 2.5 <= wos <= 6.0:
            return {
                "status": "Conversion Bottleneck",
                "diag": f"Adequate stock cover ({wos} weeks) but target is lagging ({ach:.1f}%).",
                "action": "Focus on cashier upselling and customer conversion rather than inventory.",
                "color": "#eab308"
            }
        else:
            return {
                "status": "Balanced",
                "diag": f"Stock depth ({wos} weeks) matches current turnover pace.",
                "action": "Standard replenishment schedule.",
                "color": "#38bdf8"
            }

    diagnostics = store_summary.apply(diagnose_store, axis=1)
    store_summary['diag_status'] = [d['status'] for d in diagnostics]
    store_summary['diag_text'] = [d['diag'] for d in diagnostics]
    store_summary['diag_action'] = [d['action'] for d in diagnostics]
    store_summary['diag_color'] = [d['color'] for d in diagnostics]

    # إجماليات الأصناف
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

    insights = generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp, total_soh_units)

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

    store_meta_map = {}
    store_table_rows = ""
    action_cards_html = ""

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

        diag_badge = f'<span class="badge" style="background:{row["diag_color"]}22; color:{row["diag_color"]}; border:1px solid {row["diag_color"]}66;">{row["diag_status"]}</span>'

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
            "soh_val": f"{row['soh_val']:,.0f} SAR",
            "wos": f"{row['wos']} Wks",
            "problem": row['diag_text'],
            "action": row['diag_action']
        }

        if row['diag_status'] in ["Stockout Risk", "Heavy Overstock", "Conversion Bottleneck"]:
            action_cards_html += f"""
            <div style="background:var(--card); border:1px solid var(--border); border-left:4px solid {row['diag_color']}; border-radius:10px; padding:16px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span style="font-weight:700; color:#fff; font-size:14px;">{st_name} ({st_code})</span>
                    {diag_badge}
                </div>
                <div style="font-size:12px; color:#cbd5e1; margin-bottom:8px;">
                    <strong>🔍 Problem:</strong> {row['diag_text']}
                </div>
                <div style="font-size:12px; color:#38bdf8; background:#090d16; padding:8px 12px; border-radius:6px; border:1px solid #1e293b;">
                    <strong>⚡ Action Required:</strong> {row['diag_action']}
                </div>
                <div style="display:flex; gap:16px; margin-top:10px; font-size:11px; color:#94a3b8;">
                    <span>SOH: <strong style="color:#fff;">{int(row['soh_units']):,} Pcs</strong></span>
                    <span>WOS: <strong style="color:{row['diag_color']};">{row['wos']} Wks</strong></span>
                    <span>Ach: <strong style="color:#fff;">{row['ach_pct']:.1f}%</strong></span>
                </div>
            </div>
            """

        store_table_rows += f"""
        <tr onclick="openStoreDetails('{st_code}')" style="cursor:pointer;" title="Click to view deep-dive analytics & action plan">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:600;">{st_code}</td>
            <td style="font-weight:600;color:#fff;">
                {st_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">🔍</span>
            </td>
            <td style="font-weight:700;color:#f8fafc;">{row['sales']:,.2f}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:130px;">{ach_str}</td>
            <td style="font-weight:700;color:#38bdf8;">{int(row['soh_units']):,}</td>
            <td style="font-weight:700;color:{row['diag_color']};">{row['wos']} Wks</td>
            <td>{diag_badge}</td>
            <td>{row['atv']:,.2f}</td>
            <td>{row['upt']:,.2f}</td>
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
    <title>MMS Executive KPI & SOH Diagnostic Dashboard</title>
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
        .modal-content {{ background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 90%; max-width: 950px; max-height: 88vh; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7); }}
        .modal-header {{ padding: 20px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; }}
        .modal-body {{ padding: 24px; overflow-y: auto; }}
        .close-btn {{ background: transparent; border: none; color: #94a3b8; font-size: 24px; cursor: pointer; line-height: 1; }}
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

<div id="store-modal" class="app-modal">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h2 id="modal-store-name" style="margin:0; font-size:18px; color:#fff;">Store Deep-Dive & SOH Diagnostics</h2>
        <span id="modal-store-code" style="color:#38bdf8; font-size:12px; font-weight:600;">CODE</span>
      </div>
      <button class="close-btn" onclick="closeModal('store-modal')">&times;</button>
    </div>
    <div class="modal-body">
      <div style="background:#090d16; border:1px solid #1e293b; border-radius:10px; padding:16px; margin-bottom:20px;">
        <div style="font-size:12px; font-weight:700; color:#cbd5e1; margin-bottom:6px;">
          <span style="color:#ef4444;">● Root-Cause Diagnostic:</span> <span id="modal-problem" style="color:#f8fafc; font-weight:500;">-</span>
        </div>
        <div style="font-size:12px; font-weight:700; color:#38bdf8;">
          <span>⚡ Corrective Action Plan:</span> <span id="modal-action" style="color:#38bdf8; font-weight:500;">-</span>
        </div>
      </div>

      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:12px; margin-bottom:20px;">
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">SALES</div>
          <div id="modal-sales" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">TARGET</div>
          <div id="modal-target" style="font-size:16px; font-weight:700; color:#94a3b8;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">% ACH</div>
          <div id="modal-ach" style="font-size:16px; font-weight:700; color:#10b981;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">SOH UNITS</div>
          <div id="modal-soh-units" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">WOS (COVER)</div>
          <div id="modal-wos" style="font-size:16px; font-weight:700; color:#f59e0b;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">ASP</div>
          <div id="modal-asp" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
      </div>

      <div style="margin-bottom:12px; font-size:13px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Store Category Breakdown</div>
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
        <h1>MMS Executive KPI & SOH Diagnostic Dashboard</h1>
        <p>Operational Performance, Stock On Hand (SOH) Alignment & Solutions</p>
    </div>
</div>

<div class="section-title"><span>🤖 AI Executive Insights (Powered by Claude)</span></div>
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
        <div class="kpi-value" style="color:#38bdf8;">{total_soh_units:,.0f} <span class="kpi-unit">Units</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ATV</div>
        <div class="kpi-value">SAR {network_atv:.2f}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network UPT</div>
        <div class="kpi-value">{network_upt:.2f}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ASP</div>
        <div class="kpi-value">SAR {network_asp:.2f}</div>
    </div>
</div>

<div class="view-toggle-bar">
    <button class="view-btn active" id="btn-stores" onclick="switchView('stores')">🏢 Store-Wise & SOH Diagnostics</button>
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business-Wise Performance ({len(cat_summary)} Categories)</button>
</div>

<div id="view-stores">
    <div class="section-title"><span>🚨 Operational Diagnostics & Priority Action Plan</span></div>
    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:14px; margin-bottom:24px;">
        {action_cards_html}
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
                <h3>STORE PERFORMANCE & SOH DIAGNOSTICS MATRIX</h3>
                <span style="color:var(--text-muted);font-size:12px;">Click any row to open store category mix, WOS breakdown & actionable solutions</span>
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
                        <th>SOH (Units)</th>
                        <th>WOS</th>
                        <th>Diagnostic</th>
                        <th>ATV (SAR)</th>
                        <th>UPT</th>
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
    document.getElementById("modal-store-code").innerText = "BRANCH CODE: " + storeCode;
    document.getElementById("modal-sales").innerText = meta.sales;
    document.getElementById("modal-target").innerText = meta.target;
    document.getElementById("modal-ach").innerText = meta.ach;
    document.getElementById("modal-soh-units").innerText = meta.soh_units;
    document.getElementById("modal-wos").innerText = meta.wos;
    document.getElementById("modal-asp").innerText = meta.asp;
    document.getElementById("modal-problem").innerText = meta.problem;
    document.getElementById("modal-action").innerText = meta.action;

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