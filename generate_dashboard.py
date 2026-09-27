import os
import glob
import re
import json
import pandas as pd
import numpy as np
import anthropic

REPORTS_DIR = "./reports"

def get_latest_sales_file():
    files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("Summary_") 
             and not os.path.basename(f).startswith("Sep_Target") 
             and not os.path.basename(f).startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No sales excel file found in {REPORTS_DIR}")
    return max(files, key=os.path.getctime)

def load_september_targets(target_file="Sep_Target.xlsx"):
    possible_paths = [target_file, os.path.join(REPORTS_DIR, target_file)]
    target_path = next((p for p in possible_paths if os.path.exists(p)), None)
    if not target_path:
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
            m = re.search(r'\b\d{3,8}\b', str(val))
            return m.group(0) if m else str(val).strip()

        df_t['clean_code'] = df_t[store_col].apply(extract_code)
        return dict(zip(df_t['clean_code'], df_t[sep_col]))
    except Exception as e:
        print(f"Target load error: {e}")
        return {}

def generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("[!] Warning: ANTHROPIC_API_KEY not found in environment.")
        return {
            "critical": "Category skew is evident across low ATV stores. Realign front gondolas towards higher ASP lifestyle items.",
            "attention": f"Network ASP is {network_asp:.2f} SAR with primary category reliance; monitor cross-category basket penetration.",
            "opportunity": "Top performing categories should receive priority stock replenishment across regional branch clusters."
        }

    top_stores = store_summary.head(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt', 'asp']].to_dict(orient="records")
    bottom_stores = store_summary.tail(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt', 'asp']].to_dict(orient="records")
    top_categories = cat_summary.head(4)[['Category Name', 'sales', 'contribution']].to_dict(orient="records")

    prompt = f"""
    You are a Retail Operations Executive. Based on the store and product mix performance below:
    - Total Sales: {total_sales:,.2f} SAR
    - Total Target: {total_target:,.0f} SAR
    - Network Achievement: {overall_ach:.1f}%
    - Network ATV: {network_atv:.2f} SAR
    - Network UPT: {network_upt:.2f}
    - Network ASP: {network_asp:.2f} SAR
    - Top Categories Contribution: {top_categories}
    - Top Stores: {top_stores}
    - Low Performing Stores: {bottom_stores}

    Provide 3 punchy, professional, and actionable business insights (1 sentence each):
    1. Critical Issues: direct operational problem or underperformer category/store risk.
    2. Attention Required: category penetration, basket building, or traffic conversion warning.
    3. Opportunities: merchandising or replenishment leverage for top categories and stores.

    Respond ONLY with valid JSON in this exact structure:
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
            "critical": "Category imbalances are impacting underperforming stores; enforce minimum core category stock depth.",
            "attention": f"Network ASP is {network_asp:.2f} SAR; push multi-item bundles across leading product categories.",
            "opportunity": "Replicate high-performing category visual merchandising standards across all retail branches."
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
    file_path = get_latest_sales_file()
    print(f"[*] Reading sales file: {file_path}")
    targets_map = load_september_targets("Sep_Target.xlsx")

    df = pd.read_excel(file_path, skiprows=1)
    df_clean = df.iloc[:-1].copy()
    
    # تنظيف أسماء الأعمدة من المحارف غير المرئية
    df_clean.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    # التحقق من عمود التصنيف
    cat_col = 'Category Name' if 'Category Name' in df_clean.columns else ('product_category' if 'product_category' in df_clean.columns else None)
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

    # 2. حساب مساهمة التصنيفات (Category Contribution - Network)
    cat_summary = df_clean.groupby(cat_col).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().rename(columns={cat_col: 'Category Name'})
    cat_summary['contribution'] = ((cat_summary['sales'] / total_sales) * 100).round(2)
    cat_summary['asp'] = (cat_summary['sales'] / cat_summary['units'].replace(0, np.nan)).fillna(0).round(2)
    cat_summary = cat_summary.sort_values(by='sales', ascending=False).reset_index(drop=True)

    # 3. حساب أعلى تصنيف ومساهمته لكل فرع
    store_cat = df_clean.groupby(['Organization Code', cat_col])['Actual Sales Amount'].sum().reset_index()
    top_cat_per_store = {}
    for code, group in store_cat.groupby('Organization Code'):
        top_row = group.sort_values(by='Actual Sales Amount', ascending=False).iloc[0]
        st_total = group['Actual Sales Amount'].sum()
        pct = (top_row['Actual Sales Amount'] / st_total * 100) if st_total > 0 else 0
        top_cat_per_store[code] = f"{top_row[cat_col]} ({pct:.1f}%)"

    store_summary['top_category'] = store_summary['Organization Code'].map(top_cat_per_store).fillna("-")

    # مطابقة الأهداف
    def match_target(row):
        code_str = str(row['Organization Code']).strip()
        name_str = str(row['Organization Name']).strip()
        for k, v in targets_map.items():
            if str(k) in code_str or str(k) in name_str:
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

    insights = generate_claude_insights(store_summary, cat_summary, total_sales, total_target, overall_ach, network_atv, network_upt, network_asp)

    chart_stores = store_summary.head(8)
    chart_svg_markup = build_svg_bar_chart(chart_stores)

    # كروت الـ Category Contribution
    colors = ['#38bdf8', '#818cf8', '#a855f7', '#ec4899', '#f59e0b', '#10b981', '#64748b']
    cat_cards_html = ""
    for idx, r in cat_summary.head(6).iterrows():
        c_color = colors[idx % len(colors)]
        cat_cards_html += f"""
        <div style="background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; min-width:180px; flex:1;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <span style="font-size:13px; font-weight:700; color:#f8fafc;">{r['Category Name']}</span>
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

    table_rows = ""
    for idx, row in store_summary.iterrows():
        if pd.notna(row['target']):
            target_str = f"{row['target']:,.0f}"
            ach_val = row['ach_pct']
            bar_w = min(ach_val, 100)
            if ach_val >= 100:
                color = "#10b981"
                status_badge = '<span class="badge badge-success">Target Met</span>'
            elif ach_val >= 80:
                color = "#f59e0b"
                status_badge = '<span class="badge badge-warning">On Track</span>'
            else:
                color = "#ef4444"
                status_badge = '<span class="badge badge-danger">Under Target</span>'

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
            status_badge = '<span class="badge" style="background:#1e293b;color:#94a3b8;">Normal</span>'

        table_rows += f"""
        <tr>
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:500;">{row['Organization Code']}</td>
            <td style="font-weight:600;color:#fff;">{row['Organization Name']}</td>
            <td style="font-weight:700;color:#f8fafc;">{row['sales']:,.2f}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:140px;">{ach_str}</td>
            <td>{row['share']:.2f}%</td>
            <td style="color:#cbd5e1;font-weight:500;">{row['top_category']}</td>
            <td>{int(row['txns']):,}</td>
            <td>{row['atv']:,.2f}</td>
            <td>{row['upt']:,.2f}</td>
            <td style="color:#38bdf8;font-weight:600;">{row['asp']:,.2f}</td>
            <td>{status_badge}</td>
        </tr>
        """

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS Executive KPI Dashboard</title>
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
        
        .section-title {{ font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); margin-bottom: 14px; display: flex; align-items: center; justify-content:space-between; flex-wrap:wrap; gap:10px; }}
        .insights-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .insight-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; border-left: 4px solid var(--border); }}
        .insight-card.danger {{ border-left-color: #ef4444; }}
        .insight-card.warning {{ border-left-color: #f59e0b; }}
        .insight-card.success {{ border-left-color: #10b981; }}
        .insight-title {{ font-size: 13px; font-weight: 700; margin-bottom: 6px; }}
        .insight-body {{ font-size: 13px; color: var(--text-muted); line-height: 1.5; }}

        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .kpi-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 18px; }}
        .kpi-title {{ font-size: 12px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; margin-bottom: 6px; }}
        .kpi-value {{ font-size: 24px; font-weight: 700; color: #fff; }}
        .kpi-unit {{ font-size: 13px; color: var(--text-muted); font-weight: 400; }}

        .chart-container {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 22px; margin-bottom: 24px; }}

        .table-wrap {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }}
        .table-header {{ padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); flex-wrap: wrap; gap: 12px; }}
        .table-header h3 {{ margin: 0; font-size: 15px; font-weight: 700; }}
        .table-search {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #fff; outline: none; width: 260px; font-size: 13px; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }}
        th {{ background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-bottom: 1px solid var(--border); }}
        td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); }}
        tr:hover td {{ background: var(--card-hover); }}

        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
        .badge-success {{ background: rgba(16, 185, 129, 0.15); color: #10b981; }}
        .badge-warning {{ background: rgba(245, 158, 11, 0.15); color: #f59e0b; }}
        .badge-danger {{ background: rgba(239, 68, 68, 0.15); color: #ef4444; }}
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

<div class="header">
    <div>
        <h1>MMS Executive KPI Dashboard</h1>
        <p>Operational Performance, Category Mix & Target Alignment</p>
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
        <div class="kpi-value">{total_sales:,.2f} <span class="kpi-unit">SAR</span></div>
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
        <div class="kpi-title">Network UPT</div>
        <div class="kpi-value">{network_upt:.2f}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ASP</div>
        <div class="kpi-value">SAR {network_asp:.2f}</div>
    </div>
</div>

<div class="section-title"><span>📦 Category Contribution (Network Mix)</span></div>
<div style="display:flex; flex-wrap:wrap; gap:14px; margin-bottom:24px;">
    {cat_cards_html}
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
            <h3>STORE PERFORMANCE MATRIX</h3>
            <span style="color:var(--text-muted);font-size:12px;">Ranked by revenue with achievement, leading category and store ASP</span>
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
                    <th>% Ach vs Target</th>
                    <th>Share %</th>
                    <th>Top Category Contribution</th>
                    <th>Txns</th>
                    <th>ATV (SAR)</th>
                    <th>UPT</th>
                    <th>ASP (SAR)</th>
                    <th>Status</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
    </div>
</div>

<script>
  const PASS = "MMS2026";

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