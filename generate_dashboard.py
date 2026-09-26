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

def generate_claude_insights(store_summary, total_sales, total_target, overall_ach, network_atv, network_upt):
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("[!] Warning: ANTHROPIC_API_KEY not found in environment.")
        return {
            "critical": "Stores showing high footfall but low ATV require immediate cashier upselling initiatives.",
            "attention": "High ATV locations require visual merchandising optimization to drive higher walk-in conversion.",
            "opportunity": "Top performing stores continue to lead network revenue; maintain full stock availability on high-velocity items."
        }

    top_stores = store_summary.head(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt']].to_dict(orient="records")
    bottom_stores = store_summary.tail(3)[['Organization Name', 'sales', 'ach_pct', 'atv', 'upt']].to_dict(orient="records")

    prompt = f"""
    You are a Retail Operations Executive. Based on the store performance below:
    - Total Sales: {total_sales:,.2f} SAR
    - Total Target: {total_target:,.0f} SAR
    - Network Achievement: {overall_ach:.1f}%
    - Network ATV: {network_atv:.2f} SAR
    - Network UPT: {network_upt:.2f}
    - Top Stores: {top_stores}
    - Low Performing Stores: {bottom_stores}

    Provide 3 punchy, professional, and actionable business insights (1 sentence each):
    1. Critical Issues: direct operational problem or underperformer risk.
    2. Attention Required: basket size, UPT, or traffic conversion warning.
    3. Opportunities: merchandising or replenishment leverage for top volume drivers.

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
            "critical": "Underperforming locations require focused cross-selling incentives to lift transaction value.",
            "attention": "Monitor traffic to transaction conversion ratios across regional mall locations.",
            "opportunity": "Scale high-velocity display configurations from top-performing branches."
        }

def process_and_build():
    file_path = get_latest_sales_file()
    print(f"[*] Reading sales file: {file_path}")
    targets_map = load_september_targets("Sep_Target.xlsx")

    df = pd.read_excel(file_path, skiprows=1)
    df_clean = df.iloc[:-1].copy()
    df_clean.columns = [c.replace('\u200c', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

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

    insights = generate_claude_insights(store_summary, total_sales, total_target, overall_ach, network_atv, network_upt)

    chart_stores = store_summary.head(8)
    chart_labels = chart_stores['Organization Name'].tolist()
    chart_sales = chart_stores['sales'].round(2).tolist()
    chart_targets = [round(r['target'], 2) if pd.notna(r['target']) else 0 for _, r in chart_stores.iterrows()]

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
            <td>{int(row['txns']):,}</td>
            <td>{row['atv']:,.2f}</td>
            <td>{row['upt']:,.2f}</td>
            <td>{status_badge}</td>
        </tr>
        """

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS Executive KPI Dashboard</title>
    <script src="[https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js](https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js)"></script>
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
        
        .section-title {{ font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }}
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

        .chart-container {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 20px; margin-bottom: 24px; min-height: 380px; }}

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
        <p>Operational Performance, Store Target Alignment & Metrics</p>
    </div>
</div>

<div class="section-title"><span>🤖</span> AI Executive Insights (Powered by Claude)</div>
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
        <div class="kpi-title">Avg Unit Price (ASP)</div>
        <div class="kpi-value">SAR {network_asp:.2f}</div>
    </div>
</div>

<div class="chart-container">
    <div class="section-title"><span>📊</span> Top Stores: Actual Sales vs Target</div>
    <div style="position:relative; height:300px; width:100%;">
        <canvas id="salesTargetChart"></canvas>
    </div>
</div>

<div class="table-wrap">
    <div class="table-header">
        <div>
            <h3>STORE PERFORMANCE MATRIX</h3>
            <span style="color:var(--text-muted);font-size:12px;">Ranked by revenue with achievement progress and operational badges</span>
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
                    <th>Txns</th>
                    <th>ATV (SAR)</th>
                    <th>UPT</th>
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
  let chartLoaded = false;

  function renderChart() {{
    if (chartLoaded) return;
    const canvas = document.getElementById('salesTargetChart');
    if (!canvas) return;

    try {{
      const ctx = canvas.getContext('2d');
      new Chart(ctx, {{
          type: 'bar',
          data: {{
              labels: {json.dumps(chart_labels)},
              datasets: [
                  {{
                      label: 'Actual Sales (SAR)',
                      data: {json.dumps(chart_sales)},
                      backgroundColor: '#38bdf8',
                      borderRadius: 4
                  }},
                  {{
                      label: 'Target (SAR)',
                      data: {json.dumps(chart_targets)},
                      backgroundColor: '#334155',
                      borderRadius: 4
                  }}
              ]
          }},
          options: {{
              responsive: true,
              maintainAspectRatio: false,
              plugins: {{
                  legend: {{ labels: {{ color: '#94a3b8' }} }}
              }},
              scales: {{
                  x: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ display: false }} }},
                  y: {{ ticks: {{ color: '#94a3b8' }}, grid: {{ color: '#1e293b' }} }}
              }}
          }}
      }});
      chartLoaded = true;
    }} catch(e) {{
      console.error("Chart error:", e);
    }}
  }}

  function checkAccess() {{
    const val = document.getElementById("access-pass").value;
    if (val === PASS) {{
      sessionStorage.setItem("mms_auth", "ok");
      document.getElementById("auth-overlay").style.display = "none";
      setTimeout(renderChart, 100);
    }} else {{
      document.getElementById("error-msg").style.display = "block";
    }}
  }}

  document.getElementById("access-pass").addEventListener("keypress", function(e) {{
    if (e.key === "Enter") checkAccess();
  }});

  // إذا كان المستخدم قد دخل مسبقاً
  if (sessionStorage.getItem("mms_auth") === "ok") {{
    document.getElementById("auth-overlay").style.display = "none";
    window.addEventListener('DOMContentLoaded', () => {{ setTimeout(renderChart, 100); }});
    window.addEventListener('load', () => {{ setTimeout(renderChart, 100); }});
    setTimeout(renderChart, 300);
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