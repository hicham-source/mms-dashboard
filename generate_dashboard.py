import os
import glob
import json
import re
import pandas as pd
import numpy as np

# مسار مجلد التقارير
REPORTS_DIR = "./reports"

def get_latest_sales_file():
    """البحث عن أحدث ملف إكسل داخل مجلد reports"""
    files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("Summary_") and not os.path.basename(f).startswith("Sep_Target")]
    if not files:
        raise FileNotFoundError(f"لم يتم العثور على أي ملف إكسل للمبيعات داخل المجلد {REPORTS_DIR}")
    return max(files, key=os.path.getctime)

def load_september_targets(target_file="Sep_Target.xlsx"):
    """قراءة ملف التارجت واستخراج الفروع التي تملك تارجت فقط"""
    # البحث عن الملف في المجلد الرئيسي أو داخل reports
    possible_paths = [target_file, os.path.join(REPORTS_DIR, target_file)]
    target_path = next((p for p in possible_paths if os.path.exists(p)), None)
    
    if not target_path:
        print(f"[!] لم يتم العثور على ملف التارجت: {target_file}")
        return {}

    try:
        df_t = pd.read_excel(target_path)
        df_t.columns = [str(c).strip() for c in df_t.columns]
        
        store_col = [c for c in df_t.columns if "profit" in c.lower() or "cost" in c.lower()][0]
        sep_col = [c for c in df_t.columns if "sep" in c.lower()][0]

        # استبعاد سطر Total
        df_t = df_t[~df_t[store_col].astype(str).str.lower().str.contains("total")].copy()

        # تحويل الأرقام واستبعاد الفارغ والشرطات
        df_t[sep_col] = pd.to_numeric(df_t[sep_col].astype(str).str.replace(",", "").str.strip(), errors='coerce')
        df_t = df_t[df_t[sep_col].notna() & (df_t[sep_col] > 0)].copy()

        def extract_code(val):
            m = re.search(r'\b\d{3,8}\b', str(val))
            return m.group(0) if m else str(val).strip()

        df_t['clean_code'] = df_t[store_col].apply(extract_code)
        targets_dict = dict(zip(df_t['clean_code'], df_t[sep_col]))
        print(f"[*] تم تحميل التارجت لـ {len(targets_dict)} فرع بنجاح.")
        return targets_dict
    except Exception as e:
        print(f"[!] خطأ في قراءة ملف التارجت: {e}")
        return {}

def process_and_build():
    file_path = get_latest_sales_file()
    print(f"[*] Reading sales file: {file_path}")
    targets_map = load_september_targets("Sep_Target.xlsx")

    # 1. قراءة وتنظيف ملف المبيعات
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

    # 2. ملخص أداء الفروع
    store_summary = df_clean.groupby(['Organization Code', 'Organization Name']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index()

    store_summary['atv'] = (store_summary['sales'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['upt'] = (store_summary['units'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['asp'] = (store_summary['sales'] / store_summary['units'].replace(0, np.nan)).fillna(0).round(2)
    
    total_sales = store_summary['sales'].sum()
    store_summary['share'] = ((store_summary['sales'] / total_sales) * 100).round(2)
    store_summary = store_summary.sort_values(by='sales', ascending=False)

    # 3. ربط التارجت لكل فرع
    def match_target(row):
        code_str = str(row['Organization Code']).strip()
        # محاولة المطابقة المباشرة بالكود
        for k, v in targets_map.items():
            if str(k) in code_str or str(k) in str(row['Organization Name']):
                return v
        return None

    store_summary['target'] = store_summary.apply(match_target, axis=1)
    store_summary['ach_pct'] = store_summary.apply(
        lambda r: (r['sales'] / r['target']) * 100 if pd.notna(r['target']) and r['target'] > 0 else None,
        axis=1
    )

    # حساب إجمالي التارجت والنسبة للشبكة (للفروع التي تملك تارجت فقط)
    valid_targets = store_summary[store_summary['target'].notna()]
    total_target = valid_targets['target'].sum()
    sales_with_target = valid_targets['sales'].sum()
    overall_ach = (sales_with_target / total_target * 100) if total_target > 0 else 0

    # 4. بناء صفوف جدول الفروع
    table_rows = ""
    for idx, row in store_summary.iterrows():
        # التارجت ونسبة التحقيق
        if pd.notna(row['target']):
            target_str = f"{row['target']:,.0f}"
            ach_val = row['ach_pct']
            color = "#10b981" if ach_val >= 100 else ("#f59e0b" if ach_val >= 80 else "#ef4444")
            ach_str = f'<span style="color: {color}; font-weight: bold;">{ach_val:.1f}%</span>'
        else:
            target_str = "-"
            ach_str = '<span style="color: #64748b;">-</span>'

        table_rows += f"""
        <tr>
            <td style="text-align: right; font-weight: 600;">{row['Organization Name']}</td>
            <td>{row['sales']:,.2f}</td>
            <td style="color: #38bdf8; font-weight: 500;">{target_str}</td>
            <td>{ach_str}</td>
            <td>{int(row['txns']):,}</td>
            <td>{int(row['units']):,}</td>
            <td>{row['atv']:,.2f}</td>
            <td>{row['upt']:,.2f}</td>
            <td>{row['share']:.1f}%</td>
        </tr>
        """

    # 5. بناء صفحة HTML متكاملة مع الحماية والباسورد
    html_content = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS Retail OS - لوحة مؤشرات الأداء</title>
    <style>
        :root {{
            --bg: #090d16;
            --card-bg: #131b2e;
            --border: #1e293b;
            --text-main: #f8fafc;
            --text-sub: #94a3b8;
            --primary: #2563eb;
            --success: #10b981;
            --warning: #f59e0b;
        }}
        body {{
            background-color: var(--bg);
            color: var(--text-main);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            margin: 0;
            padding: 24px;
        }}
        .header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 20px;
            margin-bottom: 25px;
        }}
        .header h1 {{ margin: 0; font-size: 24px; color: #fff; }}
        .header p {{ margin: 5px 0 0 0; color: var(--text-sub); font-size: 14px; }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 30px;
        }}
        .kpi-card {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            padding: 20px;
            border-radius: 12px;
        }}
        .kpi-title {{ font-size: 13px; color: var(--text-sub); margin-bottom: 8px; }}
        .kpi-value {{ font-size: 26px; font-weight: bold; color: #fff; }}
        .table-container {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow-x: auto;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
            font-size: 14px;
        }}
        th {{
            background: #1e293b;
            color: #cbd5e1;
            padding: 14px 12px;
            font-weight: 600;
            border-bottom: 1px solid var(--border);
        }}
        td {{
            padding: 12px;
            border-bottom: 1px solid #1a2234;
            color: #cbd5e1;
        }}
        tr:hover td {{ background: #182238; }}
    </style>
</head>
<body>

<!-- شاشة القفل بكلمة المرور -->
<div id="auth-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:#090d16;z-index:999999;display:flex;align-items:center;justify-content:center;font-family:sans-serif;direction:rtl;">
  <div style="background:#131b2e;padding:32px;border-radius:14px;box-shadow:0 15px 30px rgba(0,0,0,0.6);text-align:center;width:90%;max-width:380px;border:1px solid #1e293b;">
    <h3 style="color:#fff;margin-bottom:8px;font-size:20px;">🔒 نظام مؤشرات الأداء (MMS)</h3>
    <p style="color:#94a3b8;font-size:14px;margin-bottom:20px;">أدخل رمز المرور للمتابعة</p>
    <input type="password" id="access-pass" placeholder="كلمة السر" style="width:100%;padding:12px;border-radius:8px;border:1px solid #334155;background:#090d16;color:#fff;font-size:16px;text-align:center;outline:none;box-sizing:border-box;margin-bottom:15px;">
    <button onclick="checkAccess()" style="width:100%;padding:12px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:bold;font-size:16px;cursor:pointer;">دخول</button>
    <p id="error-msg" style="color:#ef4444;font-size:13px;margin-top:12px;display:none;">كلمة السر غير صحيحة</p>
  </div>
</div>

<script>
  const PASS = "MMS2026";
  if (sessionStorage.getItem("mms_auth") === "ok") {{
    document.getElementById("auth-overlay").style.display = "none";
  }}
  function checkAccess() {{
    if (document.getElementById("access-pass").value === PASS) {{
      sessionStorage.setItem("mms_auth", "ok");
      document.getElementById("auth-overlay").style.display = "none";
    }} else {{
      document.getElementById("error-msg").style.display = "block";
    }}
  }}
  document.getElementById("access-pass").addEventListener("keypress", function(e) {{
    if (e.key === "Enter") checkAccess();
  }});
</script>

<div class="header">
    <div>
        <h1>MMS Executive KPI Dashboard</h1>
        <p>نظام متابعة العمليات وأداء الفروع اليومي</p>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-title">إجمالي المبيعات (Total Sales)</div>
        <div class="kpi-value">{total_sales:,.2f} <span style="font-size:14px;color:#94a3b8;">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">إجمالي التارجت (Target)</div>
        <div class="kpi-value">{total_target:,.0f} <span style="font-size:14px;color:#94a3b8;">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">نسبة التحقيق الإجمالية (% Ach)</div>
        <div class="kpi-value" style="color: {'#10b981' if overall_ach >= 100 else ('#f59e0b' if overall_ach >= 80 else '#ef4444')};">{overall_ach:.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">عدد الفروع المفتوحة</div>
        <div class="kpi-value">{len(store_summary)}</div>
    </div>
</div>

<div class="table-container">
    <table>
        <thead>
            <tr>
                <th style="text-align: right;">الفرع</th>
                <th>المبيعات (SAR)</th>
                <th>التارجت (SAR)</th>
                <th>نسبة التحقيق (% Ach)</th>
                <th>الفواتير</th>
                <th>القطع</th>
                <th>ATV</th>
                <th>UPT</th>
                <th>الحصة (%)</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
    </table>
</div>

</body>
</html>
    """

    out_file = os.path.join(REPORTS_DIR, "MMS_Executive_KPI_Dashboard.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"[✓] Dashboard generated successfully: {out_file}")

if __name__ == "__main__":
    process_and_build()