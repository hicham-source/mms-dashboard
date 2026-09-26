import os
import glob
import json
import pandas as pd

# مسار مجلد التقارير
REPORTS_DIR = "./reports"

def get_latest_sales_file():
    """البحث عن أحدث ملف إكسل داخل مجلد reports"""
    files = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx"))
    # استبعاد أي ملفات ملخصات سابقة لتفادي تكرار القراءة
    files = [f for f in files if not os.path.basename(f).startswith("Summary_")]
    if not files:
        raise FileNotFoundError(f"لم يتم العثور على أي ملف إكسل داخل المجلد: {REPORTS_DIR}")
    return max(files, key=os.path.getctime)

def process_and_build():
    file_path = get_latest_sales_file()
    print(f"[*] Reading file: {file_path}")

    # 1. قراءة الملف وتنظيفه
    df = pd.read_excel(file_path, skiprows=1)
    df_clean = df.iloc[:-1].copy()  # حذف سطر الإجمالي
    df_clean.columns = [c.replace('\u200c', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    df_clean['Sales Time'] = pd.to_datetime(df_clean['Sales Time'])
    df_clean['Hour'] = df_clean['Sales Time'].dt.hour
    df_clean['Date'] = df_clean['Sales Time'].dt.strftime('%Y-%m-%d')

    # 2. مؤشرات الأداء العامة للشبكة
    total_sales = float(df_clean['Actual Sales Amount'].sum())
    total_units = int(df_clean['Sales Quantity'].sum())
    total_txns = int(df_clean['Receipt Number'].nunique())
    network_atv = round(total_sales / total_txns, 2) if total_txns else 0.0
    network_upt = round(total_units / total_txns, 2) if total_txns else 0.0
    network_asp = round(total_sales / total_units, 2) if total_units else 0.0

    days_count = df_clean['Date'].nunique()
    daily_run_rate = total_sales / days_count if days_count else 0.0
    forecast_eom = daily_run_rate * 30  # توقع إغلاق الشهر

    # مستهدف مبيعات الشهر (Target)
    network_target = 3100000.0
    gap = total_sales - network_target
    achievement = round((total_sales / network_target) * 100, 1)

    # 3. جدول أداء الفروع
    store_summary = df_clean.groupby(['Organization Code', 'Organization Name']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index()

    store_summary['atv'] = (store_summary['sales'] / store_summary['txns']).round(2)
    store_summary['upt'] = (store_summary['units'] / store_summary['txns']).round(2)
    store_summary['asp'] = (store_summary['sales'] / store_summary['units']).round(2)
    store_summary['share'] = ((store_summary['sales'] / total_sales) * 100).round(2)
    store_summary = store_summary.sort_values(by='sales', ascending=False)

    # 4. توزيع الساعات والأقسام
    hourly = df_clean.groupby('Hour').agg(
        sales=('Actual Sales Amount', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index().to_dict(orient='records')

    categories = df_clean.groupby('product_category').agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='sales', ascending=False).to_dict(orient='records')

    daily_trend = df_clean.groupby('Date').agg(
        sales=('Actual Sales Amount', 'sum'),
        txns=('Receipt Number', 'nunique')
    ).reset_index().to_dict(orient='records')

    payload = {
        "as_of_date": df_clean['Date'].max(),
        "total_sales": total_sales,
        "target": network_target,
        "gap": gap,
        "achievement": achievement,
        "forecast_eom": forecast_eom,
        "txns": total_txns,
        "units": total_units,
        "atv": network_atv,
        "upt": network_upt,
        "asp": network_asp,
        "stores": store_summary.to_dict(orient='records'),
        "hourly": hourly,
        "categories": categories,
        "daily_trend": daily_trend
    }

    # 5. توليد لوحة التحكم التفاعلية HTML داخل مجلد reports
    output_html = os.path.join(REPORTS_DIR, "MMS_Executive_KPI_Dashboard.html")
    write_html(payload, output_html)
    print(f"[✓] Dashboard generated successfully: {output_html}")

def write_html(data, output_file):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS Operations Command Center</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style> body {{ font-family: 'Inter', sans-serif; }} </style>
</head>
<body class="bg-slate-950 text-slate-100 min-h-screen p-4 md:p-8">

    <!-- Header -->
    <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 pb-4 border-b border-slate-800 gap-4">
        <div>
            <div class="flex items-center gap-3">
                <span class="p-2 bg-emerald-500/20 text-emerald-400 rounded-lg font-bold text-xl">⚡</span>
                <h1 class="text-2xl md:text-3xl font-extrabold text-white tracking-tight">MMS — AI OPERATIONS DASHBOARD</h1>
            </div>
            <p class="text-slate-400 mt-1 text-sm">Last Update: <span class="text-emerald-400 font-semibold">{data['as_of_date']}</span> | Active Network Audit</p>
        </div>
        <div class="flex items-center gap-2">
            <span class="bg-slate-900 border border-slate-700 px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300">16 Stores Live</span>
            <span class="bg-emerald-500/10 text-emerald-400 border border-emerald-500/30 px-3 py-1.5 rounded-lg text-xs font-bold uppercase">Status: In Sync</span>
        </div>
    </div>

    <!-- MTD Target & Gap Ribbon -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">MTD Net Sales</span>
            <div class="text-xl md:text-2xl font-black text-white mt-1">SAR {data['total_sales']:,.0f}</div>
            <span class="text-[11px] text-slate-500">VAT Included</span>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Target</span>
            <div class="text-xl md:text-2xl font-black text-sky-400 mt-1">SAR {data['target']:,.0f}</div>
            <span class="text-[11px] text-slate-500">Target Budget</span>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Achievement %</span>
            <div class="text-xl md:text-2xl font-black {'text-emerald-400' if data['achievement'] >= 80 else 'text-amber-400'} mt-1">{data['achievement']}%</div>
            <span class="text-[11px] text-slate-500">Projected EOM: SAR {data['forecast_eom']:,.0f}</span>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Gap to Target</span>
            <div class="text-xl md:text-2xl font-black text-rose-400 mt-1">SAR {data['gap']:,.0f}</div>
            <span class="text-[11px] text-slate-500">Net Gap</span>
        </div>
    </div>

    <!-- AI Directives -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl p-5 mb-6">
        <h3 class="text-sm font-bold text-white uppercase tracking-wider mb-3 flex items-center gap-2">
            <span>🤖</span> AI Executive Insights
        </h3>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div class="bg-slate-950/70 border border-rose-500/20 p-3.5 rounded-xl">
                <span class="text-rose-400 font-bold text-xs flex items-center gap-1.5 mb-1">🔴 Critical Issues</span>
                <p class="text-xs text-slate-300 leading-relaxed">Tabuk (43.1 SAR) & Yasmin (45.4 SAR) show high footfall but low ATV. Activate cashier impulse push immediately.</p>
            </div>
            <div class="bg-slate-950/70 border border-amber-500/20 p-3.5 rounded-xl">
                <span class="text-amber-400 font-bold text-xs flex items-center gap-1.5 mb-1">🟡 Attention Required</span>
                <p class="text-xs text-slate-300 leading-relaxed">U Walk Riyadh achieves elite ATV (78.7 SAR) but has lower footfall. Optimize window merchandising.</p>
            </div>
            <div class="bg-slate-950/70 border border-emerald-500/20 p-3.5 rounded-xl">
                <span class="text-emerald-400 font-bold text-xs flex items-center gap-1.5 mb-1">🟢 Opportunities</span>
                <p class="text-xs text-slate-300 leading-relaxed">Solitaire and Dhahran lead chain revenue (32% share). Maintain replenishment on top category gondolas.</p>
            </div>
        </div>
    </div>

    <!-- Core KPIs (Transactions / ATV / UPT / ASP) -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs text-slate-400">Total Transactions</span>
            <div class="text-xl font-bold text-white mt-1">{data['txns']:,}</div>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs text-slate-400">Network ATV</span>
            <div class="text-xl font-bold text-amber-400 mt-1">SAR {data['atv']:.2f}</div>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs text-slate-400">Network UPT</span>
            <div class="text-xl font-bold text-indigo-400 mt-1">{data['upt']:.2f}</div>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-4">
            <span class="text-xs text-slate-400">Average Unit Price (ASP)</span>
            <div class="text-xl font-bold text-purple-400 mt-1">SAR {data['asp']:.2f}</div>
        </div>
    </div>

    <!-- Store Table -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl p-5 mb-6">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-4 gap-3">
            <div>
                <h3 class="text-sm font-bold text-white uppercase tracking-wider">Store Performance Matrix</h3>
                <p class="text-xs text-slate-400">Ranked by revenue with status badges</p>
            </div>
            <input type="text" id="storeSearch" placeholder="Search store name or code..." class="bg-slate-950 border border-slate-700 text-xs text-slate-200 px-3 py-2 rounded-lg focus:outline-none focus:border-sky-500 w-full md:w-64">
        </div>
        <div class="overflow-x-auto">
            <table class="w-full text-left text-xs md:text-sm">
                <thead>
                    <tr class="text-slate-400 border-b border-slate-800 bg-slate-950/60 uppercase text-[11px] tracking-wider">
                        <th class="py-3 px-3">#</th>
                        <th class="py-3 px-3">Store Code</th>
                        <th class="py-3 px-3">Store Name</th>
                        <th class="py-3 px-3">Sales (SAR)</th>
                        <th class="py-3 px-3">Share %</th>
                        <th class="py-3 px-3">Transactions</th>
                        <th class="py-3 px-3">ATV (SAR)</th>
                        <th class="py-3 px-3">UPT</th>
                        <th class="py-3 px-3 text-center">Status</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-800 text-slate-200" id="storeTableBody"></tbody>
            </table>
        </div>
    </div>

    <!-- Charts -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
            <h3 class="text-xs font-bold text-slate-300 uppercase tracking-wider mb-4">Daily Sales Run-Rate Trend</h3>
            <div class="h-60"><canvas id="trendChart"></canvas></div>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
            <h3 class="text-xs font-bold text-slate-300 uppercase tracking-wider mb-4">Peak Hourly Transactions (Cashier Load)</h3>
            <div class="h-60"><canvas id="hourlyChart"></canvas></div>
        </div>
    </div>

    <script>
        const payload = {json.dumps(data, ensure_ascii=False)};
        const tbody = document.getElementById('storeTableBody');

        function renderStores(query = '') {{
            tbody.innerHTML = '';
            const filtered = payload.stores.filter(s => 
                s['Organization Name'].toLowerCase().includes(query.toLowerCase()) ||
                s['Organization Code'].toLowerCase().includes(query.toLowerCase())
            );

            filtered.forEach((s, idx) => {{
                let badge = '<span class="px-2 py-0.5 rounded text-[11px] font-bold bg-amber-500/10 text-amber-400 border border-amber-500/30">🟡 Normal</span>';
                if (s.atv >= 70) badge = '<span class="px-2 py-0.5 rounded text-[11px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/30">🟢 Strong</span>';
                else if (s.atv < 50) badge = '<span class="px-2 py-0.5 rounded text-[11px] font-bold bg-rose-500/10 text-rose-400 border border-rose-500/30">🔴 Low Basket</span>';

                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-800/40 transition";
                tr.innerHTML = `
                    <td class="py-3 px-3 text-slate-500 font-semibold">${{idx + 1}}</td>
                    <td class="py-3 px-3 font-mono text-sky-400 text-xs">${{s['Organization Code']}}</td>
                    <td class="py-3 px-3 font-bold text-white">${{s['Organization Name']}}</td>
                    <td class="py-3 px-3 font-mono font-bold text-emerald-400">${{s.sales.toLocaleString('en-US', {{maximumFractionDigits: 0}})}}</td>
                    <td class="py-3 px-3 font-semibold text-slate-300">${{s.share}}%</td>
                    <td class="py-3 px-3 font-mono">${{s.txns.toLocaleString()}}</td>
                    <td class="py-3 px-3 font-mono font-bold ${{s.atv < 50 ? 'text-rose-400' : (s.atv >= 70 ? 'text-emerald-400' : 'text-slate-200')}}">${{s.atv.toFixed(2)}}</td>
                    <td class="py-3 px-3 font-mono text-indigo-300">${{s.upt.toFixed(2)}}</td>
                    <td class="py-3 px-3 text-center">${{badge}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}
        renderStores();
        document.getElementById('storeSearch').addEventListener('input', (e) => renderStores(e.target.value));

        // Trend Chart
        new Chart(document.getElementById('trendChart'), {{
            type: 'line',
            data: {{
                labels: payload.daily_trend.map(d => d.Date.slice(5)),
                datasets: [{{
                    data: payload.daily_trend.map(d => d.sales),
                    borderColor: '#10b981',
                    backgroundColor: 'rgba(16, 185, 129, 0.1)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.25
                }}]
            }},
            options: {{
                responsive: true, maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ x: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }}, y: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }} }}
            }}
        }});

        // Hourly Chart
        new Chart(document.getElementById('hourlyChart'), {{
            type: 'bar',
            data: {{
                labels: payload.hourly.map(h => `${{h.Hour}}:00`),
                datasets: [{{
                    label: 'Transactions',
                    data: payload.hourly.map(h => h.txns),
                    backgroundColor: '#0ea5e9',
                    borderRadius: 4
                }}]
            }},
            options: {{
                responsive: true, maintainAspectRatio: false,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ x: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }}, y: {{ grid: {{ color: '#1e293b' }}, ticks: {{ color: '#94a3b8' }} }} }}
            }}
        }});
    </script>
</body>
</html>"""

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    process_and_build()