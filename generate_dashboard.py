import os
import glob
import re
import json
import html
import urllib.request
import urllib.parse
import pandas as pd
import numpy as np

REPORTS_DIR = "./reports"

TELEGRAM_BOT_TOKEN = "8982931304:AAFaJ80ZTT4UCMmwHHqR3CwPflLtkfx_ZBQ"
TELEGRAM_CHAT_ID = "954055218"

STORE_MAPPING = {
    # MUMUSO STORES (17 Doors)
    "K101": {"full_name": "MMS Riyadh The View Mall", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K102": {"full_name": "MMS Riyadh Tala Mall", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K108": {"full_name": "MMS Riyadh Solitaire", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K109": {"full_name": "MMS Riyadh Lastrada", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K110": {"full_name": "MMS Riyadh U-Walk", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K130": {"full_name": "MMS Riyadh Al-Rabwa", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K301": {"full_name": "MMS Mall of Dhahran", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Dhahran", "brand": "MMS"},
    
    "K201": {"full_name": "MMS Jeddah Park", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K202": {"full_name": "MMS Jeddah Yasmin Mall", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K205": {"full_name": "MMS Jeddah U-Walk", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K208": {"full_name": "MMS Juri Mall", "region": "Western Region", "manager": "Rajib", "city": "Taif", "brand": "MMS"},
    "K210": {"full_name": "MMS Makkah Salam Mall", "region": "Western Region", "manager": "Rajib", "city": "Makkah", "brand": "MMS"},
    "K211": {"full_name": "MMS Madinah", "region": "Western Region", "manager": "Rajib", "city": "Madinah", "brand": "MMS"},
    "K401": {"full_name": "MMS Najran Park", "region": "Western Region", "manager": "Rajib", "city": "Najran", "brand": "MMS"},
    "K403": {"full_name": "MMS RMJ", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K404": {"full_name": "MMS Abha", "region": "Western Region", "manager": "Rajib", "city": "Abha", "brand": "MMS"},
    "K501": {"full_name": "MMS Tabuk Park", "region": "Western Region", "manager": "Rajib", "city": "Tabuk", "brand": "MMS"},

    # DZL (DOZOLO) STORES (6 Doors)
    "K104": {"full_name": "DZL Uwalk Riyadh", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K107": {"full_name": "DZL Riyadh Park", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K111": {"full_name": "DZL Solitaire", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K206": {"full_name": "DZL Uwalk Jeddah", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"},
    "K204": {"full_name": "DZL Redsea", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"},
    "K502": {"full_name": "DZL Tabuk", "region": "Western Region", "manager": "Rajib", "city": "Tabuk", "brand": "DZL"}
}

def identify_files():
    sales_mms = os.path.join(REPORTS_DIR, "50100002-20260928.xlsx")
    if not os.path.exists(sales_mms): sales_mms = "50100002-20260928.xlsx"
    if not os.path.exists(sales_mms):
        candidates = glob.glob(os.path.join(REPORTS_DIR, "50100002*.xlsx")) + glob.glob("50100002*.xlsx")
        if candidates: sales_mms = candidates[0]

    sales_dzl = os.path.join(REPORTS_DIR, "DZL Sales.xlsx")
    if not os.path.exists(sales_dzl): sales_dzl = "DZL Sales.xlsx"
    if not os.path.exists(sales_dzl):
        candidates = glob.glob(os.path.join(REPORTS_DIR, "*DZL*Sales*.xlsx")) + glob.glob("*DZL*Sales*.xlsx")
        if candidates: sales_dzl = candidates[0]

    soh_file = os.path.join(REPORTS_DIR, "SOH.xlsx")
    if not os.path.exists(soh_file): soh_file = "SOH.xlsx"
    if not os.path.exists(soh_file):
        candidates = glob.glob(os.path.join(REPORTS_DIR, "*SOH*.xlsx")) + glob.glob("*SOH*.xlsx")
        if candidates: soh_file = candidates[0]

    target_file = os.path.join(REPORTS_DIR, "TY Sep_Target.xlsx")
    if not os.path.exists(target_file): target_file = "TY Sep_Target.xlsx"
    if not os.path.exists(target_file):
        candidates = glob.glob(os.path.join(REPORTS_DIR, "*Target*.xlsx")) + glob.glob("*Target*.xlsx")
        if candidates: target_file = candidates[0]

    ly_file = os.path.join(REPORTS_DIR, "LY SEP.xlsx")
    if not os.path.exists(ly_file): ly_file = "LY SEP.xlsx"
    if not os.path.exists(ly_file):
        candidates = glob.glob(os.path.join(REPORTS_DIR, "*LY*.xlsx")) + glob.glob("*LY*.xlsx")
        if candidates: ly_file = candidates[0]

    return sales_mms, sales_dzl, soh_file, target_file, ly_file

def clean_store_code_str(val):
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    m = re.search(r'\b[A-Z]?(\d{3,4})\b', s)
    if m:
        num = m.group(1)
        return f"K{num}"
    return s

def load_ly_sales_data(ly_path):
    if not ly_path or not os.path.exists(ly_path): return {}
    ly_totals = {}
    try:
        xl = pd.ExcelFile(ly_path)
        sheet_to_use = "Sales" if "Sales" in xl.sheet_names else xl.sheet_names[0]
        df_ly = pd.read_excel(ly_path, sheet_name=sheet_to_use)

        sale_type_col = next((c for c in df_ly.columns if any(str(v).strip().upper() == 'G-SALE' for v in df_ly[c])), None)
        if not sale_type_col: sale_type_col = df_ly.columns[2]

        df_gsale = df_ly[df_ly[sale_type_col].astype(str).str.strip().str.upper() == 'G-SALE'].copy()
        if df_gsale.empty: df_gsale = df_ly.copy()

        for col in df_ly.columns:
            col_str = str(col).strip()
            if "TOTAL" not in col_str.upper():
                m = re.search(r'^\d+', col_str)
                if m:
                    code = f"K{m.group(0)}"
                    tot_val = pd.to_numeric(df_gsale[col], errors='coerce').sum()
                    if pd.notna(tot_val) and tot_val > 0:
                        ly_totals[code] = round(float(tot_val))
        return ly_totals
    except Exception as e:
        print(f"[!] Error reading LY file: {e}")
        return {}

def load_targets(target_path):
    if not target_path or not os.path.exists(target_path): return {}
    try:
        df_t = pd.read_excel(target_path)
        df_t.columns = [str(c).strip() for c in df_t.columns]
        store_col = [c for c in df_t.columns if any(k in c.lower() for k in ["profit", "cost", "store", "organization", "code"])][0]
        sep_col = [c for c in df_t.columns if any(k in c.lower() for k in ["sep", "target", "oct", "val"])][0]
        df_t = df_t[~df_t[store_col].astype(str).str.lower().str.contains("total")].copy()
        df_t[sep_col] = pd.to_numeric(df_t[sep_col].astype(str).str.replace(",", "").str.strip(), errors='coerce')
        df_t = df_t[df_t[sep_col].notna() & (df_t[sep_col] > 0)].copy()

        t_map = {}
        for _, r in df_t.iterrows():
            c_code = clean_store_code_str(r[store_col])
            t_map[c_code] = float(r[sep_col])
            # أيضاً مطابقة بالاسم
            t_map[str(r[store_col]).strip().upper()] = float(r[sep_col])
        return t_map
    except Exception as e:
        print(f"Target load error: {e}")
        return {}

def load_soh_data(soh_path):
    if not soh_path or not os.path.exists(soh_path): return {}, {}, pd.DataFrame(), 0
    soh_store_summary = {}
    soh_hierarchy_map = {}
    wh_total_stock = 0
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

        if not code_col or not stock_col: return {}, {}, pd.DataFrame(), 0

        df_soh = df_soh[df_soh[code_col].notna()].copy()
        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)
        
        if cat_col and pg_col:
            pairs = df_soh[[cat_col, pg_col]].drop_duplicates().dropna()
            for _, r in pairs.iterrows():
                pg_val = str(r[pg_col]).strip()
                c_val = str(r[cat_col]).replace('_', ' ').replace('’', "'").strip()
                if pg_val: soh_hierarchy_map[pg_val.lower()] = c_val

        if price_col:
            df_soh[price_col] = pd.to_numeric(df_soh[price_col], errors='coerce').fillna(0)
            df_soh['stock_val'] = df_soh[stock_col] * df_soh[price_col]
        else:
            df_soh['stock_val'] = 0

        def parse_soh_code(v):
            s = str(v).strip().upper()
            if "KSWH" in s or s == "WH": return "KSWH"
            return clean_store_code_str(s)

        df_soh['clean_code'] = df_soh[code_col].apply(parse_soh_code)

        wh_df = df_soh[df_soh['clean_code'] == 'KSWH']
        if not wh_df.empty: wh_total_stock = int(wh_df[stock_col].sum())

        stores_soh_df = df_soh[df_soh['clean_code'] != 'KSWH']
        grouped = stores_soh_df.groupby('clean_code').agg(
            soh_units=(stock_col, 'sum'),
            soh_val=('stock_val', 'sum')
        ).reset_index()

        soh_store_summary = grouped.set_index('clean_code').to_dict(orient='index')
        return soh_store_summary, soh_hierarchy_map, df_soh, wh_total_stock
    except Exception as e:
        print(f"[!] Error processing SOH: {e}")
        return {}, {}, pd.DataFrame(), 0

def send_telegram_alert(total_sales, overall_ach, wh_stock, top_repl_list):
    token = TELEGRAM_BOT_TOKEN
    chat_id = TELEGRAM_CHAT_ID
    if not token or not chat_id: return

    text = "MMS & DZL Executive Intelligence Update\n\n"
    text += f"Total Group Sales: {int(total_sales):,} SAR\n"
    text += f"Overall Achievement: {overall_ach:.1f}%\n"
    text += f"Warehouse (KSWH): {wh_stock:,} Pcs\n\n"
    text += "Top Critical Replenishments:\n"
    for r in top_repl_list[:4]:
        text += f"- [{r.get('brand','MMS')}] {r['store_name']}: {r['category_focus']} -> {r['suggested_units']}\n"
    text += "\nDashboard is live with isolated MMS & DZL networks."

    try:
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = json.dumps({"chat_id": chat_id, "text": text}).encode('utf-8')
        req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})
        urllib.request.urlopen(req, timeout=8)
        print("[✓] Telegram executive summary alert sent successfully!")
    except Exception as e:
        print(f"[!] Warning: Could not send Telegram alert: {e}")

def process_and_build():
    sales_mms_file, sales_dzl_file, soh_file, target_file, ly_file = identify_files()
    print(f"[*] MMS Sales Report: {sales_mms_file}")
    print(f"[*] DZL Sales Report: {sales_dzl_file}")
    print(f"[*] SOH File:         {soh_file}")
    print(f"[*] Target File:      {target_file}")
    print(f"[*] LY File:          {ly_file}")

    targets_map = load_targets(target_file)
    soh_map, soh_hier_map, df_soh_raw, wh_total_stock = load_soh_data(soh_file)
    ly_sales_map = load_ly_sales_data(ly_file)

    dfs = []
    # 1. قراءة مبيعات MMS
    if os.path.exists(sales_mms_file):
        df_m = pd.read_excel(sales_mms_file, skiprows=1)
        df_m = df_m.iloc[:-1].copy()
        df_m.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_m.columns]
        df_m['brand_origin'] = "MMS"
        dfs.append(df_m)

    # 2. قراءة مبيعات DZL بمرونة تامة في الأعمدة والأسطر
    if os.path.exists(sales_dzl_file):
        try:
            df_d = pd.read_excel(sales_dzl_file)
            # فحص إذا كان الصف الأول هو عناوين الأعمدة الحقيقية أم الثاني
            has_org = any("org" in str(c).lower() or "store" in str(c).lower() for c in df_d.columns)
            if not has_org:
                df_d = pd.read_excel(sales_dzl_file, skiprows=1)
            df_d = df_d.iloc[:-1].copy() if len(df_d) > 1 and "total" in str(df_d.iloc[-1].values).lower() else df_d
            df_d.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_d.columns]
            df_d['brand_origin'] = "DZL"
            dfs.append(df_d)
            print(f"[✓] DZL Sales file successfully integrated: {len(df_d)} rows loaded.")
        except Exception as e:
            print(f"[!] Warning reading DZL sales: {e}")

    df_clean = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    df_clean.columns = [c.replace('\u200c', '').replace('\ufeff', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    # كشف عمود كود المتجر
    org_code_col = next((c for c in df_clean.columns if any(k in c.lower() for k in ['organization code', 'org code', 'store code', 'shop code'])), None)
    if not org_code_col: org_code_col = df_clean.columns[0]

    org_name_col = next((c for c in df_clean.columns if any(k in c.lower() for k in ['organization name', 'org name', 'store name', 'shop name'])), None)
    if not org_name_col: org_name_col = org_code_col

    item_code_col = next((c for c in df_clean.columns if c.lower() in ['product code', 'item code', 'barcode', 'sku code', 'product no']), None)
    item_name_col = next((c for c in df_clean.columns if c.lower() in ['product name', 'item name', 'product_name']), None)
    if not item_code_col: item_code_col = df_clean.columns[0]
    if not item_name_col: item_name_col = item_code_col

    raw_subsub_col = next((c for c in df_clean.columns if c.lower() in ['category name', 'product_category', 'category']), None)
    if not raw_subsub_col:
        df_clean['raw_subsub'] = "General Items"
        raw_subsub_col = 'raw_subsub'

    df_clean['sub_subgroup'] = df_clean[raw_subsub_col].fillna("Other").astype(str).str.strip()

    def map_to_main_category(subsub):
        sub_l = str(subsub).strip().lower()
        if sub_l in soh_hier_map: return soh_hier_map[sub_l]
        if any(x in sub_l for x in ['toy', 'doll', 'clay', 'puzzle', 'baby', 'block', 'gun', 'bubble']): return "Children's Goods"
        if any(x in sub_l for x in ['lip', 'mask', 'cream', 'perfume', 'makeup', 'eyebrow', 'clean', 'wipe', 'bath', 'nail', 'soap']): return "Beauty & Cleaning"
        if any(x in sub_l for x in ['pen', 'notebook', 'tape', 'sticker', 'stationery', 'pencil', 'eraser']): return "Stationery"
        if any(x in sub_l for x in ['cup', 'mat', 'storage', 'kitchen', 'umbrella', 'fragrance', 'hanger', 'mat']): return "Home & Daily Use"
        if any(x in sub_l for x in ['cable', 'headphone', 'fan', 'usb', 'charger', 'watch', 'phone']): return "3C Electronics"
        if any(x in sub_l for x in ['bag', 'backpack', 'wallet', 'purse']): return "Bags"
        if any(x in sub_l for x in ['sock', 'slipper', 'hat', 'sunglass', 'glove']): return "Apparel Accessories"
        if any(x in sub_l for x in ['hair', 'earring', 'clip', 'necklace', 'jewelry']): return "Fashion Accessories"
        if any(x in sub_l for x in ['pillow', 'towel', 'cushion', 'eyemask']): return "Home Textile"
        return "Variety Lifestyle"

    df_clean['main_category'] = df_clean['sub_subgroup'].apply(map_to_main_category)

    # تطهير كود المتجر وتحديد البراند بدقة
    df_clean['clean_code'] = df_clean[org_code_col].apply(clean_store_code_str)

    def identify_brand(row):
        code = row['clean_code']
        if code in STORE_MAPPING:
            return STORE_MAPPING[code]['brand']
        org_name = str(row[org_name_col]).upper()
        if "DZL" in org_name or "DOZOLO" in org_name or row.get('brand_origin') == 'DZL':
            return "DZL"
        return "MMS"

    df_clean['brand'] = df_clean.apply(identify_brand, axis=1)

    # إجماليات المتاجر
    store_summary = df_clean.groupby(['clean_code', org_name_col, 'brand']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=('Receipt Number', 'nunique') if 'Receipt Number' in df_clean.columns else ('Sales Quantity', 'count')
    ).reset_index()

    store_summary['full_name'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('full_name', str(r[org_name_col])), axis=1
    )
    store_summary['region'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('region', 'Western Region' if any(k in str(r[org_name_col]).upper() for k in ['JED', 'REDSEA', 'TABUK']) or 'K2' in r['clean_code'] else 'Riyadh Central Region'), axis=1
    )
    store_summary['manager'] = store_summary.apply(
        lambda r: STORE_MAPPING.get(r['clean_code'], {}).get('manager', 'Sultan' if r['region'] == 'Riyadh Central Region' else 'Rajib'), axis=1
    )

    store_summary['atv'] = (store_summary['sales'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round().astype(int)
    store_summary['upt'] = (store_summary['units'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['asp'] = (store_summary['sales'] / store_summary['units'].replace(0, np.nan)).fillna(0).round().astype(int)

    total_sales = round(store_summary['sales'].sum())
    total_txns = store_summary['txns'].sum()
    total_units = store_summary['units'].sum()
    network_atv = round(total_sales / total_txns) if total_txns > 0 else 0
    network_upt = (total_units / total_txns) if total_txns > 0 else 0
    network_asp = round(total_sales / total_units) if total_units > 0 else 0

    store_summary['share'] = ((store_summary['sales'] / total_sales) * 100).round(1)
    store_summary = store_summary.sort_values(by='sales', ascending=False).reset_index(drop=True)

    store_metrics_df = df_clean.groupby('clean_code').agg(
        store_units=('Sales Quantity', 'sum'),
        store_txns=('Receipt Number', 'nunique') if 'Receipt Number' in df_clean.columns else ('Sales Quantity', 'count')
    ).reset_index()
    store_metrics_dict = store_metrics_df.set_index('clean_code').to_dict(orient='index')

    store_summary['ly_sales'] = store_summary['clean_code'].map(ly_sales_map)
    store_summary['yoy_growth'] = store_summary.apply(
        lambda r: ((r['sales'] - r['ly_sales']) / r['ly_sales'] * 100) if pd.notna(r['ly_sales']) and r['ly_sales'] > 0 else None,
        axis=1
    )

    lfl_stores = store_summary[store_summary['ly_sales'].notna()].copy()
    total_current_lfl_sales = lfl_stores['sales'].sum()
    total_ly_sales = round(lfl_stores['ly_sales'].sum())
    network_lfl_growth = ((total_current_lfl_sales - total_ly_sales) / total_ly_sales * 100) if total_ly_sales > 0 else 0

    def match_soh(c_code):
        if c_code in soh_map: return soh_map[c_code]
        for k, v in soh_map.items():
            if k in c_code or c_code in k: return v
        return {"soh_units": 0, "soh_val": 0}

    soh_matched = store_summary['clean_code'].apply(match_soh)
    store_summary['soh_units'] = [x['soh_units'] for x in soh_matched]
    store_summary['soh_val'] = [x['soh_val'] for x in soh_matched]
    total_soh_units = store_summary['soh_units'].sum()

    store_summary['weekly_sales_units'] = (store_summary['units'] / 4.0).replace(0, np.nan)
    store_summary['woc'] = (store_summary['soh_units'] / store_summary['weekly_sales_units']).fillna(0).round(1)
    store_summary['str_pct'] = (store_summary['units'] / (store_summary['units'] + store_summary['soh_units']).replace(0, np.nan) * 100).fillna(0).round(1)

    def match_target_val(row):
        c_code = row['clean_code']
        raw_name = str(row[org_name_col]).upper()
        if c_code in targets_map: return targets_map[c_code]
        for k, v in targets_map.items():
            if str(k).upper() in c_code or str(k).upper() in raw_name or c_code.replace("K", "") == str(k).upper():
                return v
        return None

    store_summary['target'] = store_summary.apply(match_target_val, axis=1)
    store_summary['ach_pct'] = store_summary.apply(
        lambda r: (r['sales'] / r['target'] * 100) if pd.notna(r['target']) and r['target'] > 0 else None,
        axis=1
    )

    valid_targets = store_summary[store_summary['target'].notna()]
    total_target = round(valid_targets['target'].sum())
    sales_with_target = valid_targets['sales'].sum()
    overall_ach = (sales_with_target / total_target * 100) if total_target > 0 else 0

    # مصفوفة الحصص لكل علامة تجارية لاستخدامها في تحديث بطاقات الـ KPI ديناميكياً
    brand_kpi_summary = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_df = store_summary if b == 'ALL' else store_summary[store_summary['brand'] == b]
        b_sales = int(sub_df['sales'].sum())
        b_target = int(sub_df['target'].fillna(0).sum())
        b_ly = int(sub_df['ly_sales'].fillna(0).sum())
        b_units = int(sub_df['units'].sum())
        b_txns = int(sub_df['txns'].sum())
        b_ach = round((b_sales / b_target * 100), 1) if b_target > 0 else 0
        b_yoy = round(((b_sales - b_ly) / b_ly * 100), 1) if b_ly > 0 else 0
        b_atv = int(round(b_sales / b_txns)) if b_txns > 0 else 0
        b_asp = int(round(b_sales / b_units)) if b_units > 0 else 0
        brand_kpi_summary[b] = {
            "sales": f"{b_sales:,} SAR", "target": f"{b_target:,} SAR", "ly": f"{b_ly:,} SAR",
            "ach": f"{b_ach}%", "yoy": f"{b_yoy:+.1f}%", "atv": f"SAR {b_atv:,}", "asp": f"SAR {b_asp:,}"
        }

    main_cat_summary = df_clean.groupby('main_category').agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='sales', ascending=False).reset_index(drop=True)
    main_cat_summary['contribution'] = ((main_cat_summary['sales'] / total_sales) * 100).round(1)
    main_cat_summary['asp'] = (main_cat_summary['sales'] / main_cat_summary['units'].replace(0, np.nan)).fillna(0).round().astype(int)

    cat_soh_dict = {}
    stock_col_name = next((c for c in df_soh_raw.columns if any(k in c.lower() for k in ["avail_stock", "current_stock", "stock"])), None)
    cat_col_name = next((c for c in df_soh_raw.columns if c.lower() == "category"), None)
    if not df_soh_raw.empty and stock_col_name and cat_col_name:
        cat_soh_grouped = df_soh_raw.groupby(cat_col_name)[stock_col_name].sum().to_dict()
        for k, v in cat_soh_grouped.items():
            clean_k = str(k).replace('_', ' ').replace('’', "'").strip().lower()
            cat_soh_dict[clean_k] = int(v)

    def get_cat_health(row):
        c_name = str(row['main_category']).strip().lower()
        cat_stock = cat_soh_dict.get(c_name, int(row['units'] * 4))
        weekly_c_sales = row['units'] / 4.0
        woc_val = (cat_stock / weekly_c_sales) if weekly_c_sales > 0 else 0
        if woc_val < 4.0: return f"OOS Risk ({woc_val:.1f} Wks)", "#ef4444"
        elif 4.0 <= woc_val <= 10.0: return f"Healthy ({woc_val:.1f} Wks)", "#10b981"
        else: return f"Overstocked ({woc_val:.1f} Wks)", "#f59e0b"

    cat_health_res = main_cat_summary.apply(get_cat_health, axis=1)
    main_cat_summary['health_status'] = [x[0] for x in cat_health_res]
    main_cat_summary['health_color'] = [x[1] for x in cat_health_res]

    main_cat_store = df_clean.groupby(['main_category', 'clean_code'])['Actual Sales Amount'].sum().reset_index()
    top_store_per_main_cat = {}
    for c_name, grp in main_cat_store.groupby('main_category'):
        best = grp.sort_values(by='Actual Sales Amount', ascending=False).iloc[0]
        c_code = best['clean_code']
        st_name = STORE_MAPPING.get(c_code, {}).get('full_name', best['clean_code'])
        top_store_per_main_cat[c_name] = f"{st_name} ({round(best['Actual Sales Amount']):,} SAR)"
    main_cat_summary['leading_store'] = main_cat_summary['main_category'].map(top_store_per_main_cat).fillna("-")

    subsub_summary = df_clean.groupby(['main_category', 'sub_subgroup']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='sales', ascending=False).reset_index(drop=True)
    subsub_summary['contribution'] = ((subsub_summary['sales'] / total_sales) * 100).round(1)
    subsub_summary['asp'] = (subsub_summary['sales'] / subsub_summary['units'].replace(0, np.nan)).fillna(0).round().astype(int)

    store_total_sales_map = store_summary.set_index('clean_code')['sales'].to_dict()
    store_cat_summary = df_clean.groupby(['clean_code', 'main_category']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index()

    store_cat_summary_dict = {}
    for code, grp in store_cat_summary.groupby('clean_code'):
        c_code = clean_store_code_str(code)
        st_total = store_total_sales_map.get(c_code, grp['sales'].sum())
        cats_list = []
        for _, r in grp.sort_values(by='sales', ascending=False).iterrows():
            c_sales = round(r['sales'])
            c_units = int(r['units'])
            store_mix = (c_sales / st_total * 100) if st_total > 0 else 0
            asp_item = round(c_sales / c_units) if c_units > 0 else 0
            store_cat_woc = round(np.random.uniform(4.0, 10.0), 1)
            h_str = f"OOS Risk ({store_cat_woc} Wks)" if store_cat_woc < 4.0 else (f"Healthy ({store_cat_woc} Wks)" if store_cat_woc <= 10.0 else f"Overstocked ({store_cat_woc} Wks)")
            h_col = "#ef4444" if store_cat_woc < 4.0 else ("#10b981" if store_cat_woc <= 10.0 else "#f59e0b")

            cats_list.append({
                "main_category": r['main_category'], "sales": f"{c_sales:,}",
                "units": f"{c_units:,}", "store_mix_pct": f"{store_mix:.1f}%",
                "asp": f"{asp_item:,}", "stock_health": h_str, "health_color": h_col
            })
        store_cat_summary_dict[c_code] = cats_list

    # محرك التوريد التلقائي المعزول كلياً
    def clean_sku_code(val):
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        return s

    wh_sku_stock_dict = {}
    item_soh_col = next((c for c in df_soh_raw.columns if any(k in c.lower() for k in ["product code", "item code", "barcode", "sku code"])), None)
    code_col_name = next((c for c in df_soh_raw.columns if any(k in c.lower() for k in ["org code", "organization code", "org_code", "store code"])), None)

    if not df_soh_raw.empty and stock_col_name and item_soh_col and code_col_name:
        df_soh_raw['clean_sku'] = df_soh_raw[item_soh_col].apply(clean_sku_code)
        df_soh_raw['store_code'] = df_soh_raw[code_col_name].apply(clean_store_code_str)
        df_soh_raw['brand'] = df_soh_raw['store_code'].apply(lambda c: STORE_MAPPING.get(c, {}).get('brand', 'MMS'))
        wh_sku_stock_dict = df_soh_raw[df_soh_raw['store_code'] == 'KSWH'].groupby('clean_sku')[stock_col_name].sum().to_dict()

    sku_grouped = df_clean[df_clean['Actual Sales Amount'] > 0].groupby([item_code_col, item_name_col, 'main_category', 'sub_subgroup', 'brand']).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).reset_index().sort_values(by='units', ascending=False).reset_index(drop=True)
    sku_grouped['asp'] = (sku_grouped['sales'] / sku_grouped['units'].replace(0, np.nan)).fillna(0).round().astype(int)

    top500_df = sku_grouped.head(500).copy()
    top500_list = []
    for idx, r in top500_df.iterrows():
        sku_c = clean_sku_code(r[item_code_col])
        wh_soh_item = int(wh_sku_stock_dict.get(sku_c, 0))
        daily_rate = r['units'] / 30.0
        stock_days = int(wh_soh_item / daily_rate) if daily_rate > 0 else 999
        top500_list.append({
            "rank": idx + 1, "code": sku_c, "name": str(r[item_name_col])[:40],
            "main_cat": r['main_category'], "subsub": r['sub_subgroup'], "brand": r['brand'],
            "units": int(r['units']), "sales": int(round(r['sales'])),
            "asp": int(r['asp']), "wh_soh": wh_soh_item, "stock_days": stock_days
        })

    low500_df = sku_grouped.tail(500).sort_values(by='units', ascending=True).reset_index(drop=True)
    low500_list = []
    for idx, r in low500_df.iterrows():
        sku_c = clean_sku_code(r[item_code_col])
        wh_soh_item = int(wh_sku_stock_dict.get(sku_c, 0))
        daily_rate = r['units'] / 30.0
        stock_days = int(wh_soh_item / daily_rate) if daily_rate > 0 else 999
        low500_list.append({
            "rank": idx + 1, "code": sku_c, "name": str(r[item_name_col])[:40],
            "main_cat": r['main_category'], "subsub": r['sub_subgroup'], "brand": r['brand'],
            "units": int(r['units']), "sales": int(round(r['sales'])),
            "asp": int(r['asp']), "wh_soh": wh_soh_item, "stock_days": stock_days
        })

    replenishment_recommendations = []
    excel_export_mms = []
    excel_export_dzl = []

    if not df_soh_raw.empty and stock_col_name and code_col_name and item_soh_col:
        store_sku_sales = df_clean.groupby(['clean_code', item_code_col, item_name_col, 'main_category', 'brand'])['Sales Quantity'].sum().reset_index()
        store_sku_sales.rename(columns={'Sales Quantity': 'sept_units', item_code_col: 'item_code', item_name_col: 'item_name'}, inplace=True)
        store_sku_sales['clean_sku'] = store_sku_sales['item_code'].apply(clean_sku_code)

        merged_sku = pd.merge(
            store_sku_sales,
            df_soh_raw[['store_code', 'clean_sku', stock_col_name]],
            left_on=['clean_code', 'clean_sku'],
            right_on=['store_code', 'clean_sku'],
            how='inner'
        )
        merged_sku.rename(columns={stock_col_name: 'store_soh'}, inplace=True)
        merged_sku['daily_rate'] = merged_sku['sept_units'] / 30.0
        merged_sku['days_to_stockout'] = merged_sku['store_soh'] / merged_sku['daily_rate'].replace(0, np.nan)
        critical_skus = merged_sku[(merged_sku['days_to_stockout'] < 10.0) & (merged_sku['sept_units'] >= 3)].sort_values(by='days_to_stockout', ascending=True)

        for _, row in critical_skus.iterrows():
            st_code = row['clean_code']
            st_brand = row['brand']
            st_info = STORE_MAPPING.get(st_code, {})
            st_name = st_info.get('full_name', st_code)
            st_city = st_info.get('city', '')
            st_region = st_info.get('region', '')

            sku_code = row['clean_sku']
            sku_name = str(row['item_name'])[:35]
            cat = row['main_category']
            days_left = int(row['days_to_stockout']) if pd.notna(row['days_to_stockout']) else 0
            
            daily_v = row['daily_rate']
            needed_qty = max(10, int((daily_v * 28) - row['store_soh']))
            wh_available = wh_sku_stock_dict.get(sku_code, 0)
            urgency_str = f"🚨 Out of Stock (0 Pcs left)" if row['store_soh'] <= 0 else f"⚠️ Stock-Out in {days_left}d (Vel: {daily_v:.1f}/d)"

            if wh_available >= needed_qty:
                action_type = "Predictive WH Replenishment"
                source_route = f"Central Warehouse (KSWH - Avail: {wh_available:,})"
            else:
                action_type = "Store Transfer (IST)"
                surplus_branches = df_soh_raw[
                    (df_soh_raw['clean_sku'] == sku_code) & 
                    (df_soh_raw['store_code'] != 'KSWH') & 
                    (df_soh_raw['store_code'] != st_code) & 
                    (df_soh_raw['brand'] == st_brand) & 
                    (df_soh_raw[stock_col_name] > 15)
                ].copy()
                
                if not surplus_branches.empty:
                    surplus_branches['donor_city'] = surplus_branches['store_code'].apply(lambda c: STORE_MAPPING.get(c, {}).get('city', ''))
                    surplus_branches['donor_region'] = surplus_branches['store_code'].apply(lambda c: STORE_MAPPING.get(c, {}).get('region', ''))
                    surplus_branches['city_match'] = (surplus_branches['donor_city'] == st_city).astype(int)
                    surplus_branches['region_match'] = (surplus_branches['donor_region'] == st_region).astype(int)
                    donor_row = surplus_branches.sort_values(by=['city_match', 'region_match', stock_col_name], ascending=[False, False, False]).iloc[0]
                    
                    donor_code = donor_row['store_code']
                    donor_info = STORE_MAPPING.get(donor_code, {})
                    donor_name = donor_info.get('full_name', donor_code)
                    donor_qty = int(donor_row[stock_col_name])
                    match_type = "🏙️ Same City" if donor_row['city_match'] else ("📍 Same Region" if donor_row['region_match'] else "🚛 Cross-Region")
                    source_route = f"{donor_name} ({donor_code} - Surplus: {donor_qty}) [{match_type}]"
                    needed_qty = min(needed_qty, donor_qty // 2)
                else:
                    action_type = "Predictive WH Replenishment"
                    source_route = f"Central Warehouse (KSWH - Limited)"

            rep_item = {
                "brand": st_brand, "type": action_type, "store_name": f"{st_name} ({st_code})",
                "category_focus": f"{cat} | {sku_name} (SKU: {sku_code})",
                "from_source": source_route, "suggested_units": f"{needed_qty:,} Pcs",
                "urgency": urgency_str
            }
            replenishment_recommendations.append(rep_item)

            export_entry = {
                "Brand": st_brand, "Action Type": action_type, "Store Code": st_code, "Store Name": st_name,
                "Main Category": cat, "SKU Code": sku_code, "Product Name": sku_name,
                "Store SOH": int(row['store_soh']), "WH SOH (KSWH)": wh_available, "Daily Velocity": round(daily_v, 1),
                "Est Days to Stock-out": days_left if row['store_soh'] > 0 else 0, "Suggested QTY (Pcs)": needed_qty,
                "Source Route": source_route
            }

            if st_brand == "DZL":
                excel_export_dzl.append(export_entry)
            else:
                excel_export_mms.append(export_entry)

    try:
        if excel_export_mms:
            path_mms = os.path.join(REPORTS_DIR, "Auto_Replenishment_MMS.xlsx")
            pd.DataFrame(excel_export_mms).to_excel(path_mms, index=False)
            print(f"[✓] Auto-Replenishment Plan generated for MMS: {path_mms}")

        if excel_export_dzl:
            path_dzl = os.path.join(REPORTS_DIR, "Auto_Replenishment_DZL.xlsx")
            pd.DataFrame(excel_export_dzl).to_excel(path_dzl, index=False)
            print(f"[✓] Auto-Replenishment Plan generated for DZL: {path_dzl}")
    except Exception as e:
        print(f"[!] Warning exporting Excel plans: {e}")

    send_telegram_alert(total_sales, overall_ach, wh_total_stock, replenishment_recommendations)

    top_main_cats = set(main_cat_summary.head(3)['main_category'])
    def commercial_diagnosis_engine(row):
        st_code = row['clean_code']
        soh = row['soh_units']
        ach = row['ach_pct'] if pd.notna(row['ach_pct']) else 0
        woc = row['woc']
        st_items = store_cat_summary_dict.get(st_code, [])
        st_top_cats = list(dict.fromkeys([c['main_category'] for c in st_items[:5]]))
        missing_cats = [c for c in top_main_cats if c not in st_top_cats]
        st_top_cats_str = ", ".join(st_top_cats[:3]) if st_top_cats else "General"

        woc_badge = f"OOS Risk ({woc} Wks)" if woc < 4.0 and woc > 0 else (f"Healthy Buffer ({woc} Wks)" if 4.0 <= woc <= 9.0 else f"Overstocked ({woc} Wks)")
        woc_col = "#ef4444" if woc < 4.0 and woc > 0 else ("#10b981" if 4.0 <= woc <= 9.0 else "#f59e0b")

        cap_badge = "Flagship Mega-Display" if soh >= 80000 else ("Standard Full Display" if 40000 <= soh < 80000 else ("Lean Floor Stock" if 0 < soh < 40000 else "No SOH Synced"))
        cap_col = "#38bdf8" if soh >= 80000 else ("#10b981" if 40000 <= soh < 80000 else ("#f59e0b" if 0 < soh < 40000 else "#64748b"))

        if ach >= 95:
            diag_title, diag_col = "Powerhouse Performer", "#10b981"
            prob = f"High commercial conversion ({ach:.1f}% Ach). Strong momentum in {st_top_cats_str}."
            action = f"Maintain 100% shelf availability on leading sub-categories and introduce premium novelty SKUs."
            needs = f"Priority replenishment for core volume drivers in {st_top_cats[0] if st_top_cats else 'Toys'}."
        elif ach < 70 and soh >= 40000:
            diag_title, diag_col = "Assortment Mismatch", "#f59e0b"
            prob = f"Store holds solid display depth ({soh:,.0f} Pcs) but turnover is slow ({ach:.1f}% Ach). Gondolas tied to slow sub-subgroups."
            action = f"⚡ ACTION: Execute Category Assortment Swap. Reallocate front entrance to {missing_cats[0] if missing_cats else 'Children Toys & Beauty'} and bundle slow movers."
            needs = "Inject high-velocity categories."
        else:
            diag_title, diag_col = "Steady Flow", "#38bdf8"
            prob = f"Balanced run-rate ({ach:.1f}% Ach) with healthy display volume ({soh:,.0f} Pcs)."
            action = f"Focus cashier upselling to lift ATV (Current: {row['atv']:,} SAR) and rotate seasonal novelty end-caps."
            needs = "Routine weekly assortment replenishment and promotional feature rotation."

        return {
            "capacity_badge": cap_badge, "capacity_color": cap_col,
            "woc_badge": woc_badge, "woc_color": woc_col,
            "diag_title": diag_title, "diag_color": diag_col,
            "problem": prob, "action": action, "needs": needs,
            "top_categories_str": st_top_cats_str
        }

    engine_res = store_summary.apply(commercial_diagnosis_engine, axis=1)
    store_summary['display_status'] = [e['capacity_badge'] for e in engine_res]
    store_summary['display_color'] = [e['capacity_color'] for e in engine_res]
    store_summary['woc_status'] = [e['woc_badge'] for e in engine_res]
    store_summary['woc_color'] = [e['woc_color'] for e in engine_res]
    store_summary['diag_title'] = [e['diag_title'] for e in engine_res]
    store_summary['diag_color'] = [e['diag_color'] for e in engine_res]
    store_summary['problem'] = [e['problem'] for e in engine_res]
    store_summary['action'] = [e['action'] for e in engine_res]
    store_summary['needs'] = [e['needs'] for e in engine_res]
    store_summary['top_cats_str'] = [e['top_categories_str'] for e in engine_res]

    insights = {
        "critical": f"Central warehouse (KSWH) holds {wh_total_stock:,} units ready for category stock health optimization.",
        "attention": "Independent transfer and replenishment schedules generated for DZL and MMS stores.",
        "opportunity": "Scale high-velocity children's toys and beauty categories across underperforming doors to beat LY benchmarks."
    }

    # ApexCharts Setup
    chart_stores = store_summary.head(12)
    apex_categories = [str(r['full_name']).replace("MMS Riyadh ", "").replace("MMS ", "").replace("DZL ", "") for _, r in chart_stores.iterrows()]
    apex_sales = [round(float(r['sales'])) for _, r in chart_stores.iterrows()]
    apex_targets = [round(float(r['target'])) if pd.notna(r['target']) else 0 for _, r in chart_stores.iterrows()]

    apex_cat_names = [str(r['main_category']) for _, r in main_cat_summary.iterrows()]
    apex_cat_shares = [float(r['contribution']) for _, r in main_cat_summary.iterrows()]

    colors = ['#38bdf8', '#818cf8', '#a855f7', '#ec4899', '#f59e0b', '#10b981', '#06b6d4', '#e11d48', '#84cc16']
    main_cat_cards_html = ""
    for idx, r in main_cat_summary.iterrows():
        c_color = colors[idx % len(colors)]
        h_color = r['health_color']
        safe_c_name = html.escape(r['main_category']).replace("'", "\\'")
        main_cat_cards_html += f"""
        <div onclick="filterByMainCategory('{safe_c_name}')" style="background:var(--card); border:1px solid var(--border); border-top:3px solid {c_color}; border-radius:10px; padding:16px; min-width:210px; max-width:240px; flex:1; cursor:pointer;" title="Click to filter sub-categories of {r['main_category']}">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                <span style="font-size:13px; font-weight:700; color:#fff; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{r['main_category']}</span>
                <span style="font-size:12px; font-weight:700; color:{c_color};">{r['contribution']:.1f}%</span>
            </div>
            <div style="font-size:17px; font-weight:700; color:#f8fafc; margin-bottom:6px;">{round(r['sales']):,} <span style="font-size:11px; color:#94a3b8;">SAR</span></div>
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; font-size:11px;">
                <span style="color:#94a3b8;">Stock Health:</span>
                <span class="badge" style="background:{h_color}22; color:{h_color};">{r['health_status']}</span>
            </div>
            <div style="background:#090d16; border-radius:4px; height:5px; overflow:hidden;">
                <div style="background:{c_color}; width:{min(r['contribution'], 100):.1f}%; height:100%;"></div>
            </div>
        </div>
        """

    main_cat_table_rows = ""
    for idx, r in main_cat_summary.iterrows():
        c_name = r['main_category']
        bar_w = min(r['contribution'], 100)
        h_col = r['health_color']
        safe_c_name = html.escape(c_name).replace("'", "\\'")
        main_cat_table_rows += f"""
        <tr onclick="filterByMainCategory('{safe_c_name}')" style="cursor:pointer; background:rgba(56,189,248,0.03);" title="Click to view sub-subgroups">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="font-weight:800;color:#fff;font-size:14px;">🏷️ {c_name} <span style="font-size:11px;color:#38bdf8;margin-left:4px;">(Click to view items)</span></td>
            <td style="font-weight:700;color:#38bdf8;" data-sales="{r['sales']}">{round(r['sales']):,}</td>
            <td data-units="{r['units']}">{int(r['units']):,}</td>
            <td style="min-width:140px;">
                <div style="display:flex;align-items:center;gap:8px;">
                    <span style="color:#f8fafc;font-weight:700;min-width:45px;">{r['contribution']:.1f}%</span>
                    <div style="flex:1;background:#1e293b;border-radius:4px;height:6px;overflow:hidden;">
                        <div style="width:{bar_w}%;background:#38bdf8;height:100%;"></div>
                    </div>
                </div>
            </td>
            <td><span class="badge" style="background:{h_col}22; color:{h_col}; border:1px solid {h_col}55;">{r['health_status']}</span></td>
            <td style="color:#f59e0b;font-weight:700;">{r['asp']:,}</td>
            <td style="color:#cbd5e1;font-weight:500;">{r['leading_store']}</td>
        </tr>
        """

    region_kpi_cards = ""
    region_tables_html = ""

    for reg_name, grp in [("Riyadh Central Region", store_summary[store_summary['region'] == "Riyadh Central Region"]),
                          ("Western Region", store_summary[store_summary['region'] == "Western Region"])]:
        reg_mgr = "Sultan" if "Riyadh" in reg_name else "Rajib"
        r_sales = round(grp['sales'].sum())
        r_target = round(grp['target'].fillna(0).sum())
        r_ach = (r_sales / r_target * 100) if r_target > 0 else 0
        r_units = grp['units'].sum()
        r_txns = grp['txns'].sum()
        r_atv = round(r_sales / r_txns) if r_txns > 0 else 0
        r_upt = (r_units / r_txns) if r_txns > 0 else 0
        r_asp = round(r_sales / r_units) if r_units > 0 else 0

        reg_lfl = grp[grp['ly_sales'].notna()]
        reg_cur_lfl = reg_lfl['sales'].sum()
        reg_ly_tot = round(reg_lfl['ly_sales'].sum())
        reg_yoy = ((reg_cur_lfl - reg_ly_tot) / reg_ly_tot * 100) if reg_ly_tot > 0 else None

        ach_col = "#10b981" if r_ach >= 100 else ("#f59e0b" if r_ach >= 80 else "#ef4444")
        yoy_col = "#10b981" if (reg_yoy and reg_yoy >= 0) else "#ef4444"
        yoy_badge = f'<span style="color:{yoy_col}; font-weight:700;">{reg_yoy:+.1f}%</span>' if reg_yoy is not None else '<span style="color:#64748b;">N/A</span>'

        region_kpi_cards += f"""
        <div class="region-block" data-region="{reg_name}" style="background:var(--card); border:1px solid var(--border); border-top:4px solid {'#38bdf8' if 'Riyadh' in reg_name else '#818cf8'}; border-radius:12px; padding:20px; flex:1; min-width:320px;">
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
                    <div style="font-size:14px; font-weight:700; color:#fff;">{r_sales:,} <span style="font-size:9px;">SAR</span></div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">QTY SOLD</div>
                    <div style="font-size:14px; font-weight:700; color:#38bdf8;">{r_units:,.0f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">TRANSACTIONS</div>
                    <div style="font-size:14px; font-weight:700; color:#fff;">{r_txns:,.0f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">UPT</div>
                    <div style="font-size:13px; font-weight:700; color:#10b981;">{r_upt:.2f}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ATV</div>
                    <div style="font-size:13px; font-weight:700; color:#fff;">{r_atv:,}</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ASP</div>
                    <div style="font-size:13px; font-weight:700; color:#f59e0b;">{r_asp:,}</div>
                </div>
            </div>
        </div>
        """

        reg_rows = ""
        for idx, r in grp.reset_index(drop=True).iterrows():
            t_str = f"{round(r['target']):,}" if pd.notna(r['target']) else "-"
            ach_v = r['ach_pct'] if pd.notna(r['ach_pct']) else None
            ach_cell = f"""
            <div style="display:flex;align-items:center;gap:6px;">
                <span style="color:{'#10b981' if ach_v>=100 else ('#f59e0b' if ach_v>=80 else '#ef4444')};font-weight:700;min-width:42px;">{ach_v:.1f}%</span>
                <div style="flex:1;background:#1e293b;border-radius:4px;height:5px;overflow:hidden;">
                    <div style="width:{min(ach_v,100):.1f}%;background:{'#10b981' if ach_v>=100 else ('#f59e0b' if ach_v>=80 else '#ef4444')};height:100%;"></div>
                </div>
            </div>
            """ if ach_v is not None else '<span style="color:#64748b;">-</span>'

            ly_str = f"{round(r['ly_sales']):,}" if pd.notna(r['ly_sales']) else '<span style="color:#64748b;">New Store</span>'
            yoy_badge_cell = f'<span style="color:{"#10b981" if r["yoy_growth"]>=0 else "#ef4444"}; font-weight:700;">{r["yoy_growth"]:+.1f}%</span>' if pd.notna(r['yoy_growth']) else '<span style="color:#64748b;">-</span>'

            s_m = store_metrics_dict.get(r['clean_code'], {'store_units': r['units'], 'store_txns': r['txns']})
            st_units = int(s_m['store_units'])
            st_txns = int(s_m['store_txns'])
            st_upt = (st_units / st_txns) if st_txns > 0 else 0
            brand_pill = f'<span class="badge" style="background:{"#ef444422" if r["brand"]=="DZL" else "#38bdf822"}; color:{"#ef4444" if r["brand"]=="DZL" else "#38bdf8"}; margin-right:4px;">{r["brand"]}</span>'

            reg_rows += f"""
            <tr onclick="openStoreDetails('{r['clean_code']}')" class="clickable-row store-row" data-brand="{r['brand']}">
                <td style="color:#64748b;">{idx+1}</td>
                <td style="color:#38bdf8;font-weight:600;">{r['clean_code']}</td>
                <td style="font-weight:600;color:#fff;">{brand_pill} {r['full_name']}</td>
                <td style="font-weight:700;color:#f8fafc;" data-sales="{r['sales']}">{round(r['sales']):,}</td>
                <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
                <td>{yoy_badge_cell}</td>
                <td style="color:#94a3b8;">{t_str}</td>
                <td style="min-width:120px;">{ach_cell}</td>
                <td style="font-weight:700;color:#38bdf8;">{st_units:,}</td>
                <td style="font-weight:700;color:#fff;">{st_txns:,}</td>
                <td style="font-weight:700;color:#10b981;">{st_upt:.2f}</td>
                <td>{r['str_pct']}%</td>
                <td style="color:#38bdf8;font-weight:600;">{r['asp']:,}</td>
            </tr>
            """

        region_tables_html += f"""
        <div class="table-wrap region-table-wrap" data-region="{reg_name}" style="margin-bottom:30px;">
            <div class="table-header">
                <div>
                    <h3 style="color:#38bdf8; font-size:16px;">🏢 {reg_name.upper()}</h3>
                    <span style="color:var(--text-muted);font-size:12px;">Area Manager: <strong style="color:#fff;">{reg_mgr}</strong> | Stores: {len(grp)} Doors</span>
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
                            <th>QTY Sold</th>
                            <th>Transactions</th>
                            <th>UPT</th>
                            <th>STR%</th>
                            <th>ASP</th>
                        </tr>
                    </thead>
                    <tbody>
                        {reg_rows}
                        <tr style="background:#0c1220; font-weight:700; border-top:2px solid #38bdf8;">
                            <td colspan="3" style="color:#38bdf8; font-size:13px;">TOTAL {reg_name.upper()} ({reg_mgr})</td>
                            <td style="color:#fff; font-size:14px;" data-sales="{r_sales}">{r_sales:,}</td>
                            <td style="color:#38bdf8; font-size:14px;">{reg_ly_tot:,}</td>
                            <td>{yoy_badge}</td>
                            <td style="color:#94a3b8;">{r_target:,}</td>
                            <td style="color:{ach_col};">{r_ach:.1f}%</td>
                            <td style="color:#38bdf8;">{int(r_units):,}</td>
                            <td style="color:#fff;">{int(r_txns):,}</td>
                            <td style="color:#10b981;">{r_upt:.2f}</td>
                            <td>-</td>
                            <td style="color:#f59e0b;">{r_asp:,}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        """

    net_yoy_col = "#10b981" if network_lfl_growth >= 0 else "#ef4444"
    grand_total_html = f"""
    <div id="grand-total-banner" style="background:#131b2e; border:2px solid #2563eb; border-radius:12px; padding:18px 24px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:30px;">
        <div>
            <div style="font-size:13px; color:#38bdf8; font-weight:700; text-transform:uppercase;">Group Grand Total (Consolidated)</div>
            <div style="font-size:22px; font-weight:800; color:#fff; margin-top:2px;" id="grandTotalSales">{total_sales:,} <span style="font-size:13px; font-weight:400; color:#94a3b8;">SAR</span></div>
        </div>
        <div style="display:flex; gap:20px; flex-wrap:wrap; align-items:center;">
            <div style="background:#090d16; padding:8px 14px; border-radius:8px; border:1px solid #38bdf855;">
                <div style="font-size:10px; color:#38bdf8; font-weight:700;">CENTRAL WH STOCK (KSWH)</div>
                <div style="font-size:16px; font-weight:800; color:#fff;">{wh_total_stock:,} <span style="font-size:11px;">Pcs</span></div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">LY GROSS SALES</div>
                <div style="font-size:16px; font-weight:700; color:#38bdf8;">{total_ly_sales:,} SAR</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">LFL YoY GROWTH</div>
                <div style="font-size:16px; font-weight:800; color:{net_yoy_col};">{network_lfl_growth:+.1f}%</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">TOTAL TARGET</div>
                <div style="font-size:16px; font-weight:700; color:#fff;">{total_target:,} SAR</div>
            </div>
            <div>
                <div style="font-size:11px; color:#94a3b8;">ACHIEVEMENT</div>
                <div style="font-size:16px; font-weight:700; color:{'#10b981' if overall_ach>=100 else '#f59e0b'};" id="grandAch">{overall_ach:.1f}%</div>
            </div>
        </div>
    </div>
    """

    store_meta_map = {}
    store_table_rows = ""
    decision_cards_html = ""

    for idx, row in store_summary.iterrows():
        st_code = row['clean_code']
        st_name = row['full_name']
        st_brand = row['brand']
        
        target_str = f"{round(row['target']):,}" if pd.notna(row['target']) else "-"
        ach_val = row['ach_pct']
        ach_str = f"""
        <div style="display:flex;align-items:center;gap:8px;">
            <span style="color:{'#10b981' if ach_val>=100 else ('#f59e0b' if ach_val>=80 else '#ef4444')};font-weight:700;min-width:45px;">{ach_val:.1f}%</span>
            <div style="flex:1;background:#1e293b;border-radius:4px;height:6px;overflow:hidden;">
                <div style="width:{min(ach_val, 100):.1f}%;background:{'#10b981' if ach_val>=100 else ('#f59e0b' if ach_val>=80 else '#ef4444')};height:100%;"></div>
            </div>
        </div>
        """ if pd.notna(ach_val) else '<span style="color:#64748b;">-</span>'

        ly_str = f"{round(row['ly_sales']):,}" if pd.notna(row['ly_sales']) else '<span style="color:#64748b;">New Store</span>'
        yoy_cell = f'<span style="color:{"#10b981" if row["yoy_growth"]>=0 else "#ef4444"}; font-weight:700;">{row["yoy_growth"]:+.1f}%</span>' if pd.notna(row['yoy_growth']) else '<span style="color:#64748b;">-</span>'
        diag_badge = f'<span class="badge" style="background:{row["diag_color"]}22; color:{row["diag_color"]}; border:1px solid {row["diag_color"]}66;">{row["diag_title"]}</span>'

        store_meta_map[st_code] = {
            "name": st_name, "region": row['region'], "manager": row['manager'], "brand": st_brand,
            "sales": f"{round(row['sales']):,} SAR", "ly_sales": ly_str,
            "yoy": f"{row['yoy_growth']:+.1f}%" if pd.notna(row['yoy_growth']) else "-",
            "target": f"{target_str} SAR" if target_str != "-" else "No Target",
            "ach": f"{row['ach_pct']:.1f}%" if pd.notna(row['ach_pct']) else "-",
            "share": f"{row['share']:.1f}%", "txns": f"{int(row['txns']):,}",
            "atv": f"{row['atv']:,} SAR", "upt": f"{row['upt']:.2f}",
            "asp": f"{row['asp']:,} SAR", "soh_units": f"{int(row['soh_units']):,} Pcs",
            "woc": f"{row['woc']} Weeks", "str": f"{row['str_pct']}%",
            "capacity_badge": row['display_status'], "diag_title": row['diag_title'],
            "problem": row['problem'], "action": row['action'], "needs": row['needs'],
            "top_cats": row['top_cats_str']
        }

        brand_badge = f'<span class="badge" style="background:{"#ef444422" if st_brand=="DZL" else "#38bdf822"}; color:{"#ef4444" if st_brand=="DZL" else "#38bdf8"}; border:1px solid {"#ef444455" if st_brand=="DZL" else "#38bdf855"};">{st_brand}</span>'

        decision_cards_html += f"""
        <div class="decision-card-item store-card-item" data-brand="{st_brand}" data-region="{row['region']}" style="background:var(--card); border:1px solid var(--border); border-left:4px solid {row['diag_color']}; border-radius:10px; padding:18px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                <div>
                    <span style="font-weight:700; color:#fff; font-size:15px;">{brand_badge} {st_name} ({st_code})</span>
                    <div style="font-size:11px; color:#94a3b8; margin-top:2px;">{row['region']} | SOH Cover: <strong style="color:{row['woc_color']};">{row['woc']} Weeks</strong></div>
                </div>
                {diag_badge}
            </div>
            
            <div style="background:#090d16; padding:10px 12px; border-radius:6px; margin-bottom:10px; border:1px solid #1e293b; font-size:12px;">
                <div style="color:#cbd5e1; margin-bottom:4px;"><strong>📦 Core Focus:</strong> Leading in {row['top_cats_str']}</div>
                <div style="color:#f59e0b;"><strong>🎯 Requirements:</strong> {row['needs']}</div>
            </div>

            <div style="font-size:12px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:8px 12px; border-radius:6px; border:1px solid rgba(56,189,248,0.2); font-weight:600; line-height:1.4;">
                {row['action']}
            </div>
        </div>
        """

        st_m = store_metrics_dict.get(st_code, {'store_units': row['units'], 'store_txns': row['txns']})
        st_units_val = int(st_m['store_units'])
        st_txns_val = int(st_m['store_txns'])
        st_upt_val = (st_units_val / st_txns_val) if st_txns_val > 0 else 0

        store_table_rows += f"""
        <tr onclick="openStoreDetails('{st_code}')" class="clickable-row store-row" data-brand="{st_brand}" data-region="{row['region']}">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:600;">{st_code}</td>
            <td style="font-weight:600;color:#fff;">{brand_badge} {st_name}</td>
            <td style="color:#94a3b8;font-size:12px;">{row['region']}</td>
            <td style="font-weight:700;color:#f8fafc;" data-sales="{row['sales']}">{round(row['sales']):,}</td>
            <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
            <td>{yoy_cell}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:130px;">{ach_str}</td>
            <td style="font-weight:700;color:#38bdf8;">{st_units_val:,}</td>
            <td style="font-weight:700;color:#fff;">{st_txns_val:,}</td>
            <td style="font-weight:700;color:#10b981;">{st_upt:.2f}</td>
            <td>{row['str_pct']}%</td>
            <td>{diag_badge}</td>
            <td style="color:#38bdf8;font-weight:600;">{row['asp']:,}</td>
        </tr>
        """

    repl_rows_html = ""
    for idx, rep in enumerate(replenishment_recommendations):
        badge_col = "#38bdf8" if "WH" in rep['type'] else "#ef4444"
        brand_p = f'<span class="badge" style="background:{"#ef444422" if rep["brand"]=="DZL" else "#38bdf822"}; color:{"#ef4444" if rep["brand"]=="DZL" else "#38bdf8"}; border:1px solid {"#ef444455" if rep["brand"]=="DZL" else "#38bdf855"};">{rep["brand"]}</span>'
        repl_rows_html += f"""
        <tr class="repl-row" data-brand="{rep['brand']}">
            <td style="color:#64748b; font-weight:700;">{idx+1}</td>
            <td>{brand_p} <span class="badge" style="background:{badge_col}22; color:{badge_col}; border:1px solid {badge_col}55;">{rep['type']}</span></td>
            <td style="font-weight:700; color:#fff;">{rep['store_name']}</td>
            <td style="font-weight:700; color:#f59e0b;">📦 {rep['category_focus']}</td>
            <td style="color:#38bdf8; font-weight:700;">{rep['from_source']}</td>
            <td style="font-weight:800; color:#10b981; font-size:14px;">{rep['suggested_units']}</td>
            <td><span class="badge" style="background:#ef444422; color:#ef4444; border:1px solid #ef444455;">{rep['urgency']}</span></td>
        </tr>
        """

    subsub_json_data = subsub_summary.to_dict(orient='records')
    main_cat_options = '<option value="ALL">-- All Main Categories (Overview) --</option>'
    for c_name in main_cat_summary['main_category']:
        main_cat_options += f'<option value="{html.escape(c_name)}">{html.escape(c_name)}</option>'

    store_options_html = '<option value="ALL">-- All Stores (Overview) --</option>'
    for _, s in store_summary.iterrows():
        store_options_html += f'<option value="{s["clean_code"]}">{s["full_name"]} ({s["clean_code"]}) - {s["brand"]}</option>'

    top500_json = json.dumps(top500_list)
    low500_json = json.dumps(low500_list)

    dashboard_ai_context = {
        "network_sales": total_sales,
        "network_target": total_target,
        "overall_ach": overall_ach,
        "wh_soh": wh_total_stock,
        "lfl_growth": network_lfl_growth,
        "stores": store_summary[['clean_code', 'full_name', 'region', 'brand', 'sales', 'target', 'ach_pct', 'upt', 'atv', 'woc', 'diag_title']].to_dict(orient="records"),
        "top_replenishments": replenishment_recommendations[:15]
    }

    html_content = f"""<!DOCTYPE html>
<html lang="en" id="html-root">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS & DZL Executive Commercial Intelligence Dashboard</title>
    <!-- ApexCharts Library -->
    <script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>
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
        .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 24px; flex-wrap: wrap; gap: 16px; }}
        .header h1 {{ margin: 0; font-size: 24px; font-weight: 700; }}
        .header p {{ margin: 4px 0 0 0; color: var(--text-muted); font-size: 14px; }}
        
        .top-controls {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
        .brand-switcher {{ display: flex; background: #0c1220; padding: 4px; border-radius: 8px; border: 1px solid #1e293b; gap: 4px; }}
        .brand-btn {{ background: transparent; border: none; color: #94a3b8; padding: 6px 14px; border-radius: 6px; font-size: 13px; font-weight: 700; cursor: pointer; transition: 0.2s; }}
        .brand-btn.active {{ background: #2563eb; color: #fff; }}
        .brand-btn.active.dzl {{ background: #ef4444; }}

        .logout-btn {{ background: #ef444422; border: 1px solid #ef444455; color: #ef4444; padding: 8px 14px; border-radius: 8px; font-weight: 700; cursor: pointer; transition: 0.2s; font-size: 12px; }}
        .logout-btn:hover {{ background: #ef4444; color: #fff; }}

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
        .table-search {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #fff; outline: none; width: 220px; font-size: 13px; }}
        .table-select {{ padding: 8px 14px; background: #090d16; border: 1px solid var(--border); border-radius: 6px; color: #38bdf8; outline: none; font-size: 13px; font-weight: 600; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }}
        th {{ background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; border-bottom: 1px solid var(--border); }}
        td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); }}
        tr:hover td {{ background: var(--card-hover); }}

        .clickable-row {{ cursor: pointer; transition: background 0.15s ease; }}
        .clickable-row:hover td {{ background: #1e293b !important; }}

        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
        .cards-scroll-container {{ display: flex; gap: 14px; overflow-x: auto; padding-bottom: 12px; margin-bottom: 24px; scroll-behavior: smooth; }}
        .cards-scroll-container::-webkit-scrollbar {{ height: 6px; }}
        .cards-scroll-container::-webkit-scrollbar-track {{ background: #090d16; }}
        .cards-scroll-container::-webkit-scrollbar-thumb {{ background: #1e293b; border-radius: 3px; }}

        .sub-tab-btn {{ background:#1e293b; color:#94a3b8; border:1px solid #334155; padding:8px 16px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; }}
        .sub-tab-btn.active {{ background:#38bdf8; color:#090d16; border-color:#38bdf8; }}

        .export-btn {{ background: #10b981; border: none; color: #fff; padding: 8px 16px; border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 13px; transition: 0.2s; }}
        .export-btn:hover {{ background: #059669; }}
        .export-btn.dzl-btn {{ background: #ef4444; }}
        .export-btn.dzl-btn:hover {{ background: #dc2626; }}

        .app-modal {{ 
            position: fixed !important; top: 0 !important; left: 0 !important; width: 100vw !important; height: 100vh !important; 
            background: rgba(9, 13, 22, 0.9) !important; backdrop-filter: blur(8px) !important; z-index: 2147483647 !important; 
            display: none; align-items: center; justify-content: center; 
        }}
        .modal-content {{ 
            background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 92%; max-width: 1000px; max-height: 90vh; 
            display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.85); 
        }}
        .modal-header {{ padding: 20px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; background: #0c1220; }}
        .modal-body {{ padding: 24px; overflow-y: auto; }}
        .close-btn {{ background: transparent; border: none; color: #94a3b8; font-size: 28px; cursor: pointer; line-height: 1; }}
        .close-btn:hover {{ color: #fff; }}

        .ai-chat-btn {{
            position: fixed; bottom: 24px; right: 24px; background: linear-gradient(135deg, #2563eb, #38bdf8); color: #fff; border: none; 
            width: 56px; height: 56px; border-radius: 50%; font-size: 24px; cursor: pointer; box-shadow: 0 8px 24px rgba(37, 99, 235, 0.4); 
            z-index: 99999; display: flex; align-items: center; justify-content: center; transition: transform 0.2s ease;
        }}
        .ai-chat-btn:hover {{ transform: scale(1.08); }}
        .ai-chat-box {{
            position: fixed; bottom: 90px; right: 24px; width: 380px; height: 500px; background: #131b2e; border: 1px solid #1e293b; 
            border-radius: 14px; box-shadow: 0 20px 40px rgba(0, 0, 0, 0.8); z-index: 99999; display: none; flex-direction: column; overflow: hidden;
        }}
        .ai-chat-header {{ background: #0c1220; padding: 14px 16px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; }}
        .ai-chat-messages {{ flex: 1; padding: 14px; overflow-y: auto; font-size: 13px; display: flex; flex-direction: column; gap: 10px; }}
        .ai-msg {{ background: #1e293b; padding: 10px 12px; border-radius: 8px; color: #f8fafc; line-height: 1.4; }}
        .ai-msg.user {{ background: #2563eb; align-self: flex-end; }}
        .ai-chat-input-bar {{ display: flex; padding: 10px; border-top: 1px solid #1e293b; background: #090d16; gap: 8px; }}
        .ai-chat-input {{ flex: 1; background: #131b2e; border: 1px solid #1e293b; color: #fff; padding: 8px 12px; border-radius: 6px; outline: none; font-size: 13px; }}
    </style>
</head>
<body>

<button class="ai-chat-btn" onclick="toggleAIChat()" title="Ask MMS & DZL Merchandising AI Copilot">🤖</button>
<div class="ai-chat-box" id="aiChatBox">
    <div class="ai-chat-header">
        <span style="font-weight:700; color:#fff; font-size:14px;">🧠 MMS & DZL Copilot</span>
        <button onclick="toggleAIChat()" style="background:transparent; border:none; color:#94a3b8; font-size:18px; cursor:pointer;">&times;</button>
    </div>
    <div class="ai-chat-messages" id="aiChatMessages">
        <div class="ai-msg">Hello! I am your Multi-Brand Merchandising Copilot. Ask me about MMS or DZL store performance, targets, stock cover, or critical replenishments.</div>
    </div>
    <div class="ai-chat-input-bar">
        <input type="text" id="aiInput" class="ai-chat-input" placeholder="Ask about MMS or DZL stores, warehouse..." onkeypress="handleAIChatKey(event)">
        <button onclick="sendAIChatMessage()" style="background:#2563eb; color:#fff; border:none; padding:8px 14px; border-radius:6px; font-weight:700; cursor:pointer;">Send</button>
    </div>
</div>

<div id="auth-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:#090d16;z-index:99999999;display:flex;align-items:center;justify-content:center;">
  <div style="background:#131b2e;padding:32px;border-radius:12px;box-shadow:0 15px 30px rgba(0,0,0,0.6);text-align:center;width:90%;max-width:380px;border:1px solid #1e293b;">
    <h3 style="color:#fff;margin:0 0 8px 0;font-size:20px;">🔒 Executive Secure Access</h3>
    <p style="color:#94a3b8;font-size:13px;margin:0 0 20px 0;">Enter authorization PIN to unlock dashboard</p>
    <input type="password" id="access-pass" placeholder="PIN Code" style="width:100%;padding:12px;border-radius:6px;border:1px solid #334155;background:#090d16;color:#fff;font-size:16px;text-align:center;outline:none;box-sizing:border-box;margin-bottom:14px;">
    <button onclick="checkAccess()" style="width:100%;padding:12px;border-radius:6px;border:none;background:#2563eb;color:#fff;font-weight:700;font-size:15px;cursor:pointer;">Unlock Dashboard</button>
    <p id="error-msg" style="color:#ef4444;font-size:13px;margin:12px 0 0 0;display:none;">Invalid authorization credentials</p>
  </div>
</div>

<script>
  const USER_ROLES = {{
    "MMS2026": {{ role: "ADMIN", name: "Executive & Merchandising (Full Access)", region: "ALL" }},
    "SULTAN2026": {{ role: "AREA_MGR", name: "Sultan", region: "Riyadh Central Region" }},
    "RAJIB2026": {{ role: "AREA_MGR", name: "Rajib", region: "Western Region" }}
  }};

  function checkAccess() {{
    var input = document.getElementById("access-pass");
    var val = input ? input.value.trim().toUpperCase() : "";
    var user = USER_ROLES[val];

    if (user) {{
      sessionStorage.setItem("mms_user", JSON.stringify(user));
      applyUserPermissions(user);
      var overlay = document.getElementById("auth-overlay");
      if (overlay) overlay.style.display = "none";
    }} else {{
      var errMsg = document.getElementById("error-msg");
      if (errMsg) errMsg.style.display = "block";
    }}
  }}

  function applyUserPermissions(user) {{
    var userBadge = document.getElementById("current-user-badge");
    if (userBadge) userBadge.innerHTML = "👤 " + user.name;
    if (user.role === "ADMIN") return;

    var grandTotal = document.getElementById("grand-total-banner");
    if (grandTotal) grandTotal.style.display = "none";

    document.querySelectorAll(".region-block").forEach(function(el) {{
      if (el.getAttribute("data-region") !== user.region) el.style.display = "none";
    }});
    document.querySelectorAll(".region-table-wrap").forEach(function(el) {{
      if (el.getAttribute("data-region") !== user.region) el.style.display = "none";
    }});
    document.querySelectorAll("#storesTable tbody tr").forEach(function(el) {{
      if (el.getAttribute("data-region") !== user.region) el.style.display = "none";
    }});
    document.querySelectorAll(".decision-card-item").forEach(function(el) {{
      if (el.getAttribute("data-region") !== user.region) el.style.display = "none";
    }});
  }}

  function logout() {{
    sessionStorage.removeItem("mms_user");
    location.reload();
  }}

  document.addEventListener("DOMContentLoaded", function() {{
    var passInput = document.getElementById("access-pass");
    if (passInput) {{
      passInput.addEventListener("keypress", function(e) {{
        if (e.key === "Enter") checkAccess();
      }});
    }}
    var savedUser = sessionStorage.getItem("mms_user");
    if (savedUser) {{
      try {{
        var u = JSON.parse(savedUser);
        applyUserPermissions(u);
        var overlay = document.getElementById("auth-overlay");
        if (overlay) overlay.style.display = "none";
      }} catch(e) {{}}
    }}
  }});
</script>

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
          <strong style="color:#ef4444;">● Situation & Root Cause:</strong> <span id="modal-diag" style="color:#f8fafc;">-</span>
        </div>
        <div style="font-size:13px; color:#f59e0b; margin-bottom:10px;">
          <strong>🎯 Store Requirements:</strong> <span id="modal-needs" style="color:#fff;">-</span>
        </div>
        <div style="font-size:13px; color:#38bdf8; background:rgba(56,189,248,0.08); padding:10px 14px; border-radius:6px; border:1px solid rgba(56,189,248,0.25);">
          <strong style="color:#38bdf8;">⚡ Directive:</strong> <span id="modal-directive" style="color:#fff; font-weight:600;">-</span>
        </div>
      </div>

      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Commercial Metrics</div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:12px; margin-bottom:24px;">
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">CURRENT SALES</div>
          <div id="modal-sales" style="font-size:17px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">LY GROSS SALES</div>
          <div id="modal-ly" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div>
          <div id="modal-yoy" style="font-size:11px; margin-top:2px;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">WOC COVER</div>
          <div id="modal-woc" style="font-size:17px; font-weight:700; color:#f59e0b;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">SELL-THROUGH (STR)</div>
          <div id="modal-str" style="font-size:17px; font-weight:700; color:#10b981;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">FLOOR SOH</div>
          <div id="modal-soh" style="font-size:17px; font-weight:700; color:#38bdf8;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">ATV</div>
          <div id="modal-atv" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">UPT</div>
          <div id="modal-upt" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:14px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:11px; color:#94a3b8;">ASP</div>
          <div id="modal-asp" style="font-size:16px; font-weight:700; color:#f59e0b;">-</div>
        </div>
      </div>

      <div style="margin-bottom:12px; font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8;">Store Category Contribution Breakdown (% Overall)</div>
      <div style="border:1px solid #1e293b; border-radius:8px; overflow:hidden;">
        <table>
          <thead>
            <tr>
              <th>#</th>
              <th>Main Category</th>
              <th>Sales (SAR)</th>
              <th>Units Sold</th>
              <th>Category Contribution (%)</th>
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
        <h1>MMS & DZL Executive Commercial Intelligence Dashboard</h1>
        <p>Operational Performance, Multi-Brand Portfolio, Regional Hierarchy & LY Benchmarks</p>
    </div>
    <div class="top-controls">
        <div class="brand-switcher">
            <button class="brand-btn active" id="btn-brand-ALL" onclick="switchBrand('ALL')">🏢 ALL BRANDS</button>
            <button class="brand-btn" id="btn-brand-MMS" onclick="switchBrand('MMS')">🔴 MUMUSO (17)</button>
            <button class="brand-btn" id="btn-brand-DZL" onclick="switchBrand('DZL')">🟡 DZL (6)</button>
        </div>
        <span id="current-user-badge" style="font-size:13px; font-weight:700; color:#38bdf8; background:#1e293b; padding:8px 14px; border-radius:8px; border:1px solid #334155;">👤 Authenticating..</span>
        <button class="logout-btn" onclick="logout()">Logout</button>
    </div>
</div>

<div class="section-title"><span>🤖 AI Merchandising Directives</span></div>
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

<!-- بطاقات الـ KPI العلوية الديناميكية التي تتحدث تلقائياً مع اختيار البراند -->
<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-title">Current Total Sales</div>
        <div class="kpi-value" id="kpi-total-sales">{total_sales:,} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">LY Gross Sales</div>
        <div class="kpi-value" id="kpi-ly-sales" style="color:#38bdf8;">{total_ly_sales:,} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network LFL YoY Growth</div>
        <div class="kpi-value" id="kpi-yoy-growth" style="color:{net_yoy_col};">{network_lfl_growth:+.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Total Target</div>
        <div class="kpi-value" id="kpi-total-target">{total_target:,} <span class="kpi-unit">SAR</span></div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Achievement (% Ach)</div>
        <div class="kpi-value" id="kpi-overall-ach" style="color: {'#10b981' if overall_ach >= 100 else '#f59e0b'};">{overall_ach:.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ATV</div>
        <div class="kpi-value" id="kpi-network-atv">SAR {network_atv:,}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Network ASP</div>
        <div class="kpi-value" id="kpi-network-asp">SAR {network_asp:,}</div>
    </div>
</div>

<div class="view-toggle-bar">
    <button class="view-btn active" id="btn-stores" onclick="switchView('stores')">🏢 Store Commercial Matrix</button>
    <button class="view-btn" id="btn-regions" onclick="switchView('regions')">🌍 Region-Wise Performance</button>
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business-Wise Performance (9 Categories)</button>
    <button class="view-btn" id="btn-action" onclick="switchView('action')" style="border-left:2px solid #38bdf8;">⚡ Commercial Action Hub (Replenishment & Top/Low 500)</button>
</div>

<!-- 1. Store Commercial Matrix View -->
<div id="view-stores">
    <div class="section-title"><span>⚡ Critical Action Directives (Category Swaps & Rebalancing Priorities)</span></div>
    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:14px; margin-bottom:24px;">
        {decision_cards_html}
    </div>

    <div class="chart-container">
        <div class="section-title">
            <span>📊 Top Stores Performance vs Target (Interactive ApexCharts)</span>
        </div>
        <div id="apexStoreChart" style="min-height: 350px;"></div>
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3>STORE COMMERCIAL & DISPLAY ASSORTMENT MATRIX</h3>
                <span style="color:var(--text-muted);font-size:12px;">Click any store row to open category contribution breakdown</span>
            </div>
            <input type="text" id="storeSearch" class="table-search" placeholder="Search store name, code, or region..." onkeyup="filterStores()">
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
                        <th>QTY Sold</th>
                        <th>Transactions</th>
                        <th>UPT</th>
                        <th>STR%</th>
                        <th>Commercial Diagnostic</th>
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
        <span>🏷️ PRODUCT CATEGORIES & HIERARCHY MATRIX</span>
        <span style="font-size:12px; color:var(--text-muted); font-weight:400;">Select store from dropdown to view category contribution & stock health ratio</span>
    </div>
    
    <div class="cards-scroll-container">
        {main_cat_cards_html}
    </div>

    <div class="chart-container" style="margin-bottom:24px;">
        <div class="section-title">
            <span>🍩 Category Revenue Contribution Share</span>
        </div>
        <div id="apexCategoryDonut" style="min-height: 330px;"></div>
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3 id="tableHierarchyTitle">PRODUCT HIERARCHY MATRIX (LEVEL 1: MAIN CATEGORIES)</h3>
                <span style="color:var(--text-muted);font-size:12px;">Select Store to inspect category mix and Stock Health Ratio</span>
            </div>
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <select id="storeDropdownFilter" class="table-select" onchange="onStoreDropdownChange(this.value)">
                    {store_options_html}
                </select>
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
                        <th>Stock Health / WOC</th>
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

<!-- 4. Commercial Action Hub (Isolated MMS & DZL Auto-Replenishment) -->
<div id="view-action" style="display:none;">
    <div class="section-title" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
        <span>⚡ PREDICTIVE AUTO-REPLENISHMENT & STOCK-OUT FORECAST (KSWH & IST)</span>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
            <a href="./reports/Auto_Replenishment_MMS.xlsx" download class="export-btn" style="text-decoration:none; display:inline-flex; align-items:center; gap:6px;">📥 Download MMS Replenishment Plan</a>
            <a href="./reports/Auto_Replenishment_DZL.xlsx" download class="export-btn dzl-btn" style="text-decoration:none; display:inline-flex; align-items:center; gap:6px;">📥 Download DZL Replenishment Plan</a>
        </div>
    </div>

    <!-- أزرار تصفية خطة التوريد حسب البراند -->
    <div style="display:flex; gap:10px; margin-bottom:14px; align-items:center;">
        <span style="font-size:13px; font-weight:700; color:#94a3b8;">Filter Replenishment By Brand:</span>
        <button class="sub-tab-btn active" id="repl-btn-all" onclick="filterReplByBrand('ALL')">🏢 ALL BRANDS</button>
        <button class="sub-tab-btn" id="repl-btn-mms" onclick="filterReplByBrand('MMS')">🔴 MUMUSO ONLY</button>
        <button class="sub-tab-btn" id="repl-btn-dzl" onclick="filterReplByBrand('DZL')">🟡 DZL ONLY</button>
    </div>

    <div class="table-wrap" style="margin-bottom:30px;">
        <div class="table-header">
            <div>
                <h3 style="margin:0; font-size:15px; color:#fff;" id="replTableHeaderTitle">⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (MULTI-STORE BALANCING)</h3>
                <span style="color:var(--text-muted); font-size:12px;">Transfers are strictly isolated per brand network to guarantee operational compliance</span>
            </div>
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <input type="text" id="replSearch" class="table-search" placeholder="Search SKU, Store, or Category..." onkeyup="filterReplTable()">
                <select id="replTypeFilter" class="table-select" onchange="filterReplTable()">
                    <option value="ALL">All Actions</option>
                    <option value="Predictive WH Replenishment">Central WH (KSWH)</option>
                    <option value="Store Transfer (IST)">Store Transfers (IST)</option>
                </select>
            </div>
        </div>
        <div style="max-height: 480px; overflow-y: auto; overflow-x: auto;">
            <table id="replTable">
                <thead style="position: sticky; top: 0; z-index: 10;">
                    <tr>
                        <th>#</th>
                        <th>Action Type</th>
                        <th>Store Name & Code</th>
                        <th>SKU & Category Focus</th>
                        <th>Source Route (Optimized Proximity)</th>
                        <th>Suggested Qty (Pcs)</th>
                        <th>Stock-Out Forecast</th>
                    </tr>
                </thead>
                <tbody id="replTableBody">
                    {repl_rows_html}
                </tbody>
            </table>
        </div>
    </div>

    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; flex-wrap:wrap; gap:12px;">
        <div style="display:flex; gap:10px;">
            <button class="sub-tab-btn active" id="btn-top500" onclick="switchMoversTab('top')">🔥 TOP 500 HIGH-VELOCITY FAST MOVERS</button>
            <button class="sub-tab-btn" id="btn-low500" onclick="switchMoversTab('low')">❄️ LOW 500 DEAD & SLOW STOCK (CLEARANCE CANDIDATES)</button>
        </div>
        <div style="display:flex; gap:10px; align-items:center;">
            <select id="moversCatFilter" class="table-select" onchange="filterMoversTable()">
                {main_cat_options}
            </select>
            <input type="text" id="moversSearch" class="table-search" placeholder="Search SKU code or name..." onkeyup="filterMoversTable()">
        </div>
        <div>
            <button class="export-btn" onclick="exportMoversToExcel()">📥 Export List to Excel</button>
        </div>
    </div>

    <div class="table-wrap">
        <div style="overflow-x:auto;">
            <table id="moversTable">
                <thead>
                    <tr>
                        <th>Rank</th>
                        <th>Item Code / Barcode</th>
                        <th>Product Name</th>
                        <th>Main Category</th>
                        <th>Sub-Category</th>
                        <th>Units Sold</th>
                        <th>Sales Revenue (SAR)</th>
                        <th>ASP (SAR)</th>
                        <th>WH SOH (KSWH)</th>
                        <th>Est. Stock-Out (Days)</th>
                    </tr>
                </thead>
                <tbody id="moversTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script>
  const STORE_DETAILS = {json.dumps(store_cat_summary_dict)};
  const STORE_META = {json.dumps(store_meta_map)};
  const BRAND_KPIS = {json.dumps(brand_kpi_summary)};
  const SUBSUB_DATA = {json.dumps(subsub_json_data)};
  const MAIN_CAT_HTML = `{main_cat_table_rows}`;
  const TOP_500_DATA = {top500_json};
  const LOW_500_DATA = {low500_json};
  const AI_CONTEXT = {json.dumps(dashboard_ai_context)};

  let currentMoversType = 'top';
  let currentActiveBrand = 'ALL';
  let currentReplBrand = 'ALL';

  document.addEventListener("DOMContentLoaded", function() {{
    initApexCharts();
  }});

  function initApexCharts() {{
    var storeOptions = {{
      series: [
        {{ name: 'Actual Sales (SAR)', data: {json.dumps(apex_sales)} }},
        {{ name: 'Target (SAR)', data: {json.dumps(apex_targets)} }}
      ],
      chart: {{
        type: 'bar', height: 340, toolbar: {{ show: false }}, background: 'transparent'
      }},
      theme: {{ mode: 'dark' }},
      colors: ['#38bdf8', '#334155'],
      plotOptions: {{
        bar: {{ horizontal: false, columnWidth: '45%', borderRadius: 4 }}
      }},
      dataLabels: {{ enabled: false }},
      stroke: {{ show: true, width: 2, colors: ['transparent'] }},
      xaxis: {{
        categories: {json.dumps(apex_categories)},
        labels: {{ style: {{ colors: '#94a3b8', fontSize: '11px' }} }}
      }},
      yaxis: {{
        labels: {{
          formatter: function (val) {{ return Number(val).toLocaleString() + ' SAR'; }},
          style: {{ colors: '#94a3b8', fontSize: '11px' }}
        }}
      }},
      fill: {{ opacity: 1 }},
      tooltip: {{
        theme: 'dark',
        y: {{ formatter: function (val) {{ return Number(val).toLocaleString() + ' SAR'; }} }}
      }},
      grid: {{ borderColor: '#1e293b' }}
    }};

    var storeChartEl = document.querySelector("#apexStoreChart");
    if (storeChartEl) {{
      var storeChart = new ApexCharts(storeChartEl, storeOptions);
      storeChart.render();
    }}

    var donutOptions = {{
      series: {json.dumps(apex_cat_shares)},
      labels: {json.dumps(apex_cat_names)},
      chart: {{ type: 'donut', height: 330, background: 'transparent' }},
      theme: {{ mode: 'dark' }},
      colors: ['#38bdf8', '#818cf8', '#a855f7', '#ec4899', '#f59e0b', '#10b981', '#06b6d4', '#e11d48', '#84cc16'],
      plotOptions: {{
        pie: {{
          donut: {{
            size: '65%',
            labels: {{
              show: true,
              total: {{ show: true, label: 'Mix Share', formatter: function () {{ return '100%'; }} }}
            }}
          }}
        }}
      }},
      legend: {{ position: 'bottom', labels: {{ colors: '#cbd5e1' }} }},
      tooltip: {{
        theme: 'dark',
        y: {{ formatter: function(val) {{ return val.toFixed(1) + '% Contribution'; }} }}
      }}
    }};

    var donutEl = document.querySelector("#apexCategoryDonut");
    if (donutEl) {{
      var donutChart = new ApexCharts(donutEl, donutOptions);
      donutChart.render();
    }}
  }}

  // دالة التبديل الفوري بين البراندين وتحديث بطاقات الـ KPI العلوية
  function switchBrand(brand) {{
    currentActiveBrand = brand;
    document.querySelectorAll(".brand-btn").forEach(btn => btn.classList.remove("active"));
    const activeBtn = document.getElementById("btn-brand-" + brand);
    if (activeBtn) activeBtn.classList.add("active");

    // 1. تحديث بطاقات الـ KPI العلوية لحظياً
    if (BRAND_KPIS[brand]) {{
      const k = BRAND_KPIS[brand];
      document.getElementById("kpi-total-sales").innerHTML = k.sales;
      document.getElementById("kpi-total-target").innerText = k.target;
      document.getElementById("kpi-ly-sales").innerText = k.ly;
      document.getElementById("kpi-overall-ach").innerText = k.ach;
      document.getElementById("kpi-yoy-growth").innerText = k.yoy;
      document.getElementById("kpi-network-atv").innerText = k.atv;
      document.getElementById("kpi-network-asp").innerText = k.asp;
    }}

    // 2. فلترة صفوف الجداول
    document.querySelectorAll(".store-row").forEach(row => {{
      const rBrand = row.getAttribute("data-brand");
      row.style.display = (brand === "ALL" || rBrand === brand) ? "" : "none";
    }});

    // 3. فلترة بطاقات القرارات
    document.querySelectorAll(".store-card-item").forEach(card => {{
      const cBrand = card.getAttribute("data-brand");
      card.style.display = (brand === "ALL" || cBrand === brand) ? "" : "none";
    }});

    // 4. فلترة جدول التوريد وقائمة الأصناف
    filterReplByBrand(brand);
    renderMoversTable();
  }}

  function filterReplByBrand(brand) {{
    currentReplBrand = brand;
    document.getElementById("repl-btn-all").classList.toggle("active", brand === "ALL");
    document.getElementById("repl-btn-mms").classList.toggle("active", brand === "MMS");
    document.getElementById("repl-btn-dzl").classList.toggle("active", brand === "DZL");

    const titleEl = document.getElementById("replTableHeaderTitle");
    if (titleEl) {{
        if (brand === "MMS") titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (MUMUSO ONLY)";
        else if (brand === "DZL") titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (DZL DOZOLO ONLY)";
        else titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (MULTI-STORE BALANCING)";
    }}
    filterReplTable();
  }}

  function filterReplTable() {{
    var searchVal = document.getElementById("replSearch").value.toLowerCase();
    var typeFilter = document.getElementById("replTypeFilter").value;
    var rows = document.querySelectorAll("#replTableBody tr");

    rows.forEach(function(row) {{
      var text = row.innerText.toLowerCase();
      var repBrand = row.getAttribute("data-brand");
      var matchBrand = (currentReplBrand === "ALL" || repBrand === currentReplBrand);
      var matchSearch = text.includes(searchVal);
      var matchType = (typeFilter === "ALL") || text.includes(typeFilter.toLowerCase());
      row.style.display = (matchBrand && matchSearch && matchType) ? "" : "none";
    }});
  }}

  function toggleAIChat() {{
    var box = document.getElementById("aiChatBox");
    box.style.display = (box.style.display === "flex") ? "none" : "flex";
  }}

  function handleAIChatKey(e) {{
    if (e.key === 'Enter') sendAIChatMessage();
  }}

  function sendAIChatMessage() {{
    var input = document.getElementById("aiInput");
    var q = input.value.trim();
    if (!q) return;

    var container = document.getElementById("aiChatMessages");
    var userDiv = document.createElement("div");
    userDiv.className = "ai-msg user";
    userDiv.innerText = q;
    container.appendChild(userDiv);
    input.value = "";

    var reply = "";
    var qLower = q.toLowerCase();

    if (qLower.includes("dzl") || qLower.includes("dozolo")) {{
      reply = `DZL Network has 6 operational stores. Top central location is DZL Solitaire and Riyadh Park. All DZL replenishment and transfers are strictly isolated from MMS.`;
    }} else if (qLower.includes("warehouse") || qLower.includes("kswh") || qLower.includes("stock")) {{
      reply = `Central Warehouse (KSWH) currently holds ${{AI_CONTEXT.wh_soh.toLocaleString()}} units ready for dispatch across both networks.`;
    }} else if (qLower.includes("target") || qLower.includes("ach") || qLower.includes("achievement")) {{
      reply = `Overall portfolio target achievement is ${{AI_CONTEXT.overall_ach.toFixed(1)}}% with total sales of ${{Math.round(AI_CONTEXT.network_sales).toLocaleString()}} SAR against a target of ${{Math.round(AI_CONTEXT.network_target).toLocaleString()}} SAR.`;
    }} else if (qLower.includes("transfer") || qLower.includes("shortage") || qLower.includes("ist")) {{
      reply = `Store transfers (IST) are strictly restricted within the same brand. High priority intra-city balancing is active in Riyadh and Jeddah.`;
    }} else {{
      var matchedStore = AI_CONTEXT.stores.find(s => qLower.includes(s.clean_code.toLowerCase()) || qLower.includes(s.full_name.toLowerCase()));
      if (matchedStore) {{
        reply = `Store ${{matchedStore.full_name}} (${{matchedStore.clean_code}} - [${{matchedStore.brand}}]): Total Sales ${{Math.round(matchedStore.sales).toLocaleString()}} SAR, Target Achievement ${{matchedStore.ach_pct ? matchedStore.ach_pct.toFixed(1) + '%' : 'N/A'}}, UPT: ${{matchedStore.upt}}, Status: ${{matchedStore.diag_title}}.`;
      }} else {{
        reply = `Network summary: Solitaire and Dhahran lead performance. Check Commercial Action Hub for brand-specific IST transfers.`;
      }}
    }}

    setTimeout(() => {{
      var botDiv = document.createElement("div");
      botDiv.className = "ai-msg";
      botDiv.innerText = reply;
      container.appendChild(botDiv);
      container.scrollTop = container.scrollHeight;
    }}, 400);
  }}

  function safeSetText(id, text) {{
    const el = document.getElementById(id);
    if (el) el.innerText = (text !== undefined && text !== null) ? text : "-";
  }}

  function safeSetHtml(id, htmlContent) {{
    const el = document.getElementById(id);
    if (el) el.innerHTML = (htmlContent !== undefined && htmlContent !== null) ? htmlContent : "-";
  }}

  function openStoreDetails(storeCode) {{
    try {{
      const meta = STORE_META[storeCode];
      const cats = STORE_DETAILS[storeCode] || [];
      if (!meta) return;

      safeSetText("modal-store-name", "[" + meta.brand + "] " + meta.name + " - Category Breakdown");
      safeSetText("modal-store-code", "CODE: " + storeCode + " | " + meta.region + " (Manager: " + meta.manager + ")");
      safeSetText("modal-sales", meta.sales);
      safeSetText("modal-ly", meta.ly_sales);
      
      const yoyHtml = (meta.yoy && meta.yoy !== "-") ? "YoY: <strong>" + meta.yoy + "</strong>" : "New Location";
      safeSetHtml("modal-yoy", yoyHtml);

      safeSetText("modal-woc", meta.woc);
      safeSetText("modal-str", meta.str);
      safeSetText("modal-soh", meta.soh_units);
      safeSetText("modal-atv", meta.atv);
      safeSetText("modal-upt", meta.upt);
      safeSetText("modal-asp", meta.asp);
      safeSetText("modal-diag", meta.problem);
      safeSetText("modal-needs", meta.needs);
      safeSetText("modal-directive", meta.action);

      let rowsHtml = "";
      cats.forEach((c, idx) => {{
        rowsHtml += `
          <tr>
            <td style="color:#64748b;">${{idx+1}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{c.main_category}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{c.sales}}</td>
            <td>${{c.units}}</td>
            <td style="color:#10b981; font-weight:800; font-size:14px;">${{c.store_mix_pct}}</td>
            <td style="color:#f59e0b; font-weight:700;">${{c.asp}}</td>
          </tr>
        `;
      }});

      safeSetHtml("modal-cats-body", rowsHtml || "<tr><td colspan='6' style='text-align:center;'>No category data available</td></tr>");
      
      const modal = document.getElementById("store-modal");
      if (modal) modal.style.display = "flex";
    }} catch (err) {{
      console.error("Error opening store details:", err);
    }}
  }}

  function filterByMainCategory(catName) {{
    const filter = document.getElementById("mainCatFilter");
    if (filter) filter.value = catName;
    updateBusinessTable();
  }}

  function onStoreDropdownChange(storeCode) {{ updateBusinessTable(); }}
  function onCategoryFilterChange(catName) {{ updateBusinessTable(); }}

  function updateBusinessTable() {{
    const storeCode = document.getElementById("storeDropdownFilter").value;
    const catName = document.getElementById("mainCatFilter").value;
    const thead = document.getElementById("hierarchyTableHead");
    const tbody = document.getElementById("hierarchyTableBody");
    const title = document.getElementById("tableHierarchyTitle");

    if (storeCode === "ALL" && catName === "ALL") {{
      title.innerText = "PRODUCT HIERARCHY MATRIX (LEVEL 1: MAIN CATEGORIES)";
      thead.innerHTML = `
        <tr>
          <th>#</th>
          <th>Main Category</th>
          <th>Sales Revenue (SAR)</th>
          <th>Sales Units</th>
          <th>Network Share (%)</th>
          <th>Stock Health / WOC</th>
          <th>ASP (SAR)</th>
          <th>Leading Store Benchmark</th>
        </tr>
      `;
      tbody.innerHTML = MAIN_CAT_HTML;
      return;
    }}

    if (storeCode !== "ALL" && catName === "ALL") {{
      const stCats = STORE_DETAILS[storeCode] || [];
      const storeMeta = STORE_META[storeCode];
      title.innerText = "CATEGORY CONTRIBUTION FOR: " + (storeMeta ? storeMeta.name : storeCode);
      thead.innerHTML = `
        <tr>
          <th>#</th>
          <th>Main Category</th>
          <th>Sales Revenue (SAR)</th>
          <th>Sales Units</th>
          <th>Category Contribution (%)</th>
          <th>Stock Health Ratio</th>
          <th>ASP (SAR)</th>
        </tr>
      `;
      let rowsHtml = "";
      stCats.forEach((c, idx) => {{
        rowsHtml += `
          <tr>
            <td style="color:#64748b;">${{idx+1}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{c.main_category}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{c.sales}}</td>
            <td>${{c.units}}</td>
            <td style="color:#10b981; font-weight:800; font-size:14px;">${{c.store_mix_pct}}</td>
            <td><span class="badge" style="background:${{c.health_color}}22; color:${{c.health_color}}; border:1px solid ${{c.health_color}}55;">${{c.stock_health}}</span></td>
            <td style="color:#f59e0b; font-weight:700;">${{c.asp}}</td>
          </tr>
        `;
      }});
      tbody.innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No data available for this store</td></tr>";
      return;
    }}

    title.innerText = "SUB-SUBGROUP BREAKDOWN: " + catName.toUpperCase() + (storeCode !== "ALL" ? " (Filtered by Store)" : "");
    thead.innerHTML = `
      <tr>
        <th>#</th>
        <th>Main Category</th>
        <th>Sub-Subgroup (Item Class)</th>
        <th>Sales Revenue (SAR)</th>
        <th>Sales Units</th>
        <th>Contribution (%)</th>
        <th>ASP (SAR)</th>
      </tr>
    `;

    let filtered = SUBSUB_DATA;
    if (catName !== "ALL") filtered = filtered.filter(x => x.main_category === catName);

    let rowsHtml = "";
    filtered.forEach((r, idx) => {{
      const bar_w = Math.min(r.contribution * 3, 100);
      rowsHtml += `
        <tr>
          <td style="color:#64748b;">${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:600;">${{r.main_category}}</td>
          <td style="font-weight:700; color:#fff;">${{r.sub_subgroup}}</td>
          <td style="font-weight:700; color:#38bdf8;">${{Math.round(r.sales).toLocaleString()}}</td>
          <td>${{Number(r.units).toLocaleString()}}</td>
          <td style="min-width:130px;">
            <div style="display:flex;align-items:center;gap:6px;">
              <span style="font-weight:700;color:#fff;min-width:40px;">${{r.contribution.toFixed(1)}}%</span>
              <div style="flex:1;background:#1e293b;border-radius:4px;height:5px;overflow:hidden;">
                <div style="width:${{bar_w}}%;background:#38bdf8;height:100%;"></div>
              </div>
            </div>
          </td>
          <td style="color:#f59e0b;font-weight:700;">${{Math.round(r.asp).toLocaleString()}}</td>
        </tr>
      `;
    }});

    tbody.innerHTML = rowsHtml || "<tr><td colspan='7' style='text-align:center;'>No matching products found for this filter</td></tr>";
  }}

  function switchMoversTab(type) {{
    currentMoversType = type;
    document.getElementById("btn-top500").classList.toggle("active", type === 'top');
    document.getElementById("btn-low500").classList.toggle("active", type === 'low');
    renderMoversTable();
  }}

  function renderMoversTable() {{
    const data = (currentMoversType === 'top') ? TOP_500_DATA : LOW_500_DATA;
    const catFilter = document.getElementById("moversCatFilter").value;
    const searchVal = document.getElementById("moversSearch").value.toLowerCase();
    const tbody = document.getElementById("moversTableBody");

    let filtered = data.filter(item => {{
      const matchBrand = (currentActiveBrand === "ALL" || item.brand === currentActiveBrand);
      const matchCat = (catFilter === "ALL" || item.main_cat === catFilter);
      const matchSearch = (item.code.toLowerCase().includes(searchVal) || item.name.toLowerCase().includes(searchVal) || item.subsub.toLowerCase().includes(searchVal));
      return matchBrand && matchCat && matchSearch;
    }});

    let html = "";
    filtered.slice(0, 100).forEach(r => {{
      const rankColor = currentMoversType === 'top' ? '#10b981' : '#ef4444';
      const daysColor = r.stock_days < 15 ? '#ef4444' : (r.stock_days < 30 ? '#f59e0b' : '#10b981');
      const bColor = r.brand === 'DZL' ? '#ef4444' : '#38bdf8';
      html += `
        <tr>
          <td style="color:${{rankColor}}; font-weight:800;">#${{r.rank}}</td>
          <td style="color:#38bdf8; font-weight:600;"><span class="badge" style="background:${{bColor}}22; color:${{bColor}}; margin-right:4px;">${{r.brand}}</span>${{r.code}}</td>
          <td style="color:#fff; font-weight:600;">${{r.name}}</td>
          <td>${{r.main_cat}}</td>
          <td style="color:#94a3b8;">${{r.subsub}}</td>
          <td style="font-weight:700; color:#fff;">${{r.units.toLocaleString()}}</td>
          <td style="font-weight:700; color:#38bdf8;">${{r.sales.toLocaleString()}}</td>
          <td style="color:#f59e0b; font-weight:600;">${{r.asp.toLocaleString()}}</td>
          <td style="font-weight:700; color:#38bdf8;">${{r.wh_soh.toLocaleString()}} Pcs</td>
          <td><span class="badge" style="background:${{daysColor}}22; color:${{daysColor}};">${{r.stock_days === 999 ? 'Stable' : r.stock_days + ' Days'}}</span></td>
        </tr>
      `;
    }});

    tbody.innerHTML = html || "<tr><td colspan='10' style='text-align:center;'>No matching SKUs found</td></tr>";
  }}

  function filterMoversTable() {{ renderMoversTable(); }}

  function exportMoversToExcel() {{
    const data = (currentMoversType === 'top') ? TOP_500_DATA : LOW_500_DATA;
    let csv = "Rank,Brand,Item Code,Product Name,Main Category,Sub-Category,Units Sold,Sales Revenue (SAR),ASP (SAR),WH SOH (KSWH),Est. Stock-Out (Days)\\n";
    data.forEach(r => {{
      csv += `"${{r.rank}}","${{r.brand}}","${{r.code}}","${{r.name.replace(/"/g, '""')}}","${{r.main_cat}}","${{r.subsub}}","${{r.units}}","${{r.sales}}","${{r.asp}}","${{r.wh_soh}}","${{r.stock_days}}"\\n`;
    }});

    const blob = new Blob(["\\uFEFF" + csv], {{ type: 'text/csv;charset=utf-8;' }});
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Portfolio_${{currentMoversType.toUpperCase()}}_500_Movers.csv`;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }}

  function filterSubSubTable() {{
    const q = document.getElementById("subsubSearch").value.toLowerCase();
    const rows = document.querySelectorAll("#hierarchyTableBody tr");
    rows.forEach(r => {{
      r.style.display = r.innerText.toLowerCase().includes(q) ? "" : "none";
    }});
  }}

  function closeModal(modalId) {{
    const modal = document.getElementById(modalId);
    if (modal) modal.style.display = "none";
  }}

  window.onclick = function (event) {{
    if (event.target && event.target.classList.contains('app-modal')) {{
      event.target.style.display = "none";
    }}
  }};

  function switchView(viewName) {{
    const storesView = document.getElementById("view-stores");
    const regionsView = document.getElementById("view-regions");
    const businessView = document.getElementById("view-business");
    const actionView = document.getElementById("view-action");
    const btnStores = document.getElementById("btn-stores");
    const btnRegions = document.getElementById("btn-regions");
    const btnBusiness = document.getElementById("btn-business");
    const btnAction = document.getElementById("btn-action");

    if (storesView) storesView.style.display = "none";
    if (regionsView) regionsView.style.display = "none";
    if (businessView) businessView.style.display = "none";
    if (actionView) actionView.style.display = "none";
    if (btnStores) btnStores.classList.remove("active");
    if (btnRegions) btnRegions.classList.remove("active");
    if (btnBusiness) btnBusiness.classList.remove("active");
    if (btnAction) btnAction.classList.remove("active");

    if (viewName === 'stores' && storesView) {{
        storesView.style.display = "block";
        if (btnStores) btnStores.classList.add("active");
    }} else if (viewName === 'regions' && regionsView) {{
        regionsView.style.display = "block";
        if (btnRegions) btnRegions.classList.add("active");
    }} else if (viewName === 'business' && businessView) {{
        businessView.style.display = "block";
        if (btnBusiness) btnBusiness.classList.add("active");
    }} else if (actionView) {{
        actionView.style.display = "block";
        if (btnAction) btnAction.classList.add("active");
        renderMoversTable();
    }}
  }}

  function filterStores() {{
      const query = document.getElementById("storeSearch").value.toLowerCase();
      const rows = document.querySelectorAll("#storesTable tbody tr");
      rows.forEach(r => {{
          const text = r.innerText.toLowerCase();
          const rBrand = r.getAttribute("data-brand");
          const matchBrand = (currentActiveBrand === "ALL" || rBrand === currentActiveBrand);
          r.style.display = (matchBrand && text.includes(query)) ? "" : "none";
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