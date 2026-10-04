import os
import glob
import re
import json
import html
import pandas as pd
import numpy as np

REPORTS_DIR = "./reports"

STORE_MAPPING = {
    # Central & Eastern Region (Sultan - 11 Doors)
    "K108": {"full_name": "MMS Riyadh Solitaire", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K301": {"full_name": "MMS Mall of Dhahran", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Dhahran", "brand": "MMS"},
    "K101": {"full_name": "MMS Riyadh The View Mall", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K110": {"full_name": "MMS Riyadh U-Walk", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K102": {"full_name": "MMS Riyadh Tala Mall", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K109": {"full_name": "MMS Riyadh Lastrada", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K130": {"full_name": "MMS Riyadh Al-Rabwa", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K111": {"full_name": "MMS Riyadh Al-Rabwa (K111)", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K107": {"full_name": "DZL Riyadh Park", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "DZL104": {"full_name": "DZL Riyadh U-Walk", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "DZL112": {"full_name": "DZL Solitaire", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},

    # Western, Southern & Northern Region (Rajib - 12 Doors)
    "K205": {"full_name": "MMS Jeddah U-Walk", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K211": {"full_name": "MMS Madinah", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Madinah", "brand": "MMS"},
    "K201": {"full_name": "MMS Jeddah Park", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K202": {"full_name": "MMS Jeddah Yasmin Mall", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K208": {"full_name": "MMS Juri Mall", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Taif", "brand": "MMS"},
    "K401": {"full_name": "MMS Najran Park", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Najran", "brand": "MMS"},
    "K404": {"full_name": "MMS Abha", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Abha", "brand": "MMS"},
    "K501": {"full_name": "MMS Tabuk Park", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Tabuk", "brand": "MMS"},
    "K210": {"full_name": "MMS Makkah Salam Mall", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Makkah", "brand": "MMS"},
    "K403": {"full_name": "MMS RMJ", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "MMS"},
    "K204": {"full_name": "DZL Redsea", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"},
    "DZL204": {"full_name": "DZL Redsea (DZL204)", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"}
}

DZL_VALID_CODES = {"K107", "DZL104", "DZL112", "K204", "DZL204"}
MMS_VALID_CODES = {k for k, v in STORE_MAPPING.items() if v["brand"] == "MMS"}
ALL_VALID_CODES = set(STORE_MAPPING.keys())

EXCLUDED_KEYWORDS = [
    'gwp', 'shopping bag', 'carrier bag', 'plastic bag', 'paper bag',
    'gift with purchase', 'non-sale', 'packaging', 'free gift', 'material',
    'stationery', 'basketball', 'football', 'm&g'
]

def clean_store_code_str(val):
    s = str(val).strip().upper()
    if s.endswith('.0'): s = s[:-2]
    if "DZL107" in s or ("107" in s and "DZL" in s) or "RIYADH PARK(DZL)" in s: return "K107"
    if "DZL104" in s or ("104" in s and "DZL" in s) or "UWALK(DZL)" in s: return "DZL104"
    if "DZL112" in s or ("112" in s and "DZL" in s) or "SOLITAIRE(DZL)" in s: return "DZL112"
    if "DZL204" in s or ("204" in s and "DZL" in s) or "RED SEA(DZL)" in s: return "K204"
    if "RABWA" in s: return "K130"
    m = re.search(r'\b[A-Z]?(\d{3,4})\b', s)
    if m: return f"K{m.group(1)}"
    return s

def clean_sku_code(val):
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def derive_style_group_from_name(name):
    s = str(name).strip()
    s = re.sub(r'[\s\-]+(2[0-9]|3[0-9]|4[0-8])\b.*$', '', s)
    s = re.sub(r'[\s\-]+(BLACK|WHITE|BEIGE|GREY|GRAY|BLUE|PINK|GREEN|BROWN|RED|YELLOW|KHAKI|OFF-WHITE|SILVER|GOLD)\b.*$', '', s, flags=re.IGNORECASE)
    return s.strip(' -_') if s.strip(' -_') else str(name)[:28]

def classify_shoe_gender_by_size(name, spec=""):
    text = f"{name} {spec}".upper()
    nums = re.findall(r'\b(2[0-9]|3[0-9]|4[0-8])\b', text)
    if nums:
        size = int(nums[-1])
        if 23 <= size <= 34: return 'Kids'
        elif 35 <= size <= 39: return 'Women'
        elif 40 <= size <= 48: return 'Men'
            
    s = " " + text.replace('-', ' ').replace('_', ' ').replace('/', ' ') + " "
    if any(k in s for k in [' KID ', ' KIDS ', ' BOY ', ' BOYS ', ' GIRL ', ' GIRLS ', ' CHILD ', ' CHILDREN ']): return 'Kids'
    if any(k in s for k in [' WOMAN ', ' WOMEN ', ' WOMENS ', ' LADY ', ' LADIES ', ' FEMALE ']): return 'Women'
    if any(k in s for k in [' MAN ', ' MEN ', ' MENS ', ' MALE ']): return 'Men'
    return 'Women'

def load_october_phasing():
    phasing_files = glob.glob(os.path.join(REPORTS_DIR, "*Phasing*.xlsx")) + glob.glob("*Phasing*.xlsx")
    if not phasing_files:
        return {}, {}, {}
    phasing_file = phasing_files[0]
    oct_targets, oct_mtd_targets, oct_today_targets = {}, {}, {}
    try:
        df_p = pd.read_excel(phasing_file, sheet_name=0)
        row_target = df_p.iloc[2].values
        row_stores = df_p.iloc[3].values

        col_store_map = {}
        for c_idx in range(4, len(row_stores) - 1):
            h_text = str(row_stores[c_idx]).replace('\n', ' ').strip()
            c_code = clean_store_code_str(h_text)
            col_store_map[c_idx] = c_code
            full_t = row_target[c_idx]
            if pd.notna(full_t):
                oct_targets[c_code] = float(full_t)

        df_days = df_p.iloc[4:].copy()
        df_days = df_days[df_days.iloc[:, 0].astype(str).str.contains('2026-10-')].copy()

        day_3_row = df_days[df_days.iloc[:, 0].astype(str).str.contains('2026-10-03')]
        day_1_to_3 = df_days[df_days.iloc[:, 0].astype(str).str.contains('2026-10-01|2026-10-02|2026-10-03')]

        for c_idx, c_code in col_store_map.items():
            if not day_3_row.empty:
                oct_today_targets[c_code] = float(day_3_row.iloc[0, c_idx])
            if not day_1_to_3.empty:
                oct_mtd_targets[c_code] = float(day_1_to_3.iloc[:, c_idx].sum())

        return oct_targets, oct_mtd_targets, oct_today_targets
    except Exception as e:
        print(f"[!] Phasing load error: {e}")
        return {}, {}, {}

def load_ly_sales_data(target_date_str="2026-10-03"):
    ly_files = glob.glob(os.path.join(REPORTS_DIR, "*LY*OCT*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*LY*.xlsx")) + glob.glob("*LY*OCT*.xlsx") + glob.glob("*LY*.xlsx")
    if not ly_files: 
        return {}, {}
    
    ly_mtd_totals, ly_today_totals = {}, {}
    try:
        xl = pd.ExcelFile(ly_files[0])
        sheet_to_use = "Sales" if "Sales" in xl.sheet_names else xl.sheet_names[0]
        df_ly = pd.read_excel(ly_files[0], sheet_name=sheet_to_use)
        
        df_ly['Date_ffill'] = df_ly['Unnamed: 0'].ffill()
        df_ly['Date_parsed'] = pd.to_datetime(df_ly['Date_ffill'], errors='coerce')
        
        cur_dt = pd.to_datetime(target_date_str)
        ly_cur_date = cur_dt.replace(year=2025)
        
        df_gsale = df_ly[df_ly['Unnamed: 2'].astype(str).str.strip().str.upper() == 'G-SALE'].copy()
        if df_gsale.empty: df_gsale = df_ly.copy()

        df_ly_mtd = df_gsale[df_gsale['Date_parsed'] <= ly_cur_date]
        df_ly_today = df_gsale[df_gsale['Date_parsed'] == ly_cur_date]

        for col in df_ly.columns[3:]:
            if "TOTAL" not in str(col).upper():
                code = clean_store_code_str(col)
                mtd_val = pd.to_numeric(df_ly_mtd[col], errors='coerce').sum()
                if pd.notna(mtd_val) and mtd_val > 0:
                    ly_mtd_totals[code] = round(float(mtd_val))
                
                if not df_ly_today.empty:
                    today_val = pd.to_numeric(df_ly_today[col], errors='coerce').sum()
                    if pd.notna(today_val):
                        ly_today_totals[code] = round(float(today_val))
                            
        return ly_mtd_totals, ly_today_totals
    except Exception as e:
        print(f"[!] Error reading LY: {e}")
        return {}, {}

def load_soh_data():
    soh_files = glob.glob(os.path.join(REPORTS_DIR, "*SOH*.xlsx")) + glob.glob("*SOH*.xlsx")
    if not soh_files: return {}, {}, pd.DataFrame(), 0
    soh_store_summary, sku_to_cat_map = {}, {}
    wh_total_stock = 0
    try:
        xl = pd.ExcelFile(soh_files[0])
        sheet_to_use = xl.sheet_names[0]
        df_soh = pd.read_excel(soh_files[0], sheet_name=sheet_to_use, skiprows=1)
        if not any(k in [str(c).lower() for c in df_soh.columns] for k in ["avail_stock", "current_stock", "stock"]):
            df_soh = pd.read_excel(soh_files[0], sheet_name=sheet_to_use)

        code_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["org code", "store code", "organization"])), None)
        stock_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["avail_stock", "current_stock", "stock", "qty"])), None)
        barcode_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["barcode", "bar code", "upc"])), None)
        cat_col = next((c for c in df_soh.columns if str(c).lower() == "category"), None)

        if not code_col or not stock_col: return {}, {}, pd.DataFrame(), 0
        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)
        df_soh['clean_code'] = df_soh[code_col].apply(lambda v: "KSWH" if "KSWH" in str(v).upper() else clean_store_code_str(v))

        if barcode_col:
            df_soh['clean_barcode'] = df_soh[barcode_col].apply(clean_sku_code)
            if cat_col:
                for _, r in df_soh[[barcode_col, cat_col]].dropna().drop_duplicates().iterrows():
                    sku_to_cat_map[clean_sku_code(r[barcode_col])] = str(r[cat_col]).strip()

        wh_df = df_soh[df_soh['clean_code'] == 'KSWH']
        wh_total_stock = int(wh_df[stock_col].sum()) if not wh_df.empty else 0

        stores_df = df_soh[df_soh['clean_code'].isin(ALL_VALID_CODES)]
        grouped = stores_df.groupby('clean_code')[stock_col].sum()
        soh_store_summary = grouped.to_dict()

        return soh_store_summary, sku_to_cat_map, df_soh, wh_total_stock
    except Exception:
        return {}, {}, pd.DataFrame(), 0

def process_and_build():
    oct_targets, oct_mtd_targets, oct_today_targets = load_october_phasing()
    ly_mtd_map, ly_today_map = load_ly_sales_data("2026-10-03")
    soh_map, sku_to_cat, df_soh_raw, wh_total_stock = load_soh_data()

    sales_candidates = glob.glob(os.path.join(REPORTS_DIR, "*.xlsx")) + glob.glob("*.xlsx")
    mms_files = [f for f in sales_candidates if "50100002" in f and "ARCHIVE" not in f.upper()]
    dzl_files = [f for f in sales_candidates if "dzl" in f.lower() and "sales" in f.lower() and "ARCHIVE" not in f.upper()]

    dfs = []
    if mms_files:
        df_m = pd.read_excel(mms_files[0], skiprows=1).iloc[:-1].copy()
        df_m.columns = [str(c).strip() for c in df_m.columns]
        df_m['brand_origin'] = "MMS"
        dfs.append(df_m)

    if dzl_files:
        try:
            df_d = pd.read_excel(dzl_files[0])
            if not any("org" in str(c).lower() for c in df_d.columns):
                df_d = pd.read_excel(dzl_files[0], skiprows=1)
            df_d.columns = [str(c).strip() for c in df_d.columns]
            df_d['brand_origin'] = "DZL"
            dfs.append(df_d)
        except Exception: pass

    df_clean = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    for col in ['Sales Quantity', 'Actual Sales Amount']:
        if col in df_clean.columns: df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    org_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['organization code', 'org code', 'store code'])), df_clean.columns[0])
    item_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['product code', 'item code', 'barcode'])), df_clean.columns[0])
    name_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['product name', 'item name'])), item_col)
    txn_col = next((c for c in df_clean.columns if 'receipt' in str(c).lower()), item_col)
    cat_col = next((c for c in df_clean.columns if str(c).lower() in ['category name', 'category']), None)

    df_clean['clean_code'] = df_clean[org_col].apply(clean_store_code_str)
    df_clean = df_clean[df_clean['clean_code'].isin(ALL_VALID_CODES)].copy()
    df_clean['brand'] = df_clean['clean_code'].apply(lambda c: STORE_MAPPING[c]['brand'])
    df_clean['clean_sku'] = df_clean[item_col].apply(clean_sku_code)
    df_clean['clean_name'] = df_clean[name_col].fillna("Item").astype(str)
    df_clean['main_category'] = df_clean[cat_col].fillna("General").astype(str) if cat_col else "General"
    df_clean['gender'] = df_clean.apply(lambda r: classify_shoe_gender_by_size(r['clean_name']), axis=1)

    # تحديد وتوحيد عمود المبيعات والكمية تلقائياً لتفادي KeyError
    col_sales_match = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['actual sales amount', 'actual sales', 'actual amount', 'sales amount', 'amount', 'revenue'])), None)
    col_qty_match = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['sales quantity', 'quantity', 'qty', 'units'])), None)

    if col_sales_match:
        df_clean['Actual Sales Amount'] = pd.to_numeric(df_clean[col_sales_match], errors='coerce').fillna(0)
    else:
        df_clean['Actual Sales Amount'] = 0.0

    if col_qty_match:
        df_clean['Sales Quantity'] = pd.to_numeric(df_clean[col_qty_match], errors='coerce').fillna(0)
    else:
        df_clean['Sales Quantity'] = 0

    df_clean = df_clean[(df_clean['Actual Sales Amount'] > 0) & (df_clean['Sales Quantity'] > 0)].copy()
    for kw in EXCLUDED_KEYWORDS:
        df_clean = df_clean[~df_clean['clean_name'].str.lower().str.contains(kw, regex=False)]

    store_rows_data = []
    for code, info in STORE_MAPPING.items():
        st_df = df_clean[df_clean['clean_code'] == code]
        sales = round(st_df['Actual Sales Amount'].sum())
        units = int(st_df['Sales Quantity'].sum())
        txns = st_df[txn_col].nunique() if txn_col in st_df.columns else len(st_df)

        ly_s = ly_mtd_map.get(code, None)
        yoy = ((sales - ly_s) / ly_s * 100) if ly_s and ly_s > 0 else None

        target = oct_mtd_targets.get(code, oct_targets.get(code, 0))
        ach = (sales / target * 100) if target > 0 else 0

        atv = round(sales / txns) if txns > 0 else 0
        upt = round(units / txns, 2) if txns > 0 else 0
        asp = round(sales / units) if units > 0 else 0

        soh_units = soh_map.get(code, 0)
        str_pct = round((units / (units + soh_units) * 100), 1) if (units + soh_units) > 0 else 0

        diag = "Powerhouse Performer" if ach >= 95 else ("Assortment Mismatch" if ach < 70 and soh_units > 10000 else "Steady Flow")
        diag_col = "#10b981" if diag == "Powerhouse Performer" else ("#f59e0b" if diag == "Assortment Mismatch" else "#38bdf8")

        store_rows_data.append({
            "code": code, "name": info["full_name"], "region": info["region"], "manager": info["manager"], "brand": info["brand"],
            "sales": sales, "ly_sales": ly_s, "yoy": yoy, "target": target, "ach": ach,
            "units": units, "txns": txns, "upt": upt, "atv": atv, "asp": asp, "str_pct": str_pct,
            "diag": diag, "diag_col": diag_col, "soh_units": soh_units
        })

    perf_df = pd.DataFrame(store_rows_data).sort_values(by='sales', ascending=False).reset_index(drop=True)

    store_table_rows = ""
    for idx, r in perf_df.iterrows():
        yoy_str = f'<span style="color:{"#10b981" if r["yoy"]>=0 else "#ef4444"}; font-weight:700;">{r["yoy"]:+.1f}%</span>' if pd.notna(r["yoy"]) else '<span style="color:#64748b;">New Store</span>'
        ly_str = f"{r['ly_sales']:,}" if pd.notna(r['ly_sales']) else '<span style="color:#64748b;">New Store</span>'
        ach_col = "#10b981" if r['ach'] >= 100 else ("#f59e0b" if r['ach'] >= 80 else "#ef4444")
        diag_badge = f'<span class="badge" style="background:{r["diag_col"]}22; color:{r["diag_col"]}; border:1px solid {r["diag_col"]}55;">{r["diag"]}</span>'
        brand_badge = f'<span class="badge" style="background:{"#ef444422" if r["brand"]=="DZL" else "#38bdf822"}; color:{"#ef4444" if r["brand"]=="DZL" else "#38bdf8"};">{r["brand"]}</span>'

        store_table_rows += f"""
        <tr class="clickable-row store-row" data-brand="{r['brand']}" data-region="{r['region']}" onclick="openStoreModal('{r['code']}')">
            <td style="color:#64748b; font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8; font-weight:700;">{r['code']}</td>
            <td style="color:#fff; font-weight:600;">{brand_badge} {r['name']}</td>
            <td style="color:#94a3b8; font-size:12px;">{r['region']}</td>
            <td style="color:#f8fafc; font-weight:700;">{r['sales']:,}</td>
            <td style="color:#38bdf8;">{ly_str}</td>
            <td>{yoy_str}</td>
            <td style="color:#94a3b8;">{round(r['target']):,}</td>
            <td style="color:{ach_col}; font-weight:700;">{r['ach']:.1f}%</td>
            <td style="color:#38bdf8; font-weight:700;">{r['units']:,}</td>
            <td style="color:#fff; font-weight:700;">{r['txns']:,}</td>
            <td style="color:#10b981; font-weight:700;">{r['upt']:.2f}</td>
            <td style="color:#38bdf8; font-weight:700;">{r['str_pct']}</td>
            <td>{diag_badge}</td>
            <td style="color:#f59e0b; font-weight:700;">{r['asp']}</td>
        </tr>
        """

    tot_sales = perf_df['sales'].sum()
    tot_target = perf_df['target'].sum()
    tot_ach = (tot_sales / tot_target * 100) if tot_target > 0 else 0
    tot_units = perf_df['units'].sum()
    tot_txns = perf_df['txns'].sum()
    net_atv = round(tot_sales / tot_txns) if tot_txns > 0 else 0
    net_asp = round(tot_sales / tot_units) if tot_units > 0 else 0
    valid_ly = perf_df[perf_df['ly_sales'].notna()]
    tot_ly = round(valid_ly['ly_sales'].sum())
    tot_yoy = ((valid_ly['sales'].sum() - tot_ly) / tot_ly * 100) if tot_ly > 0 else 0

    repl_rows_html = ""
    repl_export = []
    dzl_shoes = df_clean[(df_clean['brand'] == 'DZL') & (df_clean['clean_name'].str.lower().str.contains('shoe|runner|trainer|sneaker|loafer', regex=True))].copy()
    if not dzl_shoes.empty:
        sku_agg = dzl_shoes.groupby(['clean_code', 'clean_sku', 'clean_name'], as_index=False)['Sales Quantity'].sum()
        for idx, r in sku_agg.head(25).iterrows():
            st_c = r['clean_code']
            st_info = STORE_MAPPING[st_c]
            other_dzl = [c for c in DZL_VALID_CODES if c != st_c]
            donor_code = other_dzl[0]
            donor_info = STORE_MAPPING[donor_code]
            match_type = "🏙️ Same City" if donor_info['city'] == st_info['city'] else "🚛 Inter-City"

            repl_rows_html += f"""
            <tr data-brand="DZL">
                <td style="color:#64748b;">{idx+1}</td>
                <td><span class="badge" style="background:#ef444422; color:#ef4444;">Store Transfer (IST - Opportunity)</span></td>
                <td style="color:#fff; font-weight:700;">{st_info['full_name']} ({st_c})</td>
                <td style="color:#f59e0b; font-weight:600;">👟 [SHOE PRIORITY 1] {r['clean_name'][:30]} (SKU: {r['clean_sku']})</td>
                <td style="color:#38bdf8;">{donor_info['full_name']} ({donor_code}) [{match_type}]</td>
                <td style="color:#10b981; font-weight:700;">{int(r['Sales Quantity'] * 2)} Pcs</td>
                <td><span class="badge" style="background:#ef444422; color:#ef4444;">🚨 Size-Run Recovery</span></td>
            </tr>
            """
            repl_export.append({
                "Brand": "DZL", "Store": st_info['full_name'], "SKU": r['clean_sku'], "Product Name": r['clean_name'],
                "Donor Store": donor_info['full_name'], "Transfer Qty": int(r['Sales Quantity'] * 2), "Type": match_type
            })

    if repl_export:
        try:
            pd.DataFrame(repl_export).to_excel(os.path.join(REPORTS_DIR, "Auto_Replenishment_DZL.xlsx"), index=False)
        except Exception: pass

    # قوائم MMS Top 500 و Low 500
    mms_only = df_clean[df_clean['brand'] == 'MMS'].groupby(['clean_sku', 'clean_name', 'main_category'], as_index=False).agg(
        units=('Sales Quantity', 'sum'), sales=('Actual Sales Amount', 'sum')
    ).sort_values(by='units', ascending=False).reset_index(drop=True)
    mms_only['asp'] = (mms_only['sales'] / mms_only['units']).fillna(0).round().astype(int)
    mms_top500 = mms_only.head(500).to_dict(orient='records')
    mms_low500 = mms_only.tail(500).sort_values(by='units', ascending=True).to_dict(orient='records')

    # قراءة بيانات الفئات والـ Gender Mix
    cat_df = df_clean.groupby('main_category', as_index=False).agg(sales=('Actual Sales Amount', 'sum')).sort_values(by='sales', ascending=False)
    cat_df['contribution'] = (cat_df['sales'] / tot_sales * 100).round(1) if tot_sales > 0 else 0
    category_data = cat_df.to_dict(orient='records')

    dzl_only_shoes = df_clean[(df_clean['brand'] == 'DZL') & (df_clean['clean_name'].str.lower().str.contains('shoe|runner|trainer|sneaker|loafer', regex=True))]
    dzl_g_sales = dzl_only_shoes.groupby('gender', as_index=False).agg(sales=('Actual Sales Amount', 'sum'))
    dzl_g_tot = dzl_g_sales['sales'].sum()
    dzl_g_sales['share'] = (dzl_g_sales['sales'] / dzl_g_tot * 100).round(1) if dzl_g_tot > 0 else 0
    dzl_gender_data = {
        "labels": dzl_g_sales['gender'].tolist(),
        "series": dzl_g_sales['share'].tolist(),
        "records": dzl_g_sales.to_dict(orient='records')
    }

    store_meta_map = {r['code']: r for r in perf_df.to_dict(orient='records')}

    # بناء الصفحة باستخدام نظام القوالب الآمن (Safe String Template) لمنع أخطاء f-string نهائياً
    template_html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS & DZL Executive Commercial Intelligence Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>
    <style>
        :root {
            --bg: #090d16; --card: #131b2e; --border: #1e293b; --primary: #38bdf8; --text-muted: #94a3b8;
        }
        * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
        body { background: var(--bg); color: #fff; margin: 0; padding: 24px; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 24px; flex-wrap: wrap; gap: 16px; }
        .top-controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
        .brand-switcher { display: flex; background: #0c1220; padding: 4px; border-radius: 8px; border: 1px solid var(--border); gap: 4px; }
        .brand-btn { background: transparent; border: none; color: var(--text-muted); padding: 6px 14px; border-radius: 6px; font-size: 13px; font-weight: 700; cursor: pointer; }
        .brand-btn.active { background: #2563eb; color: #fff; }
        .month-select { background: #0c1220; border: 1px solid var(--border); color: #38bdf8; padding: 6px 14px; border-radius: 8px; font-weight: 700; outline: none; }
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 24px; }
        .kpi-card { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }
        .kpi-title { font-size: 11px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; margin-bottom: 6px; }
        .kpi-value { font-size: 22px; font-weight: 700; color: #fff; }
        .view-toggle-bar { display: flex; background: #0c1220; padding: 4px; border-radius: 10px; border: 1px solid var(--border); margin-bottom: 24px; width: fit-content; gap: 4px; }
        .view-btn { background: transparent; border: none; color: var(--text-muted); padding: 10px 22px; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; }
        .view-btn.active { background: #2563eb; color: #fff; }
        .table-wrap { background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 24px; }
        .table-header { padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; }
        th { background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; border-bottom: 1px solid var(--border); white-space: nowrap; }
        td { padding: 12px 14px; border-bottom: 1px solid var(--border); white-space: nowrap; }
        tr:hover td { background: #19233c; }
        .clickable-row { cursor: pointer; }
        .badge { padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
        .sub-tab-btn { background:#1e293b; color:#94a3b8; border:1px solid #334155; padding:8px 16px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; }
        .sub-tab-btn.active { background:#38bdf8; color:#090d16; border-color:#38bdf8; }
        .chart-container { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 22px; margin-bottom: 24px; }
        .app-modal { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(9, 13, 22, 0.9); z-index: 2147483647; display: none; align-items: center; justify-content: center; }
        .modal-content { background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 92%; max-width: 950px; max-height: 90vh; display: flex; flex-direction: column; overflow: hidden; }
        .modal-header { padding: 18px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; background: #0c1220; }
        .modal-body { padding: 24px; overflow-y: auto; }
        .close-btn { background: transparent; border: none; color: #94a3b8; font-size: 28px; cursor: pointer; }
    </style>
</head>
<body>

<div id="auth-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:#090d16;z-index:99999999;display:flex;align-items:center;justify-content:center;">
  <div style="background:#131b2e;padding:32px;border-radius:12px;text-align:center;width:90%;max-width:380px;border:1px solid #1e293b;">
    <h3 style="color:#fff;margin:0 0 8px 0;font-size:20px;">🔒 Executive Secure Access</h3>
    <p style="color:#94a3b8;font-size:13px;margin:0 0 20px 0;">Enter authorization PIN to unlock dashboard</p>
    <input type="password" id="access-pass" placeholder="PIN Code" style="width:100%;padding:12px;border-radius:6px;border:1px solid #334155;background:#090d16;color:#fff;font-size:16px;text-align:center;outline:none;margin-bottom:14px;">
    <button onclick="checkAccess()" style="width:100%;padding:12px;border-radius:6px;border:none;background:#2563eb;color:#fff;font-weight:700;font-size:15px;cursor:pointer;">Unlock Dashboard</button>
    <p id="error-msg" style="color:#ef4444;font-size:13px;margin:12px 0 0 0;display:none;">Invalid authorization credentials</p>
  </div>
</div>

<script>
  const USER_ROLES = {
    "MMS2026": { role: "ADMIN", name: "Executive & Merchandising (Full Access)", region: "ALL" },
    "SULTAN2026": { role: "AREA_MGR", name: "Sultan", region: "Central & Eastern Region" },
    "RAJIB2026": { role: "AREA_MGR", name: "Rajib", region: "Western, Southern & Northern Region" }
  };

  function checkAccess() {
    var val = document.getElementById("access-pass").value.trim().toUpperCase();
    var user = USER_ROLES[val];
    if (user) {
      sessionStorage.setItem("mms_user", JSON.stringify(user));
      document.getElementById("auth-overlay").style.display = "none";
      document.getElementById("current-user-badge").innerHTML = "👤 " + user.name;
    } else {
      document.getElementById("error-msg").style.display = "block";
    }
  }

  function logout() { sessionStorage.removeItem("mms_user"); location.reload(); }

  document.addEventListener("DOMContentLoaded", function() {
    var u = sessionStorage.getItem("mms_user");
    if (u) {
      document.getElementById("auth-overlay").style.display = "none";
      document.getElementById("current-user-badge").innerHTML = "👤 " + JSON.parse(u).name;
    }
    document.getElementById("access-pass").addEventListener("keypress", function(e) {
      if (e.key === "Enter") checkAccess();
    });
  });
</script>

<div id="store-modal" class="app-modal">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h2 id="modal-store-name" style="margin:0; font-size:18px; color:#fff;">Store Performance Analysis</h2>
        <span id="modal-store-code" style="color:#38bdf8; font-size:12px; font-weight:700;">CODE</span>
      </div>
      <button class="close-btn" onclick="closeModal()">&times;</button>
    </div>
    <div class="modal-body">
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:10px; margin-bottom:20px;">
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">CURRENT SALES</div><div id="modal-sales" style="font-size:16px; font-weight:700; color:#fff;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">TARGET (% ACH)</div><div id="modal-target" style="font-size:16px; font-weight:700; color:#10b981;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">SOH STOCK</div><div id="modal-soh" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">ATV</div><div id="modal-atv" style="font-size:16px; font-weight:700; color:#fff;">-</div></div>
      </div>
    </div>
  </div>
</div>

<div class="header">
    <div>
        <h1 style="margin:0; font-size:22px;">MMS & DZL Executive Commercial Intelligence Dashboard</h1>
        <p style="margin:4px 0 0 0; color:var(--text-muted); font-size:13px;">October 2026 Daily Phasing & Commercial Performance Tracking</p>
    </div>
    <div class="top-controls">
        <select class="month-select" id="monthDropdown" onchange="switchMonth(this.value)">
            <option value="OCT" selected>📅 October 2026 (Active)</option>
            <option value="SEP">📅 September 2026 (Archived)</option>
        </select>
        <div class="brand-switcher">
            <button class="brand-btn active" id="btn-ALL" onclick="switchBrand('ALL')">🏢 ALL BRANDS</button>
            <button class="brand-btn" id="btn-MMS" onclick="switchBrand('MMS')">🔴 MUMUSO (17)</button>
            <button class="brand-btn" id="btn-DZL" onclick="switchBrand('DZL')">🟡 DZL (3)</button>
        </div>
        <span id="current-user-badge" style="font-size:13px; font-weight:700; color:#38bdf8; background:#1e293b; padding:8px 14px; border-radius:8px;">👤 Authenticating..</span>
        <button onclick="logout()" style="background:#ef444422; border:1px solid #ef444455; color:#ef4444; padding:8px 14px; border-radius:8px; font-weight:700; cursor:pointer;">Logout</button>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Current Total Sales</div><div class="kpi-value">__TOT_SALES__ <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">LY Gross Sales</div><div class="kpi-value" style="color:#38bdf8;">__TOT_LY__ <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">Network LFL YoY Growth</div><div class="kpi-value" style="color:__TOT_YOY_COL__; font-weight:800;">__TOT_YOY__</div></div>
    <div class="kpi-card"><div class="kpi-title">Total Target</div><div class="kpi-value">__TOT_TARGET__ <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">Achievement (% Ach)</div><div class="kpi-value" style="color:__TOT_ACH_COL__;">__TOT_ACH__</div></div>
    <div class="kpi-card"><div class="kpi-title">Network ATV</div><div class="kpi-value">SAR __NET_ATV__</div></div>
    <div class="kpi-card"><div class="kpi-title">Network ASP</div><div class="kpi-value">SAR __NET_ASP__</div></div>
</div>

<div class="view-toggle-bar">
    <button class="view-btn active" id="btn-stores" onclick="switchView('stores')">🏢 Store Matrix</button>
    <button class="view-btn" id="btn-regions" onclick="switchView('regions')">🌍 Region-Wise</button>
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business & Gender</button>
    <button class="view-btn" id="btn-action" onclick="switchView('action')">⚡ Commercial Action Hub</button>
</div>

<!-- 1. Store Commercial Matrix View -->
<div id="view-stores">
    <div class="table-wrap">
        <div class="table-header">
            <h3 style="margin:0; font-size:15px; color:#fff;">STORE COMMERCIAL & DISPLAY ASSORTMENT MATRIX</h3>
            <input type="text" id="storeSearch" placeholder="Search store..." onkeyup="filterStores()" style="background:#090d16; border:1px solid var(--border); color:#fff; padding:6px 12px; border-radius:6px;">
        </div>
        <div style="overflow-x:auto;">
            <table id="storesTable">
                <thead>
                    <tr>
                        <th>#</th><th>Code</th><th>Store Name</th><th>Region</th><th>Sales (SAR)</th><th>LY Sales</th><th>YoY</th>
                        <th>Target</th><th>% Ach</th><th>Units</th><th>Trans</th><th>UPT</th><th>STR%</th><th>Diagnostic</th><th>ASP</th>
                    </tr>
                </thead>
                <tbody>
                    __STORE_TABLE_ROWS__
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 2. Region-Wise View -->
<div id="view-regions" style="display:none;">
    <div class="table-wrap">
        <div class="table-header"><h3 style="margin:0; font-size:15px; color:#38bdf8;">🏢 CENTRAL & EASTERN REGION (Sultan - 11 Doors)</h3></div>
        <div style="overflow-x:auto;">
            <table id="regionCentralTable">
                <thead><tr><th>#</th><th>Code</th><th>Store Name</th><th>Sales</th><th>Target</th><th>% Ach</th><th>Units</th><th>Trans</th></tr></thead>
                <tbody></tbody>
            </table>
        </div>
    </div>

    <div class="table-wrap">
        <div class="table-header"><h3 style="margin:0; font-size:15px; color:#818cf8;">🏢 WESTERN, SOUTHERN & NORTHERN REGION (Rajib - 12 Doors)</h3></div>
        <div style="overflow-x:auto;">
            <table id="regionWesternTable">
                <thead><tr><th>#</th><th>Code</th><th>Store Name</th><th>Sales</th><th>Target</th><th>% Ach</th><th>Units</th><th>Trans</th></tr></thead>
                <tbody></tbody>
            </table>
        </div>
    </div>
</div>

<!-- 3. Business & Gender Mix View -->
<div id="view-business" style="display:none;">
    <div style="display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px;">
        <div class="chart-container" style="flex:1; min-width:320px;">
            <div style="font-weight:700; margin-bottom:12px; font-size:14px;">🍩 Category Contribution Share</div>
            <div id="apexCategoryDonut" style="min-height: 330px;"></div>
        </div>
        <div class="chart-container" id="genderChartWrapper" style="flex:1; min-width:320px;">
            <div style="font-weight:700; margin-bottom:12px; font-size:14px;">👟 Footwear Gender Mix Share (Kids 23-34 | Women 35-39 | Men 40-48)</div>
            <div id="apexGenderDonut" style="min-height: 330px;"></div>
        </div>
    </div>
</div>

<!-- 4. Commercial Action Hub -->
<div id="view-action" style="display:none;">
    <div class="table-wrap" style="margin-bottom:24px;">
        <div class="table-header">
            <h3 style="margin:0; font-size:15px; color:#10b981;">⚡ PREDICTIVE AUTO-REPLENISHMENT & IST OPPORTUNITY ROUTING</h3>
            <a href="./reports/Auto_Replenishment_DZL.xlsx" download style="background:#10b981; color:#fff; padding:6px 14px; border-radius:6px; text-decoration:none; font-weight:700; font-size:12px;">📥 Download Plan</a>
        </div>
        <div style="max-height:350px; overflow-y:auto;">
            <table>
                <thead><tr><th>#</th><th>Action Type</th><th>Store Target</th><th>SKU Focus</th><th>Source Route</th><th>Quantity</th><th>Urgency</th></tr></thead>
                <tbody>__REPL_ROWS_HTML__</tbody>
            </table>
        </div>
    </div>

    <div class="table-wrap">
        <div class="table-header">
            <div style="display:flex; gap:10px; align-items:center;">
                <button class="sub-tab-btn active" id="btn-top500" onclick="switchMovers('top')">🔥 MMS TOP 500 FAST MOVERS</button>
                <button class="sub-tab-btn" id="btn-low500" onclick="switchMovers('low')">❄️ MMS LOW 500 CLEARANCE CANDIDATES</button>
            </div>
            <input type="text" id="moversSearch" placeholder="Search SKU..." onkeyup="filterMovers()" style="background:#090d16; border:1px solid var(--border); color:#fff; padding:6px 12px; border-radius:6px;">
        </div>
        <div style="max-height:400px; overflow-y:auto;">
            <table id="moversTable">
                <thead><tr><th>Rank</th><th>SKU Code</th><th>Product Name</th><th>Category</th><th>Units Sold</th><th>Sales (SAR)</th><th>ASP</th></tr></thead>
                <tbody id="moversTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script>
  let activeBrand = 'ALL';
  const STORE_META = __STORE_META_JSON__;
  const MMS_TOP500 = __MMS_TOP500_JSON__;
  const MMS_LOW500 = __MMS_LOW500_JSON__;
  const CATEGORY_DATA = __CATEGORY_DATA_JSON__;
  const DZL_GENDER = __DZL_GENDER_JSON__;
  let moversType = 'top';
  let donutChart = null;
  let genderChart = null;

  function switchBrand(b) {
    activeBrand = b;
    document.querySelectorAll('.brand-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById('btn-' + b).classList.add('active');

    document.querySelectorAll('.store-row').forEach(row => {
      const rBrand = row.getAttribute('data-brand');
      row.style.display = (b === 'ALL' || rBrand === b) ? '' : 'none';
    });
    renderRegionTables();
  }

  function filterStores() {
    const q = document.getElementById('storeSearch').value.toLowerCase();
    document.querySelectorAll('.store-row').forEach(row => {
      const text = row.innerText.toLowerCase();
      const rBrand = row.getAttribute('data-brand');
      const matchBrand = (activeBrand === 'ALL' || rBrand === activeBrand);
      row.style.display = (matchBrand && text.includes(q)) ? '' : 'none';
    });
  }

  function renderRegionTables() {
    const centralTbody = document.querySelector("#regionCentralTable tbody");
    const westernTbody = document.querySelector("#regionWesternTable tbody");
    let cRows = "", wRows = "";
    let cIdx = 1, wIdx = 1;

    Object.values(STORE_META).forEach(r => {
      if (activeBrand !== 'ALL' && r.brand !== activeBrand) return;
      const rowHtml = `<tr>
        <td style="color:#64748b;">${r.region.includes('Central') ? cIdx++ : wIdx++}</td>
        <td style="color:#38bdf8; font-weight:700;">${r.code}</td>
        <td style="color:#fff;">${r.name}</td>
        <td style="color:#fff; font-weight:700;">${r.sales.toLocaleString()}</td>
        <td style="color:#94a3b8;">${Math.round(r.target).toLocaleString()}</td>
        <td style="color:${r.ach >= 100 ? '#10b981' : '#f59e0b'}; font-weight:700;">${r.ach.toFixed(1)}%</td>
        <td style="color:#38bdf8;">${r.units.toLocaleString()}</td>
        <td>${r.txns.toLocaleString()}</td>
      </tr>`;
      if (r.region.includes('Central')) cRows += rowHtml;
      else wRows += rowHtml;
    });

    if (centralTbody) centralTbody.innerHTML = cRows;
    if (westernTbody) westernTbody.innerHTML = wRows;
  }

  function renderCharts() {
    const colors = ['#38bdf8', '#f59e0b', '#10b981', '#ec4899', '#818cf8', '#a855f7', '#06b6d4'];
    const catSeries = CATEGORY_DATA.map(c => c.contribution);
    const catLabels = CATEGORY_DATA.map(c => c.main_category);

    const donutEl = document.querySelector("#apexCategoryDonut");
    if (donutEl) {
      if (donutChart) donutChart.destroy();
      donutChart = new ApexCharts(donutEl, {
        series: catSeries, labels: catLabels, chart: { type: 'donut', height: 330, background: 'transparent' },
        theme: { mode: 'dark' }, colors: colors, legend: { position: 'bottom', labels: { colors: '#cbd5e1' } }
      });
      donutChart.render();
    }

    const genderEl = document.querySelector("#apexGenderDonut");
    if (genderEl) {
      if (genderChart) genderChart.destroy();
      genderChart = new ApexCharts(genderEl, {
        series: DZL_GENDER.series, labels: DZL_GENDER.labels, chart: { type: 'donut', height: 330, background: 'transparent' },
        theme: { mode: 'dark' }, colors: ['#ec4899', '#38bdf8', '#10b981'], legend: { position: 'bottom', labels: { colors: '#cbd5e1' } }
      });
      genderChart.render();
    }
  }

  function switchView(viewName) {
    document.getElementById("view-stores").style.display = (viewName === 'stores') ? 'block' : 'none';
    document.getElementById("view-regions").style.display = (viewName === 'regions') ? 'block' : 'none';
    document.getElementById("view-business").style.display = (viewName === 'business') ? 'block' : 'none';
    document.getElementById("view-action").style.display = (viewName === 'action') ? 'block' : 'none';
    document.querySelectorAll('.view-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById('btn-' + viewName).classList.add('active');

    if (viewName === 'regions') renderRegionTables();
    if (viewName === 'business') {
      setTimeout(() => {
        renderCharts();
        window.dispatchEvent(new Event('resize'));
      }, 50);
    }
    if (viewName === 'action') renderMovers();
  }

  function openStoreModal(code) {
    const r = STORE_META[code];
    if (!r) return;
    document.getElementById("modal-store-name").innerText = "[" + r.brand + "] " + r.name;
    document.getElementById("modal-store-code").innerText = "CODE: " + code + " | " + r.region;
    document.getElementById("modal-sales").innerText = r.sales.toLocaleString() + " SAR";
    document.getElementById("modal-target").innerText = Math.round(r.target).toLocaleString() + " SAR (" + r.ach.toFixed(1) + "%)";
    document.getElementById("modal-soh").innerText = r.soh_units.toLocaleString() + " Pcs";
    document.getElementById("modal-atv").innerText = r.atv.toLocaleString() + " SAR";
    document.getElementById("store-modal").style.display = "flex";
  }

  function closeModal() { document.getElementById("store-modal").style.display = "none"; }

  function switchMovers(type) {
    moversType = type;
    document.getElementById("btn-top500").classList.toggle('active', type === 'top');
    document.getElementById("btn-low500").classList.toggle('active', type === 'low');
    renderMovers();
  }

  function renderMovers() {
    const data = (moversType === 'top') ? MMS_TOP500 : MMS_LOW500;
    const q = document.getElementById("moversSearch").value.toLowerCase();
    const tbody = document.getElementById("moversTableBody");
    let html = "";
    data.filter(r => r.clean_sku.toLowerCase().includes(q) || r.clean_name.toLowerCase().includes(q)).slice(0, 50).forEach((r, idx) => {
      html += `<tr>
        <td style="color:${moversType === 'top' ? '#10b981' : '#ef4444'};">#${idx+1}</td>
        <td style="color:#38bdf8; font-weight:700;">${r.clean_sku}</td>
        <td style="color:#fff;">${r.clean_name}</td>
        <td style="color:#94a3b8;">${r.main_category}</td>
        <td style="color:#38bdf8; font-weight:700;">${r.units.toLocaleString()}</td>
        <td style="color:#fff; font-weight:700;">${Math.round(r.sales).toLocaleString()}</td>
        <td style="color:#f59e0b;">${r.asp}</td>
      </tr>`;
    });
    tbody.innerHTML = html;
  }

  function filterMovers() { renderMovers(); }

  function switchMonth(m) {
    if (m === "SEP") {
      alert("Loading September 2026 Archived monthly benchmarks.");
    } else {
      location.reload();
    }
  }

  renderRegionTables();
</script>

</body>
</html>
"""

    # استبدال المتغيرات بأمان كامل وبدون أي f-string
    final_html = template_html.replace("__TOT_SALES__", f"{tot_sales:,}")
    final_html = final_html.replace("__TOT_LY__", f"{tot_ly:,}")
    final_html = final_html.replace("__TOT_YOY__", f"{tot_yoy:+.1f}%")
    final_html = final_html.replace("__TOT_YOY_COL__", "#10b981" if tot_yoy >= 0 else "#ef4444")
    final_html = final_html.replace("__TOT_TARGET__", f"{round(tot_target):,}")
    final_html = final_html.replace("__TOT_ACH__", f"{tot_ach:.1f}%")
    final_html = final_html.replace("__TOT_ACH_COL__", "#10b981" if tot_ach >= 100 else "#f59e0b")
    final_html = final_html.replace("__NET_ATV__", f"{net_atv:,}")
    final_html = final_html.replace("__NET_ASP__", f"{net_asp:,}")
    final_html = final_html.replace("__STORE_TABLE_ROWS__", store_table_rows)
    final_html = final_html.replace("__REPL_ROWS_HTML__", repl_rows_html)
    final_html = final_html.replace("__STORE_META_JSON__", json.dumps(store_meta_map))
    final_html = final_html.replace("__MMS_TOP500_JSON__", json.dumps(mms_top500))
    final_html = final_html.replace("__MMS_LOW500_JSON__", json.dumps(mms_low500))
    final_html = final_html.replace("__CATEGORY_DATA_JSON__", json.dumps(category_data))
    final_html = final_html.replace("__DZL_GENDER_JSON__", json.dumps(dzl_gender_data))

    out_file = os.path.join(REPORTS_DIR, "MMS_Executive_KPI_Dashboard.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(final_html)

    print(f"[✓] Dashboard generated successfully without syntax errors: {out_file}")

if __name__ == "__main__":
    process_and_build()