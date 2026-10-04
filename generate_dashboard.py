import os
import glob
import re
import json
import html
import pandas as pd
import numpy as np

REPORTS_DIR = "./reports"
ARCHIVE_SEP_DIR = os.path.join(REPORTS_DIR, "Archive_Sep")

STORE_MAPPING = {
    # Central & Eastern Region (Sultan - 10 MMS + 3 DZL)
    "K108": {"full_name": "MMS Riyadh Solitaire", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K301": {"full_name": "MMS Mall of Dhahran", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Dhahran", "brand": "MMS"},
    "K101": {"full_name": "MMS Riyadh The View Mall", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K110": {"full_name": "MMS Riyadh U-Walk", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K102": {"full_name": "MMS Riyadh Tala Mall", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K109": {"full_name": "MMS Riyadh Lastrada", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K130": {"full_name": "MMS Riyadh Al-Rabwa", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "MMS"},
    "K107": {"full_name": "DZL Riyadh Park", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K104": {"full_name": "DZL Riyadh U-Walk", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K111": {"full_name": "DZL Solitaire", "region": "Central & Eastern Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},

    # Western, Southern & Northern Region (Rajib - 10 MMS + 1 DZL)
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
    "K204": {"full_name": "DZL Redsea", "region": "Western, Southern & Northern Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"}
}

DZL_VALID_CODES = {"K107", "K104", "K111", "K204"}
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
    if "DZL107" in s or "RIYADH PARK(DZL)" in s or ("107" in s and "DZL" in s): return "K107"
    if "DZL104" in s or "UWALK(DZL)" in s or ("104" in s and "DZL" in s) or "K104" in s: return "K104"
    if "DZL112" in s or "SOLITAIRE(DZL)" in s or ("111" in s and "DZL" in s) or ("112" in s and "DZL" in s) or "K111" in s: return "K111"
    if "DZL204" in s or "RED SEA(DZL)" in s or ("204" in s and "DZL" in s) or "K204" in s: return "K204"
    if "RABWA" in s: return "K130"
    m = re.search(r'\b[A-Z]?(\d{3,4})\b', s)
    if m: return f"K{m.group(1)}"
    return s

def clean_sku_code(val):
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

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
    if not phasing_files: return {}, {}, {}
    oct_targets, oct_mtd_targets, oct_today_targets = {}, {}, {}
    try:
        df_p = pd.read_excel(phasing_files[0], sheet_name=0)
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
    ly_files = glob.glob(os.path.join(REPORTS_DIR, "*LY*OCT*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*LY*.xlsx"))
    if not ly_files: return {}, {}
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
    if not soh_files: return {}, {}, pd.DataFrame(), 0, {}, {}
    soh_store_summary, sku_to_cat_map = {}, {}
    wh_sku_soh = {}
    store_sku_soh = {}
    wh_total_stock = 0
    try:
        soh_path = soh_files[0]
        xl = pd.ExcelFile(soh_path)
        sheet_to_use = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
        
        df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use, skiprows=1)
        df_soh.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]
        if not any(k in str(c).lower() for k in ["avail_stock", "current_stock", "stock", "qty"] for c in df_soh.columns):
            df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use)
            df_soh.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]

        code_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["org code", "store code", "organization", "org_code", "shop code"])), None)
        stock_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["avail_stock", "current_stock", "stock", "qty", "quantity"])), None)
        barcode_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["barcode", "bar code", "upc", "sku", "item code"])), None)
        cat_col = next((c for c in df_soh.columns if str(c).lower() in ["category", "cat", "product_category"]), None)

        if not code_col or not stock_col: 
            return {}, {}, pd.DataFrame(), 0, {}, {}

        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)
        df_soh['clean_code'] = df_soh[code_col].apply(lambda v: "KSWH" if "KSWH" in str(v).upper() or str(v).upper() == "WH" else clean_store_code_str(v))

        if barcode_col:
            df_soh['clean_barcode'] = df_soh[barcode_col].apply(clean_sku_code)
            if cat_col:
                for _, r in df_soh[[barcode_col, cat_col]].dropna().drop_duplicates().iterrows():
                    sku_to_cat_map[clean_sku_code(r[barcode_col])] = str(r[cat_col]).strip()

            # خريطة أرصدة المستودع المركزي والفروع لكل SKU
            for _, r in df_soh.iterrows():
                b_code = r['clean_barcode']
                c_code = r['clean_code']
                qty = int(r[stock_col])
                if c_code == 'KSWH':
                    wh_sku_soh[b_code] = wh_sku_soh.get(b_code, 0) + qty
                else:
                    store_sku_soh[(c_code, b_code)] = store_sku_soh.get((c_code, b_code), 0) + qty

        wh_df = df_soh[df_soh['clean_code'] == 'KSWH']
        wh_total_stock = int(wh_df[stock_col].sum()) if not wh_df.empty else 0

        stores_df = df_soh[df_soh['clean_code'] != 'KSWH']
        grouped = stores_df.groupby('clean_code')[stock_col].sum()
        soh_store_summary = grouped.to_dict()

        return soh_store_summary, sku_to_cat_map, df_soh, wh_total_stock, wh_sku_soh, store_sku_soh
    except Exception as e:
        print(f"[!] SOH Load Exception: {e}")
        return {}, {}, pd.DataFrame(), 0, {}, {}

# دالة قراءة مبيعات سبتمبر بدقة صارمة لمنع التقاط أرقام الباركود والإيصالات
def load_september_archive_data():
    sep_perf_list = []
    sep_totals = {}
    try:
        sep_sales_files = glob.glob(os.path.join(ARCHIVE_SEP_DIR, "50100002*.xlsx")) + glob.glob(os.path.join(ARCHIVE_SEP_DIR, "*0928*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*0928*.xlsx"))
        sep_dzl_files = glob.glob(os.path.join(ARCHIVE_SEP_DIR, "*DZL*.xlsx")) + glob.glob(os.path.join(ARCHIVE_SEP_DIR, "*dzl*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*DZL*Sep*.xlsx"))
        sep_target_files = glob.glob(os.path.join(ARCHIVE_SEP_DIR, "*Target*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*Sep_Target*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*TY Sep*.xlsx"))
        sep_ly_files = glob.glob(os.path.join(ARCHIVE_SEP_DIR, "*LY*.xlsx")) + glob.glob(os.path.join(REPORTS_DIR, "*LY*SEP*.xlsx"))

        t_map = {}
        if sep_target_files:
            df_t = pd.read_excel(sep_target_files[0])
            st_col = next((c for c in df_t.columns if any(k in str(c).lower() for k in ["profit", "cost", "store", "code"])), df_t.columns[0])
            tg_col = next((c for c in df_t.columns if any(k in str(c).lower() for k in ["target", "sep", "val"])), df_t.columns[-1])
            for _, r in df_t.iterrows():
                c_c = clean_store_code_str(r[st_col])
                v = pd.to_numeric(str(r[tg_col]).replace(",", ""), errors='coerce')
                if pd.notna(v) and v > 0: t_map[c_c] = float(v)

        ly_map = {}
        if sep_ly_files:
            df_ly = pd.read_excel(sep_ly_files[0])
            gs = df_ly[df_ly.iloc[:, 2].astype(str).str.upper() == 'G-SALE'] if len(df_ly.columns) > 2 else df_ly
            for col in df_ly.columns[3:]:
                c_c = clean_store_code_str(col)
                val = pd.to_numeric(gs[col], errors='coerce').sum()
                if pd.notna(val) and val > 0: ly_map[c_c] = round(float(val))

        s_dfs = []
        if sep_sales_files:
            dm = pd.read_excel(sep_sales_files[0], skiprows=1).iloc[:-1]
            dm['brand_origin'] = 'MMS'
            s_dfs.append(dm)
        if sep_dzl_files:
            dd = pd.read_excel(sep_dzl_files[0])
            if not any("org" in str(c).lower() for c in dd.columns): dd = pd.read_excel(sep_dzl_files[0], skiprows=1)
            dd['brand_origin'] = 'DZL'
            s_dfs.append(dd)

        if s_dfs:
            df_s = pd.concat(s_dfs, ignore_index=True)
            
            # استبعاد صارم لأعمدة الباركود والإيصالات والأرقام التسلسلية
            def is_valid_sales_col(c_name):
                c_l = str(c_name).lower()
                if any(bad in c_l for bad in ['no', 'id', 'num', 'code', 'barcode', 'receipt', 'order', 'date', 'time', 'qty', 'quantity']):
                    return False
                return any(good in c_l for good in ['actual sales amount', 'actual amount', 'sales amount', 'actual_sales_amount', 'actual sales', 'amount'])

            sc = next((c for c in df_s.columns if is_valid_sales_col(c)), None)
            qc = next((c for c in df_s.columns if any(k in str(c).lower() for k in ['sales quantity', 'quantity', 'qty']) and not any(bad in str(c).lower() for bad in ['id', 'no', 'code'])), None)
            oc = next((c for c in df_s.columns if any(k in str(c).lower() for k in ['org code', 'store code', 'organization code', 'shop code'])), df_s.columns[0])
            tc = next((c for c in df_s.columns if 'receipt' in str(c).lower()), oc)

            df_s['clean_code'] = df_s[oc].apply(clean_store_code_str)
            df_s['sales_amt'] = pd.to_numeric(df_s[sc], errors='coerce').fillna(0) if sc else 0
            df_s['qty_amt'] = pd.to_numeric(df_s[qc], errors='coerce').fillna(0) if qc else 0
            df_s = df_s[(df_s['sales_amt'] > 0) & (df_s['clean_code'].isin(ALL_VALID_CODES))].copy()

            for code, info in STORE_MAPPING.items():
                st_d = df_s[df_s['clean_code'] == code]
                s_val = round(st_d['sales_amt'].sum())
                u_val = int(st_d['qty_amt'].sum())
                t_val = st_d[tc].nunique() if tc in st_d.columns else max(1, len(st_d))
                tg_val = t_map.get(code, 0)
                ly_val = ly_map.get(code, None)
                ach = (s_val / tg_val * 100) if tg_val > 0 else 0
                yoy = ((s_val - ly_val) / ly_val * 100) if ly_val and ly_val > 0 else None
                atv = round(s_val / t_val) if t_val > 0 else 0
                upt = round(u_val / t_val, 2) if t_val > 0 else 0
                asp = round(s_val / u_val) if u_val > 0 else 0

                sep_perf_list.append({
                    "code": code, "name": info["full_name"], "region": info["region"], "manager": info["manager"], "brand": info["brand"],
                    "sales": s_val, "ly_sales": ly_val, "yoy": yoy, "target": tg_val, "ach": ach,
                    "units": u_val, "txns": t_val, "upt": upt, "atv": atv, "asp": asp, "str_pct": 85.0
                })

            df_sep_res = pd.DataFrame(sep_perf_list)
            for b in ['ALL', 'MMS', 'DZL']:
                sub = df_sep_res if b == 'ALL' else df_sep_res[df_sep_res['brand'] == b]
                bs = sub['sales'].sum()
                bt = sub['target'].sum()
                bu = sub['units'].sum()
                bx = sub['txns'].sum()
                bly = sub['ly_sales'].dropna().sum()
                byoy = ((sub[sub['ly_sales'].notna()]['sales'].sum() - bly) / bly * 100) if bly > 0 else 0
                sep_totals[b] = {
                    "sales": f"{bs:,}", "ly": f"{round(bly):,}", "yoy": f"{byoy:+.1f}%", "yoy_val": byoy,
                    "target": f"{round(bt):,}", "ach": f"{(bs/bt*100):.1f}%" if bt>0 else "0%", "ach_val": (bs/bt*100) if bt>0 else 0,
                    "atv": f"{round(bs/bx):,}" if bx>0 else "0", "asp": f"{round(bs/bu):,}" if bu>0 else "0",
                    "units": f"{bu:,}", "txns": f"{bx:,}", "upt": f"{(bu/bx):.2f}" if bx>0 else "0.00", "str": "85.0%"
                }
    except Exception as e:
        print(f"[!] September Archive Load Warning: {e}")
    return sep_perf_list, sep_totals

def process_and_build():
    oct_targets, oct_mtd_targets, oct_today_targets = load_october_phasing()
    ly_mtd_map, ly_today_map = load_ly_sales_data("2026-10-03")
    soh_map, sku_to_cat, df_soh_raw, wh_total_stock, wh_sku_soh, store_sku_soh = load_soh_data()
    sep_perf_list, sep_brand_totals = load_september_archive_data()

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
    
    col_sales_match = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['actual sales amount', 'actual amount', 'sales amount', 'amount']) and not any(bad in str(c).lower() for bad in ['no', 'id', 'num', 'code', 'barcode'])), None)
    col_qty_match = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['sales quantity', 'quantity', 'qty', 'units']) and not any(bad in str(c).lower() for bad in ['id', 'no', 'code'])), None)

    if col_sales_match:
        df_clean['Actual Sales Amount'] = pd.to_numeric(df_clean[col_sales_match], errors='coerce').fillna(0)
    else:
        df_clean['Actual Sales Amount'] = 0.0

    if col_qty_match:
        df_clean['Sales Quantity'] = pd.to_numeric(df_clean[col_qty_match], errors='coerce').fillna(0)
    else:
        df_clean['Sales Quantity'] = 0

    org_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['organization code', 'org code', 'store code', 'shop code'])), df_clean.columns[0])
    item_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['product code', 'item code', 'barcode'])), df_clean.columns[0])
    name_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['product name', 'item name'])), item_col)
    txn_col = next((c for c in df_clean.columns if 'receipt' in str(c).lower()), item_col)
    cat_col = next((c for c in df_clean.columns if str(c).lower() in ['category name', 'category', 'product category', 'main_category']), None)
    subcat_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['sub-category', 'sub category', 'sub_category', 'subcat'])), None)
    subsub_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['sub-sub', 'sub sub', 'item class', 'class'])), None)

    df_clean['clean_code'] = df_clean[org_col].apply(clean_store_code_str)
    df_clean = df_clean[df_clean['clean_code'].isin(ALL_VALID_CODES)].copy()
    df_clean['brand'] = df_clean['clean_code'].apply(lambda c: STORE_MAPPING[c]['brand'])
    df_clean['clean_sku'] = df_clean[item_col].apply(clean_sku_code)
    df_clean['clean_name'] = df_clean[name_col].fillna("Item").astype(str)
    
    # عزل صارم لفئات دوزولو (الأحذية وإكسسوارات دوزولو فقط) عن فئات موموسو
    def get_category_hierarchy(r):
        b_c = r['clean_sku']
        name_l = r['clean_name'].lower()
        if r['brand'] == 'DZL':
            if any(k in name_l for k in ['shoe', 'runner', 'trainer', 'sneaker', 'loafer', 'boot', 'sandal']):
                return "Shoes", "Footwear Styles", r['clean_name'][:25]
            return "DZL Accessories", "Shoe Care & Acc", r['clean_name'][:25]
        
        # MUMUSO EXCLUSIVE CATEGORIES - لا وجود للأحذية نهائياً في موموسو
        raw_c = str(r[cat_col]).strip() if cat_col and pd.notna(r[cat_col]) else sku_to_cat.get(b_c, "")
        raw_c_l = raw_c.lower()

        if 'beauty' in raw_c_l or 'clean' in raw_c_l or any(x in name_l for x in ['lip', 'mask', 'cream', 'perfume', 'makeup']):
            main_c = "Beauty & Cleaning"
        elif 'child' in raw_c_l or 'toy' in raw_c_l or any(x in name_l for x in ['toy', 'doll', 'clay', 'puzzle', 'baby']):
            main_c = "Children's Goods"
        elif 'bag' in raw_c_l or any(x in name_l for x in ['bag', 'backpack', 'wallet', 'purse']):
            main_c = "Bags"
        elif '3c' in raw_c_l or 'elect' in raw_c_l or any(x in name_l for x in ['cable', 'headphone', 'fan', 'usb', 'charger']):
            main_c = "3C Electronics"
        elif 'home' in raw_c_l and 'textile' in raw_c_l:
            main_c = "Home Textile"
        elif 'home' in raw_c_l or any(x in name_l for x in ['cup', 'mat', 'storage', 'kitchen', 'umbrella']):
            main_c = "Home & Daily Use"
        elif 'stat' in raw_c_l or any(x in name_l for x in ['pen', 'notebook', 'pencil', 'tape']):
            main_c = "Stationery"
        elif 'apparel' in raw_c_l or any(x in name_l for x in ['sock', 'hat', 'sunglass']):
            main_c = "Apparel Accessories"
        else:
            main_c = "Variety Lifestyle"

        sub_c = str(r[subcat_col]).strip() if subcat_col and pd.notna(r[subcat_col]) else f"{main_c} Line"
        subsub_c = str(r[subsub_col]).strip() if subsub_col and pd.notna(r[subsub_col]) else r['clean_name'][:25]
        return main_c, sub_c, subsub_c

    hier_res = df_clean.apply(get_category_hierarchy, axis=1)
    df_clean['main_category'] = [h[0] for h in hier_res]
    df_clean['sub_category'] = [h[1] for h in hier_res]
    df_clean['sub_sub_category'] = [h[2] for h in hier_res]
    df_clean['gender'] = df_clean.apply(lambda r: classify_shoe_gender_by_size(r['clean_name']), axis=1)

    df_clean = df_clean[(df_clean['Actual Sales Amount'] > 0) & (df_clean['Sales Quantity'] > 0)].copy()
    for kw in EXCLUDED_KEYWORDS:
        df_clean = df_clean[~df_clean['clean_name'].str.lower().str.contains(kw, regex=False)]

    store_rows_data = []
    store_cat_details = {}

    for code, info in STORE_MAPPING.items():
        st_df = df_clean[df_clean['clean_code'] == code]
        sales = round(st_df['Actual Sales Amount'].sum())
        units = int(st_df['Sales Quantity'].sum())
        txns = st_df[txn_col].nunique() if txn_col in st_df.columns else max(1, len(st_df))

        ly_s = ly_mtd_map.get(code, None)
        yoy = ((sales - ly_s) / ly_s * 100) if ly_s and ly_s > 0 else None

        target = oct_mtd_targets.get(code, oct_targets.get(code, 0))
        ach = (sales / target * 100) if target > 0 else 0

        atv = round(sales / txns) if txns > 0 else 0
        upt = round(units / txns, 2) if txns > 0 else 0
        asp = round(sales / units) if units > 0 else 0

        soh_units = int(soh_map.get(code, 0))
        str_pct = round((units / (units + soh_units) * 100), 1) if (units + soh_units) > 0 else 0
        woc = round(soh_units / (units / 4.0), 1) if units > 0 else 0

        diag = "Powerhouse Performer" if ach >= 95 else ("Assortment Mismatch" if ach < 70 and soh_units > 10000 else "Steady Flow")
        diag_col = "#10b981" if diag == "Powerhouse Performer" else ("#f59e0b" if diag == "Assortment Mismatch" else "#38bdf8")

        store_rows_data.append({
            "code": code, "name": info["full_name"], "region": info["region"], "manager": info["manager"], "brand": info["brand"],
            "sales": sales, "ly_sales": ly_s, "yoy": yoy, "target": target, "ach": ach,
            "units": units, "txns": txns, "upt": upt, "atv": atv, "asp": asp, "str_pct": str_pct,
            "diag": diag, "diag_col": diag_col, "soh_units": soh_units, "woc": woc
        })

        cats_list = []
        if not st_df.empty:
            c_grp = st_df.groupby('main_category', as_index=False).agg(cs=('Actual Sales Amount', 'sum'), cu=('Sales Quantity', 'sum'))
            for _, cr in c_grp.sort_values(by='cs', ascending=False).iterrows():
                cmix = round(cr['cs'] / sales * 100, 1) if sales > 0 else 0
                cats_list.append({
                    "main_category": cr['main_category'],
                    "sales": f"{round(cr['cs']):,}",
                    "units": f"{int(cr['cu']):,}",
                    "store_mix_pct": f"{cmix}%",
                    "asp": f"{round(cr['cs']/cr['cu']) if cr['cu']>0 else 0:,}"
                })
        store_cat_details[code] = cats_list

    perf_df = pd.DataFrame(store_rows_data).sort_values(by='sales', ascending=False).reset_index(drop=True)

    # حساب الـ Totals الدقيقة لكل اختيار براند لشهر أكتوبر
    brand_totals = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub = perf_df if b == 'ALL' else perf_df[perf_df['brand'] == b]
        b_sales = sub['sales'].sum()
        b_target = sub['target'].sum()
        b_ach = (b_sales / b_target * 100) if b_target > 0 else 0
        b_units = sub['units'].sum()
        b_txns = sub['txns'].sum()
        b_atv = round(b_sales / b_txns) if b_txns > 0 else 0
        b_asp = round(b_sales / b_units) if b_units > 0 else 0
        b_soh = sub['soh_units'].sum()
        b_str = round(b_units / (b_units + b_soh) * 100, 1) if (b_units + b_soh) > 0 else 0
        b_ly = round(sub['ly_sales'].dropna().sum())
        b_yoy = ((sub[sub['ly_sales'].notna()]['sales'].sum() - b_ly) / b_ly * 100) if b_ly > 0 else 0

        brand_totals[b] = {
            "sales": f"{b_sales:,}", "ly": f"{b_ly:,}", "yoy": f"{b_yoy:+.1f}%", "yoy_val": b_yoy,
            "target": f"{round(b_target):,}", "ach": f"{b_ach:.1f}%", "ach_val": b_ach,
            "atv": f"{b_atv:,}", "asp": f"{b_asp:,}", "units": f"{b_units:,}", "txns": f"{b_txns:,}",
            "str": f"{b_str}%", "upt": f"{(b_units/b_txns):.2f}" if b_txns>0 else "0.00"
        }

    # بناء الأسطر في الجدول الرئيسي
    store_table_rows = ""
    for idx, r in perf_df.iterrows():
        yoy_str = f'<span style="color:{"#10b981" if r["yoy"]>=0 else "#ef4444"}; font-weight:700;">{r["yoy"]:+.1f}%</span>' if pd.notna(r["yoy"]) else '<span style="color:#64748b;">-</span>'
        ly_str = f"{r['ly_sales']:,}" if pd.notna(r['ly_sales']) else '<span style="color:#64748b;">-</span>'
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

    # هيكلية الـ Drill-down المتسلسلة لـ Business & Gender لكل براند
    hierarchy_tree = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_c = df_clean if b == 'ALL' else df_clean[df_clean['brand'] == b]
        tree = {}
        for main_c, m_grp in sub_c.groupby('main_category'):
            m_sales = round(m_grp['Actual Sales Amount'].sum())
            m_units = int(m_grp['Sales Quantity'].sum())
            tree[main_c] = {
                "sales": m_sales, "units": m_units,
                "asp": round(m_sales/m_units) if m_units>0 else 0,
                "subs": {}
            }
            for sub_c, s_grp in m_grp.groupby('sub_category'):
                s_sales = round(s_grp['Actual Sales Amount'].sum())
                s_units = int(s_grp['Sales Quantity'].sum())
                tree[main_c]["subs"][sub_c] = {
                    "sales": s_sales, "units": s_units,
                    "asp": round(s_sales/s_units) if s_units>0 else 0,
                    "subsubs": []
                }
                for subsub_c, ss_grp in s_grp.groupby('sub_sub_category'):
                    ss_sales = round(ss_grp['Actual Sales Amount'].sum())
                    ss_units = int(ss_grp['Sales Quantity'].sum())
                    tree[main_c]["subs"][sub_c]["subsubs"].append({
                        "name": subsub_c, "sales": ss_sales, "units": ss_units,
                        "asp": round(ss_sales/ss_units) if ss_units>0 else 0
                    })
        hierarchy_tree[b] = tree

    # توزيع الأحذية لـ DZL
    dzl_only_shoes = df_clean[(df_clean['brand'] == 'DZL') & (df_clean['main_category'] == 'Shoes')]
    dzl_g_sales = dzl_only_shoes.groupby('gender', as_index=False).agg(sales=('Actual Sales Amount', 'sum'), units=('Sales Quantity', 'sum'))
    dzl_g_tot = dzl_g_sales['sales'].sum()
    dzl_g_sales['share'] = (dzl_g_sales['sales'] / dzl_g_tot * 100).round(1) if dzl_g_tot > 0 else 0
    dzl_g_sales['asp'] = (dzl_g_sales['sales'] / dzl_g_sales['units']).fillna(0).round().astype(int)
    dzl_gender_data = {
        "labels": dzl_g_sales['gender'].tolist(),
        "series": dzl_g_sales['share'].tolist(),
        "records": dzl_g_sales.to_dict(orient='records')
    }

    # =========================================================================
    # محرك التوريد الآلي المتطور (Auto-Replenishment Engine with WH vs IST logic)
    # =========================================================================
    repl_data_list = []
    
    # 1. أوامر توريد ومناقلات DZL للأحذية
    dzl_shoes = dzl_only_shoes.copy()
    if not dzl_shoes.empty:
        sku_agg_dzl = dzl_shoes.groupby(['clean_code', 'clean_sku', 'clean_name'], as_index=False)['Sales Quantity'].sum()
        for idx, r in sku_agg_dzl.head(30).iterrows():
            st_c = r['clean_code']
            sku = r['clean_sku']
            name = r['clean_name']
            sold_qty = int(r['Sales Quantity'])
            st_info = STORE_MAPPING[st_c]

            wh_qty = int(wh_sku_soh.get(sku, 0))
            st_soh = int(store_sku_soh.get((st_c, sku), 0))
            sugg_qty = max(4, sold_qty * 2)

            if wh_qty > 0:
                action_type = "Warehouse Push (WH -> Store)"
                source_route = f"🏭 Central WH (KSWH) [Avail: {wh_qty} Pcs]"
                urgency = "⚡ WH Replenish (Immediate)"
                actual_qty = min(sugg_qty, wh_qty)
            else:
                action_type = "Store Transfer (IST - Opportunity)"
                # البحث عن متجر مانح يملك رصيداً
                donor_candidates = []
                for other_c in DZL_VALID_CODES:
                    if other_c != st_c:
                        d_stock = store_sku_soh.get((other_c, sku), 0)
                        donor_candidates.append((other_c, d_stock))
                
                donor_candidates.sort(key=lambda x: x[1], reverse=True)
                donor_code = donor_candidates[0][0] if donor_candidates else [c for c in DZL_VALID_CODES if c != st_c][0]
                donor_stock = donor_candidates[0][1] if donor_candidates else 0
                donor_info = STORE_MAPPING[donor_code]
                match_type = "🏙️ Same City" if donor_info['city'] == st_info['city'] else "🚛 Inter-City"

                source_route = f"{donor_info['full_name']} ({donor_code}) [{match_type}] [SOH: {donor_stock} Pcs]"
                urgency = "🚨 Broken Size Recovery (IST)"
                actual_qty = sugg_qty

            repl_data_list.append({
                "brand": "DZL",
                "action": action_type,
                "store": f"{st_info['full_name']} ({st_c})",
                "focus": f"👟 [SHOE] {name[:28]} (SKU: {sku})",
                "sold_qty": sold_qty,
                "store_soh": st_soh,
                "wh_soh": wh_qty,
                "qty": f"{actual_qty} Pcs",
                "source": source_route,
                "urgency": urgency
            })

    # 2. أوامر توريد ومناقلات MMS للأصناف سريعة الحركة
    mms_only_items = df_clean[df_clean['brand'] == 'MMS']
    if not mms_only_items.empty:
        sku_agg_mms = mms_only_items.groupby(['clean_code', 'clean_sku', 'clean_name', 'main_category'], as_index=False)['Sales Quantity'].sum()
        for idx, r in sku_agg_mms.sort_values(by='Sales Quantity', ascending=False).head(40).iterrows():
            st_c = r['clean_code']
            sku = r['clean_sku']
            name = r['clean_name']
            cat = r['main_category']
            sold_qty = int(r['Sales Quantity'])
            st_info = STORE_MAPPING[st_c]

            wh_qty = int(wh_sku_soh.get(sku, 0))
            st_soh = int(store_sku_soh.get((st_c, sku), 0))
            sugg_qty = max(6, sold_qty)

            if wh_qty > 0:
                action_type = "Warehouse Push (WH -> Store)"
                source_route = f"🏭 Central WH (KSWH) [Avail: {wh_qty} Pcs]"
                urgency = "⚡ WH Replenish (Fast Sell)"
                actual_qty = min(sugg_qty, wh_qty)
            else:
                action_type = "Store Transfer (IST - Opportunity)"
                # البحث عن متجر MMS مانح
                donor_candidates = []
                for other_c in MMS_VALID_CODES:
                    if other_c != st_c:
                        d_stock = store_sku_soh.get((other_c, sku), 0)
                        donor_candidates.append((other_c, d_stock))
                
                donor_candidates.sort(key=lambda x: (STORE_MAPPING[x[0]]['city'] == st_info['city'], x[1]), reverse=True)
                donor_code = donor_candidates[0][0] if donor_candidates else [c for c in MMS_VALID_CODES if c != st_c][0]
                donor_stock = donor_candidates[0][1] if donor_candidates else 0
                donor_info = STORE_MAPPING[donor_code]
                match_type = "🏙️ Same City" if donor_info['city'] == st_info['city'] else "🚛 Inter-City"

                source_route = f"{donor_info['full_name']} ({donor_code}) [{match_type}] [SOH: {donor_stock} Pcs]"
                urgency = "🚨 Out-of-Stock Risk (IST)"
                actual_qty = sugg_qty

            repl_data_list.append({
                "brand": "MMS",
                "action": action_type,
                "store": f"{st_info['full_name']} ({st_c})",
                "focus": f"📦 [{cat}] {name[:24]} (SKU: {sku})",
                "sold_qty": sold_qty,
                "store_soh": st_soh,
                "wh_soh": wh_qty,
                "qty": f"{actual_qty} Pcs",
                "source": source_route,
                "urgency": urgency
            })

    # قوائم MMS Top/Low 500
    mms_only = df_clean[df_clean['brand'] == 'MMS'].groupby(['clean_sku', 'clean_name', 'main_category'], as_index=False).agg(
        units=('Sales Quantity', 'sum'), sales=('Actual Sales Amount', 'sum')
    ).sort_values(by='units', ascending=False).reset_index(drop=True)
    mms_only['asp'] = (mms_only['sales'] / mms_only['units']).fillna(0).round().astype(int)
    mms_top500 = mms_only.head(500).to_dict(orient='records')
    mms_low500 = mms_only.tail(500).sort_values(by='units', ascending=True).to_dict(orient='records')

    # قوائم DZL Top/Low 20 Shoes لكل متجر
    dzl_store_movers = {}
    for st_c in ["ALL"] + list(DZL_VALID_CODES):
        st_sub = dzl_shoes if st_c == "ALL" else dzl_shoes[dzl_shoes['clean_code'] == st_c]
        if st_sub.empty:
            dzl_store_movers[st_c] = {"top": [], "low": []}
            continue
        st_grp = st_sub.groupby(['clean_sku', 'clean_name', 'gender'], as_index=False).agg(
            units=('Sales Quantity', 'sum'), sales=('Actual Sales Amount', 'sum')
        ).sort_values(by='units', ascending=False).reset_index(drop=True)
        st_grp['asp'] = (st_grp['sales'] / st_grp['units']).fillna(0).round().astype(int)
        dzl_store_movers[st_c] = {
            "top": st_grp.head(20).to_dict(orient='records'),
            "low": st_grp.tail(20).sort_values(by='units', ascending=True).to_dict(orient='records')
        }

    # استبدال NaN بـ None لضمان التوافق مع معيار JSON الصارم
    for item in store_rows_data:
        for k, v in item.items():
            if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                item[k] = None

    for item in sep_perf_list:
        for k, v in item.items():
            if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
                item[k] = None

    store_meta_map = {r['code']: r for r in store_rows_data}
    sep_meta_dict = {r['code']: r for r in sep_perf_list} if sep_perf_list else store_meta_map
    sep_totals_dict = sep_brand_totals if sep_brand_totals else brand_totals

    def safe_json(obj):
        return json.dumps(obj).replace("NaN", "null").replace("</", "<\\/")

    cur_all = brand_totals.get("ALL", {})
    init_sales = cur_all.get("sales", "0")
    init_ly = cur_all.get("ly", "0")
    init_yoy = cur_all.get("yoy", "0.0%")
    init_yoy_val = cur_all.get("yoy_val", 0)
    init_yoy_col = "#10b981" if init_yoy_val >= 0 else "#ef4444"
    init_target = cur_all.get("target", "0")
    init_ach = cur_all.get("ach", "0.0%")
    init_ach_val = cur_all.get("ach_val", 0)
    init_ach_col = "#10b981" if init_ach_val >= 100 else "#f59e0b"
    init_atv = cur_all.get("atv", "0")
    init_asp = cur_all.get("asp", "0")
    init_units = cur_all.get("units", "0")
    init_txns = cur_all.get("txns", "0")
    init_upt = cur_all.get("upt", "0.00")
    init_str = cur_all.get("str", "0.0%")

    template_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS & DZL Executive Commercial Intelligence Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>
    <style>
        :root {{
            --bg: #090d16; --card: #131b2e; --border: #1e293b; --primary: #38bdf8; --text-muted: #94a3b8;
        }}
        * {{ box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
        body {{ background: var(--bg); color: #fff; margin: 0; padding: 24px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); padding-bottom: 20px; margin-bottom: 24px; flex-wrap: wrap; gap: 16px; }}
        .top-controls {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
        .brand-switcher {{ display: flex; background: #0c1220; padding: 4px; border-radius: 8px; border: 1px solid var(--border); gap: 4px; }}
        .brand-btn {{ background: transparent; border: none; color: var(--text-muted); padding: 6px 14px; border-radius: 6px; font-size: 13px; font-weight: 700; cursor: pointer; }}
        .brand-btn.active {{ background: #2563eb; color: #fff; }}
        .month-select {{ background: #0c1220; border: 1px solid var(--border); color: #38bdf8; padding: 6px 14px; border-radius: 8px; font-weight: 700; outline: none; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 24px; }}
        .kpi-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 16px; }}
        .kpi-title {{ font-size: 11px; color: var(--text-muted); font-weight: 600; text-transform: uppercase; margin-bottom: 6px; }}
        .kpi-value {{ font-size: 22px; font-weight: 700; color: #fff; }}
        .view-toggle-bar {{ display: flex; background: #0c1220; padding: 4px; border-radius: 10px; border: 1px solid var(--border); margin-bottom: 24px; width: fit-content; gap: 4px; }}
        .view-btn {{ background: transparent; border: none; color: var(--text-muted); padding: 10px 22px; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; }}
        .view-btn.active {{ background: #2563eb; color: #fff; }}
        .table-wrap {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 24px; }}
        .table-header {{ padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); flex-wrap: wrap; gap: 10px; }}
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; }}
        th {{ background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; border-bottom: 1px solid var(--border); white-space: nowrap; }}
        td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); white-space: nowrap; }}
        tr:hover td {{ background: #19233c; }}
        .clickable-row {{ cursor: pointer; }}
        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
        .sub-tab-btn {{ background:#1e293b; color:#94a3b8; border:1px solid #334155; padding:8px 16px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; }}
        .sub-tab-btn.active {{ background:#38bdf8; color:#090d16; border-color:#38bdf8; }}
        .chart-container {{ background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 22px; margin-bottom: 24px; }}
        .app-modal {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: rgba(9, 13, 22, 0.9); z-index: 2147483647; display: none; align-items: center; justify-content: center; }}
        .modal-content {{ background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 92%; max-width: 950px; max-height: 90vh; display: flex; flex-direction: column; overflow: hidden; }}
        .modal-header {{ padding: 18px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; background: #0c1220; }}
        .modal-body {{ padding: 24px; overflow-y: auto; }}
        .close-btn {{ background: transparent; border: none; color: #94a3b8; font-size: 28px; cursor: pointer; }}
        .drill-crumb {{ display: inline-block; padding: 4px 10px; background: #1e293b; border-radius: 4px; margin-right: 6px; font-weight: 700; font-size: 12px; color: #38bdf8; cursor: pointer; }}
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
  const USER_ROLES = {{
    "MMS2026": {{ role: "ADMIN", name: "Executive & Merchandising (Full Access)", region: "ALL" }},
    "SULTAN2026": {{ role: "AREA_MGR", name: "Sultan", region: "Central & Eastern Region" }},
    "RAJIB2026": {{ role: "AREA_MGR", name: "Rajib", region: "Western, Southern & Northern Region" }}
  }};

  function checkAccess() {{
    var val = document.getElementById("access-pass").value.trim().toUpperCase();
    var user = USER_ROLES[val];
    if (user) {{
      sessionStorage.setItem("mms_user", JSON.stringify(user));
      document.getElementById("auth-overlay").style.display = "none";
      document.getElementById("current-user-badge").innerHTML = "👤 " + user.name;
    }} else {{
      document.getElementById("error-msg").style.display = "block";
    }}
  }}

  function logout() {{ sessionStorage.removeItem("mms_user"); location.reload(); }}

  document.addEventListener("DOMContentLoaded", function() {{
    var u = sessionStorage.getItem("mms_user");
    if (u) {{
      document.getElementById("auth-overlay").style.display = "none";
      document.getElementById("current-user-badge").innerHTML = "👤 " + JSON.parse(u).name;
    }}
    document.getElementById("access-pass").addEventListener("keypress", function(e) {{
      if (e.key === "Enter") checkAccess();
    }});
  }});
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
      <div style="background:#090d16; border:1px solid #1e293b; border-radius:10px; padding:16px; margin-bottom:18px;">
        <div style="font-size:13px; color:#f59e0b; margin-bottom:6px;"><strong>🎯 Focus Requirements:</strong> <span id="modal-needs" style="color:#fff;">-</span></div>
        <div style="font-size:13px; color:#38bdf8;"><strong>⚡ Operational Directive:</strong> <span id="modal-directive" style="color:#fff;">-</span></div>
      </div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:10px; margin-bottom:20px;">
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">CURRENT SALES</div><div id="modal-sales" style="font-size:16px; font-weight:700; color:#fff;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">TARGET (% ACH)</div><div id="modal-target" style="font-size:16px; font-weight:700; color:#10b981;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">SOH STOCK</div><div id="modal-soh" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">WOC COVER</div><div id="modal-woc" style="font-size:16px; font-weight:700; color:#f59e0b;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">ATV</div><div id="modal-atv" style="font-size:16px; font-weight:700; color:#fff;">-</div></div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;"><div style="font-size:10px; color:#94a3b8;">UPT</div><div id="modal-upt" style="font-size:16px; font-weight:700; color:#fff;">-</div></div>
      </div>
      <div style="font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8; margin-bottom:8px;">Category Contribution for this Door</div>
      <table><thead><tr><th>Category Name</th><th>Sales (SAR)</th><th>Units</th><th>Mix %</th><th>ASP</th></tr></thead><tbody id="modal-cats-body"></tbody></table>
    </div>
  </div>
</div>

<div class="header">
    <div>
        <h1 style="margin:0; font-size:22px;">MMS & DZL Executive Commercial Intelligence Dashboard</h1>
        <p style="margin:4px 0 0 0; color:var(--text-muted); font-size:13px;" id="headerSubtitle">October 2026 Daily Phasing & Commercial Performance Tracking</p>
    </div>
    <div class="top-controls">
        <select class="month-select" id="monthDropdown" onchange="switchMonth(this.value)">
            <option value="OCT" selected>📅 October 2026 (Active)</option>
            <option value="SEP">📅 September 2026 (Archived)</option>
        </select>
        <div class="brand-switcher">
            <button class="brand-btn active" id="btn-ALL" onclick="switchBrand('ALL')">🏢 ALL BRANDS</button>
            <button class="brand-btn" id="btn-MMS" onclick="switchBrand('MMS')">🔴 MUMUSO (17)</button>
            <button class="brand-btn" id="btn-DZL" onclick="switchBrand('DZL')">🟡 DZL (4)</button>
        </div>
        <span id="current-user-badge" style="font-size:13px; font-weight:700; color:#38bdf8; background:#1e293b; padding:8px 14px; border-radius:8px;">👤 Authenticating..</span>
        <button onclick="logout()" style="background:#ef444422; border:1px solid #ef444455; color:#ef4444; padding:8px 14px; border-radius:8px; font-weight:700; cursor:pointer;">Logout</button>
    </div>
</div>

<div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-title">Current Total Sales</div><div class="kpi-value" id="kpi-sales">{init_sales} <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">LY Gross Sales</div><div class="kpi-value" id="kpi-ly" style="color:#38bdf8;">{init_ly} <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">Network LFL YoY Growth</div><div class="kpi-value" id="kpi-yoy" style="color:{init_yoy_col}; font-weight:800;">{init_yoy}</div></div>
    <div class="kpi-card"><div class="kpi-title">Total Target</div><div class="kpi-value" id="kpi-target">{init_target} <span style="font-size:12px; color:var(--text-muted);">SAR</span></div></div>
    <div class="kpi-card"><div class="kpi-title">Achievement (% Ach)</div><div class="kpi-value" id="kpi-ach" style="color:{init_ach_col};">{init_ach}</div></div>
    <div class="kpi-card"><div class="kpi-title">Network ATV</div><div class="kpi-value" id="kpi-atv">SAR {init_atv}</div></div>
    <div class="kpi-card"><div class="kpi-title">Network ASP</div><div class="kpi-value" id="kpi-asp">SAR {init_asp}</div></div>
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
                <tbody id="storesTableBody">
                    {store_table_rows}
                </tbody>
                <tfoot>
                    <tr id="totalPortfolioRow" style="background:#0c1220; font-weight:800; border-top:3px solid #38bdf8; font-size:13px;">
                        <td colspan="4" style="color:#38bdf8;" id="totalRowTitle">TOTAL PORTFOLIO (ALL DOORS)</td>
                        <td style="color:#fff;" id="tot-sales">{init_sales}</td>
                        <td style="color:#38bdf8;" id="tot-ly">{init_ly}</td>
                        <td id="tot-yoy" style="color:{init_yoy_col};">{init_yoy}</td>
                        <td style="color:#94a3b8;" id="tot-target">{init_target}</td>
                        <td id="tot-ach" style="color:{init_ach_col};">{init_ach}</td>
                        <td style="color:#38bdf8;" id="tot-units">{init_units}</td>
                        <td style="color:#fff;" id="tot-txns">{init_txns}</td>
                        <td style="color:#10b981;" id="tot-upt">{init_upt}</td>
                        <td style="color:#38bdf8;" id="tot-str">{init_str}</td>
                        <td>-</td>
                        <td style="color:#f59e0b;" id="tot-asp">{init_asp}</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>
</div>

<!-- 2. Region-Wise View -->
<div id="view-regions" style="display:none;">
    <div id="centralRegionOverview" class="kpi-grid" style="margin-bottom:14px;"></div>
    <div class="table-wrap" style="margin-bottom:30px;">
        <div class="table-header"><h3 style="margin:0; font-size:15px; color:#38bdf8;">🏢 CENTRAL & EASTERN REGION (Sultan - 11 Doors)</h3></div>
        <div style="overflow-x:auto;">
            <table id="regionCentralTable">
                <thead><tr><th>#</th><th>Code</th><th>Store Name</th><th>Sales</th><th>LY Sales</th><th>YoY</th><th>Target</th><th>% Ach</th><th>Units</th><th>Trans</th><th>UPT</th><th>ASP</th></tr></thead>
                <tbody></tbody>
                <tfoot>
                    <tr id="centralTotalRow" style="background:#0c1220; font-weight:800; border-top:2px solid #38bdf8;">
                        <td colspan="3" style="color:#38bdf8;">TOTAL CENTRAL REGION</td>
                        <td style="color:#fff;" id="c-tot-sales">-</td>
                        <td style="color:#38bdf8;" id="c-tot-ly">-</td>
                        <td id="c-tot-yoy">-</td>
                        <td style="color:#94a3b8;" id="c-tot-target">-</td>
                        <td id="c-tot-ach">-</td>
                        <td style="color:#38bdf8;" id="c-tot-units">-</td>
                        <td style="color:#fff;" id="c-tot-txns">-</td>
                        <td style="color:#10b981;" id="c-tot-upt">-</td>
                        <td style="color:#f59e0b;" id="c-tot-asp">-</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>

    <div id="westernRegionOverview" class="kpi-grid" style="margin-bottom:14px;"></div>
    <div class="table-wrap">
        <div class="table-header"><h3 style="margin:0; font-size:15px; color:#818cf8;">🏢 WESTERN, SOUTHERN & NORTHERN REGION (Rajib - 12 Doors)</h3></div>
        <div style="overflow-x:auto;">
            <table id="regionWesternTable">
                <thead><tr><th>#</th><th>Code</th><th>Store Name</th><th>Sales</th><th>LY Sales</th><th>YoY</th><th>Target</th><th>% Ach</th><th>Units</th><th>Trans</th><th>UPT</th><th>ASP</th></tr></thead>
                <tbody></tbody>
                <tfoot>
                    <tr id="westernTotalRow" style="background:#0c1220; font-weight:800; border-top:2px solid #818cf8;">
                        <td colspan="3" style="color:#818cf8;">TOTAL WESTERN REGION</td>
                        <td style="color:#fff;" id="w-tot-sales">-</td>
                        <td style="color:#38bdf8;" id="w-tot-ly">-</td>
                        <td id="w-tot-yoy">-</td>
                        <td style="color:#94a3b8;" id="w-tot-target">-</td>
                        <td id="w-tot-ach">-</td>
                        <td style="color:#38bdf8;" id="w-tot-units">-</td>
                        <td style="color:#fff;" id="w-tot-txns">-</td>
                        <td style="color:#10b981;" id="w-tot-upt">-</td>
                        <td style="color:#f59e0b;" id="w-tot-asp">-</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>
</div>

<!-- 3. Business & Gender Mix View -->
<div id="view-business" style="display:none;">
    <div style="margin-bottom:14px; display:flex; justify-content:space-between; align-items:center; background:#131b2e; padding:12px 18px; border-radius:8px; border:1px solid #1e293b; flex-wrap:wrap; gap:10px;">
        <span style="font-weight:700; color:#38bdf8; font-size:14px;">🔍 Select Brand for Assortment Analysis:</span>
        <select id="businessBrandSelect" class="month-select" onchange="switchBusinessBrand(this.value)">
            <option value="MMS" selected>🔴 MUMUSO Categories Only (No Shoes)</option>
            <option value="DZL">🟡 DZL (Shoes & Accessories Only)</option>
            <option value="ALL">🏢 ALL BRANDS COMBINED</option>
        </select>
    </div>

    <div style="display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px;">
        <div class="chart-container" style="flex:1; min-width:320px;">
            <div style="font-weight:700; margin-bottom:12px; font-size:14px;" id="catDonutTitle">🍩 Category Contribution Share (MMS)</div>
            <div id="apexCategoryDonut" style="min-height: 330px;"></div>
        </div>
        <div class="chart-container" id="genderChartWrapper" style="flex:1; min-width:320px; display:none;">
            <div style="font-weight:700; margin-bottom:12px; font-size:14px;">👟 Footwear Gender Mix Share (Kids 23-34 | Women 35-39 | Men 40-48)</div>
            <div id="apexGenderDonut" style="min-height: 330px;"></div>
        </div>
    </div>

    <!-- 3-Level Category Drill-down Matrix -->
    <div class="table-wrap">
        <div class="table-header">
            <div>
                <h3 id="drillTitle" style="margin:0; font-size:15px; color:#38bdf8;">📦 CATEGORY HIERARCHY DRILL-DOWN</h3>
                <div id="drillBreadcrumbs" style="margin-top:6px;"></div>
            </div>
        </div>
        <div style="overflow-x:auto;">
            <table id="drillTable">
                <thead id="drillTableHead"></thead>
                <tbody id="drillTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<!-- 4. Commercial Action Hub (مع أعمدة المخزون والمبيعات الكاملة) -->
<div id="view-action" style="display:none;">
    <div class="table-wrap" style="margin-bottom:24px;">
        <div class="table-header">
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <h3 style="margin:0; font-size:15px; color:#10b981;">⚡ PREDICTIVE AUTO-REPLENISHMENT & IST ROUTING</h3>
                <button class="sub-tab-btn active" id="btn-ist-all" onclick="filterISTBrand('ALL')">ALL REPLENISHMENT</button>
                <button class="sub-tab-btn" id="btn-ist-dzl" onclick="filterISTBrand('DZL')">DZL (Shoes)</button>
                <button class="sub-tab-btn" id="btn-ist-mms" onclick="filterISTBrand('MMS')">MMS (Fast Movers)</button>
            </div>
            <div style="display:flex; gap:8px;">
                <button onclick="downloadISTPlan('DZL')" style="background:#ef4444; color:#fff; border:none; padding:8px 14px; border-radius:6px; font-weight:700; font-size:12px; cursor:pointer;">📥 Export DZL Plan</button>
                <button onclick="downloadISTPlan('MMS')" style="background:#2563eb; color:#fff; border:none; padding:8px 14px; border-radius:6px; font-weight:700; font-size:12px; cursor:pointer;">📥 Export MMS Plan</button>
            </div>
        </div>
        <div style="max-height:420px; overflow-y:auto;">
            <table>
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Brand</th>
                        <th>Action Type</th>
                        <th>Store Target</th>
                        <th>SKU / Style Focus</th>
                        <th>Sold (MTD)</th>
                        <th>Store SOH</th>
                        <th>Central WH SOH</th>
                        <th>Sugg Qty</th>
                        <th>Source Route</th>
                        <th>Urgency</th>
                    </tr>
                </thead>
                <tbody id="replTableBody"></tbody>
            </table>
        </div>
    </div>

    <!-- DZL Top/Low 20 Shoes per Store -->
    <div id="dzlMoversBlock" class="table-wrap" style="margin-bottom:24px;">
        <div class="table-header">
            <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
                <button class="sub-tab-btn active" id="btn-dzl-top" onclick="switchDZLMovers('top')">🔥 DZL TOP 20 SHOES</button>
                <button class="sub-tab-btn" id="btn-dzl-low" onclick="switchDZLMovers('low')">❄️ DZL LOW 20 SHOES</button>
                <select id="dzlStoreSelect" onchange="renderDZLMovers()" class="month-select">
                    <option value="ALL">All DZL Stores</option>
                    <option value="K107">DZL Riyadh Park (K107)</option>
                    <option value="K104">DZL Riyadh U-Walk (K104)</option>
                    <option value="K111">DZL Solitaire (K111)</option>
                    <option value="K204">DZL Redsea (K204)</option>
                </select>
            </div>
        </div>
        <div style="overflow-x:auto;">
            <table>
                <thead><tr><th>#</th><th>SKU</th><th>Product Name</th><th>Gender</th><th>Units</th><th>Sales (SAR)</th><th>ASP</th></tr></thead>
                <tbody id="dzlMoversBody"></tbody>
            </table>
        </div>
    </div>

    <!-- MMS Top/Low 500 Fast Movers -->
    <div id="mmsMoversBlock" class="table-wrap">
        <div class="table-header">
            <div style="display:flex; gap:10px; align-items:center;">
                <button class="sub-tab-btn active" id="btn-mms-top" onclick="switchMMSMovers('top')">🔥 MMS TOP 500 FAST MOVERS</button>
                <button class="sub-tab-btn" id="btn-mms-low" onclick="switchMMSMovers('low')">❄️ MMS LOW 500 CLEARANCE</button>
            </div>
            <input type="text" id="mmsSearch" placeholder="Search SKU..." onkeyup="renderMMSMovers()" style="background:#090d16; border:1px solid var(--border); color:#fff; padding:6px 12px; border-radius:6px;">
        </div>
        <div style="max-height:400px; overflow-y:auto;">
            <table>
                <thead><tr><th>Rank</th><th>SKU Code</th><th>Product Name</th><th>Category</th><th>Units Sold</th><th>Sales (SAR)</th><th>ASP</th></tr></thead>
                <tbody id="mmsMoversTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script>
  let activeMonth = 'OCT';
  let activeBrand = 'ALL';
  let businessBrand = 'MMS';
  let istFilterBrand = 'ALL';

  const OCT_BRAND_TOTALS = {safe_json(brand_totals)};
  const SEP_BRAND_TOTALS = {safe_json(sep_totals_dict)};
  const OCT_STORE_META = {safe_json(store_meta_map)};
  const SEP_STORE_META = {safe_json(sep_meta_dict)};
  
  let currentStoreMeta = OCT_STORE_META;
  let currentBrandTotals = OCT_BRAND_TOTALS;

  const STORE_CATS = {safe_json(store_cat_details)};
  const MMS_TOP500 = {safe_json(mms_top500)};
  const MMS_LOW500 = {safe_json(mms_low500)};
  const DZL_MOVERS = {safe_json(dzl_store_movers)};
  const HIERARCHY_TREE = {safe_json(hierarchy_tree)};
  const DZL_GENDER = {safe_json(dzl_gender_data)};
  const REPL_ITEMS = {safe_json(repl_data_list)};

  let currentDrillLevel = 1;
  let selectedMainCat = null;
  let selectedSubCat = null;
  let dzlMoversType = 'top';
  let mmsMoversType = 'top';
  let donutChart = null;
  let genderChart = null;

  document.addEventListener("DOMContentLoaded", function() {{
    try {{ renderRegionTables(); }} catch(e) {{ console.error(e); }}
    try {{ renderDrillDown(); }} catch(e) {{ console.error(e); }}
    try {{ renderDZLMovers(); }} catch(e) {{ console.error(e); }}
    try {{ renderMMSMovers(); }} catch(e) {{ console.error(e); }}
    try {{ renderISTTable(); }} catch(e) {{ console.error(e); }}
  }});

  function updateKPICards(b) {{
    try {{
      const d = currentBrandTotals[b] || currentBrandTotals['ALL'];
      if (!d) return;
      document.getElementById("kpi-sales").innerHTML = d.sales + " <span style='font-size:12px; color:var(--text-muted);'>SAR</span>";
      document.getElementById("kpi-ly").innerHTML = d.ly + " <span style='font-size:12px; color:var(--text-muted);'>SAR</span>";
      const yoyEl = document.getElementById("kpi-yoy");
      yoyEl.innerText = d.yoy;
      yoyEl.style.color = (d.yoy_val >= 0) ? "#10b981" : "#ef4444";

      document.getElementById("kpi-target").innerHTML = d.target + " <span style='font-size:12px; color:var(--text-muted);'>SAR</span>";
      const achEl = document.getElementById("kpi-ach");
      achEl.innerText = d.ach;
      achEl.style.color = (d.ach_val >= 100) ? "#10b981" : "#f59e0b";

      document.getElementById("kpi-atv").innerText = "SAR " + d.atv;
      document.getElementById("kpi-asp").innerText = "SAR " + d.asp;

      document.getElementById("totalRowTitle").innerText = (b === 'ALL') ? "TOTAL PORTFOLIO (ALL DOORS)" : (b === 'MMS' ? "TOTAL MUMUSO NETWORK (17 DOORS)" : "TOTAL DZL DOZOLO (4 DOORS)");
      document.getElementById("tot-sales").innerText = d.sales;
      document.getElementById("tot-ly").innerText = d.ly;
      const totYoy = document.getElementById("tot-yoy");
      totYoy.innerText = d.yoy;
      totYoy.style.color = (d.yoy_val >= 0) ? "#10b981" : "#ef4444";

      document.getElementById("tot-target").innerText = d.target;
      const totAch = document.getElementById("tot-ach");
      totAch.innerText = d.ach;
      totAch.style.color = (d.ach_val >= 100) ? "#10b981" : "#f59e0b";

      document.getElementById("tot-units").innerText = d.units;
      document.getElementById("tot-txns").innerText = d.txns;
      document.getElementById("tot-upt").innerText = d.upt;
      document.getElementById("tot-str").innerText = d.str || "0.0%";
      document.getElementById("tot-asp").innerText = d.asp;
    }} catch(e) {{ console.error("updateKPICards error:", e); }}
  }}

  function switchBrand(b) {{
    try {{
      activeBrand = b;
      document.querySelectorAll('.brand-btn').forEach(btn => btn.classList.remove('active'));
      var el = document.getElementById('btn-' + b);
      if (el) el.classList.add('active');

      document.querySelectorAll('.store-row').forEach(row => {{
        const rBrand = row.getAttribute('data-brand');
        row.style.display = (b === 'ALL' || rBrand === b) ? '' : 'none';
      }});

      updateKPICards(b);
      renderRegionTables();
      
      businessBrand = (b === 'DZL') ? 'DZL' : 'MMS';
      var sel = document.getElementById("businessBrandSelect");
      if (sel) sel.value = businessBrand;
      renderDrillDown();
      renderCharts();

      const dzlBlock = document.getElementById("dzlMoversBlock");
      const mmsBlock = document.getElementById("mmsMoversBlock");
      if (b === 'DZL') {{
        if (dzlBlock) dzlBlock.style.display = "block";
        if (mmsBlock) mmsBlock.style.display = "none";
      }} else if (b === 'MMS') {{
        if (dzlBlock) dzlBlock.style.display = "none";
        if (mmsBlock) mmsBlock.style.display = "block";
      }} else {{
        if (dzlBlock) dzlBlock.style.display = "block";
        if (mmsBlock) mmsBlock.style.display = "block";
      }}
    }} catch(e) {{ console.error(e); }}
  }}

  function switchBusinessBrand(val) {{
    businessBrand = val;
    currentDrillLevel = 1;
    selectedMainCat = null;
    selectedSubCat = null;
    renderDrillDown();
    renderCharts();
  }}

  function filterStores() {{
    const q = document.getElementById('storeSearch').value.toLowerCase();
    document.querySelectorAll('.store-row').forEach(row => {{
      const text = row.innerText.toLowerCase();
      const rBrand = row.getAttribute('data-brand');
      const matchBrand = (activeBrand === 'ALL' || rBrand === activeBrand);
      row.style.display = (matchBrand && text.includes(q)) ? '' : 'none';
    }});
  }}

  function renderRegionTables() {{
    try {{
      const centralTbody = document.querySelector("#regionCentralTable tbody");
      const westernTbody = document.querySelector("#regionWesternTable tbody");
      let cRows = "", wRows = "";
      let cIdx = 1, wIdx = 1;
      let cSales = 0, cTarget = 0, cUnits = 0, cTxns = 0, cLy = 0;
      let wSales = 0, wTarget = 0, wUnits = 0, wTxns = 0, wLy = 0;

      Object.values(currentStoreMeta).forEach(r => {{
        if (activeBrand !== 'ALL' && r.brand !== activeBrand) return;
        const isCentral = r.region.includes('Central');
        const yoyStr = (r.yoy !== null && !isNaN(r.yoy)) ? `<span style="color:${{r.yoy>=0?'#10b981':'#ef4444'}}">${{r.yoy.toFixed(1)}}%</span>` : '-';
        const rowHtml = `<tr class="clickable-row" onclick="openStoreModal('${{r.code}}')">
          <td style="color:#64748b;">${{isCentral ? cIdx++ : wIdx++}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.code}}</td>
          <td style="color:#fff;">${{r.name}}</td>
          <td style="color:#fff; font-weight:700;">${{r.sales.toLocaleString()}}</td>
          <td style="color:#38bdf8;">${{r.ly_sales ? r.ly_sales.toLocaleString() : '-'}}</td>
          <td>${{yoyStr}}</td>
          <td style="color:#94a3b8;">${{Math.round(r.target).toLocaleString()}}</td>
          <td style="color:${{r.ach >= 100 ? '#10b981' : '#f59e0b'}}; font-weight:700;">${{r.ach.toFixed(1)}}%</td>
          <td style="color:#38bdf8;">${{r.units.toLocaleString()}}</td>
          <td>${{r.txns.toLocaleString()}}</td>
          <td style="color:#10b981;">${{r.upt.toFixed(2)}}</td>
          <td style="color:#f59e0b;">${{r.asp}}</td>
        </tr>`;

        if (isCentral) {{
          cRows += rowHtml; cSales += r.sales; cTarget += r.target; cUnits += r.units; cTxns += r.txns;
          if (r.ly_sales) cLy += r.ly_sales;
        }} else {{
          wRows += rowHtml; wSales += r.sales; wTarget += r.target; wUnits += r.units; wTxns += r.txns;
          if (r.ly_sales) wLy += r.ly_sales;
        }}
      }});

      if (centralTbody) centralTbody.innerHTML = cRows;
      if (westernTbody) westernTbody.innerHTML = wRows;

      const cAch = (cTarget > 0) ? (cSales / cTarget * 100).toFixed(1) : 0;
      const wAch = (wTarget > 0) ? (wSales / wTarget * 100).toFixed(1) : 0;
      const cYoy = (cLy > 0) ? ((cSales - cLy) / cLy * 100).toFixed(1) : 0;
      const wYoy = (wLy > 0) ? ((wSales - wLy) / wLy * 100).toFixed(1) : 0;

      document.getElementById("centralRegionOverview").innerHTML = `
        <div class="kpi-card"><div class="kpi-title">CENTRAL SALES</div><div class="kpi-value">${{cSales.toLocaleString()}} <span style="font-size:11px;">SAR</span></div></div>
        <div class="kpi-card"><div class="kpi-title">CENTRAL TARGET</div><div class="kpi-value">${{Math.round(cTarget).toLocaleString()}} <span style="font-size:11px;">SAR</span></div></div>
        <div class="kpi-card"><div class="kpi-title">ACHIEVEMENT</div><div class="kpi-value" style="color:${{cAch>=100?'#10b981':'#f59e0b'}};">${{cAch}}%</div></div>
        <div class="kpi-card"><div class="kpi-title">QUANTITY</div><div class="kpi-value" style="color:#38bdf8;">${{cUnits.toLocaleString()}}</div></div>
        <div class="kpi-card"><div class="kpi-title">ACTIVE DOORS</div><div class="kpi-value">${{cIdx-1}}</div></div>
      `;

      document.getElementById("westernRegionOverview").innerHTML = `
        <div class="kpi-card"><div class="kpi-title">WESTERN SALES</div><div class="kpi-value">${{wSales.toLocaleString()}} <span style="font-size:11px;">SAR</span></div></div>
        <div class="kpi-card"><div class="kpi-title">WESTERN TARGET</div><div class="kpi-value">${{Math.round(wTarget).toLocaleString()}} <span style="font-size:11px;">SAR</span></div></div>
        <div class="kpi-card"><div class="kpi-title">ACHIEVEMENT</div><div class="kpi-value" style="color:${{wAch>=100?'#10b981':'#f59e0b'}};">${{wAch}}%</div></div>
        <div class="kpi-card"><div class="kpi-title">QUANTITY</div><div class="kpi-value" style="color:#38bdf8;">${{wUnits.toLocaleString()}}</div></div>
        <div class="kpi-card"><div class="kpi-title">ACTIVE DOORS</div><div class="kpi-value">${{wIdx-1}}</div></div>
      `;

      document.getElementById("c-tot-sales").innerText = cSales.toLocaleString();
      document.getElementById("c-tot-ly").innerText = cLy.toLocaleString();
      document.getElementById("c-tot-yoy").innerHTML = `<span style="color:${{cYoy>=0?'#10b981':'#ef4444'}}">${{cYoy>0?'+':''}}${{cYoy}}%</span>`;
      document.getElementById("c-tot-target").innerText = Math.round(cTarget).toLocaleString();
      document.getElementById("c-tot-ach").innerHTML = `<span style="color:${{cAch>=100?'#10b981':'#f59e0b'}}">${{cAch}}%</span>`;
      document.getElementById("c-tot-units").innerText = cUnits.toLocaleString();
      document.getElementById("c-tot-txns").innerText = cTxns.toLocaleString();
      document.getElementById("c-tot-upt").innerText = (cTxns>0?(cUnits/cTxns).toFixed(2):"0.00");
      document.getElementById("c-tot-asp").innerText = (cUnits>0?Math.round(cSales/cUnits):0);

      document.getElementById("w-tot-sales").innerText = wSales.toLocaleString();
      document.getElementById("w-tot-ly").innerText = wLy.toLocaleString();
      document.getElementById("w-tot-yoy").innerHTML = `<span style="color:${{wYoy>=0?'#10b981':'#ef4444'}}">${{wYoy>0?'+':''}}${{wYoy}}%</span>`;
      document.getElementById("w-tot-target").innerText = Math.round(wTarget).toLocaleString();
      document.getElementById("w-tot-ach").innerHTML = `<span style="color:${{wAch>=100?'#10b981':'#f59e0b'}}">${{wAch}}%</span>`;
      document.getElementById("w-tot-units").innerText = wUnits.toLocaleString();
      document.getElementById("w-tot-txns").innerText = wTxns.toLocaleString();
      document.getElementById("w-tot-upt").innerText = (wTxns>0?(wUnits/wTxns).toFixed(2):"0.00");
      document.getElementById("w-tot-asp").innerText = (wUnits>0?Math.round(wSales/wUnits):0);
    }} catch(e) {{ console.error("renderRegionTables error:", e); }}
  }}

  function renderDrillDown() {{
    try {{
      const tree = HIERARCHY_TREE[businessBrand] || {{}};
      const thead = document.getElementById("drillTableHead");
      const tbody = document.getElementById("drillTableBody");
      const crumbs = document.getElementById("drillBreadcrumbs");
      let cHtml = `<span class="drill-crumb" onclick="drillGoLevel(1)">🏷️ All Categories (${{businessBrand}})</span>`;

      if (currentDrillLevel === 1) {{
        crumbs.innerHTML = cHtml;
        thead.innerHTML = `<tr><th>#</th><th>Main Category</th><th>Sales Revenue (SAR)</th><th>Units Sold</th><th>ASP (SAR)</th><th>Action</th></tr>`;
        let bHtml = "";
        let idx = 1;
        for (const [mCat, data] of Object.entries(tree)) {{
          bHtml += `<tr>
            <td style="color:#64748b;">${{idx++}}</td>
            <td style="color:#fff; font-weight:700;">🏷️ ${{mCat}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{data.sales.toLocaleString()}}</td>
            <td>${{data.units.toLocaleString()}}</td>
            <td style="color:#f59e0b;">${{data.asp}}</td>
            <td><button class="sub-tab-btn" onclick="drillIntoMainCat('${{mCat.replace("'", "\\'")}}')">View Sub-Categories ▼</button></td>
          </tr>`;
        }}
        tbody.innerHTML = bHtml || `<tr><td colspan="6" style="text-align:center;">No data available</td></tr>`;
      }} else if (currentDrillLevel === 2) {{
        cHtml += ` <span style="color:#64748b;">></span> <span class="drill-crumb" onclick="drillGoLevel(2)">📁 ${{selectedMainCat}}</span>`;
        crumbs.innerHTML = cHtml;
        thead.innerHTML = `<tr><th>#</th><th>Sub-Category</th><th>Sales Revenue (SAR)</th><th>Units Sold</th><th>ASP (SAR)</th><th>Action</th></tr>`;
        const subs = tree[selectedMainCat]?.subs || {{}};
        let bHtml = "";
        let idx = 1;
        for (const [sCat, data] of Object.entries(subs)) {{
          bHtml += `<tr>
            <td style="color:#64748b;">${{idx++}}</td>
            <td style="color:#fff; font-weight:700;">📁 ${{sCat}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{data.sales.toLocaleString()}}</td>
            <td>${{data.units.toLocaleString()}}</td>
            <td style="color:#f59e0b;">${{data.asp}}</td>
            <td><button class="sub-tab-btn" onclick="drillIntoSubCat('${{sCat.replace("'", "\\'")}}')">View Items ▼</button></td>
          </tr>`;
        }}
        tbody.innerHTML = bHtml || `<tr><td colspan="6" style="text-align:center;">No sub-categories</td></tr>`;
      }} else if (currentDrillLevel === 3) {{
        cHtml += ` <span style="color:#64748b;">></span> <span class="drill-crumb" onclick="drillGoLevel(2)">📁 ${{selectedMainCat}}</span> <span style="color:#64748b;">></span> <span class="drill-crumb">📦 ${{selectedSubCat}}</span>`;
        crumbs.innerHTML = cHtml;
        thead.innerHTML = `<tr><th>#</th><th>Sub-Sub / Item Class</th><th>Sales Revenue (SAR)</th><th>Units Sold</th><th>ASP (SAR)</th></tr>`;
        const subsubs = tree[selectedMainCat]?.subs[selectedSubCat]?.subsubs || [];
        let bHtml = "";
        subsubs.forEach((item, idx) => {{
          bHtml += `<tr>
            <td style="color:#64748b;">${{idx+1}}</td>
            <td style="color:#fff; font-weight:600;">📦 ${{item.name}}</td>
            <td style="color:#38bdf8; font-weight:700;">${{item.sales.toLocaleString()}}</td>
            <td>${{item.units.toLocaleString()}}</td>
            <td style="color:#f59e0b;">${{item.asp}}</td>
          </tr>`;
        }});
        tbody.innerHTML = bHtml || `<tr><td colspan="5" style="text-align:center;">No items found</td></tr>`;
      }}
    }} catch(e) {{ console.error("renderDrillDown error:", e); }}
  }}

  function drillGoLevel(lvl) {{
    currentDrillLevel = lvl;
    if (lvl === 1) {{ selectedMainCat = null; selectedSubCat = null; }}
    if (lvl === 2) {{ selectedSubCat = null; }}
    renderDrillDown();
  }}

  function drillIntoMainCat(m) {{ selectedMainCat = m; currentDrillLevel = 2; renderDrillDown(); }}
  function drillIntoSubCat(s) {{ selectedSubCat = s; currentDrillLevel = 3; renderDrillDown(); }}

  function renderCharts() {{
    try {{
      const tree = HIERARCHY_TREE[businessBrand] || {{}};
      const catLabels = Object.keys(tree);
      const catSeries = catLabels.map(k => tree[k].sales);
      const colors = ['#38bdf8', '#f59e0b', '#10b981', '#ec4899', '#818cf8', '#a855f7', '#06b6d4', '#e11d48', '#6366f1', '#14b8a6'];

      var titleEl = document.getElementById("catDonutTitle");
      if (titleEl) titleEl.innerText = `🍩 Category Contribution Share (${{businessBrand}})`;
      
      const donutEl = document.querySelector("#apexCategoryDonut");
      if (donutEl) {{
        donutEl.innerHTML = "";
        if (donutChart) {{ try {{ donutChart.destroy(); }} catch(e){{}} }}
        if (catSeries.length > 0) {{
          donutChart = new ApexCharts(donutEl, {{
            series: catSeries, labels: catLabels,
            chart: {{ type: 'donut', height: 330, background: 'transparent' }},
            theme: {{ mode: 'dark' }}, colors: colors, legend: {{ position: 'bottom', labels: {{ colors: '#cbd5e1' }} }}
          }});
          donutChart.render();
        }}
      }}

      const genderWrapper = document.getElementById("genderChartWrapper");
      if (businessBrand === 'DZL') {{
        if (genderWrapper) genderWrapper.style.display = "block";
        const genderEl = document.querySelector("#apexGenderDonut");
        if (genderEl) {{
          genderEl.innerHTML = "";
          if (genderChart) {{ try {{ genderChart.destroy(); }} catch(e){{}} }}
          if (DZL_GENDER.series.length > 0) {{
            genderChart = new ApexCharts(genderEl, {{
              series: DZL_GENDER.series, labels: DZL_GENDER.labels,
              chart: {{ type: 'donut', height: 330, background: 'transparent' }},
              theme: {{ mode: 'dark' }}, colors: ['#ec4899', '#38bdf8', '#10b981'], legend: {{ position: 'bottom', labels: {{ colors: '#cbd5e1' }} }}
            }});
            genderChart.render();
          }}
        }}
      }} else {{
        if (genderWrapper) genderWrapper.style.display = "none";
      }}
    }} catch(e) {{ console.error("renderCharts error:", e); }}
  }}

  function switchView(viewName) {{
    try {{
      document.getElementById("view-stores").style.display = (viewName === 'stores') ? 'block' : 'none';
      document.getElementById("view-regions").style.display = (viewName === 'regions') ? 'block' : 'none';
      document.getElementById("view-business").style.display = (viewName === 'business') ? 'block' : 'none';
      document.getElementById("view-action").style.display = (viewName === 'action') ? 'block' : 'none';
      document.querySelectorAll('.view-btn').forEach(btn => btn.classList.remove('active'));
      var b = document.getElementById('btn-' + viewName);
      if (b) b.classList.add('active');

      if (viewName === 'regions') renderRegionTables();
      if (viewName === 'business') {{
        setTimeout(() => {{
          renderCharts();
          renderDrillDown();
          window.dispatchEvent(new Event('resize'));
        }}, 60);
      }}
      if (viewName === 'action') renderISTTable();
    }} catch(e) {{ console.error("switchView error:", e); }}
  }}

  function openStoreModal(code) {{
    try {{
      const r = currentStoreMeta[code];
      if (!r) return;
      document.getElementById("modal-store-name").innerText = "[" + r.brand + "] " + r.name;
      document.getElementById("modal-store-code").innerText = "CODE: " + code + " | " + r.region + " (Manager: " + r.manager + ")";
      document.getElementById("modal-sales").innerText = r.sales.toLocaleString() + " SAR";
      document.getElementById("modal-target").innerText = Math.round(r.target).toLocaleString() + " SAR (" + r.ach.toFixed(1) + "%)";
      document.getElementById("modal-soh").innerText = (r.soh_units || 0).toLocaleString() + " Pcs";
      document.getElementById("modal-woc").innerText = (r.woc || 0) + " Wks";
      document.getElementById("modal-atv").innerText = r.atv.toLocaleString() + " SAR";
      document.getElementById("modal-upt").innerText = r.upt.toFixed(2);

      let needs = "Maintain standard assortment and monitor broken sizes.";
      let directive = "Weekly routine replenishment.";
      if (r.ach >= 95) {{
        needs = "High demand velocity. Priority supply for fast movers.";
        directive = "⚡ Maintain 100% floor availability on leading drivers.";
      }} else if (r.ach < 70 && (r.soh_units || 0) > 10000) {{
        needs = "Store holds heavy display depth but slow sell-through.";
        directive = "⚡ Reallocate front gondolas to high-velocity impulse items and initiate clearance.";
      }}
      document.getElementById("modal-needs").innerText = needs;
      document.getElementById("modal-directive").innerText = directive;

      const cats = STORE_CATS[code] || [];
      let html = "";
      cats.forEach(c => {{
        html += `<tr>
          <td style="color:#38bdf8; font-weight:700;">${{c.main_category}}</td>
          <td>${{c.sales}}</td>
          <td>${{c.units}}</td>
          <td style="color:#10b981; font-weight:700;">${{c.store_mix_pct}}</td>
          <td>${{c.asp}}</td>
        </tr>`;
      }});
      document.getElementById("modal-cats-body").innerHTML = html || `<tr><td colspan="5" style="text-align:center;">No data</td></tr>`;
      document.getElementById("store-modal").style.display = "flex";
    }} catch(e) {{ console.error("openStoreModal error:", e); }}
  }}

  function closeModal() {{ document.getElementById("store-modal").style.display = "none"; }}

  function filterISTBrand(b) {{
    istFilterBrand = b;
    document.getElementById("btn-ist-all").classList.toggle('active', b === 'ALL');
    document.getElementById("btn-ist-dzl").classList.toggle('active', b === 'DZL');
    document.getElementById("btn-ist-mms").classList.toggle('active', b === 'MMS');
    renderISTTable();
  }}

  function renderISTTable() {{
    try {{
      const tbody = document.getElementById("replTableBody");
      let html = "";
      let idx = 1;
      REPL_ITEMS.forEach(r => {{
        if (istFilterBrand !== 'ALL' && r.brand !== istFilterBrand) return;
        const brandBadge = `<span class="badge" style="background:${{r.brand==='DZL'?'#ef444422':'#38bdf822'}}; color:${{r.brand==='DZL'?'#ef4444':'#38bdf8'}};">${{r.brand}}</span>`;
        const actionCol = r.action.includes("Warehouse") ? "#38bdf8" : "#f59e0b";
        html += `<tr>
          <td style="color:#64748b;">${{idx++}}</td>
          <td>${{brandBadge}}</td>
          <td><span class="badge" style="background:${{actionCol}}22; color:${{actionCol}};">${{r.action}}</span></td>
          <td style="color:#fff; font-weight:700;">${{r.store}}</td>
          <td style="color:#fff; font-weight:600;">${{r.focus}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.sold_qty}}</td>
          <td style="color:#f59e0b; font-weight:700;">${{r.store_soh}}</td>
          <td style="color:${{r.wh_soh>0?'#10b981':'#ef4444'}}; font-weight:700;">${{r.wh_soh}}</td>
          <td style="color:#10b981; font-weight:700;">${{r.qty}}</td>
          <td style="color:#38bdf8; font-size:11px;">${{r.source}}</td>
          <td><span class="badge" style="background:#ef444422; color:#ef4444;">${{r.urgency}}</span></td>
        </tr>`;
      }});
      tbody.innerHTML = html || `<tr><td colspan="11" style="text-align:center;">No recommendations available</td></tr>`;
    }} catch(e) {{ console.error("renderISTTable error:", e); }}
  }}

  function downloadISTPlan(b) {{
    try {{
      let rows = [["Brand", "Action Type", "Target Store", "SKU / Focus", "Sold MTD", "Store SOH", "WH SOH", "Replenish Qty", "Source Route", "Urgency"]];
      REPL_ITEMS.forEach(r => {{
        if (b === 'ALL' || r.brand === b) {{
          rows.push([r.brand, r.action, r.store, r.focus.replace(/,/g, ' '), r.sold_qty, r.store_soh, r.wh_soh, r.qty, r.source.replace(/,/g, ' '), r.urgency]);
        }}
      }});
      let csvContent = "data:text/csv;charset=utf-8,\uFEFF" + rows.map(e => e.join(",")).join("\\n");
      let encodedUri = encodeURI(csvContent);
      let link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", `Replenishment_Plan_${{b}}_2026.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }} catch(e) {{ console.error("downloadISTPlan error:", e); }}
  }}

  function switchDZLMovers(t) {{
    dzlMoversType = t;
    document.getElementById("btn-dzl-top").classList.toggle('active', t === 'top');
    document.getElementById("btn-dzl-low").classList.toggle('active', t === 'low');
    renderDZLMovers();
  }}

  function renderDZLMovers() {{
    try {{
      const st = document.getElementById("dzlStoreSelect").value;
      const items = DZL_MOVERS[st]?.[dzlMoversType] || DZL_MOVERS["ALL"]?.[dzlMoversType] || [];
      const tbody = document.getElementById("dzlMoversBody");
      let html = "";
      items.forEach((r, idx) => {{
        html += `<tr>
          <td style="color:${{dzlMoversType==='top'?'#10b981':'#ef4444'}}; font-weight:700;">#${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.clean_sku}}</td>
          <td style="color:#fff;">${{r.clean_name}}</td>
          <td><span class="badge" style="background:#ec489922; color:#ec4899;">${{r.gender}}</span></td>
          <td style="color:#38bdf8; font-weight:700;">${{r.units}}</td>
          <td style="color:#fff; font-weight:700;">${{r.sales.toLocaleString()}}</td>
          <td style="color:#f59e0b;">${{r.asp}}</td>
        </tr>`;
      }});
      tbody.innerHTML = html || `<tr><td colspan="7" style="text-align:center;">No data available</td></tr>`;
    }} catch(e) {{ console.error("renderDZLMovers error:", e); }}
  }}

  function switchMMSMovers(t) {{
    mmsMoversType = t;
    document.getElementById("btn-mms-top").classList.toggle('active', t === 'top');
    document.getElementById("btn-mms-low").classList.toggle('active', t === 'low');
    renderMMSMovers();
  }}

  function renderMMSMovers() {{
    try {{
      const data = (mmsMoversType === 'top') ? MMS_TOP500 : MMS_LOW500;
      const q = (document.getElementById("mmsSearch").value || "").toLowerCase();
      const tbody = document.getElementById("mmsMoversTableBody");
      let html = "";
      data.filter(r => r.clean_sku.toLowerCase().includes(q) || r.clean_name.toLowerCase().includes(q)).slice(0, 50).forEach((r, idx) => {{
        html += `<tr>
          <td style="color:${{mmsMoversType==='top'?'#10b981':'#ef4444'}}; font-weight:700;">#${{idx+1}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.clean_sku}}</td>
          <td style="color:#fff;">${{r.clean_name}}</td>
          <td style="color:#94a3b8;">${{r.main_category}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.units.toLocaleString()}}</td>
          <td style="color:#fff; font-weight:700;">${{Math.round(r.sales).toLocaleString()}}</td>
          <td style="color:#f59e0b;">${{r.asp}}</td>
        </tr>`;
      }});
      tbody.innerHTML = html;
    }} catch(e) {{ console.error("renderMMSMovers error:", e); }}
  }}

  function switchMonth(m) {{
    try {{
      activeMonth = m;
      if (m === "SEP") {{
        currentBrandTotals = SEP_BRAND_TOTALS;
        currentStoreMeta = SEP_STORE_META;
        document.getElementById("headerSubtitle").innerText = "September 2026 Full Monthly Performance & Benchmarking (Archived)";
      }} else {{
        currentBrandTotals = OCT_BRAND_TOTALS;
        currentStoreMeta = OCT_STORE_META;
        document.getElementById("headerSubtitle").innerText = "October 2026 Daily Phasing & Commercial Performance Tracking";
      }}

      let rowsHtml = "";
      let idx = 1;
      Object.values(currentStoreMeta).forEach(r => {{
        const yoyStr = (r.yoy !== null && !isNaN(r.yoy)) ? `<span style="color:${{r.yoy>=0?'#10b981':'#ef4444'}}; font-weight:700;">${{r.yoy.toFixed(1)}}%</span>` : `<span style="color:#64748b;">-</span>`;
        const lyStr = r.ly_sales ? r.ly_sales.toLocaleString() : `<span style="color:#64748b;">-</span>`;
        const achCol = (r.ach >= 100) ? "#10b981" : ((r.ach >= 80) ? "#f59e0b" : "#ef4444");
        const brandBadge = `<span class="badge" style="background:${{r.brand==='DZL'?'#ef444422':'#38bdf822'}}; color:${{r.brand==='DZL'?'#ef4444':'#38bdf8'}};">${{r.brand}}</span>`;

        rowsHtml += `<tr class="clickable-row store-row" data-brand="${{r.brand}}" data-region="${{r.region}}" onclick="openStoreModal('${{r.code}}')">
          <td style="color:#64748b; font-weight:600;">${{idx++}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.code}}</td>
          <td style="color:#fff; font-weight:600;">${{brandBadge}} ${{r.name}}</td>
          <td style="color:#94a3b8; font-size:12px;">${{r.region}}</td>
          <td style="color:#f8fafc; font-weight:700;">${{r.sales.toLocaleString()}}</td>
          <td style="color:#38bdf8;">${{lyStr}}</td>
          <td>${{yoyStr}}</td>
          <td style="color:#94a3b8;">${{Math.round(r.target).toLocaleString()}}</td>
          <td style="color:${{achCol}}; font-weight:700;">${{r.ach.toFixed(1)}}%</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.units.toLocaleString()}}</td>
          <td style="color:#fff; font-weight:700;">${{r.txns.toLocaleString()}}</td>
          <td style="color:#10b981; font-weight:700;">${{r.upt.toFixed(2)}}</td>
          <td style="color:#38bdf8; font-weight:700;">${{r.str_pct}}%</td>
          <td><span class="badge" style="background:#10b98122; color:#10b981;">Archived</span></td>
          <td style="color:#f59e0b; font-weight:700;">${{r.asp}}</td>
        </tr>`;
      }});
      document.getElementById("storesTableBody").innerHTML = rowsHtml;

      updateKPICards(activeBrand);
      renderRegionTables();
    }} catch(e) {{ console.error("switchMonth error:", e); }}
  }}
</script>

</body>
</html>
"""

    out_file = os.path.join(REPORTS_DIR, "MMS_Executive_KPI_Dashboard.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(template_html)

    print(f"[✓] Dashboard generated successfully: {out_file}")

if __name__ == "__main__":
    process_and_build()