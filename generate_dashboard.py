import os
import glob
import re
import json
import html
import pandas as pd
import numpy as np

REPORTS_DIR = "./reports"

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

    # DZL (DOZOLO) STORES (3 Active Doors)
    "K107": {"full_name": "DZL Riyadh Park", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K111": {"full_name": "DZL Solitaire", "region": "Riyadh Central Region", "manager": "Sultan", "city": "Riyadh", "brand": "DZL"},
    "K204": {"full_name": "DZL Redsea", "region": "Western Region", "manager": "Rajib", "city": "Jeddah", "brand": "DZL"}
}

DZL_VALID_CODES = {"K107", "K111", "K204"}
MMS_VALID_CODES = {k for k, v in STORE_MAPPING.items() if v["brand"] == "MMS"}
ALL_VALID_CODES = DZL_VALID_CODES.union(MMS_VALID_CODES)

EXCLUDED_KEYWORDS = [
    'gwp', 'shopping bag', 'carrier bag', 'plastic bag', 'paper bag',
    'gift with purchase', 'non-sale', 'packaging', 'free gift', 'material',
    'stationery', 'basketball', 'football', 'm&g'
]

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
        if 23 <= size <= 34:
            return 'Kids'
        elif 35 <= size <= 39:
            return 'Women'
        elif 40 <= size <= 48:
            return 'Men'
            
    s = " " + text.replace('-', ' ').replace('_', ' ').replace('/', ' ') + " "
    if any(k in s for k in [' KID ', ' KIDS ', ' BOY ', ' BOYS ', ' GIRL ', ' GIRLS ', ' CHILD ', ' CHILDREN ', ' GS ']):
        return 'Kids'
    if any(k in s for k in [' WOMAN ', ' WOMEN ', ' WOMENS ', ' LADY ', ' LADIES ', ' FEMALE ', ' WMNS ']):
        return 'Women'
    if any(k in s for k in [' MAN ', ' MEN ', ' MENS ', ' MALE ']):
        return 'Men'
        
    return 'Women'

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
                    is_dzl_col = "(DZL)" in col_str.upper()
                    if is_dzl_col and code in DZL_VALID_CODES:
                        tot_val = pd.to_numeric(df_gsale[col], errors='coerce').sum()
                        if pd.notna(tot_val) and tot_val > 0:
                            ly_totals[code] = round(float(tot_val))
                    elif not is_dzl_col and code in MMS_VALID_CODES:
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
            if c_code in ALL_VALID_CODES:
                t_map[c_code] = float(r[sep_col])
                t_map[str(r[store_col]).strip().upper()] = float(r[sep_col])
        return t_map
    except Exception as e:
        print(f"Target load error: {e}")
        return {}

def load_soh_data(soh_path):
    if not soh_path or not os.path.exists(soh_path):
        return {}, {}, {}, {}, pd.DataFrame(), 0
    soh_store_summary = {}
    soh_hierarchy_map = {}
    sku_to_pg_map = {}
    sku_to_cat_map = {}
    wh_total_stock = 0
    try:
        xl = pd.ExcelFile(soh_path)
        sheet_to_use = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
        
        df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use, skiprows=1)
        df_soh.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]
        
        if not any(k in [str(c).lower() for c in df_soh.columns] for k in ["avail_stock", "current_stock", "stock"]):
            df_soh = pd.read_excel(soh_path, sheet_name=sheet_to_use)
            df_soh.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_soh.columns]

        code_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["org code", "organization code", "org_code", "store code", "org_no"])), None)
        stock_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["avail_stock", "current_stock", "stock", "qty"])), None)
        price_col = next((c for c in df_soh.columns if "retail_price" in str(c).lower() or "price" in str(c).lower()), None)
        cat_col = next((c for c in df_soh.columns if str(c).lower() == "category"), None)
        pg_col = next((c for c in df_soh.columns if str(c).lower() in ["product_group", "product group"]), None)
        desc_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["product name", "item name", "description"])), None)
        item_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["item code", "item_code", "sku code"])), None)
        barcode_col = next((c for c in df_soh.columns if any(k in str(c).lower() for k in ["barcode", "bar code", "upc"])), None)

        if not code_col or not stock_col:
            return {}, {}, {}, {}, pd.DataFrame(), 0

        df_soh = df_soh[df_soh[code_col].notna()].copy()
        df_soh[stock_col] = pd.to_numeric(df_soh[stock_col], errors='coerce').fillna(0)

        if desc_col:
            for kw in EXCLUDED_KEYWORDS:
                df_soh = df_soh[~df_soh[desc_col].astype(str).str.lower().str.contains(kw, regex=False)]
        if cat_col:
            for kw in ['gwp', 'material', 'me+']:
                df_soh = df_soh[~df_soh[cat_col].astype(str).str.lower().str.contains(kw, regex=False)]

        if price_col:
            df_soh[price_col] = pd.to_numeric(df_soh[price_col], errors='coerce').fillna(0)
            df_soh['soh_val'] = df_soh[stock_col] * df_soh[price_col]
        else:
            df_soh['soh_val'] = 0.0

        if pg_col:
            if barcode_col:
                for _, r in df_soh[[barcode_col, pg_col]].dropna().drop_duplicates().iterrows():
                    b_c = clean_sku_code(r[barcode_col])
                    pg_v = str(r[pg_col]).strip()
                    if b_c and pg_v: sku_to_pg_map[b_c] = pg_v
            if item_col:
                for _, r in df_soh[[item_col, pg_col]].dropna().drop_duplicates().iterrows():
                    i_c = clean_sku_code(r[item_col])
                    pg_v = str(r[pg_col]).strip()
                    if i_c and pg_v: sku_to_pg_map[i_c] = pg_v

        if cat_col:
            if barcode_col:
                for _, r in df_soh[[barcode_col, cat_col]].dropna().drop_duplicates().iterrows():
                    b_c = clean_sku_code(r[barcode_col])
                    c_v = str(r[cat_col]).replace('_', ' ').strip()
                    if b_c and c_v: sku_to_cat_map[b_c] = c_v
            if item_col:
                for _, r in df_soh[[item_col, cat_col]].dropna().drop_duplicates().iterrows():
                    i_c = clean_sku_code(r[item_col])
                    c_v = str(r[cat_col]).replace('_', ' ').strip()
                    if i_c and c_v: sku_to_cat_map[i_c] = c_v

        def parse_soh_code(v):
            s = str(v).strip().upper()
            if "KSWH" in s or s == "WH": return "KSWH"
            return clean_store_code_str(s)

        df_soh['clean_code'] = df_soh[code_col].apply(parse_soh_code)

        wh_df = df_soh[df_soh['clean_code'] == 'KSWH']
        if not wh_df.empty: wh_total_stock = int(wh_df[stock_col].sum())

        stores_soh_df = df_soh[(df_soh['clean_code'] != 'KSWH') & (df_soh['clean_code'].isin(ALL_VALID_CODES))]
        grouped = stores_soh_df.groupby('clean_code', as_index=False).agg({
            stock_col: 'sum',
            'soh_val': 'sum'
        })
        grouped.rename(columns={stock_col: 'soh_units'}, inplace=True)
        soh_store_summary = grouped.set_index('clean_code').to_dict(orient='index')
        return soh_store_summary, soh_hierarchy_map, sku_to_pg_map, sku_to_cat_map, df_soh, wh_total_stock
    except Exception as e:
        print(f"[!] Error processing SOH: {e}")
        return {}, {}, {}, {}, pd.DataFrame(), 0

def process_and_build():
    replenishment_recommendations = []
    excel_export_mms = []
    excel_export_dzl = []
    store_table_rows = ""
    decision_cards_html = ""
    region_kpi_cards = ""
    region_tables_html = ""
    repl_rows_html = ""

    sales_mms_file, sales_dzl_file, soh_file, target_file, ly_file = identify_files()
    targets_map = load_targets(target_file)
    soh_map, soh_hier_map, sku_to_pg_map, sku_to_cat_map, df_soh_raw, wh_total_stock = load_soh_data(soh_file)
    ly_sales_map = load_ly_sales_data(ly_file)

    dfs = []
    if os.path.exists(sales_mms_file):
        df_m = pd.read_excel(sales_mms_file, skiprows=1)
        df_m = df_m.iloc[:-1].copy()
        df_m.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_m.columns]
        df_m['brand_origin'] = "MMS"
        dfs.append(df_m)

    if os.path.exists(sales_dzl_file):
        try:
            df_d = pd.read_excel(sales_dzl_file)
            has_org = any("org" in str(c).lower() or "store" in str(c).lower() for c in df_d.columns)
            if not has_org:
                df_d = pd.read_excel(sales_dzl_file, skiprows=1)
            df_d = df_d.iloc[:-1].copy() if len(df_d) > 1 and "total" in str(df_d.iloc[-1].values).lower() else df_d
            df_d.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_d.columns]
            df_d['brand_origin'] = "DZL"
            dfs.append(df_d)
        except Exception as e:
            print(f"[!] Warning reading DZL sales: {e}")

    df_clean = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    df_clean.columns = [str(c).replace('\u200c', '').replace('\ufeff', '').strip() for c in df_clean.columns]

    numeric_cols = [
        'Sales Quantity', 'Selling Price', 'Sales Revenue', 'Discount Amount',
        'Actual Sales Amount', 'Tax-excluded Actual Sales', 'Tax Amount'
    ]
    for col in numeric_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)

    org_code_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['organization code', 'org code', 'store code', 'shop code'])), df_clean.columns[0])
    org_name_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['organization name', 'org name', 'store name', 'shop name'])), org_code_col)
    item_code_col = next((c for c in df_clean.columns if str(c).lower() in ['product code', 'item code', 'sku code', 'product no', 'item_code']), df_clean.columns[0])
    item_name_col = next((c for c in df_clean.columns if str(c).lower() in ['product name', 'item name', 'product_name']), item_code_col)
    barcode_col = next((c for c in df_clean.columns if any(k in str(c).lower() for k in ['barcode', 'bar code', 'upc'])), item_code_col)
    raw_subsub_col = next((c for c in df_clean.columns if str(c).lower() in ['category name', 'product_category', 'category']), None)

    df_clean['clean_code'] = df_clean[org_code_col].apply(clean_store_code_str)
    df_clean = df_clean[df_clean['clean_code'].isin(ALL_VALID_CODES)].copy()
    df_clean['brand'] = df_clean['clean_code'].apply(lambda c: STORE_MAPPING[c]['brand'])
    df_clean['clean_org_name'] = df_clean[org_name_col].astype(str).str.strip()
    df_clean['clean_sku'] = df_clean[item_code_col].apply(clean_sku_code)
    df_clean['clean_barcode'] = df_clean[barcode_col].apply(clean_sku_code)
    df_clean['clean_item_name'] = df_clean[item_name_col].fillna("Item").astype(str).str.strip()
    df_clean['sub_subgroup'] = df_clean[raw_subsub_col].fillna("Other").astype(str).str.strip() if raw_subsub_col else "General"

    # حذف مبيعات الصفر و GWP والأكياس نهائياً
    df_clean = df_clean[(df_clean['Actual Sales Amount'] > 0) & (df_clean['Sales Quantity'] > 0)].copy()
    for kw in EXCLUDED_KEYWORDS:
        df_clean = df_clean[~df_clean['clean_item_name'].str.lower().str.contains(kw, regex=False)]
        df_clean = df_clean[~df_clean['sub_subgroup'].str.lower().str.contains(kw, regex=False)]

    def resolve_style_group(row):
        b_c = row['clean_barcode']
        i_c = row['clean_sku']
        if b_c in sku_to_pg_map: return sku_to_pg_map[b_c]
        if i_c in sku_to_pg_map: return sku_to_pg_map[i_c]
        return derive_style_group_from_name(row['clean_item_name'])

    df_clean['style_group'] = df_clean.apply(resolve_style_group, axis=1)

    def map_to_main_category(row):
        b_c = row['clean_barcode']
        i_c = row['clean_sku']
        soh_cat = sku_to_cat_map.get(b_c, sku_to_cat_map.get(i_c, "")).lower()
        sub_l = str(row['sub_subgroup']).lower()
        item_l = str(row['clean_item_name']).lower()
        comb = f"{soh_cat} {sub_l} {item_l}"

        if row['brand'] == 'DZL':
            if 'access' in soh_cat or any(x in comb for x in ['sock', 'insole', 'foot loop', 'bag', 'hat', 'cap', 'belt', 'wallet', 'shoelace', 'cleaner', 'care']):
                return "Accessories"
            return "Shoes"

        if any(x in comb for x in ['toy', 'doll', 'clay', 'puzzle', 'baby', 'block', 'gun', 'bubble']): return "Children's Goods"
        if any(x in comb for x in ['lip', 'mask', 'cream', 'perfume', 'makeup', 'eyebrow', 'clean', 'wipe', 'bath', 'nail', 'soap']): return "Beauty & Cleaning"
        if any(x in comb for x in ['pen', 'notebook', 'tape', 'sticker', 'stationery', 'pencil', 'eraser']): return "Stationery"
        if any(x in comb for x in ['cup', 'mat', 'storage', 'kitchen', 'umbrella', 'fragrance', 'hanger']): return "Home & Daily Use"
        if any(x in comb for x in ['cable', 'headphone', 'fan', 'usb', 'charger', 'watch', 'phone']): return "3C Electronics"
        if any(x in comb for x in ['bag', 'backpack', 'wallet', 'purse']): return "Bags"
        if any(x in comb for x in ['sock', 'slipper', 'hat', 'sunglass', 'glove']): return "Apparel Accessories"
        if any(x in comb for x in ['hair', 'earring', 'clip', 'necklace', 'jewelry']): return "Fashion Accessories"
        if any(x in comb for x in ['pillow', 'towel', 'cushion', 'eyemask']): return "Home Textile"
        return "Variety Lifestyle"

    df_clean['main_category'] = df_clean.apply(map_to_main_category, axis=1)
    df_clean['gender'] = df_clean.apply(lambda r: classify_shoe_gender_by_size(r['clean_item_name']), axis=1)

    txn_col = 'Receipt Number' if 'Receipt Number' in df_clean.columns else 'clean_sku'
    txn_agg_func = 'nunique' if 'Receipt Number' in df_clean.columns else 'count'

    store_summary = df_clean.groupby(['clean_code', 'brand'], as_index=False).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum'),
        txns=(txn_col, txn_agg_func)
    )

    store_summary['full_name'] = store_summary['clean_code'].apply(lambda c: STORE_MAPPING[c]['full_name'])
    store_summary['region'] = store_summary['clean_code'].apply(lambda c: STORE_MAPPING[c]['region'])
    store_summary['manager'] = store_summary['clean_code'].apply(lambda c: STORE_MAPPING[c]['manager'])

    store_summary['atv'] = (store_summary['sales'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round().astype(int)
    store_summary['upt'] = (store_summary['units'] / store_summary['txns'].replace(0, np.nan)).fillna(0).round(2)
    store_summary['asp'] = (store_summary['sales'] / store_summary['units'].replace(0, np.nan)).fillna(0).round().astype(int)

    total_sales = round(store_summary['sales'].sum())
    total_txns = int(store_summary['txns'].sum())
    total_units = int(store_summary['units'].sum())
    network_atv = round(total_sales / total_txns) if total_txns > 0 else 0
    network_upt = (total_units / total_txns) if total_txns > 0 else 0
    network_asp = round(total_sales / total_units) if total_units > 0 else 0

    store_summary['share'] = ((store_summary['sales'] / total_sales) * 100).round(1) if total_sales > 0 else 0
    store_summary = store_summary.sort_values(by='sales', ascending=False).reset_index(drop=True)

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
        return {"soh_units": 0, "soh_val": 0.0}

    soh_matched = store_summary['clean_code'].apply(match_soh)
    store_summary['soh_units'] = [int(x.get('soh_units', 0)) for x in soh_matched]
    store_summary['soh_val'] = [float(x.get('soh_val', 0.0)) for x in soh_matched]

    store_summary['weekly_sales_units'] = (store_summary['units'] / 4.0).replace(0, np.nan)
    store_summary['woc'] = (store_summary['soh_units'] / store_summary['weekly_sales_units']).fillna(0).round(1)
    store_summary['str_pct'] = (store_summary['units'] / (store_summary['units'] + store_summary['soh_units']).replace(0, np.nan) * 100).fillna(0).round(1)

    def match_target_val(row):
        c_code = row['clean_code']
        if c_code in targets_map: return targets_map[c_code]
        for k, v in targets_map.items():
            if str(k).upper() in c_code or c_code.replace("K", "") == str(k).upper():
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

    brand_kpi_summary = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_df = store_summary if b == 'ALL' else store_summary[store_summary['brand'] == b]
        b_sales = int(sub_df['sales'].sum())
        b_target = int(sub_df['target'].fillna(0).sum())
        
        sub_lfl = sub_df[sub_df['ly_sales'].notna()]
        b_ly = int(sub_lfl['ly_sales'].sum())
        b_cur_lfl = sub_lfl['sales'].sum()
        
        b_units = int(sub_df['units'].sum())
        b_txns = int(sub_df['txns'].sum())
        b_soh = int(sub_df['soh_units'].sum())
        b_ach = round((b_sales / b_target * 100), 1) if b_target > 0 else 0
        b_yoy_val = round(((b_cur_lfl - b_ly) / b_ly * 100), 1) if b_ly > 0 else 0
        b_atv = int(round(b_sales / b_txns)) if b_txns > 0 else 0
        b_upt = round(b_units / b_txns, 2) if b_txns > 0 else 0
        b_asp = int(round(b_sales / b_units)) if b_units > 0 else 0
        b_str = round((b_units / (b_units + b_soh) * 100), 1) if (b_units + b_soh) > 0 else 0
        
        brand_kpi_summary[b] = {
            "sales": f"{b_sales:,}", "target": f"{b_target:,}", "ly": f"{b_ly:,}",
            "ach": f"{b_ach}%", "yoy": f"{b_yoy_val:+.1f}%", "yoy_val": b_yoy_val,
            "atv": f"{b_atv:,}", "units": f"{b_units:,}", "txns": f"{b_txns:,}",
            "upt": f"{b_upt:.2f}", "str": f"{b_str}%", "asp": f"{b_asp:,}"
        }

    store_chart_data = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_st = store_summary if b == 'ALL' else store_summary[store_summary['brand'] == b]
        limit = 12 if b != 'DZL' else 3
        sub_chart = sub_st.sort_values(by='sales', ascending=False).head(limit)
        store_chart_data[b] = {
            "categories": [str(r['full_name']).replace("MMS Riyadh ", "").replace("MMS ", "").replace("DZL ", "") for _, r in sub_chart.iterrows()],
            "sales": [round(float(r['sales'])) for _, r in sub_chart.iterrows()],
            "targets": [round(float(r['target'])) if pd.notna(r['target']) else 0 for _, r in sub_chart.iterrows()]
        }

    region_kpis_by_brand = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_st = store_summary if b == 'ALL' else store_summary[store_summary['brand'] == b]
        region_kpis_by_brand[b] = {}
        for reg in ["Riyadh Central Region", "Western Region"]:
            grp = sub_st[sub_st['region'] == reg]
            r_sales = int(grp['sales'].sum())
            r_target = int(grp['target'].fillna(0).sum())
            r_units = int(grp['units'].sum())
            r_txns = int(grp['txns'].sum())
            r_ach = round((r_sales / r_target * 100), 1) if r_target > 0 else 0
            
            grp_lfl = grp[grp['ly_sales'].notna()]
            r_ly = int(grp_lfl['ly_sales'].sum())
            r_cur_lfl = grp_lfl['sales'].sum()
            r_yoy_val = round(((r_cur_lfl - r_ly) / r_ly * 100), 1) if r_ly > 0 else 0
            
            r_atv = int(round(r_sales / r_txns)) if r_txns > 0 else 0
            r_upt = round(r_units / r_txns, 2) if r_txns > 0 else 0
            r_asp = int(round(r_sales / r_units)) if r_units > 0 else 0

            region_kpis_by_brand[b][reg] = {
                "sales": f"{r_sales:,}", "target": f"{r_target:,}", "ly": f"{r_ly:,}",
                "ach": f"{r_ach}%", "yoy": f"{r_yoy_val:+.1f}%", "yoy_val": r_yoy_val,
                "units": f"{r_units:,}", "txns": f"{r_txns:,}", "atv": f"{r_atv:,}",
                "upt": f"{r_upt:.2f}", "asp": f"{r_asp:,}", "stores_count": len(grp)
            }

    category_data_by_brand = {}
    for b in ['ALL', 'MMS', 'DZL']:
        sub_c = df_clean if b == 'ALL' else df_clean[df_clean['brand'] == b]
        b_tot_sales = sub_c['Actual Sales Amount'].sum()
        b_cat_df = sub_c.groupby('main_category', as_index=False).agg(
            sales=('Actual Sales Amount', 'sum'),
            units=('Sales Quantity', 'sum')
        ).sort_values(by='sales', ascending=False).reset_index(drop=True)
        b_cat_df['contribution'] = ((b_cat_df['sales'] / b_tot_sales) * 100).round(1) if b_tot_sales > 0 else 0
        b_cat_df['asp'] = (b_cat_df['sales'] / b_cat_df['units'].replace(0, np.nan)).fillna(0).round().astype(int)
        category_data_by_brand[b] = b_cat_df.to_dict(orient='records')

    dzl_shoes_only = df_clean[(df_clean['brand'] == 'DZL') & (df_clean['main_category'] == 'Shoes')]
    dzl_shoe_sales = dzl_shoes_only['Actual Sales Amount'].sum()
    dzl_gender_df = dzl_shoes_only.groupby('gender', as_index=False).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).sort_values(by='sales', ascending=False).reset_index(drop=True)
    dzl_gender_df['share'] = ((dzl_gender_df['sales'] / dzl_shoe_sales) * 100).round(1) if dzl_shoe_sales > 0 else 0
    dzl_gender_df['asp'] = (dzl_gender_df['sales'] / dzl_gender_df['units'].replace(0, np.nan)).fillna(0).round().astype(int)
    dzl_gender_data = {
        "labels": dzl_gender_df['gender'].tolist(),
        "series": dzl_gender_df['share'].tolist(),
        "records": dzl_gender_df.to_dict(orient='records')
    }

    wh_barcode_stock_dict = {}
    store_barcode_stock_dict = {}
    barcode_soh_col = next((c for c in df_soh_raw.columns if any(k in str(c).lower() for k in ["barcode", "bar code", "upc"])), None)
    stock_col_name = next((c for c in df_soh_raw.columns if any(k in str(c).lower() for k in ["avail_stock", "current_stock", "stock", "qty"])), None)
    code_col_name = next((c for c in df_soh_raw.columns if any(k in str(c).lower() for k in ["org code", "organization code", "org_code", "store code"])), None)

    if not df_soh_raw.empty and stock_col_name and code_col_name and barcode_soh_col:
        df_soh_raw['store_code'] = df_soh_raw[code_col_name].apply(clean_store_code_str)
        df_soh_raw['clean_barcode'] = df_soh_raw[barcode_soh_col].apply(clean_sku_code)
        wh_grouped = df_soh_raw[df_soh_raw['store_code'] == 'KSWH'].groupby('clean_barcode', as_index=False)[stock_col_name].sum()
        wh_barcode_stock_dict = dict(zip(wh_grouped['clean_barcode'], wh_grouped[stock_col_name]))
        
        st_grouped = df_soh_raw[df_soh_raw['store_code'].isin(ALL_VALID_CODES)].groupby('clean_barcode', as_index=False)[stock_col_name].sum()
        store_barcode_stock_dict = dict(zip(st_grouped['clean_barcode'], st_grouped[stock_col_name]))

    store_category_details = {}
    for code, grp in df_clean.groupby('clean_code'):
        st_c = clean_store_code_str(code)
        st_tot = grp['Actual Sales Amount'].sum()
        cats_list = []
        c_agg = grp.groupby('main_category', as_index=False).agg(
            c_sales=('Actual Sales Amount', 'sum'),
            c_units=('Sales Quantity', 'sum')
        ).sort_values(by='c_sales', ascending=False)
        for _, r in c_agg.iterrows():
            cs = round(r['c_sales'])
            cu = int(r['c_units'])
            cmix = round((cs / st_tot * 100), 1) if st_tot > 0 else 0
            casp = round(cs / cu) if cu > 0 else 0
            cats_list.append({
                "main_category": r['main_category'],
                "sales": f"{cs:,}",
                "units": f"{cu:,}",
                "store_mix_pct": f"{cmix}%",
                "asp": f"{casp:,}"
            })
        store_category_details[st_c] = cats_list

    drilldown_items = df_clean.groupby(['brand', 'main_category', 'gender', 'style_group', 'clean_barcode', 'clean_item_name'], as_index=False).agg(
        sales=('Actual Sales Amount', 'sum'),
        units=('Sales Quantity', 'sum')
    ).sort_values(by='sales', ascending=False)

    drilldown_items['wh_soh'] = drilldown_items['clean_barcode'].map(wh_barcode_stock_dict).fillna(0).astype(int)
    drilldown_items['asp'] = (drilldown_items['sales'] / drilldown_items['units'].replace(0, np.nan)).fillna(0).round().astype(int)
    business_drilldown_data = drilldown_items.to_dict(orient='records')

    dzl_shoes_df = df_clean[(df_clean['brand'] == 'DZL') & (df_clean['main_category'] == 'Shoes') & (df_clean['Actual Sales Amount'] > 0)].copy()
    dzl_top20_groups = []
    dzl_low20_groups = []

    if not dzl_shoes_df.empty:
        dzl_tot_shoe_sales = dzl_shoes_df['Actual Sales Amount'].sum()
        dzl_pg_summary = dzl_shoes_df.groupby('style_group', as_index=False).agg(
            sales=('Actual Sales Amount', 'sum'),
            units=('Sales Quantity', 'sum')
        )
        dzl_pg_summary['asp'] = (dzl_pg_summary['sales'] / dzl_pg_summary['units'].replace(0, np.nan)).fillna(0).round().astype(int)
        dzl_pg_summary['share_pct'] = ((dzl_pg_summary['sales'] / dzl_tot_shoe_sales) * 100).round(1) if dzl_tot_shoe_sales > 0 else 0
        
        sku_details = dzl_shoes_df.groupby(['style_group', 'clean_barcode', 'clean_sku', 'clean_item_name', 'gender'], as_index=False).agg(
            sku_sales=('Actual Sales Amount', 'sum'),
            sku_units=('Sales Quantity', 'sum')
        ).sort_values(by='sku_units', ascending=False)

        def build_group_data(df_slice):
            out_list = []
            for idx, r in df_slice.reset_index(drop=True).iterrows():
                pg_name = str(r['style_group'])
                child_skus = sku_details[sku_details['style_group'] == pg_name]
                child_list = []
                tot_wh_stock = 0
                tot_st_stock = 0
                gender_val = child_skus['gender'].iloc[0] if not child_skus.empty else "Women"
                for _, cr in child_skus.iterrows():
                    b_code = str(cr['clean_barcode'])
                    wh_s = int(wh_barcode_stock_dict.get(b_code, 0))
                    st_s = int(store_barcode_stock_dict.get(b_code, 0))
                    tot_wh_stock += wh_s
                    tot_st_stock += st_s
                    child_list.append({
                        "barcode": b_code,
                        "item_code": str(cr['clean_sku']),
                        "name": str(cr['clean_item_name'])[:45],
                        "units": int(cr['sku_units']),
                        "sales": int(round(cr['sku_sales'])),
                        "store_soh": st_s,
                        "wh_soh": wh_s
                    })
                daily_v = r['units'] / 30.0
                days_left = int((tot_st_stock + tot_wh_stock) / daily_v) if daily_v > 0 else 999
                out_list.append({
                    "rank": idx + 1,
                    "style_group": pg_name,
                    "gender": gender_val,
                    "units": int(r['units']),
                    "sales": int(round(r['sales'])),
                    "share_pct": float(r['share_pct']),
                    "asp": int(r['asp']),
                    "store_soh": tot_st_stock,
                    "wh_soh": tot_wh_stock,
                    "stock_days": days_left,
                    "skus_count": len(child_list),
                    "skus": child_list
                })
            return out_list

        dzl_top20_groups = build_group_data(dzl_pg_summary.sort_values(by=['units', 'sales'], ascending=[False, False]).head(20))
        dzl_low20_groups = build_group_data(dzl_pg_summary.sort_values(by=['units', 'sales'], ascending=[True, True]).head(20))

    if not df_soh_raw.empty and stock_col_name and code_col_name and barcode_soh_col:
        store_sku_sales = df_clean.groupby(['clean_code', 'clean_barcode', 'clean_sku', 'clean_item_name', 'style_group', 'main_category', 'brand'], as_index=False).agg(
            sept_units=('Sales Quantity', 'sum')
        )

        soh_subset = df_soh_raw.groupby(['store_code', 'clean_barcode'], as_index=False)[stock_col_name].sum()
        merged_sku = pd.merge(
            store_sku_sales,
            soh_subset,
            left_on=['clean_code', 'clean_barcode'],
            right_on=['store_code', 'clean_barcode'],
            how='inner'
        )
        merged_sku.rename(columns={stock_col_name: 'store_soh'}, inplace=True)
        merged_sku['daily_rate'] = merged_sku['sept_units'] / 30.0
        merged_sku['days_to_stockout'] = merged_sku['store_soh'] / merged_sku['daily_rate'].replace(0, np.nan)
        
        critical_shoes = merged_sku[(merged_sku['main_category'] == 'Shoes') & ((merged_sku['days_to_stockout'] < 12.0) | (merged_sku['store_soh'] <= 2)) & (merged_sku['sept_units'] >= 1)].copy()
        critical_shoes['priority_rank'] = 1
        
        critical_acc = merged_sku[(merged_sku['main_category'] != 'Shoes') & (merged_sku['days_to_stockout'] < 8.0) & (merged_sku['sept_units'] >= 2)].copy()
        critical_acc['priority_rank'] = 2

        all_critical = pd.concat([critical_shoes, critical_acc], ignore_index=True)
        all_critical = all_critical.sort_values(by=['priority_rank', 'days_to_stockout', 'sept_units'], ascending=[True, True, False])

        for _, row in all_critical.iterrows():
            st_code = row['clean_code']
            st_brand = row['brand']
            st_info = STORE_MAPPING[st_code]
            st_name = st_info['full_name']
            st_city = st_info['city']

            b_val = row['clean_barcode']
            sku_code = row['clean_sku']
            sku_name = str(row['clean_item_name'])[:35]
            style_grp = str(row['style_group'])[:25]
            cat = row['main_category']
            days_left = int(row['days_to_stockout']) if pd.notna(row['days_to_stockout']) else 0
            
            daily_v = row['daily_rate']
            needed_qty = max(6 if cat == 'Shoes' else 4, int((daily_v * 28) - row['store_soh']))
            wh_available = int(wh_barcode_stock_dict.get(b_val, 0))
            
            if row['store_soh'] <= 0:
                urgency_str = "🚨 Out of Stock (0 Pcs left)"
            elif row['store_soh'] <= 2 and cat == 'Shoes':
                urgency_str = f"⚠️ Broken Size Run ({int(row['store_soh'])} Pcs left)"
            else:
                urgency_str = f"⚠️ Stock-Out in {days_left}d (Vel: {daily_v:.1f}/d)"

            if wh_available >= needed_qty:
                action_type = "Predictive WH Replenishment"
                source_route = f"Central Warehouse (KSWH - Avail: {wh_available:,})"
            else:
                action_type = "Store Transfer (IST)"
                surplus_branches = df_soh_raw[
                    (df_soh_raw['clean_barcode'] == b_val) & 
                    (df_soh_raw['store_code'] != 'KSWH') & 
                    (df_soh_raw['store_code'] != st_code) & 
                    (df_soh_raw['store_code'].isin(DZL_VALID_CODES if st_brand == 'DZL' else MMS_VALID_CODES)) &
                    (df_soh_raw[stock_col_name] > (4 if cat == 'Shoes' else 8))
                ].copy()
                
                if not surplus_branches.empty:
                    surplus_branches['donor_city'] = surplus_branches['store_code'].apply(lambda c: STORE_MAPPING.get(c, {}).get('city', ''))
                    surplus_branches['city_match'] = (surplus_branches['donor_city'] == st_city).astype(int)
                    donor_row = surplus_branches.sort_values(by=['city_match', stock_col_name], ascending=[False, False]).iloc[0]
                    donor_code = donor_row['store_code']
                    donor_name = STORE_MAPPING[donor_code]['full_name']
                    donor_qty = int(donor_row[stock_col_name])
                    match_type = "🏙️ Same City" if donor_row['city_match'] else "🚛 Inter-City"
                    source_route = f"{donor_name} ({donor_code} - Surplus: {donor_qty}) [{match_type}]"
                    needed_qty = min(needed_qty, max(2, donor_qty // 2))
                else:
                    action_type = "Predictive WH Replenishment"
                    source_route = f"Central Warehouse (KSWH - Limited: {wh_available})"

            p_icon = "👟 [SHOE PRIORITY 1]" if cat == 'Shoes' else "👜 [ACCESSORY]"
            rep_item = {
                "brand": st_brand, "type": action_type, "store_name": f"{st_name} ({st_code})",
                "category_focus": f"{p_icon} {cat} | [{style_grp}] {sku_name} (Barcode: {b_val})",
                "from_source": source_route, "suggested_units": f"{needed_qty:,} Pcs",
                "urgency": urgency_str
            }
            replenishment_recommendations.append(rep_item)

            export_entry = {
                "Brand": st_brand, "Priority": "1 - High (Shoes)" if cat == 'Shoes' else "2 - Routine (Accessory)",
                "Action Type": action_type, "Store Code": st_code, "Store Name": st_name,
                "Product Group / Style": style_grp, "Barcode / SKC": b_val, "Item Code": sku_code,
                "Product Name": sku_name, "Category": cat,
                "Store SOH": int(row['store_soh']), "WH SOH (KSWH)": wh_available, "Daily Velocity": round(daily_v, 1),
                "Urgency Status": urgency_str, "Suggested QTY (Pcs)": needed_qty,
                "Source Route": source_route
            }

            if st_brand == "DZL":
                excel_export_dzl.append(export_entry)
            else:
                excel_export_mms.append(export_entry)

    try:
        if excel_export_mms:
            pd.DataFrame(excel_export_mms).to_excel(os.path.join(REPORTS_DIR, "Auto_Replenishment_MMS.xlsx"), index=False)
        if excel_export_dzl:
            pd.DataFrame(excel_export_dzl).to_excel(os.path.join(REPORTS_DIR, "Auto_Replenishment_DZL.xlsx"), index=False)
    except Exception as e:
        print(f"[!] Warning: Could not overwrite Excel files: {e}")

    def commercial_diagnosis_engine(row):
        st_code = row['clean_code']
        soh = row['soh_units']
        ach = row['ach_pct'] if pd.notna(row['ach_pct']) else 0
        woc = row['woc']
        st_items = df_clean[df_clean['clean_code'] == st_code]['main_category'].value_counts().index.tolist()
        st_top_cats_str = ", ".join(st_items[:2]) if st_items else "Core"

        woc_badge = f"OOS Risk ({woc} Wks)" if woc < 4.0 and woc > 0 else (f"Healthy ({woc} Wks)" if 4.0 <= woc <= 9.0 else f"Overstocked ({woc} Wks)")
        woc_col = "#ef4444" if woc < 4.0 and woc > 0 else ("#10b981" if 4.0 <= woc <= 9.0 else "#f59e0b")

        if ach >= 95:
            diag_title, diag_col = "Powerhouse Performer", "#10b981"
            action = f"Maintain 100% size-run availability on leading drivers in {st_top_cats_str}."
            needs = "Priority replenishment for fast-moving SKCs."
        elif ach < 70 and soh >= 5000:
            diag_title, diag_col = "Assortment Mismatch", "#f59e0b"
            action = f"⚡ ACTION: Reallocate front display to Top 20 high-velocity styles and balance broken sizes."
            needs = "Inject high-velocity categories & core sizes."
        else:
            diag_title, diag_col = "Steady Flow", "#38bdf8"
            action = f"Focus cashier upselling to lift ATV (Current: {row['atv']:,} SAR) and rotate feature end-caps."
            needs = "Routine weekly replenishment on core size runs."

        return {
            "woc_badge": woc_badge, "woc_color": woc_col,
            "diag_title": diag_title, "diag_color": diag_col,
            "action": action, "needs": needs, "top_categories_str": st_top_cats_str
        }

    engine_res = store_summary.apply(commercial_diagnosis_engine, axis=1)
    store_summary['woc_status'] = [e['woc_badge'] for e in engine_res]
    store_summary['woc_color'] = [e['woc_color'] for e in engine_res]
    store_summary['diag_title'] = [e['diag_title'] for e in engine_res]
    store_summary['diag_color'] = [e['diag_color'] for e in engine_res]
    store_summary['action'] = [e['action'] for e in engine_res]
    store_summary['needs'] = [e['needs'] for e in engine_res]
    store_summary['top_cats_str'] = [e['top_categories_str'] for e in engine_res]

    net_yoy_col = "#10b981" if network_lfl_growth >= 0 else "#ef4444"
    store_table_rows = ""
    decision_cards_html = ""
    store_meta_map = {}

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
        brand_badge = f'<span class="badge" style="background:{"#ef444422" if st_brand=="DZL" else "#38bdf822"}; color:{"#ef4444" if st_brand=="DZL" else "#38bdf8"}; border:1px solid {"#ef444455" if st_brand=="DZL" else "#38bdf855"};">{st_brand}</span>'

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
            "diag_title": row['diag_title'], "action": row['action'], "needs": row['needs'],
            "top_cats": row['top_cats_str']
        }

        decision_cards_html += f"""
        <div class="decision-card-item store-card-item" data-brand="{st_brand}" data-region="{row['region']}" style="background:var(--card); border:1px solid var(--border); border-left:4px solid {row['diag_color']}; border-radius:10px; padding:18px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                <div>
                    <span style="font-weight:700; color:#fff; font-size:15px;">{brand_badge} {st_name} ({st_code})</span>
                    <div style="font-size:11px; color:#94a3b8; margin-top:2px;">{row['region']} | SOH: <strong style="color:#38bdf8;">{int(row['soh_units']):,} Pcs</strong> ({row['woc']} Wks)</div>
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

        store_table_rows += f"""
        <tr onclick="openStoreDetails('{st_code}')" class="clickable-row store-row" data-brand="{st_brand}" data-region="{row['region']}" title="Click to view detailed store breakdown">
            <td style="color:#64748b;font-weight:600;">{idx+1}</td>
            <td style="color:#38bdf8;font-weight:600;">{st_code}</td>
            <td style="font-weight:600;color:#fff;">{brand_badge} {st_name}</td>
            <td style="color:#94a3b8;font-size:12px;">{row['region']}</td>
            <td style="font-weight:700;color:#f8fafc;">{round(row['sales']):,}</td>
            <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
            <td>{yoy_cell}</td>
            <td style="color:#94a3b8;">{target_str}</td>
            <td style="min-width:130px;">{ach_str}</td>
            <td style="font-weight:700;color:#38bdf8;">{int(row['units']):,}</td>
            <td style="font-weight:700;color:#fff;">{int(row['txns']):,}</td>
            <td style="font-weight:700;color:#10b981;">{row['upt']:.2f}</td>
            <td style="font-weight:700;color:#38bdf8;">{row['atv']:,}</td>
            <td>{row['str_pct']}%</td>
            <td>{diag_badge}</td>
            <td style="color:#f59e0b;font-weight:600;">{row['asp']:,}</td>
        </tr>
        """

    for reg_name, grp in [("Riyadh Central Region", store_summary[store_summary['region'] == "Riyadh Central Region"]),
                          ("Western Region", store_summary[store_summary['region'] == "Western Region"])]:
        reg_mgr = "Sultan" if "Riyadh" in reg_name else "Rajib"
        reg_id = "riyadh" if "Riyadh" in reg_name else "western"

        region_kpi_cards += f"""
        <div class="region-block" id="reg-card-{reg_id}" data-region="{reg_name}" style="background:var(--card); border:1px solid var(--border); border-top:4px solid {'#38bdf8' if 'Riyadh' in reg_name else '#818cf8'}; border-radius:12px; padding:20px; flex:1; min-width:320px;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                <div>
                    <h3 style="margin:0; font-size:17px; color:#fff;">{reg_name}</h3>
                    <span style="font-size:12px; color:#38bdf8; font-weight:600;">Area Manager: {reg_mgr}</span>
                </div>
                <div style="text-align:right;">
                    <span class="badge" id="{reg_id}-ach-badge" style="background:#10b98122; color:#10b981; font-size:12px; font-weight:700;">-</span>
                    <div style="font-size:11px; margin-top:4px;">YoY: <span id="{reg_id}-yoy-badge" style="font-weight:700;">-</span></div>
                </div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap:10px; margin-top:14px; background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
                <div>
                    <div style="font-size:11px; color:#94a3b8;">CURRENT SALES</div>
                    <div style="font-size:14px; font-weight:700; color:#fff;" id="{reg_id}-kpi-sales">-</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">QTY SOLD</div>
                    <div style="font-size:14px; font-weight:700; color:#38bdf8;" id="{reg_id}-kpi-units">-</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">TRANSACTIONS</div>
                    <div style="font-size:14px; font-weight:700; color:#fff;" id="{reg_id}-kpi-txns">-</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">UPT</div>
                    <div style="font-size:13px; font-weight:700; color:#10b981;" id="{reg_id}-kpi-upt">-</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ATV</div>
                    <div style="font-size:14px; font-weight:700; color:#fff;" id="{reg_id}-kpi-atv">-</div>
                </div>
                <div>
                    <div style="font-size:11px; color:#94a3b8;">ASP</div>
                    <div style="font-size:14px; font-weight:700; color:#f59e0b;" id="{reg_id}-kpi-asp">-</div>
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
            brand_pill = f'<span class="badge" style="background:{"#ef444422" if r["brand"]=="DZL" else "#38bdf822"}; color:{"#ef4444" if r["brand"]=="DZL" else "#38bdf8"}; margin-right:4px;">{r["brand"]}</span>'

            reg_rows += f"""
            <tr onclick="openStoreDetails('{r['clean_code']}')" class="clickable-row store-row" data-brand="{r['brand']}" title="Click to view detailed store breakdown">
                <td style="color:#64748b;">{idx+1}</td>
                <td style="color:#38bdf8;font-weight:600;">{r['clean_code']}</td>
                <td style="font-weight:600;color:#fff;">{brand_pill} {r['full_name']}</td>
                <td style="font-weight:700;color:#f8fafc;">{round(r['sales']):,}</td>
                <td style="color:#38bdf8;font-weight:600;">{ly_str}</td>
                <td>{yoy_badge_cell}</td>
                <td style="color:#94a3b8;">{t_str}</td>
                <td style="min-width:120px;">{ach_cell}</td>
                <td style="font-weight:700;color:#38bdf8;">{int(r['units']):,}</td>
                <td style="font-weight:700;color:#fff;">{int(r['txns']):,}</td>
                <td style="font-weight:700;color:#10b981;">{r['upt']:.2f}</td>
                <td style="font-weight:700;color:#38bdf8;">{r['atv']:,}</td>
                <td>{r['str_pct']}%</td>
                <td style="color:#f59e0b;font-weight:600;">{r['asp']:,}</td>
            </tr>
            """

        region_tables_html += f"""
        <div class="table-wrap region-table-wrap" id="reg-table-{reg_id}" data-region="{reg_name}" style="margin-bottom:30px;">
            <div class="table-header">
                <div>
                    <h3 style="color:#38bdf8; font-size:16px;">🏢 {reg_name.upper()}</h3>
                    <span style="color:var(--text-muted);font-size:12px;">Area Manager: <strong style="color:#fff;">{reg_mgr}</strong> | Stores Active: <strong id="{reg_id}-doors-count" style="color:#38bdf8;">{len(grp)} Doors</strong></span>
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
                            <th>ATV (SAR)</th>
                            <th>STR%</th>
                            <th>ASP</th>
                        </tr>
                    </thead>
                    <tbody>
                        {reg_rows}
                        <tr id="{reg_id}-tot-row" style="background:#0c1220; font-weight:800; border-top:2px solid #38bdf8; font-size:13px;">
                            <td colspan="3" style="color:#38bdf8;" id="{reg_id}-tot-label">TOTAL {reg_name.upper()} ({reg_mgr})</td>
                            <td style="color:#fff;" id="{reg_id}-tot-sales">-</td>
                            <td style="color:#38bdf8;" id="{reg_id}-tot-ly">-</td>
                            <td id="{reg_id}-tot-yoy">-</td>
                            <td style="color:#94a3b8;" id="{reg_id}-tot-target">-</td>
                            <td id="{reg_id}-tot-ach">-</td>
                            <td style="color:#38bdf8;" id="{reg_id}-tot-units">-</td>
                            <td style="color:#fff;" id="{reg_id}-tot-txns">-</td>
                            <td style="color:#10b981;" id="{reg_id}-tot-upt">-</td>
                            <td style="color:#38bdf8;" id="{reg_id}-tot-atv">-</td>
                            <td>-</td>
                            <td style="color:#f59e0b;" id="{reg_id}-tot-asp">-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        """

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

    # بناء قالب الـ HTML كاملاً ومضبوطاً لتفادي أي خطأ
    html_content = f"""<!DOCTYPE html>
<html lang="en" id="html-root">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MMS & DZL Executive Commercial Intelligence Dashboard</title>
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

        .logout-btn {{ background: #ef444422; border: 1px solid #ef444455; color: #ef4444; padding: 8px 14px; border-radius: 8px; font-weight: 700; cursor: pointer; transition: 0.2s; font-size: 12px; }}

        .view-toggle-bar {{ display: flex; background: #0c1220; padding: 4px; border-radius: 10px; border: 1px solid var(--border); margin-bottom: 24px; width: fit-content; gap: 4px; flex-wrap: wrap; }}
        .view-btn {{ background: transparent; border: none; color: var(--text-muted); padding: 10px 22px; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; transition: 0.2s; }}
        .view-btn.active {{ background: #2563eb; color: #fff; }}

        .section-title {{ font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted); margin-bottom: 14px; display: flex; align-items: center; justify-content:space-between; flex-wrap:wrap; gap:10px; }}
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
        table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }}
        th {{ background: #0c1220; color: var(--text-muted); padding: 12px 14px; font-weight: 600; text-transform: uppercase; font-size: 11px; border-bottom: 1px solid var(--border); }}
        td {{ padding: 12px 14px; border-bottom: 1px solid var(--border); }}
        tr:hover td {{ background: var(--card-hover); }}

        .clickable-row {{ cursor: pointer; transition: background 0.15s ease; }}
        .clickable-row:hover td {{ background: #1e293b !important; }}

        .badge {{ padding: 4px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }}
        .cards-scroll-container {{ display: flex; gap: 14px; overflow-x: auto; padding-bottom: 12px; margin-bottom: 24px; }}

        .sub-tab-btn {{ background:#1e293b; color:#94a3b8; border:1px solid #334155; padding:8px 16px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; }}
        .sub-tab-btn.active {{ background:#38bdf8; color:#090d16; border-color:#38bdf8; }}

        .export-btn {{ background: #10b981; border: none; color: #fff; padding: 8px 16px; border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 13px; }}
        .export-btn.dzl-btn {{ background: #ef4444; }}
        .sku-drawer {{ background: #090d16; display: none; }}

        .app-modal {{ 
            position: fixed !important; top: 0 !important; left: 0 !important; width: 100vw !important; height: 100vh !important; 
            background: rgba(9, 13, 22, 0.9) !important; backdrop-filter: blur(8px) !important; z-index: 2147483647 !important; 
            display: none; align-items: center; justify-content: center; 
        }}
        .modal-content {{ 
            background: #131b2e; border: 1px solid #1e293b; border-radius: 14px; width: 92%; max-width: 950px; max-height: 90vh; 
            display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.85); 
        }}
        .modal-header {{ padding: 18px 24px; border-bottom: 1px solid #1e293b; display: flex; justify-content: space-between; align-items: center; background: #0c1220; }}
        .modal-body {{ padding: 24px; overflow-y: auto; }}
        .close-btn {{ background: transparent; border: none; color: #94a3b8; font-size: 28px; cursor: pointer; line-height: 1; }}
        .close-btn:hover {{ color: #fff; }}
    </style>
</head>
<body>

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

<!-- Store Modal for Detailed Breakdown -->
<div id="store-modal" class="app-modal">
  <div class="modal-content">
    <div class="modal-header">
      <div>
        <h2 id="modal-store-name" style="margin:0; font-size:18px; color:#fff;">Store Performance Analysis</h2>
        <span id="modal-store-code" style="color:#38bdf8; font-size:12px; font-weight:700;">CODE</span>
      </div>
      <button class="close-btn" onclick="closeModal('store-modal')">&times;</button>
    </div>
    <div class="modal-body">
      <div style="background:#090d16; border:1px solid #1e293b; border-radius:10px; padding:16px; margin-bottom:18px;">
        <div style="font-size:13px; color:#f59e0b; margin-bottom:6px;"><strong>🎯 Focus Needs:</strong> <span id="modal-needs" style="color:#fff;">-</span></div>
        <div style="font-size:13px; color:#38bdf8;"><strong>⚡ Operational Directive:</strong> <span id="modal-directive" style="color:#fff;">-</span></div>
      </div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(130px, 1fr)); gap:10px; margin-bottom:20px;">
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">CURRENT SALES</div>
          <div id="modal-sales" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">TARGET (% ACH)</div>
          <div id="modal-target" style="font-size:16px; font-weight:700; color:#10b981;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">WOC COVER</div>
          <div id="modal-woc" style="font-size:16px; font-weight:700; color:#f59e0b;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">SOH STOCK</div>
          <div id="modal-soh" style="font-size:16px; font-weight:700; color:#38bdf8;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">ATV</div>
          <div id="modal-atv" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
        <div style="background:#090d16; padding:12px; border-radius:8px; border:1px solid #1e293b;">
          <div style="font-size:10px; color:#94a3b8;">UPT</div>
          <div id="modal-upt" style="font-size:16px; font-weight:700; color:#fff;">-</div>
        </div>
      </div>
      <div style="font-size:12px; font-weight:700; text-transform:uppercase; color:#94a3b8; margin-bottom:8px;">Category Contribution for this Door</div>
      <table style="width:100%; border:1px solid #1e293b; border-radius:6px; overflow:hidden;">
        <thead>
          <tr>
            <th>Category Name</th>
            <th>Sales (SAR)</th>
            <th>Units</th>
            <th>Mix %</th>
            <th>ASP</th>
          </tr>
        </thead>
        <tbody id="modal-cats-body"></tbody>
      </table>
    </div>
  </div>
</div>

<div class="header">
    <div>
        <h1>MMS & DZL Executive Commercial Intelligence Dashboard</h1>
        <p>Portfolio Performance, Product Groups & SKC Size-Run Analytics (GWP & Non-Commercial Stripped)</p>
    </div>
    <div class="top-controls">
        <div class="brand-switcher">
            <button class="brand-btn active" id="btn-brand-ALL" onclick="switchBrand('ALL')">🏢 ALL BRANDS</button>
            <button class="brand-btn" id="btn-brand-MMS" onclick="switchBrand('MMS')">🔴 MUMUSO (17)</button>
            <button class="brand-btn" id="btn-brand-DZL" onclick="switchBrand('DZL')">🟡 DZL (3)</button>
        </div>
        <span id="current-user-badge" style="font-size:13px; font-weight:700; color:#38bdf8; background:#1e293b; padding:8px 14px; border-radius:8px; border:1px solid #334155;">👤 Authenticating..</span>
        <button class="logout-btn" onclick="logout()">Logout</button>
    </div>
</div>

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
        <div class="kpi-value" id="kpi-yoy-growth" style="color:{net_yoy_col}; font-weight:800;">{network_lfl_growth:+.1f}%</div>
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
    <button class="view-btn" id="btn-business" onclick="switchView('business')">📦 Business-Wise & Gender Mix</button>
    <button class="view-btn" id="btn-action" onclick="switchView('action')" style="border-left:2px solid #38bdf8;">⚡ Commercial Action Hub (Top/Low 20 Shoes & Replenishment)</button>
</div>

<!-- 1. Store Commercial Matrix View -->
<div id="view-stores">
    <div class="section-title"><span>⚡ Critical Action Directives (Store Diagnostics & SOH Coverage)</span></div>
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
                <span style="color:var(--text-muted);font-size:12px;">Active Stores Performance Summary (Click any row to view store details)</span>
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
                        <th>ATV (SAR)</th>
                        <th>STR%</th>
                        <th>Commercial Diagnostic</th>
                        <th>ASP</th>
                    </tr>
                </thead>
                <tbody>
                    {store_table_rows}
                    <tr id="storesTableTotalRow" style="background:#0c1220; font-weight:800; border-top:3px solid #38bdf8; font-size:13px;">
                        <td colspan="4" style="color:#38bdf8; text-transform:uppercase;" id="storesTotalTitle">TOTAL PORTFOLIO (ALL DOORS)</td>
                        <td style="color:#fff;" id="tot-sales">{total_sales:,}</td>
                        <td style="color:#38bdf8;" id="tot-ly">{total_ly_sales:,}</td>
                        <td id="tot-yoy" style="color:{net_yoy_col}; font-weight:800;">{network_lfl_growth:+.1f}%</td>
                        <td style="color:#94a3b8;" id="tot-target">{total_target:,}</td>
                        <td id="tot-ach" style="color:{'#10b981' if overall_ach>=100 else '#f59e0b'};">{overall_ach:.1f}%</td>
                        <td style="color:#38bdf8;" id="tot-units">{total_units:,}</td>
                        <td style="color:#fff;" id="tot-txns">{total_txns:,}</td>
                        <td style="color:#10b981;" id="tot-upt">{network_upt:.2f}</td>
                        <td style="color:#38bdf8;" id="tot-atv">{network_atv:,}</td>
                        <td id="tot-str">-</td>
                        <td style="color:#64748b;">-</td>
                        <td style="color:#f59e0b;" id="tot-asp">{network_asp:,}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- 2. Region-Wise Performance View (Fully Restored & Dynamic per Brand) -->
<div id="view-regions" style="display:none;">
    <div class="section-title"><span>🌍 REGIONAL LEADERSHIP & AREA MANAGER OVERVIEW</span></div>
    <div style="display:flex; flex-wrap:wrap; gap:16px; margin-bottom:24px;">
        {region_kpi_cards}
    </div>

    {region_tables_html}
</div>

<!-- 3. Business-Wise View (2 Charts: Category Contribution + Gender-Wise Contribution + Drill-down) -->
<div id="view-business" style="display:none;">
    <div class="section-title">
        <span id="businessViewTitle">🏷️ CATEGORY PORTFOLIO CONTRIBUTION</span>
        <span style="font-size:12px; color:var(--text-muted); font-weight:400;">100% Commercial Assortment (Click any row/chart to drill down)</span>
    </div>
    
    <div class="cards-scroll-container" id="categoryCardsContainer"></div>

    <div style="display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px;">
        <div class="chart-container" style="flex:1; min-width:320px;">
            <div class="section-title">
                <span id="chart1Title">🍩 Category Revenue Contribution Share</span>
            </div>
            <div id="apexCategoryDonut" style="min-height: 330px;"></div>
        </div>

        <div class="chart-container" id="genderChartWrapper" style="flex:1; min-width:320px;">
            <div class="section-title">
                <span>👟 Footwear Gender Mix Share (Kids: 23-34 | Women: 35-39 | Men: 40-48)</span>
            </div>
            <div id="apexGenderDonut" style="min-height: 330px;"></div>
        </div>
    </div>

    <div style="display:flex; gap:16px; flex-wrap:wrap; margin-bottom:24px;">
        <div class="table-wrap" style="flex:1; min-width:340px;">
            <div class="table-header">
                <h3 id="tableHierarchyTitle">CATEGORY CONTRIBUTION BREAKDOWN</h3>
                <span style="font-size:11px; color:#38bdf8;">(Click to filter items below)</span>
            </div>
            <div style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Category Name</th>
                            <th>Sales Revenue (SAR)</th>
                            <th>Units Sold</th>
                            <th>Mix Share (%)</th>
                            <th>ASP (SAR)</th>
                        </tr>
                    </thead>
                    <tbody id="hierarchyTableBody"></tbody>
                </table>
            </div>
        </div>

        <div class="table-wrap" id="genderTableWrap" style="flex:1; min-width:340px;">
            <div class="table-header">
                <h3>GENDER-WISE FOOTWEAR CONTRIBUTION</h3>
                <span style="font-size:11px; color:#ec4899;">(Click to filter items below)</span>
            </div>
            <div style="overflow-x:auto;">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Gender Segment</th>
                            <th>Sales Revenue (SAR)</th>
                            <th>Units Sold</th>
                            <th>Contribution (%)</th>
                            <th>ASP (SAR)</th>
                        </tr>
                    </thead>
                    <tbody id="genderTableBody"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Drill-Down Table in Business-Wise -->
    <div class="table-wrap" id="businessDrillDownSection" style="margin-top:24px;">
        <div class="table-header">
            <div>
                <h3 style="color:#38bdf8;" id="drilldownTitle">🔍 DRILL-DOWN ITEM & STYLE MATRIX: ALL ITEMS</h3>
                <span style="font-size:12px; color:var(--text-muted);" id="drilldownSubtitle">Detailed style-level & SKU breakdown</span>
            </div>
            <input type="text" id="drilldownSearch" class="table-search" placeholder="Filter by Style, Name or Barcode..." onkeyup="filterDrilldownTable()">
        </div>
        <div style="max-height: 480px; overflow-y: auto; overflow-x: auto;">
            <table>
                <thead style="position: sticky; top: 0; z-index: 10;">
                    <tr>
                        <th>#</th>
                        <th>Product Group / Style</th>
                        <th>SKU / Barcode</th>
                        <th>Item Description</th>
                        <th>Category</th>
                        <th>Gender</th>
                        <th>Units Sold</th>
                        <th>Sales Revenue (SAR)</th>
                        <th>ASP (SAR)</th>
                        <th>KSWH SOH</th>
                    </tr>
                </thead>
                <tbody id="drilldownTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<!-- 4. Commercial Action Hub (Top 20 Shoes & Low 20 Shoes Group-Wise with SKU Drill-down) -->
<div id="view-action" style="display:none;">
    <div class="section-title" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
        <span>⚡ PREDICTIVE AUTO-REPLENISHMENT & STOCK-OUT FORECAST (KSWH & IST)</span>
        <div style="display:flex; gap:10px; flex-wrap:wrap;">
            <a href="./reports/Auto_Replenishment_MMS.xlsx" download class="export-btn" style="text-decoration:none; display:inline-flex; align-items:center; gap:6px;">📥 Download MMS Replenishment Plan</a>
            <a href="./reports/Auto_Replenishment_DZL.xlsx" download class="export-btn dzl-btn" style="text-decoration:none; display:inline-flex; align-items:center; gap:6px;">📥 Download DZL Replenishment Plan</a>
        </div>
    </div>

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
                <span style="color:var(--text-muted); font-size:12px;">Transfers are strictly isolated per brand network</span>
            </div>
            <input type="text" id="replSearch" class="table-search" placeholder="Search Barcode, SKU, Store..." onkeyup="filterReplTable()">
        </div>
        <div style="max-height: 440px; overflow-y: auto; overflow-x: auto;">
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
            <button class="sub-tab-btn active" id="btn-top-shoes" onclick="switchDZLMovers('top')">🔥 TOP 20 SHOES (GROUP-WISE / PRODUCT GROUP)</button>
            <button class="sub-tab-btn" id="btn-low-shoes" onclick="switchDZLMovers('low')">❄️ LOW 20 SHOES (SLOW MOVERS GROUP-WISE)</button>
        </div>
        <div>
            <span style="color:#38bdf8; font-size:12px; font-weight:700;">💡 Click any Product Group row to open SKU / Barcode & Size-Run details</span>
        </div>
    </div>

    <div class="table-wrap">
        <div style="overflow-x:auto;">
            <table id="shoesMoversTable">
                <thead>
                    <tr>
                        <th>Rank</th>
                        <th>Product Group (Shoe Style)</th>
                        <th>Gender</th>
                        <th>Units Sold</th>
                        <th>Sales Revenue (SAR)</th>
                        <th>Mix %</th>
                        <th>ASP (SAR)</th>
                        <th>Stores SOH</th>
                        <th>Central WH SOH</th>
                        <th>SKUs / Sizes</th>
                        <th>Drill-Down</th>
                    </tr>
                </thead>
                <tbody id="shoesMoversTableBody"></tbody>
            </table>
        </div>
    </div>
</div>

<script>
  const BRAND_KPIS = {json.dumps(brand_kpi_summary)};
  const STORE_CHART_DATA = {json.dumps(store_chart_data)};
  const REGION_KPIS = {json.dumps(region_kpis_by_brand)};
  const STORE_META = {json.dumps(store_meta_map)};
  const STORE_DETAILS = {json.dumps(store_category_details)};
  const CATEGORY_DATA_BY_BRAND = {json.dumps(category_data_by_brand)};
  const DZL_GENDER_DATA = {json.dumps(dzl_gender_data)};
  const DZL_TOP_20 = {json.dumps(dzl_top20_groups)};
  const DZL_LOW_20 = {json.dumps(dzl_low20_groups)};
  const BUSINESS_DRILLDOWN = {json.dumps(business_drilldown_data)};

  let currentActiveBrand = 'ALL';
  let currentReplBrand = 'ALL';
  let currentDZLMoversType = 'top';
  let currentDrillFilter = {{ type: 'ALL', value: 'ALL' }};
  let storeChartInstance = null;
  let donutChart = null;
  let genderChart = null;

  document.addEventListener("DOMContentLoaded", function() {{
    initApexCharts();
    switchBrand('ALL');
  }});

  function initApexCharts() {{
    const d = STORE_CHART_DATA['ALL'];
    var storeOptions = {{
      series: [
        {{ name: 'Actual Sales (SAR)', data: d.sales }},
        {{ name: 'Target (SAR)', data: d.targets }}
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
        categories: d.categories,
        labels: {{ style: {{ colors: '#94a3b8', fontSize: '11px' }} }}
      }},
      yaxis: {{
        labels: {{
          formatter: function (val) {{ return Number(val).toLocaleString() + ' SAR'; }},
          style: {{ colors: '#94a3b8', fontSize: '11px' }}
        }}
      }},
      grid: {{ borderColor: '#1e293b' }}
    }};

    var storeChartEl = document.querySelector("#apexStoreChart");
    if (storeChartEl) {{
      storeChartInstance = new ApexCharts(storeChartEl, storeOptions);
      storeChartInstance.render();
    }}
  }}

  function openStoreDetails(storeCode) {{
    const meta = STORE_META[storeCode];
    if (!meta) return;

    document.getElementById("modal-store-name").innerText = "[" + meta.brand + "] " + meta.name;
    document.getElementById("modal-store-code").innerText = "CODE: " + storeCode + " | " + meta.region + " (Manager: " + meta.manager + ")";
    document.getElementById("modal-needs").innerText = meta.needs;
    document.getElementById("modal-directive").innerText = meta.action;
    document.getElementById("modal-sales").innerText = meta.sales;
    document.getElementById("modal-target").innerText = meta.target + " (" + meta.ach + ")";
    document.getElementById("modal-woc").innerText = meta.woc;
    document.getElementById("modal-soh").innerText = meta.soh_units;
    document.getElementById("modal-atv").innerText = meta.atv;
    document.getElementById("modal-upt").innerText = meta.upt;

    const cats = STORE_DETAILS[storeCode] || [];
    let html = "";
    cats.forEach(c => {{
      html += `
        <tr>
          <td style="color:#38bdf8; font-weight:700;">${{c.main_category}}</td>
          <td style="color:#fff; font-weight:700;">${{c.sales}}</td>
          <td>${{c.units}}</td>
          <td style="color:#10b981; font-weight:700;">${{c.store_mix_pct}}</td>
          <td style="color:#f59e0b;">${{c.asp}}</td>
        </tr>
      `;
    }});
    document.getElementById("modal-cats-body").innerHTML = html || "<tr><td colspan='5' style='text-align:center;'>No category data</td></tr>";
    document.getElementById("store-modal").style.display = "flex";
  }}

  function closeModal(id) {{
    document.getElementById(id).style.display = "none";
  }}

  function updateRegionMetrics(brand) {{
    const rData = REGION_KPIS[brand];
    if (!rData) return;

    ["riyadh", "western"].forEach(regId => {{
      const regName = regId === "riyadh" ? "Riyadh Central Region" : "Western Region";
      const d = rData[regName];
      if (!d) return;

      document.getElementById(regId + "-kpi-sales").innerText = d.sales + " SAR";
      document.getElementById(regId + "-kpi-units").innerText = d.units;
      document.getElementById(regId + "-kpi-txns").innerText = d.txns;
      document.getElementById(regId + "-kpi-upt").innerText = d.upt;
      document.getElementById(regId + "-kpi-atv").innerText = d.atv + " SAR";
      document.getElementById(regId + "-kpi-asp").innerText = d.asp + " SAR";
      
      const achBadge = document.getElementById(regId + "-ach-badge");
      achBadge.innerText = d.ach + " Ach";
      achBadge.style.color = (parseFloat(d.ach) >= 100) ? "#10b981" : ((parseFloat(d.ach) >= 80) ? "#f59e0b" : "#ef4444");

      const yoyBadge = document.getElementById(regId + "-yoy-badge");
      yoyBadge.innerText = d.yoy;
      yoyBadge.style.color = (d.yoy_val >= 0) ? "#10b981" : "#ef4444";

      document.getElementById(regId + "-doors-count").innerText = d.stores_count + " Doors";

      document.getElementById(regId + "-tot-sales").innerText = d.sales;
      document.getElementById(regId + "-tot-ly").innerText = d.ly;
      const rYoy = document.getElementById(regId + "-tot-yoy");
      rYoy.innerText = d.yoy;
      rYoy.style.color = (d.yoy_val >= 0) ? "#10b981" : "#ef4444";

      document.getElementById(regId + "-tot-target").innerText = d.target;
      document.getElementById(regId + "-tot-ach").innerText = d.ach;
      document.getElementById(regId + "-tot-units").innerText = d.units;
      document.getElementById(regId + "-tot-txns").innerText = d.txns;
      document.getElementById(regId + "-tot-upt").innerText = d.upt;
      document.getElementById(regId + "-tot-atv").innerText = d.atv;
      document.getElementById(regId + "-tot-asp").innerText = d.asp;

      const pfx = brand === "ALL" ? "TOTAL " + regName.toUpperCase() : (brand === "MMS" ? "TOTAL MMS " + regName.toUpperCase() : "TOTAL DZL " + regName.toUpperCase());
      document.getElementById(regId + "-tot-label").innerText = pfx;
    }});
  }}

  function renderCategorySection(brand) {{
    const cats = CATEGORY_DATA_BY_BRAND[brand] || [];
    const colors = ['#38bdf8', '#f59e0b', '#10b981', '#ec4899', '#818cf8', '#a855f7', '#06b6d4', '#e11d48'];

    let cardsHtml = "";
    cats.forEach((c, idx) => {{
      const clr = colors[idx % colors.length];
      const safeCat = html.escape(c.main_category).replace("'", "\\'");
      cardsHtml += `
        <div onclick="setBusinessDrillDown('category', '${{safeCat}}')" class="clickable-row" style="background:var(--card); border:1px solid var(--border); border-top:3px solid ${{clr}}; border-radius:10px; padding:16px; min-width:210px; max-width:260px; flex:1;" title="Click to drill down into ${{c.main_category}} items">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
            <span style="font-size:13px; font-weight:700; color:#fff;">${{c.main_category}}</span>
            <span style="font-size:12px; font-weight:700; color:${{clr}};">${{c.contribution}}%</span>
          </div>
          <div style="font-size:18px; font-weight:700; color:#f8fafc; margin-bottom:4px;">${{Math.round(c.sales).toLocaleString()}} <span style="font-size:11px; color:#94a3b8;">SAR</span></div>
          <div style="font-size:11px; color:#94a3b8; margin-bottom:8px;">Units: <strong style="color:#fff;">${{Number(c.units).toLocaleString()}}</strong> | ASP: <strong style="color:#f59e0b;">${{c.asp}}</strong></div>
          <div style="background:#090d16; border-radius:4px; height:5px; overflow:hidden;">
            <div style="background:${{clr}}; width:${{Math.min(c.contribution, 100)}}%; height:100%;"></div>
          </div>
        </div>
      `;
    }});
    document.getElementById("categoryCardsContainer").innerHTML = cardsHtml;

    let tableHtml = "";
    cats.forEach((c, idx) => {{
      const safeCat = html.escape(c.main_category).replace("'", "\\'");
      tableHtml += `
        <tr onclick="setBusinessDrillDown('category', '${{safeCat}}')" class="clickable-row" title="Click to drill down into ${{c.main_category}} items">
          <td style="color:#64748b; font-weight:600;">${{idx+1}}</td>
          <td style="font-weight:700; color:#fff;">🏷️ ${{c.main_category}} <span style="font-size:11px; color:#38bdf8;">(Drill-down ▼)</span></td>
          <td style="font-weight:700; color:#38bdf8;">${{Math.round(c.sales).toLocaleString()}}</td>
          <td>${{Number(c.units).toLocaleString()}}</td>
          <td style="font-weight:700; color:#10b981;">${{c.contribution}}%</td>
          <td style="color:#f59e0b; font-weight:700;">${{c.asp}}</td>
        </tr>
      `;
    }});
    document.getElementById("hierarchyTableBody").innerHTML = tableHtml;

    let gTableHtml = "";
    (DZL_GENDER_DATA.records || []).forEach((g, idx) => {{
      const safeG = html.escape(g.gender).replace("'", "\\'");
      gTableHtml += `
        <tr onclick="setBusinessDrillDown('gender', '${{safeG}}')" class="clickable-row" title="Click to drill down into ${{g.gender}} items">
          <td style="color:#64748b; font-weight:600;">${{idx+1}}</td>
          <td style="font-weight:700; color:#fff;">👟 ${{g.gender}} <span style="font-size:11px; color:#ec4899;">(Drill-down ▼)</span></td>
          <td style="font-weight:700; color:#38bdf8;">${{Math.round(g.sales).toLocaleString()}}</td>
          <td>${{Number(g.units).toLocaleString()}}</td>
          <td style="font-weight:700; color:#ec4899;">${{g.share}}%</td>
          <td style="color:#f59e0b; font-weight:700;">${{g.asp}}</td>
        </tr>
      `;
    }});
    document.getElementById("genderTableBody").innerHTML = gTableHtml;

    const series = cats.map(c => c.contribution);
    const labels = cats.map(c => c.main_category);

    if (donutChart) donutChart.destroy();
    donutChart = new ApexCharts(document.querySelector("#apexCategoryDonut"), {{
      series: series,
      labels: labels,
      chart: {{ type: 'donut', height: 330, background: 'transparent' }},
      theme: {{ mode: 'dark' }},
      colors: colors,
      legend: {{ position: 'bottom', labels: {{ colors: '#cbd5e1' }} }}
    }});
    donutChart.render();

    const genderWrapper = document.getElementById("genderChartWrapper");
    const genderTableWrap = document.getElementById("genderTableWrap");
    if (brand === 'DZL' || brand === 'ALL') {{
      genderWrapper.style.display = "block";
      genderTableWrap.style.display = "block";
      if (genderChart) genderChart.destroy();
      genderChart = new ApexCharts(document.querySelector("#apexGenderDonut"), {{
        series: DZL_GENDER_DATA.series,
        labels: DZL_GENDER_DATA.labels,
        chart: {{ type: 'donut', height: 330, background: 'transparent' }},
        theme: {{ mode: 'dark' }},
        colors: ['#ec4899', '#38bdf8', '#10b981'],
        legend: {{ position: 'bottom', labels: {{ colors: '#cbd5e1' }} }}
      }});
      genderChart.render();
    }} else {{
      genderWrapper.style.display = "none";
      genderTableWrap.style.display = "none";
    }}

    renderBusinessDrillDown();
  }}

  function setBusinessDrillDown(type, val) {{
    currentDrillFilter = {{ type: type, value: val }};
    renderBusinessDrillDown();
    const target = document.getElementById("businessDrillDownSection");
    if (target) target.scrollIntoView({{ behavior: 'smooth' }});
  }}

  function renderBusinessDrillDown() {{
    const tbody = document.getElementById("drilldownTableBody");
    const searchVal = document.getElementById("drilldownSearch").value.toLowerCase();
    
    let filtered = BUSINESS_DRILLDOWN.filter(item => {{
      const matchBrand = (currentActiveBrand === "ALL" || item.brand === currentActiveBrand);
      let matchFilter = true;
      if (currentDrillFilter.type === 'category' && currentDrillFilter.value !== 'ALL') {{
        matchFilter = (item.main_category === currentDrillFilter.value);
      }} else if (currentDrillFilter.type === 'gender' && currentDrillFilter.value !== 'ALL') {{
        matchFilter = (item.gender === currentDrillFilter.value && item.main_category === 'Shoes');
      }}
      const matchSearch = (
        item.style_group.toLowerCase().includes(searchVal) ||
        item.clean_barcode.toLowerCase().includes(searchVal) ||
        item.clean_item_name.toLowerCase().includes(searchVal)
      );
      return matchBrand && matchFilter && matchSearch;
    }});

    const titleEl = document.getElementById("drilldownTitle");
    const subEl = document.getElementById("drilldownSubtitle");
    if (currentDrillFilter.type === 'category') {{
      titleEl.innerText = `🔍 DRILL-DOWN ITEM & STYLE MATRIX: CATEGORY (${{currentDrillFilter.value.toUpperCase()}})`;
      subEl.innerText = `Showing all items under ${{currentDrillFilter.value}} (${{filtered.length}} items)`;
    }} else if (currentDrillFilter.type === 'gender') {{
      titleEl.innerText = `🔍 DRILL-DOWN ITEM & STYLE MATRIX: GENDER (${{currentDrillFilter.value.toUpperCase()}} FOOTWEAR)`;
      subEl.innerText = `Showing footwear items under size run for ${{currentDrillFilter.value}} (${{filtered.length}} items)`;
    }} else {{
      titleEl.innerText = "🔍 DRILL-DOWN ITEM & STYLE MATRIX: ALL ITEMS";
      subEl.innerText = "Click on any category or gender above to narrow drill-down";
    }}

    let html = "";
    filtered.slice(0, 100).forEach((r, idx) => {{
      const bColor = r.brand === 'DZL' ? '#ef4444' : '#38bdf8';
      html += `
        <tr>
          <td style="color:#64748b; font-weight:600;">${{idx+1}}</td>
          <td style="font-weight:700; color:#fff;">${{r.style_group}}</td>
          <td style="color:#38bdf8; font-weight:600;"><span class="badge" style="background:${{bColor}}22; color:${{bColor}}; margin-right:4px;">${{r.brand}}</span>${{r.clean_barcode}}</td>
          <td style="color:#fff;">${{r.clean_item_name}}</td>
          <td style="color:#94a3b8;">${{r.main_category}}</td>
          <td><span class="badge" style="background:#ec489922; color:#ec4899;">${{r.gender}}</span></td>
          <td style="font-weight:700; color:#fff;">${{r.units}} Pcs</td>
          <td style="font-weight:700; color:#38bdf8;">${{Math.round(r.sales).toLocaleString()}}</td>
          <td style="color:#f59e0b; font-weight:700;">${{r.asp}}</td>
          <td style="color:#10b981; font-weight:700;">${{r.wh_soh}} Pcs</td>
        </tr>
      `;
    }});
    tbody.innerHTML = html || "<tr><td colspan='10' style='text-align:center;'>No matching items found for drill-down</td></tr>";
  }}

  function filterDrilldownTable() {{
    renderBusinessDrillDown();
  }}

  function switchDZLMovers(type) {{
    currentDZLMoversType = type;
    document.getElementById("btn-top-shoes").classList.toggle("active", type === 'top');
    document.getElementById("btn-low-shoes").classList.toggle("active", type === 'low');
    renderDZLShoesTable();
  }}

  function renderDZLShoesTable() {{
    const data = (currentDZLMoversType === 'top') ? DZL_TOP_20 : DZL_LOW_20;
    const tbody = document.getElementById("shoesMoversTableBody");
    let html = "";

    data.forEach((r, idx) => {{
      const rankColor = currentDZLMoversType === 'top' ? '#10b981' : '#ef4444';
      const rowId = "drawer-" + idx;
      
      let skuRows = "";
      r.skus.forEach(s => {{
        skuRows += `
          <tr style="border-bottom:1px solid #1e293b;">
            <td style="color:#38bdf8; font-weight:700;">${{s.barcode}}</td>
            <td style="color:#94a3b8;">${{s.item_code}}</td>
            <td style="color:#fff; font-weight:600;">${{s.name}}</td>
            <td style="font-weight:700; color:#fff;">${{s.units}} Pcs</td>
            <td style="font-weight:700; color:#38bdf8;">${{s.sales.toLocaleString()}} SAR</td>
            <td style="font-weight:700; color:#f59e0b;">${{s.store_soh}} Pcs</td>
            <td style="font-weight:700; color:#10b981;">${{s.wh_soh}} Pcs</td>
          </tr>
        `;
      }});

      html += `
        <tr class="clickable-row" onclick="toggleDrawer('${{rowId}}')">
          <td style="color:${{rankColor}}; font-weight:800;">#${{r.rank}}</td>
          <td style="font-weight:700; color:#fff; font-size:14px;">👟 ${{r.style_group}}</td>
          <td><span class="badge" style="background:#ec489922; color:#ec4899;">${{r.gender}}</span></td>
          <td style="font-weight:700; color:#38bdf8;">${{r.units.toLocaleString()}}</td>
          <td style="font-weight:700; color:#fff;">${{r.sales.toLocaleString()}}</td>
          <td style="font-weight:700; color:#10b981;">${{r.share_pct}}%</td>
          <td style="color:#f59e0b; font-weight:700;">${{r.asp}}</td>
          <td style="color:#f59e0b; font-weight:700;">${{r.store_soh.toLocaleString()}} Pcs</td>
          <td style="color:#10b981; font-weight:700;">${{r.wh_soh.toLocaleString()}} Pcs</td>
          <td><span class="badge" style="background:#2563eb22; color:#38bdf8;">${{r.skus_count}} SKUs</span></td>
          <td><button style="background:#2563eb; color:#fff; border:none; border-radius:4px; padding:5px 10px; font-size:11px; font-weight:700; cursor:pointer;">View SKUs ▼</button></td>
        </tr>
        <tr id="${{rowId}}" class="sku-drawer">
          <td colspan="11" style="padding:14px 24px; background:#090d16;">
            <div style="font-size:12px; font-weight:700; color:#38bdf8; margin-bottom:8px;">📦 SKU / BARCODE LEVEL BREAKDOWN FOR PRODUCT GROUP: ${{r.style_group}}</div>
            <table style="width:100%; background:#0c1220; border:1px solid #1e293b; border-radius:6px; overflow:hidden;">
              <thead>
                <tr>
                  <th>Barcode (SKC)</th>
                  <th>Item Code</th>
                  <th>SKU / Size & Color Name</th>
                  <th>Units Sold</th>
                  <th>Sales Revenue</th>
                  <th>Stores SOH</th>
                  <th>Warehouse SOH (KSWH)</th>
                </tr>
              </thead>
              <tbody>
                ${{skuRows}}
              </tbody>
            </table>
          </td>
        </tr>
      `;
    }});

    tbody.innerHTML = html || "<tr><td colspan='11' style='text-align:center;'>No footwear data available</td></tr>";
  }}

  function toggleDrawer(id) {{
    const el = document.getElementById(id);
    if (el) el.style.display = (el.style.display === "table-row") ? "none" : "table-row";
  }}

  function switchBrand(brand) {{
    currentActiveBrand = brand;
    document.querySelectorAll(".brand-btn").forEach(btn => btn.classList.remove("active"));
    const activeBtn = document.getElementById("btn-brand-" + brand);
    if (activeBtn) activeBtn.classList.add("active");

    if (storeChartInstance && STORE_CHART_DATA[brand]) {{
      const d = STORE_CHART_DATA[brand];
      storeChartInstance.updateOptions({{
        xaxis: {{ categories: d.categories }}
      }});
      storeChartInstance.updateSeries([
        {{ name: 'Actual Sales (SAR)', data: d.sales }},
        {{ name: 'Target (SAR)', data: d.targets }}
      ]);
    }}

    if (BRAND_KPIS[brand]) {{
      const k = BRAND_KPIS[brand];
      document.getElementById("kpi-total-sales").innerHTML = k.sales + " <span class='kpi-unit'>SAR</span>";
      document.getElementById("kpi-total-target").innerText = k.target + " SAR";
      document.getElementById("kpi-ly-sales").innerText = k.ly + " SAR";
      document.getElementById("kpi-overall-ach").innerText = k.ach;

      const yoyEl = document.getElementById("kpi-yoy-growth");
      yoyEl.innerText = k.yoy;
      yoyEl.style.setProperty("color", k.yoy_val < 0 ? "#ef4444" : "#10b981", "important");

      document.getElementById("kpi-network-atv").innerText = "SAR " + k.atv;
      document.getElementById("kpi-network-asp").innerText = "SAR " + k.asp;

      const titlePrefix = brand === "ALL" ? "TOTAL PORTFOLIO (ALL DOORS)" : (brand === "MMS" ? "TOTAL MUMUSO NETWORK (17 DOORS)" : "TOTAL DZL DOZOLO (3 DOORS)");
      document.getElementById("storesTotalTitle").innerText = titlePrefix;
      document.getElementById("tot-sales").innerText = k.sales;
      document.getElementById("tot-ly").innerText = k.ly;
      document.getElementById("tot-target").innerText = k.target;
      document.getElementById("tot-ach").innerText = k.ach;
      document.getElementById("tot-units").innerText = k.units;
      document.getElementById("tot-txns").innerText = k.txns;
      document.getElementById("tot-upt").innerText = k.upt;
      document.getElementById("tot-atv").innerText = k.atv;
      document.getElementById("tot-str").innerText = k.str;
      document.getElementById("tot-asp").innerText = k.asp;

      const totYoyEl = document.getElementById("tot-yoy");
      if (totYoyEl) {{
        totYoyEl.innerText = k.yoy;
        totYoyEl.style.setProperty("color", k.yoy_val < 0 ? "#ef4444" : "#10b981", "important");
      }}
    }}

    document.querySelectorAll(".store-row").forEach(row => {{
      const rBrand = row.getAttribute("data-brand");
      row.style.display = (brand === "ALL" || rBrand === brand) ? "" : "none";
    }});

    document.querySelectorAll(".store-card-item").forEach(card => {{
      const cBrand = card.getAttribute("data-brand");
      card.style.display = (brand === "ALL" || cBrand === brand) ? "" : "none";
    }});

    updateRegionMetrics(brand);
    renderCategorySection(brand);
    filterReplByBrand(brand);
    renderDZLShoesTable();
  }}

  function filterReplByBrand(brand) {{
    currentReplBrand = brand;
    document.getElementById("repl-btn-all").classList.toggle("active", brand === "ALL");
    document.getElementById("repl-btn-mms").classList.toggle("active", brand === "MMS");
    document.getElementById("repl-btn-dzl").classList.toggle("active", brand === "DZL");

    const titleEl = document.getElementById("replTableHeaderTitle");
    if (titleEl) {{
        if (brand === "MMS") titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (MUMUSO ONLY)";
        else if (brand === "DZL") titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (DZL DOZOLO ONLY - SHOES PRIORITY 1)";
        else titleEl.innerText = "⚡ ACTIONABLE REPLENISHMENT DIRECTIVES (MULTI-STORE BALANCING)";
    }}
    filterReplTable();
  }}

  function filterReplTable() {{
    var searchVal = document.getElementById("replSearch").value.toLowerCase();
    var rows = document.querySelectorAll("#replTableBody tr");

    rows.forEach(function(row) {{
      var text = row.innerText.toLowerCase();
      var repBrand = row.getAttribute("data-brand");
      var matchBrand = (currentReplBrand === "ALL" || repBrand === currentReplBrand);
      var matchSearch = text.includes(searchVal);
      row.style.display = (matchBrand && matchSearch) ? "" : "none";
    }});
  }}

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
        updateRegionMetrics(currentActiveBrand);
    }} else if (viewName === 'business' && businessView) {{
        businessView.style.display = "block";
        if (btnBusiness) btnBusiness.classList.add("active");
        renderCategorySection(currentActiveBrand);
    }} else if (actionView) {{
        actionView.style.display = "block";
        if (btnAction) btnAction.classList.add("active");
        renderDZLShoesTable();
    }}
  }}

  function filterStores() {{
      const query = document.getElementById("storeSearch").value.toLowerCase();
      const rows = document.querySelectorAll("#storesTable tbody tr.store-row");
      rows.forEach(r => {{
          const title = r.innerText.toLowerCase();
          const rBrand = r.getAttribute("data-brand");
          const matchBrand = (currentActiveBrand === "ALL" || rBrand === currentActiveBrand);
          r.style.display = (matchBrand && title.includes(query)) ? "" : "none";
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