import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="訂單自動分類整理工具", page_icon="📋", layout="wide")

st.title("📋 訂單自動分類與欄位整理工具")
st.markdown("上傳原始訂單 Excel/CSV 後，系統會自動按物流方式分類並整理成專屬欄位！")

# 台灣地址拆分與清理輔助函式
def parse_taiwan_address(address):
    if not isinstance(address, str) or not address:
        return "", "", "", ""
    
    address = address.strip()
    address = re.sub(r'[\,\s]*台灣[\,\s]*$', '', address, flags=re.IGNORECASE)
    address = re.sub(r'^\d{3,5}\s*', '', address.strip())
    address = address.replace("台", "臺")
    
    city, district, road, rest = "", "", "", ""
    
    city_match = re.search(r'(.+?[縣市])', address)
    if city_match:
        city = city_match.group(1)
        address = address[len(city):]
    
    dist_match = re.search(r'(.+?[區鎮鄉市])', address)
    if dist_match:
        district = dist_match.group(1)
        address = address[len(district):]
        
    road_match = re.search(r'(.+?[路街大道(段)])', address)
    if road_match:
        road = road_match.group(1)
        rest = address[len(road):]
    else:
        rest = address
        
    return city, district, road, rest

# 超商門市與店號精準拆分輔助函式
def parse_store_info(store_str):
    if not isinstance(store_str, str) or not store_str:
        return "", ""
    
    store_str = store_str.strip()
    
    code_match = re.search(r'\d{5,8}', store_str)
    store_code = code_match.group(0) if code_match else ""
    
    store_name = re.sub(r'\[.*?\]', '', store_str)
    store_name = re.sub(r'\(.*?\)', '', store_name)
    store_name = re.sub(r'（.*?）', '', store_name)
    store_name = re.sub(r'代碼[\:：]?\d+', '', store_name)
    store_name = store_name.replace("台", "臺").strip()
    
    return store_code, store_name

# 電話號碼格式化輔助函式（專治 Excel 號碼開頭 0 消失）
def format_phone_number(phone_val):
    if pd.isna(phone_val) or phone_val is None:
        return ""
    
    # 轉為純文字並去除小數點 (如浮點數轉過來的 .0)
    phone_str = str(phone_val).split('.')[0].strip()
    phone_str = re.sub(r'[^\d]', '', phone_str) # 只保留數字
    
    # 如果是 9 位數（開頭 0 被吃掉的手機號碼），自動在最前面補 0
    if len(phone_str) == 9 and phone_str.startswith('9'):
        phone_str = '0' + phone_str
        
    return phone_str

# 移除指定多餘欄位輔助函式（適用於順豐與海外）
def clean_overseas_columns(df_in, col_phone_name):
    df_out = df_in.copy()
    
    cols_to_drop = []
    for col in df_out.columns:
        col_str = str(col)
        if any(keyword in col_str for keyword in ["付款方式", "取貨門市", "取貨日期", "發票"]):
            cols_to_drop.append(col)
            
    if cols_to_drop:
        df_out = df_out.drop(columns=cols_to_drop)
        
    df_out.columns = [str(c).replace("台", "臺") for c in df_out.columns]
    
    # 處理所有文字內容，並對電話號碼補 0
    for col in df_out.columns:
        if col == col_phone_name or "電話" in col or "手機" in col:
            df_out[col] = df_out[col].apply(format_phone_number)
        else:
            df_out[col] = df_out[col].map(lambda x: str(x).replace("台", "臺") if isinstance(x, str) else x)
            
    return df_out

uploaded_file = st.file_uploader("請上傳訂單檔案 (支援 .xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # 讀取檔案時強制將所有欄位先視為字串，避免讀取階段 0 就消失
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)
            
        st.success(f"成功讀取檔案！共 {len(df)} 筆訂單。")
        st.markdown("---")
        
        st.subheader("⚙️ 欄位對應設定（請確認與你的表格標題一致）")
        cols = list(df.columns)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            col_ship = st.selectbox("1. 物流方式欄位", cols, index=0 if "物流" not in cols else cols.index("物流"))
            col_name = st.selectbox("2. 收件人姓名欄位", cols, index=0 if "姓名" not in cols else cols.index("姓名"))
        with c2:
            col_phone = st.selectbox("3. 收件人電話欄位", cols, index=0 if "電話" not in cols else cols.index("電話"))
            col_addr = st.selectbox("4. 收件地址欄位（會自動拆分為 縣市/分區/路名/剩餘地址）", cols, index=0 if "地址" not in cols else cols.index("地址"))
        with c3:
            col_store = st.selectbox("5. 超商門市欄位（會自動精準拆分為純數字店號與純門市名稱）", cols, index=0 if "門市" not in cols and "店" not in cols else (cols.index("門市") if "門市" in cols else cols.index("店")))

        if st.button("🚀 開始分類與整理"):
            # 1. 物流關鍵字比對
            df_post = df[df[col_ship].astype(str).str.contains("郵局|宅配|快捷|包裹", na=False)].copy()
            df_711 = df[df[col_ship].astype(str).str.contains("711|7-11|交貨便|統一", na=False)].copy()
            df_familymart = df[df[col_ship].astype(str).str.contains("全家|店到店", na=False)].copy()
            df_sf_hk = df[df[col_ship].astype(str).str.contains("順豐|香港|SF", na=False)].copy()
            
            known_keywords = "郵局|宅配|快捷|包裹|711|7-11|交貨便|統一|全家|店到店|順豐|香港|SF"
            df_overseas = df[~df[col_ship].astype(str).str.contains(known_keywords, na=False)].copy()

            # 整理郵局包裹 / 宅配
            post_data = []
            for _, row in df_post.iterrows():
                c, d, r, rest = parse_taiwan_address(str(row.get(col_addr, '')))
                name_clean = str(row.get(col_name, '')).replace("台", "臺")
                phone_clean = format_phone_number(row.get(col_phone, ''))
                post_data.append({
                    "收件人姓名": name_clean,
                    "收件人電話": phone_clean,
                    "完整地址": str(row.get(col_addr, '')).replace("台", "臺"),
                    "縣市": c,
                    "鄉鎮市區": d,
                    "路名": r,
                    "剩餘地址": rest
                })
            df_post_out = pd.DataFrame(post_data)

            # 整理 7-11 店到店
            seven_data = []
            for _, row in df_711.iterrows():
                code, name = parse_store_info(str(row.get(col_store, '')))
                name_clean = str(row.get(col_name, '')).replace("台", "臺")
                phone_clean = format_phone_number(row.get(col_phone, ''))
                seven_data.append({
                    "收件人姓名": name_clean,
                    "收件人電話": phone_clean,
                    "原始門市資訊": str(row.get(col_store, '')).replace("台", "臺"),
                    "收件店號": code,
                    "收件店名": name,
                    "寄件店號": "252975",
                    "寄件店名": "永吉門市"
                })
            df_711_out = pd.DataFrame(seven_data)

            # 整理 全家店到店
            family_data = []
            for _, row in df_familymart.iterrows():
                code, name = parse_store_info(str(row.get(col_store, '')))
                name_clean = str(row.get(col_name, '')).replace("台", "臺")
                phone_clean = format_phone_number(row.get(col_phone, ''))
                family_data.append({
                    "收件人姓名": name_clean,
                    "收件人電話": phone_clean,
                    "原始門市資訊": str(row.get(col_store, '')).replace("台", "臺"),
                    "收件店號": code,
                    "收件店名": name,
                    "寄件店號": "024502",
                    "寄件店名": "永吉二店"
                })
            df_familymart_out = pd.DataFrame(family_data)

            # 4. 香港順豐 & 海外
            df_sf_hk_out = clean_overseas_columns(df_sf_hk, col_phone)
            df_overseas_out = clean_overseas_columns(df_overseas, col_phone)

            # 輸出多 Sheet Excel（強制寫入為字串格式）
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_post_out.to_excel(writer, sheet_name='郵局包裹寄送', index=False)
                df_711_out.to_excel(writer, sheet_name='711店到店', index=False)
                df_familymart_out.to_excel(writer, sheet_name='全家店到店', index=False)
                df_sf_hk_out.to_excel(writer, sheet_name='香港順豐到付', index=False)
                df_overseas_out.to_excel(writer, sheet_name='其他海外寄送', index=False)
                
            output.seek(0)

            st.success("🎉 電話號碼已被修正！開頭的 0 已完全保留與補齊。")

            t1, t2, t3, t4, t5 = st.tabs(["郵局包裹", "7-11店到店", "全家店到店", "香港順豐", "其他海外"])
            with t1: st.dataframe(df_post_out)
            with t2: st.dataframe(df_711_out)
            with t3: st.dataframe(df_familymart_out)
            with t4: st.dataframe(df_sf_hk_out)
            with t5: st.dataframe(df_overseas_out)

            orig_name = uploaded_file.name.rsplit('.', 1)[0]
            st.download_button(
                label="💾 下載分類整理後的 Excel 檔 (.xlsx)",
                data=output,
                file_name=f"{orig_name}_已分類出貨單.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")
