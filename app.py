import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="訂單自動分類整理工具", page_icon="📋", layout="wide")

st.title("📋 訂單自動分類與欄位整理工具")
st.markdown("上傳原始訂單 Excel/CSV 後，系統會自動按物流方式分類並整理成專屬欄位！")

# 台灣地址拆分輔助函式（縣市、鄉鎮市區、路名、剩餘地址）
def parse_taiwan_address(address):
    if not isinstance(address, str) or not address:
        return "", "", "", ""
    
    # 移除郵遞區號與多餘空白
    address = re.sub(r'^\d{3,5}\s*', '', address.strip())
    
    city, district, road, rest = "", "", "", ""
    
    # 抓取縣市
    city_match = re.search(r'(.+?[縣市])', address)
    if city_match:
        city = city_match.group(1)
        address = address[len(city):]
    
    # 抓取鄉鎮市區
    dist_match = re.search(r'(.+?[區鎮鄉市])', address)
    if dist_match:
        district = dist_match.group(1)
        address = address[len(district):]
        
    # 抓取路段/街/大道
    road_match = re.search(r'(.+?[路街大道(段)])', address)
    if road_match:
        road = road_match.group(1)
        rest = address[len(road):]
    else:
        rest = address
        
    return city, district, road, rest

# 店號與店名拆分輔助函式
def parse_store_info(store_str):
    if not isinstance(store_str, str) or not store_str:
        return "", ""
    store_str = store_str.strip()
    
    # 尋找 5-8 位的連續數字作為店號
    code_match = re.search(r'\b\d{5,8}\b', store_str)
    if code_match:
        store_code = code_match.group(0)
        # 清除店號及常見分隔符，剩餘字串即為店名
        store_name = re.sub(r'\b\d{5,8}\b', '', store_str)
        store_name = re.sub(r'[\(\)（）\-\_\s]', '', store_name)
        return store_code, store_name
    else:
        # 若無數字，嘗試用常見符號切割
        parts = re.split(r'[\(\)（）\-\_\s]+', store_str)
        parts = [p for p in parts if p]
        if len(parts) >= 2:
            return parts[0], parts[1]
        return "", store_str

uploaded_file = st.file_uploader("請上傳訂單檔案 (支援 .xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
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
            col_store = st.selectbox("5. 超商門市欄位（店名/店號在同欄位，會自動拆分）", cols, index=0 if "門市" not in cols and "店" not in cols else (cols.index("門市") if "門市" in cols else cols.index("店")))

        if st.button("🚀 開始分類與整理"):
            # 1. 物流關鍵字比對（郵局包含：郵局、宅配、快捷、包裹）
            df_post = df[df[col_ship].astype(str).str.contains("郵局|宅配|快捷|包裹", na=False)].copy()
            df_711 = df[df[col_ship].astype(str).str.contains("711|7-11|交貨便|統一", na=False)].copy()
            df_familymart = df[df[col_ship].astype(str).str.contains("全家|店到店", na=False)].copy()
            df_sf_hk = df[df[col_ship].astype(str).str.contains("順豐|香港|SF", na=False)].copy()
            
            known_keywords = "郵局|宅配|快捷|包裹|711|7-11|交貨便|統一|全家|店到店|順豐|香港|SF"
            df_overseas = df[~df[col_ship].astype(str).str.contains(known_keywords, na=False)].copy()

            # 整理郵局包裹
            post_data = []
            for _, row in df_post.iterrows():
                c, d, r, rest = parse_taiwan_address(str(row.get(col_addr, '')))
                post_data.append({
                    "收件人姓名": row.get(col_name, ''),
                    "收件人電話": row.get(col_phone, ''),
                    "完整地址": row.get(col_addr, ''),
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
                seven_data.append({
                    "收件人姓名": row.get(col_name, ''),
                    "收件人電話": row.get(col_phone, ''),
                    "原始門市資訊": row.get(col_store, ''),
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
                family_data.append({
                    "收件人姓名": row.get(col_name, ''),
                    "收件人電話": row.get(col_phone, ''),
                    "原始門市資訊": row.get(col_store, ''),
                    "收件店號": code,
                    "收件店名": name,
                    "寄件店號": "024502",
                    "寄件店名": "永吉二店"
                })
            df_familymart_out = pd.DataFrame(family_data)

            # 4. 香港順豐 & 海外（保留完整原始資料）
            df_sf_hk_out = df_sf_hk
            df_overseas_out = df_overseas

            # 輸出成多個 Sheet 的 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_post_out.to_excel(writer, sheet_name='郵局包裹寄送', index=False)
                df_711_out.to_excel(writer, sheet_name='711店到店', index=False)
                df_familymart_out.to_excel(writer, sheet_name='全家店到店', index=False)
                df_sf_hk_out.to_excel(writer, sheet_name='香港順豐到付', index=False)
                df_overseas_out.to_excel(writer, sheet_name='其他海外寄送', index=False)
                
            output.seek(0)

            st.success("🎉 分類與欄位拆分完成！您可以點擊下方頁籤預覽，或直接下載 Excel 檔。")

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
