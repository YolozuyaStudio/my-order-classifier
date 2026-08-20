import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="訂單自動分類整理工具", page_icon="📋", layout="wide")

st.title("📋 訂單自動分類與欄位整理工具")
st.markdown("上傳原始訂單 Excel/CSV 後，系統會自動按物流方式分類並整理成專屬欄位！")

# 台灣地址拆分輔助函式
def parse_taiwan_address(address):
    if not isinstance(address, str) or not address:
        return "", "", "", ""
    address = re.sub(r'^\d{3,5}\s*', '', address.strip())
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
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            col_ship = st.selectbox("物流方式欄位", cols, index=0 if "物流" not in cols else cols.index("物流"))
            col_name = st.selectbox("收件人姓名欄位", cols, index=0 if "姓名" not in cols else cols.index("姓名"))
        with c2:
            col_phone = st.selectbox("收件人電話欄位", cols, index=0 if "電話" not in cols else cols.index("電話"))
            col_addr = st.selectbox("收件地址欄位", cols, index=0 if "地址" not in cols else cols.index("地址"))
        with c3:
            col_store_code = st.selectbox("門市店號欄位", cols, index=0 if "店號" not in cols else cols.index("店號"))
        with c4:
            col_store_name = st.selectbox("門市店名欄位", cols, index=0 if "店名" not in cols else cols.index("店名"))

        if st.button("🚀 開始分類與整理"):
            # 分類邏輯
            df_post = df[df[col_ship].astype(str).str.contains("郵局|快捷|包裹", na=False)].copy()
            df_711 = df[df[col_ship].astype(str).str.contains("711|7-11|交貨便|統一", na=False)].copy()
            df_familymart = df[df[col_ship].astype(str).str.contains("全家|店到店", na=False)].copy()
            df_sf_hk = df[df[col_ship].astype(str).str.contains("順豐|香港|SF", na=False)].copy()
            
            known_keywords = "郵局|快捷|包裹|711|7-11|交貨便|統一|全家|店到店|順豐|香港|SF"
            df_overseas = df[~df[col_ship].astype(str).str.contains(known_keywords, na=False)].copy()

            # 1. 整理郵局
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

            # 2. 整理 7-11
            df_711_out = pd.DataFrame({
                "收件人姓名": df_711[col_name],
                "收件人電話": df_711[col_phone],
                "收件店號": df_711[col_store_code],
                "收件店名": df_711[col_store_name],
                "寄件店號": "252975",
                "寄件店名": "永吉門市"
            })

            # 3. 整理 全家
            df_familymart_out = pd.DataFrame({
                "收件人姓名": df_familymart[col_name],
                "收件人電話": df_familymart[col_phone],
                "收件店號": df_familymart[col_store_code],
                "收件店名": df_familymart[col_store_name],
                "寄件店號": "024502",
                "寄件店名": "永吉二店"
            })

            # 4. 香港順豐 & 海外
            df_sf_hk_out = df_sf_hk
            df_overseas_out = df_overseas

            # 輸出多 Sheet Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_post_out.to_excel(writer, sheet_name='郵局包裹寄送', index=False)
                df_711_out.to_excel(writer, sheet_name='711店到店', index=False)
                df_familymart_out.to_excel(writer, sheet_name='全家店到店', index=False)
                df_sf_hk_out.to_excel(writer, sheet_name='香港順豐到付', index=False)
                df_overseas_out.to_excel(writer, sheet_name='其他海外寄送', index=False)
                
            output.seek(0)

            st.success("🎉 分類完成！您可以點擊下方頁籤預覽，或直接下載 Excel 檔。")

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
