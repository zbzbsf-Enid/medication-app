import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime

# ---------------------------------------------------------
# 1. 頁面基本設定
# ---------------------------------------------------------
st.set_page_config(
    page_title="衛保組藥品庫存管理系統",
    page_icon="💊",
    layout="wide"
)

# 初始化 Google Sheets 連線
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"❌ 初始化 GSheetsConnection 失敗，請確認 secrets.toml 設定: {e}")
    st.stop()

# 取得原生 gspread Spreadsheet 物件 (解包 GSheetsServiceAccountClient)
def get_spreadsheet():
    try:
        url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        client = getattr(conn.client, "_client", conn.client)
        return client.open_by_url(url)
    except Exception as e:
        st.error(f"❌ 開啟雲端試算表失敗: {e}")
        return None

# 讀取特定分頁資料
def load_sheet_data(worksheet_name: str) -> pd.DataFrame:
    sh = get_spreadsheet()
    if not sh:
        return pd.DataFrame()
    try:
        ws = sh.worksheet(worksheet_name)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception:
        return pd.DataFrame()

# 原生 gspread 寫入邏輯（含自動建立分頁與預設標題）
def safe_update_sheet(worksheet_name: str, df: pd.DataFrame, default_headers: list = None):
    sh = get_spreadsheet()
    if not sh:
        return False, "無法連接至 Google 試算表"
    
    try:
        try:
            ws = sh.worksheet(worksheet_name)
        except Exception:
            ws = sh.add_worksheet(title=worksheet_name, rows="1000", cols="20")
            if default_headers and df.empty:
                df = pd.DataFrame(columns=default_headers)

        clean_df = df.fillna("").copy()
        for col in clean_df.columns:
            clean_df[col] = clean_df[col].astype(str)

        ws.clear()
        if not clean_df.empty:
            ws.update([clean_df.columns.values.tolist()] + clean_df.values.tolist())
        elif default_headers:
            ws.update([default_headers])
            
        return True, "成功寫入雲端"
    except Exception as err:
        return False, str(err)

# ---------------------------------------------------------
# 2. 左側邊欄功能選單與維護區
# ---------------------------------------------------------
st.sidebar.title("💊 衛保組藥品系統")
st.sidebar.markdown("---")

# 主功能選單
page = st.sidebar.radio(
    "📌 請選擇功能：",
    ["📦 藥品庫存清單", "📋 領用登記", "🚚 進貨登記", "📜 歷史紀錄"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 系統設定與維護")

# 顯示雲端連線狀態與分頁資訊
sh = get_spreadsheet()
if sh:
    try:
        existing_sheets = [ws.title for ws in sh.worksheets()]
        st.sidebar.success("✅ Google Sheet 連線成功")
        st.sidebar.write("🔍 目前偵測到的雲端分頁：", existing_sheets)
    except Exception as e:
        st.sidebar.warning(f"⚠️ 無法讀取分頁列表: {e}")

# 手動刷新快取按鈕
if st.sidebar.button("🔄 手動刷新雲端資料", use_container_width=True):
    st.cache_data.clear()
    st.sidebar.success("已刷新資料快取！")
    st.rerun()

# ---------------------------------------------------------
# 3. 主頁面內容控制（根據左側選單切換）
# ---------------------------------------------------------

# --- 頁面 1: 藥品庫存清單 ---
if page == "📦 藥品庫存清單":
    st.title("📦 目前庫存狀態")
    inventory_df = load_sheet_data("庫存")
    
    if not inventory_df.empty:
        search_kw = st.text_input("🔍 搜尋藥品（英文或中文名稱）：", "")
        if search_kw:
            filtered_df = inventory_df[
                inventory_df["藥品名稱"].astype(str).str.contains(search_kw, case=False, na=False) |
                inventory_df["中文名稱"].astype(str).str.contains(search_kw, case=False, na=False)
            ]
        else:
            filtered_df = inventory_df
            
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    else:
        st.info("目前「庫存」分頁尚無資料或試算表載入中。")

# --- 頁面 2: 領用登記 ---
elif page == "📋 領用登記":
    st.title("📋 藥品領用登記")
    inventory_df = load_sheet_data("庫存")
    
    if inventory_df.empty or "藥品名稱" not in inventory_df.columns:
        st.warning("⚠️ 請先確保「庫存」分頁有包含「藥品名稱」的欄位資料。")
    else:
        with st.form("usage_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            med_options = inventory_df.apply(
                lambda row: f"{row.get('藥品名稱', '')} ({row.get('中文名稱', '')})", axis=1
            ).tolist()
            
            selected_med_str = col1.selectbox("選擇領用藥品", med_options)
            use_qty = col2.number_input("領用數量", min_value=1, step=1, value=1)
            
            use_date = col1.date_input("領用日期", datetime.now())
            remarks = col2.text_input("備註 / 領用單位或個人", "")
            
            submit_btn = st.form_submit_button("確認無誤，寫入雲端並更新庫存", use_container_width=True)
            
            if submit_btn:
                idx = med_options.index(selected_med_str)
                selected_row = inventory_df.iloc[idx]
                med_name = selected_row.get("藥品名稱", "")
                zh_name = selected_row.get("中文名稱", "")
                
                # 1. 寫入「領用紀錄」
                usage_df = load_sheet_data("領用紀錄")
                new_usage_row = pd.DataFrame([{
                    "領用時間": str(use_date),
                    "藥品名稱": med_name,
                    "中文名稱": zh_name,
                    "領用數量": use_qty,
                    "備註": remarks
                }])
                
                updated_usage_df = pd.concat([usage_df, new_usage_row], ignore_index=True)
                
                success_usage, msg1 = safe_update_sheet(
                    "領用紀錄", 
                    updated_usage_df, 
                    default_headers=["領用時間", "藥品名稱", "中文名稱", "領用數量", "備註"]
                )
                
                # 2. 更新「庫存」
                if success_usage:
                    curr_qty = pd.to_numeric(inventory_df.at[idx, "現有庫存"], errors='coerce')
                    if pd.isna(curr_qty):
                        curr_qty = pd.to_numeric(inventory_df.at[idx, "剩餘庫存"], errors='coerce')
                    curr_qty = 0 if pd.isna(curr_qty) else int(curr_qty)
                    
                    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else "剩餘庫存"
                    inventory_df.at[idx, target_col] = max(0, curr_qty - use_qty)
                    
                    success_inv, msg2 = safe_update_sheet("庫存", inventory_df)
                    
                    if success_inv:
                        st.success(f"✅ 成功登記領用：{zh_name} {use_qty} 個！庫存已同步扣減。")
                        st.cache_data.clear()
                    else:
                        st.error(f"❌ 領用紀錄已寫入，但庫存更新失敗：{msg2}")
                else:
                    st.error(f"❌ 寫入「領用紀錄」失敗：{msg1}")

# --- 頁面 3: 進貨登記 ---
elif page == "🚚 進貨登記":
    st.title("🚚 藥品進貨登記")
    inventory_df = load_sheet_data("庫存")
    
    with st.form("restock_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        med_list = inventory_df["藥品名稱"].tolist() if not inventory_df.empty and "藥品名稱" in inventory_df.columns else []
        med_list.insert(0, "+ 新增未在庫存的藥品")
        
        selected_option = col1.selectbox("選擇或新增藥品", med_list)
        
        if selected_option == "+ 新增未在庫存的藥品":
            med_name = col1.text_input("藥品英文名稱", "")
            zh_name = col2.text_input("藥品中文名稱", "")
        else:
            med_name = selected_option
            matched = inventory_df[inventory_df["藥品名稱"] == med_name]
            zh_name = matched.iloc[0]["中文名稱"] if not matched.empty else ""
            col2.text_input("藥品中文名稱", value=zh_name, disabled=True)
            
        restock_qty = col1.number_input("進貨數量", min_value=1, step=1, value=100)
        restock_date = col2.date_input("進貨日期", datetime.now())
        restock_remarks = st.text_input("進貨備註 / 廠商資訊", "")
        
        restock_btn = st.form_submit_button("確認進貨並更新雲端庫存", use_container_width=True)
        
        if restock_btn:
            if not med_name:
                st.error("請輸入藥品名稱！")
            else:
                restock_df = load_sheet_data("進貨紀錄")
                new_restock_row = pd.DataFrame([{
                    "進貨時間": str(restock_date),
                    "藥品名稱": med_name,
                    "中文名稱": zh_name,
                    "進貨數量": restock_qty,
                    "備註": restock_remarks
                }])
                updated_restock_df = pd.concat([restock_df, new_restock_row], ignore_index=True)
                
                success_restock, msg1 = safe_update_sheet(
                    "進貨紀錄", 
                    updated_restock_df,
                    default_headers=["進貨時間", "藥品名稱", "中文名稱", "進貨數量", "備註"]
                )
                
                if success_restock:
                    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else "剩餘庫存"
                    
                    if not inventory_df.empty and med_name in inventory_df["藥品名稱"].values:
                        idx = inventory_df[inventory_df["藥品名稱"] == med_name].index[0]
                        curr_qty = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                        curr_qty = 0 if pd.isna(curr_qty) else int(curr_qty)
                        inventory_df.at[idx, target_col] = curr_qty + restock_qty
                    else:
                        new_inv_row = pd.DataFrame([{
                            "藥品名稱": med_name,
                            "中文名稱": zh_name,
                            target_col: restock_qty,
                            "備註": ""
                        }])
                        inventory_df = pd.concat([inventory_df, new_inv_row], ignore_index=True)
                        
                    success_inv, msg2 = safe_update_sheet("庫存", inventory_df)
                    if success_inv:
                        st.success(f"✅ 成功進貨：{zh_name or med_name} {restock_qty} 個！庫存已更新。")
                        st.cache_data.clear()
                    else:
                        st.error(f"❌ 進貨紀錄已寫入，但庫存更新失敗：{msg2}")
                else:
                    st.error(f"❌ 寫入「進貨紀錄」失敗：{msg1}")

# --- 頁面 4: 歷史紀錄 ---
elif page == "📜 歷史紀錄":
    st.title("📜 歷史異動紀錄")
    sub_tab1, sub_tab2 = st.tabs(["📋 領用紀錄歷史", "🚚 進貨紀錄歷史"])
    
    with sub_tab1:
        u_df = load_sheet_data("領用紀錄")
        if not u_df.empty:
            st.dataframe(u_df, use_container_width=True, hide_index=True)
        else:
            st.info("目前無領用紀錄。")
            
    with sub_tab2:
        r_df = load_sheet_data("進貨紀錄")
        if not r_df.empty:
            st.dataframe(r_df, use_container_width=True, hide_index=True)
        else:
            st.info("目前無進貨紀錄。")
