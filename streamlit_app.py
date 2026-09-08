import streamlit as st
import pandas as pd
from datetime import datetime
import os

# 1. 頁面基本配置
st.set_page_config(
    page_title="北大衛保組 - 藥品庫存管理系統",
    page_icon="🏥",
    layout="wide"
)

# 2. 實體檔案路徑設定
INV_FILE = "inventory.csv"
LOG_FILE = "logs.csv"

# 3. 資料讀取與自動初始化函數
def load_data():
    if os.path.exists(INV_FILE):
        df_inv = pd.read_csv(INV_FILE)
    else:
        # 初次使用時自動建置預設藥品資料庫
        df_inv = pd.DataFrame([
            {"藥品名稱": "Actein (愛克痰)", "適應症": "化痰", "目前庫存": 803, "效期": "2028-04-01"},
            {"藥品名稱": "Fexofenadine (飛敏耐)", "適應症": "抗組織胺/過敏", "目前庫存": 1054, "效期": "2026-08-15"},
            {"藥品名稱": "Amoxicillin 500mg", "適應症": "抗生素", "目前庫存": 1556, "效期": "2026-08-30"},
            {"藥品名稱": "Purfen (普服芬)", "適應症": "解熱/消炎/止痛", "目前庫存": 1267, "效期": "2027-10-10"},
            {"藥品名稱": "Biofermin (表飛鳴)", "適應症": "整腸健胃", "目前庫存": 1062, "效期": "2027-05-20"},
            {"藥品名稱": "C.B. oint (強力施美藥膏)", "適應症": "止癢", "目前庫存": 45, "效期": "2026-11-01"}
        ])
        df_inv.to_csv(INV_FILE, index=False, encoding="utf-8-sig")

    if os.path.exists(LOG_FILE):
        df_logs = pd.read_csv(LOG_FILE)
    else:
        df_logs = pd.DataFrame(columns=["登記時間", "日期", "藥品名稱", "消耗數量", "用途分類"])
        df_logs.to_csv(LOG_FILE, index=False, encoding="utf-8-sig")

    return df_inv, df_logs

# 4. 存檔至硬碟函數
def save_data(df_inv, df_logs):
    df_inv.to_csv(INV_FILE, index=False, encoding="utf-8-sig")
    df_logs.to_csv(LOG_FILE, index=False, encoding="utf-8-sig")

# 5. 狀態初始化
if "inventory" not in st.session_state or "logs" not in st.session_state:
    st.session_state.inventory, st.session_state.logs = load_data()

# 標題區塊
st.title("🏥 國立臺北大學衛保組 - 藥品庫存與消耗管理系統")
st.caption("🔒 單機自動實體存檔版｜每次登記皆即時同步寫入電腦硬碟")

# 6. 分頁介面規劃
tab1, tab2, tab3 = st.tabs(["⚡ 每日發藥登記", "📦 庫存與效期查詢", "📊 報表與歷史紀錄"])

# --- TAB 1: 每日發藥登記 ---
with tab1:
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📝 快速發藥登錄")
        with st.form(key="dispense_form", clear_on_submit=True):
            drug_list = st.session_state.inventory["藥品名稱"].tolist()
            selected_drug = st.selectbox("1. 選擇或搜尋藥品", drug_list)
            qty = st.number_input("2. 輸入發放數量", min_value=1, value=1, step=1)
            category = st.radio("3. 用途分類", ["一般消耗", "公藥使用", "過期報銷"], horizontal=True)
            
            submit = st.form_submit_button("確認送出並自動扣庫存 (Enter)")
            
            if submit:
                now = datetime.now()
                today_str = now.strftime("%Y-%m-%d")
                datetime_str = now.strftime("%Y-%m-%d %H:%M:%S")
                
                # 搜尋選擇的藥品索引
                idx = st.session_state.inventory[st.session_state.inventory["藥品名稱"] == selected_drug].index[0]
                current_qty = st.session_state.inventory.at[idx, "目前庫存"]
                
                if current_qty < qty:
                    st.error(f"⚠️ 庫存不足！{selected_drug} 目前僅剩 {current_qty}")
                else:
                    # 扣減庫存
                    st.session_state.inventory.at[idx, "目前庫存"] -= qty
                    
                    # 新增消耗日誌
                    new_log = pd.DataFrame([{
                        "登記時間": datetime_str,
                        "日期": today_str,
                        "藥品名稱": selected_drug,
                        "消耗數量": qty,
                        "用途分類": category
                    }])
                    st.session_state.logs = pd.concat([new_log, st.session_state.logs], ignore_index=True)
                    
                    # 寫入實體檔案存檔
                    save_data(st.session_state.inventory, st.session_state.logs)
                    st.success(f"✅ 已成功扣除 {selected_drug} 共 {qty} 單位，資料已自動寫入硬碟！")
                    st.rerun()

    with col2:
        st.subheader("📋 今日發藥紀錄")
        today_date = datetime.now().strftime("%Y-%m-%d")
        if not st.session_state.logs.empty:
            today_logs = st.session_state.logs[st.session_state.logs["日期"] == today_date]
            if not today_logs.empty:
                st.dataframe(today_logs[["登記時間", "藥品名稱", "消耗數量", "用途分類"]], use_container_width=True, hide_index=True)
            else:
                st.info("今日尚無發藥紀錄")
        else:
            st.info("尚無歷史紀錄")

# --- TAB 2: 庫存與效期查詢 ---
with tab2:
    st.subheader("📦 目前藥品庫存清單")
    
    # 低庫存警報提示 (低於 100 單位)
    low_stock = st.session_state.inventory[st.session_state.inventory["目前庫存"] < 100]
    if not low_stock.empty:
        st.warning("⚠️ 注意：以下藥品庫存低於安全水位 (100 單位)：")
        st.dataframe(low_stock[["藥品名稱", "目前庫存", "效期"]], hide_index=True)
        
    st.dataframe(st.session_state.inventory, use_container_width=True, hide_index=True)

# --- TAB 3: 報表與歷史紀錄 ---
with tab3:
    st.subheader("📜 完整歷史消耗日誌")
    st.dataframe(st.session_state.logs, use_container_width=True, hide_index=True)
    
    st.divider()
    st.subheader("📥 匯出 Excel 月報表")
    
    if st.button("生成本月簡化 Excel 報表"):
        filename = f"衛保組藥品月報表_{datetime.now().strftime('%Y%m')}.xlsx"
        
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            st.session_state.inventory.to_excel(writer, sheet_name="月結算總表", index=False)
            st.session_state.logs.to_excel(writer, sheet_name="每日發藥明細", index=False)
            
        with open(filename, "rb") as f:
            st.download_button(
                label="點此下載 Excel 月報表",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
