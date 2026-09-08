import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

# 1. 頁面基本配置
st.set_page_config(
    page_title="衛保組管理系統",
    page_icon="💊",
    layout="wide"
)

# 2. 初始化 Session State 資料庫
if "inventory" not in st.session_state:
    st.session_state.inventory = pd.DataFrame([
        {"藥品名稱 (英文)": "Actein 600mg", "中文名稱": "愛克痰發泡錠", "目前庫存": 168, "有效期限": "2028-04-30", "用途/備註": "去痰"},
        {"藥品名稱 (英文)": "Actein 600mg", "中文名稱": "愛克痰發泡錠", "目前庫存": 461, "有效期限": "2028-05-31", "用途/備註": "去痰"},
        {"藥品名稱 (英文)": "Amoxicillin 500mg", "中文名稱": "安莫西林", "目前庫存": 400, "有效期限": "2028-02-28", "用途/備註": "抗生素/葡萄球菌/鏈球菌/肺"},
        {"藥品名稱 (英文)": "Amoxicillin 500mg", "中文名稱": "安莫西林", "目前庫存": 1000, "有效期限": "2028-09-03", "用途/備註": "抗生素/葡萄球菌/鏈球菌/肺"},
        {"藥品名稱 (英文)": "Amoxicillin 250mg", "中文名稱": "安莫西林", "目前庫存": 0, "有效期限": "2026-12-31", "用途/備註": "抗生素/葡萄球菌/鏈球菌/肺"},
        {"藥品名稱 (英文)": "Ancogen", "中文名稱": "安可腱", "目前庫存": 380, "有效期限": "2027-09-30", "用途/備註": "骨骼肌肉鬆弛/腰椎/脊椎/關"},
        {"藥品名稱 (英文)": "Biofermin", "中文名稱": "表飛鳴", "目前庫存": 604, "有效期限": "2028-05-31", "用途/備註": "腹瀉/整腸/便祕"},
        {"藥品名稱 (英文)": "Meclizine 25mg", "中文名稱": "美克利靜", "目前庫存": 948, "有效期限": "2026-09-12", "用途/備註": "暈眩/緩動暈症"}
    ])

if "logs" not in st.session_state:
    st.session_state.logs = pd.DataFrame(columns=[
        "領用日期", "登記時間", "藥品名稱", "中文名稱", "領用數量", "用途分類", "備註"
    ])

# 3. 側邊欄控制與資料同步
st.sidebar.title("🏥 衛保組管理系統")
page = st.sidebar.radio("📍 請選擇功能頁面", ["💊 藥品領用與紀錄", "📦 庫存盤點與校正", "☁️ 雲端報表匯出"])

st.sidebar.markdown("---")
st.sidebar.subheader("🔗 串接 Google 雲端試算表")
gsheet_url = st.sidebar.text_input(
    "貼上 Google 試算表連結：",
    value="https://docs.google.com/spreadsheets/d/1gv_1Fz0iR9kUFVyj9P_IN0dJvz-drnv5v-wQ_UY8qvN0/edit?usp=sharing"
)

if st.sidebar.button("🔄 同步雲端試算表資料"):
    try:
        # 自動轉換 edit 連結為 csv 匯出網址
        if "/edit" in gsheet_url:
            csv_url = gsheet_url.split("/edit")[0] + "/export?format=csv"
        else:
            csv_url = gsheet_url
        
        df_cloud = pd.read_csv(csv_url)
        st.session_state.inventory = df_cloud
        st.sidebar.success("✅ 已成功同步雲端資料庫！")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"❌ 讀取失敗，請確認共享權限：{e}")

st.sidebar.markdown("---")
st.sidebar.subheader("📤 上傳本地 CSV / Excel 檔")
uploaded_file = st.sidebar.file_uploader("手動上傳藥品清單", type=["csv", "xlsx"])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            st.session_state.inventory = pd.read_csv(uploaded_file)
        else:
            st.session_state.inventory = pd.read_excel(uploaded_file)
        st.sidebar.success("✅ 檔案載入成功！")
    except Exception as e:
        st.sidebar.error(f"❌ 檔案格式錯誤：{e}")


# 4. 先進先出 (FIFO) 自動扣庫存函數
def fifo_deduct(df_inv, drug_eng_name, qty_needed):
    matches = df_inv[df_inv['藥品名稱 (英文)'] == drug_eng_name].copy()
    if matches.empty:
        return df_inv, False, f"找不到藥品：{drug_eng_name}"
    
    total_stock = matches['目前庫存'].sum()
    if total_stock < qty_needed:
        return df_inv, False, f"{drug_eng_name} 總庫存不足 (現有: {total_stock}, 需要: {qty_needed})"
    
    # 依照有效期限由近至遠排序 (先進先出)
    matches = matches.sort_values(by='有效期限')
    rem = qty_needed
    for idx in matches.index:
        current_stock = df_inv.at[idx, '目前庫存']
        if current_stock >= rem:
            df_inv.at[idx, '目前庫存'] = current_stock - rem
            rem = 0
            break
        else:
            rem -= current_stock
            df_inv.at[idx, '目前庫存'] = 0
            
    return df_inv, True, "扣減成功"


# --- 頁面 1: 藥品領用與紀錄 ---
if page == "💊 藥品領用與紀錄":
    col_left, col_right = st.columns([1.3, 1])
    
    with col_left:
        st.subheader("💊 藥品領用與登記")
        st.caption("點選下方搜尋欄可選擇一種或多種藥品，設定數量與領用日期後即可一次完成登記與庫存扣減。")
        
        # 1. 領用日期與類別（已新增領用日期選擇器）
        c_date, c_cat = st.columns([1, 1])
        with c_date:
            issue_date = st.date_input("📅 領用日期", value=date.today())
        with c_cat:
            category = st.selectbox("🏷️ 用途分類", ["一般消耗", "公藥使用", "過期報銷", "其他"])
            
        # 2. 藥品清單選單 (去除重複英文名稱)
        unique_drugs = st.session_state.inventory.drop_duplicates(subset=['藥品名稱 (英文)'])
        drug_options = [
            f"{row['藥品名稱 (英文)']} ({row['中文名稱']})" 
            for _, row in unique_drugs.iterrows()
        ]
        
        selected_options = st.multiselect(
            "選擇本次領取的所有藥品 (可同時選擇多項)",
            options=drug_options,
            placeholder="請點擊或輸入藥名/中文名稱進行搜尋..."
        )
        
        if not selected_options:
            st.info("💡 請先在上方的選單中點選或搜尋要領取的藥品。")
        else:
            st.markdown("---")
            st.markdown("##### 🔢 設定發放數量與細項")
            
            # 動態建立多項藥品的數量輸入框
            quantities = {}
            for opt in selected_options:
                drug_eng = opt.split(" (")[0]
                # 計算該藥品總庫存
                total_k = st.session_state.inventory[
                    st.session_state.inventory['藥品名稱 (英文)'] == drug_eng
                ]['目前庫存'].sum()
                
                q = st.number_input(
                    f"數量 - {opt} [總庫存: {total_k}]", 
                    min_value=1, 
                    value=1, 
                    step=1, 
                    key=f"qty_{opt}"
                )
                quantities[opt] = q
                
            note = st.text_input("📝 備註說明 (選填)", placeholder="例如：學生領用、研討會備藥...")
            
            if st.button("確認登記並扣減庫存", type="primary", use_container_width=True):
                success_all = True
                messages = []
                
                # 執行多筆扣量與日誌紀錄
                for opt, qty in quantities.items():
                    drug_eng = opt.split(" (")[0]
                    chinese_name = opt.split(" (")[1].replace(")", "") if "(" in opt else ""
                    
                    df_inv, ok, msg = fifo_deduct(st.session_state.inventory, drug_eng, qty)
                    if ok:
                        st.session_state.inventory = df_inv
                        # 新增日誌 (包含選擇的領用日期與系統登記時間)
                        new_log = pd.DataFrame([{
                            "領用日期": issue_date.strftime("%Y-%m-%d"),
                            "登記時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "藥品名稱": drug_eng,
                            "中文名稱": chinese_name,
                            "領用數量": qty,
                            "用途分類": category,
                            "備註": note
                        }])
                        st.session_state.logs = pd.concat([new_log, st.session_state.logs], ignore_index=True)
                    else:
                        success_all = False
                        messages.append(msg)
                
                if success_all:
                    st.success("✅ 藥品領用登記成功！庫存已完成先進先出 (FIFO) 自動扣減。")
                    st.rerun()
                else:
                    st.error("⚠️ 登記過程發生錯誤：" + "；".join(messages))

    with col_right:
        st.subheader("📋 當前藥品庫存總覽")
        st.dataframe(
            st.session_state.inventory,
            use_container_width=True,
            hide_index=True,
            height=450
        )

    # 底部顯示發藥紀錄
    st.markdown("---")
    st.subheader("📜 歷史領用與發藥紀錄")
    if not st.session_state.logs.empty:
        st.dataframe(st.session_state.logs, use_container_width=True, hide_index=True)
    else:
        st.info("目前尚無任何領用紀錄。")


# --- 頁面 2: 庫存盤點與校正 ---
elif page == "📦 庫存盤點與校正":
    st.subheader("📦 庫存盤點與資料手動校正")
    st.caption("您可以直接在下方表格中修改庫存數量、有效期限或新增批號，完成後點擊「儲存盤點結果」。")
    
    edited_df = st.data_editor(
        st.session_state.inventory,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )
    
    if st.button("💾 儲存盤點校正結果", type="primary"):
        st.session_state.inventory = edited_df
        st.success("✅ 庫存資料已成功更新！")
        st.rerun()


# --- 頁面 3: 雲端報表匯出 ---
elif page == "☁️ 雲端報表匯出":
    st.subheader("☁️ 報表預覽與匯出")
    st.caption("匯出排版乾淨、包含「月結算總表」與「每日發藥明細」的 Excel 報表。")
    
    tab_summary, tab_detail = st.tabs(["月結算總表預覽", "每日發藥明細預覽"])
    
    with tab_summary:
        st.dataframe(st.session_state.inventory, use_container_width=True, hide_index=True)
        
    with tab_detail:
        st.dataframe(st.session_state.logs, use_container_width=True, hide_index=True)
        
    st.markdown("---")
    
    # 匯出 Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        st.session_state.inventory.to_excel(writer, sheet_name="月結算總表", index=False)
        st.session_state.logs.to_excel(writer, sheet_name="每日發藥明細", index=False)
        
    st.download_button(
        label="📥 下載完整 Excel 月報表",
        data=buffer.getvalue(),
        file_name=f"衛保組藥品月報表_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
