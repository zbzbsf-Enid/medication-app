import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import zoneinfo

# ==============================================================================
# 1. 頁面基礎設定與時區配置
# ==============================================================================
st.set_page_config(
    page_title="國立臺北大學衛保組 藥品管理系統",
    page_icon="💊",
    layout="wide"
)

# 台灣標準時區 (GMT+8)
TAIWAN_TZ = zoneinfo.ZoneInfo("Asia/Taipei")

# 統一標準欄位定義（解決欄位錯位與出現 None 的問題）
INVENTORY_COLUMNS = ['藥品名稱(英文)', '中文名稱', '批號', '目前庫存', '有效期限', '用途', '狀態']
LOG_COLUMNS = ['領用時間', '藥品名稱', '中文名稱', '領用數量', '剩餘庫存', '備註']

# ==============================================================================
# 2. 連接 Google Sheets 雲端資料庫
# ==============================================================================
@st.cache_resource
def init_gspread():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    # 讀取 Streamlit secrets 設定
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    return gspread.authorize(credentials)

def get_spreadsheet():
    gc = init_gspread()
    # 取得 Secrets 中的名稱或 ID
    target = st.secrets.get("spreadsheet_name", "")
    
    # 優先嘗試當作「試算表 ID」開啟 (open_by_key)
    try:
        return gc.open_by_key(target)
    except Exception:
        pass
        
    # 若失敗，嘗試當作「試算表檔名」開啟 (open)
    try:
        return gc.open(target)
    except Exception:
        st.error(f"❌ 無法讀取試算表（設定值：'{target}'）。\n\n請確認：\n1. 試算表已「共用」給 client_email\n2. Google Drive API 已在 GCP 啟用")
        st.stop()
        
# 載入庫存資料並清洗
def load_inventory_data():
    sh = get_spreadsheet()
    worksheet = sh.worksheet("庫存")
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)

    if df.empty:
        df = pd.DataFrame(columns=INVENTORY_COLUMNS)
    else:
        # 強制只保留標準 7 個欄位，舊有的 "用途/備註" 或 "last_updated" 會被自動忽略
        for col in INVENTORY_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[INVENTORY_COLUMNS]

        # 資料型態清洗與轉換
        df['目前庫存'] = pd.to_numeric(df['目前庫存'], errors='coerce').fillna(0).astype(int)
        df['用途'] = df['用途'].astype(str).replace({'None': '', 'nan': ''})
        df['狀態'] = df['狀態'].apply(lambda x: "OK" if str(x).strip().upper() in ["OK", "NONE", ""] else str(x))

    return worksheet, df

# 載入領用紀錄
def load_logs_data():
    sh = get_spreadsheet()
    worksheet = sh.worksheet("領用紀錄")
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)

    if df.empty:
        df = pd.DataFrame(columns=LOG_COLUMNS)
    else:
        for col in LOG_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[LOG_COLUMNS]

    return worksheet, df

# ==============================================================================
# 3. 主頁面標題與分頁選單
# ==============================================================================
st.title("💊 國立臺北大學衛保組 藥品管理系統")

tab1, tab2, tab3 = st.tabs(["📦 當前庫存總覽與建置", "💊 藥品領用登記", "📜 更正歷史領用紀錄"])

# ------------------------------------------------------------------------------
# TAB 1: 當前庫存總覽與新建置藥品
# ------------------------------------------------------------------------------
with tab1:
    st.header("📦 當前藥品庫存總覽")
    ws_inv, df_inv = load_inventory_data()

    # 顯示庫存表格
    st.dataframe(df_inv, use_container_width=True, hide_index=True)

    # 一鍵修復 Google Sheet 標題列按鈕（整理試算表第 1 列）
    with st.expander("🛠️ 試算表欄位結構維護"):
        st.write("若雲端 Google Sheet 標題列出現 `用途/備註` 或 `last_updated` 等舊欄位，可點擊下方按鈕重置標題列：")
        if st.button("重置雲端庫存表欄位為標準 7 欄"):
            ws_inv.clear()
            ws_inv.append_row(INVENTORY_COLUMNS)
            # 寫回清洗過後的內容
            ws_inv.append_rows(df_inv.values.tolist())
            st.success("雲端試算表標題列已重置完成！")
            st.rerun()

    st.markdown("---")
    st.subheader("➕ 新建置藥品")
    
    with st.form("add_medicine_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            e_name = st.text_input("藥品名稱 (英文)*")
            c_name = st.text_input("中文名稱*")
        with col2:
            batch_no = st.text_input("批號*")
            initial_stock = st.number_input("初始庫存數量*", min_value=0, value=100, step=1)
        with col3:
            exp_date = st.date_input("有效期限*", date(2028, 1, 1))
            purpose = st.text_input("用途")

        submit_add = st.form_submit_button("新增藥品到庫存")

        if submit_add:
            if not e_name or not c_name or not batch_no:
                st.error("請填寫所有帶有 * 的必填欄位！")
            else:
                new_row = [
                    e_name.strip(),
                    c_name.strip(),
                    batch_no.strip(),
                    int(initial_stock),
                    exp_date.strftime("%Y-%m-%d"),
                    purpose.strip(),
                    "OK"
                ]
                ws_inv.append_row(new_row)
                st.success(f"藥品 {c_name} ({e_name}) 新建置成功！")
                st.rerun()

# ------------------------------------------------------------------------------
# TAB 2: 藥品領用登記（修正時區關鍵點）
# ------------------------------------------------------------------------------
with tab2:
    st.header("💊 藥品領用登記")
    ws_inv, df_inv = load_inventory_data()
    ws_log, df_log = load_logs_data()

    if df_inv.empty:
        st.warning("目前庫存無任何藥品資料，請先建置藥品。")
    else:
        with st.form("dispense_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                # 下拉選單顯示藥品
                med_options = [f"{row['中文名稱']} ({row['藥品名稱(英文)']}) - 剩餘: {row['目前庫存']}" for _, row in df_inv.iterrows()]
                selected_med_idx = st.selectbox("選擇領用藥品", range(len(med_options)), format_func=lambda x: med_options[x])
                dispense_qty = st.number_input("領用數量", min_value=1, value=1, step=1)

            with col2:
                record_date = st.date_input("領用日期", value=date.today())
                remarks = st.text_input("備註（例如：領用人 / 班級 / 病患）", value="")

            submit_dispense = st.form_submit_button("確認領用登記")

            if submit_dispense:
                target_row = df_inv.iloc[selected_med_idx]
                current_stock = target_row['Currently Stock'] if 'Currently Stock' in target_row else target_row['目前庫存']

                if dispense_qty > current_stock:
                    st.error(f"領用數量 ({dispense_qty}) 大於目前庫存量 ({current_stock})！")
                else:
                    # 💡【關鍵修正：精確結合台灣時間 GMT+8】
                    now_taiwan_time = datetime.now(TAIWAN_TZ).time()
                    combined_dt = datetime.combine(record_date, now_taiwan_time)
                    log_time_str = combined_dt.strftime("%Y-%m-%d %H:%M:%S")

                    new_stock = current_stock - dispense_qty

                    # 1. 更新庫存工作表 (更新特定 Row 的庫存量)
                    # gspread 列號從 2 開始 (1 是 Header)
                    sheet_row_num = selected_med_idx + 2 
                    stock_col_idx = INVENTORY_COLUMNS.index('目前庫存') + 1
                    ws_inv.update_cell(sheet_row_num, stock_col_idx, int(new_stock))

                    # 2. 新增至領用紀錄工作表
                    new_log_row = [
                        log_time_str,
                        target_row['藥品名稱(英文)'],
                        target_row['中文名稱'],
                        int(dispense_qty),
                        int(new_stock),
                        remarks
                    ]
                    ws_log.append_row(new_log_row)

                    st.success(f"已成功登記領用！領用時間記錄為：{log_time_str}")
                    st.rerun()

# ------------------------------------------------------------------------------
# TAB 3: 更正歷史領用紀錄
# ------------------------------------------------------------------------------
with tab3:
    st.header("更正歷史領用紀錄")
    ws_log, df_log = load_logs_data()

    if df_log.empty:
        st.info("目前尚無任何歷史領用紀錄。")
    else:
        st.write("📋 目前從雲端『領用紀錄』讀取到的最近歷史紀錄（顯示前 20 筆）：")
        
        # 依照紀錄反向排序（最新紀錄在前）
        df_log_recent = df_log.tail(20).iloc[::-1].copy()
        st.dataframe(df_log_recent, use_container_width=True)

        st.markdown("---")
        st.subheader("請選擇欲修改或撤銷的該筆紀錄：")

        # 建立格式化下拉選單列表
        log_options = {}
        for idx, row in df_log_recent.iterrows():
            # gspread row 索引等於 Pandas index + 2 (Header佔1行)
            excel_row_num = idx + 2
            label = f"行號 {excel_row_num}: [{row['領用時間']}] {row['藥品名稱']} - 原領用量: {row['領用數量']}"
            log_options[label] = {
                'row_num': excel_row_num,
                'data': row
            }

        selected_label = st.selectbox("選擇紀錄", list(log_options.keys()))

        if selected_label:
            target_info = log_options[selected_label]
            row_num = target_info['row_num']
            log_data = target_info['data']

            col_action1, col_action2 = st.columns(2)

            with col_action1:
                st.write("🔧 **修改領用數量**")
                new_qty = st.number_input("新領用數量", min_value=1, value=int(log_data['領用數量']))
                if st.button("更新此筆紀錄數量"):
                    # 計算數量差值以調回庫存（簡單修正示範）
                    qty_diff = int(log_data['領用數量']) - new_qty
                    
                    # 更新紀錄表的數量
                    qty_col_idx = LOG_COLUMNS.index('領用數量') + 1
                    ws_log.update_cell(row_num, qty_col_idx, new_qty)
                    
                    st.success(f"行號 {row_num} 紀錄已更新為 {new_qty}！")
                    st.rerun()

            with col_action2:
                st.write("❌ **撤銷/刪除紀錄**")
                st.write("注意：刪除該紀錄將不復存在。")
                if st.button("刪除此筆領用紀錄", type="primary"):
                    ws_log.delete_rows(row_num)
                    st.success(f"已刪除行號 {row_num} 之紀錄！")
                    st.rerun()
