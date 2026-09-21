import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import zoneinfo

# 頁面基本設定
st.set_page_config(
    page_title="國立臺北大學衛保組 藥品管理系統",
    page_icon="💊",
    layout="wide"
)

# 系統定義標準欄位
STANDARD_COLUMNS = [
    "藥品名稱(英文)", "中文名稱", "批號", "目前庫存", "有效期限", "用途", "狀態"
]
LOG_COLUMNS = [
    "紀錄時間", "藥品名稱(英文)", "中文名稱", "批號", "領用數量", "領用人/用途", "操作類型", "備註"
]

# 取得台北時間
def get_taipei_now():
    tz = zoneinfo.ZoneInfo("Asia/Taipei")
    return datetime.now(tz)

# 1. 初始化 gspread 連線 (快取機制)
@st.cache_resource
def init_gspread():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # 彈性存取 Secrets
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    elif "type" in st.secrets and st.secrets["type"] == "service_account":
        creds_dict = dict(st.secrets)
    else:
        st.error("❌ 尚未在 Streamlit Secrets 中設定 `[gcp_service_account]`！")
        st.stop()

    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(credentials)

# 2. 開啟 Google 試算表 (支援 ID 與 檔名雙重開啟)
def get_spreadsheet():
    gc = init_gspread()
    target = st.secrets.get("spreadsheet_name", "").strip()
    
    if not target:
        st.error("❌ 未在 Secrets 中設定 `spreadsheet_name`！")
        st.stop()
        
    # 優先嘗試用 試算表 ID (open_by_key) 開啟
    try:
        return gc.open_by_key(target)
    except Exception:
        pass

    # 若失敗，嘗試用 試算表檔名 (open) 開啟
    try:
        return gc.open(target)
    except Exception:
        email = st.secrets.get("gcp_service_account", {}).get("client_email", "您的機器人 Email")
        st.error(
            f"❌ 無法存取試算表（設定值：`{target}`）。\n\n"
            f"請確認以下事項：\n"
            f"1. 已將 Google 試算表「共用」給：`{email}` (權限：編輯者)\n"
            f"2. Secrets 中的 `spreadsheet_name` 為純 ID (不含 `/edit` 或 `#gid=`) 或精確檔名\n"
            f"3. Google Cloud Console 已啟用 Google Drive API 與 Google Sheets API"
        )
        st.stop()

# 3. 讀取庫存資料
@st.cache_data(ttl=5)
def load_inventory_data():
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet("庫存總覽")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="庫存總覽", rows=100, cols=10)
        ws.append_row(STANDARD_COLUMNS)
        
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    
    # 動態補充缺失的標準欄位
    for col in STANDARD_COLUMNS:
        if col not in df.columns:
            df[col] = ""
            
    df = df[STANDARD_COLUMNS]
    df["目前庫存"] = pd.to_numeric(df["目前庫存"], errors="coerce").fillna(0).astype(int)
    return ws, df

# 4. 讀取領用異動紀錄
@st.cache_data(ttl=5)
def load_logs_data():
    sh = get_spreadsheet()
    try:
        ws = sh.worksheet("異動紀錄")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="異動紀錄", rows=100, cols=10)
        ws.append_row(LOG_COLUMNS)
        
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    for col in LOG_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return ws, df[LOG_COLUMNS]

# ----------------- 主介面選單 -----------------
st.title("💊 國立臺北大學衛保組 藥品管理系統")

tab1, tab2, tab3 = st.tabs([
    "📦 當前庫存總覽與建置", 
    "✍️ 藥品領用登記", 
    "📜 更正歷史領用紀錄"
])

# ==================== Tab 1: 當前庫存總覽與建置 ====================
with tab1:
    st.subheader("📦 當前藥品庫存總覽")
    ws_inv, df_inv = load_inventory_data()

    if df_inv.empty:
        st.info("目前庫存表中無資料。")
    else:
        # 計算過期與預警狀態
        today_str = get_taipei_now().strftime("%Y-%m-%d")
        
        def update_status(row):
            exp = str(row["有效期限"]).strip()
            if not exp or exp == "None":
                return "正常"
            try:
                exp_date = datetime.strptime(exp, "%Y-%m-%d").date()
                today_date = date.today()
                days_left = (exp_date - today_date).days
                if days_left < 0:
                    return "⚠️ 已過期"
                elif days_left <= 90:
                    return "⚡ 即將到期"
                else:
                    return "正常"
            except ValueError:
                return "格式錯誤"

        df_inv["狀態"] = df_inv.apply(update_status, axis=1)
        
        # 庫存過低或過期警告
        expired_count = (df_inv["狀態"] == "⚠️ 已過期").sum()
        warning_count = (df_inv["狀態"] == "⚡ 即將到期").sum()
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("藥品品項總數", len(df_inv))
        col_m2.metric("即將到期品項 (90天內)", warning_count, delta_color="inverse")
        col_m3.metric("已過期品項", expired_count, delta_color="inverse")

        st.dataframe(df_inv, use_container_width=True)

    st.markdown("---")
    
    # 新增藥品品項
    with st.expander("➕ 新增藥品 / 新增批號登記"):
        with st.form("add_drug_form", clear_on_submit=True):
            col_a1, col_a2 = st.columns(2)
            eng_name = col_a1.text_input("藥品名稱 (英文) *").strip()
            chn_name = col_a2.text_input("中文名稱").strip()
            
            col_b1, col_b2, col_b3 = st.columns(3)
            batch_no = col_b1.text_input("批號 *").strip()
            qty = col_b2.number_input("初始庫存數量 *", min_value=1, value=1, step=1)
            exp_date_val = col_b3.date_input("有效期限", value=date.today())
            
            usage_note = st.text_input("用途說明").strip()
            
            submitted = st.form_submit_button("確認新增藥品")
            if submitted:
                if not eng_name or not batch_no:
                    st.error("請務必填寫「英文名稱」與「批號」！")
                else:
                    exp_str = exp_date_val.strftime("%Y-%m-%d")
                    new_row = [eng_name, chn_name, batch_no, int(qty), exp_str, usage_note, "正常"]
                    
                    ws_inv.append_row(new_row)
                    
                    # 寫入異動紀錄
                    ws_log, _ = load_logs_data()
                    now_str = get_taipei_now().strftime("%Y-%m-%d %H:%M:%S")
                    ws_log.append_row([now_str, eng_name, chn_name, batch_no, qty, "建置入庫", "新增藥品", usage_note])
                    
                    st.cache_data.clear()
                    st.success(f"成功新增藥品：{eng_name} ({batch_no})，數量：{qty}")
                    st.rerun()

    # 智慧維護工具
    with st.expander("🛠️ 系統維護工具"):
        st.write("若試算表欄位順序錯亂或標頭遺失，可使用以下按鈕重新修復與對齊第一行標頭。")
        if st.button("🔄 智慧校正並修復試算表標頭"):
            ws_inv.update("A1:G1", [STANDARD_COLUMNS])
            ws_log, _ = load_logs_data()
            ws_log.update("A1:H1", [LOG_COLUMNS])
            st.cache_data.clear()
            st.success("表頭欄位已修復對齊為標準格式！")
            st.rerun()

# ==================== Tab 2: 藥品領用登記 ====================
with tab2:
    st.subheader("✍️ 藥品領用登記")
    ws_inv, df_inv = load_inventory_data()

    if df_inv.empty:
        st.warning("目前庫存表中無可領用的藥品資料。")
    else:
        # 過濾出有庫存的藥品選項
        df_available = df_inv[df_inv["目前庫存"] > 0].copy()
        
        if df_available.empty:
            st.warning("當前所有藥品庫存皆為 0，請先進行藥品建置補貨。")
        else:
            # 建立選單格式: 英文名稱 | 中文名稱 | 批號 (剩餘: X)
            df_available["select_label"] = df_available.apply(
                lambda r: f"{r['藥品名稱(英文)']} ({r['中文名稱']}) - 批號:{r['批號']} [剩餘: {r['目前庫存']}]", axis=1
            )
            
            selected_label = st.selectbox("選擇領用藥品品項 *", df_available["select_label"].tolist())
            selected_row = df_available[df_available["select_label"] == selected_label].iloc[0]
            
            current_stock = int(selected_row["目前庫存"])
            eng_name = selected_row["藥品名稱(英文)"]
            chn_name = selected_row["中文名稱"]
            batch_no = selected_row["批號"]
            
            with st.form("dispense_form", clear_on_submit=True):
                col_d1, col_d2 = st.columns(2)
                dispense_qty = col_d1.number_input("領用數量 *", min_value=1, max_value=current_stock, value=1, step=1)
                recipient = col_d2.text_input("領用人 / 用途說明 *").strip()
                log_remark = st.text_input("備註 (可填寫病歷號或細項)").strip()
                
                btn_dispense = st.form_submit_button("確認扣庫並登記領用")
                
                if btn_dispense:
                    if not recipient:
                        st.error("請填寫領用人或用途說明！")
                    else:
                        # 搜尋試算表中對應的列數 (Row index)
                        cell = ws_inv.find(batch_no, in_column=3) # 第3欄為批號
                        if cell:
                            row_idx = cell.row
                            new_stock = current_stock - int(dispense_qty)
                            
                            # 更新庫存欄 (第 4 欄)
                            ws_inv.update_cell(row_idx, 4, new_stock)
                            
                            # 寫入異動紀錄
                            ws_log, _ = load_logs_data()
                            now_str = get_taipei_now().strftime("%Y-%m-%d %H:%M:%S")
                            ws_log.append_row([
                                now_str, eng_name, chn_name, batch_no, int(dispense_qty), recipient, "領用扣庫", log_remark
                            ])
                            
                            st.cache_data.clear()
                            st.success(f"領用登記成功！{eng_name} 扣減 {dispense_qty}，剩餘庫存：{new_stock}")
                            st.rerun()
                        else:
                            st.error("更新失敗：找不到該藥品批號所在的列。")

# ==================== Tab 3: 更正歷史領用紀錄 ====================
with tab3:
    st.subheader("📜 歷史領用異動紀錄與更正")
    ws_log, df_log = load_logs_data()
    
    if df_log.empty:
        st.info("目前尚無任何領用或異動紀錄。")
    else:
        # 反轉順序，最新紀錄顯示在最上面
        df_display = df_log.iloc[::-1].reset_index(drop=True)
        st.dataframe(df_display, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🔄 紀錄沖銷 / 錯誤更正")
        st.caption("如先前領用登記數量或品項填寫錯誤，可在下方選擇紀錄進行「庫存回補與沖銷紀錄」。")
        
        # 篩選出可更正的紀錄 (領用扣庫)
        dispense_logs = df_log[df_log["操作類型"] == "領用扣庫"].copy()
        
        if not dispense_logs.empty:
            dispense_logs["revert_label"] = dispense_logs.apply(
                lambda r: f"[{r['紀錄時間']}] {r['藥品名稱(英文)']} - 批號:{r['批號']} - 領用:{r['領用數量']} (領用人:{r['領用人/用途']})", axis=1
            )
            
            selected_revert_label = st.selectbox("選擇欲更正沖銷的歷史紀錄", dispense_logs["revert_label"].tolist())
            revert_item = dispense_logs[dispense_logs["revert_label"] == selected_revert_label].iloc[0]
            
            revert_reason = st.text_input("請輸入更正/沖銷原因 (必填)").strip()
            
            if st.button("確認沖銷此紀錄並歸還庫存"):
                if not revert_reason:
                    st.error("請務必填寫更正/沖銷原因！")
                else:
                    target_batch = revert_item["批號"]
                    qty_to_restore = int(revert_item["領用數量"])
                    eng_n = revert_item["藥品名稱(英文)"]
                    chn_n = revert_item["中文名稱"]
                    
                    # 找到庫存表位置並歸還庫存
                    ws_inv, df_inv = load_inventory_data()
                    cell = ws_inv.find(target_batch, in_column=3)
                    
                    if cell:
                        row_idx = cell.row
                        curr_qty = int(ws_inv.cell(row_idx, 4).value or 0)
                        updated_qty = curr_qty + qty_to_restore
                        
                        # 加回庫存
                        ws_inv.update_cell(row_idx, 4, updated_qty)
                        
                        # 寫入沖銷紀錄
                        now_str = get_taipei_now().strftime("%Y-%m-%d %H:%M:%S")
                        ws_log.append_row([
                            now_str, eng_n, chn_n, target_batch, qty_to_restore, "系統沖銷歸還", "錯誤更正", f"沖銷原紀錄({revert_item['紀錄時間']}) - 原因: {revert_reason}"
                        ])
                        
                        st.cache_data.clear()
                        st.success(f"已成功沖銷！{eng_n} (批號:{target_batch}) 已歸還庫存 {qty_to_restore}，目前總庫存為：{updated_qty}")
                        st.rerun()
                    else:
                        st.error("歸還失敗：在當前庫存表中找不到對應的批號。")
