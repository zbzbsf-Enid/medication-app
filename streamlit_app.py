import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime, timedelta, date
import calendar
import io
import xlsxwriter

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

# 取得原生 gspread Spreadsheet 物件
def get_spreadsheet():
    try:
        url = st.secrets["connections"]["gsheets"]["spreadsheet"]
        client = getattr(conn.client, "_client", conn.client)
        return client.open_by_url(url)
    except Exception as e:
        st.error(f"❌ 開啟雲端試算表失敗: {e}")
        return None

# 🚀 【關鍵修正】對資料讀取加上快取 (Cache)，60 秒內不重複發送請求給 Google API
@st.cache_data(ttl=60, show_spinner=False)
def load_sheet_data(worksheet_name: str, expected_cols: list = None) -> pd.DataFrame:
    sh = get_spreadsheet()
    if expected_cols is None:
        expected_cols = []
        
    if not sh:
        return pd.DataFrame(columns=expected_cols)
        
    try:
        ws = sh.worksheet(worksheet_name)
        data = ws.get_all_records()
        df = pd.DataFrame(data)
    except Exception:
        df = pd.DataFrame(columns=expected_cols)

    # 自動補齊缺少的欄位
    for col in expected_cols:
        if col not in df.columns:
            if "數量" in col or "庫存" in col:
                df[col] = 0
            else:
                df[col] = ""
                
    return df

# 🚀 【關鍵修正】對試算表分頁清單加上快取 (5 分鐘)
@st.cache_data(ttl=300, show_spinner=False)
def get_existing_sheets():
    sh = get_spreadsheet()
    if sh:
        try:
            return [ws.title for ws in sh.worksheets()]
        except Exception:
            return []
    return []

# 原生 gspread 寫入邏輯
def safe_update_sheet(worksheet_name: str, df: pd.DataFrame, default_headers: list = None):
    sh = get_spreadsheet()
    if not sh:
        return False, "無法連接至 Google 試算表"
    
    try:
        try:
            ws = sh.worksheet(worksheet_name)
        except Exception:
            ws = sh.add_worksheet(title=worksheet_name, rows="1000", cols="30")
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
            
        # 🚀 寫入成功後自動清空快取，確保下次讀取到最新資料
        st.cache_data.clear()
        return True, "成功寫入雲端"
    except Exception as err:
        return False, str(err)

# ---------------------------------------------------------
# 效期檢查輔助函式
# ---------------------------------------------------------
def parse_exp_date(exp_str):
    """解析日期字串，支援 YYYY-MM-DD 與 YYYY-MM 格式"""
    if not exp_str or str(exp_str).strip() in ["", "nan", "None", "NaT"]:
        return None
    clean_str = str(exp_str).strip().replace('/', '-')
    try:
        if len(clean_str) == 7:
            parts = clean_str.split('-')
            y, m = int(parts[0]), int(parts[1])
            _, last_day = calendar.monthrange(y, m)
            return date(y, m, last_day)
        elif len(clean_str) >= 10:
            clean_str = clean_str[:10]
            return datetime.strptime(clean_str, "%Y-%m-%d").date()
    except Exception:
        return None
    return None

def get_expiration_status(exp_str):
    """回傳 (狀態, 解析後的日期物件) : status 可為 'EXPIRED', 'WARNING', 'NORMAL'"""
    exp_date = parse_exp_date(exp_str)
    if not exp_date:
        return "NORMAL", None
    
    today = datetime.now().date()
    warning_limit = today + timedelta(days=30) # 30 天內（一個月）預警
    
    if exp_date < today:
        return "EXPIRED", exp_date
    elif today <= exp_date <= warning_limit:
        return "WARNING", exp_date
    else:
        return "NORMAL", exp_date

# 預設標準欄位定義
INV_COLS = ["藥品名稱", "中文名稱", "現有庫存", "批號", "有效期限", "備註"]
USAGE_COLS = ["領用時間", "藥品名稱", "中文名稱", "批號", "領用數量", "領用類別", "備註"]
RESTOCK_COLS = ["進貨時間", "藥品名稱", "中文名稱", "批號", "有效期限", "進貨數量", "備註"]

# ---------------------------------------------------------
# 2. 左側邊欄功能選單與效期即時預警
# ---------------------------------------------------------
st.sidebar.title("💊 衛保組藥品系統")
st.sidebar.markdown("---")

page = st.sidebar.radio(
    "📌 請選擇功能：",
    [
        "📋 藥品領用登記", 
        "📦 藥品庫存清單(可編修庫存/批號/效期)", 
        "🚚 進貨登記", 
        "📜 歷史紀錄(修改/刪除/同步庫存)", 
        "📊 月報表下載"
    ],
    index=0
)

st.sidebar.markdown("---")

# 🚨 側邊欄全域藥品效期監控視窗
inv_alert_df = load_sheet_data("庫存", INV_COLS)
expired_items = []
warning_items = []

if not inv_alert_df.empty:
    for _, r in inv_alert_df.iterrows():
        m_name = str(r.get("藥品名稱", "")).strip()
        if not m_name: continue
        z_name = str(r.get("中文名稱", "")).strip()
        disp_name = f"{m_name} ({z_name})" if z_name else m_name
        exp_s = r.get("有效期限", "")
        st_code, exp_d = get_expiration_status(exp_s)
        
        if st_code == "EXPIRED":
            expired_items.append(f"❌ {disp_name} (過期日: {exp_d})")
        elif st_code == "WARNING":
            warning_items.append(f"⚠️ {disp_name} (到期日: {exp_d})")

if expired_items or warning_items:
    with st.sidebar.expander("🚨 藥品效期警示總覽", expanded=True):
        if expired_items:
            st.error(f"**已過期藥品 ({len(expired_items)} 項，已鎖定領用)：**\n" + "\n".join([f"- {i}" for i in expired_items]))
        if warning_items:
            st.warning(f"**一個月內即將到期 ({len(warning_items)} 項)：**\n" + "\n".join([f"- {i}" for i in warning_items]))

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 系統設定")

existing_sheets = get_existing_sheets()
if existing_sheets:
    st.sidebar.success("✅ Google Sheet 連線正常")
    st.sidebar.write("🔍 目前雲端分頁：", existing_sheets)

if st.sidebar.button("🔄 手動刷新雲端資料", use_container_width=True):
    st.cache_data.clear()
    st.sidebar.success("已刷新資料快取！")
    st.rerun()

# 初始化領用購物車
if "claim_cart" not in st.session_state:
    st.session_state.claim_cart = []

# ---------------------------------------------------------
# 3. 主頁面內容控制
# ---------------------------------------------------------

# --- 頁面 1: 藥品領用登記 ---
if page == "📋 藥品領用登記":
    st.title("📋 藥品領用登記")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    
    if inventory_df.empty or inventory_df["藥品名稱"].str.strip().eq("").all():
        st.warning("⚠️ 目前「庫存」清單中無藥品資料，請先至「進貨登記」或「藥品庫存清單」新增藥品。")
    else:
        target_col = "現有庫存" if "現有庫存" in inventory_df.columns else ("剩餘庫存" if "剩餘庫存" in inventory_df.columns else "現有庫存")
        
        st.subheader("1. 搜尋與選擇藥品")
        med_options = []
        for _, row in inventory_df.iterrows():
            m_name = str(row.get("藥品名稱", "")).strip()
            if not m_name: continue
            z_name = str(row.get("中文名稱", "")).strip()
            batch_no = str(row.get("批號", "")).strip()
            exp_str = str(row.get("有效期限", "")).strip()
            stock_val = row.get(target_col, 0)
            
            # 計算效期標籤
            status_code, exp_d = get_expiration_status(exp_str)
            exp_tag = ""
            if status_code == "EXPIRED":
                exp_tag = f" ❌ [已過期 {exp_d}]"
            elif status_code == "WARNING":
                exp_tag = f" ⚠️ [快到期 {exp_d}]"
            elif exp_str:
                exp_tag = f" (效期: {exp_str})"
                
            batch_str = f" | 批號: {batch_no}" if batch_no else ""
            med_options.append(f"{m_name} ({z_name}){exp_tag}{batch_str} | 目前庫存: {stock_val}")
            
        if med_options:
            col_sel1, col_sel2, col_sel3 = st.columns([3, 1, 1])
            with col_sel1:
                selected_med_str = st.selectbox("搜尋或下拉選擇藥品：", med_options, key="med_selectbox")
                
            selected_idx = med_options.index(selected_med_str)
            selected_row = inventory_df.iloc[selected_idx]
            sel_med_name = selected_row.get("藥品名稱", "")
            sel_zh_name = selected_row.get("中文名稱", "")
            sel_batch_no = selected_row.get("批號", "")
            sel_exp_str = str(selected_row.get("有效期限", "")).strip()
            
            curr_stock_num = pd.to_numeric(selected_row.get(target_col, 0), errors='coerce')
            curr_stock_num = 0 if pd.isna(curr_stock_num) else int(curr_stock_num)
            
            sel_status, sel_exp_d = get_expiration_status(sel_exp_str)
            
            with col_sel2:
                add_qty = st.number_input("輸入領用數量", min_value=1, value=1, step=1, key="add_qty_input")
                
            with col_sel3:
                st.write(" ")
                st.write(" ")
                if sel_status == "EXPIRED":
                    st.button("🚫 已過期(禁領)", disabled=True, use_container_width=True)
                else:
                    if st.button("➕ 加入領用清單", use_container_width=True):
                        existing_item = next((item for item in st.session_state.claim_cart if item["藥品名稱"] == sel_med_name), None)
                        if existing_item:
                            existing_item["領用數量"] += add_qty
                            st.toast(f"已更新 {sel_zh_name or sel_med_name} 的數量為 {existing_item['領用數量']} 個！")
                        else:
                            st.session_state.claim_cart.append({
                                "藥品名稱": sel_med_name,
                                "中文名稱": sel_zh_name,
                                "批號": sel_batch_no,
                                "目前庫存": curr_stock_num,
                                "領用數量": add_qty
                            })
                            st.toast(f"已將 {sel_zh_name or sel_med_name} 加入領用清單！")

            if sel_status == "EXPIRED":
                st.error(f"🛑 **警告：【{sel_zh_name or sel_med_name}】已於 {sel_exp_d} 到期！過期藥品禁止領用。**")
            elif sel_status == "WARNING":
                st.warning(f"⚠️ **提醒：【{sel_zh_name or sel_med_name}】將於 {sel_exp_d} 到期（剩餘一個月內），請注意優先領用！**")

        st.markdown("---")
        st.subheader("2. 本次領用清單預覽與填寫資訊")
        
        if st.session_state.claim_cart:
            cart_df = pd.DataFrame(st.session_state.claim_cart)
            edited_cart_df = st.data_editor(
                cart_df,
                column_config={
                    "藥品名稱": st.column_config.Column("藥品名稱", disabled=True),
                    "中文名稱": st.column_config.Column("中文名稱", disabled=True),
                    "批號": st.column_config.Column("批號", disabled=True),
                    "目前庫存": st.column_config.Column("目前庫存", disabled=True),
                    "領用數量": st.column_config.NumberColumn("領用數量", min_value=1, step=1)
                },
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                key="cart_data_editor"
            )
            
            col_info1, col_info2 = st.columns([1, 2])
            use_date = col_info1.date_input("領用日期", datetime.now())
            with col_info2:
                st.write(" ")
                st.write(" ")
                is_public_med = st.checkbox("🏥 勾選為「公藥領用」", value=False)
            
            claim_type = "公藥" if is_public_med else "個人"
            
            col_btn1, col_btn2 = st.columns([1, 4])
            if col_btn1.button("🗑️ 清空領用清單", use_container_width=True):
                st.session_state.claim_cart = []
                st.rerun()
                
            confirm_check = col_btn2.checkbox(f"✅ 我已仔細核對上述所有領用品項與數量，確認為【{claim_type}】領用無誤。")
            
            if st.button("🚀 確認無誤，寫入雲端並扣減庫存", type="primary", use_container_width=True):
                if not confirm_check:
                    st.error("⚠️ 請先勾選「我已仔細核對上述所有領用品項與數量，確認無誤。」")
                elif edited_cart_df.empty:
                    st.error("⚠️ 領用清單目前為空，請先加入藥品！")
                else:
                    usage_df = load_sheet_data("領用紀錄", USAGE_COLS)
                    new_rows = []
                    
                    for _, row in edited_cart_df.iterrows():
                        u_qty = int(row["領用數量"])
                        m_name = row["藥品名稱"]
                        z_name = row["中文名稱"]
                        b_no = row.get("批號", "")
                        
                        new_rows.append({
                            "領用時間": str(use_date),
                            "藥品名稱": m_name,
                            "中文名稱": z_name,
                            "批號": b_no,
                            "領用數量": u_qty,
                            "領用類別": claim_type,
                            "備註": ""
                        })
                        
                        m_idx = inventory_df[inventory_df["藥品名稱"] == m_name].index
                        if not m_idx.empty:
                            idx = m_idx[0]
                            c_qty = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                            c_qty = 0 if pd.isna(c_qty) else int(c_qty)
                            inventory_df.at[idx, target_col] = max(0, c_qty - u_qty)
                    
                    new_usage_df = pd.concat([usage_df, pd.DataFrame(new_rows)], ignore_index=True)
                    ok1, msg1 = safe_update_sheet("領用紀錄", new_usage_df, USAGE_COLS)
                    ok2, msg2 = safe_update_sheet("庫存", inventory_df)
                    
                    if ok1 and ok2:
                        st.success(f"🎉 成功登記 {len(new_rows)} 項【{claim_type}】藥品領用，庫存已同步扣減！")
                        st.session_state.claim_cart = []
                    else:
                        st.error(f"❌ 更新失敗: 領用紀錄({msg1}) / 庫存({msg2})")
        else:
            st.info("🛒 目前領用清單為空。請由上方選單選擇藥品後，點擊「➕ 加入領用清單」。")

# --- 頁面 2: 藥品庫存清單 ---
elif page == "📦 藥品庫存清單(可編修庫存/批號/效期)":
    st.title("📦 藥品庫存、批號與有效期限管理")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else ("剩餘庫存" if "剩餘庫存" in inventory_df.columns else "現有庫存")
    inventory_df[target_col] = pd.to_numeric(inventory_df[target_col], errors='coerce').fillna(0).astype(int)
    
    st.caption("💡 有效期限請輸入 `YYYY-MM-DD`（例如 `2026-10-31`）。系統會自動於到期前一個月發出預警並鎖定過期品。")
    
    edited_inv_df = st.data_editor(
        inventory_df,
        column_config={
            "藥品名稱": st.column_config.TextColumn("藥品名稱（英文名）", required=True),
            "中文名稱": st.column_config.TextColumn("中文名稱"),
            target_col: st.column_config.NumberColumn(f"庫存數量 ({target_col})", min_value=0, step=1, required=True),
            "批號": st.column_config.TextColumn("批號 (Batch No.)"),
            "有效期限": st.column_config.TextColumn("有效期限 (YYYY-MM-DD)", help="請填寫YYYY-MM-DD格式"),
            "備註": st.column_config.TextColumn("備註 / 單位")
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="inventory_data_editor"
    )
    
    if st.button("💾 儲存庫存、批號與有效期限變更至雲端", type="primary", use_container_width=True):
        ok, msg = safe_update_sheet("庫存", edited_inv_df, INV_COLS)
        if ok:
            st.success("✅ 藥品庫存與有效期限資料已成功更新至 Google Sheets！")
            st.rerun()
        else:
            st.error(f"❌ 儲存失敗：{msg}")

# --- 頁面 3: 進貨登記 ---
elif page == "🚚 進貨登記":
    st.title("🚚 藥品進貨登記")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else ("剩餘庫存" if "剩餘庫存" in inventory_df.columns else "現有庫存")
    
    with st.form("restock_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        med_list = inventory_df["藥品名稱"].dropna().astype(str).tolist() if not inventory_df.empty else []
        med_list = [m for m in med_list if m.strip()]
        med_list.insert(0, "+ 新增未在庫存的藥品")
        
        selected_option = col1.selectbox("選擇或新增藥品", med_list)
        if selected_option == "+ 新增未在庫存的藥品":
            med_name = col1.text_input("藥品英文名稱", "")
            zh_name = col2.text_input("藥品中文名稱", "")
            batch_no = col1.text_input("批號 (Batch No.)", "")
            exp_date_val = col2.date_input("有效期限", datetime.now().date() + timedelta(days=365))
        else:
            med_name = selected_option
            matched = inventory_df[inventory_df["藥品名稱"] == med_name]
            zh_name = matched.iloc[0]["中文名稱"] if not matched.empty and "中文名稱" in matched.columns else ""
            exist_batch = matched.iloc[0]["批號"] if not matched.empty and "批號" in matched.columns else ""
            exist_exp = matched.iloc[0]["有效期限"] if not matched.empty and "有效期限" in matched.columns else ""
            
            col2.text_input("藥品中文名稱", value=str(zh_name), disabled=True)
            batch_no = col1.text_input("批號 (Batch No.)", value=str(exist_batch))
            
            default_d = parse_exp_date(exist_exp) or (datetime.now().date() + timedelta(days=365))
            exp_date_val = col2.date_input("有效期限", value=default_d)
            
        restock_qty = col2.number_input("進貨數量", min_value=1, step=1, value=100)
        restock_date = col1.date_input("進貨日期", datetime.now())
        restock_remarks = col2.text_input("進貨備註 / 廠商資訊", "")
        
        if st.form_submit_button("確認進貨並更新雲端庫存", use_container_width=True):
            if not med_name.strip():
                st.error("請輸入藥品名稱！")
            else:
                exp_date_str = str(exp_date_val)
                restock_df = load_sheet_data("進貨紀錄", RESTOCK_COLS)
                new_restock_row = pd.DataFrame([{
                    "進貨時間": str(restock_date),
                    "藥品名稱": med_name,
                    "中文名稱": zh_name,
                    "批號": batch_no,
                    "有效期限": exp_date_str,
                    "進貨數量": restock_qty,
                    "備註": restock_remarks
                }])
                updated_restock_df = pd.concat([restock_df, new_restock_row], ignore_index=True)
                success_restock, msg1 = safe_update_sheet("進貨紀錄", updated_restock_df, RESTOCK_COLS)
                
                if success_restock:
                    if not inventory_df.empty and med_name in inventory_df["藥品名稱"].values:
                        idx = inventory_df[inventory_df["藥品名稱"] == med_name].index[0]
                        curr_qty = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                        curr_qty = 0 if pd.isna(curr_qty) else int(curr_qty)
                        inventory_df.at[idx, target_col] = curr_qty + restock_qty
                        if batch_no:
                            inventory_df.at[idx, "批號"] = batch_no
                        inventory_df.at[idx, "有效期限"] = exp_date_str
                    else:
                        new_inv_row = pd.DataFrame([{
                            "藥品名稱": med_name,
                            "中文名稱": zh_name,
                            target_col: restock_qty,
                            "批號": batch_no,
                            "有效期限": exp_date_str,
                            "備註": ""
                        }])
                        inventory_df = pd.concat([inventory_df, new_inv_row], ignore_index=True)
                        
                    success_inv, msg2 = safe_update_sheet("庫存", inventory_df)
                    if success_inv:
                        st.success(f"✅ 成功進貨：{zh_name or med_name} {restock_qty} 個！有效期限已設為 {exp_date_str}")
                    else:
                        st.error(f"❌ 庫存更新失敗：{msg2}")
                else:
                    st.error(f"❌ 進貨紀錄寫入失敗：{msg1}")

# --- 頁面 4: 歷史紀錄維護 ---
elif page == "📜 歷史紀錄(修改/刪除/同步庫存)":
    st.title("📜 歷史紀錄維護（具備庫存自動同步機制）")
    tab_rec1, tab_rec2 = st.tabs(["📋 領用紀錄編輯", "🚚 進貨紀錄編輯"])
    inventory_df = load_sheet_data("庫存", INV_COLS)
    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else ("剩餘庫存" if "剩餘庫存" in inventory_df.columns else "現有庫存")
    
    with tab_rec1:
        u_df = load_sheet_data("領用紀錄", USAGE_COLS)
        if not u_df.empty:
            edited_u_df = st.data_editor(u_df, num_rows="dynamic", use_container_width=True, key="usage_editor")
            if st.button("💾 儲存「領用紀錄」修改並自動同步庫存", type="primary"):
                u_df["領用數量"] = pd.to_numeric(u_df["領用數量"], errors='coerce').fillna(0)
                old_totals = u_df.groupby("藥品名稱")["領用數量"].sum().to_dict()
                
                if not edited_u_df.empty and "領用數量" in edited_u_df.columns and "藥品名稱" in edited_u_df.columns:
                    edited_u_df["領用數量"] = pd.to_numeric(edited_u_df["領用數量"], errors='coerce').fillna(0)
                    new_totals = edited_u_df.groupby("藥品名稱")["領用數量"].sum().to_dict()
                else:
                    new_totals = {}
                
                for med in set(old_totals.keys()).union(set(new_totals.keys())):
                    if not med: continue
                    diff = old_totals.get(med, 0) - new_totals.get(med, 0)
                    if diff != 0 and not inventory_df.empty and "藥品名稱" in inventory_df.columns:
                        m_idx = inventory_df[inventory_df["藥品名稱"] == med].index
                        if not m_idx.empty:
                            idx = m_idx[0]
                            curr_stock = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                            curr_stock = 0 if pd.isna(curr_stock) else int(curr_stock)
                            inventory_df.at[idx, target_col] = max(0, curr_stock + diff)
                
                ok1, msg1 = safe_update_sheet("領用紀錄", edited_u_df, USAGE_COLS)
                ok2, msg2 = safe_update_sheet("庫存", inventory_df)
                if ok1 and ok2:
                    st.success("✅ 領用紀錄已更新，庫存已精準同步加減！")
                    st.rerun()
        else:
            st.info("目前無領用紀錄。")
            
    with tab_rec2:
        r_df = load_sheet_data("進貨紀錄", RESTOCK_COLS)
        if not r_df.empty:
            edited_r_df = st.data_editor(r_df, num_rows="dynamic", use_container_width=True, key="restock_editor")
            if st.button("💾 儲存「進貨紀錄」修改並自動同步庫存", type="primary"):
                r_df["進貨數量"] = pd.to_numeric(r_df["進貨數量"], errors='coerce').fillna(0)
                old_restock = r_df.groupby("藥品名稱")["進貨數量"].sum().to_dict()
                
                if not edited_r_df.empty and "進貨數量" in edited_r_df.columns and "藥品名稱" in edited_r_df.columns:
                    edited_r_df["進貨數量"] = pd.to_numeric(edited_r_df["進貨數量"], errors='coerce').fillna(0)
                    new_restock = edited_r_df.groupby("藥品名稱")["進貨數量"].sum().to_dict()
                else:
                    new_restock = {}
                
                for med in set(old_restock.keys()).union(set(new_restock.keys())):
                    if not med: continue
                    diff = new_restock.get(med, 0) - old_restock.get(med, 0)
                    if diff != 0 and not inventory_df.empty and "藥品名稱" in inventory_df.columns:
                        m_idx = inventory_df[inventory_df["藥品名稱"] == med].index
                        if not m_idx.empty:
                            idx = m_idx[0]
                            curr_stock = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                            curr_stock = 0 if pd.isna(curr_stock) else int(curr_stock)
                            inventory_df.at[idx, target_col] = max(0, curr_stock + diff)
                            
                ok1, msg1 = safe_update_sheet("進貨紀錄", edited_r_df, RESTOCK_COLS)
                ok2, msg2 = safe_update_sheet("庫存", inventory_df)
                if ok1 and ok2:
                    st.success("✅ 進貨紀錄已更新，庫存已同步調整！")
                    st.rerun()
        else:
            st.info("目前無進貨紀錄。")

# --- 頁面 5: 月報表下載 ---
elif page == "📊 月報表下載":
    st.title("📊 藥品使用月報表與統計表繪出")
    st.write("產出格式符合國立臺北大學衛保組月報表標準格式。")
    
    col_y, col_m = st.columns(2)
    selected_year = col_y.number_input("選擇年份(西元)", min_value=2020, max_value=2030, value=datetime.now().year)
    selected_month = col_m.selectbox("選擇月份", list(range(1, 13)), index=datetime.now().month - 1)
    
    roc_year = selected_year - 1911
    acad_year = roc_year - 1 if selected_month < 8 else roc_year
    semester = "下學期" if 1 <= selected_month <= 7 else "上學期"
    
    prev_month = 12 if selected_month == 1 else selected_month - 1
    prev_year = selected_year - 1 if selected_month == 1 else selected_year
    prev_roc_year = prev_year - 1911
    
    if st.button("📥 產生並預覽月報表", type="primary", use_container_width=True):
        inventory_df = load_sheet_data("庫存", INV_COLS)
        usage_df = load_sheet_data("領用紀錄", USAGE_COLS)
        restock_df = load_sheet_data("進貨紀錄", RESTOCK_COLS)
        
        target_col = "現有庫存" if "現有庫存" in inventory_df.columns else ("剩餘庫存" if "剩餘庫存" in inventory_df.columns else "現有庫存")
        
        usage_df["領用時間"] = usage_df["領用時間"].astype(str)
        usage_df["藥品名稱"] = usage_df["藥品名稱"].astype(str)
        restock_df["進貨時間"] = restock_df["進貨時間"].astype(str)
        restock_df["藥品名稱"] = restock_df["藥品名稱"].astype(str)
        
        _, num_days = calendar.monthrange(selected_year, selected_month)
        
        report_rows = []
        
        for _, inv_row in inventory_df.iterrows():
            med_name = str(inv_row.get("藥品名稱", "")).strip()
            if not med_name:
                continue
                
            zh_name = str(inv_row.get("中文名稱", "")).strip()
            full_name = f"{med_name}({zh_name})" if zh_name else med_name
            curr_stock = pd.to_numeric(inv_row.get(target_col, 0), errors='coerce')
            curr_stock = 0 if pd.isna(curr_stock) else int(curr_stock)
            
            row_dict = {
                "藥品名稱\n(商品名/中文)": full_name,
                f"{prev_roc_year}年{prev_month}月\n剩餘量": curr_stock
            }
            
            daily_total = 0
            for d in range(1, num_days + 1):
                day_str = f"{selected_year}-{selected_month:02d}-{d:02d}"
                col_name = f"{selected_month}/{d}"
                
                matched_u = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["領用時間"] == day_str)]
                u_qty = pd.to_numeric(matched_u["領用數量"], errors='coerce').sum() if not matched_u.empty else 0
                    
                row_dict[col_name] = int(u_qty) if not pd.isna(u_qty) else 0
                daily_total += row_dict[col_name]
                
            m_start = f"{selected_year}-{selected_month:02d}-01"
            m_end = f"{selected_year}-{selected_month:02d}-{num_days:02d}"
            
            m_usage = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["領用時間"] >= m_start) & (usage_df["領用時間"] <= m_end)]
            public_qty = pd.to_numeric(m_usage[m_usage["領用類別"] == "公藥"]["領用數量"], errors='coerce').sum() if not m_usage.empty else 0
                
            matched_r = restock_df[(restock_df["藥品名稱"] == med_name) & (restock_df["進貨時間"] >= m_start) & (restock_df["進貨時間"] <= m_end)]
            restock_qty = pd.to_numeric(matched_r["進貨數量"], errors='coerce').sum() if not matched_r.empty else 0
                
            row_dict["當月使用\n總量"] = daily_total
            row_dict["購入量"] = int(restock_qty)
            row_dict["過期報銷"] = 0
            row_dict["公藥使用"] = int(public_qty)
            row_dict[f"{roc_year}年{selected_month}月\n期末剩餘量"] = curr_stock
            
            report_rows.append(row_dict)
            
        report_df = pd.DataFrame(report_rows)
        
        if report_df.empty:
            st.warning("⚠️ 目前沒有可呈現的藥品資料，請確認「庫存」分頁中是否有項目。")
        else:
            st.subheader("📋 報表線上預覽")
            st.dataframe(report_df, use_container_width=True, hide_index=True)
            
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('月報表')
            
            font_family = '微軟正黑體'
            
            fmt_title = workbook.add_format({
                'font_name': font_family, 'font_size': 14, 'bold': True, 'valign': 'vcenter'
            })
            
            fmt_header_navy = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
                'bg_color': '#1F4E78', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                'border': 1, 'border_color': '#BFBFBF'
            })
            
            fmt_header_blue = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#FFFFFF',
                'bg_color': '#2F5597', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                'border': 1, 'border_color': '#BFBFBF'
            })
            
            fmt_header_daily = workbook.add_format({
                'font_name': font_family, 'font_size': 9, 'bold': True, 'font_color': '#000000',
                'bg_color': '#F2F2F2', 'align': 'center', 'valign': 'vcenter',
                'border': 1, 'border_color': '#BFBFBF'
            })
            
            fmt_data_left = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'align': 'left', 'valign': 'vcenter',
                'border': 1, 'border_color': '#D9D9D9'
            })
            
            fmt_data_center = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
                'border': 1, 'border_color': '#D9D9D9'
            })
            
            fmt_data_red = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'bold': True, 'font_color': '#C00000',
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'
            })
            
            fmt_data_yellow = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'bold': True, 'bg_color': '#FFF2CC',
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'
            })
            
            fmt_data_blue = workbook.add_format({
                'font_name': font_family, 'font_size': 10, 'bg_color': '#D9E1F2',
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D9D9D9'
            })

            title_text = f"國立臺北大學衛保組 {acad_year}學年度{semester}藥品使用月報與全學期統計表 ({roc_year}學年度{selected_month}月起)"
            worksheet.set_row(0, 32)
            worksheet.write(0, 0, title_text, fmt_title)
            
            worksheet.set_row(1, 38)
            headers = list(report_df.columns)
            
            for col_idx, h in enumerate(headers):
                if col_idx in [0, 1, len(headers) - 1]:
                    worksheet.write(1, col_idx, h, fmt_header_navy)
                elif col_idx in [len(headers) - 5, len(headers) - 4, len(headers) - 3, len(headers) - 2]:
                    worksheet.write(1, col_idx, h, fmt_header_blue)
                else:
                    worksheet.write(1, col_idx, h, fmt_header_daily)

            for row_idx, row_data in enumerate(report_df.values):
                current_row = row_idx + 2
                worksheet.set_row(current_row, 22)
                
                for col_idx, val in enumerate(row_data):
                    if col_idx == 0:
                        worksheet.write(current_row, col_idx, str(val), fmt_data_left)
                    elif col_idx in [1, len(headers) - 1]:
                        worksheet.write(current_row, col_idx, int(val), fmt_data_red)
                    elif col_idx == len(headers) - 5:
                        worksheet.write(current_row, col_idx, int(val), fmt_data_yellow)
                    elif col_idx in [len(headers) - 4, len(headers) - 3, len(headers) - 2]:
                        worksheet.write(current_row, col_idx, int(val), fmt_data_blue)
                    else:
                        worksheet.write(current_row, col_idx, int(val), fmt_data_center)

            worksheet.set_column(0, 0, 38)
            worksheet.set_column(1, 1, 13)
            worksheet.set_column(2, len(headers) - 6, 5.5)
            worksheet.set_column(len(headers) - 5, len(headers) - 1, 13)
            
            workbook.close()
            excel_data = output.getvalue()
            
            file_name = f"國立臺北大學衛保組_{acad_year}學年度_{selected_month}月藥品使用月報表.xlsx"
            
            st.download_button(
                label="📥 點此下載標準彩色 Excel 月報表 (.xlsx)",
                data=excel_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
