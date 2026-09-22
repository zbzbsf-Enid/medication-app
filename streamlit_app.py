import streamlit as st
import pandas as pd
import io
import calendar
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# 頁面基本設定
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="國立臺北大學衛保組 藥品管理系統",
    page_icon="💊",
    layout="wide"
)

st.title("💊 國立臺北大學衛保組 藥品管理系統")

# -----------------------------------------------------------------------------
# 1. 雲端 Google Sheets 連線與動態頁籤偵測
# -----------------------------------------------------------------------------
def get_connection():
    return st.connection("gsheets", type=GSheetsConnection)

conn = get_connection()

@st.cache_data(ttl=60, show_spinner="讀取雲端資料庫中...")
def load_base_data():
    """讀取庫存與領用紀錄流水帳，並自動記錄真實分頁名稱"""
    df_inventory = None
    df_logs = None
    inv_sheet_name = "庫存"
    log_sheet_name = "領用紀錄"

    try:
        # 1. 偵測並讀取『庫存』工作表
        for inv_name in ["庫存", "Sheet1", "工作表1", "Inventory", None]:
            try:
                tmp = conn.read(worksheet=inv_name, ttl=60) if inv_name else conn.read(ttl=60)
                if tmp is not None and not tmp.empty:
                    cols_str = " ".join([str(c) for c in tmp.columns])
                    if any(k in cols_str for k in ['藥品', '品名', '名稱', '現有庫存', '剩餘量']):
                        df_inventory = tmp
                        inv_sheet_name = inv_name if inv_name else "Sheet1"
                        break
            except Exception:
                continue

        # 2. 偵測並讀取『領用紀錄』流水帳
        for log_name in ["領用紀錄", "Sheet2", "工作表2", "用藥紀錄", "紀錄", "Logs"]:
            try:
                tmp_log = conn.read(worksheet=log_name, ttl=60)
                if tmp_log is not None and not tmp_log.empty:
                    df_logs = tmp_log
                    log_sheet_name = log_name
                    break
            except Exception:
                continue

        if df_logs is None or df_logs.empty:
            df_logs = pd.DataFrame(columns=['日期', '藥品名稱', '批號', '領用數量', '備註'])

        # 庫存表欄位標準化
        if df_inventory is not None and not df_inventory.empty:
            col_map = {}
            for col in df_inventory.columns:
                c_str = str(col).strip()
                if any(k in c_str for k in ['藥品', '品名', '名稱']):
                    if '藥品名稱' not in col_map.values(): col_map[col] = '藥品名稱'
                elif any(k in c_str for k in ['批號', '批次']):
                    if '批號' not in col_map.values(): col_map[col] = '批號'
                elif any(k in c_str for k in ['效期', '有效日期', '有效期限', '到期日']):
                    if '有效日期' not in col_map.values(): col_map[col] = '有效日期'
                elif any(k in c_str for k in ['現有庫存', '目前庫存', '當前庫存', '剩餘量', '庫存', '剩餘']):
                    if '現有庫存' not in col_map.values(): col_map[col] = '現有庫存'

            df_inventory = df_inventory.rename(columns=col_map)
            if '有效日期' not in df_inventory.columns:
                df_inventory['有效日期'] = '2099-12-31'
            if '現有庫存' in df_inventory.columns:
                df_inventory['現有庫存'] = pd.to_numeric(df_inventory['現有庫存'], errors='coerce').fillna(0).astype(int)

        # 領用紀錄表欄位標準化
        if df_logs is not None and not df_logs.empty:
            log_col_map = {}
            for col in df_logs.columns:
                c_str = str(col).strip()
                if any(k in c_str for k in ['日期', '時間', 'Date']):
                    if '日期' not in log_col_map.values(): log_col_map[col] = '日期'
                elif any(k in c_str for k in ['藥品', '品名', '名稱']):
                    if '藥品名稱' not in log_col_map.values(): log_col_map[col] = '藥品名稱'
                elif any(k in c_str for k in ['批號', '批次']):
                    if '批號' not in log_col_map.values(): log_col_map[col] = '批號'
                elif any(k in c_str for k in ['數量', '領用', '扣減']):
                    if '領用數量' not in log_col_map.values(): log_col_map[col] = '領用數量'
                elif any(k in c_str for k in ['備註', '說明']):
                    if '備註' not in log_col_map.values(): log_col_map[col] = '備註'
            df_logs = df_logs.rename(columns=log_col_map)
            if '批號' not in df_logs.columns:
                df_logs['批號'] = ""

        return df_inventory, df_logs, inv_sheet_name, log_sheet_name
    except Exception as e:
        st.error(f"❌ 讀取雲端資料失敗：{e}")
        st.stop()

@st.cache_data(ttl=60, show_spinner="讀取試算表原檔中...")
def load_raw_monthly_sheet(sheet_name=None):
    """直接讀取試算表原檔分頁"""
    try:
        if sheet_name:
            df = conn.read(worksheet=sheet_name, ttl=60)
        else:
            df = conn.read(ttl=60)
        return df
    except Exception:
        return None

df_inventory, df_logs, inv_sheet_name, log_sheet_name = load_base_data()

# 安全寫入輔助函式
def safe_update_sheet(worksheet_target, data_df):
    """確保寫入雲端時格式乾淨無誤"""
    clean_df = data_df.fillna("").copy()
    for col in clean_df.columns:
        clean_df[col] = clean_df[col].astype(str)
    try:
        conn.update(worksheet=worksheet_target, data=clean_df)
        return True, "OK"
    except Exception as err:
        return False, str(err)

# -----------------------------------------------------------------------------
# 2. 先進先出 (FIFO) 庫存扣減邏輯
# -----------------------------------------------------------------------------
def deduct_inventory_fifo(inventory_df, drug_name, req_qty, log_date, note=""):
    df_inv = inventory_df.copy()
    mask = (df_inv['藥品名稱'] == drug_name) & (df_inv['現有庫存'] > 0)
    available = df_inv[mask].copy()

    if available.empty:
        return df_inv, [], f"❌ {drug_name} 目前已無可用庫存！"

    available['exp_dt'] = pd.to_datetime(available['有效日期'], errors='coerce')
    available = available.sort_values(by='exp_dt', ascending=True)

    remaining_qty = req_qty
    new_logs = []

    for idx in available.index:
        if remaining_qty <= 0:
            break

        current_stock = df_inv.loc[idx, '現有庫存']
        batch_no = df_inv.loc[idx, '批號']

        if current_stock >= remaining_qty:
            df_inv.loc[idx, '現有庫存'] = current_stock - remaining_qty
            deducted = remaining_qty
            remaining_qty = 0
        else:
            df_inv.loc[idx, '現有庫存'] = 0
            deducted = current_stock
            remaining_qty -= current_stock

        new_logs.append({
            '日期': str(log_date),
            '藥品名稱': drug_name,
            '批號': batch_no,
            '領用數量': int(deducted),
            '備註': note
        })

    status_msg = "SUCCESS"
    if remaining_qty > 0:
        status_msg = f"⚠️ {drug_name} 庫存不足，尚有 {remaining_qty} 顆未完成扣減。"

    return df_inv, new_logs, status_msg

# -----------------------------------------------------------------------------
# 3. 側邊欄與功能選單
# -----------------------------------------------------------------------------
st.sidebar.title("📌 功能選單")

if st.sidebar.button("🔄 手動刷新雲端資料"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

menu = st.sidebar.radio(
    "請選擇功能頁面",
    [
        "📋 多項藥品領用登記",
        "🏥 藥品進貨/建檔登記",
        "🛠️ 紀錄修改與庫存微調",
        "📊 當前庫存總覽",
        "🗓️ 官方月報表與學期統計 (全月份動態產生)"
    ]
)

# -----------------------------------------------------------------------------
# 頁面 1：多項藥品領用登記
# -----------------------------------------------------------------------------
if menu == "📋 多項藥品領用登記":
    st.header("📋 批量藥品領用登記")

    col1, col2 = st.columns(2)
    with col1:
        log_date = st.date_input("領用日期", value=datetime.now())
    with col2:
        note = st.text_input("領用備註", value="無")

    drug_list = sorted(df_inventory['藥品名稱'].dropna().unique().tolist()) if df_inventory is not None else []
    selected_drugs = st.multiselect("請選擇欲領用的藥品名稱", options=drug_list)

    qty_dict = {}
    if selected_drugs:
        st.subheader("領用數量設定")
        cols = st.columns(min(len(selected_drugs), 3))
        
        for i, drug in enumerate(selected_drugs):
            col = cols[i % 3]
            total_stock = df_inventory[df_inventory['藥品名稱'] == drug]['現有庫存'].sum()
            qty_dict[drug] = col.number_input(
                f"{drug} (總庫存: {total_stock})",
                min_value=1,
                value=1,
                step=1,
                key=f"qty_{drug}"
            )

        st.markdown("---")
        if st.button("確認無誤，寫入雲端並更新庫存", type="primary"):
            updated_inv = df_inventory.copy()
            all_new_logs = []
            warning_messages = []

            for drug, req_qty in qty_dict.items():
                updated_inv, logs, msg = deduct_inventory_fifo(
                    updated_inv, drug, req_qty, log_date, note
                )
                if logs:
                    all_new_logs.extend(logs)
                if msg != "SUCCESS":
                    warning_messages.append(msg)

            if warning_messages:
                for msg in warning_messages:
                    st.warning(msg)

            if all_new_logs:
                # 1. 更新庫存
                ok1, err1 = safe_update_sheet(inv_sheet_name, updated_inv)
                
                # 2. 更新領用紀錄
                logs_df_new = pd.concat([df_logs, pd.DataFrame(all_new_logs)], ignore_index=True)
                ok2, err2 = safe_update_sheet(log_sheet_name, logs_df_new)

                if ok1 and ok2:
                    st.success("✅ 庫存與領用紀錄更新成功！已同步至雲端。")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    if not ok1:
                        st.error(f"❌ 寫入庫存分頁 ({inv_sheet_name}) 失敗：{err1}")
                    if not ok2:
                        st.error(f"❌ 寫入領用紀錄分頁 ({log_sheet_name}) 失敗：{err2}")
                        st.info("💡 提示：請確認您的 Google Sheet 中是否有建立名為『領用紀錄』或『Sheet2』的分頁喔！")

# -----------------------------------------------------------------------------
# 頁面 2：藥品進貨/建檔登記
# -----------------------------------------------------------------------------
elif menu == "🏥 藥品進貨/建檔登記":
    st.header("🏥 藥品進貨與建檔登記")

    entry_type = st.radio("請選擇操作類型", ["既存藥品進貨 (加庫存/新批號)", "新增全新藥品品項"])

    if entry_type == "既存藥品進貨 (加庫存/新批號)":
        drug_list = sorted(df_inventory['藥品名稱'].dropna().unique().tolist())
        selected_drug = st.selectbox("選擇藥品名稱", options=drug_list)

        col1, col2, col3 = st.columns(3)
        with col1:
            batch_no = st.text_input("進貨批號")
        with col2:
            exp_date = st.date_input("有效日期")
        with col3:
            add_qty = st.number_input("進貨數量", min_value=1, value=100)

        if st.button("提交進貨紀錄", type="primary"):
            if not batch_no:
                st.warning("請填寫批號！")
            else:
                updated_inv = df_inventory.copy()
                mask = (updated_inv['藥品名稱'] == selected_drug) & (updated_inv['批號'].astype(str) == str(batch_no))
                if mask.any():
                    updated_inv.loc[mask, '現有庫存'] += add_qty
                    updated_inv.loc[mask, '有效日期'] = str(exp_date)
                else:
                    new_row = {col: "" for col in updated_inv.columns}
                    new_row['藥品名稱'] = selected_drug
                    new_row['批號'] = str(batch_no)
                    new_row['有效日期'] = str(exp_date)
                    new_row['現有庫存'] = add_qty
                    updated_inv = pd.concat([updated_inv, pd.DataFrame([new_row])], ignore_index=True)

                ok, err = safe_update_sheet(inv_sheet_name, updated_inv)
                if ok:
                    st.success(f"✅ 已成功為 {selected_drug} 新增庫存 {add_qty} 顆！")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"❌ 寫入失敗：{err}")

    else:
        col1, col2 = st.columns(2)
        with col1:
            new_drug_name = st.text_input("全新藥品名稱 (商品名/中文)")
            new_batch = st.text_input("批號")
        with col2:
            new_exp = st.date_input("有效日期")
            new_qty = st.number_input("初始庫存數量", min_value=1, value=100)

        if st.button("建立新藥品品項", type="primary"):
            if not new_drug_name or not new_batch:
                st.warning("請輸入藥品名稱與批號！")
            else:
                updated_inv = df_inventory.copy()
                new_row = {col: "" for col in updated_inv.columns}
                new_row['藥品名稱'] = new_drug_name
                new_row['批號'] = str(new_batch)
                new_row['有效日期'] = str(new_exp)
                new_row['現有庫存'] = new_qty
                updated_inv = pd.concat([updated_inv, pd.DataFrame([new_row])], ignore_index=True)

                ok, err = safe_update_sheet(inv_sheet_name, updated_inv)
                if ok:
                    st.success(f"✅ 已建立新藥品 {new_drug_name}！")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"❌ 寫入失敗：{err}")

# -----------------------------------------------------------------------------
# 頁面 3：紀錄修改與庫存微調
# -----------------------------------------------------------------------------
elif menu == "🛠️ 紀錄修改與庫存微調":
    st.header("🛠️ 庫存數量手動微調")
    st.info("💡 在此頁面可以直接修正盤點後的庫存數量。")

    edited_inv = st.data_editor(df_inventory, use_container_width=True, num_rows="dynamic")

    if st.button("儲存庫存微調變更", type="primary"):
        ok, err = safe_update_sheet(inv_sheet_name, edited_inv)
        if ok:
            st.success("✅ 庫存資料微調成功並同步至 Google Sheet！")
            st.cache_data.clear()
            st.rerun()
        else:
            st.error(f"❌ 儲存失敗：{err}")

# -----------------------------------------------------------------------------
# 頁面 4：當前庫存總覽
# -----------------------------------------------------------------------------
elif menu == "📊 當前庫存總覽":
    st.header("📊 當前庫存總覽")
    st.dataframe(df_inventory, use_container_width=True)

# -----------------------------------------------------------------------------
# 頁面 5：官方月報表與學期統計
# -----------------------------------------------------------------------------
elif menu == "🗓️ 官方月報表與學期統計 (全月份動態產生)":
    st.header("🗓️ 國立臺北大學校園門診藥品統計表")

    tab1, tab2 = st.tabs(["✨ 選取月份生成官方格式月報表 (可匯出 Excel)", "📄 雲端 Google Sheet 原檔備份"])

    with tab1:
        c1, c2, c3 = st.columns([1.5, 1.5, 3])
        with c1:
            roc_year = st.number_input("民國年份", min_value=110, max_value=130, value=115, step=1)
            ad_year = roc_year + 1911
        with c2:
            sel_month = st.selectbox("選擇月份", list(range(1, 13)), index=8) # 預設 9 月
        with c3:
            term_title = st.text_input("學期報表標題", value=f"{roc_year}-1 國立臺北大學校園門診藥品統計表")

        if sel_month == 1:
            prev_month = 12
            prev_roc_year = roc_year - 1
        else:
            prev_month = sel_month - 1
            prev_roc_year = roc_year

        prev_month_label = f"{prev_roc_year}年{prev_month}月剩餘量"
        curr_month_label = f"{roc_year}年{sel_month}月剩餘量"

        _, num_days = calendar.monthrange(ad_year, sel_month)

        if df_logs is not None and not df_logs.empty and '日期' in df_logs.columns:
            df_logs_calc = df_logs.copy()
            df_logs_calc['dt'] = pd.to_datetime(df_logs_calc['日期'], errors='coerce')
            df_month_logs = df_logs_calc[
                (df_logs_calc['dt'].dt.year == ad_year) & 
                (df_logs_calc['dt'].dt.month == sel_month)
            ].copy()
            df_month_logs['領用數量'] = pd.to_numeric(df_month_logs['領用數量'], errors='coerce').fillna(0)
        else:
            df_month_logs = pd.DataFrame()

        unique_drugs = sorted(df_inventory['藥品名稱'].dropna().unique().tolist()) if df_inventory is not None and not df_inventory.empty else []

        rows = []
        for drug in unique_drugs:
            row_data = {}
            row_data[f'藥品 / {sel_month}月用量'] = drug
            row_data[prev_month_label] = ""

            for d in range(1, num_days + 1):
                d_str = f"{sel_month}/{d}"
                if not df_month_logs.empty:
                    d_qty = df_month_logs[(df_month_logs['藥品名稱'] == drug) & (df_month_logs['dt'].dt.day == d)]['領用數量'].sum()
                    row_data[d_str] = int(d_qty) if d_qty > 0 else ""
                else:
                    row_data[d_str] = ""

            if not df_month_logs.empty:
                consumed = df_month_logs[df_month_logs['藥品名稱'] == drug]['領用數量'].sum()
                row_data[f'{sel_month}月消耗總量'] = int(consumed) if consumed > 0 else 0
            else:
                row_data[f'{sel_month}月消耗總量'] = 0

            row_data['購入量'] = ""
            row_data['過期報銷'] = ""
            row_data[f'{prev_month}月購入量'] = ""
            row_data[f'{sel_month}月購入量'] = ""

            curr_stock = df_inventory[df_inventory['藥品名稱'] == drug]['現有庫存'].sum() if df_inventory is not None and not df_inventory.empty else 0
            row_data[curr_month_label] = int(curr_stock)
            row_data[f'{sel_month}月總盤點'] = int(curr_stock)

            rows.append(row_data)

        official_df = pd.DataFrame(rows)

        st.subheader(f"📊 {term_title} ({roc_year}年{sel_month}月)")
        st.dataframe(official_df, use_container_width=True)

        output = io.BytesIO()
        sheet_tag = f"{sel_month}月"
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            official_df.to_excel(writer, sheet_name=sheet_tag, index=False, startrow=2)
            worksheet = writer.sheets[sheet_tag]
            worksheet.cell(row=1, column=1, value=term_title)

        excel_data = output.getvalue()

        st.download_button(
            label=f"📥 下載 {roc_year}年{sel_month}月 官方格式 Excel 月報表 (.xlsx)",
            data=excel_data,
            file_name=f"{term_title}_{sel_month}月.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    with tab2:
        col_input, _ = st.columns([2, 1])
        with col_input:
            custom_sheet = st.text_input(
                "📌 輸入 Google Sheet 標籤名稱 (若留空則預設讀取第一頁)：", 
                value="",
                placeholder="例如: Sheet1 或 9月"
            )

        target_sheet = custom_sheet.strip() if custom_sheet.strip() else None
        df_monthly_raw = load_raw_monthly_sheet(target_sheet)

        if df_monthly_raw is not None and not df_monthly_raw.empty:
            display_name = target_sheet if target_sheet else "試算表預設第一頁"
            st.success(f"✅ 成功載入【{display_name}】月報表紀錄！")
            st.dataframe(df_monthly_raw, use_container_width=True)
        else:
            st.warning("未找到指定的 Google Sheet 工作表分頁。")
