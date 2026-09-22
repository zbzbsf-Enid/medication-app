import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
from datetime import datetime
import calendar
import io

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

page = st.sidebar.radio(
    "📌 請選擇功能：",
    ["📦 藥品庫存清單", "📋 多品項領用登記", "🚚 進貨登記", "📜 歷史紀錄(修改/刪除)", "📊 月報表下載"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 系統設定與維護")

sh = get_spreadsheet()
if sh:
    try:
        existing_sheets = [ws.title for ws in sh.worksheets()]
        st.sidebar.success("✅ Google Sheet 連線成功")
        st.sidebar.write("🔍 目前雲端分頁：", existing_sheets)
    except Exception as e:
        st.sidebar.warning(f"⚠️ 無法讀取分頁列表: {e}")

if st.sidebar.button("🔄 手動刷新雲端資料", use_container_width=True):
    st.cache_data.clear()
    st.sidebar.success("已刷新資料快取！")
    st.rerun()

# ---------------------------------------------------------
# 3. 主頁面內容控制
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

# --- 頁面 2: 多品項領用登記 ---
elif page == "📋 多品項領用登記":
    st.title("📋 多品項藥品領用登記")
    inventory_df = load_sheet_data("庫存")
    
    if inventory_df.empty or "藥品名稱" not in inventory_df.columns:
        st.warning("⚠️ 請先確保「庫存」分頁有包含「藥品名稱」欄位。")
    else:
        st.subheader("1. 請在表格中直接填入欲領用的數量")
        
        # 準備供編輯的表格數據
        target_col = "現有庫存" if "現有庫存" in inventory_df.columns else "剩餘庫存"
        
        display_cols = ["藥品名稱", "中文名稱", target_col]
        for col in ["批號", "有效日期", "用途/備註"]:
            if col in inventory_df.columns:
                display_cols.append(col)
                
        cart_df = inventory_df[display_cols].copy()
        cart_df["領用數量"] = 0  # 預設領用數量為0
        
        # 移動「領用數量」欄位到前面
        cols = ["領用數量"] + [c for c in cart_df.columns if c != "領用數量"]
        cart_df = cart_df[cols]

        edited_df = st.data_editor(
            cart_df,
            column_config={
                "領用數量": st.column_config.NumberColumn(
                    "領用數量",
                    help="直接輸入預計領取的數量",
                    min_value=0,
                    step=1,
                    default=0
                )
            },
            disabled=[c for c in cart_df.columns if c != "領用數量"],
            hide_index=True,
            use_container_width=True,
            key="multi_claim_editor"
        )
        
        # 篩選出數量 > 0 的項目
        selected_items = edited_df[edited_df["領用數量"] > 0].copy()
        
        st.markdown("---")
        st.subheader("2. 領用資訊與二次確認")
        
        col_info1, col_info2 = st.columns(2)
        use_date = col_info1.date_input("領用日期", datetime.now())
        remarks = col_info2.text_input("備註 / 領用單位或個人", "")
        
        if not selected_items.empty:
            st.info(f"🛒 **目前已選擇 {len(selected_items)} 項藥品：**")
            st.dataframe(selected_items[["藥品名稱", "中文名稱", "領用數量", target_col]], hide_index=True, use_container_width=True)
            
            # 安全防呆確認選項
            confirm_check = st.checkbox("✅ 我已仔細核對上述藥品品項與數量，確認無誤。")
            
            if st.button("確認寫入雲端並扣減庫存", type="primary", use_container_width=True):
                if not confirm_check:
                    st.error("⚠️ 請先勾選「我已仔細核對上述藥品品項與數量，確認無誤。」進行確認！")
                else:
                    # 開始寫入領用紀錄與更新庫存
                    usage_df = load_sheet_data("領用紀錄")
                    
                    new_rows = []
                    for _, row in selected_items.iterrows():
                        new_rows.append({
                            "領用時間": str(use_date),
                            "藥品名稱": row["藥品名稱"],
                            "中文名稱": row["中文名稱"],
                            "領用數量": int(row["領用數量"]),
                            "備註": remarks
                        })
                        
                        # 同步更新原庫存 dataframe
                        m_idx = inventory_df[inventory_df["藥品名稱"] == row["藥品名稱"]].index
                        if not m_idx.empty:
                            idx = m_idx[0]
                            curr_q = pd.to_numeric(inventory_df.at[idx, target_col], errors='coerce')
                            curr_q = 0 if pd.isna(curr_q) else int(curr_q)
                            inventory_df.at[idx, target_col] = max(0, curr_q - int(row["領用數量"]))
                    
                    new_usage_df = pd.concat([usage_df, pd.DataFrame(new_rows)], ignore_index=True)
                    
                    # 寫入領用紀錄
                    ok1, msg1 = safe_update_sheet("領用紀錄", new_usage_df, ["領用時間", "藥品名稱", "中文名稱", "領用數量", "備註"])
                    # 寫入庫存
                    ok2, msg2 = safe_update_sheet("庫存", inventory_df)
                    
                    if ok1 and ok2:
                        st.success(f"🎉 成功登記 {len(new_rows)} 項藥品領用，庫存已同步扣減！")
                        st.cache_data.clear()
                    else:
                        st.error(f"❌ 更新失敗: 領用紀錄({msg1}) / 庫存({msg2})")
        else:
            st.write("💡 請在上方表格的「領用數量」欄位輸入數字即可進行領用預覽。")

# --- 頁面 3: 進貨登記 ---
elif page == "🚚 進貨登記":
    st.title("🚚 藥品進貨登記")
    inventory_df = load_sheet_data("庫存")
    target_col = "現有庫存" if "現有庫存" in inventory_df.columns else "剩餘庫存"
    
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

# --- 頁面 4: 歷史紀錄(修改/刪除) ---
elif page == "📜 歷史紀錄(修改/刪除)":
    st.title("📜 歷史紀錄維護（修改與整筆刪除）")
    st.caption("💡 在下方表格中可直接修改內容或點選列進行刪除，完成後點擊「儲存更新至雲端」。")
    
    tab_rec1, tab_rec2 = st.tabs(["📋 領用紀錄編輯", "🚚 進貨紀錄編輯"])
    
    with tab_rec1:
        u_df = load_sheet_data("領用紀錄")
        if not u_df.empty:
            edited_u_df = st.data_editor(
                u_df,
                num_rows="dynamic",
                use_container_width=True,
                key="usage_editor"
            )
            if st.button("💾 儲存領用紀錄修改至雲端", type="primary"):
                ok, msg = safe_update_sheet("領用紀錄", edited_u_df)
                if ok:
                    st.success("✅ 領用紀錄已順利同步更新至 Google Sheets！")
                    st.cache_data.clear()
                else:
                    st.error(f"❌ 儲存失敗: {msg}")
        else:
            st.info("目前無領用紀錄。")
            
    with tab_rec2:
        r_df = load_sheet_data("進貨紀錄")
        if not r_df.empty:
            edited_r_df = st.data_editor(
                r_df,
                num_rows="dynamic",
                use_container_width=True,
                key="restock_editor"
            )
            if st.button("💾 儲存進貨紀錄修改至雲端", type="primary"):
                ok, msg = safe_update_sheet("進貨紀錄", edited_r_df)
                if ok:
                    st.success("✅ 進貨紀錄已順利同步更新至 Google Sheets！")
                    st.cache_data.clear()
                else:
                    st.error(f"❌ 儲存失敗: {msg}")
        else:
            st.info("目前無進貨紀錄。")

# --- 頁面 5: 月報表下載 ---
elif page == "📊 月報表下載":
    st.title("📊 藥品使用月報表與統計表繪出")
    st.write("產出格式符合國立臺北大學衛保組月報表標準格式。")
    
    col_y, col_m = st.columns(2)
    selected_year = col_y.number_input("選擇年份(西元)", min_value=2020, max_value=2030, value=datetime.now().year)
    selected_month = col_m.selectbox("選擇月份", list(range(1, 13)), index=datetime.now().month - 1)
    
    # 計算民國年與學年度
    roc_year = selected_year - 1911
    acad_year = roc_year - 1 if selected_month < 8 else roc_year
    semester = "下學期" if 1 <= selected_month <= 7 else "上學期"
    
    prev_month = 12 if selected_month == 1 else selected_month - 1
    prev_year = selected_year - 1 if selected_month == 1 else selected_year
    prev_roc_year = prev_year - 1911
    
    if st.button("📥 產生並預覽月報表", type="primary", use_container_width=True):
        inventory_df = load_sheet_data("庫存")
        usage_df = load_sheet_data("領用紀錄")
        restock_df = load_sheet_data("進貨紀錄")
        
        target_col = "現有庫存" if "現有庫存" in inventory_df.columns else "剩餘庫存"
        
        # 取得當月天數
        _, num_days = calendar.monthrange(selected_year, selected_month)
        day_cols = [f"{selected_month}/{d}" for d in range(1, num_days + 1)]
        
        report_rows = []
        
        for _, inv_row in inventory_df.iterrows():
            med_name = inv_row.get("藥品名稱", "")
            zh_name = inv_row.get("中文名稱", "")
            full_name = f"{med_name}({zh_name})" if zh_name else med_name
            curr_stock = pd.to_numeric(inv_row.get(target_col, 0), errors='coerce')
            curr_stock = 0 if pd.isna(curr_stock) else int(curr_stock)
            
            row_dict = {
                "藥品名稱\n(商品名/中文)": full_name,
                f"{prev_roc_year}年{prev_month}月\n剩餘量": curr_stock  # 簡化統計
            }
            
            # 計算每日領用量
            daily_total = 0
            for d in range(1, num_days + 1):
                day_str = f"{selected_year}-{selected_month:02d}-{d:02d}"
                col_name = f"{selected_month}/{d}"
                
                if not usage_df.empty and "領用時間" in usage_df.columns:
                    matched_u = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["領用時間"] == day_str)]
                    u_qty = pd.to_numeric(matched_u["領用數量"], errors='coerce').sum()
                else:
                    u_qty = 0
                    
                row_dict[col_name] = int(u_qty) if not pd.isna(u_qty) else 0
                daily_total += row_dict[col_name]
                
            # 進貨數量計算
            if not restock_df.empty and "進貨時間" in restock_df.columns:
                m_start = f"{selected_year}-{selected_month:02d}-01"
                m_end = f"{selected_year}-{selected_month:02d}-{num_days:02d}"
                matched_r = restock_df[(restock_df["藥品名稱"] == med_name) & (restock_df["進貨時間"] >= m_start) & (restock_df["進貨時間"] <= m_end)]
                restock_qty = pd.to_numeric(matched_r["進貨數量"], errors='coerce').sum()
            else:
                restock_qty = 0
                
            row_dict["當月使用\n總量"] = daily_total
            row_dict["購入量"] = int(restock_qty)
            row_dict["過期報銷"] = 0
            row_dict["公藥使用"] = 0
            row_dict[f"{roc_year}年{selected_month}月\n期末剩餘量"] = curr_stock
            
            report_rows.append(row_dict)
            
        report_df = pd.DataFrame(report_rows)
        
        st.subheader("📋 報表預覽")
        st.dataframe(report_df, use_container_width=True, hide_index=True)
        
        # 匯出至 Excel (含完美標題標頭)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            title_text = f"國立臺北大學衛保組 {acad_year}學年度{semester}藥品使用月報與全學期統計表 ({roc_year}學年度{selected_month}月起)"
            
            # 先將 Title 寫在第 1 行
            workbook = writer.book
            worksheet = workbook.add_worksheet('月報表')
            writer.sheets['月報表'] = worksheet
            
            title_format = workbook.add_format({'bold': True, 'font_size': 14})
            worksheet.write(0, 0, title_text, title_format)
            
            # 將資料表從第 2 行開始寫入
            report_df.to_excel(writer, sheet_name='月報表', startrow=1, index=False)
            
        excel_data = output.getvalue()
        
        file_name = f"國立臺北大學衛保組_{acad_year}學年度_{selected_month}月藥品使用月報表.xlsx"
        st.download_button(
            label="📥 點此下載標準 Excel 月報表 (.xlsx)",
            data=excel_data,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
