import streamlit as st
import pandas as pd
import datetime
import calendar
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# 1. 頁面設定與系統初始化
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="衛保組藥品庫存管理系統",
    page_icon="💊",
    layout="wide"
)

# 取得台灣時間 (UTC+8)
def get_tw_now():
    return datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(hours=8)

# 定義各工作表標準欄位
INV_COLS = ["藥品名稱", "中文名稱", "單位", "現有庫存", "批號", "有效期限", "備註"]
USAGE_COLS = ["領用時間", "領用人", "藥品名稱", "中文名稱", "批號", "領用數量", "領用類別", "備註"]
RESTOCK_COLS = ["進貨時間", "藥品名稱", "中文名稱", "批號", "進貨數量", "廠商/來源", "備註"]

# Google Sheets 連線物件
conn = st.connection("gsheets", type=GSheetsConnection)

def load_sheet_data(worksheet_name, expected_cols):
    """讀取雲端試算表並自動補齊欠缺欄位"""
    try:
        df = conn.read(worksheet=worksheet_name, ttl="0s")
        if df is None or df.empty:
            return pd.DataFrame(columns=expected_cols)
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""
        return df[expected_cols]
    except Exception as e:
        st.error(f"讀取工作表 【{worksheet_name}】 失敗: {e}")
        return pd.DataFrame(columns=expected_cols)

def save_sheet_data(df, worksheet_name):
    """儲存資料至雲端試算表並重置快取"""
    try:
        conn.update(worksheet=worksheet_name, data=df)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"儲存工作表 【{worksheet_name}】 失敗: {e}")
        return False

# -----------------------------------------------------------------------------
# 2. 側邊欄與頁面選單
# -----------------------------------------------------------------------------
st.sidebar.title("🏥 衛保組藥品系統")
tw_now = get_tw_now()
st.sidebar.write(f"🕒 系統時間：{tw_now.strftime('%Y-%m-%d %H:%M')} (UTC+8)")

page = st.sidebar.radio(
    "📌 請選擇功能：",
    [
        "💊 藥品領用登記",
        "📦 藥品庫存清單(可編輯修正/批號/效期)",
        "🚚 進貨登記",
        "📜 歷史紀錄(修改/刪除/同步庫存)",
        "📊 月報表下載"
    ]
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 手動刷新雲端資料", use_container_width=True):
    st.cache_data.clear()
    st.success("已重置系統快取，成功讀取最新資料！")
    st.rerun()

# =============================================================================
# 頁面 1: 藥品領用登記 (用完自動隱藏 + A-Z 排序)
# =============================================================================
if page == "💊 藥品領用登記":
    st.title("💊 藥品領用登記")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    
    if inventory_df.empty:
        st.warning("⚠️ 目前「庫存」清單中無藥品資料，請先至「進貨登記」或「藥品庫存清單」新增藥品。")
    else:
        # 資料清理與格式轉換
        inventory_df["藥品名稱"] = inventory_df["藥品名稱"].astype(str).str.strip()
        inventory_df["批號"] = inventory_df["批號"].astype(str).str.strip()
        inventory_df["現有庫存"] = pd.to_numeric(inventory_df["現有庫存"], errors='coerce').fillna(0).astype(int)
        
        # 💡 過濾條件 1：僅顯示現有庫存 > 0 的藥品（庫存為 0 的批號自動隱藏）
        available_inv = inventory_df[inventory_df["現有庫存"] > 0].copy()
        
        # 💡 排序條件 2：自動依英文藥名 (A~Z) 與批號排序
        available_inv = available_inv.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
        
        if available_inv.empty:
            st.warning("⚠️ 目前所有藥品庫存皆為 0，無可用藥品供領用。")
        else:
            st.subheader("1. 搜尋與選擇藥品")
            
            # 建立多選下拉選單選項標籤
            available_inv["display_label"] = available_inv.apply(
                lambda r: f"{r['藥品名稱']} | 中文: {r['中文名稱']} | 批號: {r['批號']} | 庫存: {r['現有庫存']} {r['單位']}", axis=1
            )
            
            selected_labels = st.multiselect(
                "可一次搜尋並選擇多款藥品（已自動依藥名 A~Z 順序排列）：",
                options=available_inv["display_label"].tolist(),
                placeholder="請點擊或輸入關鍵字搜尋藥品..."
            )
            
            if selected_labels:
                st.subheader("2. 本次領用清單與預覽填寫資訊")
                with st.form("usage_form"):
                    col1, col2 = st.columns(2)
                    registrant = col1.text_input("領用人 / 經手人姓名", value="")
                    use_type = col2.selectbox("領用類別", ["一般領用", "公藥"])
                    usage_date = st.date_input("領用日期", value=tw_now.date())
                    
                    items_to_submit = []
                    for label in selected_labels:
                        item_row = available_inv[available_inv["display_label"] == label].iloc[0]
                        st.markdown(f"**📍 {item_row['藥品名稱']}** (批號: `{item_row['批號']}` | 當前庫存: {item_row['現有庫存']} {item_row['單位']})")
                        c_qty, c_note = st.columns([1, 2])
                        qty = c_qty.number_input(
                            f"領用數量 ({item_row['單位']})",
                            min_value=1,
                            max_value=int(item_row['現有庫存']),
                            value=1,
                            key=f"qty_{item_row['藥品名稱']}_{item_row['批號']}"
                        )
                        note = c_note.text_input("備註 (可填學生姓名/學號等)", value="", key=f"note_{item_row['藥品名稱']}_{item_row['批號']}")
                        
                        items_to_submit.append({
                            "藥品名稱": item_row['藥品名稱'],
                            "中文名稱": item_row['中文名稱'],
                            "批號": item_row['批號'],
                            "單位": item_row['單位'],
                            "領用數量": qty,
                            "備註": note
                        })
                        st.markdown("---")
                        
                    submitted = st.form_submit_button("🚀 確認送出領用登記", type="primary", use_container_width=True)
                    
                    if submitted:
                        if not registrant.strip():
                            st.error("❌ 請填寫領用人姓名！")
                        else:
                            usage_df = load_sheet_data("領用紀錄", USAGE_COLS)
                            new_usage_rows = []
                            
                            for item in items_to_submit:
                                med = item["藥品名稱"]
                                batch = item["批號"]
                                qty = item["領用數量"]
                                
                                # 扣減庫存表對應批號數量
                                mask = (inventory_df["藥品名稱"] == med) & (inventory_df["批號"] == batch)
                                if mask.any():
                                    inventory_df.loc[mask, "現有庫存"] -= qty
                                    
                                # 建立領用紀錄列
                                new_usage_rows.append({
                                    "領用時間": usage_date.strftime("%Y-%m-%d"),
                                    "領用人": registrant.strip(),
                                    "藥品名稱": med,
                                    "中文名稱": item["中文名稱"],
                                    "批號": batch,
                                    "領用數量": qty,
                                    "領用類別": use_type,
                                    "備註": item["備註"]
                                })
                            
                            # 儲存至 Google Sheets
                            if save_sheet_data(inventory_df[INV_COLS], "庫存"):
                                updated_usage = pd.concat([usage_df, pd.DataFrame(new_usage_rows)], ignore_index=True)
                                save_sheet_data(updated_usage[USAGE_COLS], "領用紀錄")
                                st.success("🎉 領用登記成功！庫存量已自動扣減。")
                                st.rerun()

# =============================================================================
# 頁面 2: 藥品庫存清單 (A-Z 排序 + 完整保留所有批號)
# =============================================================================
elif page == "📦 藥品庫存清單(可編輯修正/批號/效期)":
    st.title("📦 藥品庫存清單")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    
    if not inventory_df.empty:
        inventory_df["藥品名稱"] = inventory_df["藥品名稱"].astype(str).str.strip()
        inventory_df["批號"] = inventory_df["批號"].astype(str).str.strip()
        inventory_df["現有庫存"] = pd.to_numeric(inventory_df["現有庫存"], errors='coerce').fillna(0).astype(int)
        
        # 💡 自動依藥品英文名稱 (A~Z) 與批號進行整體排序
        inventory_df = inventory_df.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
        
        st.subheader("📋 目前庫存總表 (已自動依藥品名稱 A~Z 排序)")
        
        edited_df = st.data_editor(
            inventory_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "現有庫存": st.column_config.NumberColumn("現有庫存", min_value=0, step=1)
            }
        )
        
        if st.button("💾 儲存庫存表變更", type="primary"):
            edited_df["藥品名稱"] = edited_df["藥品名稱"].astype(str).str.strip()
            edited_df["批號"] = edited_df["批號"].astype(str).str.strip()
            edited_df = edited_df.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
            
            if save_sheet_data(edited_df[INV_COLS], "庫存"):
                st.success("✅ 庫存修改已成功儲存！")
                st.rerun()

# =============================================================================
# 頁面 3: 進貨登記
# =============================================================================
elif page == "🚚 進貨登記":
    st.title("🚚 藥品進貨登記")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    
    with st.form("restock_form"):
        col1, col2 = st.columns(2)
        med_name = col1.text_input("藥品英文名稱 (如: Actein 600mg)").strip()
        zh_name = col2.text_input("中文名稱 (如: 愛克痰發泡錠)").strip()
        
        col3, col4, col5 = st.columns(3)
        batch = col3.text_input("批號 (Batch No.)").strip()
        unit = col4.text_input("包裝單位", value="顆").strip()
        qty = col5.number_input("進貨數量", min_value=1, value=100)
        
        col6, col7 = st.columns(2)
        exp_date = col6.date_input("有效期限")
        restock_date = col7.date_input("進貨日期", value=tw_now.date())
        
        supplier = st.text_input("廠商 / 來源", value="")
        note = st.text_input("備註", value="")
        
        submitted = st.form_submit_button("📥 送出進貨登記", type="primary", use_container_width=True)
        
        if submitted:
            if not med_name or not batch:
                st.error("❌ 藥品英文名稱與批號為必填欄位！")
            else:
                restock_df = load_sheet_data("進貨紀錄", RESTOCK_COLS)
                
                inventory_df["藥品名稱"] = inventory_df["藥品名稱"].astype(str).str.strip()
                inventory_df["批號"] = inventory_df["批號"].astype(str).str.strip()
                
                mask = (inventory_df["藥品名稱"] == med_name) & (inventory_df["批號"] == batch)
                
                if mask.any():
                    curr_stock = pd.to_numeric(inventory_df.loc[mask, "現有庫存"], errors='coerce').fillna(0).astype(int).values[0]
                    inventory_df.loc[mask, "現有庫存"] = curr_stock + qty
                    inventory_df.loc[mask, "有效期限"] = str(exp_date)
                else:
                    new_inv_row = {
                        "藥品名稱": med_name,
                        "中文名稱": zh_name,
                        "單位": unit,
                        "現有庫存": qty,
                        "批號": batch,
                        "有效期限": str(exp_date),
                        "備註": note
                    }
                    inventory_df = pd.concat([inventory_df, pd.DataFrame([new_inv_row])], ignore_index=True)
                
                # 重新按 A~Z 字母順序排序庫存
                inventory_df = inventory_df.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
                
                new_restock_row = {
                    "進貨時間": restock_date.strftime("%Y-%m-%d"),
                    "藥品名稱": med_name,
                    "中文名稱": zh_name,
                    "批號": batch,
                    "進貨數量": qty,
                    "廠商/來源": supplier,
                    "備註": note
                }
                updated_restock = pd.concat([restock_df, pd.DataFrame([new_restock_row])], ignore_index=True)
                
                if save_sheet_data(inventory_df[INV_COLS], "庫存") and save_sheet_data(updated_restock[RESTOCK_COLS], "進貨紀錄"):
                    st.success(f"🎉 藥品「{med_name}」進貨成功！現有庫存已增加。")
                    st.rerun()

# =============================================================================
# 頁面 4: 歷史紀錄
# =============================================================================
elif page == "📜 歷史紀錄(修改/刪除/同步庫存)":
    st.title("📜 歷史領用與進貨紀錄")
    tab1, tab2 = st.tabs(["領用紀錄", "進貨紀錄"])
    
    with tab1:
        usage_df = load_sheet_data("領用紀錄", USAGE_COLS)
        st.dataframe(usage_df, use_container_width=True)
        
    with tab2:
        restock_df = load_sheet_data("進貨紀錄", RESTOCK_COLS)
        st.dataframe(restock_df, use_container_width=True)

# =============================================================================
# 頁面 5: 月報表下載 (🌟 已包含雙重批號比對 + 期初庫存正確回推 + A-Z 排序)
# =============================================================================
elif page == "📊 月報表下載":
    st.title("📊 藥品使用月報表與統計表繪出")
    st.write("產出格式符合國立臺北大學衛保組月報表標準格式。")
    
    tw_now = get_tw_now()
    col_y, col_m = st.columns(2)
    selected_year = col_y.number_input("選擇年份(西元)", min_value=2020, max_value=2030, value=tw_now.year)
    selected_month = col_m.selectbox("選擇月份", list(range(1, 13)), index=tw_now.month - 1)
    
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
        
        # 💡 庫存清單先依藥品名稱 (A~Z) 與批號進行自動排序
        inventory_df["藥品名稱"] = inventory_df["藥品名稱"].astype(str).str.strip()
        inventory_df["批號"] = inventory_df["批號"].astype(str).str.strip() if "批號" in inventory_df.columns else ""
        inventory_df = inventory_df.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
        
        usage_df["領用時間"] = usage_df["領用時間"].astype(str)
        usage_df["藥品名稱"] = usage_df["藥品名稱"].astype(str).str.strip()
        usage_df["批號"] = usage_df["批號"].astype(str).str.strip() if "批號" in usage_df.columns else ""
            
        restock_df["進貨時間"] = restock_df["進貨時間"].astype(str)
        restock_df["藥品名稱"] = restock_df["藥品名稱"].astype(str).str.strip()
        restock_df["批號"] = restock_df["批號"].astype(str).str.strip() if "批號" in restock_df.columns else ""
        
        _, num_days = calendar.monthrange(selected_year, selected_month)
        m_start = f"{selected_year}-{selected_month:02d}-01"
        m_end = f"{selected_year}-{selected_month:02d}-{num_days:02d}"
        
        report_rows = []
        
        for _, inv_row in inventory_df.iterrows():
            med_name = str(inv_row.get("藥品名稱", "")).strip()
            if not med_name:
                continue
                
            zh_name = str(inv_row.get("中文名稱", "")).strip()
            batch_no = str(inv_row.get("批號", "")).strip()
            full_name = f"{med_name}({zh_name})" if zh_name else med_name
            
            # 當前庫存表中紀錄的實時庫存（即當期期末剩餘量）
            curr_stock = pd.to_numeric(inv_row.get(target_col, 0), errors='coerce')
            curr_stock = 0 if pd.isna(curr_stock) else int(curr_stock)
            
            # 1. 計算當月每日領用量 (雙重驗證: 藥品名稱 + 批號)
            daily_quantities = {}
            daily_total = 0
            for d in range(1, num_days + 1):
                day_str = f"{selected_year}-{selected_month:02d}-{d:02d}"
                col_name = f"{selected_month}/{d}"
                
                if batch_no:
                    matched_u = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["批號"] == batch_no) & (usage_df["領用時間"] == day_str)]
                else:
                    matched_u = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["領用時間"] == day_str)]
                    
                u_qty = pd.to_numeric(matched_u["領用數量"], errors='coerce').sum() if not matched_u.empty else 0
                qty_val = int(u_qty) if not pd.isna(u_qty) else 0
                daily_quantities[col_name] = qty_val
                daily_total += qty_val

            # 2. 計算當月總公藥領用與進貨量 (雙重驗證: 藥品名稱 + 批號)
            if batch_no:
                m_usage = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["批號"] == batch_no) & (usage_df["領用時間"] >= m_start) & (usage_df["領用時間"] <= m_end)]
                matched_r = restock_df[(restock_df["藥品名稱"] == med_name) & (restock_df["批號"] == batch_no) & (restock_df["進貨時間"] >= m_start) & (restock_df["進貨時間"] <= m_end)]
            else:
                m_usage = usage_df[(usage_df["藥品名稱"] == med_name) & (usage_df["領用時間"] >= m_start) & (usage_df["領用時間"] <= m_end)]
                matched_r = restock_df[(restock_df["藥品名稱"] == med_name) & (restock_df["進貨時間"] >= m_start) & (restock_df["進貨時間"] <= m_end)]
                
            public_qty = pd.to_numeric(m_usage[m_usage["領用類別"] == "公藥"]["領用數量"], errors='coerce').sum() if not m_usage.empty else 0
            restock_qty = pd.to_numeric(matched_r["進貨數量"], errors='coerce').sum() if not matched_r.empty else 0
            
            # 3. 🌟 自動回推計算期初庫存量：期初剩餘量 = 期末現有庫存 + 當月使用量 - 當月進貨量
            start_stock = curr_stock + daily_total - restock_qty

            row_dict = {
                "藥品名稱\n(商品名/中文)": full_name,
                f"{prev_roc_year}年{prev_month}月\n剩餘量": int(start_stock)
            }
            row_dict.update(daily_quantities)
            row_dict["當月使用\n總量"] = int(daily_total)
            row_dict["購入量"] = int(restock_qty)
            row_dict["過期報銷"] = 0
            row_dict["公藥使用"] = int(public_qty)
            row_dict[f"{roc_year}年{selected_month}月\n期末剩餘量"] = int(curr_stock)
            
            report_rows.append(row_dict)

        # 轉換為 DataFrame 並進行 A~Z 最終排序確認
        report_df = pd.DataFrame(report_rows)
        if not report_df.empty:
            report_df = report_df.sort_values(by="藥品名稱\n(商品名/中文)", ascending=True).reset_index(drop=True)

        st.subheader(f"📋 國立臺北大學衛保組 {acad_year}學年度{semester}藥品使用月報表 ({roc_year}年{selected_month}月)")
        st.dataframe(report_df, use_container_width=True)
        
        # 下載 CSV 檔按鈕
        csv_data = report_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載月報表 (CSV 檔 / 可於 Excel 開啟)",
            data=csv_data,
            file_name=f"衛保組_{roc_year}年{selected_month}月藥品使用月報表.csv",
            mime="text/csv",
            type="primary"
        )
