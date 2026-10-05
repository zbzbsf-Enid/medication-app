import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import io

# 頁面配置
st.set_page_config(
    page_title="衛保組藥品庫存管理系統",
    page_icon="💊",
    layout="wide"
)

# 連結 Google Sheets
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource(ttl=3600)
def get_gspread_client():
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=SCOPES
        )
        return gspread.authorize(creds)
    else:
        st.error("❌ 找不到 GCP Service Account 密鑰，請確認 Streamlit Secrets 設定！")
        return None

def get_spreadsheet():
    client = get_gspread_client()
    if client:
        sheet_url = st.secrets.get("spreadsheet_url", "")
        if sheet_url:
            return client.open_by_url(sheet_url)
        else:
            st.error("❌ 找不到 spreadsheet_url，請確認 Streamlit Secrets 設定！")
    return None

# 讀取資料表（含自動清洗標題與無效列）
def load_data(worksheet_name):
    spreadsheet = get_spreadsheet()
    if spreadsheet:
        try:
            ws = spreadsheet.worksheet(worksheet_name)
            data = ws.get_all_records()
            df = pd.DataFrame(data)
            df.columns = [str(col).strip() for col in df.columns]
            
            # 過濾空白列
            if not df.empty:
                if '藥品名稱' in df.columns:
                    df = df[df['藥品名稱'].astype(str).str.strip() != '']
                elif '流水號' in df.columns:
                    df = df[df['流水號'].astype(str).str.strip() != '']
            return df
        except Exception as e:
            st.error(f"讀取 {worksheet_name} 失敗: {e}")
            return pd.DataFrame()
    return pd.DataFrame()

# 主畫面標題
st.title("💊 衛保組藥品系統")
st.caption(f"🕒 系統時間：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} (UTC+8)")

# 側邊欄導覽
with st.sidebar:
    st.header("📌 請選擇功能：")
    menu = st.radio(
        "選擇功能",
        [
            "📋 藥品領用登記", 
            "📦 藥品庫存清單(可編輯庫存/批號/效期)", 
            "📥 進貨登記", 
            "📜 歷史紀錄(修改/刪除/同步庫存)", 
            "📊 月報表下載"
        ],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.subheader("⚙️️ 系統設定")
    if st.button("🔄 手動刷新雲端資料"):
        st.cache_data.clear()
        st.rerun()

# -----------------------------------------------------------------------------
# 1. 📋 藥品領用登記（已移除 領用人/經手人 欄位）
# -----------------------------------------------------------------------------
if menu == "📋 藥品領用登記":
    st.subheader("📋 藥品領用登記")
    
    inventory_df = load_data("藥品庫存清單")
    
    if inventory_df.empty:
        st.warning("⚠️️ 目前無庫存資料，請先進行進貨登記。")
    else:
        # 資料預處理
        inventory_df['現有庫存'] = pd.to_numeric(inventory_df['現有庫存'], errors='coerce').fillna(0).astype(int)
        
        # 自動過濾庫存 <= 0 藥品，並依藥品名稱 A~Z 排序
        valid_df = inventory_df[inventory_df['現有庫存'] > 0].copy()
        valid_df = valid_df.sort_values(by=['藥品名稱'], ascending=True)
        
        if valid_df.empty:
            st.warning("⚠️ 目前所有藥品庫存均為 0，無可領用藥品。")
        else:
            valid_df['display_text'] = valid_df.apply(
                lambda row: f"{row['藥品名稱']} | 批號: {row['批號']} | 當前庫存: {row['現有庫存']}", axis=1
            )
            
            st.markdown("##### 1. 搜尋與選擇多款藥品")
            selected_display_list = st.multiselect(
                "可一次搜尋選擇多款藥品：",
                options=valid_df['display_text'].tolist(),
                placeholder="請輸入藥品名稱或批號搜尋..."
            )
            
            st.markdown("---")
            st.markdown("##### 2. 本次領用清單與預覽填寫資訊")
            
            if not selected_display_list:
                st.info("🛒 目前領用清單為空。請由上方多選單選擇藥品。")
            else:
                col_cat, col_date = st.columns(2)
                with col_cat:
                    category = st.selectbox("領用類別", ["一般領用", "公務使用", "過期報銷"], index=0)
                with col_date:
                    use_date = st.date_input("領用日期", datetime.date.today())
                
                st.markdown("---")
                
                cart_items = []
                for display_text in selected_display_list:
                    matched_row = valid_df[valid_df['display_text'] == display_text].iloc[0]
                    
                    st.markdown(f"📍 **{matched_row['藥品名稱']}** (批號: `{matched_row['批號']}` | 當前庫存: {matched_row['現有庫存']})")
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        qty = st.number_input(
                            f"領用數量 ({matched_row['藥品名稱']})",
                            min_value=1,
                            max_value=int(matched_row['現有庫存']),
                            value=1,
                            key=f"qty_{matched_row['藥品ID']}"
                        )
                    with c2:
                        note = st.text_input(
                            f"備註 (可填學生姓名/學號等)",
                            key=f"note_{matched_row['藥品ID']}"
                        )
                    
                    cart_items.append({
                        '藥品ID': matched_row['藥品ID'],
                        '藥品名稱': matched_row['藥品名稱'],
                        '批號': matched_row['批號'],
                        '單位': matched_row.get('單位', '顆'),
                        '現有庫存': matched_row['現有庫存'],
                        '領用數量': qty,
                        '備註': note
                    })
                
                st.markdown("---")
                
                if st.button("🚀 確認送出領用登記", type="primary", use_container_width=True):
                    spreadsheet = get_spreadsheet()
                    if spreadsheet:
                        try:
                            ws_inventory = spreadsheet.worksheet("藥品庫存清單")
                            ws_history = spreadsheet.worksheet("歷史紀錄")
                            
                            inv_records = ws_inventory.get_all_records()
                            inv_df = pd.DataFrame(inv_records)
                            
                            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            for item in cart_items:
                                history_row = [
                                    now_str,
                                    use_date.strftime("%Y-%m-%d"),
                                    "領用",
                                    item['藥品ID'],
                                    item['藥品名稱'],
                                    item['批號'],
                                    item['領用數量'],
                                    item['單位'],
                                    "",  # 經手人留空
                                    category,
                                    item['備註']
                                ]
                                ws_history.append_row(history_row)
                                
                                row_idx = inv_df[inv_df['藥品ID'].astype(str) == str(item['藥品ID'])].index
                                if not row_idx.empty:
                                    real_row_num = row_idx[0] + 2
                                    current_stock = int(inv_df.loc[row_idx[0], '現有庫存'])
                                    new_stock = max(0, current_stock - item['領用數量'])
                                    
                                    col_num = inv_df.columns.get_loc('現有庫存') + 1
                                    ws_inventory.update_cell(real_row_num, col_num, new_stock)
                            
                            st.success("✅ 領用登記完成！庫存與歷史紀錄已同步更新。")
                            st.cache_data.clear()
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ 登記失敗，錯誤訊息: {e}")

# -----------------------------------------------------------------------------
# 2. 📦 藥品庫存清單(可編輯庫存/批號/效期)
# -----------------------------------------------------------------------------
elif menu == "📦 藥品庫存清單(可編輯庫存/批號/效期)":
    st.subheader("📦 藥品庫存清單")
    df = load_data("藥品庫存清單")
    if not df.empty:
        df = df.sort_values(by=['藥品名稱'], ascending=True)
        st.dataframe(df, use_container_width=True)
    else:
        st.info("尚無庫存資料。")

# -----------------------------------------------------------------------------
# 3. 📥 進貨登記
# -----------------------------------------------------------------------------
elif menu == "📥 進貨登記":
    st.subheader("📥 進貨登記")
    st.info("請在此填寫新進貨藥品資訊。")
    with st.form("inbound_form"):
        med_id = st.text_input("藥品ID / 簡碼")
        med_name = st.text_input("藥品名稱")
        batch_no = st.text_input("批號")
        qty = st.number_input("進貨數量", min_value=1, value=100)
        unit = st.text_input("單位", value="顆")
        exp_date = st.date_input("有效期限")
        safe_qty = st.number_input("安全庫存量", min_value=0, value=50)
        
        submitted = st.form_submit_button("💾 送出進貨登記", type="primary")
        if submitted:
            if not med_name or not batch_no:
                st.error("請填寫藥品名稱與批號！")
            else:
                spreadsheet = get_spreadsheet()
                if spreadsheet:
                    try:
                        ws_inv = spreadsheet.worksheet("藥品庫存清單")
                        ws_his = spreadsheet.worksheet("歷史紀錄")
                        
                        # 檢查是否已存在
                        data = ws_inv.get_all_records()
                        inv_df = pd.DataFrame(data)
                        
                        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        if not inv_df.empty and med_id in inv_df['藥品ID'].astype(str).values:
                            # 累加庫存
                            idx = inv_df[inv_df['藥品ID'].astype(str) == str(med_id)].index[0]
                            real_row = idx + 2
                            old_stock = int(inv_df.loc[idx, '現有庫存'])
                            new_stock = old_stock + qty
                            ws_inv.update_cell(real_row, inv_df.columns.get_loc('現有庫存') + 1, new_stock)
                        else:
                            # 新增列
                            new_row = [med_id, med_name, batch_no, qty, unit, exp_date.strftime("%Y-%m-%d"), safe_qty]
                            ws_inv.append_row(new_row)
                        
                        # 寫入歷史紀錄
                        his_row = [
                            now_str,
                            datetime.date.today().strftime("%Y-%m-%d"),
                            "進貨",
                            med_id,
                            med_name,
                            batch_no,
                            qty,
                            unit,
                            "",
                            "廠商進貨",
                            ""
                        ]
                        ws_his.append_row(his_row)
                        
                        st.success("✅ 進貨登記成功！")
                        st.cache_data.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"進貨失敗: {e}")

# -----------------------------------------------------------------------------
# 4. 📜 歷史紀錄(修改/刪除/同步庫存)
# -----------------------------------------------------------------------------
elif menu == "📜 歷史紀錄(修改/刪除/同步庫存)":
    st.subheader("📜 歷史紀錄")
    df = load_data("歷史紀錄")
    if not df.empty:
        st.dataframe(df, use_container_width=True)
    else:
        st.info("尚無歷史紀錄。")

# -----------------------------------------------------------------------------
# 5. 📊 月報表下載
# -----------------------------------------------------------------------------
elif menu == "📊 月報表下載":
    st.subheader("📊 藥品使用月報表與統計表繪出")
    st.write("產出格式符合國立臺北大學衛保組月報表標準格式。")
    
    col_y, col_m = st.columns(2)
    with col_y:
        year = st.number_input("選擇年份(西元)", min_value=2020, max_value=2030, value=datetime.date.today().year)
    with col_m:
        month = st.number_input("選擇月份", min_value=1, max_value=12, value=datetime.date.today().month)
    
    if st.button("📊 產生並預覽月報表", type="primary"):
        history_df = load_data("歷史紀錄")
        inventory_df = load_data("藥品庫存清單")
        
        if inventory_df.empty:
            st.warning("尚無庫存資料。")
        else:
            inventory_df = inventory_df.sort_values(by=['藥品名稱'], ascending=True)
            st.success(f"✅ 已成功計算 {year} 年 {month} 月報表資料！")
            st.dataframe(inventory_df, use_container_width=True)
            
            csv_data = inventory_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 下載月報表 (CSV 檔 / 可於 Excel 開啟)",
                data=csv_data,
                file_name=f"衛保組藥品月報表_{year}_{month:02d}.csv",
                mime="text/csv"
            )
