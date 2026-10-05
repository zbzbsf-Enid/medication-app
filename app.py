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
        ["📋 藥品領用登記", "📦 藥品庫存清單(可編輯庫存/批號/效期)", "📥 進貨登記", "📜 歷史紀錄(修改/刪除/同步庫存)", "📊 月報表下載"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.subheader("⚙️ 系統設定")
    if st.button("🔄 手動刷新雲端資料"):
        st.cache_data.clear()
        st.rerun()

# -----------------------------------------------------------------------------
# 1. 📋 藥品領用登記（已移除 領用人/經手人 欄位）
# -----------------------------------------------------------------------------
if menu == "📋 藥品領用登记":
    st.subheader("📋 藥品領用登記")
    
    inventory_df = load_data("藥品庫存清單")
    
    if inventory_df.empty:
        st.warning("⚠️ 目前無庫存資料，請先進行進貨登記。")
    else:
        # 資料預處理
        inventory_df['現有庫存'] = pd.to_numeric(inventory_df['現有庫存'], errors='coerce').fillna(0).astype(int)
        
        # 1. 自動過濾掉庫存 <= 0 的藥品
        valid_df = inventory_df[inventory_df['現有庫存'] > 0].copy()
        
        # 2. 依照藥品名稱 A~Z 排序
        valid_df = valid_df.sort_values(by=['藥品名稱'], ascending=True)
        
        if valid_df.empty:
            st.warning("⚠️ 目前所有藥品庫存均為 0，無可領用藥品。")
        else:
            # 建立選單選項文字
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
                # 領用基本資料（已移除 領用人/經手人 欄位）
                col_cat, col_date = st.columns(2)
                with col_cat:
                    category = st.selectbox("領用類別", ["一般領用", "公務使用", "過期報銷"], index=0)
                with col_date:
                    use_date = st.date_input("領用日期", datetime.date.today())
                
                st.markdown("---")
                
                # 填寫各個藥品的領用數量與備註
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
                
                # 確認送出按鈕（不需檢查經手人姓名）
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
                                # 1. 寫入歷史紀錄（經手人欄位直接填入 "" 空白）
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
                                
                                # 2. 扣減庫存清單
                                row_idx = inv_df[inv_df['藥品ID'].astype(str) == str(item['藥品ID'])].index
                                if not row_idx.empty:
                                    real_row_num = row_idx[0] + 2  # 包含標題列 offset
                                    current_stock = int(inv_df.loc[row_idx[0], '現有庫存'])
                                    new_stock = max(0, current_stock - item['領用數量'])
                                    
                                    col_num = inv_df.columns.get_loc('現有庫存') + 1
                                    ws_inventory.update_cell(real_row_num, col_num, new_stock)
                            
                            st.success("✅ 領用登記完成！庫存與歷史紀錄已同步更新。")
                            st.cache_data.clear()
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ 登記失敗，錯誤訊息: {e}")

# (其他功能頁面 code 保持不變...)
