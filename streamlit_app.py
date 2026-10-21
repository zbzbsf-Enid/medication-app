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
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    conn = None

def load_sheet_data(worksheet_name, expected_cols):
    """讀取雲端試算表並自動補齊欠缺欄位"""
    try:
        if conn is None:
            return pd.DataFrame(columns=expected_cols)
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
        if conn is None:
            st.error("連線物件不存在，無法儲存！")
            return False
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
# 頁面 1: 藥品領用登記
# =============================================================================
if page == "💊 藥品領用登記":
    st.title("💊 藥品領用登記")
    inventory_df = load_sheet_data("庫存", INV_COLS)
    
    if inventory_df.empty:
        st.warning("⚠️ 目前「庫存」清單中無藥品資料，請先至「進貨登記」或「藥品庫存清單」新增藥品。")
    else:
        inventory_df["藥品名稱"] = inventory_df["藥品名稱"].astype(str).str.strip()
        inventory_df["批號"] = inventory_df["批號"].astype(str).str.strip()
        inventory_df["現有庫存"] = pd.to_numeric(inventory_df["現有庫存"], errors='coerce').fillna(0).astype(int)
        
        available_inv = inventory_df[inventory_df["現有庫存"] > 0].copy()
        available_inv = available_inv.sort_values(by=["藥品名稱", "批號"], ascending=[True, True]).reset_index(drop=True)
        
        if available_inv.empty:
            st.warning("⚠️ 目前所有藥品庫存皆為 0，無可用藥品供領用。")
        else:
            st.subheader("1. 搜尋與選擇藥品")
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
                    use_type = col1.selectbox("領用類別", ["一般領用", "公藥"])
                    usage_date = col2.date_input("領用日期", value=tw_now.date())
                    
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
                        note =
