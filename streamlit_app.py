import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime

# 頁面基本設定
st.set_page_config(
    page_title="衛保組藥品管理系統",
    page_icon="💊",
    layout="wide"
)

# 注入清新明亮主題 CSS 與 微軟正黑體 16號字體 (16px)
st.markdown("""
    <style>
    /* 全域字型設定 (微軟正黑體 16px) */
    html, body, [class*="css"], .stApp {
        font-family: 'Microsoft JhengHei', '微軟正黑體', 'PingFang TC', sans-serif !important;
        font-size: 16px !important;
        background-color: #f8fafc !important;
        color: #1e293b !important;
    }

    /* 標題樣式調整 */
    h1, h2, h3, h4 {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        color: #0f172a !important;
        font-weight: 700 !important;
    }

    /* 輸入框、下拉選單與數字輸入框文字大小 */
    .stTextInput input, .stNumberInput input, .stSelectbox, .stMultiSelect {
        font-size: 16px !important;
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
    }

    /* 清新卡片容器設計 */
    div[data-testid="stForm"] {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
    }

    /* 清新明亮風格按鈕 (薄荷青綠色調) */
    .stButton > button, div[data-testid="stForm"] button {
        font-size: 16px !important;
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        background-color: #0d9488 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 8px 20px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease-in-out;
    }
    .stButton > button:hover, div[data-testid="stForm"] button:hover {
        background-color: #0f766e !important;
        box-shadow: 0 2px 8px rgba(13, 148, 136, 0.3);
    }

    /* 側邊欄色彩優化 */
    section[data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
        border-right: 1px solid #e2e8f0;
    }
    </style>
""", unsafe_allow_html=True)

st.title("💊 衛保組藥品管理系統")

# 建立 Google Connection
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"❌ 無法建立 Google 連線，請檢查 Streamlit Cloud 的 Secrets 設定。錯誤訊息: {e}")
    st.stop()

# 讀取試算表資料
def load_data():
    try:
        df = conn.read(worksheet="庫存", ttl=0)
        return df
    except Exception as e:
        st.error(f"❌ 讀取試算表失敗，請確認標題與權限。錯誤：{e}")
        st.stop()

# 側邊選單
st.sidebar.title("📌 功能選單")
if st.sidebar.button("🔄 手動刷新雲端最新資料"):
    st.cache_data.clear()
    st.rerun()

menu = st.sidebar.radio("請選擇功能頁面", ["💊 多項藥品領用與登記", "📦 當前庫存總覽"])

df_inventory = load_data()

if menu == "💊 多項藥品領用與登記":
    st.header("📋 批量藥品領用登記")
    
    # 建立多選下拉菜單標籤
    df_inventory['display_name'] = df_inventory['藥品名稱(英文)'].fillna('') + " | " + df_inventory['中文名稱'].fillna('') + " (批號: " + df_inventory['批號'].astype(str) + ")"
    options = df_inventory['display_name'].tolist()

    selected_items = st.multiselect(
        "請點擊或輸入關鍵字選擇欲領取的藥品（可同時選擇多項）：",
        options=options,
        placeholder="搜尋或選擇藥品..."
    )

    if selected_items:
        st.markdown("---")
        st.subheader("✏️ 請鍵入各藥品的領用數量")
        
        with st.form("batch_checkout_form"):
            quantities = {}
            for item in selected_items:
                row_data = df_inventory[df_inventory['display_name'] == item].iloc[0]
                curr_stock = int(row_data['目前庫存'])
                
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.markdown(f"**{item}**")
                    st.caption(f"目前剩餘庫存：`{curr_stock}`")
                with col2:
                    # 可直接輸入/鍵入數字的選項
                    qty = st.number_input(
                        f"領用數量",
                        min_value=1,
                        max_value=max(curr_stock, 1),
                        value=1,
                        step=1,
                        key=f"input_{item}"
                    )
                    quantities[item] = qty
                st.markdown("<hr style='margin: 8px 0; border-color: #f1f5f9;'>", unsafe_allow_html=True)

            remarks = st.text_input("備註（選填，例如：學生領用 / 活動備用）：", placeholder="請輸入領用備註...")
            submit_btn = st.form_submit_button("✅ 一鍵完成登記與庫存扣減")

            if submit_btn:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_logs = []

                # 扣減庫存並建立 Log 紀錄
                for item, qty in quantities.items():
                    idx = df_inventory[df_inventory['display_name'] == item].index[0]
                    old_stock = int(df_inventory.loc[idx, '目前庫存'])
                    new_stock = max(0, old_stock - qty)
                    
                    df_inventory.loc[idx, '目前庫存'] = new_stock
                    
                    new_logs.append({
                        "領用時間": now_str,
                        "藥品名稱": df_inventory.loc[idx, '藥品名稱(英文)'],
                        "中文名稱": df_inventory.loc[idx, '中文名稱'],
                        "領用數量": qty,
                        "剩餘庫存": new_stock,
                        "備註": remarks
                    })

                # 更新 Google Sheets
                try:
                    # 1. 更新庫存分頁
                    df_save = df_inventory.drop(columns=['display_name'])
                    conn.update(worksheet="庫存", data=df_save)
                    
                    # 2. 寫入領用紀錄分頁
                    try:
                        df_logs_existing = conn.read(worksheet="領用紀錄", ttl=0)
                        df_logs_new = pd.DataFrame(new_logs)
                        df_logs_updated = pd.concat([df_logs_existing, df_logs_new], ignore_index=True)
                    except Exception:
                        df_logs_updated = pd.DataFrame(new_logs)
                        
                    conn.update(worksheet="領用紀錄", data=df_logs_updated)
                    
                    st.success("🎉 批量領用登記成功！試算表庫存與 Log 已完成自動更新。")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 更新試算表失敗，原因：{e}")

elif menu == "📦 當前庫存總覽":
    st.header("📦 當前藥品庫存總覽")
    st.dataframe(
        df_inventory.drop(columns=['display_name'], errors='ignore'),
        use_container_width=True,
        hide_index=True
    )
