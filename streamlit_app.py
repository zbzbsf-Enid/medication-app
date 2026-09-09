import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
import io
import openpyxl

# 頁面基本設定
st.set_page_config(
    page_title="國立臺北大學衛保組 - 藥品管理與月報系統",
    page_icon="💊",
    layout="wide"
)

# 注入清新明亮主題 CSS 與 微軟正黑體 16px 字體
st.markdown("""
    <style>
    html, body, [class*="css"], .stApp {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        font-size: 16px !important;
        background-color: #f8fafc !important;
        color: #1e293b !important;
    }
    h1, h2, h3, h4 {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        color: #0f172a !important;
        font-weight: 700 !important;
    }
    .stTextInput input, .stNumberInput input, .stSelectbox, .stMultiSelect {
        font-size: 16px !important;
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
    }
    div[data-testid="stForm"] {
        background-color: #ffffff;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
    }
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
    section[data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
        border-right: 1px solid #e2e8f0;
    }
    </style>
""", unsafe_allow_html=True)

st.title("💊 國立臺北大學衛保組 藥品管理系統")

# 建立 Google Connection
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"❌ 無法建立 Google 連線：{e}")
    st.stop()

def load_data():
    try:
        df = conn.read(worksheet="庫存", ttl=0)
        return df
    except Exception as e:
        st.error(f"❌ 讀取試算表失敗：{e}")
        st.stop()

# 側邊欄選單
st.sidebar.title("📌 功能選單")
if st.sidebar.button("🔄 手動刷新雲端資料"):
    st.cache_data.clear()
    st.rerun()

menu = st.sidebar.radio(
    "請選擇功能頁面", 
    ["💊 多項藥品領用登記", "📦 當前庫存總覽", "📊 用藥月報與學期統計表"]
)

df_inventory = load_data()

# 頁面 1：多項藥品領用登記
if menu == "💊 多項藥品領用登記":
    st.header("📋 批量藥品領用登記")
    df_inventory['display_name'] = (
        df_inventory['藥品名稱(英文)'].fillna('') + " (" + 
        df_inventory['中文名稱'].fillna('') + ") - 批號:" + 
        df_inventory['批號'].astype(str)
    )
    options = df_inventory['display_name'].tolist()

    selected_items = st.multiselect(
        "請選擇或搜尋欲領取的藥品（可多選）：",
        options=options,
        placeholder="點擊選擇藥品..."
    )

    if selected_items:
        st.markdown("---")
        st.subheader("✏️ 請輸入領用數量")
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
                    qty = st.number_input("領用數量", min_value=1, max_value=max(curr_stock, 1), value=1, step=1, key=f"input_{item}")
                    quantities[item] = qty
                st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

            remarks = st.text_input("領用備註/用途：", placeholder="例如：衛保組公用 / 門診備用")
            submit_btn = st.form_submit_button("✅ 完成登記並更新庫存")

            if submit_btn:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_logs = []
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

                try:
                    df_save = df_inventory.drop(columns=['display_name'])
                    conn.update(worksheet="庫存", data=df_save)
                    try:
                        df_logs_existing = conn.read(worksheet="領用紀錄", ttl=0)
                        df_logs_updated = pd.concat([df_logs_existing, pd.DataFrame(new_logs)], ignore_index=True)
                    except Exception:
                        df_logs_updated = pd.DataFrame(new_logs)
                    conn.update(worksheet="領用紀錄", data=df_logs_updated)
                    st.success("🎉 批量領用登記成功！庫存與 Log 已更新。")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 更新失敗：{e}")

# 頁面 2：當前庫存總覽
elif menu == "📦 當前庫存總覽":
    st.header("📦 當前藥品庫存總覽")
    st.dataframe(df_inventory.drop(columns=['display_name'], errors='ignore'), use_container_width=True, hide_index=True)

# 頁面 3：用藥月報與學期統計表 (對齊臺北大學官方 Excel 格式)
elif menu == "📊 用藥月報與學期統計表":
    st.header("📊 國立臺北大學衛保組 藥品使用月報與全學期統計表")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_year = st.selectbox("學年度/年份", ["115學年度", "114學年度"], index=0)
    with col_sel2:
        selected_month = st.selectbox("統計月份", ["9月", "10月", "11月", "12月", "1月"], index=0)

    # 建立與 Excel 一致的 37 欄報表結構
    report_rows = []
    days = [f"9/{d}" for d in [1, 2, 3, 4, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29, 30]]
    
    for _, row in df_inventory.iterrows():
        eng_name = str(row.get('藥品名稱(英文)', ''))
        cht_name = str(row.get('中文名稱', ''))
        combined_name = f"{eng_name} ({cht_name})" if cht_name else eng_name
        stock = int(row.get('目前庫存', 0))
        expiry = str(row.get('有效期限', ''))

        r_dict = {"藥品名稱\n(商品名/中文)": combined_name, "上月剩餘量": stock}
        for d in days:
            r_dict[d] = 0
        r_dict.update({
            "當月使用總量": 0, "購入量": 0, "過期報銷": 0, "公藥使用": 0,
            "期末剩餘量": stock, "實體盤點數量": stock, "有效期限": expiry,
            "9月消耗量": 0, "10月消耗量": 0, "11月消耗量": 0, "12月消耗量": 0, "1月消耗量": 0,
            "全學期使用總量": 0
        })
        report_rows.append(r_dict)

    df_report = pd.DataFrame(report_rows)

    # 線上預覽表格
    st.subheader(f"📄 {selected_year} 上學期用藥月報表 ({selected_month}) 預覽")
    st.dataframe(df_report, use_container_width=True, hide_index=True)

    # 動態生成 openpyxl Excel 檔案供一鍵下載
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"115年{selected_month}用藥月報表"

    # Row 1: 大標題
    title_text = f"國立臺北大學衛保組 {selected_year}上學期藥品使用月報與全學期統計表 ({selected_year}9月起)"
    ws.append([title_text])
    
    # Row 2: 37 欄標準表頭
    excel_headers = [
        "藥品名稱\n(商品名/中文)", "115年8月\n剩餘量",
        *days,
        "當月使用\n總量", "購入量", "過期報銷", "公藥使用", "115年9月\n期末剩餘量", "實體盤點\n數量", "有效期限",
        "9月\n消耗量", "10月\n消耗量", "11月\n消耗量", "12月\n消耗量", "1月\n消耗量", "全學期\n使用總量"
    ]
    ws.append(excel_headers)

    # Row 3+: 填入藥品資料與自動公式
    for idx, row in df_report.iterrows():
        r_idx = idx + 3
        data_row = [
            row["藥品名稱\n(商品名/中文)"], row["上月剩餘量"],
            *[0]*len(days),
            f"=SUM(C{r_idx}:X{r_idx})", # 當月使用總量
            0, 0, 0,                   # 購入量、過期報銷、公藥使用
            f"=B{r_idx}+Z{r_idx}-Y{r_idx}-AA{r_idx}-AB{r_idx}", # 期末剩餘量公式
            f"=AC{r_idx}",              # 實體盤點數量
            row["有效期限"],
            f"=Y{r_idx}",              # 9月消耗量
            0, 0, 0, 0,                # 10月~1月消耗量
            f"=SUM(AF{r_idx}:AJ{r_idx})" # 全學期使用總量公式
        ]
        ws.append(data_row)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 一鍵下載標準 Excel 月報表 (.xlsx)",
        data=output,
        file_name=f"國立臺北大學衛保組_{selected_year}_上學期用藥月報與學期統計表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
