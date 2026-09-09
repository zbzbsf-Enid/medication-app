import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 頁面基本設定
st.set_page_config(
    page_title="國立臺北大學衛保組 - 藥品管理與月報系統",
    page_icon="💊",
    layout="wide"
)

# 高對比度與日曆專用 CSS 修正
st.markdown("""
    <style>
    /* 1. 全域背景與字型 (微軟正黑體 16px) */
    html, body, [class*="css"], .stApp {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        font-size: 16px !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }

    /* 2. 強制所有文字、標籤與標題為深色高對比 */
    p, span, label, h1, h2, h3, h4, .stMarkdown, div[data-testid="stMarkdownContainer"] * {
        color: #0f172a !important;
        opacity: 1 !important;
    }

    /* 3. 側邊欄樣式與選項強化 */
    section[data-testid="stSidebar"] {
        background-color: #e2e8f0 !important;
        border-right: 2px solid #cbd5e1 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #0f172a !important;
        font-weight: 500 !important;
    }

    /* 4. 下拉選單與普通輸入框樣式 */
    div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #94a3b8 !important;
        border-radius: 8px !important;
    }
    div[data-baseweb="select"] * {
        color: #0f172a !important;
    }

    /* 5. 日期輸入框本體 (stDateInput) 強制白底黑字 */
    div[data-testid="stDateInput"] {
        background-color: transparent !important;
    }
    div[data-testid="stDateInput"] div[data-baseweb="input"] {
        background-color: #ffffff !important;
        border: 1px solid #94a3b8 !important;
        border-radius: 6px !important;
    }
    div[data-testid="stDateInput"] input {
        background-color: #ffffff !important;
        color: #0f172a !important;
        font-weight: 600 !important;
    }

    /* 6. 徹底修正日曆彈窗 (Popover & Calendar) 高對比白底黑字 */
    div[data-baseweb="popover"],
    div[data-baseweb="calendar"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15) !important;
    }

    /* 強制日曆內部所有文字（星期、月份、數字）為深黑色 */
    div[data-baseweb="calendar"] *,
    div[data-baseweb="calendar"] button,
    div[data-baseweb="calendar"] div {
        color: #0f172a !important;
        background-color: #ffffff !important;
    }

    /* 日曆日期 hover 效果 */
    div[data-baseweb="calendar"] button:hover {
        background-color: #e2e8f0 !important;
    }

    /* 被選中的日期（標示為綠色背景 + 純白文字） */
    div[data-baseweb="calendar"] [aria-selected="true"],
    div[data-baseweb="calendar"] [aria-selected="true"] * {
        background-color: #0d9488 !important;
        color: #ffffff !important;
        font-weight: bold !important;
        border-radius: 50% !important;
    }

    /* 7. 表單與按鈕樣式 */
    div[data-testid="stForm"] {
        background-color: #ffffff !important;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        border: 1px solid #cbd5e1 !important;
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
    }
    .stButton > button:hover, div[data-testid="stForm"] button:hover {
        background-color: #0f766e !important;
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

st.sidebar.title("📌 功能選單")
if st.sidebar.button("🔄 手動刷新雲端資料"):
    st.cache_data.clear()
    st.rerun()

menu = st.sidebar.radio(
    "請選擇功能頁面", 
    ["💊 多項藥品領用登記", "📥 藥品進貨/補貨登記", "📦 當前庫存總覽", "📊 用藥月報與學期統計表"]
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
        placeholder="點擊或輸入藥品名稱關鍵字..."
    )

    if selected_items:
        st.markdown("---")
        st.subheader("✏️ 請輸入領用資訊與數量")
        with st.form("batch_checkout_form"):
            col_d1, col_d2 = st.columns([1, 2])
            with col_d1:
                record_date = st.date_input(
                    "📅 實際領用/補登日期：", 
                    value=datetime.now().date(),
                    help="若為補登昨日或過去的紀錄，請在此修改日期"
                )
            with col_d2:
                remarks = st.text_input("領用備註/用途：", placeholder="例如：衛保組公用 / 門診備用")

            st.markdown("<hr style='margin: 12px 0;'>", unsafe_allow_html=True)

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

            submit_btn = st.form_submit_button("✅ 完成登記並更新庫存")

            if submit_btn:
                log_time_str = f"{record_date.strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M:%S')}"
                new_logs = []
                for item, qty in quantities.items():
                    idx = df_inventory[df_inventory['display_name'] == item].index[0]
                    old_stock = int(df_inventory.loc[idx, '目前庫存'])
                    new_stock = max(0, old_stock - qty)
                    df_inventory.loc[idx, '目前庫存'] = new_stock
                    new_logs.append({
                        "領用時間": log_time_str,
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
                    st.success(f"🎉 領用登記成功！紀錄日期為 `{record_date}`，庫存與 Log 已更新。")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 更新失敗：{e}")

# 頁面 2：新增藥品進貨/補貨登記
elif menu == "📥 藥品進貨/補貨登記":
    st.header("📥 藥品購入與進貨登記")
    st.caption("在此輸入買入的藥品數量、新批號與有效期限，系統將自動累加庫存。")

    df_inventory['display_name'] = (
        df_inventory['藥品名稱(英文)'].fillna('') + " (" + 
        df_inventory['中文名稱'].fillna('') + ")"
    )
    med_list = df_inventory['display_name'].tolist()

    selected_med = st.selectbox("請選擇進貨藥品：", options=med_list)

    if selected_med:
        row_info = df_inventory[df_inventory['display_name'] == selected_med].iloc[0]
        
        st.info(f"📌 當前庫存：`{row_info['目前庫存']}` | 目前批號：`{row_info['批號']}` | 目前有效期限：`{row_info['有效期限']}`")

        with st.form("purchase_inbound_form"):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                inbound_date = st.date_input("📅 進貨日期：", value=datetime.now().date())
                purchase_qty = st.number_input("📦 購入數量：", min_value=1, value=100, step=1)
            with col_p2:
                new_batch = st.text_input("🏷️ 新藥品批號：", value=str(row_info['批號']) if pd.notna(row_info['批號']) else "")
                new_expiry = st.date_input("⏳ 新有效期限：", value=datetime.now().date())

            vendor_remark = st.text_input("🏢 廠商/採購備註：", placeholder="例如：衛福部撥發 / 某某藥局採購")

            submit_purchase = st.form_submit_button("✅ 確認進貨並更新庫存與批號")

            if submit_purchase:
                idx = df_inventory[df_inventory['display_name'] == selected_med].index[0]
                old_qty = int(df_inventory.loc[idx, '目前庫存'])
                new_qty = old_qty + int(purchase_qty)
                expiry_str = new_expiry.strftime("%Y-%m-%d")
                inbound_time_str = f"{inbound_date.strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M:%S')}"

                # 1. 更新庫存主表
                df_inventory.loc[idx, '目前庫存'] = new_qty
                df_inventory.loc[idx, '批號'] = new_batch
                df_inventory.loc[idx, '有效期限'] = expiry_str

                # 2. 建立進貨 Log
                purchase_log = {
                    "進貨時間": inbound_time_str,
                    "藥品名稱": df_inventory.loc[idx, '藥品名稱(英文)'],
                    "中文名稱": df_inventory.loc[idx, '中文名稱'],
                    "購入數量": purchase_qty,
                    "新批號": new_batch,
                    "有效期限": expiry_str,
                    "更新後總庫存": new_qty,
                    "備註": vendor_remark
                }

                try:
                    df_save = df_inventory.drop(columns=['display_name'])
                    conn.update(worksheet="庫存", data=df_save)
                    
                    try:
                        try:
                            df_inbound_existing = conn.read(worksheet="進貨紀錄", ttl=0)
                            df_inbound_updated = pd.concat([df_inbound_existing, pd.DataFrame([purchase_log])], ignore_index=True)
                        except Exception:
                            df_inbound_updated = pd.DataFrame([purchase_log])
                        conn.update(worksheet="進貨紀錄", data=df_inbound_updated)
                    except Exception:
                        st.warning("⚠️ 庫存已成功更新！但寫入『進貨紀錄』分頁失敗，請確認 Google 試算表中是否已建立名為 『進貨紀錄』 的分頁。")

                    st.success(f"🎉 進貨完成！`{selected_med}` 庫存已由 {old_qty} 增加至 {new_qty}，批號更新為 `{new_batch}`。")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 庫存更新失敗：{e}")

# 頁面 3：當前庫存總覽
elif menu == "📦 當前庫存總覽":
    st.header("📦 當前藥品庫存總覽")
    st.dataframe(df_inventory.drop(columns=['display_name'], errors='ignore'), use_container_width=True, hide_index=True)

# 頁面 4：用藥月報與學期統計表
elif menu == "📊 用藥月報與學期統計表":
    st.header("📊 國立臺北大學衛保組 藥品使用月報與全學期統計表")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_year = st.selectbox("學年度/年份", ["115學年度", "114學年度"], index=0)
    with col_sel2:
        selected_month = st.selectbox("統計月份", ["9月", "10月", "11月", "12月", "1月"], index=0)

    days = [f"9/{d}" for d in [1, 2, 3, 4, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29, 30]]
    
    report_rows = []
    for _, row in df_inventory.iterrows():
        eng_name = str(row.get('藥品名稱(英文)', ''))
        cht_name = str(row.get('中文名稱', ''))
        combined_name = f"{eng_name} ({cht_name})" if cht_name else eng_name
        stock = int(row.get('目前庫存', 0))
        expiry = str(row.get('有效期限', ''))

        r_dict = {"藥品名稱\n(商品名/中文)": combined_name, "115年8月\n剩餘量": stock}
        for d in days:
            r_dict[d] = 0
        r_dict.update({
            "當月使用\n總量": 0, "購入量": 0, "過期報銷": 0, "公藥使用": 0,
            "115年9月\n期末剩餘量": stock, "實體盤點\n數量": stock, "有效期限": expiry,
            "9月\n消耗量": 0, "10月\n消耗量": 0, "11月\n消耗量": 0, "12月\n消耗量": 0, "1月\n消耗量": 0,
            "全學期\n使用總量": 0
        })
        report_rows.append(r_dict)

    df_report = pd.DataFrame(report_rows)
    st.subheader(f"📄 {selected_year} 上學期用藥月報表 ({selected_month}) 預覽")
    st.dataframe(df_report, use_container_width=True, hide_index=True)

    # 產生 Excel 檔
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"115年{selected_month}用藥月報表"

    title_text = f"國立臺北大學衛保組 {selected_year}上學期藥品使用月報與全學期統計表 ({selected_year}9月起)"
    ws.append([title_text])
    ws.cell(row=1, column=1).font = Font(name="微軟正黑體", size=13, bold=True, color="1F4E78")

    excel_headers = [
        "藥品名稱\n(商品名/中文)", "115年8月\n剩餘量",
        *days,
        "當月使用\n總量", "購入量", "過期報銷", "公藥使用", "115年9月\n期末剩餘量", "實體盤點\n數量", "有效期限",
        "9月\n消耗量", "10月\n消耗量", "11月\n消耗量", "12月\n消耗量", "1月\n消耗量", "全學期\n使用總量"
    ]
    ws.append(excel_headers)

    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_blue = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    fill_orange = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")
    fill_data_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_data_orange = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    font_white_bold = Font(name="微軟正黑體", size=9, bold=True, color="FFFFFF")
    font_gray_bold = Font(name="微軟正黑體", size=9, bold=True, color="333333")
    font_red = Font(name="微軟正黑體", size=9, color="C00000")
    font_navy = Font(name="微軟正黑體", size=9, color="002060")
    font_orange = Font(name="微軟正黑體", size=9, color="C65911")
    font_default = Font(name="微軟正黑體", size=9)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, 38):
        cell = ws.cell(row=2, column=col_idx)
        cell.alignment = align_center
        cell.border = thin_border
        if col_idx in [1, 2]:
            cell.fill = fill_navy
            cell.font = font_white_bold
        elif 3 <= col_idx <= 24:
            cell.fill = fill_gray
            cell.font = font_gray_bold
        elif 25 <= col_idx <= 31:
            cell.fill = fill_blue
            cell.font = font_white_bold
        else:
            cell.fill = fill_orange
            cell.font = font_white_bold

    for idx, row in df_report.iterrows():
        r_idx = idx + 3
        data_row = [
            row["藥品名稱\n(商品名/中文)"], row["115年8月\n剩餘量"],
            *[0]*len(days),
            f"=SUM(C{r_idx}:X{r_idx})",
            0, 0, 0,
            f"=B{r_idx}+Z{r_idx}-Y{r_idx}-AA{r_idx}-AB{r_idx}",
            f"=AC{r_idx}",
            row["有效期限"],
            f"=Y{r_idx}",
            0, 0, 0, 0,
            f"=SUM(AF{r_idx}:AJ{r_idx})"
        ]
        ws.append(data_row)

        for c_idx in range(1, 38):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.alignment = align_center if c_idx > 1 else align_left
            cell.font = font_default

            if c_idx in [2, 29]:
                cell.font = font_red
            elif c_idx == 25:
                cell.fill = fill_data_yellow
                cell.font = font_navy
            elif c_idx == 37:
                cell.fill = fill_data_orange
                cell.font = font_orange

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    for c in range(3, 25):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 5.5
    ws.column_dimensions['Y'].width = 12
    for c in range(26, 31):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 12
    ws.column_dimensions['AE'].width = 14
    for c in range(32, 38):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 12

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 一鍵下載標準 Excel 月報表 (.xlsx)",
        data=output,
        file_name=f"國立臺北大學衛保組_{selected_year}_上學期用藥月報與學期統計表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
