import streamlit as st
import pandas as pd
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
# 1. 雲端 Google Sheets 連線與安全讀取防護
# -----------------------------------------------------------------------------
@st.cache_resource(ttl=60)
def get_connection():
    return st.connection("gsheets", type=GSheetsConnection)

conn = get_connection()

def load_data():
    try:
        # 讀取『庫存』工作表
        df_inventory = conn.read(worksheet="庫存", ttl=0)
        
        # 讀取『領用紀錄』工作表 (若不存在則建立空表格)
        try:
            df_logs = conn.read(worksheet="領用紀錄", ttl=0)
        except Exception:
            df_logs = pd.DataFrame(columns=['日期', '藥品名稱', '批號', '領用數量', '備註'])

        # 🚨 安全防護 1：若表格讀取為空，直接停止，保護雲端資料不被覆蓋
        if df_inventory is None or df_inventory.empty:
            st.error("❌ 讀取『庫存』工作表失敗或資料為空，請檢查 Google Sheet 權限與工作表名稱！")
            st.stop()

        # 🚨 安全防護 2：檢查必要欄位是否存在
        required_cols = ['藥品名稱', '批號', '有效日期', '現有庫存']
        for col in required_cols:
            if col not in df_inventory.columns:
                st.error(f"❌ 庫存工作表缺乏必要欄位：『{col}』，請檢查 Google Sheet 標頭！")
                st.stop()

        # 格式轉型
        df_inventory['現有庫存'] = pd.to_numeric(df_inventory['現有庫存'], errors='coerce').fillna(0).astype(int)
        
        return df_inventory, df_logs

    except Exception as e:
        st.error(f"❌ 連線失敗或無法讀取資料：{e}")
        st.stop()

df_inventory, df_logs = load_data()

# -----------------------------------------------------------------------------
# 2. 先進先出 (FIFO) 庫存扣減函數
# -----------------------------------------------------------------------------
def deduct_inventory_fifo(inventory_df, drug_name, req_qty, log_date, note=""):
    """
    先進先出 (FIFO) 邏輯：按「有效日期」舊到新進行扣減
    """
    df_inv = inventory_df.copy()
    
    # 1. 搜尋該藥品且庫存 > 0 的批號
    mask = (df_inv['藥品名稱'] == drug_name) & (df_inv['現有庫存'] > 0)
    available = df_inv[mask].copy()

    if available.empty:
        return df_inv, [], f"❌ {drug_name} 目前已無可用庫存！"

    # 2. 按「有效日期」排序 (舊效期/舊進貨排前面先扣)
    available['exp_dt'] = pd.to_datetime(available['有效日期'], errors='coerce')
    available = available.sort_values(by='exp_dt', ascending=True)

    remaining_qty = req_qty
    new_logs = []

    # 3. 逐筆批號扣除庫存
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

        # 記錄每筆實際被扣除的批號與數量
        new_logs.append({
            '日期': str(log_date),
            '藥品名稱': drug_name,
            '批號': batch_no,
            '領用數量': deducted,
            '備註': note
        })

    status_msg = "SUCCESS"
    if remaining_qty > 0:
        status_msg = f"⚠️ {drug_name} 庫存不足，尚有 {remaining_qty} 顆未完成扣減。"

    return df_inv, new_logs, status_msg

# -----------------------------------------------------------------------------
# 3. 側邊欄與選單
# -----------------------------------------------------------------------------
st.sidebar.title("📌 功能選單")

if st.sidebar.button("🔄 手動刷新雲端資料"):
    st.cache_resource.clear()
    st.rerun()

menu = st.sidebar.radio(
    "請選擇功能頁面",
    ["📋 多項藥品領用登記", "📊 當前庫存總覽", "🗓️ 用藥月報表與統計"]
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

    # 取得去重後的藥品名稱選單
    drug_list = sorted(df_inventory['藥品名稱'].dropna().unique().tolist())
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

            # 執行 FIFO 庫存扣減
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
                try:
                    # 1. 寫入『庫存』工作表
                    conn.update(worksheet="庫存", data=updated_inv)

                    # 2. 寫入『領用紀錄』工作表
                    logs_df_new = pd.concat([df_logs, pd.DataFrame(all_new_logs)], ignore_index=True)
                    conn.update(worksheet="領用紀錄", data=logs_df_new)

                    st.success("✅ 庫存更新成功！已精準扣除舊效期庫存並同步至雲端。")
                    st.cache_resource.clear()
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ 庫存更新失敗：{e}")

# -----------------------------------------------------------------------------
# 頁面 2：當前庫存總覽
# -----------------------------------------------------------------------------
elif menu == "📊 當前庫存總覽":
    st.header("📊 當前庫存總覽")
    st.dataframe(df_inventory, use_container_width=True)

# -----------------------------------------------------------------------------
# 頁面 3：用藥月報表與統計 (按 [藥名 + 批號] 精準統計)
# -----------------------------------------------------------------------------
elif menu == "🗓️ 用藥月報表與統計":
    st.header("🗓️ 用藥月報表與學期統計")

    if df_logs.empty:
        st.info("目前尚無任何領用紀錄。")
    else:
        # 資料預處理
        df_logs['日期_str'] = pd.to_datetime(df_logs['日期'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_logs['領用數量'] = pd.to_numeric(df_logs['領用數量'], errors='coerce').fillna(0)

        # 基礎欄位
        report_df = df_inventory[['藥品名稱', '批號', '有效日期', '現有庫存']].copy()
        
        # 取得所有發生過領用的不重複日期
        dates = sorted(df_logs['日期_str'].dropna().unique().tolist())

        # 🟢 關鍵精準統計：每一列依據 [藥品名稱] 與 [批號] 計算各日期領用量
        for d in dates:
            daily_quantities = []
            for _, row in report_df.iterrows():
                drug = row['藥品名稱']
                batch = row['批號']

                # 比對「藥名 + 批號 + 日期」
                mask = (
                    (df_logs['藥品名稱'] == drug) & 
                    (df_logs['批號'] == batch) & 
                    (df_logs['日期_str'] == d)
                )
                qty = df_logs[mask]['領用數量'].sum()
                daily_quantities.append(int(qty))

            report_df[d] = daily_quantities

        # 計算累計使用量
        report_df['當月使用總量'] = report_df[dates].sum(axis=1) if dates else 0

        st.dataframe(report_df, use_container_width=True)
