import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime

# 1. 頁面基本設定
st.set_page_config(
    page_title="衛保組管理系統",
    page_icon="💊",
    layout="wide"
)

st.title("💊 衛保組藥品管理系統")

# 2. 建立 Google Sheets 雲端連線 (讀取 Secrets)
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"❌ 無法建立 Google 連線，請檢查 Streamlit Cloud 的 Secrets 設定。錯誤訊息: {e}")
    st.stop()

# 3. 資料載入與處理解析函式
@st.cache_data(ttl=5)
def load_inventory_data():
    try:
        # 優先讀取名為「庫存」的工作表，若不存在則讀取第一個 Sheet
        df = conn.read(worksheet="庫存", ttl=0)
    except Exception:
        df = conn.read(ttl=0)
    
    # 欄位檢查與格式修正
    required_cols = ["藥品名稱(英文)", "中文名稱", "目前庫存", "有效期限", "用途/備註"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""
            
    df["目前庫存"] = pd.to_numeric(df["目前庫存"], errors="coerce").fillna(0).astype(int)
    return df[required_cols]

def load_log_data():
    try:
        df = conn.read(worksheet="領用紀錄", ttl=0)
    except Exception:
        df = pd.DataFrame(columns=["領用時間", "藥品名稱", "中文名稱", "領用數量", "剩餘庫存", "備註"])
    return df

def save_inventory_data(df):
    try:
        conn.update(worksheet="庫存", data=df)
    except Exception:
        conn.update(data=df)

def save_log_data(df_log):
    try:
        conn.update(worksheet="領用紀錄", data=df_log)
    except Exception:
        pass

# 4. 讀取最新庫存資料
try:
    df_inventory = load_inventory_data()
except Exception as e:
    st.error(f"❌ 讀取試算表資料失敗，請確認試算表是否有共用給服務帳號 Email。錯誤：{e}")
    st.stop()

# 5. 側邊欄選單
st.sidebar.title("📌 功能選單")
page = st.sidebar.radio(
    "請選擇功能頁面",
    ["💊 藥品領用與紀錄", "📦 庫存盤點與校正", "☁️ 雲端報表匯出"]
)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 手動刷新雲端最新資料"):
    st.cache_data.clear()
    st.rerun()

# -----------------------------------------------------------------------------
# 功能頁面 1：藥品領用與紀錄
# -----------------------------------------------------------------------------
if page == "💊 藥品領用與紀錄":
    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.subheader("💊 藥品領用與登記")
        st.caption("點選下方搜尋欄可選擇一種或多種藥品，設定數量後即可一次完成登記與庫存扣減。")

        options = [f"{row['藥品名稱(英文)']} | {row['中文名稱']}" for _, row in df_inventory.iterrows()]
        selected_meds = st.multiselect(
            "選擇本次領取的所有藥品 (可同時選擇多項)",
            options=options,
            placeholder="請點擊或輸入藥名/中文名稱進行搜尋..."
        )

        if selected_meds:
            st.write("### 📝 設定領取數量")
            inputs = {}
            with st.form("dispense_form"):
                for item in selected_meds:
                    med_eng = item.split(" | ")[0]
                    current_stock = df_inventory.loc[df_inventory["藥品名稱(英文)"] == med_eng, "目前庫存"].values[0]
                    
                    num = st.number_input(
                        f"【{item}】領用數量 (目前庫存: {current_stock})",
                        min_value=1,
                        max_value=int(current_stock) if current_stock > 0 else 1,
                        value=1,
                        key=f"num_{med_eng}"
                    )
                    inputs[med_eng] = num
                
                note = st.text_input("領用備註 (選填，例如：學生領用/傷口處置)", value="")
                submit_button = st.form_submit_button("✅ 確認登記領用")

                if submit_button:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    df_logs = load_log_data()

                    new_logs = []
                    for med_eng, qty in inputs.items():
                        idx = df_inventory[df_inventory["藥品名稱(英文)"] == med_eng].index[0]
                        med_chi = df_inventory.loc[idx, "中文名稱"]
                        cur_qty = df_inventory.loc[idx, "目前庫存"]
                        
                        new_qty = max(0, cur_qty - qty)
                        df_inventory.loc[idx, "目前庫存"] = new_qty
                        
                        new_logs.append({
                            "領用時間": now_str,
                            "藥品名稱": med_eng,
                            "中文名稱": med_chi,
                            "領用數量": qty,
                            "剩餘庫存": new_qty,
                            "備註": note
                        })

                    # 更新寫回雲端試算表
                    save_inventory_data(df_inventory)
                    if new_logs:
                        df_new_logs = pd.DataFrame(new_logs)
                        df_logs = pd.concat([df_logs, df_new_logs], ignore_index=True)
                        save_log_data(df_logs)

                    st.success("🎉 領用登記成功！資料已同步寫入 Google 雲端試算表。")
                    st.cache_data.clear()
                    st.rerun()
        else:
            st.info("💡 請先在上方的選單中點選或搜尋要領取的藥品。")

    with col2:
        st.subheader("📋 當前藥品庫存總覽")
        st.dataframe(df_inventory, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# 功能頁面 2：庫存盤點與校正
# -----------------------------------------------------------------------------
elif page == "📦 庫存盤點與校正":
    st.subheader("📦 庫存盤點與數字校正")
    st.caption("您可以直接雙擊下方表格中的欄位修改數值（例如盤點後的實際數量），完成後點擊「儲存修改」寫回雲端。")

    edited_df = st.data_editor(
        df_inventory,
        num_rows="dynamic",
        use_container_width=True,
        key="inventory_editor"
    )

    if st.button("💾 儲存修改至 Google 雲端試算表"):
        save_inventory_data(edited_df)
        st.success("✅ 庫存校正成功！雲端資料庫已更新。")
        st.cache_data.clear()
        st.rerun()

# -----------------------------------------------------------------------------
# 功能頁面 3：雲端報表匯出
# -----------------------------------------------------------------------------
elif page == "☁️ 雲端報表匯出":
    st.subheader("☁️ 領用歷史紀錄與報表")
    try:
        df_logs = load_log_data()
        st.dataframe(df_logs, use_container_width=True, hide_index=True)

        csv_data = df_logs.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載領用歷史紀錄 (CSV 檔)",
            data=csv_data,
            file_name=f"medication_logs_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    except Exception as e:
        st.warning(f"目前無領用紀錄或讀取失敗：{e}")
