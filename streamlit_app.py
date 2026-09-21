import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ==========================================
# 頁面基本設定
# ==========================================
st.set_page_config(
    page_title="國立臺北大學衛保組 藥品管理系統",
    page_icon="💊",
    layout="wide"
)

# ==========================================
# Google Sheets 資料讀取與寫入函式
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    """從 Google Sheets 讀取『庫存』與『異動紀錄』兩個工作表"""
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    # 讀取庫存資料表
    try:
        df_inventory = conn.read(worksheet="Inventory", ttl=0)
    except Exception:
        # 若工作表不存在或空白，初始化預設欄位
        df_inventory = pd.DataFrame(columns=[
            "藥品代碼", "藥品名稱", "規格", "單位", "目前庫存", "安全庫存", "分類"
        ])
    
    # 讀取異動紀錄表
    try:
        df_logs = conn.read(worksheet="Logs", ttl=0)
    except Exception:
        df_logs = pd.DataFrame(columns=[
            "時間", "類型", "對象/系所", "經手人", "藥品名稱", "變動數量", "備註"
        ])

    # 確保數值型別正確
    if not df_inventory.empty:
        df_inventory["目前庫存"] = pd.to_numeric(df_inventory["目前庫存"], errors="coerce").fillna(0).astype(int)
        df_inventory["安全庫存"] = pd.to_numeric(df_inventory["安全庫存"], errors="coerce").fillna(0).astype(int)

    return df_inventory, df_logs


def save_inventory(df_inventory):
    """更新庫存資料至 Google Sheets"""
    conn = st.connection("gsheets", type=GSheetsConnection)
    conn.update(worksheet="Inventory", data=df_inventory)
    st.cache_data.clear()


def save_logs(df_logs):
    """更新異動紀錄至 Google Sheets"""
    conn = st.connection("gsheets", type=GSheetsConnection)
    conn.update(worksheet="Logs", data=df_logs)
    st.cache_data.clear()


# 載入資料
try:
    df_inventory, df_logs = load_data()
except Exception as e:
    st.error(f"⚠️ 連線至 Google Sheets 時發生錯誤：{e}")
    st.info("請檢查 Secrets 設定中的 [connections.gsheets] 憑證資訊是否正確。")
    st.stop()


# ==========================================
# 側邊欄設計
# ==========================================
st.sidebar.title("📌 功能選單")

# 手動刷新按鈕
if st.sidebar.button("🔄 手動刷新雲端資料", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")

page = st.sidebar.radio(
    "請選擇功能頁面",
    [
        "💊 多項藥品領用登記",
        "🏥 藥品進貨/建檔登記",
        "🛠️ 紀錄修改與庫存微調",
        "📦 當前庫存總覽",
        "📊 用藥月報與學期統計表"
    ]
)

# 頁面主標題
st.title("💊 國立臺北大學衛保組 藥品管理系統")

# ==========================================
# 頁面 1：多項藥品領用登記
# ==========================================
if page == "💊 多項藥品領用登記":
    st.subheader("📋 領用登記單")

    if df_inventory.empty:
        st.warning("目前庫存無任何藥品，請先至『藥品進貨/建檔登記』建立藥品資料。")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            recipient = st.text_input("領用對象 / 學號 / 姓名", placeholder="例如：張同學 (41112345)")
        with col2:
            department = st.text_input("單位 / 系所", placeholder="例如：企管系 / 行政人員")
        with col3:
            handler = st.text_input("經手人 (護理師)", value="衛保組護理師")

        st.markdown("##### 選擇領用藥品與數量")
        
        # 選擇領用藥品
        selected_drugs = st.multiselect(
            "請選擇要領用的藥品（可多選）：",
            options=df_inventory["藥品名稱"].tolist()
        )

        items_to_issue = []
        if selected_drugs:
            for drug in selected_drugs:
                row = df_inventory[df_inventory["藥品名稱"] == drug].iloc[0]
                current_stock = row["目前庫存"]
                unit = row["單位"]

                c1, c2, c3 = st.columns([3, 2, 3])
                with c1:
                    st.write(f"**{drug}** （單位: {unit}）")
                with c2:
                    st.caption(f"目前庫存：{current_stock}")
                with c3:
                    qty = st.number_input(
                        f"領用數量 ({drug})",
                        min_value=1,
                        max_value=max(1, int(current_stock)),
                        value=1,
                        key=f"issue_{drug}"
                    )
                    items_to_issue.append({"藥品名稱": drug, "數量": qty, "當前庫存": current_stock})

            note = st.text_input("備註說明", placeholder="例如：急性頭痛領藥")

            if st.button("🚀 確認送出領用紀錄", type="primary"):
                # 驗證庫存
                has_error = False
                for item in items_to_issue:
                    if item["數量"] > item["當前庫存"]:
                        st.error(f"❌ {item['藥品名稱']} 庫存不足！(剩餘 {item['當前庫存']})")
                        has_error = True

                if not has_error:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_logs = []

                    # 更新庫存並紀錄
                    for item in items_to_issue:
                        drug_name = item["藥品名稱"]
                        qty = item["數量"]

                        # 扣減庫存
                        df_inventory.loc[df_inventory["藥品名稱"] == drug_name, "目前庫存"] -= qty

                        # 新增日誌
                        new_logs.append({
                            "時間": now_str,
                            "類型": "領用",
                            "對象/系所": f"{recipient} ({department})",
                            "經手人": handler,
                            "藥品名稱": drug_name,
                            "變動數量": -qty,
                            "備註": note
                        })

                    # 寫入雲端
                    save_inventory(df_inventory)
                    df_new_logs = pd.concat([df_logs, pd.DataFrame(new_logs)], ignore_index=True)
                    save_logs(df_new_logs)

                    st.success("✅ 領用登記成功！庫存已同步更新。")
                    st.rerun()

# ==========================================
# 頁面 2：藥品進貨/建檔登記
# ==========================================
elif page == "🏥 藥品進貨/建檔登記":
    st.subheader("🏥 藥品進貨與新品建檔")
    tab1, tab2 = st.tabs(["📦 已有藥品進貨", "✨ 新增藥品品項"])

    # Tab 1: 已有藥品進貨
    with tab1:
        if df_inventory.empty:
            st.info("目前尚無藥品，請使用「新增藥品品項」建立第一筆藥品。")
        else:
            with st.form("restock_form"):
                drug_to_add = st.selectbox("選擇進貨藥品", options=df_inventory["藥品名稱"].tolist())
                add_qty = st.number_input("進貨數量", min_value=1, value=10)
                handler = st.text_input("經手人", value="衛保組護理師")
                supplier = st.text_input("廠商 / 來源", placeholder="例如：中央健康保險局 / 某某藥局")
                submit_restock = st.form_submit_button("📥 寫入進貨紀錄")

                if submit_restock:
                    # 增加庫存
                    df_inventory.loc[df_inventory["藥品名稱"] == drug_to_add, "目前庫存"] += add_qty
                    
                    # 記錄日誌
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_log = {
                        "時間": now_str,
                        "類型": "進貨",
                        "對象/系所": supplier,
                        "經手人": handler,
                        "藥品名稱": drug_to_add,
                        "變動數量": add_qty,
                        "備註": "藥品補充進貨"
                    }
                    df_new_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)

                    save_inventory(df_inventory)
                    save_logs(df_new_logs)
                    st.success(f"✅ {drug_to_add} 已成功進貨 {add_qty} 單位！")
                    st.rerun()

    # Tab 2: 新增藥品品項
    with tab2:
        with st.form("new_drug_form"):
            code = st.text_input("藥品代碼", placeholder="例如：MED001")
            name = st.text_input("藥品名稱 *", placeholder="例如：普拿疼 (Panadol)")
            spec = st.text_input("規格", placeholder="例如：500mg/錠")
            unit = st.text_input("單位", placeholder="例如：錠、顆、瓶")
            init_stock = st.number_input("初始庫存量", min_value=0, value=0)
            safe_stock = st.number_input("安全庫存警戒值", min_value=1, value=20)
            category = st.selectbox("分類", ["解熱鎮痛", "腸胃用藥", "外用藥品", "呼吸道用藥", "包材/衛材", "其他"])
            
            submit_new_drug = st.form_submit_button("✨ 建立新藥品")

            if submit_new_drug:
                if not name:
                    st.error("❌ 藥品名稱為必填欄位！")
                elif name in df_inventory["藥品名稱"].values:
                    st.error("❌ 該藥品名稱已存在，請直接切換至「已有藥品進貨」進行補充。")
                else:
                    new_row = {
                        "藥品代碼": code,
                        "藥品名稱": name,
                        "規格": spec,
                        "單位": unit,
                        "目前庫存": init_stock,
                        "安全庫存": safe_stock,
                        "分類": category
                    }
                    df_updated_inv = pd.concat([df_inventory, pd.DataFrame([new_row])], ignore_index=True)
                    
                    # 初始庫存異動紀錄
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_log = {
                        "時間": now_str,
                        "類型": "建檔",
                        "對象/系所": "系統建檔",
                        "經手人": "衛保組護理師",
                        "藥品名稱": name,
                        "變動數量": init_stock,
                        "備註": "新藥品建檔"
                    }
                    df_new_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)

                    save_inventory(df_updated_inv)
                    save_logs(df_new_logs)
                    st.success(f"🎉 成功新增藥品：{name}！")
                    st.rerun()

# ==========================================
# 頁面 3：紀錄修改與庫存微調
# ==========================================
elif page == "🛠️ 紀錄修改與庫存微調":
    st.subheader("🛠️ 庫存盤點與微調")
    
    if df_inventory.empty:
        st.info("目前尚無藥品可供微調。")
    else:
        with st.form("adjust_form"):
            selected_drug = st.selectbox("選擇要微調的藥品", options=df_inventory["藥品名稱"].tolist())
            current_val = int(df_inventory[df_inventory["藥品名稱"] == selected_drug]["目前庫存"].values[0])
            
            st.info(f"該藥品目前紀錄庫存為：**{current_val}**")
            new_val = st.number_input("校正後的真實庫存數量", min_value=0, value=current_val)
            reason = st.text_input("微調原因 / 盤點備註", placeholder="例如：定期盤點損耗、過期報廢")
            handler = st.text_input("經手人", value="衛保組護理師")

            submit_adjust = st.form_submit_button("⚠️ 確認校正庫存")

            if submit_adjust:
                diff = new_val - current_val
                if diff == 0:
                    st.warning("數量未變更，未進行任何更新。")
                else:
                    df_inventory.loc[df_inventory["藥品名稱"] == selected_drug, "目前庫存"] = new_val
                    
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    new_log = {
                        "時間": now_str,
                        "類型": "盤點微調",
                        "對象/系所": "盤點校正",
                        "經手人": handler,
                        "藥品名稱": selected_drug,
                        "變動數量": diff,
                        "備註": reason if reason else "盤點微調"
                    }
                    df_new_logs = pd.concat([df_logs, pd.DataFrame([new_log])], ignore_index=True)

                    save_inventory(df_inventory)
                    save_logs(df_new_logs)
                    st.success(f"✅ {selected_drug} 庫存已校正為 {new_val}！")
                    st.rerun()

    st.markdown("---")
    st.subheader("📜 歷史異動紀錄總覽")
    if not df_logs.empty:
        st.dataframe(df_logs.sort_values(by="時間", ascending=False), use_container_width=True)
    else:
        st.info("目前無任何歷史異動紀錄。")

# ==========================================
# 頁面 4：當前庫存總覽
# ==========================================
elif page == "📦 當前庫存總覽":
    st.subheader("📦 當前庫存總覽與預警")

    if df_inventory.empty:
        st.info("目前庫存為空。")
    else:
        # 指標卡片
        total_items = len(df_inventory)
        low_stock_df = df_inventory[df_inventory["目前庫存"] <= df_inventory["安全庫存"]]
        low_stock_count = len(low_stock_df)

        col1, col2 = st.columns(2)
        col1.metric("藥品品項總數", f"{total_items} 種")
        col2.metric("低於安全庫存警戒品項", f"{low_stock_count} 種", delta_color="inverse")

        # 低庫存警告
        if low_stock_count > 0:
            st.error(f"⚠️ 共有 {low_stock_count} 項藥品庫存偏低，請儘速補充進貨！")
            st.dataframe(
                low_stock_df[["藥品代碼", "藥品名稱", "目前庫存", "安全庫存", "單位", "分類"]],
                use_container_width=True
            )

        st.markdown("### 📋 完整庫存清單")
        
        # 搜尋篩選
        search_kw = st.text_input("🔍 搜尋藥品名稱或分類", placeholder="輸入關鍵字...")
        display_df = df_inventory.copy()
        
        if search_kw:
            display_df = display_df[
                display_df["藥品名稱"].str.contains(search_kw, na=False) |
                display_df["分類"].str.contains(search_kw, na=False)
            ]

        # 高亮顯示低庫存列
        def highlight_low_stock(row):
            if row["目前庫存"] <= row["安全庫存"]:
                return ['background-color: #ffcccc'] * len(row)
            return [''] * len(row)

        st.dataframe(
            display_df.style.apply(highlight_low_stock, axis=1),
            use_container_width=True
        )

# ==========================================
# 頁面 5：用藥月報與學期統計表
# ==========================================
elif page == "📊 用藥月報與學期統計表":
    st.subheader("📊 用藥統計分析")

    if df_logs.empty:
        st.info("尚無異動紀錄資料可供統計分析。")
    else:
        # 僅過濾領用紀錄
        usage_df = df_logs[df_logs["類型"] == "領用"].copy()
        
        if usage_df.empty:
            st.info("目前尚無藥品領用紀錄。")
        else:
            usage_df["數量"] = usage_df["變動數量"].abs()
            
            st.markdown("#### 🏆 熱門領用藥品排行榜")
            summary_drug = usage_df.groupby("藥品名稱")["數量"].sum().reset_index()
            summary_drug = summary_drug.sort_values(by="數量", ascending=False)

            st.bar_chart(data=summary_drug, x="藥品名稱", y="數量", use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("##### 藥品領用總數統計")
                st.dataframe(summary_drug, use_container_width=True)
            
            with col2:
                st.markdown("##### 領用紀錄明細表")
                st.dataframe(usage_df[["時間", "對象/系所", "藥品名稱", "數量", "經手人"]], use_container_width=True)
