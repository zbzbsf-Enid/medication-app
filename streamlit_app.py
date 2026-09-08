import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

# 1. 頁面基本配置
st.set_page_config(
    page_title="衛保組管理系統",
    page_icon="💊",
    layout="wide"
)

# 輔助函數：清洗與轉化庫存欄位型態
def clean_and_parse_inventory(df):
    if df is None or df.empty:
        return pd.DataFrame()
    
    df_clean = df.copy()
    df_clean.columns = df_clean.columns.astype(str).str.strip()
    
    # 搜尋期初庫存欄位名稱
    stock_col = None
    for possible_name in ['115年8月剩餘量', '期初庫存', '目前庫存', '庫存', '剩餘量']:
        if possible_name in df_clean.columns:
            stock_col = possible_name
            break
            
    if stock_col:
        df_clean[stock_col] = pd.to_numeric(df_clean[stock_col], errors='coerce').fillna(0)
    
    return df_clean

def get_init_stock_col(df):
    for possible_name in ['115年8月剩餘量', '期初庫存', '目前庫存', '庫存', '剩餘量']:
        if possible_name in df.columns:
            return possible_name
    return '115年8月剩餘量'

# 2. 初始化 Session State 資料庫
if "inventory" not in st.session_state:
    raw_df = pd.DataFrame([
        {"藥品名稱": "Actein 600mg (愛克痰發泡錠)", "115年8月剩餘量": 168, "有效期限": "2028-04-30"},
        {"藥品名稱": "Actein 600mg (愛克痰發泡錠)", "115年8月剩餘量": 461, "有效期限": "2028-05-31"},
        {"藥品名稱": "Amoxicillin 500mg (安莫西林)", "115年8月剩餘量": 400, "有效期限": "2028-02-28"},
        {"藥品名稱": "Fexofenadine 60mg (飛敏耐膜衣錠)", "115年8月剩餘量": 0, "有效期限": "2027-11-25"},
        {"藥品名稱": "Biofermin (表飛鳴)", "115年8月剩餘量": 604, "有效期限": "2028-05-31"}
    ])
    st.session_state.inventory = clean_and_parse_inventory(raw_df)

if "logs" not in st.session_state:
    st.session_state.logs = pd.DataFrame(columns=[
        "領用日期", "登記時間", "藥品名稱", "領用數量", "用途分類", "備註"
    ])

# 3. 側邊欄控制與資料同步
st.sidebar.title("🏥 衛保組管理系統")
page = st.sidebar.radio("📍 請選擇功能頁面", ["💊 藥品領用與登記", "📦 庫存盤點與校正", "☁️ 月報表加減統計與匯出"])

st.sidebar.markdown("---")
st.sidebar.subheader("🔗 串接 Google 雲端試算表")
gsheet_url = st.sidebar.text_input(
    "貼上 Google 試算表連結：",
    value="https://docs.google.com/spreadsheets/d/1gv_1Fz0iR9kUFVyj9P_IN0dJvz-drnv5v-wQ_UY8qvN0/edit?usp=sharing"
)

if st.sidebar.button("🔄 同步雲端試算表資料"):
    try:
        if "/edit" in gsheet_url:
            csv_url = gsheet_url.split("/edit")[0] + "/export?format=csv"
        else:
            csv_url = gsheet_url
        
        df_cloud = pd.read_csv(csv_url)
        st.session_state.inventory = clean_and_parse_inventory(df_cloud)
        st.sidebar.success("✅ 已成功同步雲端資料庫！")
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"❌ 讀取失敗，請確認共享權限：{e}")


# 4. 先進先出 (FIFO) 自動扣庫存函數
def fifo_deduct(df_inv, drug_eng_name, qty_needed):
    stock_col = get_init_stock_col(df_inv)
    matches = df_inv[df_inv['藥品名稱'].str.contains(drug_eng_name, regex=False, na=False)].copy()
    if matches.empty:
        return df_inv, False, f"找不到藥品：{drug_eng_name}"
    
    total_stock = matches[stock_col].sum()
    if total_stock < qty_needed:
        return df_inv, False, f"{drug_eng_name} 總庫存不足 (現有: {int(total_stock)}, 需要: {qty_needed})"
    
    if '有效期限' in matches.columns:
        matches = matches.sort_values(by='有效期限', na_position='last')
        
    rem = qty_needed
    for idx in matches.index:
        current_stock = df_inv.at[idx, stock_col]
        if current_stock >= rem:
            df_inv.at[idx, stock_col] = current_stock - rem
            rem = 0
            break
        else:
            rem -= current_stock
            df_inv.at[idx, stock_col] = 0
            
    return df_inv, True, "扣減成功"


# 5. 月報表動態加減統計計算引擎
def build_monthly_report_analysis(inv_df, logs_df):
    report_df = inv_df.copy()
    init_col = get_init_stock_col(report_df)
    
    # 確保統計欄位存在且為數字型態
    stat_cols = ['購入量', '過期報銷', '公藥使用', '當月使用總量', '115年9月期末剩餘量', '實體盤點數量', '9月消耗量']
    for c in stat_cols:
        if c not in report_df.columns:
            report_df[c] = 0
        else:
            report_df[c] = pd.to_numeric(report_df[c], errors='coerce').fillna(0)
            
    date_cols = [c for c in report_df.columns if c.startswith('9/') or c.startswith('10/') or c.startswith('11/')]
    for c in date_cols:
        report_df[c] = 0

    if not logs_df.empty:
        logs_temp = logs_df.copy()
        logs_temp['領用日期'] = pd.to_datetime(logs_temp['領用日期'], errors='coerce')
        
        for idx, row in report_df.iterrows():
            drug_name = str(row.get('藥品名稱', ''))
            matched_logs = logs_temp[logs_temp['藥品名稱'].apply(lambda x: str(x) in drug_name or drug_name in str(x))]
            
            if not matched_logs.empty:
                # 分類歸總：公藥使用與過期報銷
                pub_used = matched_logs[matched_logs['用途分類'] == '🏛️ 公藥使用']['領用數量'].sum()
                exp_used = matched_logs[matched_logs['用途分類'] == '🗑️ 過期報銷']['領用數量'].sum()
                
                report_df.at[idx, '公藥使用'] = pub_used
                report_df.at[idx, '過期報銷'] = exp_used
                
                # 每日一般領用填入日期欄
                general_logs = matched_logs[~matched_logs['用途分類'].isin(['🏛️ 公藥使用', '🗑️ 過期報銷'])]
                for _, log_row in general_logs.iterrows():
                    if pd.notna(log_row['領用日期']):
                        d_str = f"{log_row['領用日期'].month}/{log_row['領用日期'].day}"
                        if d_str in report_df.columns:
                            report_df.at[idx, d_str] += log_row['領用數量']

    # 🧮 核心加減統計公式計算
    date_sum = report_df[date_cols].sum(axis=1) if date_cols else 0
    report_df['當月使用總量'] = date_sum + report_df['公藥使用'] + report_df['過期報銷']
    report_df['115年9月期末剩餘量'] = report_df[init_col] + report_df['購入量'] - report_df['當月使用總量']
    report_df['實體盤點數量'] = report_df['115年9月期末剩餘量']
    report_df['9月消耗量'] = report_df['當月使用總量']

    return report_df


# --- 頁面 1: 藥品領用與登記 ---
if page == "💊 藥品領用與登記":
    col_left, col_right = st.columns([1.3, 1])
    
    st.session_state.inventory = clean_and_parse_inventory(st.session_state.inventory)
    stock_col = get_init_stock_col(st.session_state.inventory)
    
    with col_left:
        st.subheader("💊 藥品領用與登記")
        st.caption("點選下方搜尋欄選擇藥品，設定數量與用途（如公藥使用），發藥紀錄將自動計入月報表統計。")
        
        c_date, c_cat = st.columns([1, 1])
        with c_date:
            issue_date = st.date_input("📅 領用日期", value=date.today())
        with c_cat:
            # 🔑 增加獨立且顯眼的「公藥使用」選項
            category = st.selectbox(
                "🏷️ 用途分類 (請特別留意選擇)", 
                ["📋 一般消耗/學生領用", "🏛️ 公藥使用", "🗑️ 過期報銷", "其他"]
            )
            
        available_inventory = st.session_state.inventory[st.session_state.inventory[stock_col] > 0]
        
        drug_options = []
        for _, row in available_inventory.iterrows():
            d_name = str(row.get('藥品名稱', '')).strip()
            exp_date = str(row.get('有效期限', '')).strip() if pd.notna(row.get('有效期限')) else ""
            stk_val = int(row.get(stock_col, 0))
            
            if exp_date and exp_date.lower() != 'nan':
                opt_str = f"{d_name} | 效期:{exp_date} (庫存:{stk_val})"
            else:
                opt_str = f"{d_name} (庫存:{stk_val})"
            drug_options.append(opt_str)
        
        selected_options = st.multiselect(
            "選擇本次領取的所有藥品 (可同時選擇多項)",
            options=drug_options,
            placeholder="請點擊或輸入藥名進行搜尋..."
        )
        
        if not selected_options:
            st.info("💡 請先在上方的選單中點選或搜尋要領取的藥品。")
        else:
            st.markdown("---")
            st.markdown("##### 🔢 設定發放數量與細項")
            
            quantities = {}
            for opt in selected_options:
                drug_eng = opt.split(" | ")[0].split(" (")[0]
                total_k = available_inventory[
                    available_inventory['藥品名稱'].str.contains(drug_eng, regex=False, na=False)
                ][stock_col].sum()
                
                q = st.number_input(
                    f"數量 - {opt.split(' | ')[0]} [可用庫存: {int(total_k)}]", 
                    min_value=1, 
                    value=1, 
                    step=1, 
                    key=f"qty_{opt}"
                )
                quantities[opt] = q
                
            note = st.text_input("📝 備註說明 (選填)", placeholder="例如：衛保組活動備藥、公藥領用、研討會備用...")
            
            if st.button("確認登記並扣減庫存", type="primary", use_container_width=True):
                success_all = True
                messages = []
                
                for opt, qty in quantities.items():
                    drug_eng = opt.split(" | ")[0].split(" (")[0]
                    
                    df_inv, ok, msg = fifo_deduct(st.session_state.inventory, drug_eng, qty)
                    if ok:
                        st.session_state.inventory = df_inv
                        new_log = pd.DataFrame([{
                            "領用日期": issue_date.strftime("%Y-%m-%d"),
                            "登記時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "藥品名稱": drug_eng,
                            "領用數量": qty,
                            "用途分類": category,
                            "備註": note
                        }])
                        st.session_state.logs = pd.concat([new_log, st.session_state.logs], ignore_index=True)
                    else:
                        success_all = False
                        messages.append(msg)
                
                if success_all:
                    st.success("✅ 藥品領用登記成功！庫存與紀錄已即時更新。")
                    st.rerun()
                else:
                    st.error("⚠️ 登記過程發生錯誤：" + "；".join(messages))

    with col_right:
        st.subheader("📋 當前藥品庫存總覽")
        st.dataframe(
            st.session_state.inventory,
            use_container_width=True,
            hide_index=True,
            height=450
        )

    st.markdown("---")
    st.subheader("📜 歷史領用與發藥紀錄")
    if not st.session_state.logs.empty:
        st.dataframe(st.session_state.logs, use_container_width=True, hide_index=True)
    else:
        st.info("目前尚無任何領用紀錄。")


# --- 頁面 2: 庫存盤點與校正 ---
elif page == "📦 庫存盤點與校正":
    st.subheader("📦 庫存盤點與資料手動校正")
    st.caption("您可以直接在下方表格中修改庫存數量。")
    
    edited_df = st.data_editor(
        st.session_state.inventory,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )
    
    if st.button("💾 儲存盤點校正結果", type="primary"):
        st.session_state.inventory = clean_and_parse_inventory(edited_df)
        st.success("✅ 庫存資料已成功更新！")
        st.rerun()


# --- 頁面 3: 月報表加減統計與匯出 ---
elif page == "☁️ 月報表加減統計與匯出":
    st.subheader("📊 國立臺北大學衛保組 - 月報表動態加減統計與匯出")
    st.caption("系統會依據發藥紀錄，自動進行『每日消耗量 + 公藥使用 + 過期報銷 = 當月使用總量』與期末結算加減分析。")
    
    # 進行動態加減統計分析計算
    report_analysis = build_monthly_report_analysis(st.session_state.inventory, st.session_state.logs)
    
    tab_summary, tab_detail = st.tabs(["📊 全學期月報表加減統計分析", "📜 發藥領用歷史明細"])
    
    with tab_summary:
        st.dataframe(report_analysis, use_container_width=True, hide_index=True)
        
    with tab_detail:
        st.dataframe(st.session_state.logs, use_container_width=True, hide_index=True)
        
    st.markdown("---")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        report_analysis.to_excel(writer, sheet_name="115學年度月報表統計", index=False)
        st.session_state.logs.to_excel(writer, sheet_name="發藥領用紀錄", index=False)
        
    st.download_button(
        label="📥 下載計算完成的完整 Excel 月報表",
        data=buffer.getvalue(),
        file_name=f"國立臺北大學衛保組_115學年度上學期藥品使用月報表_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
