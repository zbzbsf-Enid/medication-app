import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# 1. 頁面基本設定與高對比主題 CSS
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title='國立臺北大學衛保組 - 藥品管理與月報系統',
    page_icon='💊',
    layout='wide',
)

st.markdown(
    """
    <style>
    :root { color-scheme: light !important; }
    html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif !important;
        font-size: 16px !important;
        background-color: #f8fafc !important;
        color: #0f172a !important;
    }
    p, span, label, h1, h2, h3, h4, .stMarkdown, div[data-testid="stMarkdownContainer"] * {
        color: #0f172a !important; opacity: 1 !important;
    }
    div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="select"] > div,
    .stTextInput input, .stNumberInput input, .stDateInput input {
        background-color: #ffffff !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
    }
    div[data-baseweb="tag"] {
        background-color: #e0f2fe !important;
        border: 1px solid #bae6fd !important;
        border-radius: 6px !important;
    }
    div[data-baseweb="tag"] span, div[data-baseweb="tag"] * {
        color: #0369a1 !important;
        -webkit-text-fill-color: #0369a1 !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="popover"], div[data-baseweb="popover"] > div,
    div[data-baseweb="calendar"], div[data-baseweb="menu"], ul[role="listbox"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1) !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #f1f5f9 !important;
        border-right: 1px solid #e2e8f0 !important;
    }
    .stButton > button, div[data-testid="stForm"] button {
        background-color: #0d9488 !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: 600 !important;
        padding: 8px 20px !important;
    }
    .stButton > button:hover, div[data-testid="stForm"] button:hover {
        background-color: #0f766e !important;
    }
    .confirm-card {
        background-color: #ffffff;
        border: 2px solid #0d9488;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.1);
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title('💊 國立臺北大學衛保組 藥品管理系統')

# Session State 初始化
if 'checkout_stage' not in st.session_state:
    st.session_state.checkout_stage = 'input'
if 'inbound_stage' not in st.session_state:
    st.session_state.inbound_stage = 'input'

# -----------------------------------------------------------------------------
# 2. 資料庫連線與資料讀取輔助函式
# -----------------------------------------------------------------------------
creds_dict = {}

if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
    conn_secrets = dict(st.secrets["connections"]["gsheets"])
    if "service_account" in conn_secrets:
        creds_dict = dict(conn_secrets["service_account"])
    else:
        creds_dict = conn_secrets
elif "gcp_service_account" in st.secrets:
    creds_dict = dict(st.secrets["gcp_service_account"])

# 自動修復私鑰換行問題
if "private_key" in creds_dict and isinstance(creds_dict["private_key"], str):
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1fqR5nvOGTOnKljryhMwfbAUaVZo5L11Jtsm823Hf8hU/edit"

try:
    # 💡 關鍵：將憑證透過 service_account= 傳入，避免 direct kwargs 參數撞名
    conn = st.connection(
        'medication_app',
        type=GSheetsConnection,
        service_account=creds_dict
    )
except Exception as e:
    st.error(f'❌ 無法建立 Google 連線：{e}')
    st.stop()
# 效期自動預警與警告通知
def check_expiration_warnings(df):
    if df is None or df.empty:
        return

    exp_col = None
    for col in df.columns:
        if '期限' in str(col) or '效期' in str(col):
            exp_col = col
            break

    if not exp_col:
        return

    today = datetime.now().date()
    expiring_soon = []
    already_expired = []

    for _, row in df.iterrows():
        exp_str = row.get(exp_col)
        if (
            pd.notna(exp_str)
            and str(exp_str).strip() != ''
            and str(exp_str).lower() != 'nan'
            and str(exp_str).lower() != 'none'
        ):
            try:
                exp_date = pd.to_datetime(exp_str).date()
                days_left = (exp_date - today).days

                eng_name = str(row.get('藥品名稱(英文)', '')).strip()
                cht_name = str(row.get('中文名稱', '')).strip()
                med_name = f'{eng_name} ({cht_name})' if cht_name else eng_name
                batch = str(row.get('批號', '無'))
                stock = int(row.get('目前庫存', 0))

                if days_left < 0:
                    already_expired.append({
                        'name': med_name,
                        'batch': batch,
                        'expiry': exp_date,
                        'stock': stock,
                        'days': abs(days_left),
                    })
                elif 0 <= days_left <= 30:
                    expiring_soon.append({
                        'name': med_name,
                        'batch': batch,
                        'expiry': exp_date,
                        'stock': stock,
                        'days': days_left,
                    })
            except Exception:
                continue

    if already_expired:
        st.error(
            '🚨 **【嚴重警告】系統檢測到有藥品已過期！提醒人員儘速處置與辦理報銷！**'
        )
        with st.expander('🚨 查看已過期藥品明細 (請儘速處置與報銷)', expanded=True):
            for item in already_expired:
                st.markdown(
                    f"❌ **{item['name']}** ｜ 批號：`{item['batch']}` ｜ 有效期限：`{item['expiry']}` "
                    f"（已過期 **{item['days']}** 天） ｜ 當前庫存：`{item['stock']}`"
                )

    if expiring_soon:
        with st.expander(
            '⚠️ 【到期預警】以下藥品將於 1 個月內到期！請留意使用狀況與準備報銷',
            expanded=True,
        ):
            for item in expiring_soon:
                st.warning(
                    f"⚠️ **{item['name']}** ｜ 批號：`{item['batch']}` ｜ 有效期限：`{item['expiry']}` "
                    f"（剩餘 **{item['days']}** 天） ｜ 當前庫存：`{item['stock']}`"
                )


def append_to_log_sheet(connection, possible_sheet_names, new_logs_df):
    for sheet_name in possible_sheet_names:
        try:
            try:
                df_existing = connection.read(
                    spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl='5m'
                )
                df_updated = pd.concat([df_existing, new_logs_df], ignore_index=True)
            except Exception:
                df_updated = new_logs_df

            connection.update(
                spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, data=df_updated
            )
            return True, sheet_name
        except Exception:
            continue
    return False, possible_sheet_names[0]


def get_log_sheet_data(connection, possible_sheet_names):
    for sheet_name in possible_sheet_names:
        try:
            df = connection.read(
                spreadsheet=SPREADSHEET_URL, worksheet=sheet_name, ttl='5m'
            )
            return df, sheet_name
        except Exception:
            continue
    return None, possible_sheet_names[0]


# -----------------------------------------------------------------------------
# 3. 側邊欄與頁面選擇
# -----------------------------------------------------------------------------
st.sidebar.title('📌 功能選單')
if st.sidebar.button('🔄 手動刷新雲端資料'):
    st.cache_data.clear()
    st.rerun()

menu = st.sidebar.radio(
    '請選擇功能頁面',
    [
        '💊 多項藥品領用登記',
        '📥 藥品進貨/建檔登記',
        '🛠️ 紀錄修改與庫存微調',
        '📦 當前庫存總覽',
        '📊 用藥月報與學期統計表',
    ],
)

df_inventory = load_data()
check_expiration_warnings(df_inventory)

# -----------------------------------------------------------------------------
# 頁面 1：多項藥品領用登記
# -----------------------------------------------------------------------------
if menu == '💊 多項藥品領用登記':
    st.header('📋 批量藥品領用登記')
    df_inventory['display_name'] = (
        df_inventory['藥品名稱(英文)'].fillna('')
        + ' ('
        + df_inventory['中文名稱'].fillna('')
        + ') - 批號:'
        + df_inventory['批號'].astype(str)
    )
    options = df_inventory['display_name'].tolist()

    if st.session_state.checkout_stage == 'input':
        selected_items = st.multiselect(
            '請選擇或搜尋欲領取的藥品（可多選）：',
            options=options,
            placeholder='點擊或輸入藥品名稱關鍵字...',
        )

        if selected_items:
            st.markdown('---')
            st.subheader('✏️ 請輸入領用資訊與數量')

            col_d1, col_d2 = st.columns([1, 2])
            with col_d1:
                record_date = st.date_input(
                    '📅 實際領用/補登日期：', value=datetime.now().date()
                )
            with col_d2:
                remarks = st.text_input(
                    '領用備註/用途：', placeholder='例如：衛保組公用 / 門診備用'
                )

            st.markdown("<hr style='margin: 12px 0;'>", unsafe_allow_html=True)

            quantities = {}
            for item in selected_items:
                row_data = df_inventory[df_inventory['display_name'] == item].iloc[0]
                curr_stock = int(row_data['目前庫存'])
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.markdown(f'**{item}**')
                    st.caption(f'目前剩餘庫存：`{curr_stock}`')
                with col2:
                    qty = st.number_input(
                        '領用數量',
                        min_value=1,
                        max_value=max(curr_stock, 1),
                        value=1,
                        step=1,
                        key=f'input_{item}',
                    )
                    quantities[item] = qty
                st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

            if st.button('🔍 預覽並檢查領用明細', type='primary'):
                st.session_state.checkout_data = {
                    'record_date': record_date,
                    'remarks': remarks,
                    'quantities': quantities,
                }
                st.session_state.checkout_stage = 'confirm'
                st.rerun()

    elif st.session_state.checkout_stage == 'confirm':
        c_data = st.session_state.checkout_data
        st.markdown("<div class='confirm-card'>", unsafe_allow_html=True)
        st.warning('⚠️ **請再次核對領用內容，確認無誤後點選下方按鈕提交：**')
        st.write(f"📅 **領用日期：** `{c_data['record_date']}`")
        st.write(
            f"📝 **領用備註：** `{c_data['remarks'] if c_data['remarks'] else '無'}`"
        )

        preview_rows = []
        for item, qty in c_data['quantities'].items():
            row_data = df_inventory[df_inventory['display_name'] == item].iloc[0]
            curr_stock = int(row_data['目前庫存'])
            after_stock = max(0, curr_stock - qty)
            preview_rows.append({
                '藥品品名': item,
                '目前庫存': curr_stock,
                '預計領用數量': qty,
                '領用後剩餘庫存': after_stock,
            })
        st.table(pd.DataFrame(preview_rows))
        st.markdown('</div>', unsafe_allow_html=True)

        col_c1, col_c2 = st.columns([1, 1])
        with col_c1:
            if st.button('✅ 確認無誤，寫入雲端並更新庫存', type='primary'):
                log_time_str = f"{c_data['record_date'].strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M:%S')}"
                new_logs = []
                for item, qty in c_data['quantities'].items():
                    idx = df_inventory[df_inventory['display_name'] == item].index[0]
                    old_stock = int(df_inventory.loc[idx, '目前庫存'])
                    new_stock = max(0, old_stock - qty)
                    df_inventory.loc[idx, '目前庫存'] = new_stock
                    new_logs.append({
                        '領用時間': str(log_time_str),
                        '藥品名稱': str(df_inventory.loc[idx, '藥品名稱(英文)']),
                        '中文名稱': str(df_inventory.loc[idx, '中文名稱']),
                        '領用數量': int(qty),
                        '剩餘庫存': int(new_stock),
                        '備註': str(c_data['remarks']),
                    })

                try:
                    df_save = df_inventory.drop(
                        columns=['display_name'], errors='ignore'
                    )
                    conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save)

                    df_logs_new = pd.DataFrame(new_logs)
                    success, target_sheet = append_to_log_sheet(
                        conn, ['領用記錄', '領用紀錄'], df_logs_new
                    )
                    st.cache_data.clear()
                    st.session_state.checkout_stage = 'input'
                    if success:
                        st.success(
                            f"🎉 領用登記成功！紀錄已同步寫入『{target_sheet}』分頁。"
                        )
                        st.rerun()
                    else:
                        st.warning('⚠️ 庫存已扣減，但寫入歷史紀錄失敗，請檢查權限。')
                except Exception as e_inv:
                    st.error(f'❌ 庫存更新失敗：{e_inv}')

        with col_c2:
            if st.button('✏️ 返回修改數量/品項'):
                st.session_state.checkout_stage = 'input'
                st.rerun()

# -----------------------------------------------------------------------------
# 頁面 2：藥品進貨/建檔登記
# -----------------------------------------------------------------------------
elif menu == '📥 藥品進貨/建檔登記':
    st.header('📥 藥品進貨與建檔管理')

    tab_inbound, tab_create = st.tabs(['📦 現有藥品補貨/進貨', '➕ 建立全新藥品品項'])

    with tab_inbound:
        df_inventory['display_name'] = (
            df_inventory['藥品名稱(英文)'].fillna('')
            + ' ('
            + df_inventory['中文名稱'].fillna('')
            + ')'
        )
        med_list = df_inventory['display_name'].tolist()

        if st.session_state.inbound_stage == 'input':
            selected_med = st.selectbox(
                '請選擇補貨藥品：', options=med_list, key='select_existing_med'
            )

            if selected_med:
                row_info = df_inventory[
                    df_inventory['display_name'] == selected_med
                ].iloc[0]
                st.info(
                    f"📌 當前庫存：`{row_info['目前庫存']}` | 目前批號：`{row_info.get('批號', '無')}` |"
                    f" 目前有效期限：`{row_info.get('有效期限', '無')}`"
                )

                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    inbound_date = st.date_input(
                        '📅 進貨日期：',
                        value=datetime.now().date(),
                        key='inbound_date',
                    )
                    purchase_qty = st.number_input(
                        '📦 購入數量：', min_value=1, value=100, step=1, key='purchase_qty'
                    )
                with col_p2:
                    new_batch = st.text_input(
                        '🏷️ 新藥品批號：',
                        value=str(row_info.get('批號', ''))
                        if pd.notna(row_info.get('批號'))
                        else '',
                        key='new_batch',
                    )
                    new_expiry = st.date_input(
                        '⏳ 新有效期限：',
                        value=datetime.now().date(),
                        key='new_expiry',
                    )

                vendor_remark = st.text_input(
                    '🏢 廠商/採購備註：',
                    placeholder='例如：衛福部撥發 / 某某藥局採購',
                    key='vendor_remark',
                )

                if st.button('🔍 預覽進貨明細', type='primary'):
                    st.session_state.inbound_data = {
                        'selected_med': selected_med,
                        'inbound_date': inbound_date,
                        'purchase_qty': purchase_qty,
                        'new_batch': new_batch,
                        'new_expiry': new_expiry,
                        'vendor_remark': vendor_remark,
                        'old_qty': int(row_info['目前庫存']),
                    }
                    st.session_state.inbound_stage = 'confirm'
                    st.rerun()

        elif st.session_state.inbound_stage == 'confirm':
            i_data = st.session_state.inbound_data
            st.markdown("<div class='confirm-card'>", unsafe_allow_html=True)
            st.warning('⚠️ **請再次核對進貨資訊：**')
            st.write(f"💊 **進貨藥品：** `{i_data['selected_med']}`")
            st.write(f"📅 **進貨日期：** `{i_data['inbound_date']}`")
            st.write(
                f"📦 **進貨數量：** `{i_data['purchase_qty']}` （原有庫存：{i_data['old_qty']}"
                f" ➔ **進貨後總庫存：{i_data['old_qty'] + i_data['purchase_qty']}**）"
            )
            st.write(
                f"🏷️ **新批號：** `{i_data['new_batch']}` | ⏳ **新有效期限：**"
                f" `{i_data['new_expiry']}`"
            )
            st.write(
                f"📝 **採購備註：** `{i_data['vendor_remark'] if i_data['vendor_remark'] else '無'}`"
            )
            st.markdown('</div>', unsafe_allow_html=True)

            col_i1, col_i2 = st.columns([1, 1])
            with col_i1:
                if st.button('✅ 確認進貨並同步更新雲端', type='primary'):
                    idx = df_inventory[
                        df_inventory['display_name'] == i_data['selected_med']
                    ].index[0]
                    new_qty = i_data['old_qty'] + int(i_data['purchase_qty'])
                    expiry_str = i_data['new_expiry'].strftime('%Y-%m-%d')
                    inbound_time_str = f"{i_data['inbound_date'].strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M:%S')}"

                    df_inventory.loc[idx, '目前庫存'] = new_qty
                    if '批號' in df_inventory.columns:
                        df_inventory.loc[idx, '批號'] = i_data['new_batch']

                    for col in df_inventory.columns:
                        if '期限' in str(col) or '效期' in str(col):
                            df_inventory.loc[idx, col] = expiry_str
                            break

                    purchase_log = {
                        '進貨時間': str(inbound_time_str),
                        '藥品名稱': str(df_inventory.loc[idx, '藥品名稱(英文)']),
                        '中文名稱': str(df_inventory.loc[idx, '中文名稱']),
                        '購入數量': int(i_data['purchase_qty']),
                        '新批號': str(i_data['new_batch']),
                        '有效期限': str(expiry_str),
                        '更新後總庫存': int(new_qty),
                        '備註': str(i_data['vendor_remark']),
                    }

                    try:
                        df_save = df_inventory.drop(
                            columns=['display_name'], errors='ignore'
                        )
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save)

                        df_inbound_new = pd.DataFrame([purchase_log])
                        append_to_log_sheet(
                            conn, ['進貨紀錄', '進貨記錄'], df_inbound_new
                        )

                        st.cache_data.clear()
                        st.session_state.inbound_stage = 'input'
                        st.success(
                            f"🎉 進貨完成！`{i_data['selected_med']}` 庫存已由 {i_data['old_qty']} 增加至"
                            f" {new_qty}。"
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f'❌ 庫存更新失敗：{e}')

            with col_i2:
                if st.button('✏️ 返回修改進貨內容'):
                    st.session_state.inbound_stage = 'input'
                    st.rerun()

    with tab_create:
        st.subheader('➕ 新增全新藥品至雲端庫存清單')
        st.caption('用於建立試算表中尚未存在的新藥品資料。')

        col_n1, col_n2 = st.columns(2)
        with col_n1:
            new_eng_name = st.text_input(
                '💊 藥品英文名稱/商品名（必填）：',
                placeholder='例如：Amoxicillin 500mg',
            )
            new_cht_name = st.text_input(
                '💊 藥品中文名稱：', placeholder='例如：安莫西林膠囊'
            )
            new_batch_code = st.text_input('🏷️ 批號：', value='DEFAULT')
        with col_n2:
            new_init_stock = st.number_input(
                '📦 初期庫存數量：', min_value=0, value=0, step=1
            )
            new_exp_date = st.date_input(
                '⏳ 有效期限：', value=datetime.now().date()
            )
            new_usage = st.text_input(
                '📝 作用/適應症/用途說明：', placeholder='例如：抗生素/感冒發炎用藥'
            )

        if st.button('✨ 建立新藥品並寫入雲端', type='primary'):
            if not new_eng_name.strip():
                st.error('❌ 藥品英文名稱為必填欄位！')
            else:
                existing_match = df_inventory[
                    df_inventory['藥品名稱(英文)']
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    == new_eng_name.strip().lower()
                ]
                if not existing_match.empty:
                    st.warning(
                        f"⚠️ 雲端庫存中已存在名稱類似的藥品『{new_eng_name}』，建議直接使用「現有藥品補貨」功能。"
                    )
                else:
                    df_save = df_inventory.drop(
                        columns=['display_name'], errors='ignore'
                    ).copy()

                    new_drug_row = {}
                    for col in df_save.columns:
                        col_s = str(col).strip()
                        if '英文' in col_s or '藥品名稱' in col_s:
                            new_drug_row[col] = new_eng_name.strip()
                        elif '中文' in col_s:
                            new_drug_row[col] = new_cht_name.strip()
                        elif '批號' in col_s:
                            new_drug_row[col] = new_batch_code.strip()
                        elif '庫存' in col_s:
                            new_drug_row[col] = int(new_init_stock)
                        elif '期限' in col_s or '效期' in col_s:
                            new_drug_row[col] = new_exp_date.strftime('%Y-%m-%d')
                        elif '用途' in col_s or '備註' in col_s:
                            new_drug_row[col] = new_usage.strip()
                        elif '狀態' in col_s or 'last' in col_s or 'updated' in col_s:
                            new_drug_row[col] = 'OK'
                        else:
                            new_drug_row[col] = ''

                    df_save = pd.concat(
                        [df_save, pd.DataFrame([new_drug_row])], ignore_index=True
                    )

                    try:
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save)
                        st.cache_data.clear()
                        st.success(
                            f"🎉 成功建立新藥品『{new_eng_name} ({new_cht_name})』！已正確對齊雲端欄位並同步更新。"
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f'❌ 寫入雲端失敗：{e}')

# -----------------------------------------------------------------------------
# 頁面 3：紀錄修改與庫存微調
# -----------------------------------------------------------------------------
elif menu == '🛠️ 紀錄修改與庫存微調':
    st.header('🛠️ 歷史紀錄更正與庫存校正')
    st.caption(
        '若發現先前輸入的數量有誤，可在本頁面進行「修改」、「紀錄撤銷」或「直接庫存校正」。'
    )

    tab1, tab2 = st.tabs(['✏️ 領用紀錄更正/撤銷', '🎯 直接盤點庫存校正'])

    with tab1:
        st.subheader('更正歷史領用紀錄')
        df_logs, sheet_used = get_log_sheet_data(conn, ['領用記錄', '領用紀錄'])

        if df_logs is not None and not df_logs.empty:
            st.write(
                f"📋 目前從雲端『{sheet_used}』讀取到的最近歷史紀錄（顯示前 20 筆）："
            )
            recent_logs = df_logs.tail(20).copy().iloc[::-1]
            st.dataframe(recent_logs, use_container_width=True)

            log_indices = recent_logs.index.tolist()
            log_options = [
                f"行號 {idx+2}: [{recent_logs.loc[idx, '領用時間']}]"
                f" {recent_logs.loc[idx, '藥品名稱']} - 原領用量:"
                f" {recent_logs.loc[idx, '領用數量']}"
                for idx in log_indices
            ]

            selected_log_str = st.selectbox(
                '請選擇欲修改或撤銷的該筆紀錄：', options=log_options
            )

            if selected_log_str:
                selected_idx = (
                    int(selected_log_str.split(':')[0].replace('行號 ', '')) - 2
                )
                target_log = df_logs.loc[selected_idx]
                med_name = target_log['藥品名稱']
                orig_qty = int(target_log['領用數量'])

                st.markdown('---')
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    st.markdown(f'**目前選擇項目：** `{med_name}`')
                    st.markdown(f'**原登記領用數量：** `{orig_qty}`')
                    new_log_qty = st.number_input(
                        '✏️ 修改為正確領用數量：',
                        min_value=0,
                        value=orig_qty,
                        step=1,
                    )
                with col_e2:
                    st.write('🔧 **處置動作：**')
                    btn_update_log = st.button(
                        '💾 更新此筆領用數量並調整庫存', type='primary'
                    )
                    btn_delete_log = st.button(
                        '🗑️ 徹底撤銷此筆紀錄（數量全數加回庫存）'
                    )

                if btn_update_log:
                    diff = new_log_qty - orig_qty
                    inv_match = df_inventory[df_inventory['藥品名稱(英文)'] == med_name]
                    if not inv_match.empty:
                        inv_idx = inv_match.index[0]
                        curr_stk = int(df_inventory.loc[inv_idx, '目前庫存'])
                        adjusted_stk = max(0, curr_stk - diff)

                        df_inventory.loc[inv_idx, '目前庫存'] = adjusted_stk
                        df_logs.loc[selected_idx, '領用數量'] = new_log_qty
                        df_logs.loc[selected_idx, '剩餘庫存'] = adjusted_stk
                        df_logs.loc[selected_idx, '備註'] = (
                            str(df_logs.loc[selected_idx, '備註']) + ' (已更正)'
                        )

                        df_save_inv = df_inventory.drop(
                            columns=['display_name'], errors='ignore'
                        )
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save_inv)
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_used, data=df_logs)

                        st.cache_data.clear()
                        st.success(
                            f'🎉 紀錄已更新！`{med_name}` 庫存已相應調整為 `{adjusted_stk}`。'
                        )
                        st.rerun()
                    else:
                        st.error('❌ 找不到對應的藥品庫存項目。')

                if btn_delete_log:
                    inv_match = df_inventory[df_inventory['藥品名稱(英文)'] == med_name]
                    if not inv_match.empty:
                        inv_idx = inv_match.index[0]
                        curr_stk = int(df_inventory.loc[inv_idx, '目前庫存'])
                        adjusted_stk = curr_stk + orig_qty

                        df_inventory.loc[inv_idx, '目前庫存'] = adjusted_stk
                        df_logs = df_logs.drop(index=selected_idx)

                        df_save_inv = df_inventory.drop(
                            columns=['display_name'], errors='ignore'
                        )
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save_inv)
                        conn.update(spreadsheet=SPREADSHEET_URL, worksheet=sheet_used, data=df_logs)

                        st.cache_data.clear()
                        st.success(
                            f'🎉 已成功撤銷該筆紀錄！領用的 `{orig_qty}`'
                            f' 單位已自動加回庫存，最新庫存為 `{adjusted_stk}`。'
                        )
                        st.rerun()
        else:
            st.info('目前尚未有任何領用紀錄。')

    with tab2:
        st.subheader('🎯 直接盤點與庫存微調')
        st.caption('用於盤點發現數量不符時，直接手動修正最終庫存數。')

        df_inventory['display_name'] = (
            df_inventory['藥品名稱(英文)'].fillna('')
            + ' ('
            + df_inventory['中文名稱'].fillna('')
            + ')'
        )
        med_list = df_inventory['display_name'].tolist()

        selected_cal_med = st.selectbox(
            '請選擇欲校正的藥品：', options=med_list, key='cal_select'
        )
        if selected_cal_med:
            cal_row = df_inventory[
                df_inventory['display_name'] == selected_cal_med
            ].iloc[0]
            c_idx = df_inventory[
                df_inventory['display_name'] == selected_cal_med
            ].index[0]

            st.write(f"目前雲端記錄庫存量：`{cal_row['目前庫存']}`")
            correct_stock = st.number_input(
                '🎯 輸入實體盤點後的真實庫存數量：',
                min_value=0,
                value=int(cal_row['目前庫存']),
                step=1,
            )

            if st.button('✅ 強制覆蓋寫入正確庫存'):
                df_inventory.loc[c_idx, '目前庫存'] = correct_stock
                df_save = df_inventory.drop(columns=['display_name'], errors='ignore')
                conn.update(spreadsheet=SPREADSHEET_URL, worksheet='庫存', data=df_save)

                st.cache_data.clear()
                st.success(f'🎉 `{selected_cal_med}` 庫存已修正為 `{correct_stock}`！')
                st.rerun()

# -----------------------------------------------------------------------------
# 頁面 4：當前庫存總覽
# -----------------------------------------------------------------------------
elif menu == '📦 當前庫存總覽':
    st.header('📦 當前藥品庫存總覽')
    display_df = df_inventory.drop(
        columns=['display_name'], errors='ignore'
    ).copy()
    st.dataframe(display_df, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# 頁面 5：用藥月報與學期統計表
# -----------------------------------------------------------------------------
elif menu == '📊 用藥月報與學期統計表':
    st.header('📊 國立臺北大學衛保組 藥品使用月報與全學期統計表')

    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        selected_year = st.selectbox(
            '學年度/年份', ['115學年度', '114學年度'], index=0
        )
    with col_sel2:
        selected_month = st.selectbox(
            '統計月份', ['9月', '10月', '11月', '12月', '1月'], index=0
        )

    month_num = int(selected_month.replace('月', ''))
    days_list = [
        1, 2, 3, 4, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 21, 22, 23, 24, 25, 28, 29, 30
    ]
    days = [f'{month_num}/{d}' for d in days_list]

    try:
        roc_num = int(selected_year.replace('學年度', '').strip())
        report_year = (roc_num + 1911) if month_num >= 8 else (roc_num + 1911 + 1)
    except Exception:
        report_year = datetime.now().year

    df_logs, _ = get_log_sheet_data(conn, ['領用記錄', '領用紀錄'])
    daily_usage_map = {}

    if df_logs is not None and not df_logs.empty:
        try:
            df_logs['dt'] = pd.to_datetime(df_logs['領用時間'], errors='coerce')
            df_logs_filtered = df_logs[
                df_logs['dt'].dt.month == month_num
            ].copy()
            df_logs_filtered['day_str'] = df_logs_filtered['dt'].apply(
                lambda x: f'{x.month}/{x.day}' if pd.notnull(x) else ''
            )

            grouped = (
                df_logs_filtered.groupby(['藥品名稱', 'day_str'])['領用數量']
                .sum()
                .reset_index()
            )
            for _, g_row in grouped.iterrows():
                med_name = str(g_row['藥品名稱']).strip()
                d_str = str(g_row['day_str']).strip()
                qty = int(g_row['領用數量'])
                daily_usage_map[(med_name, d_str)] = qty

            if '中文名稱' in df_logs_filtered.columns:
                grouped_cht = (
                    df_logs_filtered.groupby(['中文名稱', 'day_str'])['領用數量']
                    .sum()
                    .reset_index()
                )
                for _, g_row in grouped_cht.iterrows():
                    cht_name = str(g_row['中文名稱']).strip()
                    d_str = str(g_row['day_str']).strip()
                    qty = int(g_row['領用數量'])
                    if cht_name:
                        daily_usage_map[(cht_name, d_str)] = qty
        except Exception as e:
            st.warning(f'⚠️ 讀取領用紀錄計算每日用量時提醒：{e}')

    report_rows = []
    exp_col = None
    for col in df_inventory.columns:
        if '期限' in str(col) or '效期' in str(col):
            exp_col = col
            break

    for idx, row in df_inventory.iterrows():
        eng_name = str(row.get('藥品名稱(英文)', '')).strip()
        cht_name = str(row.get('中文名稱', '')).strip()
        combined_name = f'{eng_name} ({cht_name})' if cht_name else eng_name
        stock = int(row.get('目前庫存', 0))
        expiry = str(row.get(exp_col, '')).strip() if exp_col else ''

        day_quantities = []
        r_dict = {
            '藥品名稱\n(商品名/中文)': combined_name,
            '115年8月\n剩餘量': stock,
        }

        for d in days:
            qty_used = daily_usage_map.get((eng_name, d), 0)
            if qty_used == 0 and cht_name:
                qty_used = daily_usage_map.get((cht_name, d), 0)

            r_dict[d] = qty_used
            day_quantities.append(qty_used)

        monthly_used_sum = sum(day_quantities)

        expired_writeoff = 0
        if (
            pd.notna(expiry)
            and expiry != ''
            and expiry.lower() != 'nan'
            and expiry.lower() != 'none'
        ):
            try:
                exp_date = pd.to_datetime(expiry).date()
                if (exp_date.year == report_year and exp_date.month == month_num) or (
                    exp_date < datetime(report_year, month_num, 1).date()
                ):
                    expired_writeoff = max(0, stock - monthly_used_sum)
            except Exception:
                expired_writeoff = 0

        rem_stock = max(0, stock - monthly_used_sum - expired_writeoff)

        r_dict.update({
            '當月使用\n總量': monthly_used_sum,
            '當月剩餘量': rem_stock,
            '購入量': 0,
            '過期報銷': expired_writeoff,
            '公藥使用': 0,
            '實體盤點\n數量': rem_stock,
            '有效期限': expiry,
            '9月\n消耗量': monthly_used_sum if month_num == 9 else 0,
            '10月\n消耗量': monthly_used_sum if month_num == 10 else 0,
            '11月\n消耗量': monthly_used_sum if month_num == 11 else 0,
            '12月\n消耗量': monthly_used_sum if month_num == 12 else 0,
            '1月\n消耗量': monthly_used_sum if month_num == 1 else 0,
            '全學期\n使用總量': monthly_used_sum,
        })
        report_rows.append(r_dict)

    df_report = pd.DataFrame(report_rows)
    st.subheader(f'📄 {selected_year} 上學期用藥月報表 ({selected_month}) 預覽')
    st.dataframe(df_report, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------------------
    # openpyxl 匯出 Excel
    # -------------------------------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'115年{selected_month}用藥月報表'

    title_text = (
        f'國立臺北大學衛保組 {selected_year}上學期藥品使用月報與全學期統計表'
        f' ({selected_year}{selected_month}起)'
    )
    ws.append([title_text])
    ws.cell(row=1, column=1).font = Font(
        name='微軟正黑體', size=13, bold=True, color='1F4E78'
    )

    excel_headers = [
        '藥品名稱\n(商品名/中文)',
        '115年8月\n剩餘量',
        *days,
        '當月使用\n總量',
        '當月剩餘量',
        '購入量',
        '過期報銷',
        '公藥使用',
        '實體盤點\n數量',
        '有效期限',
        '9月\n消耗量',
        '10月\n消耗量',
        '11月\n消耗量',
        '12月\n消耗量',
        '1月\n消耗量',
        '全學期\n使用總量',
    ]
    ws.append(excel_headers)

    fill_navy = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_gray = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_blue = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
    fill_orange = PatternFill(start_color='C65911', end_color='C65911', fill_type='solid')

    font_white_bold = Font(name='微軟正黑體', size=9, bold=True, color='FFFFFF')
    font_gray_bold = Font(name='微軟正黑體', size=9, bold=True, color='333333')
    font_default = Font(name='微軟正黑體', size=9)

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9'),
    )

    for col_idx, _ in enumerate(excel_headers, start=1):
        cell = ws.cell(row=2, column=col_idx)
        cell.alignment = align_center
        cell.border = thin_border
        if col_idx in [1, 2]:
            cell.fill = fill_navy
            cell.font = font_white_bold
        elif 3 <= col_idx <= 2 + len(days):
            cell.fill = fill_gray
            cell.font = font_gray_bold
        elif col_idx in [3 + len(days), 4 + len(days)]:
            cell.fill = fill_blue
            cell.font = font_white_bold
        else:
            cell.fill = fill_orange
            cell.font = font_white_bold

    for r_idx, r_data in enumerate(report_rows, start=3):
        row_values = [
            r_data['藥品名稱\n(商品名/中文)'],
            r_data['115年8月\n剩餘量'],
            *[r_data[d] for d in days],
            r_data['當月使用\n總量'],
            r_data['當月剩餘量'],
            r_data['購入量'],
            r_data['過期報銷'],
            r_data['公藥使用'],
            r_data['實體盤點\n數量'],
            r_data['有效期限'],
            r_data['9月\n消耗量'],
            r_data['10月\n消耗量'],
            r_data['11月\n消耗量'],
            r_data['12月\n消耗量'],
            r_data['1月\n消耗量'],
            r_data['全學期\n使用總量'],
        ]
        ws.append(row_values)
        for c_idx, _ in enumerate(row_values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            cell.font = font_default
            cell.alignment = align_left if c_idx == 1 else align_center

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        label='📥 下載用藥月報與全學期統計 Excel 檔案',
        data=buffer,
        file_name=f'國立臺北大學衛保組_{selected_year}_{selected_month}_用藥月報表.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
