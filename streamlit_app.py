@st.cache_resource
def init_gspread():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # 彈性讀取各種可能設定名稱，避免 KeyError
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
    elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        creds_dict = dict(st.secrets["connections"]["gsheets"])
    elif "type" in st.secrets and st.secrets["type"] == "service_account":
        creds_dict = dict(st.secrets)
    else:
        st.error("❌ 尚未在 Streamlit Secrets 中設定 Google Service Account 金鑰！請至 Streamlit 儀表板設定。")
        st.stop()

    # 處理 private_key 換行符號問題
    if "private_key" in creds_dict:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    credentials = Credentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )
    return gspread.authorize(credentials)
