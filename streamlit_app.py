# 💡 智慧型對齊欄位新增邏輯（防止試算表欄位錯位）
        if st.button("✨ 建立新藥品並寫入雲端", type="primary"):
            if not new_eng_name.strip():
                st.error("❌ 藥品英文名稱為必填欄位！")
            else:
                # 檢查是否已存在相同英文名稱
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
                    # 準備複製舊表的欄位結構
                    df_save = df_inventory.drop(
                        columns=['display_name'], errors='ignore'
                    ).copy()

                    # 動態比對現有表格的欄位名稱
                    new_drug_row = {}
                    for col in df_save.columns:
                        col_clean = str(col).strip()

                        if col_clean in ['藥品名稱(英文)', '英文名稱']:
                            new_drug_row[col] = new_eng_name.strip()
                        elif col_clean in ['中文名稱']:
                            new_drug_row[col] = new_cht_name.strip()
                        elif col_clean in ['批號']:
                            new_drug_row[col] = new_batch_code.strip()
                        elif col_clean in ['目前庫存', '庫存']:
                            new_drug_row[col] = int(new_init_stock)
                        elif col_clean in ['有效期限', '效期']:
                            new_drug_row[col] = new_exp_date.strftime(
                                '%Y-%m-%d'
                            )
                        elif '用途' in col_clean or '備註' in col_clean:
                            new_drug_row[col] = new_usage.strip()
                        elif col_clean in ['狀態', 'last_updated']:
                            new_drug_row[col] = 'OK'
                        else:
                            new_drug_row[col] = ''

                    # 將正確對齊的新資料併入表格
                    df_save = pd.concat(
                        [df_save, pd.DataFrame([new_drug_row])],
                        ignore_index=True,
                    )

                    try:
                        conn.update(worksheet='庫存', data=df_save)
                        st.cache_data.clear()
                        st.success(
                            f'🎉 成功建立新藥品『{new_eng_name} ({new_cht_name})』！已正確對齊欄位並寫入 Google 試算表。'
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f'❌ 寫入雲端失敗：{e}')
