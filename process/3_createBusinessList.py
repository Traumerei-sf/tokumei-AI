import pandas as pd
import urllib.parse
import json
from typing import Dict, Tuple
from process.u_accessGemini import exe_gemini_withGoogleSearch_and_structure
import streamlit as st

def create_business_list(df_journal: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """営業先リストを作成する（スプレッドシートのB2セルを使用）"""
    target_accounts = ["売上高", "売掛金", "受取手形"]
    return _create_list_common(df_journal, row_idx=1, col_idx=1, target_accounts=target_accounts, fallback_name="営業先")

def create_supplier_list(df_journal: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """仕入先リストを作成する（スプレッドシートのB3セルを使用）"""
    target_accounts = ["外注費", "仕入高"]
    return _create_list_common(df_journal, row_idx=2, col_idx=1, target_accounts=target_accounts, fallback_name="仕入先")

def _create_list_common(df_journal: pd.DataFrame, row_idx: int, col_idx: int, target_accounts: list, fallback_name: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    営業先・仕入先リスト作成の共通ロジック。
    """
    # 1. 取引先一覧を抽出
    partners = set()
    # 借方・貸方のいずれかに指定科目が含まれる場合、その行の取引先を取得
    mask_debit = df_journal["debit_account"].fillna("").str.contains("|".join(target_accounts))
    partners.update(df_journal[mask_debit]["debit_partner"].dropna().unique())
    mask_credit = df_journal["credit_account"].fillna("").str.contains("|".join(target_accounts))
    partners.update(df_journal[mask_credit]["credit_partner"].dropna().unique())
    
    partners = {str(p) for p in partners if pd.notna(p) and str(p).strip() != ""}
    partner_list_str = "\n".join(sorted(list(partners)))
    
    # 2. Googleスプレッドシートからプロンプトを取得
    spreadsheet_id = st.secrets["SPREADSHEET_ID"]
    worksheet_name = "AIプロンプト"
    encoded_worksheet = urllib.parse.quote(worksheet_name)
    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_worksheet}"
    
    try:
        df_prompt = pd.read_csv(csv_url, header=None)
        base_prompt = df_prompt.iloc[row_idx, col_idx] if df_prompt.shape[0] > row_idx and df_prompt.shape[1] > col_idx else ""
    except Exception as e:
        print(f"Error fetching prompt: {e}")
        base_prompt = f"以下の取引先一覧から、今後の{fallback_name}候補を提案してください。"

    # 3. プロンプト結合
    full_prompt = f"{base_prompt}\n\n【既存取引先一覧】\n{partner_list_str}"
    
    # 4. Gemini呼び出し
    structured_json_str = exe_gemini_withGoogleSearch_and_structure(full_prompt)
    
    # 5. DataFrame化
    try:
        data = json.loads(structured_json_str)
        business_list_data = data.get("business_list", [])
        df_full = pd.DataFrame(business_list_data)
    except Exception as e:
        print(f"Error parsing Gemini JSON: {e}")
        df_full = pd.DataFrame(columns=["会社名", "ホームページURL", "業種", "事業内容", "登記地域"])

    expected_cols = ["会社名", "ホームページURL", "メールアドレス", "業種", "事業内容", "登記地域"]
    for col in expected_cols:
        if col not in df_full.columns:
            df_full[col] = ""
    df_full = df_full[expected_cols]

    # 免責事項の追加
    disclaimer = "※本リストは生成AIが作成したものであり、誤りが含まれる可能性があります。参考程度にご利用ください。"
    disclaimer_row = pd.DataFrame([{expected_cols[0]: disclaimer}])
    df_full = pd.concat([df_full, disclaimer_row], ignore_index=True)

    # プレビュー用3件
    df_preview = df_full.head(min(3, len(df_full)-1)) if len(df_full) > 1 else df_full.head(3)
    
    return df_full, df_preview
