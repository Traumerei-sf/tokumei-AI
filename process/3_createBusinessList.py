import pandas as pd
import urllib.parse
import json
from typing import Dict, Tuple
from process.u_accessGemini import exe_gemini_withGoogleSearch_and_structure

def create_business_list(df_journal: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    営業先リストを作成する。
    1. 仕訳帳から売上先を抽出
    2. スプレッドシートからプロンプトを取得
    3. Geminiを呼び出してリストを生成
    4. DataFrame化して返す
    
    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: (全件データ, プレビュー用3件)
    """
    
    # 1. 「この会社の売上先」を抽出
    # 売上高、売掛金、受取手形などの勘定科目に関連する取引先を抽出
    sales_accounts = ["売上高", "売掛金", "受取手形"]
    
    partners = set()
    
    # 借方
    mask_debit = df_journal["debit_account"].fillna("").str.contains("|".join(sales_accounts))
    partners.update(df_journal[mask_debit]["debit_partner"].dropna().unique())
    
    # 貸方
    mask_credit = df_journal["credit_account"].fillna("").str.contains("|".join(sales_accounts))
    partners.update(df_journal[mask_credit]["credit_partner"].dropna().unique())
    
    # 空文字やNAを除外
    partners = {str(p) for p in partners if pd.notna(p) and str(p).strip() != ""}
    
    partner_list_str = "\n".join(sorted(list(partners)))
    print(f"DEBUG: Extracted partners: {list(partners)[:5]}...") # 最初の5件のみ
    
    # 2. Googleスプレッドシートからプロンプトを取得
    # auth.pyと同じスプシID。シート名「AIプロンプト」、2行目B列 (index 1, col 1)
    import streamlit as st
    spreadsheet_id = st.secrets["SPREADSHEET_ID"]
    worksheet_name = "AIプロンプト"
    encoded_worksheet = urllib.parse.quote(worksheet_name)
    csv_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&sheet={encoded_worksheet}"
    
    try:
        df_prompt = pd.read_csv(csv_url, header=None)
        # 2行目B列 -> index[1][1]
        base_prompt = df_prompt.iloc[1, 1] if df_prompt.shape[0] > 1 and df_prompt.shape[1] > 1 else ""
    except Exception as e:
        print(f"Error fetching prompt: {e}")
        base_prompt = "以下の取引先一覧から、今後の営業先候補を10件提案してください。"

    # 3. プロンプト結合
    full_prompt = f"{base_prompt}\n\n【既存取引先一覧】\n{partner_list_str}"
    
    # 4. Gemini呼び出し (Google検索 + 構造化)
    structured_json_str = exe_gemini_withGoogleSearch_and_structure(full_prompt)
    print(f"DEBUG: Gemini raw response: {structured_json_str[:200]}...") # 最初のみ
    
    # 5. DataFrame化
    try:
        data = json.loads(structured_json_str)
        # スキーマに基づき business_list キー配下のリストを取得
        business_list_data = data.get("business_list", [])
        df_full = pd.DataFrame(business_list_data)
    except Exception as e:
        print(f"Error parsing Gemini JSON: {e}")
        # フォールバック: ダミーデータ
        df_full = pd.DataFrame(columns=["会社名", "ホームページURL", "業種", "事業内容", "登記地域"])

    # 列順の保証 (要件: 会社名, ホームページURL, 業種, 事業内容, 登記地域)
    expected_cols = ["会社名", "ホームページURL", "業種", "事業内容", "登記地域"]
    for col in expected_cols:
        if col not in df_full.columns:
            df_full[col] = ""
    df_full = df_full[expected_cols]

    # 10件に制限（要件通り）
    df_full = df_full.head(10)
    
    # プレビュー用3件
    df_preview = df_full.head(3)
    
    return df_full, df_preview
