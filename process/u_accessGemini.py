# プロンプトは、Google スプレッドシートから取得する予定
# REST API形式で書きたい。

import requests
import json
import streamlit as st

# APIの設定
# APIキーは .streamlit/secrets.toml または Streamlit Cloud の Secrets から取得
API_KEY = st.secrets["GEMINI_API_KEY"]
# モデル名は、検索機能と構造化出力を一回のリクエストで実行するためにgemini 3以上が必要
MODEL_NAME = "gemini-3-flash-preview" 
URL = f"https://generativelanguage.googleapis.com/v1beta/models/{MODEL_NAME}:generateContent"

def exe_gemini_withGoogleSearch_and_structure(prompt: str, schema: dict = None) -> str:
    """
    Google検索機能と構造化出力を一回のリクエストで実行する
    """
    
    # デフォルトの営業先リスト用スキーマ
    if schema is None:
        schema = {
            "type": "OBJECT",
            "properties": {
                "business_list": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "会社名": { "type": "STRING" },
                            "ホームページURL": { "type": "STRING" },
                            "メールアドレス": { "type": "STRING" },
                            "業種": { "type": "STRING" },
                            "事業内容": { "type": "STRING" },
                            "登記地域": { "type": "STRING" }
                        },
                        "required": ["会社名", "ホームページURL", "メールアドレス", "業種", "事業内容", "登記地域"]
                    }
                }
            },
            "required": ["business_list"]
        }

    payload = {
        "contents": [
            {
                "parts": [
                    { "text": prompt }
                ]
            }
        ],
        "tools": [
            { "google_search": {} }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": schema
        }
    }

    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": API_KEY
    }

    try:
        response = requests.post(URL, headers=headers, json=payload)
        response.raise_for_status()
        
        data = response.json()
        
        # モデルの応答テキストを取得（JSON形式の文字列として返ってくるはず）
        if "candidates" in data and len(data["candidates"]) > 0:
            content = data["candidates"][0].get("content", {})
            parts = content.get("parts", [])
            if parts:
                return parts[0].get("text", "")
        
        return "{}"
    except Exception as e:
        print(f"DEBUG: Gemini API Error: {e}")
        if 'response' in locals():
            print(f"DEBUG: Status Code: {response.status_code}")
            print(f"DEBUG: Response Text: {response.text}")
        return "{}"


def exe_gemini_structure_forJournal(prompt: str) -> str:
    """
    仕訳帳のCSV/テキストから、必要な列のインデックスとデータ開始行を特定する
    """
    schema = {
        "type": "OBJECT",
        "properties": {
            "column_mapping": {
                "type": "OBJECT",
                "properties": {
                    "date": { "type": "INTEGER", "description": "取引日の列番号(0開始)。不明ならnull" },
                    "debit_account": { "type": "INTEGER", "description": "借方科目の列番号(0開始)。不明ならnull" },
                    "debit_amount": { "type": "INTEGER", "description": "借方金額の列番号(0開始)。不明ならnull" },
                    "credit_account": { "type": "INTEGER", "description": "貸方科目の列番号(0開始)。不明ならnull" },
                    "credit_amount": { "type": "INTEGER", "description": "貸方金額の列番号(0開始)。不明ならnull" },
                    "partner": { "type": "INTEGER", "description": "取引先名称や摘要（相手先）の列番号(0開始)。借方・貸方などの区別は不要で、1つだけ特定してください。不明ならnull" },
                    "created_at": { "type": "INTEGER", "description": "作成日時/入力日時の列番号(0開始)。不明ならnull" }
                },
                "required": ["date", "debit_account", "debit_amount", "credit_account", "credit_amount", "partner", "created_at"]
            },
            "data_start_row": { "type": "INTEGER", "description": "実データが始まる行番号(0開始)" }
        },
        "required": ["column_mapping", "data_start_row"]
    }

    payload = {
        "contents": [
            {
                "parts": [
                    { "text": prompt }
                ]
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": schema
        }
    }

    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": API_KEY
    }

    try:
        response = requests.post(URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        
        if "candidates" in data and len(data["candidates"]) > 0:
            content = data["candidates"][0].get("content", {})
            parts = content.get("parts", [])
            if parts:
                return parts[0].get("text", "")
        
        return "{}"
    except Exception as e:
        # スレッド内で st.error は使えないため、戻り値でエラーを表現するかログに出す
        print(f"Gemini API Error: {e}")
        return "{}"

def exe_gemini_structure_forBS(prompt: str) -> str:
    """
    貸借対照表からのデータ抽出のため
    """
    schema = {
        "type": "OBJECT",
        "properties": {
            "year_month": { "type": "STRING", "description": "期末の年月（YYYY/MM形式）。不明な場合はnull" },
            "cash_amount": { "type": "NUMBER", "description": "現預金の合計金額。不明な場合はnull" }
        },
        "required": ["year_month", "cash_amount"]
    }

    payload = {
        "contents": [
            {
                "parts": [
                    { "text": prompt }
                ]
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": schema
        }
    }

    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": API_KEY
    }

    try:
        response = requests.post(URL, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        
        data = response.json()
        
        # モデルの応答テキストを取得
        if "candidates" in data and len(data["candidates"]) > 0:
            content = data["candidates"][0].get("content", {})
            parts = content.get("parts", [])
            if parts:
                text = parts[0].get("text", "")
                try:
                    import json
                    parsed = json.loads(text)
                    print("--- DEBUG: BS Gemini Result ---")
                    print(f"期末の年月: {parsed.get('year_month')}")
                    print(f"現預金の合計金額: {parsed.get('cash_amount')}")
                except:
                    pass
                return text
        
        return "{}"
    except Exception as e:
        print(f"DEBUG: Gemini API Error: {e}")
        if 'response' in locals():
            print(f"DEBUG: Status Code: {response.status_code}")
            print(f"DEBUG: Response Text: {response.text}")
        return "{}"