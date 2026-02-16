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
                            "業種": { "type": "STRING" },
                            "事業内容": { "type": "STRING" },
                            "登記地域": { "type": "STRING" }
                        },
                        "required": ["会社名", "ホームページURL", "業種", "事業内容", "登記地域"]
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


