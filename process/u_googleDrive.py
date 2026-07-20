import streamlit as st
import requests
import base64
import json

def upload_file_to_drive(file_bytes: bytes, filename: str, mime_type: str) -> str:
    """
    Google Apps Script (GAS) の Web API を経由してファイルをアップロードします。
    """
    try:
        # secrets から GAS の URL を取得
        gas_url = st.secrets.get("GAS_DRIVE_UPLOAD_URL")
        if not gas_url:
            print("Error: GAS_DRIVE_UPLOAD_URL is not set in secrets.toml")
            return None
        
        # バイナリを Base64 エンコード
        encoded_data = base64.b64encode(file_bytes).decode("utf-8")
        
        payload = {
            "fileName": filename,
            "mimeType": mime_type,
            "fileBytes": encoded_data
        }
        
        # POSTリクエスト (GASはリダイレクトを返すため allow_redirects=True が必須)
        response = requests.post(
            gas_url,
            json=payload,
            headers={'Content-Type': 'application/json'},
            allow_redirects=True,
            timeout=60
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get("status") == "success":
                file_id = result.get("fileId")
                print(f"Google Drive Upload Success (via GAS). File ID: {file_id}")
                return file_id
            else:
                print(f"Google Drive Upload Error (GAS returned error): {result.get('message')}")
                return None
        else:
            print(f"Google Drive Upload HTTP Error: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        print(f"Google Drive Upload Exception: {e}")
        return None

def upload_pdf_to_drive(pdf_bytes: bytes, filename: str) -> str:
    """
    PDFバイナリデータをGoogle Driveへアップロードします。
    """
    return upload_file_to_drive(pdf_bytes, filename, "application/pdf")
