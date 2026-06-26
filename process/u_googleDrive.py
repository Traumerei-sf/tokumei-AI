import base64
import requests
import streamlit as st

def upload_file_to_drive(file_bytes: bytes, filename: str, mime_type: str) -> str:
    """
    Google Apps Script (GAS) ウェブアプリを経由して、任意のバイナリデータをGoogle Driveへアップロードします。
    ※アップロード先のフォルダIDはGAS側で管理されています。
    
    Parameters:
        file_bytes (bytes): アップロードするファイルのバイナリデータ
        filename (str): アップロードする際のファイル名
        mime_type (str): アップロードするファイルのMIMEタイプ
        
    Returns:
        str: 成功した場合は作成されたファイルのID、失敗した場合はNone
    """
    try:
        # SecretsからGASのURLをロード
        if "GAS_UPLOAD_URL" not in st.secrets:
            print("Google Drive Link Warning: 'GAS_UPLOAD_URL' is not configured in secrets.")
            return None
        
        gas_url = st.secrets["GAS_UPLOAD_URL"]
        if not gas_url or gas_url == "YOUR_GAS_WEB_APP_URL_HERE":
            print("Google Drive Link Warning: 'GAS_UPLOAD_URL' is placeholder or empty.")
            return None

        # 送信データの構築
        payload = {
            "fileName": filename,
            "mimeType": mime_type,
            "fileBytes": base64.b64encode(file_bytes).decode("utf-8")
        }

        # GASウェブアプリへPOST送信
        response = requests.post(
            gas_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code != 200:
            print(f"Google Drive Upload Error: HTTP {response.status_code}")
            return None
            
        result = response.json()
        if result.get("status") == "success":
            file_id = result.get("fileId")
            print(f"Google Drive Upload Success via GAS. File ID: {file_id}")
            return file_id
        else:
            print(f"Google Drive Upload Error via GAS: {result.get('message')}")
            return None
            
    except Exception as e:
        print(f"Google Drive Upload Error (GAS connection): {e}")
        return None

def upload_pdf_to_drive(pdf_bytes: bytes, filename: str) -> str:
    """
    Google Apps Script (GAS) ウェブアプリを経由して、PDFバイナリデータをGoogle Driveへアップロードします。
    ※アップロード先のフォルダIDはGAS側で管理されています。
    
    Parameters:
        pdf_bytes (bytes): アップロードするPDFのバイナリデータ
        filename (str): アップロードする際のファイル名
        
    Returns:
        str: 成功した場合は作成されたファイルのID、失敗した場合はNone
    """
    return upload_file_to_drive(pdf_bytes, filename, "application/pdf")
