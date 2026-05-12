import os
import json
import streamlit as st

try:
    from googleapiclient.discovery import build
    from google.oauth2.service_account import Credentials
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = "1XgRR6AG45QwuLBghyv7htc8ICYByvs1DDxhTMzcEY6Y"

@st.cache_resource
def get_sheets_service():
    if not GOOGLE_API_AVAILABLE: return None
    try:
        has_secrets = False
        try:
            if "gcp_service_account" in st.secrets:
                has_secrets = True
        except Exception:
            pass

        if has_secrets:
            creds_info = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        elif os.path.exists("google_credentials.json"):
            creds = Credentials.from_service_account_file("google_credentials.json", scopes=SCOPES)
        else:
            return None
        return build('sheets', 'v4', credentials=creds, cache_discovery=False)
    except Exception as e:
        print("Sheets auth error:", e)
    return None

def _ensure_sheet(service, sheet_name):
    """시트(탭)가 없으면 새로 만들기"""
    try:
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        existing = [s['properties']['title'] for s in meta.get('sheets', [])]
        if sheet_name not in existing:
            body = {'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
            service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    except Exception as e:
        print(f"Sheet ensure error ({sheet_name}):", e)

def save_to_cloud(key, data):
    service = get_sheets_service()
    if not service: return False
    try:
        _ensure_sheet(service, key)
        json_str = json.dumps(data, ensure_ascii=False)
        body = {'values': [[json_str]]}
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{key}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()
        return True
    except Exception as e:
        print(f"Cloud save error ({key}):", e)
        return False

def load_from_cloud(key, default):
    service = get_sheets_service()
    if not service: return default
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{key}!A1"
        ).execute()
        values = result.get('values', [])
        if values and values[0]:
            return json.loads(values[0][0])
    except Exception as e:
        print(f"Cloud load error ({key}):", e)
    return default

# 하위 호환성 유지 (기존 코드에서 get_drive_service를 호출하는 경우)
def get_drive_service():
    return get_sheets_service()
