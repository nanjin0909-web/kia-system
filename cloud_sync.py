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

def get_spreadsheet_id():
    """Secrets에서 Spreadsheet ID를 가져오거나 기본값 반환"""
    try:
        if "google_sheets" in st.secrets and "spreadsheet_id" in st.secrets["google_sheets"]:
            return st.secrets["google_sheets"]["spreadsheet_id"]
    except:
        pass
    return "1XgRR6AG45QwuLBghyv7htc8ICYByvs1DDxhTMzcEY6Y"

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
        st.error(f"Google Sheets 인증 에러: {e}")
    return None

def _ensure_sheet(service, spreadsheet_id, sheet_name):
    """시트(탭)가 없으면 새로 만들기"""
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing = [s['properties']['title'] for s in meta.get('sheets', [])]
        if sheet_name not in existing:
            body = {'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    except Exception as e:
        pass

def save_to_cloud(key, data):
    service = get_sheets_service()
    spreadsheet_id = get_spreadsheet_id()
    if not service or not spreadsheet_id: return False
    try:
        _ensure_sheet(service, spreadsheet_id, key)
        
        # 5만자 제한 해결: 데이터를 3만자 단위로 쪼개서 여러 행에 저장
        json_str = json.dumps(data, ensure_ascii=False)
        chunk_size = 30000
        chunks = [json_str[i:i+chunk_size] for i in range(0, len(json_str), chunk_size)]
        
        # 기존 내용 삭제 후 새로 쓰기
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"{key}!A:A"
        ).execute()
        
        body = {'values': [[c] for c in chunks]}
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{key}!A1",
            valueInputOption='RAW',
            body=body
        ).execute()
        return True
    except Exception as e:
        st.error(f"Cloud 저장 에러 ({key}): {e}")
        return False

def load_from_cloud(key, default):
    service = get_sheets_service()
    spreadsheet_id = get_spreadsheet_id()
    if not service or not spreadsheet_id: return default
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{key}!A:A"
        ).execute()
        values = result.get('values', [])
        if values:
            # 여러 행에 나뉘어 저장된 데이터를 다시 합치기
            full_json_str = "".join([row[0] for row in values if row])
            return json.loads(full_json_str)
    except Exception as e:
        pass
    return default

def get_drive_service():
    return get_sheets_service()
