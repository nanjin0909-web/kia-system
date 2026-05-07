import os
import json
import io
import streamlit as st

try:
    from googleapiclient.discovery import build
    from google.oauth2.service_account import Credentials
    from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
    GOOGLE_API_AVAILABLE = True
except ImportError:
    GOOGLE_API_AVAILABLE = False

SCOPES = ['https://www.googleapis.com/auth/drive.file']
FOLDER_NAME = "Kia_System_DB"

@st.cache_resource
def get_drive_service():
    if not GOOGLE_API_AVAILABLE: return None
    try:
        if "gcp_service_account" in st.secrets:
            creds_info = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
            return build('drive', 'v3', credentials=creds, cache_discovery=False)
        elif os.path.exists("google_credentials.json"):
            creds = Credentials.from_service_account_file("google_credentials.json", scopes=SCOPES)
            return build('drive', 'v3', credentials=creds, cache_discovery=False)
    except Exception as e:
        print("Drive auth error:", e)
    return None

def get_or_create_folder(service):
    # 폴더 검색
    results = service.files().list(q=f"name='{FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'", spaces='drive').execute()
    items = results.get('files', [])
    if not items:
        # 폴더 생성
        file_metadata = {'name': FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder'}
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')
    return items[0].get('id')

def save_to_cloud(key, data):
    service = get_drive_service()
    if not service: return False
    
    try:
        folder_id = get_or_create_folder(service)
        filename = f"{key}.json"
        
        # 파일이 존재하는지 확인
        results = service.files().list(q=f"name='{filename}' and '{folder_id}' in parents", spaces='drive').execute()
        items = results.get('files', [])
        
        json_str = json.dumps(data, ensure_ascii=False)
        fh = io.BytesIO(json_str.encode('utf-8'))
        media = MediaIoBaseUpload(fh, mimetype='application/json', resumable=True)
        
        if items:
            # 기존 파일 업데이트
            file_id = items[0].get('id')
            service.files().update(fileId=file_id, media_body=media).execute()
        else:
            # 새 파일 생성
            file_metadata = {'name': filename, 'parents': [folder_id]}
            service.files().create(body=file_metadata, media_body=media).execute()
        return True
    except Exception as e:
        print(f"Cloud save error ({key}):", e)
        return False

def load_from_cloud(key, default):
    service = get_drive_service()
    if not service: return default
    
    try:
        folder_id = get_or_create_folder(service)
        filename = f"{key}.json"
        
        results = service.files().list(q=f"name='{filename}' and '{folder_id}' in parents", spaces='drive').execute()
        items = results.get('files', [])
        
        if items:
            file_id = items[0].get('id')
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            fh.seek(0)
            return json.loads(fh.read().decode('utf-8'))
    except Exception as e:
        print(f"Cloud load error ({key}):", e)
    return default
