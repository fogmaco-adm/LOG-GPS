import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- KONFIGURASI PATH ---
SCOPES = ['https://www.googleapis.com/auth/drive']
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_OUTPUT_ROOT = os.path.join(BASE_DIR, "3_Output_PDF")

def get_drive_service():
    creds = None
    token_path = os.path.join(BASE_DIR, 'token.json')
    creds_path = os.path.join(BASE_DIR, 'credentials.json')
    
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(creds_path):
                print(f"❌ File 'credentials.json' tidak ditemukan!")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0, prompt='select_account')
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            
    return build('drive', 'v3', credentials=creds)

def get_or_create_folder(service, folder_name, parent_id):
    """Mencari atau membuat subfolder di dalam parent_id tertentu secara aman"""
    clean_name = folder_name.replace("'", "").replace('"', "")
    
    try:
        query = f"name = '{clean_name}' and '{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        if items:
            return items[0]['id']
    except Exception:
        pass
        
    folder_metadata = {
        'name': folder_name,
        'parents': [parent_id],
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=folder_metadata, fields='id').execute()
    return folder.get('id')

def get_or_create_main_folder(service, folder_name):
    """Mencari atau membuat folder utama di root Drive"""
    clean_name = folder_name.replace("'", "").replace('"', "")
    try:
        query = f"name = '{clean_name}' and 'root' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        if items:
            return items[0]['id']
    except Exception:
        pass
        
    folder_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = service.files().create(body=folder_metadata, fields='id').execute()
    return folder.get('id')

def upload_all_timesheets():
    
    folder_input = input("👉 Masukkan nama folder utama di Drive (Contoh: TIMESHEET OKTOBER 2026): ").strip()
    
    if not folder_input:
        print("❌ Nama folder tidak boleh kosong! Program dihentikan.")
        return
        
    service = get_drive_service()
    if not service:
        return
        
    print(f"\n📁 Menyiapkan folder utama '{folder_input}' di Google Drive...")
    main_folder_id = get_or_create_main_folder(service, folder_input)
    
    if not os.path.exists(PATH_OUTPUT_ROOT):
        print(f"❌ Folder output lokal '{PATH_OUTPUT_ROOT}' tidak ditemukan!")
        return
        
    print(f"Memulai proses upload sesuai struktur folder...\n")
    uploaded_count = 0
    
    # Dictionary untuk menyimpan ID folder di Drive agar tidak berulang kali membuat/mencari
    folder_cache = {".": main_folder_id}
    
    for root, dirs, files in os.walk(PATH_OUTPUT_ROOT):
        relative_path = os.path.relpath(root, PATH_OUTPUT_ROOT)
        
        # Tentukan ID parent folder untuk file-file di dalam root ini
        current_parent_id = main_folder_id
        
        if relative_path != ".":
            # Pisahkan path bertingkat (misal: BMP/Admin Inventory - Mining)
            path_parts = relative_path.split(os.sep)
            accumulated_path = ""
            
            for part in path_parts:
                parent_key = accumulated_path if accumulated_path else "."
                parent_id = folder_cache[parent_key]
                
                accumulated_path = os.path.join(accumulated_path, part) if accumulated_path else part
                
                if accumulated_path not in folder_cache:
                    folder_cache[accumulated_path] = get_or_create_folder(service, part, parent_id)
                    
            current_parent_id = folder_cache[relative_path]
            subfolder_name_display = relative_path
        else:
            subfolder_name_display = "Utama"
        
        for file in files:
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                
                print(f"Mengunggah [{subfolder_name_display}] {file}...")
                
                file_metadata = {
                    'name': file,
                    'parents': [current_parent_id]
                }
                
                media = MediaFileUpload(file_path, resumable=True)
                
                try:
                    uploaded_file = service.files().create(
                        body=file_metadata,
                        media_body=media,
                        fields='id, name'
                    ).execute()
                    
                    print(f"✅ Berhasil diunggah: {uploaded_file.get('name')}")
                    uploaded_count += 1
                except Exception as e:
                    print(f"❌ Gagal mengunggah {file}: {e}")
                    
    print(f"\nSelesai! Total {uploaded_count} file PDF Timesheet berhasil diunggah dengan struktur hierarki yang rapi.")

if __name__ == "__main__":
    upload_all_timesheets()