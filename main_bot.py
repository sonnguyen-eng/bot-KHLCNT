import os
import json
import io
import re
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# 1. LẤY BIẾN MÔI TRƯỜNG TỪ GITHUB SECRETS
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')

# 2. KHỞI TẠO CLIENTS
client = genai.Client(api_key=GEMINI_API_KEY)

def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

PROMPT_KHOA_HOC = """
Bạn là một trợ lý phân tích dữ liệu khoa học / kỹ thuật chuyên nghiệp.
Hãy phân tích tài liệu được cung cấp và trích xuất thông tin theo cấu trúc JSON chuẩn sau:

{
  "metadata": {
    "model_name": "Tên Model / Thiết bị / Sản phẩm chính",
    "document_type": "Loại tài liệu (Brochure / IFU / SpecSheet / DataSheet / Manual / TestReport / Certificate)",
    "version": "Số phiên bản (VD: V1, V2...) nếu không có thì ghi V1",
    "suggested_standard_name": "Tên đặt chuẩn hóa: [Model]_[DocType]_V[Version]"
  },
  "extracted_data": {
    "title": "Tiêu đề chính của tài liệu",
    "summary": "Tóm tắt ngắn gọn nội dung khoa học/kỹ thuật cốt lõi (100-200 từ)",
    "technical_specifications": {},
    "key_notes": "Lưu ý đặc biệt, kết luận kiểm định hoặc quy chuẩn vận hành"
  }
}
Chỉ trả về JSON hợp lệ, không kèm lời giải thích văn bản.
"""

def scan_and_process_folder(service, folder_id):
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    existing_jsons = {f['name'] for f in items if f['name'].endswith('_PROCESSED.json')}

    for item in items:
        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        # Nếu là thư mục con -> Quét tiếp đệ quy
        if mime_type == 'application/vnd.google-apps.folder':
            scan_and_process_folder(service, file_id)

        # Nếu là file tài liệu khoa học
        elif file_name.lower().endswith(ALLOWED_EXTENSIONS):
            base_name = os.path.splitext(file_name)[0]
            expected_json_name = f"{base_name}_PROCESSED.json"

            if expected_json_name in existing_jsons:
                print(f"⏩ Bỏ qua (Đã xử lý): {file_name}")
                continue

            print(f"\n📄 Đang xử lý file mới: {file_name}")
            
            # Tải file tạm thời về máy chủ GitHub
            request = service.files().get_media(fileId=file_id)
            local_file_path = f"temp_{file_name}"
            with open(local_file_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

            # Gửi file sang Gemini API để xử lý
            try:
                uploaded_file = client.files.upload(file=local_file_path)
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=[uploaded_file, PROMPT_KHOA_HOC],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                client.files.delete(name=uploaded_file.name)
                
                # Lưu file JSON cục bộ
                local_json_path = f"{base_name}_PROCESSED.json"
                with open(local_json_path, 'w', encoding='utf-8') as f:
                    f.write(response.text)

                # Tải file JSON ngược lên Google Drive
                file_metadata = {'name': expected_json_name, 'parents': [folder_id]}
                media = MediaFileUpload(local_json_path, mimetype='application/json')
                service.files().create(body=file_metadata, media_body=media).execute()
                print(f"✅ Đã xuất file JSON thành công lên Drive!")

                # Dọn dẹp file tạm
                if os.path.exists(local_file_path):
                    os.remove(local_file_path)
                if os.path.exists(local_json_path):
                    os.remove(local_json_path)

            except Exception as e:
                print(f"❌ Lỗi xử lý {file_name}: {str(e)}")

def main():
    service = get_drive_service()
    scan_and_process_folder(service, DRIVE_FOLDER_ID)

if __name__ == "__main__":
    main()
