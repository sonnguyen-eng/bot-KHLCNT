import os
import json
import io
import time
import hashlib
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# 1. BIẾN MÔI TRƯỜNG & SECRETS
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
TARGET_TAGS = os.environ.get("TARGET_TAGS", "").split(",") if os.environ.get("TARGET_TAGS") else []

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')
PRIMARY_MODEL = "gemini-3.8-flash"

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
    "suggested_standard_name": "Tên đặt chuẩn hóa: [Model]_[DocType]_V[Version]",
    "search_tags": ["Danh sách 3-5 từ khóa chính phục vụ truy vấn nhanh"]
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

def calculate_md5(file_content_bytes):
    return hashlib.md5(file_content_bytes).hexdigest()

def is_matching_tags(file_name, folder_path_str):
    if not TARGET_TAGS:
        return True
    combined_text = f"{file_name} {folder_path_str}".lower()
    for tag in TARGET_TAGS:
        tag_clean = tag.strip().lower()
        if tag_clean and tag_clean in combined_text:
            return True
    return False

def call_gemini_with_retry(uploaded_file, prompt, max_retries=5):
    """Gửi request tới Gemini API với cơ chế Retry tăng dần khi chạm hạn ngạch Token / Rate Limit"""
    backoff_delays = [20, 40, 60, 90, 120]
    
    for attempt in range(max_retries):
        try:
            print(f"🤖 Đang gọi Gemini ({PRIMARY_MODEL}) - Thử lần {attempt + 1}/{max_retries}...")
            response = client.models.generate_content(
                model=PRIMARY_MODEL,
                contents=[uploaded_file, prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            return response
        except Exception as e:
            error_msg = str(e)
            print(f"⚠️ Chi tiết lỗi API: {error_msg}")
            if any(k in error_msg.lower() for k in ["429", "resource_exhausted", "quota", "503", "unavailable"]):
                wait_time = backoff_delays[min(attempt, len(backoff_delays) - 1)]
                print(f"⏳ Giới hạn Token/Rate Limit hoặc Server bận. Tạm dừng {wait_time} giây trước khi thử lại...")
                time.sleep(wait_time)
            else:
                raise e
    raise RuntimeError(f"❌ Thất bại sau {max_retries} lần thử do vượt hạn ngạch API.")

# 3. BƯỚC 1: QUÉT THƯ MỤC VÀ THU THẬP DANH SÁCH FILE CẦN XỬ LÝ
def collect_all_files(service, folder_id, current_path="", collected_list=None):
    if collected_list is None:
        collected_list = []

    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query, 
        fields="files(id, name, mimeType, md5Checksum)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    items = results.get('files', [])

    json_files = {f['name'] for f in items if f['name'].endswith('.json')}
