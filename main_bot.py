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
    """Gửi request tới Gemini API với cơ chế Retry khi gặp Quota / Server Busy"""
    backoff_delays = [15, 30, 45, 60, 90]
    
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
            if any(k in error_msg.lower() for k in ["429", "resource_exhausted", "quota", "503", "unavailable"]):
                wait_time = backoff_delays[min(attempt, len(backoff_delays) - 1)]
                print(f"⚠️ Giới hạn Token/Rate Limit hoặc Server bận. Tạm dừng {wait_time} giây trước khi thử lại...")
                time.sleep(wait_time)
            else:
                raise e
    raise RuntimeError(f"❌ Thất bại sau {max_retries} lần thử do lỗi kết nối API.")

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

    # Gom danh sách các file JSON đã tồn tại trong thư mục hiện tại để chống trùng lặp
    json_files = {f['name'] for f in items if f['name'].endswith('.json')}

    for item in items:
        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        if mime_type == 'application/vnd.google-apps.folder':
            collect_all_files(service, file_id, f"{current_path}/{file_name}", collected_list)
        elif file_name.lower().endswith(ALLOWED_EXTENSIONS):
            if not is_matching_tags(file_name, current_path):
                continue

            base_name = os.path.splitext(file_name)[0]
            expected_json_v1 = f"{base_name}_PROCESSED_V1.json"
            
            if expected_json_v1 in json_files:
                print(f"⏩ Bỏ qua (Đã xử lý trước đó): {file_name}")
                continue

            collected_list.append({
                "file_id": file_id,
                "file_name": file_name,
                "folder_id": folder_id,
                "folder_path": current_path if current_path else "Thư mục gốc",
                "expected_json": expected_json_v1
            })

    return collected_list

# 4. BƯỚC 2: XỬ LÝ DỨT ĐIỂM TUẦN TỰ TỪNG FILE
def process_single_file(service, file_info):
    file_id = file_info["file_id"]
    file_name = file_info["file_name"]
    folder_id = file_info["folder_id"]
    folder_path = file_info["folder_path"]
    expected_json_v1 = file_info["expected_json"]

    print(f"\n--------------------------------------------------")
    print(f"📄 Đang xử lý dứt điểm file: {file_name}")
    print(f"📂 Vị trí: {folder_path}")

    # 1. Tải file từ Drive
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    
    file_bytes = file_stream.getvalue()
    current_file_hash = calculate_md5(file_bytes)
    print(f"🔑 Mã Hash MD5: {current_file_hash}")

    local_file_path = f"temp_{file_name}"
    with open(local_file_path, 'wb') as f:
        f.write(file_bytes)

    uploaded_file = None
    try:
        # 2. Upload lên Gemini File API
        uploaded_file = client.files.upload(file=local_file_path)
        
        # 3. Phân tích tài liệu bằng Gemini 3.8 Flash
        response = call_gemini_with_retry(uploaded_file, PROMPT_KHOA_HOC)
        
        # 4. Tạo nội dung JSON kết quả
        json_result = json.loads(response.text)
        json_result["metadata"]["file_hash_md5"] = current_file_hash
        json_result["metadata"]["original_filename"] = file_name
        
        # 5. Ghi trực tiếp JSON kết quả lên Drive
        json_bytes = io.BytesIO(json.dumps(json_result, ensure_ascii=False, indent=2).encode('utf-8'))
        media = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=True)
        
        file_metadata = {
            'name': expected_json_v1,
            'parents': [folder_id]
        }
        
        service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        
        print(f"✅ Hoàn tất xử lý & Đã ghi file JSON: {expected_json_v1}")

    except Exception as e:
        print(f"❌ Thất bại khi xử lý file {file_name}: {str(e)}")
    
    finally:
        # 6. Dọn dẹp tài nguyên dứt điểm sau mỗi file (Tránh tràn RAM & hết Token)
        if uploaded_file is not None:
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass

        if os.path.exists(local_file_path):
            os.remove(local_file_path)

        # Nghỉ ngắn giữa các file để giải phóng Rate Limit/Token
        time.sleep(3)

def main():
    service = get_drive_service()
    if TARGET_TAGS:
        print(f"🔍 Đang chạy chế độ lọc theo từ khóa: {TARGET_TAGS}")
    else:
        print("🌐 Đang chạy chế độ quét toàn bộ thư viện...")
        
    print("⏳ Bước 1: Quét toàn bộ thư mục và lập danh sách file...")
    files_to_process = collect_all_files(service, DRIVE_FOLDER_ID)
    total_files = len(files_to_process)
    
    print(f"📊 Tìm thấy tổng cộng {total_files} file cần trích xuất.")
    print("🚀 Bước 2: Tiến hành xử lý dứt điểm từng file...")

    for index, file_info in enumerate(files_to_process, start=1):
        print(f"\n[Tiến độ: {index}/{total_files}]")
        process_single_file(service, file_info)

    print("\n🎉 Đã hoàn thành xử lý toàn bộ thư viện dữ liệu!")

if __name__ == "__main__":
    main()
