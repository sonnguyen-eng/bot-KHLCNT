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

# 1. BIẾN MÔI TRƯỜNG & SECRETS (HỖ TRỢ NHIỀU API KEYS CHỐNG HẾT QUOTA)
RAW_GEMINI_KEYS = os.environ.get("GEMINI_API_KEY", "")
GEMINI_API_KEYS = [k.strip() for k in RAW_GEMINI_KEYS.split(",") if k.strip()]
current_key_index = 0

DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
TARGET_TAGS = os.environ.get("TARGET_TAGS", "").split(",") if os.environ.get("TARGET_TAGS") else []

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')
PRIMARY_MODEL = "gemini-3.8-flash"

# 2. KHỞI TẠO SERVICES
def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

def get_current_gemini_client():
    global current_key_index, GEMINI_API_KEYS
    if not GEMINI_API_KEYS:
        raise ValueError("❌ Không tìm thấy GEMINI_API_KEY nào trong biến môi trường!")
    key = GEMINI_API_KEYS[current_key_index]
    print(f"🔑 Đang sử dụng Gemini API Key chỉ số [{current_key_index + 1}/{len(GEMINI_API_KEYS)}]")
    return genai.Client(api_key=key)

def switch_to_next_gemini_key():
    global current_key_index, GEMINI_API_KEYS
    if current_key_index + 1 < len(GEMINI_API_KEYS):
        current_key_index += 1
        print(f"🔄 Đã tự động đổi sang API Key tiếp theo: Key [{current_key_index + 1}/{len(GEMINI_API_KEYS)}]")
        return True
    else:
        print("⚠️ Tất cả các API Key Gemini trong danh sách đều đã đụng trần hạn ngạch ngày (Daily Quota)!")
        return False

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

def call_gemini_with_retry(local_file_path, prompt, max_retries=3):
    """Upload và gọi Gemini API với cơ chế tự động xoay vòng API Keys khi hết Hạn ngạch (Quota 429)"""
    global current_key_index
    
    while True:
        client = get_current_gemini_client()
        uploaded_file = None
        try:
            # 1. Upload file tạm lên Gemini bằng Client hiện tại
            uploaded_file = client.files.upload(file=local_file_path)
            
            # 2. Gọi Model Phân Tích
            print(f"🤖 Đang gọi Gemini ({PRIMARY_MODEL}) trích xuất dữ liệu...")
            response = client.models.generate_content(
                model=PRIMARY_MODEL,
                contents=[uploaded_file, prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            
            # 3. Dọn dẹp file trên Gemini sau khi gọi xong
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass
                
            return response

        except Exception as e:
            error_msg = str(e)
            print(f"⚠️ Chi tiết thông báo API: {error_msg}")
            
            # Xóa file vừa upload nếu lỗi
            if uploaded_file is not None:
                try:
                    client.files.delete(name=uploaded_file.name)
                except Exception:
                    pass

            # Nếu chạm trần Quota ngày (429 / RESOURCE_EXHAUSTED)
            if any(k in error_msg.lower() for k in ["429", "resource_exhausted", "quota"]):
                print("🛑 API Key hiện tại đã hết Hạn ngạch (Quota Exceeded).")
                if switch_to_next_gemini_key():
                    time.sleep(2)
                    continue  # Thử lại ngay với Key mới
                else:
                    raise RuntimeError("❌ Đã hết toàn bộ API Keys khả dụng do chạm trần hạn ngạch ngày.")
            
            # Lỗi server tạm thời (503 / Unavailable)
            elif any(k in error_msg.lower() for k in ["503", "unavailable"]):
                print("⏳ Máy chủ Gemini bận, tạm dừng 20 giây...")
                time.sleep(20)
            else:
                raise e

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
                print(f"⏩ Bỏ qua (Đã có file JSON kết quả): {file_name}")
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

    try:
        # 2. Upload & Phân tích bằng Gemini API (Tự xoay API Key nếu hết hạn ngạch)
        response = call_gemini_with_retry(local_file_path, PROMPT_KHOA_HOC)
        
        # 3. Tạo nội dung JSON kết quả
        json_result = json.loads(response.text)
        json_result["metadata"]["file_hash_md5"] = current_file_hash
        json_result["metadata"]["original_filename"] = file_name
        
        # 4. Ghi trực tiếp JSON kết quả lên Drive (Đã giải quyết lỗi 403 Quota)
        json_bytes = io.BytesIO(json.dumps(json_result, ensure_ascii=False, indent=2).encode('utf-8'))
        media = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=False)
        
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
        
        print(f"✅ Hoàn tất xử lý & Đã ghi file JSON thành công: {expected_json_v1}")

    except Exception as e:
        print(f"❌ Thất bại khi xử lý file {file_name}: {str(e)}")
    
    finally:
        # 5. Dọn dẹp file tạm local dứt điểm
        if os.path.exists(local_file_path):
            os.remove(local_file_path)

        # Nghỉ ngắn 3 giây giữa các file
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
