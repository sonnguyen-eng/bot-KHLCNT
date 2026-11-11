import os
import json
import io
import time
import hashlib
import sys
from datetime import datetime, timezone
from google import genai
from google.genai import types, errors
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ============ 1. BIẾN MÔI TRƯỜNG & CẤU HÌNH ============
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
TARGET_TAGS = [t.strip().lower() for t in os.environ.get("TARGET_TAGS", "").split(",") if t.strip()]

# Danh sách Model ưu tiên (Mặc định gemini-3.8-flash)
GEMINI_MODELS = [m.strip() for m in os.environ.get("GEMINI_MODELS", "gemini-3.8-flash").split(",") if m.strip()]
MAX_RETRIES = int(os.environ.get("MAX_RETRIES", "3"))

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.png', '.jpg', '.jpeg', '.txt')
FOLDER_MIME = 'application/vnd.google-apps.folder'

# ============ 2. KHỞI TẠO CLIENTS & PROMPT ============
client = genai.Client(api_key=GEMINI_API_KEY)

PROMPT_KHOA_HOC = """
Bạn là một trợ lý phân tích dữ liệu khoa học / kỹ thuật chuyên nghiệp.
Hãy phân tích tài liệu được cung cấp (bao gồm cả văn bản, hình ảnh, bảng biểu) và trích xuất thông tin theo cấu trúc JSON chuẩn sau:

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

QUY TẮC BẮT BUỘC:
- Chỉ sử dụng thông tin có trong tài liệu. Không tự suy đoán.
- Giữ nguyên số liệu, đơn vị đo, mã ký hiệu kỹ thuật như tài liệu gốc.
- Trích xuất đầy đủ các bảng biểu thông số kỹ thuật vào trường "technical_specifications".
- Chỉ trả về duy nhất chuỗi JSON hợp lệ, không kèm bất kỳ lời giải thích hay định dạng markdown bổ sung nào.
"""

def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)

def calculate_md5(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def is_matching_tags(file_name, folder_path_str):
    if not TARGET_TAGS:
        return True
    combined_text = f"{file_name} {folder_path_str}".lower()
    return any(tag in combined_text for tag in TARGET_TAGS)

# ============ 3. TẢI VÀ GHI DỮ LIỆU DRIVE ============
def download_file_bytes(service, file_id):
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return file_stream.getvalue()

def save_json_to_drive(service, folder_id, file_name, json_data, existing_file_id=None):
    json_bytes = io.BytesIO(json.dumps(json_data, ensure_ascii=False, indent=2).encode('utf-8'))
    media = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=True)
    
    if existing_file_id:
        service.files().update(
            fileId=existing_file_id,
            media_body=media,
            supportsAllDrives=True
        ).execute()
    else:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()

# ============ 4. GỌI GEMINI API ============
def call_gemini_multimodal(file_bytes, mime_type, file_name):
    """Tải trực tiếp byte dữ liệu lên Gemini qua API File và gọi Model trích xuất."""
    # Tạo file tạm địa phương phục vụ client.files.upload
    temp_path = f"temp_upload_{int(time.time())}_{file_name}"
    with open(temp_path, "wb") as f:
        f.write(file_bytes)
        
    try:
        uploaded_file = client.files.upload(file=temp_path)
        
        data = None
        used_model = None
        
        for model in GEMINI_MODELS:
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    response = client.models.generate_content(
                        model=model,
                        contents=[uploaded_file, PROMPT_KHOA_HOC],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.1
                        )
                    )
                    data = json.loads(response.text)
                    used_model = model
                    break
                except errors.APIError as e:
                    code = getattr(e, "code", None)
                    if code in (429, 500, 503):
                        wait = 5 * (2 ** (attempt - 1))
                        print(f"   ⚠️ Model {model} gặp lỗi {code} (lần {attempt}/{MAX_RETRIES}), chờ {wait}s...")
                        time.sleep(wait)
                    else:
                        print(f"   ⚠️ Model {model} lỗi {code}: {e}. Chuyển model tiếp theo.")
                        break
                except (ValueError, TypeError) as e:
                    print(f"   ⚠️ Model {model} trả JSON không hợp lệ ({e}), thử lại...")
                    time.sleep(2)
            
            if data:
                break

        # Xóa file đã upload trên Gemini Storage
        try:
            client.files.delete(name=uploaded_file.name)
        except Exception:
            pass

        if not data:
            raise RuntimeError("Tất cả mô hình Gemini đều xử lý thất bại.")
            
        return data, used_model

    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# ============ 5. QUÉT ĐỆ QUY VÀ XỬ LÝ THƯ MỤC ============
def scan_and_process_folder(service, folder_id, stats, current_path=""):
    items = []
    page_token = None
    
    # Phân trang toàn bộ item trong thư mục
    while True:
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        items.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break

    # Lập danh sách các file JSON đã tồn tại trong thư mục hiện tại
    json_files_map = {f['name']: f['id'] for f in items if f['name'].endswith('.json')}

    for item in items:
        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        # Quét đệ quy nếu là thư mục con
        if mime_type == FOLDER_MIME:
            scan_and_process_folder(service, file_id, stats, f"{current_path}/{file_name}")
            continue

        # Kiểm tra phần mở rộng định dạng hỗ trợ
        if not file_name.lower().endswith(ALLOWED_EXTENSIONS):
            continue

        # Lọc theo TARGET_TAGS nếu có cấu hình
        if not is_matching_tags(file_name, current_path):
            continue

        base_name = os.path.splitext(file_name)[0]
        expected_json_name = f"{base_name}_PROCESSED_V1.json"
        existing_json_id = json_files_map.get(expected_json_name)

        print(f"\n📄 Phát hiện tài liệu: {file_name}")
        print(f"📂 Vị trí: {current_path if current_path else 'Thư mục gốc'}")

        try:
            # Tải nội dung byte & tính toán MD5
            file_bytes = download_file_bytes(service, file_id)
            current_md5 = calculate_md5(file_bytes)

            # Kiểm tra xem file JSON đã tồn tại và so sánh Hash MD5
            if existing_json_id:
                try:
                    existing_bytes = download_file_bytes(service, existing_json_id)
                    existing_json_data = json.loads(existing_bytes.decode('utf-8'))
                    old_md5 = existing_json_data.get("metadata", {}).get("file_hash_md5")

                    if old_md5 == current_md5:
                        print(f"⏩ Bỏ qua (Nội dung không đổi - MD5 trùng khớp): {file_name}")
                        stats['skipped'] += 1
                        continue
                    else:
                        print(f"🔄 File nguồn đã thay đổi (MD5 khác biệt). Tiến hành xử lý lại...")
                except Exception:
                    print("⚠️ Không thể đọc MD5 cũ từ file JSON hiện có. Tiếp tục xử lý...")

            print(f"🔑 Mã Hash MD5: {current_md5}")
            print(f"🚀 Gửi dữ liệu sang Gemini trích xuất...")

            # Phân tích qua Gemini API Multi-modal
            json_result, used_model = call_gemini_multimodal(file_bytes, mime_type, file_name)

            # Chuẩn hóa Metadata & Đóng gói
            if "metadata" not in json_result or not isinstance(json_result["metadata"], dict):
                json_result["metadata"] = {}
            if "extracted_data" not in json_result or not isinstance(json_result["extracted_data"], dict):
                json_result["extracted_data"] = {}

            json_result["metadata"].update({
                "file_hash_md5": current_md5,
                "original_filename": file_name,
                "source_drive_id": file_id,
                "source_path": current_path or "/",
                "processed_at": datetime.now(timezone.utc).isoformat(),
                "processed_by_model": used_model
            })

            # Lưu file kết quả .json trực tiếp tại thư mục con hiện tại
            save_json_to_drive(service, folder_id, expected_json_name, json_result, existing_file_id=existing_json_id)

            print(f"✅ Đã tạo/cập nhật file thành công: {expected_json_name} (Model: {used_model})")
            stats['ok'] += 1

        except Exception as e:
            print(f"❌ Lỗi khi xử lý file {file_name}: {str(e)}")
            stats['failed'] += 1

def main():
    service = get_drive_service()
    if TARGET_TAGS:
        print(f"🔍 Đang chạy chế độ lọc theo từ khóa: {TARGET_TAGS}")
    else:
        print("🌐 Đang chạy chế độ quét toàn bộ thư viện...")
        
    print(f"🤖 Danh sách Mô hình Ưu tiên: {GEMINI_MODELS}")
    
    stats = {'ok': 0, 'skipped': 0, 'failed': 0}
    scan_and_process_folder(service, DRIVE_FOLDER_ID, stats)
    
    print(f"\n📊 TỔNG KẾT TIẾN TRÌNH:")
    print(f"   ✔️ Thành công: {stats['ok']}")
    print(f"   ⏩ Bỏ qua:     {stats['skipped']}")
    print(f"   ❌ Thất bại:   {stats['failed']}")

    if stats['failed'] > 0:
        sys.exit(1) # Báo đỏ GitHub Actions nếu có lỗi phát sinh

if __name__ == "__main__":
    main()
