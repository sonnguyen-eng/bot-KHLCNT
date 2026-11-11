import os
import json
import io
import hashlib
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# 1. LẤY BIẾN MÔI TRƯỜNG BẢO MẬT TỪ GITHUB SECRETS
# ==============================================================================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")

# Lọc từ khóa/Model chỉ định (Ví dụ: ["EP6000"]) - Nếu để trống [] sẽ quét toàn bộ
TARGET_TAGS = os.environ.get("TARGET_TAGS", "").split(",") if os.environ.get("TARGET_TAGS") else []

# Các định dạng tài liệu khoa học cho phép quét và xử lý
ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')

# ==============================================================================
# 2. KHỞI TẠO CÁC KẾT NỐI API
# ==============================================================================
client = genai.Client(api_key=GEMINI_API_KEY)

def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

# Prompt chuẩn hóa trích xuất thông tin khoa học & kỹ thuật
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
    """Tính mã băm MD5 duy nhất cho nội dung file để kiểm tra trùng lặp"""
    return hashlib.md5(file_content_bytes).hexdigest()

def is_matching_tags(file_name, folder_path_str):
    """Kiểm tra file/thư mục có khớp với từ khóa lọc (nếu có cấu hình TARGET_TAGS)"""
    if not TARGET_TAGS:
        return True
    combined_text = f"{file_name} {folder_path_str}".lower()
    for tag in TARGET_TAGS:
        tag_clean = tag.strip().lower()
        if tag_clean and tag_clean in combined_text:
            return True
    return False

# ==============================================================================
# 3. QUÉT THƯ MỤC VÀ XỬ LÝ DỮ LIỆU ĐỆ QUY (TỐI ƯU HÓA)
# ==============================================================================
def scan_and_process_folder(service, folder_id, current_path=""):
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query, 
        fields="files(id, name, mimeType, md5Checksum)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    items = results.get('files', [])

    # Thu thập toàn bộ file JSON hiện có trong thư mục để kiểm tra phiên bản
    json_files = {f['name'] for f in items if f['name'].endswith('.json')}

    for item in items:
        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        # 1. Nếu là Thư mục con -> Quét đệ quy sâu hơn
        if mime_type == 'application/vnd.google-apps.folder':
            scan_and_process_folder(service, file_id, f"{current_path}/{file_name}")

        # 2. Nếu là File tài liệu khoa học
        elif file_name.lower().endswith(ALLOWED_EXTENSIONS):
            if not is_matching_tags(file_name, current_path):
                continue

            base_name = os.path.splitext(file_name)[0]
            expected_json_v1 = f"{base_name}_PROCESSED_V1.json"
            
            # KIỂM TRA CHỐNG CHẠY LẠI: Nếu đã có file JSON kết quả tương ứng -> Bỏ qua
            if expected_json_v1 in json_files:
                print(f"⏩ Bỏ qua (Tài liệu đã được xử lý trước đó): {file_name}")
                continue

            print(f"\n📄 Đang xử lý tài liệu mới: {file_name}")
            print(f"📂 Vị trí: {current_path if current_path else 'Thư mục gốc'}")

            # Tải nội dung file về bộ nhớ tạm để tính MD5 và gửi Gemini
            request = service.files().get_media(fileId=file_id)
            file_stream = io.BytesIO()
            downloader = MediaIoBaseDownload(file_stream, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            
            file_bytes = file_stream.getvalue()
            current_file_hash = calculate_md5(file_bytes)
            print(f"🔑 Mã Hash MD5: {current_file_hash}")

            # Ghi file ra đĩa cứng ảo GitHub để đẩy qua Gemini File API
            local_file_path = f"temp_{file_name}"
            with open(local_file_path, 'wb') as f:
                f.write(file_bytes)

            try:
                # Gửi sang Gemini API (Sử dụng mô hình mới gemini-3.8-flash)
                uploaded_file = client.files.upload(file=local_file_path)
                response = client.models.generate_content(
                    model='gemini-3.8-flash',
                    contents=[uploaded_file, PROMPT_KHOA_HOC],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                client.files.delete(name=uploaded_file.name)
                
                # Bổ sung thông tin truy vết MD5 vào Metadata của JSON
                json_result = json.loads(response.text)
                json_result["metadata"]["file_hash_md5"] = current_file_hash
                json_result["metadata"]["original_filename"] = file_name
                
                # Tải file JSON trực tiếp lên Google Drive (Khắc phục lỗi 403 Storage Quota)
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
                
                print(f"✅ Đã tạo file JSON thành công: {expected_json_v1}")

                # Dọn dẹp đĩa tạm GitHub
                if os.path.exists(local_file_path):
                    os.remove(local_file_path)

            except Exception as e:
                print(f"❌ Lỗi xử lý file {file_name}: {str(e)}")

def main():
    service = get_drive_service()
    if TARGET_TAGS:
        print(f"🔍 Đang chạy chế độ lọc theo từ khóa: {TARGET_TAGS}")
    else:
        print("🌐 Đang chạy chế độ quét toàn bộ thư viện...")
        
    scan_and_process_folder(service, DRIVE_FOLDER_ID)

if __name__ == "__main__":
    main()
