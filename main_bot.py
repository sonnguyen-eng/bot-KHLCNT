import os
import sys
import re
import json
import io
import time
import hashlib
import pypdf
import docx
import openpyxl
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# 1. BIẾN MÔI TRƯỜNG & SECRETS
# ==============================================================================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "").strip()
OAUTH_TOKEN_JSON = os.environ.get("OAUTH_TOKEN_JSON", "").strip()
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "").strip()
TARGET_TAGS = [t.strip() for t in os.environ.get("TARGET_TAGS", "").split(",") if t.strip()]

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')
MODEL_PRIORITY_LIST = ['gemini-3.8-flash', 'gemini-3.8-flash-lite', 'gemini-3.8-pro']

# ==============================================================================
# 2. KHỞI TẠO CLIENTS & PROMPT CHUẨN HÓA
# ==============================================================================
client = genai.Client(api_key=GEMINI_API_KEY)


def get_drive_service():
    scopes = ['https://www.googleapis.com/auth/drive']

    if OAUTH_TOKEN_JSON:
        print("🔐 Đang xác thực Google Drive bằng OAUTH_TOKEN_JSON...")
        token_info = json.loads(OAUTH_TOKEN_JSON)
        creds = UserCredentials.from_authorized_user_info(token_info, scopes)
        creds.refresh(Request())
        print("✅ Xác thực OAuth User Token thành công!")
        return build('drive', 'v3', credentials=creds)

    if SERVICE_ACCOUNT_JSON:
        print("🔑 Đang xác thực Google Drive bằng SERVICE_ACCOUNT_JSON...")
        sa_info = json.loads(SERVICE_ACCOUNT_JSON)
        sa_info["token_uri"] = "https://oauth2.googleapis.com/token"
        creds = ServiceAccountCredentials.from_service_account_info(sa_info, scopes=scopes)
        creds.refresh(Request())
        print("✅ Xác thực Service Account thành công!")
        return build('drive', 'v3', credentials=creds)

    raise ValueError("Thiếu thông tin xác thực Google Drive (OAUTH_TOKEN_JSON hoặc SERVICE_ACCOUNT_JSON).")


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


# ==============================================================================
# 3. TRÍCH XUẤT VĂN BẢN THÔNG MINH (CHỐNG LỖI 429 TRÀN TOKEN VỚI FILE PDF LỚN)
# ==============================================================================
def extract_and_prioritize_text(file_bytes, file_ext, max_chars=350000):
    pages_text = []
    try:
        if file_ext == '.pdf':
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for idx, page in enumerate(reader.pages):
                txt = page.extract_text() or ""
                if txt.strip():
                    pages_text.append((idx + 1, f"--- TRANG {idx + 1} ---\n{txt}"))
        elif file_ext == '.docx':
            doc = docx.Document(io.BytesIO(file_bytes))
            block = []
            for para in doc.paragraphs:
                if para.text.strip():
                    block.append(para.text)
            for table in doc.tables:
                for row in table.rows:
                    block.append(" | ".join(c.text.strip().replace('\n', ' ') for c in row.cells))
            if block:
                pages_text.append((1, "\n".join(block)))
        elif file_ext in ('.xlsx', '.xls'):
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            for idx, sheet in enumerate(wb.sheetnames):
                ws = wb[sheet]
                rows = [f"=== SHEET: {sheet} ==="]
                for row in ws.iter_rows(values_only=True):
                    if any(row):
                        rows.append(" | ".join(str(v).strip() if v is not None else "" for v in row))
                pages_text.append((idx + 1, "\n".join(rows)))
    except Exception:
        return ""

    full_text = "\n".join(p[1] for p in pages_text).strip()
    if len(full_text) <= 500:
        return ""

    if len(full_text) <= max_chars:
        return full_text

    spec_keywords = re.compile(
        r"(specification|thông số|kích thước|dimension|weight|trọng lượng|power|nguồn điện|voltage|"
        r"resolution|độ phân giải|compatible|tương thích|operating|môi trường|rating|model|appendix|phụ lục)",
        re.IGNORECASE
    )
    selected = []
    current_len = 0
    head_limit = int(max_chars * 0.25)

    for _, content in pages_text:
        if current_len < head_limit:
            selected.append(content)
            current_len += len(content)
        elif spec_keywords.search(content) and current_len + len(content) <= max_chars:
            selected.append(content)
            current_len += len(content)

    return "\n".join(selected)


def generate_structured_json(file_bytes, file_name):
    file_ext = os.path.splitext(file_name)[1].lower()
    extracted_text = extract_and_prioritize_text(file_bytes, file_ext)

    uploaded_file = None
    local_file_path = None

    try:
        if extracted_text:
            print(f"⚡ Đã trích xuất văn bản sạch ({len(extracted_text):,} ký tự) -> Gửi trực tiếp siêu tốc.")
            contents_payload = [f"TÊN FILE GỐC: {file_name}\n\nNỘI DUNG TÀI LIỆU:\n{extracted_text}", PROMPT_KHOA_HOC]
        else:
            print("📦 Tài liệu dạng Scan/Hình ảnh -> Tải qua Gemini Files API...")
            local_file_path = f"temp_{hashlib.md5(file_name.encode()).hexdigest()}{file_ext}"
            with open(local_file_path, 'wb') as f:
                f.write(file_bytes)
            uploaded_file = client.files.upload(file=local_file_path)
            contents_payload = [uploaded_file, PROMPT_KHOA_HOC]

        for model_name in MODEL_PRIORITY_LIST:
            backoff_delays = [15, 30, 60]
            for attempt in range(len(backoff_delays) + 1):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=contents_payload,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.1
                        )
                    )
                    return json.loads(response.text)
                except Exception as e:
                    err_msg = str(e)
                    if "404" in err_msg or "NOT_FOUND" in err_msg:
                        print(f"⚠️ Model {model_name} báo 404, chuyển sang model dự phòng...")
                        break
                    if ("429" in err_msg or "503" in err_msg or "RESOURCE_EXHAUSTED" in err_msg) and attempt < len(backoff_delays):
                        wait_sec = backoff_delays[attempt]
                        print(f"⚠️ Model {model_name} quá tải (lần {attempt + 1}/{len(backoff_delays)}), chờ {wait_sec}s...")
                        time.sleep(wait_sec)
                    else:
                        print(f"⚠️ Lỗi trên {model_name}: {err_msg[:180]}")
                        break

        raise RuntimeError("Tất cả các mô hình Gemini thế hệ 3.8 đều quá tải hoặc bị gián đoạn.")

    finally:
        if uploaded_file:
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass
        if local_file_path and os.path.exists(local_file_path):
            os.remove(local_file_path)


# ==============================================================================
# 4. QUÉT ĐỆ QUY VÀ TRÍCH XUẤT DỮ LIỆU
# ==============================================================================
def scan_and_process_folder(service, folder_id, current_path="", stats=None):
    if stats is None:
        stats = {"success": 0, "skipped": 0, "failed": 0}

    page_token = None
    items = []
    while True:
        query = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, md5Checksum)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            pageToken=page_token
        ).execute()
        items.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break

    json_files = {f['name'] for f in items if f['name'].endswith('.json')}

    for item in items:
        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        if mime_type == 'application/vnd.google-apps.folder':
            scan_and_process_folder(service, file_id, f"{current_path}/{file_name}", stats)

        elif file_name.lower().endswith(ALLOWED_EXTENSIONS):
            if not is_matching_tags(file_name, current_path):
                continue

            base_name = os.path.splitext(file_name)[0]
            expected_json_v1 = f"{base_name}_PROCESSED_V1.json"

            if expected_json_v1 in json_files:
                print(f"⏩ Bỏ qua (Đã xử lý): {file_name}")
                stats["skipped"] += 1
                continue

            print(f"\n📄 Đang xử lý tài liệu mới: {file_name}")
            print(f"📂 Vị trí: {current_path if current_path else 'Thư mục gốc'}")

            try:
                request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                file_bytes = file_stream.getvalue()
                current_file_hash = calculate_md5(file_bytes)
                print(f"🔑 Mã Hash MD5: {current_file_hash}")

                json_result = generate_structured_json(file_bytes, file_name)
                if "metadata" not in json_result or not isinstance(json_result["metadata"], dict):
                    json_result["metadata"] = {}
                json_result["metadata"]["file_hash_md5"] = current_file_hash
                json_result["metadata"]["original_filename"] = file_name

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
                stats["success"] += 1
                time.sleep(4)

            except Exception as e:
                print(f"❌ Lỗi xử lý file {file_name}: {str(e)}")
                stats["failed"] += 1

    return stats


def main():
    if not GEMINI_API_KEY or not DRIVE_FOLDER_ID:
        print("❌ Lỗi cấu hình: Thiếu GEMINI_API_KEY hoặc DRIVE_FOLDER_ID.")
        sys.exit(1)

    service = get_drive_service()
    print(f"🤖 Danh sách Mô hình Ưu tiên: {MODEL_PRIORITY_LIST}")
    if TARGET_TAGS:
        print(f"🔍 Đang chạy chế độ lọc theo từ khóa: {TARGET_TAGS}")
    else:
        print("🌐 Đang chạy chế độ quét toàn bộ thư viện THU_VIEN_SAN_PHAM...")

    stats = scan_and_process_folder(service, DRIVE_FOLDER_ID)

    print("\n📊 BÁO CÁO TỔNG KẾT VẬN HÀNH:")
    print(f"   ✔️ Thành công: {stats['success']}")
    print(f"   ⏩ Bỏ qua:     {stats['skipped']}")
    print(f"   ❌ Thất bại:   {stats['failed']}")

    if stats["failed"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
