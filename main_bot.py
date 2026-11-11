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
# 1. CẤU HÌNH HỆ THỐNG & BIẾN MÔI TRƯỜNG
# ==============================================================================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")

# Hỗ trợ đa phương thức xác thực Drive (OAuth User Token ưu tiên để tránh lỗi 403 Storage Quota)
OAUTH_TOKEN_JSON = os.environ.get("OAUTH_TOKEN_JSON")
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET")
GOOGLE_REFRESH_TOKEN = os.environ.get("GOOGLE_REFRESH_TOKEN")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")

TARGET_TAGS = [t.strip() for t in os.environ.get("TARGET_TAGS", "").split(",") if t.strip()]

ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')
MODEL_PRIORITY_LIST = ['gemini-3.8-flash', 'gemini-3.8-flash-lite', 'gemini-3.8-pro']

client = genai.Client(api_key=GEMINI_API_KEY)

PROMPT_KHOA_HOC = """
Bạn là một trợ lý phân tích dữ liệu khoa học / kỹ thuật chuyên nghiệp.
Hãy phân tích tài liệu được cung cấp và trích xuất thông tin theo đúng cấu trúc JSON chuẩn sau:

{
  "metadata": {
    "model_name": "Tên Model / Thiết bị / Sản phẩm chính (VD: EP-6000, EP-8000)",
    "document_type": "Loại tài liệu (Brochure / IFU / SpecSheet / DataSheet / Manual / TestReport / Certificate)",
    "language": "Ngôn ngữ tài liệu (VI / EN)",
    "version": "Số phiên bản (VD: V1, V2...) nếu không có thì ghi V1",
    "suggested_standard_name": "Tên đặt chuẩn hóa: [Model]_[DocType]_[Language]_V[Version]",
    "search_tags": ["Danh sách 5-8 từ khóa kỹ thuật chính phục vụ truy vấn nhanh"]
  },
  "extracted_data": {
    "title": "Tiêu đề chính của tài liệu",
    "summary": "Tóm tắt súc tích nội dung khoa học/kỹ thuật cốt lõi (100-250 từ)",
    "technical_specifications": {
      "Nhom_thong_so": {
        "Ten_thong_so": "Gia_tri_kem_don_vi"
      }
    },
    "compatible_accessories_or_models": ["Danh sách model/phụ kiện tương thích nếu có"],
    "key_notes": "Lưu ý đặc biệt, tiêu chuẩn chất lượng, cảnh báo an toàn hoặc quy chuẩn vận hành"
  }
}
YÊU CẦU BẮT BUỘC:
- Trích xuất đầy đủ và chính xác 100% các thông số kỹ thuật có trong tài liệu vào mục technical_specifications.
- Chỉ trả về JSON hợp lệ, tuyệt đối không kèm markdown ```json hay lời giải thích văn bản bên ngoài.
"""


# ==============================================================================
# 2. KẾT NỐI GOOGLE DRIVE (CHỐNG LỖI 403 STORAGE QUOTA)
# ==============================================================================
def get_drive_service():
    scopes = ['[https://www.googleapis.com/auth/drive](https://www.googleapis.com/auth/drive)']

    # Phương án 1: OAuth Token dạng JSON nguyên khối
    if OAUTH_TOKEN_JSON:
        print("🔐 Đang xác thực Google Drive bằng OAUTH_TOKEN_JSON...")
        token_info = json.loads(OAUTH_TOKEN_JSON)
        creds = UserCredentials.from_authorized_user_info(token_info, scopes)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return build('drive', 'v3', credentials=creds)

    # Phương án 2: OAuth Token tách 3 biến (Client ID, Secret, Refresh Token)
    if GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET and GOOGLE_REFRESH_TOKEN:
        print("🔐 Đang xác thực Google Drive bằng OAuth Refresh Token...")
        creds = UserCredentials(
            token=None,
            refresh_token=GOOGLE_REFRESH_TOKEN,
            token_uri="[https://oauth2.googleapis.com/token](https://oauth2.googleapis.com/token)",
            client_id=GOOGLE_CLIENT_ID,
            client_secret=GOOGLE_CLIENT_SECRET,
            scopes=scopes
        )
        creds.refresh(Request())
        return build('drive', 'v3', credentials=creds)

    # Phương án 3: Service Account (Chỉ dùng khi thư mục nằm trên Shared Drive)
    if SERVICE_ACCOUNT_JSON:
        print("🔑 Đang xác thực Google Drive bằng Service Account...")
        sa_info = json.loads(SERVICE_ACCOUNT_JSON)
        creds = ServiceAccountCredentials.from_service_account_info(sa_info, scopes=scopes)
        return build('drive', 'v3', credentials=creds)

    raise ValueError("Thiếu thông tin xác thực Google Drive trong GitHub Secrets.")


def calculate_md5(file_content_bytes):
    return hashlib.md5(file_content_bytes).hexdigest()


def is_matching_tags(file_name, folder_path_str):
    if not TARGET_TAGS:
        return True
    combined_text = f"{file_name} {folder_path_str}".lower()
    return any(tag.lower() in combined_text for tag in TARGET_TAGS)


# ==============================================================================
# 3. TRÍCH XUẤT VĂN BẢN THÔNG MINH & BẢO TOÀN BẢNG THÔNG SỐ KỸ THUẬT
# ==============================================================================
def extract_and_prioritize_text(file_bytes, file_ext, max_chars=350000):
    """Trích xuất text từ PDF/DOCX/XLSX. Nếu quá dài, ưu tiên giữ phần đầu và các trang chứa bảng thông số."""
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
    except Exception as e:
        print(f"   ⚠️ Không thể đọc text cục bộ ({e}), chuyển sang luồng Files API.")
        return ""

    full_text = "\n".join(p[1] for p in pages_text).strip()
    if len(full_text) <= 500:
        return ""  # PDF dạng scan ảnh

    if len(full_text) <= max_chars:
        return full_text

    # Nếu tài liệu quá lớn (như IFU hàng trăm trang): Giữ 20% đầu (Mục lục/Tổng quan) + Các trang chứa từ khóa thông số
    spec_keywords = re.compile(
        r"(specification|thông số|kích thước|dimension|weight|trọng lượng|power|nguồn điện|voltage|"
        r"resolution|độ phân giải|compatible|tương thích|operating|môi trường|rating|model|appendix|phụ lục)",
        re.IGNORECASE
    )
    selected = []
    current_len = 0
    head_limit = int(max_chars * 0.25)

    for page_num, content in pages_text:
        if current_len < head_limit:
            selected.append(content)
            current_len += len(content)
        elif spec_keywords.search(content) and current_len + len(content) <= max_chars:
            selected.append(content)
            current_len += len(content)

    return "\n".join(selected)


# ==============================================================================
# 4. GỌI GEMINI API VỚI RETRY & FALLBACK CHỐNG QUÁ TẢI 429/503
# ==============================================================================
def generate_structured_json(file_bytes, file_name):
    file_ext = os.path.splitext(file_name)[1].lower()
    extracted_text = extract_and_prioritize_text(file_bytes, file_ext)

    uploaded_file = None
    local_temp_path = None

    try:
        if extracted_text:
            print(f"   ⚡ Đã tinh lọc văn bản kỹ thuật ({len(extracted_text):,} ký tự) -> Gửi trực tiếp siêu tốc.")
            contents_payload = [f"TÊN FILE GỐC: {file_name}\n\nNỘI DUNG TÀI LIỆU:\n{extracted_text}", PROMPT_KHOA_HOC]
        else:
            size_mb = len(file_bytes) / (1024 * 1024)
            print(f"   📦 Tài liệu dạng Scan/Hình ảnh ({size_mb:.1f} MB). Tải qua Gemini Files API...")
            local_temp_path = f"temp_{hashlib.md5(file_name.encode()).hexdigest()}{file_ext}"
            with open(local_temp_path, 'wb') as f:
                f.write(file_bytes)
            uploaded_file = client.files.upload(file=local_temp_path)
            contents_payload = [uploaded_file, f"TÊN FILE GỐC: {file_name}", PROMPT_KHOA_HOC]

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
                        print(f"   ⚠️ Model {model_name} báo 404, chuyển sang model dự phòng...")
                        break
                    if ("429" in err_msg or "503" in err_msg or "RESOURCE_EXHAUSTED" in err_msg) and attempt < len(backoff_delays):
                        wait_sec = backoff_delays[attempt]
                        print(f"   ⚠️ Model {model_name} quá tải (lần {attempt + 1}/{len(backoff_delays)}), chờ {wait_sec}s...")
                        time.sleep(wait_sec)
                    else:
                        print(f"   ⚠️ Lỗi trên {model_name}: {err_msg[:180]}")
                        break

        raise RuntimeError("Tất cả các mô hình Gemini thế hệ 3.8 đều quá tải hoặc bị gián đoạn.")

    finally:
        if uploaded_file:
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass
        if local_temp_path and os.path.exists(local_temp_path):
            os.remove(local_temp_path)


# ==============================================================================
# 5. QUÉT ĐỆ QUY & KIỂM SOÁT PHIÊN BẢN MD5
# ==============================================================================
def get_existing_json_info(service, folder_id, base_name, items):
    """Kiểm tra các file [base_name]_PROCESSED_V*.json hiện có để đối chiếu MD5 và xác định version tiếp theo."""
    pattern = re.compile(rf"^{re.escape(base_name)}_PROCESSED_V(\d+)\.json$")
    matched_files = []
    for item in items:
        m = pattern.match(item['name'])
        if m:
            matched_files.append((int(m.group(1)), item['id'], item['name']))

    if not matched_files:
        return None, 1

    matched_files.sort(key=lambda x: x[0], reverse=True)
    latest_ver, latest_id, latest_name = matched_files[0]

    try:
        req = service.files().get_media(fileId=latest_id, supportsAllDrives=True)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        data = json.loads(buf.getvalue().decode('utf-8'))
        saved_md5 = data.get("metadata", {}).get("file_hash_md5")
        return saved_md5, latest_ver + 1
    except Exception:
        return None, latest_ver + 1


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
            saved_md5, next_version = get_existing_json_info(service, folder_id, base_name, items)
            drive_md5 = item.get('md5Checksum')

            # So sánh nhanh MD5 từ metadata của Drive trước khi tải file để tiết kiệm băng thông
            if saved_md5 and drive_md5 and saved_md5 == drive_md5:
                print(f"⏩ Bỏ qua (Trùng khớp MD5 V{next_version - 1}): {current_path}/{file_name}")
                stats["skipped"] += 1
                continue

            print(f"\n📄 Phát hiện tài liệu cần xử lý: {file_name}")
            print(f"   📂 Vị trí: {current_path if current_path else 'Thư mục gốc'}")

            try:
                request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                file_bytes = file_stream.getvalue()
                current_file_hash = calculate_md5(file_bytes)
                print(f"   🔑 Mã Hash MD5: {current_file_hash}")

                if saved_md5 and saved_md5 == current_file_hash:
                    print(f"⏩ Bỏ qua (Trùng khớp MD5 V{next_version - 1}): {file_name}")
                    stats["skipped"] += 1
                    continue

                print("   🚀 Gửi dữ liệu sang Gemini trích xuất (Structured Output)...")
                json_result = generate_structured_json(file_bytes, file_name)

                if "metadata" not in json_result or not isinstance(json_result["metadata"], dict):
                    json_result["metadata"] = {}
                json_result["metadata"]["file_hash_md5"] = current_file_hash
                json_result["metadata"]["original_filename"] = file_name
                json_result["metadata"]["version"] = f"V{next_version}"

                output_json_name = f"{base_name}_PROCESSED_V{next_version}.json"
                json_bytes = io.BytesIO(json.dumps(json_result, ensure_ascii=False, indent=2).encode('utf-8'))
                media = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=False)

                file_metadata = {
                    'name': output_json_name,
                    'parents': [folder_id]
                }

                service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()

                print(f"   ✅ Đã tạo file JSON thành công: {output_json_name}")
                stats["success"] += 1
                time.sleep(4)

            except Exception as e:
                print(f"   ❌ Lỗi khi xử lý file {file_name}: {str(e)}")
                stats["failed"] += 1

    return stats


def main():
    if not GEMINI_API_KEY or not DRIVE_FOLDER_ID:
        print("❌ Lỗi cấu hình: Thiếu GEMINI_API_KEY hoặc DRIVE_FOLDER_ID trong GitHub Secrets.")
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
