import os
import sys
import json
import io
import time
import hashlib
import pymupdf  # PyMuPDF: trích xuất toàn bộ văn bản & nén ảnh trang PDF scan
import docx
import openpyxl
from PIL import Image
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google.oauth2.credentials import Credentials as UserCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ==============================================================================
# 1. BIẾN MÔI TRƯỜNG & CẤU HÌNH TÁCH BIỆT 2 TÀI KHOẢN (DRIVE vs GEMINI)
# ==============================================================================
# Tài khoản B (Tài khoản Gemini đề xuất để xử lý AI - hỗ trợ xoay vòng nhiều key)
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_API_KEYS = os.environ.get("GEMINI_API_KEYS", "").strip()

# Tài khoản A (Tài khoản Google Drive chứa thư mục THU_VIEN_SAN_PHAM)
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "").strip()
OAUTH_TOKEN_JSON = os.environ.get("OAUTH_TOKEN_JSON", "").strip()
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "").strip()

TARGET_TAGS = [t.strip() for t in os.environ.get("TARGET_TAGS", "").split(",") if t.strip()]
ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg')
GEMINI_MODEL = 'gemini-3.8-flash'


def build_api_key_pool():
    keys = []
    if GEMINI_API_KEYS:
        keys.extend([k.strip() for k in GEMINI_API_KEYS.split(",") if k.strip()])
    if GEMINI_API_KEY and GEMINI_API_KEY not in keys:
        keys.insert(0, GEMINI_API_KEY)
    return keys


API_KEY_POOL = build_api_key_pool()
CURRENT_KEY_IDX = 0

# ==============================================================================
# 2. KHỞI TẠO CLIENTS ĐỘC LẬP & PROMPT CHUẨN HÓA
# ==============================================================================
def get_gemini_client():
    if not API_KEY_POOL:
        raise ValueError("Thiếu GEMINI_API_KEY của tài khoản Gemini đề xuất trong GitHub Secrets.")
    active_key = API_KEY_POOL[CURRENT_KEY_IDX]
    masked_key = f"{active_key[:8]}...{active_key[-4:]}"
    pool_info = f" | Key {CURRENT_KEY_IDX + 1}/{len(API_KEY_POOL)}" if len(API_KEY_POOL) > 1 else ""
    print(f"🧠 Khởi tạo Gemini Client bằng tài khoản đề xuất (API Key: {masked_key}{pool_info}) | Model: {GEMINI_MODEL}")
    return genai.Client(api_key=active_key)


def get_drive_service():
    scopes = ['https://www.googleapis.com/auth/drive']

    if OAUTH_TOKEN_JSON:
        print("🔐 Đang kết nối Google Drive (Kho dữ liệu) bằng OAUTH_TOKEN_JSON...")
        token_info = json.loads(OAUTH_TOKEN_JSON)
        creds = UserCredentials.from_authorized_user_info(token_info, scopes)
        creds.refresh(Request())
        print("✅ Xác thực tài khoản Google Drive thành công!")
        return build('drive', 'v3', credentials=creds)

    if SERVICE_ACCOUNT_JSON:
        print("🔑 Đang kết nối Google Drive bằng SERVICE_ACCOUNT_JSON...")
        sa_info = json.loads(SERVICE_ACCOUNT_JSON)
        sa_info["token_uri"] = "https://oauth2.googleapis.com/token"
        creds = ServiceAccountCredentials.from_service_account_info(sa_info, scopes=scopes)
        creds.refresh(Request())
        print("✅ Xác thực Service Account thành công!")
        return build('drive', 'v3', credentials=creds)

    raise ValueError("Thiếu thông tin xác thực Google Drive (OAUTH_TOKEN_JSON hoặc SERVICE_ACCOUNT_JSON).")


PROMPT_KHOA_HOC = """
Bạn là một trợ lý phân tích dữ liệu khoa học / kỹ thuật chuyên nghiệp.
Hãy phân tích TOÀN BỘ tài liệu được cung cấp và trích xuất đầy đủ thông tin theo cấu trúc JSON chuẩn sau:

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
# 3. TRÍCH XUẤT 100% DỮ LIỆU TÀI LIỆU
# ==============================================================================
def build_gemini_payload(file_bytes, file_name):
    file_ext = os.path.splitext(file_name)[1].lower()

    if file_ext == '.docx':
        doc = docx.Document(io.BytesIO(file_bytes))
        lines = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                lines.append(" | ".join(c.text.strip().replace('\n', ' ') for c in row.cells))
        full_text = "\n".join(lines)
        print(f"⚡ Đã trích xuất 100% văn bản Word ({len(full_text):,} ký tự).")
        return [f"TÊN FILE: {file_name}\n\n{full_text}", PROMPT_KHOA_HOC]

    if file_ext in ('.xlsx', '.xls'):
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        lines = []
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            lines.append(f"=== SHEET: {sheet} ===")
            for row in ws.iter_rows(values_only=True):
                if any(row):
                    lines.append(" | ".join(str(v).strip() if v is not None else "" for v in row))
        full_text = "\n".join(lines)
        print(f"⚡ Đã trích xuất 100% dữ liệu Excel ({len(full_text):,} ký tự, {len(wb.sheetnames)} sheet).")
        return [f"TÊN FILE: {file_name}\n\n{full_text}", PROMPT_KHOA_HOC]

    if file_ext in ('.png', '.jpg', '.jpeg'):
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        img.thumbnail((1600, 1600))
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=80)
        print(f"🖼️ Đã tối ưu kích thước ảnh ({len(buf.getvalue()) // 1024} KB).")
        return [
            types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"),
            f"TÊN FILE: {file_name}",
            PROMPT_KHOA_HOC
        ]

    if file_ext == '.pdf':
        pdf_doc = pymupdf.open(stream=file_bytes, filetype="pdf")
        total_pages = len(pdf_doc)
        pages_text = []
        for idx, page in enumerate(pdf_doc):
            txt = page.get_text("text") or ""
            if txt.strip():
                pages_text.append(f"--- TRANG {idx + 1} ---\n{txt}")

        full_text = "\n".join(pages_text).strip()

        if len(full_text) > 400:
            print(f"⚡ Đã trích xuất 100% văn bản PDF ({len(full_text):,} ký tự / {total_pages} trang).")
            return [f"TÊN FILE: {file_name}\n\n{full_text}", PROMPT_KHOA_HOC]

        print(f"🖼️ PDF dạng hình ảnh/scan ({total_pages} trang) -> Nén toàn bộ 100% số trang...")
        parts = [f"TÊN FILE: {file_name} (Đầy đủ {total_pages}/{total_pages} trang)"]
        total_img_kb = 0
        for p_idx in range(total_pages):
            pix = pdf_doc[p_idx].get_pixmap(dpi=120)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail((1400, 1400))
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=75)
            img_bytes = buf.getvalue()
            total_img_kb += len(img_bytes) // 1024
            parts.append(types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg"))

        parts.append(PROMPT_KHOA_HOC)
        print(f"📦 Đã đóng gói toàn bộ {total_pages} trang ảnh (Tổng: {total_img_kb} KB).")
        return parts

    raise ValueError(f"Định dạng file không được hỗ trợ: {file_ext}")


def generate_structured_json(client_holder, file_bytes, file_name):
    global CURRENT_KEY_IDX
    contents_payload = build_gemini_payload(file_bytes, file_name)
    backoff_delays = [20, 45, 70]

    while CURRENT_KEY_IDX < len(API_KEY_POOL):
        for attempt in range(len(backoff_delays) + 1):
            try:
                response = client_holder["client"].models.generate_content(
                    model=GEMINI_MODEL,
                    contents=contents_payload,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.1
                    )
                )
                return json.loads(response.text)
            except Exception as e:
                err_msg = str(e)
                is_quota_429 = ("429" in err_msg) or ("RESOURCE_EXHAUSTED" in err_msg)
                is_daily_limit = ("PerDay" in err_msg) or ("per day" in err_msg.lower())

                if is_quota_429 and is_daily_limit:
                    print(f"🛑 API Key #{CURRENT_KEY_IDX + 1} đã hết hạn mức NGÀY (Free Tier Daily Limit).")
                    if CURRENT_KEY_IDX + 1 < len(API_KEY_POOL):
                        CURRENT_KEY_IDX += 1
                        print(f"🔄 Đang chuyển sang API Key dự phòng #{CURRENT_KEY_IDX + 1}...")
                        client_holder["client"] = get_gemini_client()
                        break
                    else:
                        raise RuntimeError(
                            f"Hết hạn mức ngày của {GEMINI_MODEL} (Free Tier giới hạn 20 request/ngày). "
                            f"Chi tiết: {err_msg}"
                        )

                if (is_quota_429 or "503" in err_msg) and attempt < len(backoff_delays):
                    wait_sec = backoff_delays[attempt]
                    print(f"⚠️ Model {GEMINI_MODEL} chạm giới hạn tốc độ (lần {attempt + 1}/{len(backoff_delays)}), chờ {wait_sec}s...")
                    time.sleep(wait_sec)
                else:
                    raise RuntimeError(f"Lỗi gọi {GEMINI_MODEL}: {err_msg}")


# ==============================================================================
# 4. QUÉT ĐỆ QUY VÀ TRÍCH XUẤT DỮ LIỆU
# ==============================================================================
def scan_and_process_folder(drive_service, client_holder, folder_id, current_path="", stats=None):
    if stats is None:
        stats = {"success": 0, "skipped": 0, "failed": 0, "quota_exhausted": False}

    if stats.get("quota_exhausted"):
        return stats

    page_token = None
    items = []
    while True:
        query = f"'{folder_id}' in parents and trashed = false"
        results = drive_service.files().list(
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
        if stats.get("quota_exhausted"):
            break

        file_id = item['id']
        file_name = item['name']
        mime_type = item['mimeType']

        if mime_type == 'application/vnd.google-apps.folder':
            scan_and_process_folder(drive_service, client_holder, file_id, f"{current_path}/{file_name}", stats)

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
                request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
                file_stream = io.BytesIO()
                downloader = MediaIoBaseDownload(file_stream, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

                file_bytes = file_stream.getvalue()
                current_file_hash = calculate_md5(file_bytes)
                print(f"🔑 Mã Hash MD5: {current_file_hash}")

                json_result = generate_structured_json(client_holder, file_bytes, file_name)
                if "metadata" not in json_result or not isinstance(json_result["metadata"], dict):
                    json_result["metadata"] = {}
                json_result["metadata"]["file_hash_md5"] = current_file_hash
                json_result["metadata"]["original_filename"] = file_name

                json_bytes = io.BytesIO(json.dumps(json_result, ensure_ascii=False, indent=2).encode('utf-8'))
                media = MediaIoBaseUpload(json_bytes, mimetype='application/json', resumable=False)

                file_metadata = {
                    'name': expected_json_v1,
                    'parents': [folder_id]
                }

                drive_service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id',
                    supportsAllDrives=True
                ).execute()

                print(f"✅ Đã tạo file JSON thành công: {expected_json_v1}")
                stats["success"] += 1
                time.sleep(10)

            except RuntimeError as e:
                err_text = str(e)
                print(f"❌ Lỗi xử lý file {file_name}: {err_text}")
                stats["failed"] += 1
                if "Hết hạn mức ngày" in err_text:
                    stats["quota_exhausted"] = True
                    break
            except Exception as e:
                print(f"❌ Lỗi xử lý file {file_name}: {str(e)}")
                stats["failed"] += 1

    return stats


def main():
    if not DRIVE_FOLDER_ID:
        print("❌ Lỗi cấu hình: Thiếu DRIVE_FOLDER_ID trong GitHub Secrets.")
        sys.exit(1)

    client_holder = {"client": get_gemini_client()}
    drive_service = get_drive_service()

    if TARGET_TAGS:
        print(f"🔍 Đang chạy chế độ lọc theo từ khóa: {TARGET_TAGS}")
    else:
        print("🌐 Đang chạy chế độ quét toàn bộ thư viện THU_VIEN_SAN_PHAM...")

    stats = scan_and_process_folder(drive_service, client_holder, DRIVE_FOLDER_ID)

    print("\n📊 BÁO CÁO TỔNG KẾT VẬN HÀNH:")
    print(f"   ✔️ Thành công: {stats['success']}")
    print(f"   ⏩ Bỏ qua:     {stats['skipped']}")
    print(f"   ❌ Thất bại:   {stats['failed']}")

    if stats["failed"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
