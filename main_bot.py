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

import pypdf
import pdfplumber
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import openpyxl

# ============ 1. CẤU HÌNH HỆ THỐNG & BIẾN MÔI TRƯỜNG ============
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
TARGET_TAGS = [t.strip().lower() for t in os.environ.get("TARGET_TAGS", "").split(",") if t.strip()]

GEMINI_MODELS = [m.strip() for m in os.environ.get("GEMINI_MODELS", "gemini-3.8-flash,gemini-2.5-flash").split(",") if m.strip()]
MAX_RETRIES = int(os.environ.get("MAX_RETRIES", "4"))
MAX_CHARS = int(os.environ.get("MAX_CHARS", "300000"))
SLEEP_BETWEEN_FILES = float(os.environ.get("SLEEP_BETWEEN_FILES", "6"))
MAX_INLINE_BYTES = 18 * 1024 * 1024  # Threshold 18MB cho Inline Payload

OUTPUT_SUFFIX = "_PROCESSED_V1.json"
ALLOWED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.png', '.jpg', '.jpeg', '.txt')
FOLDER_MIME = 'application/vnd.google-apps.folder'

EXPORT_MAP = {
    'application/vnd.google-apps.document': ('application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'),
    'application/vnd.google-apps.spreadsheet': ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.xlsx'),
}

DIRECT_MIME = {
    '.pdf': 'application/pdf',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
}

client = genai.Client(api_key=GEMINI_API_KEY)

# JSON Schema tương thích 100% với Gemini Developer API Mode
EXTRACTION_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "metadata": {
            "type": "OBJECT",
            "properties": {
                "model_name": {"type": "STRING", "description": "Tên Model / Thiết bị / Sản phẩm chính"},
                "document_type": {"type": "STRING", "description": "Brochure / IFU / SpecSheet / DataSheet / Manual / TestReport / Certificate / Other"},
                "version": {"type": "STRING", "description": "Số phiên bản (VD: V1, V2...). Nếu không có thì ghi V1"},
                "suggested_standard_name": {"type": "STRING", "description": "Tên đặt chuẩn hóa: [Model]_[DocType]_V[Version]"},
                "search_tags": {
                    "type": "ARRAY",
                    "items": {"type": "STRING"},
                    "description": "3-5 từ khóa chính phục vụ truy vấn nhanh"
                }
            },
            "required": ["model_name", "document_type", "version", "suggested_standard_name", "search_tags"]
        },
        "extracted_data": {
            "type": "OBJECT",
            "properties": {
                "title": {"type": "STRING", "description": "Tiêu đề chính của tài liệu"},
                "summary": {"type": "STRING", "description": "Tóm tắt ngắn gọn nội dung khoa học/kỹ thuật cốt lõi (100-200 từ)"},
                "technical_specifications": {
                    "type": "ARRAY",
                    "description": "Danh sách các thông số kỹ thuật được trích xuất",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "name": {"type": "STRING", "description": "Tên thông số kỹ thuật"},
                            "value": {"type": "STRING", "description": "Giá trị của thông số"},
                            "unit": {"type": "STRING", "nullable": True, "description": "Đơn vị đo (nếu không có ghi null)"}
                        },
                        "required": ["name", "value"]
                    }
                },
                "key_notes": {"type": "STRING", "description": "Lưu ý đặc biệt, kết luận kiểm định hoặc quy chuẩn vận hành"}
            },
            "required": ["title", "summary", "technical_specifications", "key_notes"]
        }
    },
    "required": ["metadata", "extracted_data"]
}

PROMPT_KHOA_HOC = """
Bạn là một trợ lý phân tích dữ liệu khoa học / kỹ thuật chuyên nghiệp.
Hãy phân tích tài liệu được cung cấp và trích xuất thông tin chính xác theo JSON Schema được chỉ định.

QUY TẮC BẮT BUỘC:
- Chỉ sử dụng thông tin có sẵn trong tài liệu. Không suy đoán; thông tin không có thì để null.
- Giữ nguyên toàn bộ số liệu, đơn vị đo, ký hiệu kỹ thuật như bản gốc.
- Trích xuất đầy đủ các bảng thông số kỹ thuật vào "technical_specifications".
"""

# ============ 2. KẾT NỐI VÀ THAO TÁC GOOGLE DRIVE ============
def get_drive_service():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)

def list_children(service, folder_id):
    items, token = [], None
    while True:
        res = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType, md5Checksum, modifiedTime, appProperties)",
            pageSize=1000,
            pageToken=token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        items.extend(res.get('files', []))
        token = res.get('nextPageToken')
        if not token:
            return items

def download_bytes(service, request):
    stream = io.BytesIO()
    downloader = MediaIoBaseDownload(stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return stream.getvalue()

def save_json_to_drive(service, target_folder_id, json_name, data, app_props, existing_id=None):
    """Lưu JSON trực tiếp vào đúng thư mục con target_folder_id (Tránh lỗi 403 Storage Quota)."""
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8')
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype='application/json', resumable=True)
    
    if existing_id:
        service.files().update(
            fileId=existing_id,
            body={'appProperties': app_props},
            media_body=media,
            supportsAllDrives=True
        ).execute()
    else:
        file_metadata = {
            'name': json_name,
            'parents': [target_folder_id],
            'appProperties': app_props
        }
        service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()

# ============ 3. TRÍCH XUẤT VĂN BẢN TRÊN PYTHON ============
def _rows_to_text(rows):
    lines = []
    for row in rows:
        cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
        if cells:
            lines.append(" | ".join(cells))
    return "\n".join(lines)

def extract_pdf_text_fallback(file_bytes):
    parts = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                tables = page.extract_tables() or []
                block = text
                for t in tables:
                    block += "\n[BẢNG]\n" + _rows_to_text(t)
                if block.strip():
                    parts.append(f"--- Trang {i} ---\n{block}")
    except Exception:
        pass
        
    if not parts:
        try:
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            for i, page in enumerate(reader.pages, 1):
                t = page.extract_text()
                if t and t.strip():
                    parts.append(f"--- Trang {i} ---\n{t}")
        except Exception:
            pass
    return "\n".join(parts)

def extract_docx(file_bytes):
    d = docx.Document(io.BytesIO(file_bytes))
    lines = []
    for child in d.element.body.iterchildren():
        if child.tag.endswith('}p'):
            text = Paragraph(child, d).text
            if text.strip():
                lines.append(text)
        elif child.tag.endswith('}tbl'):
            tbl = Table(child, d)
            lines.append("[BẢNG]\n" + _rows_to_text([[c.text for c in r.cells] for r in tbl.rows]))
    return "\n".join(lines)

def extract_xlsx(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    parts = []
    for ws in wb.worksheets:
        body = _rows_to_text(ws.iter_rows(values_only=True))
        if body:
            parts.append(f"--- Sheet: {ws.title} ---\n{body}")
    return "\n".join(parts)

def extract_text_from_bytes(file_bytes, ext):
    try:
        if ext == '.docx':
            return extract_docx(file_bytes).strip()
        elif ext == '.xlsx':
            return extract_xlsx(file_bytes).strip()
        elif ext == '.txt':
            return file_bytes.decode('utf-8', errors='ignore').strip()
        elif ext == '.pdf':
            return extract_pdf_text_fallback(file_bytes).strip()
    except Exception as e:
        print(f"   ⚠️ Lỗi trích xuất văn bản ({ext}): {e}")
    return ""

# ============ 4. KẾT NỐI GEMINI API & SỬA LỖI SIZE/QUOTA ============
def call_gemini_with_fallback(file_bytes, ext, file_name):
    """
    Xử lý thông minh: 
    - File <= 18MB: Dùng types.Part.from_bytes (Inline Byte Payload)
    - File > 18MB: Dùng client.files.upload (Files API) và tự dọn dẹp file tạm
    """
    uploaded_file_ref = None
    temp_file_path = None

    try:
        if ext in DIRECT_MIME:
            if len(file_bytes) <= MAX_INLINE_BYTES:
                contents = [
                    types.Part.from_bytes(data=file_bytes, mime_type=DIRECT_MIME[ext]),
                    PROMPT_KHOA_HOC
                ]
            else:
                print(f"   📦 File lớn ({len(file_bytes) / (1024*1024):.1f}MB > 18MB). Chuyển sang luồng Files API...")
                temp_file_path = f"temp_{int(time.time())}_{file_name}"
                with open(temp_file_path, "wb") as f:
                    f.write(file_bytes)
                uploaded_file_ref = client.files.upload(file=temp_file_path)
                contents = [uploaded_file_ref, PROMPT_KHOA_HOC]
        else:
            text = extract_text_from_bytes(file_bytes, ext)
            if not text:
                return None, None
            contents = f"{PROMPT_KHOA_HOC}\n\nNỘI DUNG TÀI LIỆU TRÍCH XUẤT:\n{text[:MAX_CHARS]}"

        for model in GEMINI_MODELS:
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    resp = client.models.generate_content(
                        model=model,
                        contents=contents,
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            response_schema=EXTRACTION_SCHEMA,
                            temperature=0.1,
                        ),
                    )
                    data = json.loads(resp.text)
                    if not isinstance(data, dict):
                        raise ValueError("Dữ liệu JSON không đạt chuẩn dict/object")
                    return data, model

                except errors.APIError as e:
                    code = getattr(e, "code", None)
                    if code in (429, 500, 502, 503, 504):
                        wait = 15 * (2 ** (attempt - 1)) # Tăng thời gian chờ lũy thừa (15s, 30s, 60s, 120s)
                        print(f"   ⚠️ Model {model} gặp lỗi {code} (lần {attempt}/{MAX_RETRIES}), chờ {wait}s...")
                        time.sleep(wait)
                    else:
                        print(f"   ⚠️ Model {model} báo lỗi {code}: {e}. Chuyển sang model dự phòng.")
                        break
                except (ValueError, TypeError) as e:
                    print(f"   ⚠️ Model {model} trả JSON sai cấu trúc ({e}), thử lại sau 3s...")
                    time.sleep(3)

        raise RuntimeError("Tất cả các mô hình Gemini trong cấu hình đều quá tải hoặc bị gián đoạn.")

    finally:
        # Dọn dẹp tài nguyên tạm
        if uploaded_file_ref:
            try:
                client.files.delete(name=uploaded_file_ref.name)
            except Exception:
                pass
        if temp_file_path and os.path.exists(temp_file_path):
            os.remove(temp_file_path)

def normalize_json_structure(data):
    if not isinstance(data, dict):
        data = {}
    data.setdefault("metadata", {})
    data.setdefault("extracted_data", {})
    if not isinstance(data["metadata"], dict):
        data["metadata"] = {}
    if not isinstance(data["extracted_data"], dict):
        data["extracted_data"] = {}
        
    # Chuyển đổi mảng technical_specifications thành Dict Key-Value nhất quán
    specs_raw = data["extracted_data"].get("technical_specifications")
    if isinstance(specs_raw, list):
        specs_dict = {}
        for item in specs_raw:
            if isinstance(item, dict) and "name" in item:
                spec_name = item["name"]
                specs_dict[spec_name] = {
                    "value": item.get("value"),
                    "unit": item.get("unit")
                }
        data["extracted_data"]["technical_specifications"] = specs_dict

    return data

def matches_tags(name, path):
    if not TARGET_TAGS:
        return True
    combined = f"{name} {path}".lower()
    return any(t in combined for t in TARGET_TAGS)

# ============ 5. QUÉT VÀ TRÍCH XUẤT ĐỆ QUY THƯ MỤC ============
def scan_and_process_folder(service, folder_id, stats, current_path=""):
    items = list_children(service, folder_id)

    done_by_source = {}
    for f in items:
        if f['name'].endswith(OUTPUT_SUFFIX):
            props = f.get('appProperties') or {}
            if props.get('source_id'):
                done_by_source[props['source_id']] = (f['id'], props.get('source_sig'))

    for item in items:
        fid = item['id']
        name = item['name']
        mime = item['mimeType']

        if mime == FOLDER_MIME:
            scan_and_process_folder(service, fid, stats, f"{current_path}/{name}")
            continue

        export_info = EXPORT_MAP.get(mime)
        if export_info:
            export_mime, ext = export_info
            clean_ext_name = ext.replace('.', '')
            base_name = f"{name}_{clean_ext_name}"
        else:
            if not name.lower().endswith(ALLOWED_EXTENSIONS):
                continue
            ext = os.path.splitext(name)[1].lower()
            clean_ext_name = ext.replace('.', '')
            base_name = f"{os.path.splitext(name)[0]}_{clean_ext_name}"
            export_mime = None

        if not matches_tags(name, current_path):
            continue

        json_name = f"{base_name}{OUTPUT_SUFFIX}"
        signature = item.get('md5Checksum') or item.get('modifiedTime')

        existing_json_id = None
        if fid in done_by_source:
            existing_json_id, old_sig = done_by_source[fid]
            if old_sig and old_sig == signature:
                stats['skipped'] += 1
                print(f"⏩ Bỏ qua (Dữ liệu không đổi): {name}")
                continue
            print(f"\n🔄 File nguồn đã cập nhật, tiến hành trích xuất lại: {name}")
        else:
            print(f"\n📄 Phát hiện tài liệu mới: {name}")

        print(f"   📂 Vị trí: {current_path or 'Thư mục gốc'}")

        try:
            if export_mime:
                req = service.files().export_media(fileId=fid, mimeType=export_mime)
            else:
                req = service.files().get_media(fileId=fid, supportsAllDrives=True)
                
            file_bytes = download_bytes(service, req)
            file_hash = hashlib.md5(file_bytes).hexdigest()

            print("   🚀 Gửi dữ liệu sang Gemini trích xuất (Structured Output)...")
            result_data, used_model = call_gemini_with_fallback(file_bytes, ext, name)
            
            if not result_data:
                print("   ⚠️ Không trích xuất được nội dung từ file. Bỏ qua.")
                stats['skipped'] += 1
                continue

            result_data = normalize_json_structure(result_data)

            # Cập nhật siêu dữ liệu truy vết
            result_data["metadata"].update({
                "file_hash_md5": file_hash,
                "original_filename": name,
                "source_drive_id": fid,
                "source_path": current_path or "/",
                "processed_at": datetime.now(timezone.utc).isoformat(),
                "processed_by_model": used_model,
            })

            # Ghi JSON trực tiếp vào đúng thư mục con hiện tại
            save_json_to_drive(
                service=service,
                target_folder_id=folder_id,
                json_name=json_name,
                data=result_data,
                app_props={"source_id": fid, "source_sig": signature or ""},
                existing_id=existing_json_id
            )

            stats['ok'] += 1
            print(f"   ✅ Đã tạo JSON thành công: {json_name} (Mô hình: {used_model})")
            time.sleep(SLEEP_BETWEEN_FILES)

        except Exception as e:
            stats['failed'] += 1
            print(f"   ❌ Lỗi khi xử lý file {name}: {e}")

def main():
    service = get_drive_service()
    print(f"🔍 Từ khóa lọc (TARGET_TAGS): {TARGET_TAGS}" if TARGET_TAGS else "🌐 Quét toàn bộ thư viện THU_VIEN_SAN_PHAM...")
    print(f"🤖 Danh sách Mô hình Ưu tiên: {GEMINI_MODELS}")

    stats = {'ok': 0, 'skipped': 0, 'failed': 0}
    scan_and_process_folder(service, DRIVE_FOLDER_ID, stats)

    print(f"\n📊 BÁO CÁO TỔNG KẾT VẬN HÀNH:")
    print(f"   ✔️ Thành công: {stats['ok']}")
    print(f"   ⏩ Bỏ qua:     {stats['skipped']}")
    print(f"   ❌ Thất bại:   {stats['failed']}")

    if stats['failed'] > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
