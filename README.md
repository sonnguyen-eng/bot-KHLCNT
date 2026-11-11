# Technical Data Warehouse Bot — B

BOT xử lý tài liệu từ Google Drive bằng Gemini, trích xuất dữ liệu kỹ thuật có cấu trúc và duy trì kho dữ liệu SQLite trên Google Drive.

## Luồng xử lý

Google Drive → nhận diện file → kiểm tra chữ ký → tải file thay đổi → trích xuất nội dung → Gemini Structured Output → JSON kết quả → cập nhật SQLite Warehouse → phát hiện file trùng và thông số có khả năng mâu thuẫn.

## Chức năng trong bản đề xuất

1. Quét thư mục Google Drive đệ quy, hỗ trợ Shared Drive.
2. Hỗ trợ PDF, PNG/JPG, DOCX, XLSX, TXT và export Google Docs/Sheets/Slides.
3. Dùng `md5Checksum` của Drive cho file nhị phân; dùng `modifiedTime`/metadata cho Google Workspace.
4. Lưu liên kết nguồn bằng `source_id`, không dựa vào tên file.
5. Trích xuất manufacturer, model, document type/version, tags, summary và danh sách thông số.
6. Tạo JSON kết quả cạnh tài liệu nguồn.
7. Duy trì SQLite Warehouse trên Drive: documents, models, specifications, duplicate_groups, conflicts, processing_runs.
8. Phát hiện file trùng chính xác bằng SHA-256.
9. Gắn cờ `POTENTIAL_CONFLICT` khi cùng model + thông số + đơn vị có giá trị văn bản khác nhau.
10. Không tự quy đổi đơn vị và không tự kết luận hai model khác tên là cùng một model.

## Cấu hình GitHub

Tạo các Secrets:
- `GEMINI_API_KEY`: API key Gemini.
- `DRIVE_FOLDER_ID`: ID thư mục gốc cần quét.
- `SERVICE_ACCOUNT_JSON`: nội dung JSON của Google service account.
- `TARGET_TAGS` (không bắt buộc): từ khóa phân tách bằng dấu phẩy; để trống để quét mọi tài liệu hỗ trợ.

Tạo Repository Variable:
- `GEMINI_MODELS`: danh sách model Gemini phân tách bằng dấu phẩy, ví dụ `gemini-3.8-flash`.

Tùy chọn workflow environment:
- `FORCE_REPROCESS=true`: xử lý lại toàn bộ file đã có trong warehouse (dùng sau khi thay đổi prompt/schema lớn). Mặc định `false`.

## Quyền Google Cloud / Drive

- Bật Google Drive API trong Google Cloud Project.
- Chia sẻ thư mục nguồn với email service account.
- Với Shared Drive, cấp membership/role phù hợp.
- Không đưa service account JSON hoặc API key vào Git repository.

## Cách triển khai

1. Sao lưu repository và các JSON kết quả hiện tại.
2. Thay mã nguồn bằng các file trong gói này.
3. Cấu hình Secrets/Variables như trên.
4. Chạy `workflow_dispatch` thủ công lần đầu.
5. Kiểm tra Actions logs, JSON sinh ra và file `_technical_warehouse.sqlite3` trong thư mục Drive gốc.
6. Chạy thử trên một thư mục nhỏ trước khi quét toàn bộ kho.

## Chuẩn hóa model bằng alias

`model_aliases.json` là object JSON dạng `{"manufacturer::model alias": "manufacturer::canonical model"}`.
Hai vế phải được viết theo dạng đã chuẩn hóa: chữ thường, bỏ dấu, dấu câu chuyển thành khoảng trắng, các khoảng trắng liên tiếp gộp lại.
Chỉ thêm alias sau khi người quản trị xác nhận hai tên thực sự cùng một model. Không bật fuzzy matching tự động.
Ví dụ cấu trúc:
```json
{
  "acme::pump x 100": "acme::pump x100"
}
```
Sau khi sửa file, chạy workflow; mapping được lưu lại trong bảng `model_aliases` của SQLite Warehouse.

## Cách hiểu kết quả

### Duplicate
`duplicate_groups` nhóm các file có cùng SHA-256. Đây là trùng nội dung byte chính xác; không phát hiện bản scan khác độ phân giải hoặc tài liệu cùng nội dung nhưng khác định dạng.

### Potential conflict
`conflicts` là cảnh báo bảo thủ: cùng model key, cùng tên thông số đã chuẩn hóa, cùng đơn vị nhưng có giá trị khác nhau. Đây không phải kết luận tài liệu sai. Có thể do phiên bản, điều kiện đo, cấu hình hoặc phạm vi áp dụng khác nhau; cần người dùng xác minh.

### Model identity
`model_key` được tạo từ manufacturer + model name sau chuẩn hóa ký tự. Không dùng fuzzy matching tự động vì có nguy cơ gộp nhầm model gần giống nhau. Cần bổ sung bảng alias/định danh thủ công khi có dữ liệu thực tế.

## Giới hạn / việc cần kiểm thử trước production

- SQLite được tải xuống và tải ngược lên Google Drive mỗi lần chạy. `concurrency` ngăn hai workflow trong cùng repository chạy song song, nhưng không khóa được một tiến trình khác cập nhật cùng file DB.
- Bản hiện tại xử lý tuần tự để giảm áp lực quota Gemini.
- PDF/ảnh lớn hơn 18 MB bị từ chối trong luồng inline; cần bổ sung Gemini Files API cho lô tài liệu lớn. PDF trên 50 MB bị từ chối theo ngưỡng bảo thủ.
- File media lớn hơn 18 MB bị từ chối ở luồng inline hiện tại; cần triển khai Gemini Files API trước khi đưa file lớn vào lô.
- File text dài hơn `MAX_TEXT_CHARS` bị cắt; nên bổ sung chia chunk và hợp nhất kết quả nếu tài liệu dài.
- XLSX đọc công thức (`data_only=False`); giá trị tính toán cache không được bảo đảm.
- So sánh xung đột chưa chuẩn hóa đơn vị, điều kiện thử nghiệm, dung sai hoặc version. Cần human review.
- Trước khi triển khai quy mô lớn, cần test với mẫu PDF scan, PDF nhiều bảng, DOCX có bảng, XLSX nhiều sheet, tài liệu trùng và tài liệu cùng model khác phiên bản.

## Gợi ý lộ trình tiếp theo

Giai đoạn 1: chạy thử và xác nhận chất lượng extraction.
Giai đoạn 2: thêm bảng alias model do người quản trị xác nhận.
Giai đoạn 3: bổ sung chuẩn hóa đơn vị theo từng đại lượng (không dùng quy tắc chung).
Giai đoạn 4: tạo báo cáo Excel/CSV về model, thông số, duplicate và potential conflicts.
