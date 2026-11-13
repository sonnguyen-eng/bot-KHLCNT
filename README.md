# Bot_KHLCNT - Tự động quét & phân tích KHLCNT Trang thiết bị Y tế

Bot tự động thu thập **Kế hoạch lựa chọn nhà thầu (KHLCNT)** mới đăng tải trên `muasamcong.mof.gov.vn`, quét sâu danh sách gói thầu con theo 19 từ khóa y tế tại 14 tỉnh/thành phố mục tiêu, phân tích bằng Gemini AI và gửi báo cáo qua Email.

## Cấu hình GitHub Secrets
Vào **Settings > Secrets and variables > Actions** và thêm các biến:
- `GEMINI_API_KEY`: Danh sách Gemini API Key (cách nhau bởi dấu phẩy `,` để tự động xoay vòng nếu hết quota)
- `SENDER_EMAIL`: Địa chỉ Gmail gửi báo cáo
- `SENDER_PASSWORD`: Mật khẩu ứng dụng (App Password) của Gmail gửi
- `RECEIVER_EMAIL`: Danh sách email nhận báo cáo (cách nhau bởi dấu phẩy `,`)
