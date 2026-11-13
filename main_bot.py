import os
import time
import json
import smtplib
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & DANH SÁCH 19 TỪ KHÓA Y TẾ
# ==============================================================================
RAW_GEMINI_KEYS = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_API_KEYS = [k.strip() for k in RAW_GEMINI_KEYS.split(",") if k.strip()]
current_key_index = 0

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "").strip() or SENDER_EMAIL

LAST_RUN_ENV = os.environ.get("LAST_RUN_TIMESTAMP", "").strip()

KEYWORDS = [
    "nội soi", "Bàn mổ", "đèn mổ", "hệ thống khí y tế", "vận chuyển mẫu",
    "ERCP", "Monitor theo dõi bệnh nhân", "Máy hấp tiệt trùng", "Nhiệt độ thấp",
    "máy gây mê", "Bơm tiêm điện", "máy truyền dịch", "bàn khám sản",
    "Monitor sản khoa", "máy hấp ướt", "giường ICU", "giường cấp cứu",
    "Bồn oxy lỏng", "Oxy trung tâm"
]

# ==============================================================================
# 2. DANH SÁCH 14 TỈNH / THÀNH PHỐ MỤC TIÊU
# ==============================================================================
TARGET_LOCATIONS = [
    "Thành phố Đà Nẵng", "Tỉnh Quảng Ngãi", "Tỉnh Đắc Lắc", "Tỉnh Đắk Lắk", "Tỉnh Gia Lai",
    "Tỉnh Khánh Hòa", "Tỉnh Lâm Đồng", "Thành phố Hồ Chí Minh", "Thành phố Đồng Nai",
    "Tỉnh Tây Ninh", "Tỉnh Đồng Tháp", "Tỉnh Vĩnh Long", "Tỉnh An Giang",
    "Thành phố Cần Thơ", "Tỉnh Cà Mau"
]

TZ_VN = timezone(timedelta(hours=7))
NOW_VN = datetime.now(TZ_VN)
TODAY_STR = NOW_VN.strftime("%H:%M %d/%m/%Y")

# Tính toán mốc START_TIME dựa trên lịch chạy tự động trước đó (mặc định lùi 24h)
if LAST_RUN_ENV:
    try:
        START_TIME = datetime.fromisoformat(LAST_RUN_ENV.replace("Z", "+00:00")).astimezone(TZ_VN)
    except Exception:
        START_TIME = NOW_VN - timedelta(hours=24)
else:
    START_TIME = NOW_VN - timedelta(hours=24)

BASE_DOMAIN = "muasamcong.mof.gov.vn"
PRIMARY_MODEL = "gemini-3.8-flash"

# ==============================================================================
# 3. KHỞI TẠO CHROME HEADLESS
# ==============================================================================
def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# ==============================================================================
# 4. LỌC THỜI GIAN ĐĂNG TRONG KHOẢNG [START_TIME -> NOW_VN] & ĐỊNH DẠNG TIỀN TỆ
# ==============================================================================
def is_published_in_time_window(item):
    pub_date_str = item.get("publicDate") or item.get("originalPublicDate") or ""
    if not pub_date_str:
        return True
    try:
        clean_date_str = pub_date_str.split(".")[0].replace("Z", "")
        dt_obj = datetime.strptime(clean_date_str, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc).astimezone(TZ_VN)
        return START_TIME <= dt_obj <= NOW_VN
    except Exception:
        return True


def format_date_str(date_str):
    if not date_str:
        return "Chưa xác định"
    try:
        clean_str = date_str.split(".")[0].replace("Z", "")
        dt_obj = datetime.strptime(clean_str, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc).astimezone(TZ_VN)
        return dt_obj.strftime("%H:%M %d/%m/%Y")
    except Exception:
        return str(date_str)


def format_currency(amount):
    if amount is None or amount == "":
        return "Chưa cập nhật"
    try:
        val = int(float(amount))
        return f"{val:,} VNĐ".replace(",", ".")
    except Exception:
        return str(amount)

# ==============================================================================
# 5. KIỂM TRA & CHUẨN HÓA ĐỊA PHƯƠNG
# ==============================================================================
def check_location(item, raw_str):
    found_locs = []
    investor = str(item.get("investorName") or item.get("procuringEntityName") or "").lower()
    plan_name = str(item.get("planName") or item.get("name") or "").lower()
    project_name = str(item.get("projectName") or "").lower()
    search_text = f"{investor} {plan_name} {project_name} {raw_str}"

    locations_arr = item.get("locations", [])
    if isinstance(locations_arr, list):
        for loc in locations_arr:
            if isinstance(loc, dict):
                prov_name = loc.get("provName", "")
                dist_name = loc.get("districtName", "")
                search_text += f" {prov_name} {dist_name}".lower()

    if any(k in search_text for k in ["đắk lắk", "đắc lắc", "tây nguyên", "buôn ma thuột"]):
        found_locs.append("Tỉnh Đắk Lắk")

    for target in TARGET_LOCATIONS:
        clean_kw = target.replace("Thành phố ", "").replace("Tỉnh ", "").lower()
        if clean_kw in search_text:
            if clean_kw == "đồng nai":
                found_locs.append("Thành phố Đồng Nai")
            elif clean_kw not in ["đắk lắk", "đắc lắc"]:
                found_locs.append(target)

    return list(set(found_locs))

# ==============================================================================
# 6. DEEP-SCAN LẤY CHI TIẾT DANH SÁCH GÓI THẦU BÊN TRONG KHLCNT
# ==============================================================================
def get_khlcnt_detail_values(driver, plan_id, token_val):
    js_get_khlcnt = """
    var callback = arguments[arguments.length - 1];
    var planId = arguments[0];
    var token = arguments[1];

    let url = '/o/egp-portal-contractor-selection-v2/services/plan/get-by-id';
    if (token) {
        url += '?token=' + token;
    }

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        },
        body: JSON.stringify({
            id: planId
        })
    })
    .then(response => response.json())
    .then(data => callback(data))
    .catch(err => callback({ error: err.toString() }));
    """
    try:
        res = driver.execute_async_script(js_get_khlcnt, plan_id, token_val)
        if isinstance(res, dict) and not res.get("error"):
            return res, json.dumps(res, ensure_ascii=False).lower()
    except Exception:
        pass
    return {}, ""

# ==============================================================================
# 7. TẠO URL CHI TIẾT KHLCNT CHUẨN V2
# ==============================================================================
def build_khlcnt_detail_url(plan_id):
    return (
        f"https://{BASE_DOMAIN}/web/guest/contractor-selection?"
        f"p_p_id=egpportalcontractorselectionv2_WAR_egpportalcontractorselectionv2"
        f"&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view"
        f"&_egpportalcontractorselectionv2_WAR_egpportalcontractorselectionv2_render=detail-v2"
        f"&type=es-plan-project-p"
        f"&stepCode=plan-step-1"
        f"&id={plan_id}"
        f"&notifyId={plan_id}"
        f"&inputResultId=undefined&bidOpenId=undefined&techBidOpenId=undefined&searchType=undefined"
    )

# ==============================================================================
# 8. TRÍCH XUẤT DỮ LIỆU KHLCNT THEO KHUNG THỜI GIAN
# ==============================================================================
def fetch_khlcnt_today(driver):
    print(f"-> Kết nối Session Mua Sắm Công KHLCNT (Khung quét: {START_TIME.strftime('%H:%M %d/%m/%Y')} -> {TODAY_STR})...")
    driver.get(f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index")
    time.sleep(10)

    js_get_token = """
    let token = '';
    let entries = window.performance.getEntriesByType('resource');
    for (let e of entries) {
        if (e.name && e.name.includes('token=')) {
            let match = e.name.match(/token=([^&]+)/);
            if (match) {
                token = match[1];
                break;
            }
        }
    }
    return token;
    """
    token_val = driver.execute_script(js_get_token)
    if token_val:
        print(f"-> Token KHLCNT trích xuất từ Session: {token_val[:20]}...")

    unique_matches = {}

    js_fetch_search_khlcnt = """
    var callback = arguments[arguments.length - 1];
    var kw = arguments[0];
    var token = arguments[1];

    let url = '/o/egp-portal-contractor-selection-v2/services/smart/search';
    if (token) {
        url += '?token=' + token;
    }

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        },
        body: JSON.stringify([{
            pageSize: 50,
            pageNumber: "0",
            query: [{
                index: "es-contractor-selection",
                keyWord: kw,
                matchType: "all-1",
                matchFields: ["planNo", "planName", "projectName", "investorName"],
                filters: [
                    { fieldName: "type", searchType: "in", fieldValues: ["KHLCNT"] }
                ]
            }]
        }])
    })
    .then(response => response.json())
    .then(data => callback(data))
    .catch(err => callback({ error: err.toString() }));
    """

    # Kết hợp 19 từ khóa chuyên sâu + các từ khóa bao quát để quét cả những KHLCNT đặt tên chung
    search_queries = KEYWORDS + ["thiết bị y tế", "y tế", "bệnh viện", "vật tư y tế", "trang thiết bị"]

    for kw in search_queries:
        print(f"-> Truy vấn KHLCNT cho từ khóa: '{kw}'...")
        try:
            res = driver.execute_async_script(js_fetch_search_khlcnt, kw, token_val)
            data_list = []

            if isinstance(res, list) and len(res) > 0:
                first_item = res[0]
                if isinstance(first_item, dict):
                    data_list = first_item.get("page", {}).get("content", [])
            elif isinstance(res, dict):
                data_list = res.get("page", {}).get("content", [])

            if not isinstance(data_list, list):
                data_list = []

            matched_in_kw = 0
            for item in data_list:
                if not isinstance(item, dict):
                    continue

                if not is_published_in_time_window(item):
                    continue

                plan_no = str(item.get("planNo") or item.get("notifyNo") or "")
                plan_id = str(item.get("id") or "")
                if not plan_no or plan_no in unique_matches:
                    continue

                plan_name = str(item.get("planName") or item.get("name") or "Chưa cập nhật tên KHLCNT")
                project_name = str(item.get("projectName") or "")
                investor = str(item.get("investorName") or item.get("procuringEntityName") or "")
                public_date = format_date_str(item.get("publicDate") or item.get("originalPublicDate"))
                invest_total = format_currency(item.get("investTotal") or item.get("totalInvest"))

                detail_url = build_khlcnt_detail_url(plan_id)
                raw_str = json.dumps(item, ensure_ascii=False).lower()

                # Deep-scan danh sách gói thầu con bên trong KHLCNT
                bid_packages_summary = []
                if plan_id:
                    detail_json, detail_str = get_khlcnt_detail_values(driver, plan_id, token_val)
                    if detail_str:
                        raw_str += " " + detail_str
                        # Trích xuất tên các gói thầu con để gửi cho AI phân tích chính xác
                        bid_list = (
                            detail_json.get("bidList")
                            or detail_json.get("data", {}).get("bidList")
                            or detail_json.get("packages")
                            or []
                        )
                        if isinstance(bid_list, list):
                            for b in bid_list[:20]:
                                if isinstance(b, dict):
                                    b_name = b.get("bidName") or b.get("name") or ""
                                    b_price = format_currency(b.get("bidPrice") or b.get("price"))
                                    b_time = b.get("bidStartTime") or b.get("selectStartTime") or ""
                                    if b_name:
                                        bid_packages_summary.append(f"{b_name} ({b_price} - Dự kiến: {b_time})")

                matched_kws = [k for k in KEYWORDS if k.lower() in raw_str]
                if matched_kws:
                    matched_province = check_location(item, raw_str)
                    if matched_province:
                        matched_in_kw += 1
                        header_text = f"{plan_no} {plan_name} {project_name}".lower()
                        location_detail = (
                            "Tên KHLCNT / Dự án"
                            if any(k.lower() in header_text for k in matched_kws)
                            else "Danh mục Gói thầu chi tiết"
                        )

                        unique_matches[plan_no] = {
                            "code": plan_no,
                            "plan_id": plan_id,
                            "title": plan_name,
                            "project_name": project_name,
                            "investor": investor,
                            "invest_total": invest_total,
                            "public_date": public_date,
                            "province": ", ".join(matched_province),
                            "matched_kw": ", ".join(matched_kws),
                            "location": location_detail,
                            "bid_packages": bid_packages_summary,
                            "detail_url": detail_url,
                            "snippet": raw_str[:1500]
                        }

            print(f"    + Server trả về {len(data_list)} KHLCNT -> Khớp khung thời gian, địa phương & từ khóa: {matched_in_kw} kế hoạch.")
            time.sleep(0.5)
        except Exception as e:
            print(f"[CẢNH BÁO FETCH KHLCNT] Từ khóa '{kw}': {e}")
            continue

    return list(unique_matches.values())

# ==============================================================================
# 9. PHÂN TÍCH KHLCNT VỚI GEMINI AI (XOAY VÒNG API KEY KHI HẾT QUOTA)
# ==============================================================================
def analyze_with_gemini(candidates):
    global current_key_index, GEMINI_API_KEYS
    if not candidates or not GEMINI_API_KEYS:
        return candidates

    while True:
        try:
            from google import genai
            from google.genai import types

            key = GEMINI_API_KEYS[current_key_index]
            client = genai.Client(api_key=key)

            prompt = f"""
Bạn là chuyên gia phân tích dữ liệu Kế hoạch lựa chọn nhà thầu (KHLCNT) Trang thiết bị Y tế.
Dưới đây là danh sách các KHLCNT phát hành từ {START_TIME.strftime('%H:%M %d/%m/%Y')} đến {TODAY_STR} đã lọc theo địa phương và 19 từ khóa y tế.

Yêu cầu:
1. Giữ nguyên toàn bộ danh sách đầu vào.
2. Dựa vào tên KHLCNT, tên dự án và danh sách gói thầu con (bid_packages), viết 1-2 câu tóm tắt ngắn gọn: gói thầu liên quan thiết bị y tế gì, giá gói thầu và thời gian dự kiến tổ chức lựa chọn nhà thầu.
3. Trả về đúng mảng JSON chứa các object đầy đủ thuộc tính và thêm trường "reason".

Danh sách:
{json.dumps(candidates, ensure_ascii=False, indent=2)}
Chỉ trả về chuỗi JSON thuần túy (không kèm markdown).
"""
            response = client.models.generate_content(
                model=PRIMARY_MODEL,
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )

            if response and response.text:
                content = response.text.strip()
                return json.loads(content)

        except Exception as e:
            error_msg = str(e)
            print(f"[CẢNH BÁO AI GOOGLE-GENAI] {error_msg}")
            if any(k in error_msg.lower() for k in ["429", "resource_exhausted", "quota"]):
                if current_key_index + 1 < len(GEMINI_API_KEYS):
                    current_key_index += 1
                    print(f"🔄 Đã đổi sang Gemini API Key thứ [{current_key_index + 1}/{len(GEMINI_API_KEYS)}]")
                    continue
            break

    return candidates

# ==============================================================================
# 10. GỬI EMAIL BÁO CÁO KHLCNT
# ==============================================================================
def send_email_report(matched_list):
    if not RECEIVER_EMAIL:
        print("[LỖI CẤU HÌNH] Chưa khai báo email nhận.")
        return

    receiver_list = [e.strip() for e in RECEIVER_EMAIL.replace("\n", "").split(",") if e.strip()]
    if not receiver_list:
        print("[LỖI CẤU HÌNH] Danh sách email nhận không hợp lệ.")
        return

    email_subject = f"[KHLCNT TBYT - Địa bàn phụ trách] Báo cáo Kế hoạch LCNT ({START_TIME.strftime('%H:%M %d/%m')} - {TODAY_STR})"

    if not matched_list:
        html_body = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h3 style="color: #d9534f;">BÁO CÁO KẾ HOẠCH LỰA CHỌN NHÀ THẦU (KHLCNT)</h3>
            <p>Hệ thống đã quét dữ liệu từ <b>{START_TIME.strftime('%H:%M %d/%m/%Y')}</b> đến <b>{TODAY_STR}</b> nhưng chưa phát hiện KHLCNT mới khớp 19 từ khóa y tế tại 14 địa phương theo dõi.</p>
        </div>
        """
    else:
        rows = ""
        for idx, item in enumerate(matched_list, 1):
            code = item.get("code", "")
            title = item.get("title", "")
            project_name = item.get("project_name", "")
            investor = item.get("investor", "")
            invest_total = item.get("invest_total", "Chưa cập nhật")
            public_date = item.get("public_date", "")
            province = item.get("province", "Theo dõi")
            location = item.get("location", "Danh mục Gói thầu chi tiết")
            reason = item.get("reason") or "Khớp từ khóa theo dõi y tế trong KHLCNT"
            matched_kw = item.get("matched_kw", "")
            detail_url = item.get("detail_url") or f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index"

            title_block = f"<b>{title}</b>"
            if project_name and project_name.lower() != title.lower():
                title_block += f"<br><small style='color: #555;'><i>Dự án: {project_name}</i></small>"

            rows += f"""
            <tr>
                <td align="center">{idx}</td>
                <td><b>{code}</b><br><small style="color: #666;">{public_date}</small></td>
                <td>{title_block}</td>
                <td align="right"><span style="color: #d9534f; font-weight: bold;">{invest_total}</span></td>
                <td>{investor}</td>
                <td align="center"><b><span style="color: #0275d8;">{province}</span></b></td>
                <td align="center"><span style="background-color: #ffffcc; padding: 3px 6px; border-radius: 3px;">{matched_kw}</span></td>
                <td align="center"><i>{location}</i></td>
                <td><small style="color: #333;">{reason}</small></td>
                <td align="center"><a href="{detail_url}" target="_blank" style="color: #0275d8; font-weight: bold;">Xem ngay</a></td>
            </tr>
            """

        html_body = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h2 style="color: #0275d8;">BÁO CÁO KẾ HOẠCH LỰA CHỌN NHÀ THẦU (KHLCNT)</h2>
            <p>Khung thời gian quét: từ <b>{START_TIME.strftime('%H:%M %d/%m/%Y')}</b> đến <b>{TODAY_STR}</b></p>
            <p>Đã lọc phát hiện <b>{len(matched_list)}</b> KHLCNT Y tế thuộc địa bàn phụ trách:</p>
            <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border-color: #ddd;">
                <tr style="background-color: #e6f2ff;">
                    <th align="center">STT</th>
                    <th>Mã KHLCNT / Ngày đăng</th>
                    <th>Tên KHLCNT & Dự án</th>
                    <th align="center">Tổng Mức Đầu Tư</th>
                    <th>Chủ Đầu Tư</th>
                    <th align="center">Tỉnh / TP</th>
                    <th align="center">Từ Khóa Khớp</th>
                    <th align="center">Vị Trí Khớp</th>
                    <th>Phân Tích Gói Thầu</th>
                    <th align="center">Chi Tiết</th>
                </tr>
                {rows}
            </table>
        </div>
        """

    if SENDER_EMAIL and SENDER_PASSWORD:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = email_subject
        msg["From"] = SENDER_EMAIL
        msg["To"] = ", ".join(receiver_list)
        msg.attach(MIMEText(html_body, "html"))
        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, receiver_list, msg.as_string())
            print(f"[SUCCESS] Đã gửi báo cáo KHLCNT thành công tới các email: {', '.join(receiver_list)}!")
        except Exception as e:
            print(f"[LỖI GỬI EMAIL SMTP] {e}")


def main():
    print(f"=== CRAWL & DEEP-SCAN KHLCNT (Khung quét: {START_TIME.strftime('%H:%M %d/%m/%Y')} -> {TODAY_STR}) ===")
    driver = init_driver()
    try:
        matched_list = fetch_khlcnt_today(driver)
        final_matches = analyze_with_gemini(matched_list)
        send_email_report(final_matches)
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
