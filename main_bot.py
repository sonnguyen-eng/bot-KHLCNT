import os
import re
import time
import json
import smtplib
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & BỘ TỪ KHÓA LỌC CHÍNH XÁC TBYT / PACS / DỰ ÁN BV
# ==============================================================================
RAW_GEMINI_KEYS = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_API_KEYS = [k.strip() for k in RAW_GEMINI_KEYS.split(",") if k.strip()]
current_key_index = 0

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "").strip() or SENDER_EMAIL

LAST_RUN_ENV = os.environ.get("LAST_RUN_TIMESTAMP", "").strip()
EVENT_NAME_ENV = os.environ.get("GITHUB_EVENT_NAME", "").strip()

# Nhóm 1: Từ khóa thiết bị y tế & giải pháp PACS đặc thù (Đầy đủ toàn bộ BU)
SPECIFIC_DEVICE_KEYWORDS = [
    # Nội soi - Phẫu thuật - Laser
    "nội soi", "ống soi", "optic nội soi", "dụng cụ nội soi", "ercp",
    "robot phẫu thuật", "tán sỏi laser", "máy tán sỏi", "dao mổ",
    # Phòng mổ - Gây mê - Hồi sức - Giường
    "bàn mổ", "đèn mổ", "thiết bị phòng mổ", "máy gây mê",
    "monitor theo dõi bệnh nhân", "máy theo dõi bệnh nhân",
    "bơm tiêm điện", "máy truyền dịch",
    "giường icu", "giường cấp cứu", "giường đa năng", "giường bệnh",
    "hồi sức cấp cứu", "hồi sức sơ sinh", "máy thở", "máy phá rung tim",
    # Khí y tế - Vận chuyển - Tiệt trùng
    "hệ thống khí y tế", "bồn oxy lỏng", "oxy trung tâm", "vận chuyển mẫu",
    "máy hấp tiệt trùng", "nhiệt độ thấp", "máy hấp ướt", "nồi hấp",
    # Sản khoa - Chẩn đoán hình ảnh - PACS - Thận nhân tạo
    "bàn khám sản", "monitor sản khoa", "c-arm", "c arm", "x-quang",
    "máy siêu âm", "siêu âm màu", "siêu âm chẩn đoán",
    "thận nhân tạo", "máy chạy thận", "hdf online",
    "pacs", "ris/pacs", "ris - pacs", "lưu trữ và truyền tải hình ảnh",
    # Gói mua sắm / sửa chữa thiết bị y tế tổng thể
    "thiết bị y tế chuyên môn", "thiết bị y tế chuyên dùng",
    "thiết bị y tế cấp thiết", "mua sắm thiết bị chuyên môn",
    "mua sắm thiết bị y tế", "mua thiết bị y tế",
    "cung cấp, lắp đặt thiết bị y tế", "cung cấp lắp đặt thiết bị y tế",
    "trang thiết bị y tế", "sửa chữa thiết bị y tế", "sửa chữa máy"
]

# Nhóm 2: Từ khóa có điều kiện (CHỈ khớp khi đi kèm đơn vị/công trình ngành Y tế)
CONDITIONAL_KEYWORDS = [
    "thẩm định giá", "chẩn đoán hình ảnh", "phục vụ chuyên môn"
]

# Thực thể / Công trình thuộc ngành Y tế
MEDICAL_ENTITY_TERMS = [
    "bệnh viện", "trung tâm y tế", "viện tim", "ngành y tế", "sở y tế",
    "đại học y khoa", "đại học y dược", "khối khám bệnh", "khối bệnh nhiệt đới",
    "khoa khám bệnh", "khu khám bệnh", "khu hồi sức", "trung tâm sơ sinh",
    "trung tâm y khoa", "trạm y tế", "y tế khu vực"
]

# Hành động đầu tư xây mới / nâng cấp / chuẩn bị dự án công trình BV quy mô lớn
CONSTRUCTION_PROJECT_TERMS = [
    "xây dựng mới", "chuẩn bị đầu tư", "chuẩn bị dự án",
    "báo cáo nghiên cứu khả thi", "thiết kế bản vẽ thi công",
    "nâng cấp", "cải tạo, nâng cấp", "cải tạo nâng cấp",
    "xây dựng khối", "xây dựng khu", "mua sắm thiết bị"
]

# Từ khóa loại trừ các dự án ngoài ngành y tế, dịch vụ thuê/khám sức khỏe, duy tu nhỏ
EXCLUDE_NON_MEDICAL_PATTERNS = [
    "thẩm định giá đất", "thi hành bản án", "tuần tra, kiểm tra rừng",
    "dây chuyền kiểm định phương tiện", "cầu đi bộ", "tái định cư",
    "công viên", "nhà văn hóa", "trường tiểu học", "trường thcs", "trường thpt",
    "cao đẳng kỹ thuật", "trụ sở văn phòng", "hệ thống chiếu sáng",
    "mở rộng, nâng cấp tuyến đường", "kiểm định máy x-quang",
    "kiểm xạ an toàn", "liều kế cá nhân", "khám sức khỏe",
    "dịch vụ chụp x-quang", "xe lưu động", "ốp gạch", "sơn tường",
    "chống thấm", "lợp mái", "vệ sinh công nghiệp", "giặt là", "cây xanh"
]

# Cụm từ nhận diện gói VẬT TƯ TIÊU HAO thuần túy (không phải thiết bị y tế)
CONSUMABLE_MEDICAL_TERMS = [
    "phim x-quang", "phim chụp", "bơm, kim tiêm", "bơm kim tiêm",
    "dây truyền", "găng tay", "bông, dung dịch", "băng, gạc",
    "vật tư tiêu hao", "hóa chất xét nghiệm", "sinh phẩm",
    "vật tư y tế sử dụng trong", "vật tư dùng trong"
]

# Cụm từ nhận diện gói THUÊ PACS / THUÊ THIẾT BỊ / THUÊ PHẦN MỀM (loại trừ tuyệt đối)
RENTAL_TERMS = [
    "thuê dịch vụ sử dụng phần mềm", "thuê dịch vụ phần mềm",
    "thuê dịch vụ hệ thống", "thuê hệ thống ris", "thuê hệ thống pacs",
    "thuê hệ thống phần mềm", "thuê phần mềm", "thuê pacs", "thuê ris/pacs",
    "dịch vụ hệ thống thông tin", "ca chụp/tháng",
    "thuê máy", "cho thuê máy", "thuê thiết bị", "cho thuê thiết bị",
    "sl: 12 tháng", "sl: 24 tháng", "sl: 36 tháng", "sl: 6 tháng"
]

BROAD_QUERIES = [
    "thiết bị y tế", "trang thiết bị y tế", "bệnh viện",
    "trung tâm y tế", "ngành y tế", "ban quản lý dự án",
    "chuẩn bị dự án", "chuẩn bị đầu tư", "xây dựng mới", "mua sắm"
]

# ==============================================================================
# 2. DANH SÁCH 14 TỈNH / THÀNH PHỐ MỤC TIÊU PHỤ TRÁCH
# ==============================================================================
LOCATION_ALIASES = {
    "Thành phố Đà Nẵng": ["đà nẵng", "quảng nam", "tam kỳ", "hội an", "việt - hàn đà nẵng"],
    "Tỉnh Quảng Ngãi": ["quảng ngãi", "kon tum"],
    "Tỉnh Đắk Lắk": ["đắk lắk", "đắc lắc", "tây nguyên", "buôn ma thuột", "phú yên", "tuy hòa"],
    "Tỉnh Gia Lai": ["gia lai", "pleiku", "bình định", "quy nhơn"],
    "Tỉnh Khánh Hòa": ["khánh hòa", "nha trang", "ninh thuận", "phan rang"],
    "Tỉnh Lâm Đồng": ["lâm đồng", "đà lạt", "bảo lộc", "đắk nông", "bình thuận", "phan thiết"],
    "Thành phố Hồ Chí Minh": [
        "hồ chí minh", "tp.hcm", "tphcm", "tp hcm", "sài gòn", "thủ đức",
        "bình dương", "bà rịa", "vũng tàu", "phú mỹ", "nhi đồng", "bình dân",
        "viện tim", "nhân ái", "đại học y dược", "ung bướu", "gò vấp",
        "bệnh nhiệt đới", "phạm ngọc thạch", "chấn thương", "lê văn thịnh",
        "thuận an", "tân uyên"
    ],
    "Thành phố Đồng Nai": ["đồng nai", "biên hòa", "long khánh", "bình phước", "đồng xoài", "long thành"],
    "Tỉnh Tây Ninh": ["tây ninh", "long an", "tân an", "thủ thừa", "tân hưng"],
    "Tỉnh Đồng Tháp": ["đồng tháp", "cao lãnh", "sa đéc", "tiền giang", "mỹ tho", "quân y 120"],
    "Tỉnh Vĩnh Long": ["vĩnh long", "bến tre", "trà vinh", "nguyễn đình chiểu"],
    "Tỉnh An Giang": ["an giang", "long xuyên", "châu đốc", "kiên giang", "rạch giá", "phú quốc", "ung bướu kiên giang"],
    "Thành phố Cần Thơ": ["cần thơ", "hậu giang", "sóc trăng", "thốt nốt"],
    "Tỉnh Cà Mau": ["cà mau", "bạc liêu"]
}

TZ_VN = timezone(timedelta(hours=7))
NOW_VN = datetime.now(TZ_VN)
TODAY_STR = NOW_VN.strftime("%H:%M %d/%m/%Y")

if LAST_RUN_ENV:
    try:
        parsed_last_run = datetime.fromisoformat(LAST_RUN_ENV.replace("Z", "+00:00")).astimezone(TZ_VN)
        START_TIME = parsed_last_run - timedelta(minutes=15)
    except Exception:
        START_TIME = NOW_VN - timedelta(hours=48)
else:
    START_TIME = NOW_VN - timedelta(hours=48)

BASE_DOMAIN = "muasamcong.mof.gov.vn"
PRIMARY_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"

SERVER_MAX_PAGE_SIZE = 50
SPECIFIC_MAX_RECORDS = 150
BROAD_MAX_RECORDS = 500

BID_FORM_MAP = {
    "DTRR": "Đấu thầu rộng rãi",
    "DTHC": "Đấu thầu hạn chế",
    "CDT": "Chỉ định thầu",
    "CDTRG": "Chỉ định thầu rút gọn",
    "CHCT": "Chào hàng cạnh tranh",
    "CHCTRG": "Chào hàng cạnh tranh rút gọn",
    "MSTT": "Mua sắm trực tiếp",
    "TTH": "Tự thực hiện"
}

# ==============================================================================
# 3. KHỞI TẠO CHROME HEADLESS & ĐỌC LỊCH SỬ CHỐNG TRÙNG TỪ GOOGLE SHEET
# ==============================================================================
def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

def get_existing_codes_from_gsheet():
    gsheet_id = os.environ.get("GSHEET_ID", "").strip()
    sa_json = os.environ.get("GCP_SERVICE_ACCOUNT_JSON", "").strip()
    if not gsheet_id or not sa_json:
        return set()
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_dict = json.loads(sa_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(gsheet_id)

        existing_codes = set()
        for ws in sh.worksheets():
            all_vals = ws.get_all_values()
            for row in all_vals[1:]:
                for cell in row[:4]:
                    val = str(cell).strip()
                    if val.startswith(("PL", "RQ", "PR", "IB")):
                        existing_codes.add(val)
        print(f"-> Đã tải {len(existing_codes)} mã KHLCNT lịch sử từ Google Sheet để chống trùng.")
        return existing_codes
    except Exception as e:
        print(f"[CẢNH BÁO ĐỌC LỊCH SỬ GSHEET] {e}")
        return set()

# ==============================================================================
# 4. LỌC THỜI GIAN & ĐỊNH DẠNG DỮ LIỆU
# ==============================================================================
def parse_public_date(pub_date_str):
    if not pub_date_str:
        return None
    try:
        clean_date_str = pub_date_str.split(".")[0].replace("Z", "")
        return datetime.strptime(clean_date_str, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=TZ_VN)
    except Exception:
        return None

def is_published_in_time_window(item):
    pub_date_str = item.get("publicDate") or item.get("originalPublicDate") or ""
    dt_obj = parse_public_date(pub_date_str)
    if not dt_obj:
        return True
    return START_TIME <= dt_obj <= NOW_VN

def format_date_str(date_str):
    if not date_str:
        return "Chưa xác định"
    try:
        clean_str = date_str.split(".")[0].replace("Z", "")
        dt_obj = datetime.strptime(clean_str, "%Y-%m-%dT%H:%M:%S")
        return dt_obj.strftime("%H:%M %d/%m/%Y")
    except Exception:
        return str(date_str)

def parse_numeric_amount(amount):
    if amount is None or amount == "":
        return 0.0
    try:
        if isinstance(amount, list) and len(amount) > 0:
            amount = amount[0]
        return float(amount)
    except Exception:
        return 0.0

def format_currency(amount):
    val = parse_numeric_amount(amount)
    if val <= 0:
        return "Chưa cập nhật"
    return f"{int(val):,} VND".replace(",", ".")

def format_expected_bid_time(pkg):
    bid_time_str = str(pkg.get("bidTime") or "").strip()
    unit = pkg.get("bidStartUnit", "")
    year = pkg.get("bidStartYear", "")
    quarter = pkg.get("bidStartQuarter", "")
    month = pkg.get("bidStartMonth", "")
    period_str = ""
    if unit == "Q" and quarter and year:
        period_str = f"Quý {quarter}/{year}"
    elif unit == "M" and month and year:
        period_str = f"Tháng {month}/{year}"
    elif year:
        period_str = f"Năm {year}"

    if period_str and bid_time_str:
        return f"{period_str} (TGTH: {bid_time_str})"
    return period_str or bid_time_str or "Chưa rõ thời gian"

# ==============================================================================
# 5. KIỂM TRA & CHUẨN HÓA ĐỊA PHƯƠNG TRONG 14 TỈNH / THÀNH PHỐ MỤC TIÊU
# ==============================================================================
def check_location(item, raw_str):
    found_locs = []
    investor = str(item.get("investorName") or item.get("procuringEntityName") or "").lower()
    plan_name = str(item.get("name") or item.get("pname") or "").lower()
    search_text = f"{investor} {plan_name} {raw_str}".lower()

    locations_arr = item.get("locations", [])
    if isinstance(locations_arr, list):
        for loc in locations_arr:
            if isinstance(loc, dict):
                prov_name = loc.get("provName", "")
                dist_name = loc.get("districtName", "")
                search_text += f" {prov_name} {dist_name}".lower()

    for canonical_name, aliases in LOCATION_ALIASES.items():
        if any(alias in search_text for alias in aliases):
            found_locs.append(canonical_name)

    return list(set(found_locs))

# ==============================================================================
# 6. DEEP-SCAN CHI TIẾT KHLCNT & DANH SÁCH GÓI THẦU CON
# ==============================================================================
def get_khlcnt_detail_values(driver, plan_id, token_val):
    js_get_khlcnt = """
    var callback = arguments[arguments.length - 1];
    var planId = arguments[0];
    var token = arguments[1];

    let url = '/o/egp-portal-contractor-selection-v2/services/expose/lcnt/bid-po-bidp-plan-project-view/get-by-id';
    if (token) {
        url += '?token=' + token;
    }

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        },
        body: JSON.stringify({ id: planId })
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
def build_khlcnt_detail_url(plan_id, plan_no=""):
    return (
        f"https://{BASE_DOMAIN}/web/guest/contractor-selection?"
        f"p_p_id=egpportalcontractorselectionv2_WAR_egpportalcontractorselectionv2"
        f"&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view"
        f"&_egpportalcontractorselectionv2_WAR_egpportalcontractorselectionv2_render=detail-v2"
        f"&type=es-plan-project-p"
        f"&stepCode=plan-step-1"
        f"&id={plan_id}"
        f"&notifyId=undefined"
        f"&planNo={plan_no}"
        f"&step=khlcnt"
        f"&inputResultId=undefined&bidOpenId=undefined&techBidOpenId=undefined&searchType=undefined"
    )

# ==============================================================================
# 8. TRÍCH XUẤT & LÀM MỚI RECAPTCHA TOKEN TỪ SESSION TRÌNH DUYỆT
# ==============================================================================
def get_latest_token_from_perf(driver):
    js_get_token = """
    let token = '';
    let entries = window.performance.getEntriesByType('resource');
    for (let i = entries.length - 1; i >= 0; i--) {
        let e = entries[i];
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
    return driver.execute_script(js_get_token) or ""

def refresh_session_token(driver, kw_hint="y tế"):
    driver.execute_script("window.performance.clearResourceTimings();")
    try:
        inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input[placeholder*='tìm'], input[placeholder*='Tìm'], input[placeholder*='từ khóa']")
        for inp in inputs:
            if inp.is_displayed():
                inp.click()
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(Keys.BACKSPACE)
                inp.send_keys(kw_hint)
                inp.send_keys(Keys.RETURN)
                time.sleep(5)
                new_tok = get_latest_token_from_perf(driver)
                if new_tok:
                    return new_tok
                break
    except Exception:
        pass

    driver.get(f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index")
    time.sleep(8)
    return get_latest_token_from_perf(driver)

# ==============================================================================
# 9. BỘ LỌC 2 LỚP: LOẠI BỎ VẬT TƯ TIÊU HAO, THUÊ MÁY, SỬA CHỮA NHỎ; BẮT TBYT & DA BV
# ==============================================================================
def is_consumable_package(text_lower):
    return any(c in text_lower for c in CONSUMABLE_MEDICAL_TERMS)

def is_rental_package(text_lower):
    if any(r in text_lower for r in RENTAL_TERMS):
        return True
    if re.search(r"\bthuê\s+\d*\s*(máy|thiết bị|hệ thống)", text_lower):
        return True
    return False

def is_equipment_package(text_lower):
    if is_consumable_package(text_lower) or is_rental_package(text_lower):
        return False
    if any(excl in text_lower for excl in EXCLUDE_NON_MEDICAL_PATTERNS):
        return False
    equipment_terms = [
        "thiết bị y tế", "trang thiết bị y tế", "nội soi", "ống soi", "dụng cụ nội soi",
        "bàn mổ", "đèn mổ", "thiết bị phòng mổ", "dao mổ", "gây mê", "hồi sức",
        "giường icu", "giường cấp cứu", "giường đa năng", "giường bệnh",
        "máy x-quang", "hệ thống x-quang", "c-arm", "c arm",
        "máy siêu âm", "máy thở", "tán sỏi", "khí y tế", "oxy lỏng", "oxy trung tâm",
        "tiệt trùng", "nồi hấp", "monitor", "sơ sinh", "thận nhân tạo", "máy chạy thận",
        "robot phẫu thuật", "pacs", "ris/pacs", "lưu trữ và truyền tải hình ảnh",
        "sửa chữa thiết bị y tế", "sửa chữa máy"
    ]
    return any(t in text_lower for t in equipment_terms)

def evaluate_medical_relevance(plan_name, investor, pkg_names_list, medical_pkg_names, raw_str):
    plan_lower = plan_name.lower()
    investor_lower = investor.lower()
    pkgs_lower = " ".join(pkg_names_list).lower()
    core_text = f"{plan_lower} {investor_lower} {pkgs_lower}"
    content_text = f"{plan_lower} {pkgs_lower}"

    # 1. Chốt chặn loại trừ các dự án ngoài ngành y tế, khám sức khỏe, duy tu/ốp gạch/sơn sửa nhỏ
    if any(excl in core_text for excl in EXCLUDE_NON_MEDICAL_PATTERNS):
        return False, False, []

    if any(infra in plan_lower for infra in ["cầu đi bộ", "tuyến đường", "khu tái định cư", "hệ thống chiếu sáng", "trụ sở", "nhà văn hóa", "công viên"]):
        return False, False, []

    # 2. Loại trừ TUYỆT ĐỐI các gói THUÊ máy / THUÊ thiết bị / THUÊ PACS / THUÊ phần mềm
    if is_rental_package(content_text):
        return False, False, []

    # 3. Nếu đã bóc tách được danh sách gói thầu con mà KHÔNG có gói thiết bị nào (chỉ toàn vật tư tiêu hao)
    if pkg_names_list and not medical_pkg_names:
        has_medical_entity = any(ent in f"{plan_lower} {investor_lower}" for ent in MEDICAL_ENTITY_TERMS)
        is_major_hospital_proj = has_medical_entity and any(act in plan_lower for act in CONSTRUCTION_PROJECT_TERMS)
        if not is_major_hospital_proj:
            return False, False, []

    # Nếu không có gói con (hoặc gói con rỗng) nhưng tên kế hoạch là vật tư tiêu hao / phim X-quang -> loại
    if not medical_pkg_names and is_consumable_package(content_text):
        return False, False, []

    # Nếu chỉ là phần mềm bệnh án điện tử / HIS đơn thuần mà không có mua PACS hoặc TBYT thì bỏ qua
    if ("bệnh án điện tử" in content_text or "phần mềm quản lý bệnh viện" in content_text) and not any(
        p in content_text for p in ["pacs", "ris/pacs", "ris - pacs", "lưu trữ và truyền tải hình ảnh", "thiết bị y tế"]
    ):
        return False, False, []

    # 4. Kiểm tra thực thể y tế
    has_medical_entity = any(ent in f"{plan_lower} {investor_lower}" for ent in MEDICAL_ENTITY_TERMS)

    # 5. Kiểm tra từ khóa thiết bị y tế & mua đứt PACS đặc thù (ưu tiên trên các gói thiết bị sạch + generalTasks)
    clean_scope = " ".join(medical_pkg_names).lower() if medical_pkg_names else (content_text if pkgs_lower.strip() else raw_str)
    matched_specific = [k for k in SPECIFIC_DEVICE_KEYWORDS if k in clean_scope]

    # Loại bỏ khớp nhầm "x-quang" nếu chỉ là "phim x-quang"
    if "x-quang" in matched_specific and "phim x-quang" in clean_scope and not any(
        dev in clean_scope for dev in ["máy x-quang", "hệ thống x-quang", "x-quang kỹ thuật số", "x-quang di động", "x-quang cao tần", "c-arm"]
    ):
        matched_specific = [k for k in matched_specific if k != "x-quang"]

    # 6. Kiểm tra từ khóa có điều kiện (CHỈ tính khi thuộc thực thể y tế và xuất hiện trong gói thầu/tên KH)
    matched_cond = []
    if has_medical_entity:
        for ck in CONDITIONAL_KEYWORDS:
            if ck in pkgs_lower or ck in plan_lower:
                matched_cond.append(ck)

    # 7. Kiểm tra dự án xây mới / nâng cấp / chuẩn bị đầu tư công trình Bệnh viện
    is_hospital_proj = False
    hospital_proj_label = ""
    if has_medical_entity:
        matched_actions = [act for act in CONSTRUCTION_PROJECT_TERMS if act in plan_lower]
        if matched_actions:
            is_hospital_proj = True
            hospital_proj_label = f"Dự án công trình BV ({', '.join(matched_actions[:2])})"

    all_matched_tags = list(dict.fromkeys(matched_specific + matched_cond))
    if is_hospital_proj and hospital_proj_label not in all_matched_tags:
        all_matched_tags.append(hospital_proj_label)

    is_relevant = bool(all_matched_tags)
    return is_relevant, is_hospital_proj, all_matched_tags

# ==============================================================================
# 10. TRÍCH XUẤT KHLCNT 14 TỈNH/TP (TBYT + MUA PACS + DỰ ÁN CÔNG TRÌNH BV)
# ==============================================================================
def fetch_khlcnt_target_provinces(driver):
    print(f"-> Chế độ chạy: {EVENT_NAME_ENV or 'local'} | Mốc lịch sử gốc: {LAST_RUN_ENV or 'Mặc định 48h'}")
    print(f"-> Kết nối Session Mua Sắm Công KHLCNT - 14 Tỉnh/TP (Khung quét có đệm: {START_TIME.strftime('%H:%M %d/%m/%Y')} -> {TODAY_STR})...")
    driver.get(f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index")
    time.sleep(9)

    token_val = get_latest_token_from_perf(driver)
    if not token_val:
        token_val = refresh_session_token(driver, "bệnh viện")

    if token_val:
        print(f"-> Token KHLCNT trích xuất từ Session: {token_val[:20]}...")
    else:
        print(f"-> [CẢNH BÁO SESSION] Chưa trích xuất được token (Page Title: '{driver.title}').")

    unique_matches = {}
    scanned_plan_ids = set()
    api_calls_since_refresh = 0

    from_date_str = START_TIME.strftime("%Y-%m-%dT%H:%M:%S.000")
    to_date_str = NOW_VN.strftime("%Y-%m-%dT%H:%M:%S.999")

    js_fetch_search_khlcnt = """
    var callback = arguments[arguments.length - 1];
    var kw = arguments[0];
    var pageNum = arguments[1];
    var pageSize = arguments[2];
    var fromDate = arguments[3];
    var toDate = arguments[4];
    var token = arguments[5];

    let url = '/o/egp-portal-contractor-selection-v2/services/smart/search';
    if (token) {
        url += '?token=' + token;
    }

    const baseFilter = [
        { fieldName: "type", searchType: "in", fieldValues: ["es-plan-project-p"] },
        { fieldName: "publicDate", searchType: "range", from: fromDate, to: toDate }
    ];

    const payload = [
        {
            pageSize: pageSize,
            pageNumber: String(pageNum),
            query: [{
                index: "es-contractor-selection",
                keyWord: kw,
                matchType: "all-1",
                matchFields: ["planNo", "name"],
                filters: baseFilter
            }]
        },
        {
            pageSize: pageSize,
            pageNumber: String(pageNum),
            query: [{
                index: "es-contractor-selection",
                keyWord: kw,
                matchType: "all-1",
                matchFields: ["generalTasks"],
                filters: baseFilter
            }]
        },
        {
            pageSize: pageSize,
            pageNumber: String(pageNum),
            query: [{
                index: "es-contractor-selection",
                keyWord: kw,
                matchType: "all-1",
                matchFields: ["bidName"],
                filters: baseFilter
            }]
        }
    ];

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        },
        body: JSON.stringify(payload)
    })
    .then(async response => {
        const rawText = await response.text();
        try {
            callback({
                status: response.status,
                body: JSON.parse(rawText)
            });
        } catch (e) {
            callback({
                status: response.status,
                rawText: rawText.substring(0, 300)
            });
        }
    })
    .catch(err => callback({ error: err.toString() }));
    """

    query_plan = []
    for kw in SPECIFIC_DEVICE_KEYWORDS:
        query_plan.append((kw, SPECIFIC_MAX_RECORDS))
    for kw in BROAD_QUERIES:
        query_plan.append((kw, BROAD_MAX_RECORDS))

    for kw, max_records in query_plan:
        if api_calls_since_refresh >= 10:
            new_tok = refresh_session_token(driver, kw)
            if new_tok:
                token_val = new_tok
            api_calls_since_refresh = 0

        print(f"-> Quét 14 tỉnh từ khóa: '{kw}' (Mục tiêu tối đa {max_records} bản ghi/nhóm)...")
        matched_in_kw = 0
        province_in_kw = 0
        fetched_per_block = 0
        total_fetched_in_kw = 0
        max_total_elements = None
        page_num = 0
        retried_token_for_kw = False

        while fetched_per_block < max_records:
            current_page_size = min(SERVER_MAX_PAGE_SIZE, max_records - fetched_per_block)
            try:
                raw_res = driver.execute_async_script(
                    js_fetch_search_khlcnt,
                    kw,
                    page_num,
                    current_page_size,
                    from_date_str,
                    to_date_str,
                    token_val
                )
                api_calls_since_refresh += 1

                res = raw_res.get("body") if isinstance(raw_res, dict) else raw_res
                data_list = []
                max_block_len = 0
                block_totals = []

                if isinstance(res, list):
                    for block in res:
                        if isinstance(block, dict):
                            page_obj = block.get("page", {})
                            content = page_obj.get("content", [])
                            t_el = page_obj.get("totalElements")
                            if isinstance(t_el, int):
                                block_totals.append(t_el)
                            if isinstance(content, list):
                                data_list.extend(content)
                                if len(content) > max_block_len:
                                    max_block_len = len(content)
                elif isinstance(res, dict):
                    page_obj = res.get("page", {})
                    data_list = page_obj.get("content", [])
                    max_block_len = len(data_list)
                    t_el = page_obj.get("totalElements")
                    if isinstance(t_el, int):
                        block_totals.append(t_el)

                if max_total_elements is None and block_totals:
                    max_total_elements = max(block_totals)

                if not data_list and page_num == 0 and max_total_elements is None and not retried_token_for_kw:
                    retried_token_for_kw = True
                    new_tok = refresh_session_token(driver, kw)
                    if new_tok and new_tok != token_val:
                        token_val = new_tok
                        api_calls_since_refresh = 0
                        continue

                if not data_list:
                    break

                fetched_per_block += current_page_size
                total_fetched_in_kw += len(data_list)

                for item in data_list:
                    if not isinstance(item, dict):
                        continue

                    if not is_published_in_time_window(item):
                        continue

                    plan_no = str(item.get("planNo") or item.get("planNoStand") or "").strip()
                    plan_ver = str(item.get("planVersion") or "00").strip()
                    plan_id = str(item.get("id") or "").strip()
                    if not plan_no or plan_no in unique_matches or plan_id in scanned_plan_ids:
                        continue

                    scanned_plan_ids.add(plan_id)

                    plan_name = str(item.get("name") or item.get("pname") or "Chưa cập nhật tên KHLCNT").replace("\n", " ").strip()
                    investor = str(item.get("investorName") or item.get("procuringEntityName") or "").strip()
                    public_date = format_date_str(item.get("publicDate") or item.get("originalPublicDate"))
                    numeric_invest = parse_numeric_amount(item.get("investTotal"))
                    invest_total = format_currency(numeric_invest)
                    detail_url = build_khlcnt_detail_url(plan_id, plan_no)

                    raw_str = json.dumps(item, ensure_ascii=False).lower()

                    matched_province = check_location(item, raw_str)
                    if not matched_province:
                        continue

                    province_in_kw += 1

                    decision_no = ""
                    decision_file = ""
                    bid_packages_summary = []
                    all_pkg_names = []
                    medical_pkg_names = []
                    general_tasks_list = []
                    medical_pkg_total = 0.0
                    expected_times = set()

                    if plan_id:
                        detail_json, detail_str = get_khlcnt_detail_values(driver, plan_id, token_val)
                        if detail_str:
                            raw_str += " " + detail_str
                        plan_view = detail_json.get("bidPoBidpPlanProjectDetailView") or {}
                        decision_no = str(plan_view.get("decisionNo") or "").strip()
                        if plan_ver and plan_ver != "00":
                            decision_no = f"{decision_no} (ĐC lần {plan_ver})".strip()
                        decision_file = str(plan_view.get("decisionFileName") or "").strip()

                        pkg_list = detail_json.get("bidpPlanDetailToProjectList") or []
                        if isinstance(pkg_list, list):
                            for pkg in pkg_list:
                                if not isinstance(pkg, dict):
                                    continue
                                b_no = str(pkg.get("bidNo") or pkg.get("code") or "").strip()
                                b_name = str(pkg.get("bidName") or "").replace("\n", " ").strip()
                                b_tasks = str(pkg.get("generalTasks") or "").replace("\n", " ").strip()
                                b_num_price = parse_numeric_amount(pkg.get("bidPrice"))
                                b_price = format_currency(b_num_price)
                                b_form_code = pkg.get("bidForm") or ""
                                b_form_text = BID_FORM_MAP.get(b_form_code, b_form_code)
                                b_time_expected = format_expected_bid_time(pkg)
                                if b_time_expected != "Chưa rõ thời gian":
                                    expected_times.add(b_time_expected)

                                combined_pkg_text = f"{b_name} {b_tasks}".strip()
                                if combined_pkg_text:
                                    all_pkg_names.append(combined_pkg_text)
                                    prefix_code = f"{b_no} - " if b_no else ""
                                    task_note = f" | Tóm tắt công việc: {b_tasks}" if (b_tasks and b_tasks.lower() != b_name.lower()) else ""
                                    bid_packages_summary.append(
                                        f"{prefix_code}{b_name}{task_note} ({b_price} | {b_form_text} | Dự kiến: {b_time_expected})"
                                    )
                                    if is_equipment_package(combined_pkg_text.lower()):
                                        medical_pkg_names.append(f"{prefix_code}{combined_pkg_text}")
                                        medical_pkg_total += b_num_price
                                        if b_tasks:
                                            general_tasks_list.append(b_tasks)

                    is_relevant, is_hospital_proj, kw_tags = evaluate_medical_relevance(
                        plan_name, investor, all_pkg_names, medical_pkg_names, raw_str
                    )

                    if is_relevant:
                        matched_in_kw += 1
                        final_kw_str = ", ".join(kw_tags)

                        is_early_prep = is_hospital_proj and not medical_pkg_names and any(
                            p in plan_name.lower()
                            for p in ["chuẩn bị đầu tư", "chuẩn bị dự án", "xây dựng mới", "báo cáo nghiên cứu khả thi", "thiết kế bản vẽ thi công", "nâng cấp", "cải tạo, nâng cấp", "cải tạo nâng cấp"]
                        )

                        if is_early_prep:
                            display_title = plan_name
                            display_price = invest_total if invest_total != "Chưa cập nhật" else format_currency(medical_pkg_total)
                            stage_label = "Tiền dự án (Chuẩn bị đầu tư / Xây mới BV)"
                        else:
                            display_title = "; ".join(medical_pkg_names) if medical_pkg_names else plan_name
                            display_price = format_currency(medical_pkg_total) if medical_pkg_total > 0 else invest_total
                            stage_label = "KHLCNT (Kế hoạch được duyệt)"

                        has_generic_pkg = any(
                            any(g in p.lower() for g in ["cung cấp, lắp đặt thiết bị y tế", "cung cấp lắp đặt thiết bị y tế", "mua sắm thiết bị y tế"])
                            and not any(sp in p.lower() for sp in ["nội soi", "x-quang", "siêu âm", "máy thở", "bàn mổ", "tán sỏi", "sơ sinh", "thận nhân tạo", "pacs", "dao mổ"])
                            for p in (medical_pkg_names or [plan_name])
                        )

                        if is_early_prep:
                            auto_note = "Giai đoạn chuẩn bị đầu tư / xây mới công trình BV - Theo dõi sát khâu lập danh mục TBYT"
                        elif has_generic_pkg:
                            auto_note = f"Gói thiết bị tổng hợp (chưa tách tên máy trên KHLCNT - File QĐ: {decision_file})" if decision_file else "Đang thực hiện xây danh mục, cấu hình thiết bị"
                        elif general_tasks_list:
                            auto_note = "Chi tiết danh mục máy trên KHLCNT: " + "; ".join(general_tasks_list[:3])
                        else:
                            auto_note = ""

                        header_text = f"{plan_no} {plan_name}".lower()
                        location_detail = (
                            "Tên KHLCNT"
                            if (any(k.lower() in header_text for k in kw_tags) or is_hospital_proj)
                            else "Danh mục Gói thầu / Tóm tắt công việc"
                        )

                        unique_matches[plan_no] = {
                            "code": plan_no,
                            "decision_no": decision_no,
                            "plan_id": plan_id,
                            "stage": stage_label,
                            "title": display_title,
                            "plan_name_full": plan_name,
                            "investor": investor,
                            "invest_total": display_price,
                            "public_date": public_date,
                            "expected_time": ", ".join(sorted(expected_times)) if expected_times else "Xem chi tiết",
                            "province": ", ".join(matched_province),
                            "matched_kw": final_kw_str,
                            "location": location_detail,
                            "auto_note": auto_note,
                            "bid_packages": bid_packages_summary[:25],
                            "detail_url": detail_url
                        }

                if max_block_len < current_page_size:
                    break
                if isinstance(max_total_elements, int) and fetched_per_block >= max_total_elements:
                    break

                page_num += 1
                time.sleep(0.35)
            except Exception as e:
                print(f"[CẢNH BÁO FETCH KHLCNT] Từ khóa '{kw}' (Trang {page_num}): {e}")
                break

        total_label = f" (tổng max/nhóm: {max_total_elements})" if isinstance(max_total_elements, int) else ""
        print(f"    + Đã tải {total_fetched_in_kw} bản ghi{total_label} (thuộc 14 tỉnh: {province_in_kw}) -> Khớp TBYT / Mua PACS / Công trình BV: {matched_in_kw} KHLCNT.")
        time.sleep(0.35)

    print(f"=== TỔNG KẾT QUÉT 14 TỈNH/TP: {len(unique_matches)} KẾ HOẠCH SẠCH KHỚP TIÊU CHÍ ===")
    return list(unique_matches.values())

# ==============================================================================
# 11. PHÂN TÍCH KHLCNT VỚI GEMINI AI & HẬU KIỂM LOẠI BỎ GÓI RÁC
# ==============================================================================
def analyze_batch_with_gemini(batch_items):
    global current_key_index, GEMINI_API_KEYS
    from google import genai
    from google.genai import types

    models_to_try = [
        PRIMARY_MODEL,
        FALLBACK_MODEL,
        "gemini-3.1-pro-preview"
    ]

    prompt = f"""
Bạn là chuyên gia thẩm định dữ liệu Kế hoạch lựa chọn nhà thầu (KHLCNT) Trang thiết bị Y tế, Hệ thống PACS (chỉ mua đứt/đầu tư, KHÔNG lấy gói thuê) và Dự án đầu tư xây dựng/nâng cấp Bệnh viện.
Dưới đây là danh sách các KHLCNT phát hành từ {START_TIME.strftime('%H:%M %d/%m/%Y')} đến {TODAY_STR} tại 14 tỉnh/thành phố phụ trách.

Yêu cầu BẮT BUỘC:
1. Giữ nguyên toàn bộ thuộc tính gốc của từng phần tử trong danh sách đầu vào.
2. Thêm trường boolean "is_medical_relevant":
   - Đặt true NẾU kế hoạch này liên quan đến mua sắm/sửa chữa Trang thiết bị y tế, MUA SẮM/ĐẦU TƯ hệ thống PACS (RIS/PACS), hoặc đầu tư xây mới/nâng cấp quy mô lớn Bệnh viện / Trung tâm y tế.
   - Đặt false nếu thuộc một trong các trường hợp sau:
     + Dự án ngoài ngành y tế (trường học, giao thông, công viên, UBND xã, công an, điện lực).
     + Gói THUÊ thiết bị y tế, THUÊ hệ thống RIS/PACS, dịch vụ chụp X-quang lưu động phục vụ khám sức khỏe.
     + Gói mua sắm thuần vật tư y tế tiêu hao (bơm kim tiêm, găng tay, bông băng gạc, phim X-quang, hóa chất xét nghiệm).
     + Gói sửa chữa nhỏ/duy tu cơ sở hạ tầng bệnh viện (ốp gạch, sơn tường, chống thấm, lợp mái).
3. Trong trường "reason", trình bày ngắn gọn, rõ ràng:
   - Nếu là dự án xây mới / nâng cấp / giai đoạn chuẩn bị đầu tư bệnh viện: nêu rõ quy mô công trình bệnh viện, các gói tư vấn/khảo sát/thiết kế đang triển khai và tiềm năng gói TBYT sắp tới.
   - Nếu là kế hoạch mua sắm/sửa chữa thiết bị y tế / hệ thống PACS: liệt kê rõ từng mã gói/lô con, tên thiết bị / hệ thống PACS chi tiết (lấy từ bid_packages / tóm tắt công việc), hình thức & thời gian dự kiến LCNT.
4. Nếu chưa có trường "stage", ghi "KHLCNT (Kế hoạch được duyệt)".
5. Thêm trường "bu_group": phân loại vào 1 trong các nhóm [Dự án Xây mới / Nâng cấp BV, Nội soi - Tiêu hóa, Phòng mổ - Gây mê, Hồi sức - Giường bệnh, Tiệt trùng, Khí y tế, Chẩn đoán hình ảnh, Đa chuyên khoa / TBYT tổng hợp].
6. Trả về đúng mảng JSON thuần túy (không kèm markdown).

Danh sách:
{json.dumps(batch_items, ensure_ascii=False, indent=2)}
"""

    retry_delays = [8, 15, 25]

    for model_name in models_to_try:
        key_idx = current_key_index
        while key_idx < len(GEMINI_API_KEYS):
            key = GEMINI_API_KEYS[key_idx]
            client = genai.Client(api_key=key)

            for attempt in range(len(retry_delays)):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    if response and response.text:
                        current_key_index = key_idx
                        parsed = json.loads(response.text.strip())
                        if isinstance(parsed, list):
                            return [x for x in parsed if x.get("is_medical_relevant", True) is not False]
                except Exception as e:
                    error_msg = str(e)
                    print(f"[CẢNH BÁO AI GOOGLE-GENAI ({model_name})] {error_msg}")

                    if "404" in error_msg or "not_found" in error_msg.lower():
                        key_idx = len(GEMINI_API_KEYS)
                        break

                    if any(k in error_msg.lower() for k in ["429", "resource_exhausted", "quota"]):
                        key_idx += 1
                        if key_idx < len(GEMINI_API_KEYS):
                            print(f"🔄 Đã đổi sang Gemini API Key thứ [{key_idx + 1}/{len(GEMINI_API_KEYS)}]")
                        break

                    if "503" in error_msg or "unavailable" in error_msg.lower():
                        if attempt < len(retry_delays) - 1:
                            wait_sec = retry_delays[attempt]
                            print(f"⏳ Model {model_name} đang quá tải (503), thử lại lần {attempt + 2} sau {wait_sec}s...")
                            time.sleep(wait_sec)
                            continue
                        key_idx = len(GEMINI_API_KEYS)
                        break

                    key_idx = len(GEMINI_API_KEYS)
                    break

    return batch_items

def analyze_with_gemini(candidates):
    if not candidates or not GEMINI_API_KEYS:
        return candidates

    batch_size = 15
    analyzed_results = []
    for i in range(0, len(candidates), batch_size):
        batch = candidates[i:i + batch_size]
        print(f"   Đang phân tích & thẩm định AI lô {i // batch_size + 1} ({len(batch)} KHLCNT)...")
        analyzed_batch = analyze_batch_with_gemini(batch)
        if isinstance(analyzed_batch, list):
            analyzed_results.extend(analyzed_batch)
        else:
            analyzed_results.extend(batch)
        time.sleep(1)
    return analyzed_results

# ==============================================================================
# 12. GỬI EMAIL BÁO CÁO KHLCNT ĐỊA BÀN PHỤ TRÁCH
# ==============================================================================
def send_email_report(matched_list):
    if not RECEIVER_EMAIL:
        print("[LỖI CẤU HÌNH] Chưa khai báo email nhận.")
        return

    receiver_list = [e.strip() for e in RECEIVER_EMAIL.replace("\n", "").split(",") if e.strip()]
    if not receiver_list:
        print("[LỖI CẤU HÌNH] Danh sách email nhận không hợp lệ.")
        return

    matched_list_sorted = sorted(matched_list, key=lambda x: x.get("province", ""))
    email_subject = f"[KHLCNT TBYT & Dự Án BV - Địa bàn phụ trách] Phát hiện {len(matched_list_sorted)} Kế hoạch LCNT ({START_TIME.strftime('%H:%M %d/%m')} - {TODAY_STR})"

    if not matched_list_sorted:
        html_body = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h3 style="color: #d9534f;">BÁO CÁO KẾ HOẠCH LỰA CHỌN NHÀ THẦU (KHLCNT)</h3>
            <p>Hệ thống đã quét dữ liệu từ <b>{START_TIME.strftime('%H:%M %d/%m/%Y')}</b> đến <b>{TODAY_STR}</b> nhưng chưa phát hiện KHLCNT mới khớp tiêu chí tại 14 địa phương theo dõi.</p>
        </div>
        """
    else:
        rows = ""
        for idx, item in enumerate(matched_list_sorted, 1):
            code = item.get("code", "")
            decision_no = item.get("decision_no", "")
            stage = item.get("stage", "KHLCNT")
            title = item.get("title", "")
            investor = item.get("investor", "")
            invest_total = item.get("invest_total", "Chưa cập nhật")
            public_date = item.get("public_date", "")
            expected_time = item.get("expected_time", "Xem chi tiết")
            province = item.get("province", "Theo dõi")
            location = item.get("location", "Danh mục Gói thầu chi tiết")
            auto_note = item.get("auto_note", "")
            reason = item.get("reason") or ("; ".join(item.get("bid_packages", [])[:5]) or "Khớp tiêu chí theo dõi y tế")
            if auto_note and auto_note not in reason:
                reason += f"<br><b style='color:#d9534f;'>Note: {auto_note}</b>"
            matched_kw = item.get("matched_kw", "")
            detail_url = item.get("detail_url") or f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index"

            code_block = f"<b>{code}</b><br><small style='color: #666;'>Đăng: {public_date}</small>"
            if decision_no:
                code_block += f"<br><small style='color: #0275d8;'>QĐ: {decision_no}</small>"
            code_block += f"<br><span style='font-size:11px; color: #2e7d32; font-weight:bold;'>[{stage}]</span>"

            rows += f"""
            <tr>
                <td align="center">{idx}</td>
                <td>{code_block}</td>
                <td><b>{title}</b></td>
                <td align="right"><span style="color: #d9534f; font-weight: bold;">{invest_total}</span></td>
                <td align="center"><span style="color: #2e7d32; font-weight: bold;">{expected_time}</span></td>
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
            <p>Đã lọc phát hiện <b>{len(matched_list_sorted)}</b> KHLCNT (TBYT, Mua PACS & Dự án xây mới/nâng cấp BV) thuộc 14 tỉnh/TP phụ trách:</p>
            <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border-color: #ddd;">
                <tr style="background-color: #e6f2ff;">
                    <th align="center">STT</th>
                    <th>Mã KHLCNT / Giai Đoạn</th>
                    <th>Tên Kế Hoạch / Gói Thầu</th>
                    <th align="center">Giá Trị</th>
                    <th align="center">Dự Kiến LCNT</th>
                    <th>Chủ Đầu Tư</th>
                    <th align="center">Tỉnh/TP</th>
                    <th align="center">Tiêu Chí Khớp</th>
                    <th align="center">Vị Trí Khớp</th>
                    <th>Phân Tích Chi Tiết</th>
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
            print(f"[SUCCESS] Đã gửi báo cáo KHLCNT ({len(matched_list_sorted)} kế hoạch) tới: {', '.join(receiver_list)}!")
        except Exception as e:
            print(f"[LỖI GỬI EMAIL SMTP] {e}")

def main():
    print(f"=== CRAWL & DEEP-SCAN KHLCNT 14 TỈNH/TP (Khung quét: {START_TIME.strftime('%H:%M %d/%m/%Y')} -> {TODAY_STR}) ===")
    driver = init_driver()
    try:
        matched_list = fetch_khlcnt_target_provinces(driver)
        existing_codes = get_existing_codes_from_gsheet()
        if existing_codes:
            matched_list = [m for m in matched_list if m.get("code") not in existing_codes]
            print(f"-> Sau khi đối chiếu chống trùng với Google Sheet: còn {len(matched_list)} KHLCNT mới hoàn toàn.")

        final_matches = analyze_with_gemini(matched_list)
        with open("latest_results.json", "w", encoding="utf-8") as f:
            json.dump(final_matches, f, ensure_ascii=False, indent=2)
        send_email_report(final_matches)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
