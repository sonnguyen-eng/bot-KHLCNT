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
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & BỘ TỪ KHÓA LỌC CHÍNH XÁC Y TẾ / CÔNG TRÌNH BV
# ==============================================================================
RAW_GEMINI_KEYS = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_API_KEYS = [k.strip() for k in RAW_GEMINI_KEYS.split(",") if k.strip()]
current_key_index = 0

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "").strip() or SENDER_EMAIL

LAST_RUN_ENV = os.environ.get("LAST_RUN_TIMESTAMP", "").strip()
EVENT_NAME_ENV = os.environ.get("GITHUB_EVENT_NAME", "").strip()

# Nhóm 1: Từ khóa thiết bị y tế & giải pháp PACS đặc thù (khớp trực tiếp)
SPECIFIC_DEVICE_KEYWORDS = [
    # Nội soi - Phẫu thuật - Laser
    "nội soi", "ercp", "robot phẫu thuật", "tán sỏi laser", "máy tán sỏi",
    # Phòng mổ - Gây mê - Hồi sức - Giường
    "bàn mổ", "đèn mổ", "máy gây mê", "monitor theo dõi bệnh nhân",
    "máy theo dõi bệnh nhân", "bơm tiêm điện", "máy truyền dịch",
    "giường icu", "giường cấp cứu", "giường đa năng", "giường bệnh",
    "hồi sức cấp cứu", "hồi sức sơ sinh", "máy thở", "máy phá rung tim",
    # Khí y tế - Vận chuyển - Tiệt trùng
    "hệ thống khí y tế", "bồn oxy lỏng", "oxy trung tâm", "vận chuyển mẫu",
    "máy hấp tiệt trùng", "nhiệt độ thấp", "máy hấp ướt",
    # Sản khoa - Chẩn đoán hình ảnh - PACS - Chuyên khoa
    "bàn khám sản", "monitor sản khoa", "c-arm", "c arm", "x-quang",
    "máy siêu âm", "siêu âm màu", "siêu âm chẩn đoán", "thận nhân tạo",
    "pacs", "ris/pacs", "ris - pacs", "lưu trữ và truyền tải hình ảnh",
    # Gói mua sắm thiết bị y tế tổng thể
    "thiết bị y tế chuyên môn", "thiết bị y tế chuyên dùng",
    "thiết bị y tế cấp thiết", "mua sắm thiết bị chuyên môn",
    "mua sắm thiết bị y tế", "mua thiết bị y tế",
    "cung cấp, lắp đặt thiết bị y tế", "cung cấp lắp đặt thiết bị y tế",
    "trang thiết bị y tế"
]

# Nhóm 2: Từ khóa có điều kiện (CHỈ khớp khi đi kèm đơn vị/công trình ngành Y tế)
CONDITIONAL_KEYWORDS = [
    "thẩm định giá", "chẩn đoán hình ảnh", "phục hồi chức năng"
]

# Thực thể / Công trình thuộc ngành Y tế
MEDICAL_ENTITY_TERMS = [
    "bệnh viện", "trung tâm y tế", "viện tim", "ngành y tế", "sở y tế",
    "đại học y khoa", "đại học y dược", "khối khám bệnh", "khối bệnh nhiệt đới",
    "khoa khám bệnh", "khu khám bệnh", "khu hồi sức", "trung tâm sơ sinh",
    "trung tâm y khoa", "trạm y tế", "y tế khu vực"
]

# Hành động đầu tư xây mới / cải tạo / chuẩn bị dự án công trình BV
CONSTRUCTION_PROJECT_TERMS = [
    "xây dựng mới", "chuẩn bị đầu tư", "chuẩn bị dự án",
    "báo cáo nghiên cứu khả thi", "thiết kế bản vẽ thi công",
    "cải tạo, sửa chữa", "cải tạo sửa chữa", "nâng cấp",
    "xây dựng khối", "xây dựng khu", "mua sắm thiết bị"
]

# Từ khóa loại trừ các dự án ngoài ngành y tế hoặc kiểm xạ/khám sức khỏe định kỳ
EXCLUDE_NON_MEDICAL_PATTERNS = [
    "thẩm định giá đất", "thi hành bản án", "tuần tra, kiểm tra rừng",
    "dây chuyền kiểm định phương tiện", "cầu đi bộ", "tái định cư",
    "công viên", "nhà văn hóa", "trường tiểu học", "trường thcs", "trường thpt",
    "cao đẳng kỹ thuật", "trụ sở văn phòng", "hệ thống chiếu sáng",
    "mở rộng, nâng cấp tuyến đường", "kiểm định máy x-quang",
    "kiểm xạ an toàn", "liều kế cá nhân", "khám sức khỏe định kỳ"
]

# Cụm từ nhận diện gói thuần THUÊ dịch vụ phần mềm (không phải mua đứt/đầu tư PACS)
RENTAL_SOFTWARE_TERMS = [
    "thuê dịch vụ sử dụng phần mềm", "thuê dịch vụ phần mềm",
    "thuê hệ thống phần mềm", "thuê phần mềm", "cung cấp dịch vụ phần mềm"
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
    "Thành phố Đà Nẵng": ["đà nẵng", "quảng nam", "tam kỳ", "hội an"],
    "Tỉnh Quảng Ngãi": ["quảng ngãi", "kon tum"],
    "Tỉnh Đắk Lắk": ["đắk lắk", "đắc lắc", "tây nguyên", "buôn ma thuột", "phú yên", "tuy hòa"],
    "Tỉnh Gia Lai": ["gia lai", "pleiku", "bình định", "quy nhơn"],
    "Tỉnh Khánh Hòa": ["khánh hòa", "nha trang", "ninh thuận", "phan rang"],
    "Tỉnh Lâm Đồng": ["lâm đồng", "đà lạt", "bảo lộc", "đắk nông", "bình thuận", "phan thiết"],
    "Thành phố Hồ Chí Minh": [
        "hồ chí minh", "tp.hcm", "tphcm", "tp hcm", "sài gòn", "thủ đức",
        "bình dương", "bà rịa", "vũng tàu", "phú mỹ", "nhi đồng", "bình dân",
        "viện tim", "nhân ái", "đại học y dược", "ung bướu", "gò vấp",
        "bệnh nhiệt đới", "phạm ngọc thạch", "chấn thương", "lê văn thịnh", "thuận an"
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
        START_TIME = NOW_VN - timedelta(hours=24)
else:
    START_TIME = NOW_VN - timedelta(hours=24)

BASE_DOMAIN = "muasamcong.mof.gov.vn"
PRIMARY_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-2.5-flash-lite"

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
# 3. KHỞI TẠO CHROME HEADLESS
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
    if bid_time_str:
        return bid_time_str
    unit = pkg.get("bidStartUnit", "")
    year = pkg.get("bidStartYear", "")
    quarter = pkg.get("bidStartQuarter", "")
    month = pkg.get("bidStartMonth", "")
    if unit == "Q" and quarter and year:
        return f"Quý {quarter}/{year}"
    if unit == "M" and month and year:
        return f"Tháng {month}/{year}"
    if year:
        return f"Năm {year}"
    return "Chưa rõ thời gian"

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
# 9. BỘ LỌC 2 LỚP: LOẠI BỎ DỰ ÁN NGOÀI Y TẾ & NHẬN DIỆN TBYT / PACS / CÔNG TRÌNH BV
# ==============================================================================
def evaluate_medical_relevance(plan_name, investor, pkg_names_list, raw_str):
    """
    Đánh giá độ liên quan thực sự đến TBYT, mua sắm/đầu tư PACS hoặc Dự án xây mới/nâng cấp Bệnh viện:
    - Loại bỏ tuyệt đối các gói hạ tầng giao thông, trường học, công viên, trụ sở xã,
      hoặc gói thuần THUÊ dịch vụ phần mềm (HIS/EMR/PACS cho thuê định kỳ).
    - Giữ lại các gói MUA SẮM / ĐẦU TƯ / CUNG CẤP LẮP ĐẶT hệ thống PACS, RIS/PACS.
    """
    plan_lower = plan_name.lower()
    investor_lower = investor.lower()
    pkgs_lower = " ".join(pkg_names_list).lower()
    core_text = f"{plan_lower} {investor_lower} {pkgs_lower}"

    # 1. Chốt chặn loại trừ các dự án ngoài ngành y tế hoặc kiểm xạ/khám sức khỏe định kỳ
    if any(excl in core_text for excl in EXCLUDE_NON_MEDICAL_PATTERNS):
        return False, False, []

    if any(infra in plan_lower for infra in ["cầu đi bộ", "tuyến đường", "khu tái định cư", "hệ thống chiếu sáng", "trụ sở", "nhà văn hóa", "công viên"]):
        return False, False, []

    # 2. Loại bỏ gói thuần THUÊ dịch vụ phần mềm (không phải mua đứt/đầu tư PACS)
    is_software_rental = any(r in core_text for r in RENTAL_SOFTWARE_TERMS)
    is_purchase_action = any(p in core_text for p in ["mua sắm", "mua ", "đầu tư", "cung cấp", "lắp đặt", "trang bị"])
    if is_software_rental and not is_purchase_action:
        return False, False, []

    # Nếu chỉ là mua phần mềm bệnh án điện tử / HIS đơn thuần mà không có PACS hoặc TBYT thì bỏ qua
    if ("bệnh án điện tử" in core_text or "phần mềm quản lý bệnh viện" in core_text) and not any(
        p in core_text for p in ["pacs", "ris/pacs", "ris - pacs", "lưu trữ và truyền tải hình ảnh", "thiết bị y tế"]
    ):
        return False, False, []

    # 3. Kiểm tra có thuộc thực thể / công trình y tế hay không
    has_medical_entity = any(ent in f"{plan_lower} {investor_lower}" for ent in MEDICAL_ENTITY_TERMS)

    # 4. Kiểm tra từ khóa thiết bị y tế & PACS đặc thù
    search_scope = f"{plan_lower} {pkgs_lower}" if pkgs_lower.strip() else raw_str
    matched_specific = [k for k in SPECIFIC_DEVICE_KEYWORDS if k in search_scope]

    # 5. Kiểm tra từ khóa có điều kiện (CHỈ tính khi thuộc thực thể y tế và xuất hiện trong gói thầu/tên KH)
    matched_cond = []
    if has_medical_entity:
        for ck in CONDITIONAL_KEYWORDS:
            if ck in pkgs_lower or ck in plan_lower:
                matched_cond.append(ck)

    # 6. Kiểm tra dự án xây mới / cải tạo nâng cấp / chuẩn bị đầu tư công trình Bệnh viện
    is_hospital_proj = False
    hospital_proj_label = ""
    if has_medical_entity:
        matched_actions = [act for act in CONSTRUCTION_PROJECT_TERMS if act in plan_lower or act in pkgs_lower]
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
    print(f"-> Chế độ chạy: {EVENT_NAME_ENV or 'local'} | Mốc lịch sử gốc: {LAST_RUN_ENV or 'Mặc định 24h'}")
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

                    plan_no = str(item.get("planNo") or item.get("planNoStand") or "")
                    plan_id = str(item.get("id") or "")
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
                    bid_packages_summary = []
                    all_pkg_names = []
                    medical_pkg_names = []
                    medical_pkg_total = 0.0
                    expected_times = set()

                    if plan_id:
                        detail_json, detail_str = get_khlcnt_detail_values(driver, plan_id, token_val)
                        if detail_str:
                            raw_str += " " + detail_str
                        
                        plan_view = detail_json.get("bidPoBidpPlanProjectDetailView") or {}
                        decision_no = plan_view.get("decisionNo") or ""
                        
                        pkg_list = detail_json.get("bidpPlanDetailToProjectList") or []
                        if isinstance(pkg_list, list):
                            for pkg in pkg_list:
                                if not isinstance(pkg, dict):
                                    continue
                                b_no = str(pkg.get("bidNo") or pkg.get("code") or "").strip()
                                b_name = str(pkg.get("bidName") or "").replace("\n", " ").strip()
                                b_num_price = parse_numeric_amount(pkg.get("bidPrice"))
                                b_price = format_currency(b_num_price)
                                b_form_code = pkg.get("bidForm") or ""
                                b_form_text = BID_FORM_MAP.get(b_form_code, b_form_code)
                                b_time_expected = format_expected_bid_time(pkg)
                                
                                if b_time_expected != "Chưa rõ thời gian":
                                    expected_times.add(b_time_expected)
                                    
                                if b_name:
                                    all_pkg_names.append(b_name)
                                    prefix_code = f"{b_no} - " if b_no else ""
                                    bid_packages_summary.append(
                                        f"{prefix_code}{b_name} ({b_price} | {b_form_text} | Dự kiến: {b_time_expected})"
                                    )
                                    b_lower = b_name.lower()
                                    if any(med_w in b_lower for med_w in ["thiết bị y tế", "y tế", "nội soi", "bàn mổ", "đèn mổ", "gây mê", "hồi sức", "giường", "x-quang", "siêu âm", "máy thở", "tán sỏi", "khí y tế", "tiệt trùng", "monitor", "sơ sinh", "thận nhân tạo", "robot", "pacs", "lưu trữ và truyền tải hình ảnh"]):
                                        medical_pkg_names.append(f"{prefix_code}{b_name}")
                                        medical_pkg_total += b_num_price

                    is_relevant, is_hospital_proj, kw_tags = evaluate_medical_relevance(
                        plan_name, investor, all_pkg_names, raw_str
                    )

                    if is_relevant:
                        matched_in_kw += 1
                        final_kw_str = ", ".join(kw_tags)

                        is_early_prep = is_hospital_proj and any(
                            p in plan_name.lower()
                            for p in ["chuẩn bị đầu tư", "chuẩn bị dự án", "xây dựng mới", "báo cáo nghiên cứu khả thi", "thiết kế bản vẽ thi công", "cải tạo, sửa chữa", "cải tạo sửa chữa", "nâng cấp"]
                        )

                        if is_early_prep:
                            display_title = plan_name
                            display_price = invest_total if invest_total != "Chưa cập nhật" else format_currency(medical_pkg_total)
                            stage_label = "Tiền dự án (Chuẩn bị đầu tư / Xây mới BV)"
                        else:
                            display_title = " ; ".join(medical_pkg_names) if medical_pkg_names else plan_name
                            display_price = format_currency(medical_pkg_total) if medical_pkg_total > 0 else invest_total
                            stage_label = "KHLCNT (Kế hoạch được duyệt)"

                        has_generic_pkg = any(
                            any(g in p.lower() for g in ["cung cấp, lắp đặt thiết bị y tế", "cung cấp lắp đặt thiết bị y tế", "mua sắm thiết bị y tế"])
                            and not any(sp in p.lower() for sp in ["nội soi", "x-quang", "siêu âm", "máy thở", "bàn mổ", "tán sỏi", "sơ sinh", "thận nhân tạo", "pacs"])
                            for p in (medical_pkg_names or [plan_name])
                        )

                        if is_early_prep:
                            auto_note = "Giai đoạn chuẩn bị đầu tư / xây mới công trình BV - Theo dõi sát khâu lập danh mục TBYT"
                        elif has_generic_pkg:
                            auto_note = "Đang thực hiện xây danh mục, cấu hình thiết bị"
                        else:
                            auto_note = ""

                        header_text = f"{plan_no} {plan_name}".lower()
                        location_detail = (
                            "Tên KHLCNT"
                            if (any(k.lower() in header_text for k in kw_tags) or is_hospital_proj)
                            else "Danh mục Gói thầu chi tiết"
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
        print(f"    + Đã tải {total_fetched_in_kw} bản ghi{total_label} (thuộc 14 tỉnh: {province_in_kw}) -> Khớp TBYT / PACS / Công trình BV: {matched_in_kw} KHLCNT.")
        time.sleep(0.35)

    print(f"=== TỔNG KẾT QUÉT 14 TỈNH/TP: {len(unique_matches)} KẾ HOẠCH SẠCH KHỚP TIÊU CHÍ ===")
    return list(unique_matches.values())

# ==============================================================================
# 11. PHÂN TÍCH KHLCNT VỚI GEMINI AI & HẬU KIỂM LOẠI BỎ GÓI NGOÀI NGÀNH Y TẾ
# ==============================================================================
def analyze_batch_with_gemini(batch_items):
    global current_key_index, GEMINI_API_KEYS
    from google import genai
    from google.genai import types

    models_to_try = [
        PRIMARY_MODEL,
        FALLBACK_MODEL,
        "gemini-2.5-pro",
        "gemini-2.0-flash"
    ]

    prompt = f"""
Bạn là chuyên gia thẩm định dữ liệu Kế hoạch lựa chọn nhà thầu (KHLCNT) Trang thiết bị Y tế, Hệ thống PACS và Dự án đầu tư xây dựng Bệnh viện.
Dưới đây là danh sách các KHLCNT phát hành từ {START_TIME.strftime('%H:%M %d/%m/%Y')} đến {TODAY_STR} tại 14 tỉnh/thành phố phụ trách.

Yêu cầu BẮT BUỘC:
1. Giữ nguyên toàn bộ thuộc tính gốc của từng phần tử trong danh sách đầu vào.
2. Thêm trường boolean "is_medical_relevant":
   - Đặt true NẾU kế hoạch này liên quan đến mua sắm/sửa chữa Trang thiết bị y tế, vật tư y tế, MUA SẮM/ĐẦU TƯ hệ thống PACS (RIS/PACS), hoặc đầu tư xây mới/cải tạo nâng cấp Bệnh viện / Trung tâm y tế / Cơ sở khám chữa bệnh.
   - Đặt false nếu đây là dự án ngoài ngành y tế (trường học, giao thông, công viên, UBND xã, công an, điện lực, di tích) hoặc chỉ là thuê dịch vụ phần mềm hành chính/EMR không có mua sắm PACS/TBYT.
3. Trong trường "reason", trình bày ngắn gọn, rõ ràng:
   - Nếu là dự án xây mới / cải tạo / giai đoạn chuẩn bị đầu tư bệnh viện: nêu rõ quy mô công trình bệnh viện, các gói tư vấn/khảo sát/thiết kế đang triển khai và tiềm năng gói TBYT sắp tới.
   - Nếu là kế hoạch mua sắm thiết bị y tế / hệ thống PACS: liệt kê rõ từng mã gói/lô con (nếu có) và tên thiết bị / hệ thống PACS chính, hình thức & thời gian dự kiến LCNT.
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
            reason = item.get("reason") or (" ; ".join(item.get("bid_packages", [])[:5]) or "Khớp tiêu chí theo dõi y tế")
            if auto_note and auto_note not in reason:
                reason += f"<br><b style='color:#d9534f;'>Note: {auto_note}</b>"
            matched_kw = item.get("matched_kw", "")
            detail_url = item.get("detail_url") or f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index"
            
            code_block = f"<b>{code}</b><br><small style='color: #666;'>Đăng: {public_date}</small>"
            if decision_no:
                code_block += f"<br><small style='color: #0275d8;'>QĐ: {decision_no}</small>"
            code_block += f"<br><span style='font-size:11px; color:#2e7d32; font-weight:bold;'>[{stage}]</span>"

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
            <p>Đã lọc phát hiện <b>{len(matched_list_sorted)}</b> KHLCNT (TBYT, PACS & Dự án xây mới/nâng cấp BV) thuộc 14 tỉnh/TP phụ trách:</p>
            <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border-color: #ddd;">
                <tr style="background-color: #e6f2ff;">
                    <th align="center">STT</th>
                    <th>Mã KHLCNT / Giai Đoạn</th>
                    <th>Tên Kế Hoạch / Gói Thầu</th>
                    <th align="center">Giá Trị</th>
                    <th align="center">Dự Kiến LCNT</th>
                    <th>Chủ Đầu Tư</th>
                    <th align="center">Tỉnh / TP</th>
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
        final_matches = analyze_with_gemini(matched_list)
        with open("latest_results.json", "w", encoding="utf-8") as f:
            json.dump(final_matches, f, ensure_ascii=False, indent=2)
        send_email_report(final_matches)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
