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
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & DANH SÁCH TỪ KHÓA Y TẾ
# ==============================================================================
RAW_GEMINI_KEYS = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_API_KEYS = [k.strip() for k in RAW_GEMINI_KEYS.split(",") if k.strip()]
current_key_index = 0

SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "").strip()
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", "").strip() or SENDER_EMAIL

LAST_RUN_ENV = os.environ.get("LAST_RUN_TIMESTAMP", "").strip()
EVENT_NAME_ENV = os.environ.get("GITHUB_EVENT_NAME", "").strip()

KEYWORDS = [
    "nội soi", "Bàn mổ", "đèn mổ", "hệ thống khí y tế", "vận chuyển mẫu", 
    "ERCP", "Monitor theo dõi bệnh nhân", "Máy hấp tiệt trùng", "Nhiệt độ thấp", 
    "máy gây mê", "Bơm tiêm điện", "máy truyền dịch", "bàn khám sản", 
    "Monitor sản khoa", "máy hấp ướt", "giường ICU", "giường cấp cứu", 
    "giường đa năng", "giường bệnh", "Bồn oxy lỏng", "Oxy trung tâm", "máy chạy thận", "HDF Online", "Ống soi"
]

BROAD_QUERIES = [
    "thiết bị y tế", "trang thiết bị y tế", "y tế", "bệnh viện",
    "trung tâm y tế", "sở y tế", "vật tư y tế", "mua sắm", "trang thiết bị"
]

# ==============================================================================
# 2. DANH SÁCH 14 TỈNH / THÀNH PHỐ MỤC TIÊU (BAO GỒM TÊN CŨ/MỚI & VIẾT TẮT)
# ==============================================================================
TARGET_LOCATIONS = [
    "Thành phố Đà Nẵng", "Tỉnh Quảng Ngãi", "Tỉnh Đắc Lắc", "Tỉnh Đắk Lắk", "Tỉnh Gia Lai", 
    "Tỉnh Khánh Hòa", "Tỉnh Lâm Đồng", "Thành phố Hồ Chí Minh", "Thành phố Đồng Nai", 
    "Tỉnh Tây Ninh", "Tỉnh Đồng Tháp", "Tỉnh Vĩnh Long", "Tỉnh An Giang", 
    "Thành phố Cần Thơ", "Tỉnh Cà Mau"
]

LOCATION_ALIASES = {
    "Thành phố Đà Nẵng": ["đà nẵng", "quảng nam", "tam kỳ", "hội an"],
    "Tỉnh Quảng Ngãi": ["quảng ngãi", "kon tum"],
    "Tỉnh Đắk Lắk": ["đắk lắk", "đắc lắc", "tây nguyên", "buôn ma thuột", "phú yên", "tuy hòa"],
    "Tỉnh Gia Lai": ["gia lai", "pleiku", "bình định", "quy nhơn"],
    "Tỉnh Khánh Hòa": ["khánh hòa", "nha trang", "ninh thuận", "phan rang"],
    "Tỉnh Lâm Đồng": ["lâm đồng", "đà lạt", "bảo lộc", "đắk nông", "bình thuận", "phan thiết"],
    "Thành phố Hồ Chí Minh": ["hồ chí minh", "tp.hcm", "tphcm", "tp hcm", "sài gòn", "thủ đức", "bình dương", "bà rịa", "vũng tàu"],
    "Thành phố Đồng Nai": ["đồng nai", "biên hòa", "long khánh", "bình phước", "đồng xoài"],
    "Tỉnh Tây Ninh": ["tây ninh", "long an", "tân an"],
    "Tỉnh Đồng Tháp": ["đồng tháp", "cao lãnh", "sa đéc", "tiền giang", "mỹ tho"],
    "Tỉnh Vĩnh Long": ["vĩnh long", "bến tre", "trà vinh"],
    "Tỉnh An Giang": ["an giang", "long xuyên", "châu đốc", "kiên giang", "rạch giá", "phú quốc"],
    "Thành phố Cần Thơ": ["cần thơ", "hậu giang", "sóc trăng"],
    "Tỉnh Cà Mau": ["cà mau", "bạc liêu"]
}

TZ_VN = timezone(timedelta(hours=7))
NOW_VN = datetime.now(TZ_VN)
TODAY_STR = NOW_VN.strftime("%H:%M %d/%m/%Y")

# Tính toán mốc START_TIME dựa trên lần chạy tự động (schedule) thành công gần nhất
# Lùi thêm 15 phút đệm an toàn để bù trừ độ trễ khởi động runner của GitHub Actions
if LAST_RUN_ENV:
    try:
        parsed_last_run = datetime.fromisoformat(LAST_RUN_ENV.replace("Z", "+00:00")).astimezone(TZ_VN)
        START_TIME = parsed_last_run - timedelta(minutes=15)
    except Exception:
        START_TIME = NOW_VN - timedelta(hours=24)
else:
    START_TIME = NOW_VN - timedelta(hours=24)

BASE_DOMAIN = "muasamcong.mof.gov.vn"
PRIMARY_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.8-flash-lite"

# Phân tầng giới hạn bản ghi quét (Server trả về tối đa 50 bản ghi/trang/nhóm trường):
# - Từ khóa sản phẩm cụ thể: tối đa 150 bản ghi / nhóm trường (3 trang x 50)
# - Từ khóa chung mở rộng: tối đa 500 bản ghi / nhóm trường (10 trang x 50)
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
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    return driver

# ==============================================================================
# 4. LỌC THỜI GIAN ĐĂNG & ĐỊNH DẠNG DỮ LIỆU
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

def format_currency(amount):
    if amount is None or amount == "":
        return "Chưa cập nhật"
    try:
        if isinstance(amount, list) and len(amount) > 0:
            amount = amount[0]
        val = int(float(amount))
        return f"{val:,} VNĐ".replace(",", ".")
    except Exception:
        return str(amount)

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
# 6. DEEP-SCAN CHI TIẾT KHLCNT (bid-po-bidp-plan-project-view/get-by-id)
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
# 9. TRÍCH XUẤT KHLCNT TRONG 14 TỈNH/TP VỚI PHÂN TẦNG GIỚI HẠN QUÉT (150 - 500)
# ==============================================================================
def fetch_khlcnt_target_provinces(driver):
    print(f"-> Chế độ chạy: {EVENT_NAME_ENV or 'local'} | Mốc lịch sử gốc: {LAST_RUN_ENV or 'Mặc định 24h'}")
    print(f"-> Kết nối Session Mua Sắm Công KHLCNT - 14 Tỉnh/TP (Khung quét có đệm: {START_TIME.strftime('%H:%M %d/%m/%Y')} -> {TODAY_STR})...")
    driver.get(f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index")
    time.sleep(9)

    token_val = get_latest_token_from_perf(driver)
    if not token_val:
        token_val = refresh_session_token(driver, "y tế")

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
    for kw in KEYWORDS:
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
                    invest_total = format_currency(item.get("investTotal"))
                    detail_url = build_khlcnt_detail_url(plan_id, plan_no)

                    raw_str = json.dumps(item, ensure_ascii=False).lower()

                    matched_province = check_location(item, raw_str)
                    if not matched_province:
                        continue

                    province_in_kw += 1

                    decision_no = ""
                    bid_packages_summary = []
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
                                b_name = str(pkg.get("bidName") or "").replace("\n", " ").strip()
                                b_price = format_currency(pkg.get("bidPrice"))
                                b_form_code = pkg.get("bidForm") or ""
                                b_form_text = BID_FORM_MAP.get(b_form_code, b_form_code)
                                b_time_expected = format_expected_bid_time(pkg)
                                b_capital = pkg.get("capitalDetail") or ""
                                
                                if b_time_expected != "Chưa rõ thời gian":
                                    expected_times.add(b_time_expected)
                                    
                                if b_name:
                                    bid_packages_summary.append(
                                        f"Gói: {b_name} | Giá: {b_price} | Hình thức: {b_form_text} | Dự kiến: {b_time_expected} | Nguồn vốn: {b_capital}"
                                    )

                    matched_kws = [k for k in KEYWORDS if k.lower() in raw_str]
                    if matched_kws:
                        matched_in_kw += 1
                        header_text = f"{plan_no} {plan_name}".lower()
                        location_detail = (
                            "Tên KHLCNT"
                            if any(k.lower() in header_text for k in matched_kws)
                            else "Danh mục Gói thầu chi tiết"
                        )
                        unique_matches[plan_no] = {
                            "code": plan_no,
                            "decision_no": decision_no,
                            "plan_id": plan_id,
                            "title": plan_name,
                            "investor": investor,
                            "invest_total": invest_total,
                            "public_date": public_date,
                            "expected_time": ", ".join(sorted(expected_times)) if expected_times else "Xem chi tiết",
                            "province": ", ".join(matched_province),
                            "matched_kw": ", ".join(matched_kws),
                            "location": location_detail,
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
        print(f"    + Đã tải {total_fetched_in_kw} bản ghi{total_label} (thuộc 14 tỉnh: {province_in_kw}) -> Khớp từ khóa y tế: {matched_in_kw} KHLCNT.")
        time.sleep(0.35)

    print(f"=== TỔNG KẾT QUÉT 14 TỈNH/TP: {len(unique_matches)} KHLCNT KHỚP TỪ KHÓA Y TẾ ===")
    return list(unique_matches.values())

# ==============================================================================
# 10. PHÂN TÍCH KHLCNT VỚI GEMINI AI (CHIA LÔ + RETRY 503 + FALLBACK MODEL)
# ==============================================================================
def analyze_batch_with_gemini(batch_items):
    global current_key_index, GEMINI_API_KEYS
    from google import genai
    from google.genai import types

    models_to_try = [PRIMARY_MODEL, FALLBACK_MODEL, "gemini-3.8-pro"]

    prompt = f"""
Bạn là chuyên gia phân tích dữ liệu Kế hoạch lựa chọn nhà thầu (KHLCNT) Trang thiết bị Y tế.
Dưới đây là danh sách các KHLCNT phát hành từ {START_TIME.strftime('%H:%M %d/%m/%Y')} đến {TODAY_STR} đã lọc theo 14 địa phương phụ trách và từ khóa y tế.

Yêu cầu:
1. Giữ nguyên toàn bộ danh sách đầu vào và mọi thuộc tính gốc.
2. Dựa vào danh sách gói thầu con (bid_packages), viết 1-2 câu tóm tắt ngắn gọn vào trường "reason": gói thầu liên quan thiết bị y tế gì, giá trị gói thầu, hình thức lựa chọn nhà thầu và thời gian dự kiến tổ chức.
3. Thêm trường "stage": ghi "KHLCNT (Kế hoạch được duyệt)".
4. Thêm trường "bu_group": phân loại vào 1 trong các nhóm [Nội soi - Tiêu hóa, Phòng mổ - Gây mê, Hồi sức - Giường bệnh, Tiệt trùng, Khí y tế, Vật tư & Khác].
5. Trả về đúng mảng JSON thuần túy (không kèm markdown).

Danh sách:
{json.dumps(batch_items, ensure_ascii=False, indent=2)}
"""

    for model_name in models_to_try:
        key_idx = current_key_index
        while key_idx < len(GEMINI_API_KEYS):
            key = GEMINI_API_KEYS[key_idx]
            client = genai.Client(api_key=key)

            for attempt in range(2):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=types.GenerateContentConfig(response_mime_type="application/json")
                    )
                    if response and response.text:
                        current_key_index = key_idx
                        return json.loads(response.text.strip())
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
                        if attempt == 0:
                            print(f"⏳ Model {model_name} đang quá tải, thử lại sau 10s...")
                            time.sleep(10)
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
        print(f"   Đang phân tích AI lô {i // batch_size + 1} ({len(batch)} KHLCNT)...")
        analyzed_batch = analyze_batch_with_gemini(batch)
        if isinstance(analyzed_batch, list) and len(analyzed_batch) == len(batch):
            analyzed_results.extend(analyzed_batch)
        else:
            analyzed_results.extend(batch)
        time.sleep(1)
    return analyzed_results

# ==============================================================================
# 11. GỬI EMAIL BÁO CÁO KHLCNT ĐỊA BÀN PHỤ TRÁCH
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
    email_subject = f"[KHLCNT TBYT - Địa bàn phụ trách] Phát hiện {len(matched_list_sorted)} Kế hoạch LCNT ({START_TIME.strftime('%H:%M %d/%m')} - {TODAY_STR})"

    if not matched_list_sorted:
        html_body = f"""
        <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h3 style="color: #d9534f;">BÁO CÁO KẾ HOẠCH LỰA CHỌN NHÀ THẦU (KHLCNT)</h3>
            <p>Hệ thống đã quét dữ liệu từ <b>{START_TIME.strftime('%H:%M %d/%m/%Y')}</b> đến <b>{TODAY_STR}</b> nhưng chưa phát hiện KHLCNT mới khớp từ khóa y tế tại 14 địa phương theo dõi.</p>
        </div>
        """
    else:
        rows = ""
        for idx, item in enumerate(matched_list_sorted, 1):
            code = item.get("code", "")
            decision_no = item.get("decision_no", "")
            title = item.get("title", "")
            investor = item.get("investor", "")
            invest_total = item.get("invest_total", "Chưa cập nhật")
            public_date = item.get("public_date", "")
            expected_time = item.get("expected_time", "Xem chi tiết")
            province = item.get("province", "Theo dõi")
            location = item.get("location", "Danh mục Gói thầu chi tiết")
            reason = item.get("reason") or "Khớp từ khóa theo dõi y tế trong KHLCNT"
            matched_kw = item.get("matched_kw", "")
            detail_url = item.get("detail_url") or f"https://{BASE_DOMAIN}/web/guest/contractor-selection?render=index"
            
            code_block = f"<b>{code}</b><br><small style='color: #666;'>Đăng: {public_date}</small>"
            if decision_no:
                code_block += f"<br><small style='color: #0275d8;'>QĐ: {decision_no}</small>"

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
            <p>Đã lọc phát hiện <b>{len(matched_list_sorted)}</b> KHLCNT Y tế thuộc địa bàn phụ trách:</p>
            <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border-color: #ddd;">
                <tr style="background-color: #e6f2ff;">
                    <th align="center">STT</th>
                    <th>Mã KHLCNT / Số QĐ</th>
                    <th>Tên KHLCNT</th>
                    <th align="center">Tổng Mức Đầu Tư</th>
                    <th align="center">Dự Kiến LCNT</th>
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
