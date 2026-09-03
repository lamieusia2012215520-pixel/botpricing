import json
import re
import time

class MSCAccessDeniedException(Exception):
    pass

import subprocess
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException
from datetime import datetime, timedelta
from bot_cli import parse_date_offset_days, etd_within_max, max_etd_date, max_etd_date_only
from bot_runtime_utils import is_transient_webdriver_error, switch_to_live_window

DATE_OFFSET_DAYS = parse_date_offset_days(default=4)
MSC_MIN_SCHEDULE_ETD_DAYS = int(os.environ.get("MSC_MIN_SCHEDULE_ETD_DAYS", "6"))
from remark_rules import build_subject_remark

import builtins

# Lưu lại hàm print gốc của hệ thống
original_print = builtins.print

# Tạo hàm print mới tự động chèn thêm thời gian
def timestamped_print(*args, **kwargs):
    # Lấy giờ phút giây và mili-giây (VD: 15:30:45.123)
    current_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    original_print(f"[{current_time}]", *args, **kwargs)

# Ghi đè hàm print mặc định bằng hàm mới
builtins.print = timestamped_print

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
MSC_DEBUG_PORT    = 9530
DEBUGGER_ADDRESS  = f"localhost:{MSC_DEBUG_PORT}"
DRIVER_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "msedgedriver.exe")
EMAIL             = "celine@pio-logistics.vn"
PASSWORD          = "Vankiep*21"
MYMSC_QUOTE_URL   = "https://www.mymsc.com/myMSC/instantquote"
MYMSC_TIMEOUT_STR = "errorMessage=You%20have%20been%20timed%20out"

# ─────────────────────────────────────────────
# XPATH
# ─────────────────────────────────────────────
XPATH_MYMSC_EMAIL    = "/html/body/div[1]/div/div/div[2]/div/div/div/div[1]/div/div[1]/input"
XPATH_MYMSC_NEXT_BTN = "/html/body/div[1]/div/div/div[2]/div/div/div/div[1]/div/div[2]/button"
XPATH_B2C_PASSWORD   = "/html/body/div[2]/div[2]/section/div[1]/form/div[3]/div[2]/input"
XPATH_B2C_LOGIN_BTN  = "/html/body/div[2]/div[2]/section/div[1]/form/div[3]/div[2]/button"

# ─────────────────────────────────────────────
# COUNTRY CONSTANTS
# ─────────────────────────────────────────────
# POD thuộc các nước gần VN → chọn Port cho VNSGN POL
SEA_NEAR_COUNTRIES = [
    "CHINA", "CN", "TAIWAN", "TW", "JAPAN", "JP",
    "SOUTH KOREA", "KR", "NORTH KOREA", "KP", "KOREA",
    "PHILIPPINES", "PH", "MALAYSIA", "MY", "INDONESIA", "ID",
    "THAILAND", "TH", "CAMBODIA", "KH", "TIMOR LESTE", "TL",
    "MYANMAR", "MM", "SINGAPORE", "SG",
]

# FIX: Danh sách tên cảng/port thuộc vùng SEA/Đông Á để check trực tiếp theo POD name
# Dùng khi cột Country trong Excel bị bỏ trống hoặc điền sai
SEA_NEAR_PORTS = [
    # China
    "SHANGHAI", "NINGBO", "QINGDAO", "TIANJIN", "XINGANG", "DALIAN",
    "XIAMEN", "GUANGZHOU", "SHENZHEN", "YANTIAN", "SHEKOU", "NANSHA",
    "HONG KONG",
    # Japan
    "TOKYO", "YOKOHAMA", "NAGOYA", "OSAKA", "KOBE", "HAKATA", "MOJI",
    # Korea
    "BUSAN", "PUSAN", "INCHEON", "KWANGYANG",
    # Taiwan
    "KEELUNG", "KAOHSIUNG", "TAICHUNG",
    # Philippines
    "MANILA", "CEBU", "SUBIC", "DAVAO",
    # Malaysia
    "PORT KLANG", "PORT KELANG", "PENANG", "TANJUNG PELEPAS", "PASIR GUDANG",
    # Indonesia
    "JAKARTA", "SURABAYA", "BELAWAN", "SEMARANG",
    # Thailand
    "BANGKOK", "LAEM CHABANG", "LAT KRABANG",
    # Myanmar
    "YANGON", "RANGOON", "THILAWA",
    # Singapore
    "SINGAPORE",
    # Cambodia / Vietnam adjacent
    "SIHANOUKVILLE", "PHNOM PENH",
]

# POD thuộc nước khác → chọn Door cho VNSGN POL

CHINA_COUNTRIES = ["CHINA", "CN"]
CHINA_PORTS = [
    "SHANGHAI", "NINGBO", "QINGDAO", "TIANJIN", "XINGANG", "DALIAN",
    "XIAMEN", "GUANGZHOU", "SHENZHEN", "YANTIAN", "SHEKOU", "NANSHA",
    "HONG KONG"
]
JAPAN_COUNTRIES = ["JAPAN", "JP"]
EUROPE_COUNTRIES = [
    "DE", "FR", "NL", "BE", "ES", "IT", "GB", "UK", "PT", "SE",
    "DK", "FI", "PL", "NO", "AT", "CH", "GR", "IE", "CZ", "HU",
    "GERMANY", "FRANCE", "NETHERLANDS", "BELGIUM", "SPAIN", "ITALY",
    "UNITED KINGDOM", "PORTUGAL", "SWEDEN", "DENMARK", "FINLAND",
    "POLAND", "NORWAY", "AUSTRIA", "SWITZERLAND", "GREECE", "IRELAND",
    "EUROPE",
]
INDIA_COUNTRIES = ["INDIA", "IN"]

# POD đặc biệt luôn chọn Door thay vì Port
POD_DOOR_EXCEPTIONS = ["SAJED", "RIYADH"]

# Container types: (show_details_index, tên hiển thị)
CONT_TYPES = [
    (0, "20DV"),
    (1, "40DV"),
    (2, "40HC"),
]


# ═════════════════════════════════════════════
# PHẦN 1 – DRIVER + HELPERS CHUNG
# ═════════════════════════════════════════════

def is_port_in_use(port):
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def _wait_port(port, timeout=8):
    """FIX: Poll port thay vì sleep cứng."""
    import time as _t
    end = _t.time() + timeout
    while _t.time() < end:
        if is_port_in_use(port):
            return True
        _t.sleep(0.2)
    return False

def connect_driver() -> webdriver.Edge:
    import subprocess
    import time

    if not is_port_in_use(MSC_DEBUG_PORT):
        print("[HỆ THỐNG] Edge MSC chưa mở. Đang tự động khởi động...")
        try:
            subprocess.Popen([
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                f"--remote-debugging-port={MSC_DEBUG_PORT}",
                r"--user-data-dir=C:\edge_msc",
                "--disable-background-timer-throttling",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows"
            ])
            _wait_port(MSC_DEBUG_PORT, timeout=8)
        except: pass
    else:
        print("[HỆ THỐNG] Edge MSC đã mở sẵn. Bỏ qua lệnh khởi động trình duyệt.")
    
    options = Options()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)
    try:
        driver = webdriver.Edge(service=Service(executable_path=DRIVER_PATH), options=options)
        driver.set_page_load_timeout(30)
        try:
            driver.maximize_window()
        except Exception:
            try:
                info = driver.execute_cdp_cmd("Browser.getWindowForTarget", {})
                driver.execute_cdp_cmd("Browser.setWindowBounds", {
                    "windowId": info["windowId"],
                    "bounds": {"windowState": "maximized"}
                })
            except Exception:
                pass
        print(f"Connected to Edge: {DEBUGGER_ADDRESS}")
        return driver
    except Exception as e:
        print(f"Connection failed: {e}")
        exit()  

def wait_page_load(driver, timeout=15):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except TimeoutException:
        pass

def js_click(driver, xpath, label="", timeout=10) -> bool:
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].click();", el)
        print(f"Clicked: {label}")
        return True
    except TimeoutException:
        print(f"Timeout: {label}")
        return False
    except Exception as e:
        print(f"Click error {label}: {e}")
        return False

def fill_input(driver, xpath, text, label="", timeout=10) -> bool:
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        el.click()
        time.sleep(0.3)
        el.clear()
        el.send_keys(text)
        print(f"Filled {label}: {text}")
        return True
    except Exception as e:
        print(f"Fill error {label}: {e}")
        return False

def dismiss_cookies(driver):
    for by, sel in [
        (By.ID,           "onetrust-reject-all-handler"),
        (By.CSS_SELECTOR, "button#onetrust-reject-all-handler"),
    ]:
        try:
            el = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((by, sel))
            )
            driver.execute_script("arguments[0].click();", el)
            print("Cookies rejected")
            time.sleep(1)
            return
        except:
            continue


# ═════════════════════════════════════════════
# PHẦN 2 – LOGIN
# ═════════════════════════════════════════════

def login_and_go_to_instant_quote(driver) -> bool:
    print("Starting login process")
    driver.get(MYMSC_QUOTE_URL)

    try:
        email_el = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH_MYMSC_EMAIL)))
        driver.execute_script("""
            var email = arguments[0];
            email.value = arguments[1];
            email.dispatchEvent(new Event('input', {bubbles:true}));
            var btn = document.evaluate('""" + XPATH_MYMSC_NEXT_BTN + """', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (!btn) btn = document.querySelector('button[type="submit"]') || document.querySelector('.login-next-button');
            if (btn) btn.click();
        """, email_el, EMAIL)
        print("Email submitted")
    except Exception as e:
        print(f"Email step skipped: {e}")

    try:
        pass_el = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
        driver.execute_script("""
            var pass = arguments[0];
            pass.value = arguments[1];
            pass.dispatchEvent(new Event('input', {bubbles:true}));
            var btn = document.getElementById('next') || document.querySelector('button[type="submit"]');
            if(btn) btn.click();
        """, pass_el, PASSWORD)
        print("Password submitted")
    except:
        pass

    try:
        WebDriverWait(driver, 20).until(lambda d: "instantquote" in d.current_url and "errorMessage" not in d.current_url)
        print("Login successful")
        time.sleep(2)
        return True
    except:
        return "instantquote" in driver.current_url
    
# ═════════════════════════════════════════════
# PHẦN 3 – SHADOW DOM: NHẬP LIỆU POL / POD
# ═════════════════════════════════════════════

def _shadow_fill_and_select(driver, input_id: str, search_text: str, icon_type: str = "anchor") -> bool:
    listbox_id = f"{input_id}-listbox"
    res = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var input  = shadow.getElementById(arguments[0]);
        if (!input) return 'NOT_FOUND';
        input.focus();
        input.value = '';
        var nativeSet = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        nativeSet.call(input, arguments[1]);
        input.dispatchEvent(new Event('input',  { bubbles: true }));
        return 'OK';
    """, input_id, search_text)
    
    if res != 'OK': return False

    try:
        WebDriverWait(driver, 10).until(lambda d: d.execute_script("""
            var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
            var lb = shadow.getElementById(arguments[0]);
            return lb && lb.querySelectorAll('[role="option"]').length > 0;
        """, listbox_id))
    except:
        print(f"Dropdown not found: {input_id}")
        return False

    chosen = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var listbox = shadow.getElementById(arguments[0]);
        var options = listbox.querySelectorAll('[role="option"]');
        var iconClass = arguments[1] === 'anchor' ? '.icon-anchor' : '.icon-marker';
        for (var i = 0; i < options.length; i++) {
            if (options[i].querySelector(iconClass)) {
                options[i].click();
                return options[i].textContent.trim();
            }
        }
        if (options.length > 0) { options[0].click(); return options[0].textContent.trim(); }
        return null;
    """, listbox_id, icon_type)

    if chosen:
        print(f"Selected: {chosen}")
        return True
    return False


def _shadow_get_value(driver, input_id: str) -> str:
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var input  = shadow.getElementById(arguments[0]);
        return input ? input.value : '';
    """, input_id)

def normalize_location(name: str) -> str:
    """Sửa tên cảng theo quy tắc MSC"""
    n = name.upper().strip()
    if n == "HAI PHONG": return "HAIPHONG"
    if n == "DA NANG": return "VNDAD"
    if n == "CHENNAI": return "ENNORE"
    return n

def go_back_to_search(driver):
    if "iqsearchresult" not in driver.current_url:
        print("Already on search page")
        return
    driver.execute_script("""
        var host = document.querySelector('mymsc-instantquote-app');
        if (!host || !host.shadowRoot) return;
        var shadow = host.shadowRoot;
        var btn = shadow.querySelector('footer a p') || shadow.querySelector('a[href*="instantquote"]');
        if (btn) btn.click();
    """)
    print("Clicked back to search")
    time.sleep(2)

def click_tab_by_text(driver, tab_name: str) -> bool:
    """Hàm tổng quát để click vào tab dựa trên tên hiển thị (Schedule, Free Time...)"""
    success = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var tabs = shadow.querySelectorAll('button[role="tab"]');
        var targetText = arguments[0].toLowerCase();
        
        for (var i = 0; i < tabs.length; i++) {
            var tabText = tabs[i].textContent.trim().toLowerCase();
            if (tabText.includes(targetText)) {
                tabs[i].click();
                return true;
            }
        }
        return false;
    """, tab_name)
    if success:
        time.sleep(2.5 if tab_name.strip().lower() == "free time" else 1.2)
    return success

# Các hàm gọi lại cho dễ đọc code
def click_selected_charges_tab(driver):
    return click_tab_by_text(driver, "Selected Charges")

def click_quote_conditions_tab(driver):
    return click_tab_by_text(driver, "Quote Conditions")

def click_schedule_tab(driver):
    return click_tab_by_text(driver, "Schedule")

def click_freetime_tab(driver):
    return click_tab_by_text(driver, "Free Time")

def parse_transshipment_ports(driver) -> str:
    """Lấy thông tin cảng trung chuyển và format thành Port1 + Port2"""
    text = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        // Tìm div chứa chữ "Via :"
        var divs = shadow.querySelectorAll('.transportLocation .value, .MuiTypography-body2');
        for (var i = 0; i < divs.length; i++) {
            if (divs[i].textContent.includes('Via :')) {
                return divs[i].textContent.trim();
            }
        }
        return "";
    """)
    if not text: return ""
    
    # Xử lý chuỗi: "Via : DONG NAI (T/s 1), SINGAPORE (T/s 2)" 
    # 1. Bỏ "Via :"
    clean_text = text.replace("Via :", "").strip()
    # 2. Dùng regex tìm các tên cảng (bỏ phần (T/s X))
    ports = re.findall(r'([^,()]+)\s*\(T/s \d+\)', clean_text)
    if not ports:
        # Trường hợp chỉ có 1 cảng không có dấu phẩy
        ports = [re.sub(r'\s*\(T/s \d+\)', '', clean_text).strip()]
        
    return " + ".join([p.strip() for p in ports if p.strip()])


def parse_freetime_pod(driver) -> str:
    """Lay free time POD tu tab Free Time cua MyMSC."""
    data = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        function getPanelText() {
            var panel = shadow.querySelector('[data-test-id="tabPanelFreeTimeResult"]');
            if (panel) return panel.innerText || "";
            var nodes = shadow.querySelectorAll('[role="tabpanel"], .MuiTabPanel-root, div');
            for (var i = 0; i < nodes.length; i++) {
                var t = nodes[i].innerText || "";
                if (/Free\\s*Time|Demurrage|Detention|Combined/i.test(t)) return t;
            }
            return "";
        }
        function pick(regex, text) {
            var m = text.match(regex);
            return m ? m[1] : "";
        }
        var text = getPanelText().replace(/\\s+/g, " ").trim();
        if (!text) return {value:"", text:""};

        var combined =
            pick(/Import\\s+Combined[\\s\\S]*?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/Combined[\\s\\S]{0,80}?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/(\\d+)\\s*(?:Calendar\\s*)?Days?[\\s\\S]{0,60}Combined/i, text);
        if (combined) return {value: combined + " COMBINED", text:text};

        var dem =
            pick(/Import\\s+Demurrage[\\s\\S]*?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/Demurrage[\\s\\S]{0,80}?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/\\bDEM\\b[\\s\\S]{0,50}?(\\d+)\\s*(?:CD|Calendar\\s*Days?|Days?)/i, text) ||
            pick(/(\\d+)\\s*(?:CD|Calendar\\s*Days?|Days?)[\\s\\S]{0,50}\\bDEM\\b/i, text);
        var det =
            pick(/Import\\s+Detention[\\s\\S]*?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/Detention[\\s\\S]{0,80}?Free\\s*Days\\s*:?\\s*(\\d+)/i, text) ||
            pick(/\\bDET\\b[\\s\\S]{0,50}?(\\d+)\\s*(?:CD|Calendar\\s*Days?|Days?)/i, text) ||
            pick(/(\\d+)\\s*(?:CD|Calendar\\s*Days?|Days?)[\\s\\S]{0,50}\\bDET\\b/i, text);
        if (dem && det) return {value: dem + " DEM + " + det + " DET", text:text};
        if (dem) return {value: dem + " DEM", text:text};
        if (det) return {value: det + " DET", text:text};
        return {value:"", text:text};
    """)
    if isinstance(data, dict):
        value = data.get("value") or ""
        if value:
            print(f"Free Time POD: {value}")
            return value
        raw = (data.get("text") or "").strip()
        if raw:
            print(f"Free Time POD: not parsed. Text sample: {raw[:180]}")
    return ""

def get_pol_icon(pol: str, pod_country: str) -> str:
    """
    Xác định icon_type cho POL dựa vào quốc gia POD.
    VNSGN: gần VN → Port (anchor), xa VN → Door (marker)
    Các POL khác: luôn Port
    """
    pol_upper = pol.upper()
    pod_upper = pod_country.upper()
    if "VNSGN" in pol_upper or "HO CHI MINH" in pol_upper:
        if any(c in pod_upper for c in SEA_NEAR_COUNTRIES):
            return "anchor"   # Port
        else:
            return "marker"   # Door
    return "anchor"           # POL khác: luôn Port


def fill_pol_pod(driver, pol: str, pod: str, force_icon: str = None) -> bool:
    pol_norm = normalize_location(pol)
    pod_norm = normalize_location(pod)
    pod_icon = "marker" if any(x in pod_norm for x in POD_DOOR_EXCEPTIONS) else "anchor"
    
    print(f"Input: {pol_norm} ({force_icon}) -> {pod_norm} ({pod_icon})")
    
    try:
        WebDriverWait(driver, 20).until(lambda d: d.execute_script("""
            var host = document.querySelector('mymsc-instantquote-app');
            return host && host.shadowRoot && host.shadowRoot.getElementById('origin');
        """))
    except: return False

    if not _shadow_fill_and_select(driver, "destination", pod_norm, icon_type=pod_icon): return False
    if not _shadow_fill_and_select(driver, "origin", pol_norm, icon_type=force_icon): return False
    return True

# ═════════════════════════════════════════════
# PHẦN 4 – SEARCH RATES
# ═════════════════════════════════════════════

def click_search_rates(driver) -> bool:
    print("Clicking Search Rates")
    try:
        driver.execute_script("""
            var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
            var btn = shadow.getElementById('search-rate-button');
            if (btn) btn.click();
        """)
        return True
    except Exception as e:
        print(f"Search click error: {e}")
        return False

def wait_for_search_results(driver, timeout=60) -> str:
    print("Waiting for results...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                "iqsearchresult" in d.current_url
                or "You do not have access to use this functionality" in d.page_source
                or d.execute_script("""
                    var shadow = document.querySelector('mymsc-instantquote-app')?.shadowRoot;
                    if (!shadow) return false;
                    var h1 = shadow.querySelector('h1');
                    return h1 && h1.textContent.includes('No rates found');
                """)
            )
        )
    except TimeoutException:
        print("Timeout waiting for results")
        return 'timeout'

    if "You do not have access to use this functionality" in driver.page_source:
        raise MSCAccessDeniedException("Access denied during search")

    if "iqsearchresult" in driver.current_url:
        print("Results found")
        return 'has_results'
    else:
        print("No rates found")
        return 'no_results'


# ═════════════════════════════════════════════
# PHẦN 5 – ĐỌC GIÁ & LỊCH TÀU (LOGIC MỚI)
# ═════════════════════════════════════════════
def count_shipping_windows(driver) -> int:
    """Đếm số lượng Shipping Window (Card lớn ở trên)"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        return shadow.querySelectorAll('[data-test-id^="carousel-card-base"]').length;
    """) or 0

def click_shipping_window(driver, index):
    """Bấm vào Shipping Window thứ index"""
    driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var cards = shadow.querySelectorAll('[data-test-id^="carousel-card-base"]');
        if (arguments[0] < cards.length) {
            cards[arguments[0]].click();
        }
    """, index)
    time.sleep(2) # Chờ danh sách dòng giá bên dưới load lại

def count_rate_rows(driver) -> int:
    """Đếm tổng số dòng giá (rateCardBox) đang hiển thị"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        return shadow.querySelectorAll('.rateCardBox').length;
    """) or 0

def get_row_info(driver, row_idx):
    """Lấy thông tin Cảng (POL) và loại Cont của dòng thứ row_idx"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var rows = shadow.querySelectorAll('.rateCardBox');
        if (arguments[0] >= rows.length) return null;
        var row = rows[arguments[0]];
        
        var pol = "";
        var locs = row.querySelectorAll('.transportLocation');
        locs.forEach(l => {
            if (l.querySelector('.label')?.textContent.includes('Port of Load')) {
                pol = l.querySelector('.value')?.textContent.trim();
            }
        });
        
        var eq = row.querySelector('.equipmentDescription')?.textContent.trim() || "";
        return { pol: pol, equipment: eq };
    """, row_idx)

def click_show_details_on_row(driver, row_idx):
    """Bấm Show Details của dòng thứ row_idx"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var rows = shadow.querySelectorAll('.rateCardBox');
        if (arguments[0] >= rows.length) return false;
        var btn = rows[arguments[0]].querySelector('.showDetailsIcon');
        if (btn) { btn.click(); return true; }
        return false;
    """, row_idx)

def close_popup(driver):
    driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var btn = shadow.querySelector('.breakdown-modal-container .close-btn');
        if (btn) btn.click();
    """)
    time.sleep(0.5)

# --- Các hàm xử lý Schedule & Validity ---

def parse_msc_date(date_str):
    try:
        parts = date_str.split()
        if len(parts) == 2: date_str += f" {datetime.now().year}"
        return datetime.strptime(date_str, "%d %b %Y")
    except: return None

def format_etd_display(dates_list):
    if not dates_list: return ""
    dates_list.sort()
    
    def _fmt(d):
        return f"{d.day}-{d.strftime('%b')}"
    
    if len(dates_list) == 1: 
        return _fmt(dates_list[0])
        
    if len(dates_list) == 2:
        return f"{_fmt(dates_list[0])} & {_fmt(dates_list[1])}"
        
    months = set([d.month for d in dates_list])
    if len(months) == 1:
        days = [str(d.day) for d in dates_list]
        return f"{', '.join(days[:-1])}, {days[-1]}-{dates_list[-1].strftime('%b')}"
    else:
        return ", ".join([_fmt(d) for d in dates_list])

def get_validity_from_row(driver, row_idx):
    """Lấy ngày kết thúc Shipping Window từ chính dòng (Row) đó"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var rows = shadow.querySelectorAll('.rateCardBox');
        if (arguments[0] >= rows.length) return null;
        var row = rows[arguments[0]];
        
        // Tìm Shipping Window Date cuối cùng trong row
        var dates = row.querySelectorAll('.shipping-window-date');
        if (dates.length >= 2) {
            var day = dates[1].querySelector('.sail-day').textContent.trim();
            var month = dates[1].querySelector('.sail-month').textContent.trim();
            return day + "-" + month;
        }
        // Dự phòng nếu cấu trúc khác
        var fallback = row.querySelector('[data-test-id="OceanRateValidTo"]')?.textContent.trim();
        if (fallback) {
            var p = fallback.split(' ');
            return p[0] + "-" + p[1];
        }
        return null;
    """, row_idx)

def is_dong_nai_card(driver):
    """Kiểm tra lộ trình chi tiết của Card xem có phải đi từ DONG NAI không"""
    return driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        // 1. Quét các ô Transport Location
        var locations = shadow.querySelectorAll('.transportLocation');
        for (var i = 0; i < locations.length; i++) {
            var label = locations[i].querySelector('.label');
            var value = locations[i].querySelector('.value');
            if (label && label.textContent.includes('Port of Load')) {
                if (value && value.textContent.toUpperCase().includes('DONG NAI')) {
                    return true;
                }
            }
        }
        // 2. Quét nhanh toàn bộ text trên Card
        var activeBody = shadow.querySelector('.rateCardBody');
        if (activeBody && activeBody.innerText.toUpperCase().includes('DONG NAI')) {
            return true;
        }
        return false;
    """)


def process_schedule_logic(driver, ts_ports_str="DIRECT"):
    """Đọc bảng Schedule, lọc ETD >= DATE_OFFSET_DAYS và LOẠI BỎ TÀU TRÙNG LẶP"""
    schedule_raw = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var rows = shadow.querySelectorAll('table.schedule-table tbody tr');
        var results = [];
        rows.forEach(row => {
            var vessel = row.querySelector('th')?.textContent.trim() || "";
            var cells = row.querySelectorAll('td');
            if (cells.length >= 5) {
                results.push({
                    'vessel': vessel,
                    'voyage': cells[0].textContent.trim(),
                    'etd': cells[1].textContent.trim(),
                    'tt': cells[4].textContent.trim()
                });
            }
        });
        return results;
    """)
    if not schedule_raw: return None

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    min_etd = today + timedelta(days=DATE_OFFSET_DAYS)

    # 1. Lọc theo ngày ETD >= DATE_OFFSET_DAYS và LOẠI BỎ TRÙNG LẶP TUYỆT ĐỐI
    unique_vessels = []
    seen_entries = set() # Dùng set để đánh dấu những tàu đã lấy

    for item in schedule_raw:
        dt = parse_msc_date(item['etd'])
        if dt and dt >= min_etd:
            # Tạo một mã định danh duy nhất cho mỗi dòng tàu
            # Nếu Tên, Chuyến, Ngày chạy và Transit time giống hệt nhau -> Coi như là một
            entry_id = f"{item['vessel']}|{item['voyage']}|{item['etd']}|{item['tt']}"
            
            if entry_id not in seen_entries:
                item['dt'] = dt
                unique_vessels.append(item)
                seen_entries.add(entry_id)
            
    if not unique_vessels: return None

    # 2. Sắp xếp theo ETD tăng dần
    unique_vessels.sort(key=lambda x: x['dt'])

    # 3. Format chuỗi hiển thị cho cột O
    vessel_details = []
    ts_text = ts_ports_str if ts_ports_str else "DIRECT"
    for v in unique_vessels:
        etd_str = f"{v['dt'].day}-{v['dt'].strftime('%b')}"
        vessel_details.append(f"{v['vessel']} {v['voyage']} / ETD: {etd_str} / Transit time: {v['tt']} / Transshipment Port: {ts_text}")
    
    # 4. Lấy dữ liệu cho cột I và J (Lấy ETD đầu và cuối của danh sách đã lọc)
    unique_dates = sorted(list(set([v['dt'] for v in unique_vessels])))
    tts = [int(re.search(r'\d+', v['tt']).group()) for v in unique_vessels]
    
    tt_display = f"{min(tts)}-{max(tts)}" if min(tts) != max(tts) else f"{min(tts)}"
    
    return {
        'etd_text': format_etd_display(unique_dates[:3]),
        'tt_text': tt_display,
        'full_vessels': "\n".join(vessel_details)
    }

def process_schedule_logic(driver, ts_ports_str="DIRECT"):
    """
    MSC schedule rule:
    - Prefer ETDs >= today + MSC_MIN_SCHEDULE_ETD_DAYS (default 6 days).
    - If no such ETD exists but the schedule table still shows ETDs, take exactly
      one fallback ETD: the farthest ETD that is still below that 6-day threshold.
    """
    schedule_raw = driver.execute_script("""
        var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
        var rows = shadow.querySelectorAll('table.schedule-table tbody tr');
        var results = [];
        rows.forEach(row => {
            var vessel = row.querySelector('th')?.textContent.trim() || "";
            var cells = row.querySelectorAll('td');
            if (cells.length >= 5) {
                results.push({
                    'vessel': vessel,
                    'voyage': cells[0].textContent.trim(),
                    'etd': cells[1].textContent.trim(),
                    'tt': cells[4].textContent.trim()
                });
            }
        });
        return results;
    """)
    if not schedule_raw:
        return None

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    min_etd = today + timedelta(days=MSC_MIN_SCHEDULE_ETD_DAYS)

    deduped = []
    seen_entries = set()
    for item in schedule_raw:
        dt = parse_msc_date(item.get('etd', ''))
        if not dt or dt < today or not etd_within_max(dt):
            continue

        entry_id = f"{item.get('vessel','')}|{item.get('voyage','')}|{item.get('etd','')}|{item.get('tt','')}"
        if entry_id in seen_entries:
            continue

        item['dt'] = dt
        deduped.append(item)
        seen_entries.add(entry_id)

    if not deduped:
        return None

    selected_vessels = [v for v in deduped if v['dt'] >= min_etd]

    if not selected_vessels:
        fallback_candidates = [v for v in deduped if today <= v['dt'] < min_etd]
        if not fallback_candidates:
            return None
        selected_vessels = [max(fallback_candidates, key=lambda x: x['dt'])]
        print(
            f"MSC Schedule fallback: khong co ETD >= today+{MSC_MIN_SCHEDULE_ETD_DAYS}; "
            f"lay 1 ETD xa nhat duoi nguong: {selected_vessels[0]['dt'].strftime('%d-%b')}"
        )

    selected_vessels.sort(key=lambda x: x['dt'])

    vessel_details = []
    ts_text = ts_ports_str if ts_ports_str else "DIRECT"
    for v in selected_vessels:
        etd_str = f"{v['dt'].day}-{v['dt'].strftime('%b')}"
        vessel_details.append(
            f"{v['vessel']} {v['voyage']} / ETD: {etd_str} / "
            f"Transit time: {v['tt']} / Transshipment Port: {ts_text}"
        )

    unique_dates = sorted(set(v['dt'] for v in selected_vessels))
    tts = []
    for v in selected_vessels:
        m = re.search(r'\d+', v.get('tt', ''))
        if m:
            tts.append(int(m.group()))
    if not tts:
        return None

    tt_display = f"{min(tts)}-{max(tts)}" if min(tts) != max(tts) else f"{min(tts)}"

    return {
        'etd_text': format_etd_display(unique_dates[:3]),
        'tt_text': tt_display,
        'full_vessels': "\n".join(vessel_details)
    }


def parse_amount(amount_raw: str) -> tuple:
    """'350 USD' -> (350.0, 'USD')"""
    clean = " ".join(str(amount_raw or "").replace(",", "").split()).upper()
    m = re.search(r'\b([A-Z]{3})\b\s*([+-]?\d+(?:\.\d+)?)', clean)
    if m:
        return float(m.group(2)), m.group(1)
    m = re.search(r'([+-]?\d+(?:\.\d+)?)\s*\b([A-Z]{3})\b', clean)
    if m:
        return float(m.group(1)), m.group(2)
    return 0.0, 'UNKNOWN'

def parse_charges_popup(driver) -> dict | None:
    """Đọc dữ liệu từ popup chi tiết giá"""
    raw = driver.execute_script("""
        try {
            var shadow = document.querySelector('mymsc-instantquote-app').shadowRoot;
            var modal = shadow.querySelector('.breakdown-modal-container');
            if (!modal) return JSON.stringify({error: 'no_modal'});

            var result = { charges: [], total_raw: '', has_export_thc: false, equipment: '' };
            var totalEl = modal.querySelector('[data-test-id^="chargesTotal_"]');
            if (totalEl) result.total_raw = totalEl.textContent.trim();
            
            var rows = modal.querySelectorAll('.standard-charges-row');
            rows.forEach(function(row) {
                var catEl = row.querySelector(':scope > div > strong');
                var category = catEl ? catEl.textContent.trim() : '';
                var subRows = row.querySelectorAll('.standard-charges-subrow');
                subRows.forEach(function(sub) {
                    var divs = Array.from(sub.querySelectorAll(':scope > div'));
                    if (divs.length < 2) return;
                    var charge = divs[0].textContent.trim();
                    var amtEl = sub.querySelector('.amount-field strong');
                    var amount_raw = amtEl ? amtEl.textContent.trim() : '';
                    var payments = sub.querySelector('.allowedPaymentMethods')?.textContent.trim() || '';
                    var commEl = sub.querySelector('.commentsCondition');
                    var comments = commEl ? (commEl.querySelector('[title]')?.getAttribute('title') || commEl.textContent.trim()) : '';

                    if (category.toLowerCase().includes('export') && charge.toUpperCase().includes('THC')) {
                        result.has_export_thc = true;
                    }
                    result.charges.push({
                        category: category, charge: charge, level: divs[1].textContent.trim(),
                        amount_raw: amount_raw, payments: payments, comments: comments
                    });
                });
            });
            return JSON.stringify(result);
        } catch (e) { return JSON.stringify({error: e.message}); }
    """)
    try:
        data = json.loads(raw)
        return None if 'error' in data else data
    except: return None

import requests

def get_exchange_rates():
    try:
        response = requests.get("https://api.exchangerate-api.com/v4/latest/USD")
        data = response.json()
        print(f"Rates updated (VND: {data['rates']['VND']} | EUR: {data['rates']['EUR']} | CHF: {data['rates'].get('CHF')})")
        return data['rates']
    except Exception as e:
        print(f"Rate API failed ({e}). Using fallback.")
        return {"VND": 25450, "EUR": 0.92, "CHF": 0.89, "USD": 1}

EXCHANGE_RATES = get_exchange_rates()

def convert_to_usd(amount, currency):
    curr = currency.upper().strip()
    if curr == "USD": return amount
    if curr in EXCHANGE_RATES: return amount / EXCHANGE_RATES[curr]
    print(f"Rate not found for {curr}, using raw amount")
    return amount

def is_china_destination(pod_country: str = "", pod_name: str = "") -> bool:
    country_up = str(pod_country or "").upper().strip()
    pod_up = str(pod_name or "").upper().strip()
    return (
        country_up in [c.upper() for c in CHINA_COUNTRIES]
        or any(p in pod_up for p in CHINA_PORTS)
    )

def calculate_price_and_remark(charges_data: dict, pod_country: str = "", pod_name: str = "", equipment_hint: str = "") -> tuple:
    charges = charges_data.get('charges', [])
    has_export_thc = charges_data.get('has_export_thc', False)
    is_china = is_china_destination(pod_country, pod_name)
    
    total_usd = 0.0 # Chỉ dùng 1 biến tổng USD
    formula_parts = []
    china_thc_added = False
    
    ALWAYS_INC = ['FREIGHT CHARGE', 'FREIGHT SURCHARGES', 'PRE-CARRIAGE', 'PRE - CARRIAGE', 'ON CARRIAGE', 'ON-CARRIAGE', 'ON - CARRIAGE']
    REMARK_ONLY = ['[SEL]', 'SEAL FEE', '[DOC]', 'DOCUMENTATION FEE', '[TEL]', 'TELEX']

    for c in charges:
        if c['level'] != 'Per Equipment': continue
            
        cat_up = c['category'].upper().strip()
        charge_up = c['charge'].upper().strip()
        comm_up = c['comments'].upper().strip()
        is_export_thc = 'EXPORT' in cat_up and 'THC' in charge_up
        
        if any(x in charge_up for x in REMARK_ONLY): continue
        
        should_inc = False
        if is_china and is_export_thc:
            # China sell rates must be all-in O.THC.  Use the actual MSC charge
            # and currency shown for this equipment; do not substitute a fixed fee.
            should_inc = True
        elif any(x in cat_up for x in ALWAYS_INC):
            should_inc = True
        elif 'EXPORT' in cat_up or 'IMPORT' in cat_up:
            should_inc = (c['payments'].strip() == 'Prepaid') or ('SAME TERMS' in comm_up)
            
        if 'PREPAID TERMS OF PAYMENT ONLY' in comm_up:
            should_inc = True

        if should_inc:
            amt_raw, cur_raw = parse_amount(c['amount_raw'])
            # QUY ĐỔI SANG USD TẠI ĐÂY, gồm THC/L VND của MSC.
            amt_in_usd = convert_to_usd(amt_raw, cur_raw)
            total_usd += amt_in_usd
            formula_parts.append(round(amt_in_usd, 2))
            if is_export_thc:
                china_thc_added = True
                print(f"Included MSC origin THC: {amt_in_usd:.2f} USD ({amt_raw:g} {cur_raw})")

    if is_china and has_export_thc and not china_thc_added:
        print("WARNING: MSC showed a separate origin THC row but its amount could not be parsed")

    # Làm tròn giá tiền cho đẹp (ví dụ: 1824.56 -> 1825 hoặc giữ 2 số thập phân)
    final_totals = {"USD": round(total_usd, 2)}
    if formula_parts:
        expr_parts = []
        for part in formula_parts:
            expr_parts.append(str(int(part)) if float(part).is_integer() else str(part))
        final_totals["formula"] = "=" + "+".join(expr_parts)

    othc_included = bool(is_china or not has_export_thc)
    remark = build_subject_remark(othc_included=othc_included, country=pod_country, pod=pod_name)
    
    return final_totals, remark
    
    
# --- Hàm Scrape chính (Đã tích hợp Flow lọc) ---

def scrape_all_results(driver, pod_country: str = "", pod_name: str = "") -> dict | None:
    print("Starting scraping process")
    num_windows = count_shipping_windows(driver)
    if num_windows == 0: return None

    pod_check = pod_country.upper().strip()
    is_india = (pod_check == "INDIA" or pod_check == "IN")
    
    for w_idx in range(num_windows):
        print(f"Processing Window {w_idx + 1}/{num_windows}")
        click_shipping_window(driver, w_idx)
        
        num_rows = count_rate_rows(driver)
        current_window_data = None
        target_pol = "" 

        for i in range(num_rows):
            info = get_row_info(driver, i)
            if not info: continue
            
            pol_name = info['pol'].upper()
            equipment = info['equipment']
            
            if ("20'" in equipment or "40'" in equipment) and "DONG NAI" not in pol_name:
                print(f"Checking row {i}: {pol_name} | {equipment}")
                if click_show_details_on_row(driver, i):
                    if click_schedule_tab(driver):
                        click_quote_conditions_tab(driver)
                        ts_ports = parse_transshipment_ports(driver)
                        click_schedule_tab(driver)
                        sched = process_schedule_logic(driver, ts_ports_str=ts_ports or "DIRECT")
                        if not sched:
                            print(f"Window {w_idx+1}: Schedule invalid. Skipping entire window.")
                            close_popup(driver)
                            break 

                        print("Schedule OK. Fetching data...")

                        freetime_val = "14 COMBINED" if is_india else (click_freetime_tab(driver) and parse_freetime_pod(driver) or parse_freetime_pod(driver))
                        
                        close_popup(driver) 
                        click_show_details_on_row(driver, i)
                        charges_data = parse_charges_popup(driver)
                        close_popup(driver)

                        if charges_data:
                            target_pol = info['pol'] 
                            validity = get_validity_from_row(driver, i)
                            totals, remark = calculate_price_and_remark(charges_data, pod_country, pod_name, equipment)
                            
                            current_window_data = {
                                'etd': sched['etd_text'],
                                'tt_display': sched['tt_text'],
                                'full_vessels': sched['full_vessels'],
                                'ts_ports': ts_ports,
                                'validity': validity,
                                'freetime': freetime_val,
                                'containers': {("20DV" if "20'" in equipment else ("40HC" if "High Cube" in equipment else "40DV")): {'total': totals, 'remark': remark}}
                            }
                            print(f"Selected Window {w_idx+1} - POL: {target_pol}")
                            break 
                    else:
                        close_popup(driver)

        if current_window_data:
            for i in range(num_rows):
                info = get_row_info(driver, i)
                if info and info['pol'] == target_pol:
                    eq = info['equipment']
                    cont_key = "40DV" if "40'" in eq and "High Cube" not in eq else ("40HC" if "40'" in eq and "High Cube" in eq else "20DV")
                    
                    if cont_key and cont_key not in current_window_data['containers']:
                        if click_show_details_on_row(driver, i):
                            charges_data = parse_charges_popup(driver)
                            if charges_data:
                                totals, remark = calculate_price_and_remark(charges_data, pod_country, pod_name, eq)
                                current_window_data['containers'][cont_key] = {'total': totals, 'remark': remark}
                            close_popup(driver)
                            print(f"Fetched price: {cont_key}")
            
            if is_india:
                print("Applying India pricing rules")
                conts = current_window_data['containers']
                if "20DV" in conts: conts['20DV']['total']['USD'] = conts['20DV']['total'].get('USD', 0.0) + 50.0
                if "40DV" in conts:
                    conts['40DV']['total']['USD'] = conts['40DV']['total'].get('USD', 0.0) + 50.0
                    if "40HC" in conts:
                        conts['40HC']['total']['USD'] = conts['40DV']['total']['USD']
                        conts['40HC']['remark'] = conts['40DV']['remark']
                elif "40HC" in conts:
                    conts['40HC']['total']['USD'] = conts['40HC']['total'].get('USD', 0.0) + 50.0

            return current_window_data 

    return None


# ═════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════

def solve_msc_quote(driver, pol, pod, country):
    pol_up = pol.upper()
    is_vnsgn = "VNSGN" in pol_up or "HO CHI MINH" in pol_up
    # FIX: kiểm tra cả country column VÀ tên POD trực tiếp
    # → tránh bỏ sót khi cột Country trong Excel để trống hoặc điền sai
    is_near = (
        any(c in country.upper() for c in SEA_NEAR_COUNTRIES)
        or any(p in pod.upper() for p in SEA_NEAR_PORTS)
    )
    is_exception = any(x in pod.upper() for x in POD_DOOR_EXCEPTIONS)
    
    if is_vnsgn and (not is_near or is_exception):
        print("Mode: DOOR (MARKER) only")
        if fill_pol_pod(driver, pol, pod, force_icon="marker"):
            if click_search_rates(driver) and wait_for_search_results(driver) == 'has_results':
                return scrape_all_results(driver, country, pod)
        return None

    if is_vnsgn and is_near:
        res_port = None
        res_door = None

        print("Mode 1: PORT (ANCHOR)")
        if fill_pol_pod(driver, pol, pod, force_icon="anchor"):
            if click_search_rates(driver) and wait_for_search_results(driver) == 'has_results':
                res_port = scrape_all_results(driver, country, pod)
        
        go_back_to_search(driver)
        
        print("Mode 2: DOOR (MARKER)")
        if fill_pol_pod(driver, pol, pod, force_icon="marker"):
            if click_search_rates(driver) and wait_for_search_results(driver) == 'has_results':
                res_door = scrape_all_results(driver, country, pod)

        if res_port and res_door:
            p_price = res_port['containers'].get('20DV', {}).get('total', {}).get('USD', 9999)
            d_price = res_door['containers'].get('20DV', {}).get('total', {}).get('USD', 9999)
            if p_price < d_price: return res_port
            if d_price < p_price: return res_door
            return res_port if res_port.get('tt_num', 99) <= res_door.get('tt_num', 99) else res_door
        
        return res_port or res_door

    print("Mode: PORT (ANCHOR) only")
    if fill_pol_pod(driver, pol, pod, force_icon="anchor"):
        if click_search_rates(driver) and wait_for_search_results(driver) == 'has_results':
            return scrape_all_results(driver, country, pod)
    
    return None

import pandas as pd
import openpyxl

if __name__ == "__main__":
    driver = connect_driver()
    file_path = os.environ.get("EXCEL_PATH", "input_gia.xlsx")
    FILTER_POL = os.environ.get("FILTER_POL", "").strip().upper()
    FILTER_POD = os.environ.get("FILTER_POD", "").strip().upper()
    SINGLE_ROW = os.environ.get("SINGLE_ROW", "").strip()
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        print(f"Loaded Excel sheet: {sheet.title}")
    except Exception as e:
        print(f"Error opening {file_path}: {e}")
        exit()

    # Port mapping: Excel name → MSC search name
    MSC_PORT_MAPPING = {
        "TIANJIN": "XINGANG",
        "FOS SUR MER": "FOS-SUR-MER",
    }

    target_single_row = None
    if SINGLE_ROW:
        try:
            target_single_row = int(SINGLE_ROW)
            print(f"[SINGLE_ROW] Chỉ chạy dòng {target_single_row} theo lệnh từ main.py")
        except Exception:
            print(f"[SINGLE_ROW] Không hợp lệ: {SINGLE_ROW}")

    target_rows = []
    for row_idx in range(2, sheet.max_row + 1):
        if target_single_row is not None and row_idx != target_single_row:
            continue

        carrier = str(sheet[f"E{row_idx}"].value or "").strip().upper()
        if carrier != "MSC":
            continue

        country = str(sheet[f"B{row_idx}"].value or "").strip()
        if not country:
            country = os.environ.get("FILTER_COUNTRY", "").strip()
        pol     = str(sheet[f"C{row_idx}"].value or "").strip()
        pod     = str(sheet[f"D{row_idx}"].value or "").strip()

        if not pol or not pod:
            continue
        if FILTER_POL and pol.upper() != FILTER_POL:
            continue
        if FILTER_POD and pod.upper() != FILTER_POD:
            continue

        target_rows.append((row_idx, country, pol, pod))

    total_rows = len(target_rows)
    print(f"MSC có {total_rows} dòng cần check")

    for progress_idx, (row_idx, country, pol, pod) in enumerate(target_rows, start=1):

        # Áp dụng port mapping cho web search (giữ tên gốc trong Excel)
        pol = MSC_PORT_MAPPING.get(pol.upper(), pol)
        pod = MSC_PORT_MAPPING.get(pod.upper(), pod)

        print("-" * 40)
        print(
            f"[MSC {progress_idx}/{total_rows}] "
            f"Processing Excel row {row_idx}: {pol} -> {pod} ({country})"
        )

        transient_retries = 0
        while True:
            try:
                try:
                    driver.get(MYMSC_QUOTE_URL)
                except Exception as e:
                    print(f"Page load timeout/error: {e}")
                    try:
                        driver.execute_script("window.stop();")
                    except:
                        pass
                wait_page_load(driver)
                
                if "You do not have access to use this functionality" in driver.page_source:
                    raise MSCAccessDeniedException("Access denied on load")

                if "instantquote" not in driver.current_url or "errorMessage" in driver.current_url:
                    if not login_and_go_to_instant_quote(driver):
                        print(f"Login error at row {row_idx}")
                        raise Exception("Login error")

                final_data = solve_msc_quote(driver, pol, pod, country)

                if final_data:
                    conts = final_data['containers']
                    if "20DV" in conts:
                        sheet[f"F{row_idx}"].value = conts['20DV']['total'].get('formula') or conts['20DV']['total'].get('USD', "-")
                        sheet[f"M{row_idx}"].value = conts['20DV']['remark'] 
                    else:
                        sheet[f"F{row_idx}"].value = "-"

                    sheet[f"G{row_idx}"].value = (conts['40DV']['total'].get('formula') or conts['40DV']['total'].get('USD', "-")) if "40DV" in conts else "-"
                    sheet[f"H{row_idx}"].value = (conts['40HC']['total'].get('formula') or conts['40HC']['total'].get('USD', "-")) if "40HC" in conts else "-"

                    sheet[f"I{row_idx}"].value = final_data['etd']
                    sheet[f"J{row_idx}"].value = final_data['tt_display']
                    sheet[f"K{row_idx}"].value = final_data['validity']
                    sheet[f"N{row_idx}"].value = final_data['freetime']
                    
                    sheet[f"O{row_idx}"].value = final_data['full_vessels']
                    sheet[f"O{row_idx}"].alignment = openpyxl.styles.Alignment(wrapText=True)
                    sheet[f"P{row_idx}"].value = final_data['ts_ports']

                    print(f"[MSC {progress_idx}/{total_rows}] Saved Excel row {row_idx}")
                else:
                    sheet[f"F{row_idx}"].value = "-"
                    sheet[f"G{row_idx}"].value = "-"
                    sheet[f"H{row_idx}"].value = "-"
                    sheet[f"O{row_idx}"].value = "-"
                    sheet[f"P{row_idx}"].value = "-"
                    print(
                        f"[MSC {progress_idx}/{total_rows}] "
                        f"No rates for Excel row {row_idx}, filled '-'"
                    )
                
                break # Success or handled properly, break the retry loop

            except MSCAccessDeniedException as e:
                print(f"ACCESS DENIED ERROR: {e}. Logging out and restarting row {row_idx}...")
                driver.delete_all_cookies()
                driver.get(MYMSC_QUOTE_URL)
                time.sleep(3)
                continue # Retry the loop for the same row
                
            except Exception as e:
                print(f"[ERROR] Row {row_idx} thất bại: {e}")
                if is_transient_webdriver_error(e) and transient_retries < 1:
                    transient_retries += 1
                    print(
                        f"[MSC {progress_idx}/{total_rows}] Lỗi browser tạm thời -> "
                        f"khôi phục và retry cùng row ({transient_retries}/1)"
                    )
                    try:
                        driver.execute_script("window.stop();")
                    except Exception:
                        pass
                    try:
                        switch_to_live_window(driver)
                    except Exception:
                        try:
                            driver = connect_driver()
                        except Exception as reconnect_error:
                            print(f"[MSC] Reconnect thất bại: {reconnect_error}")
                    time.sleep(1)
                    continue
                sheet[f"F{row_idx}"].value = "-"
                sheet[f"G{row_idx}"].value = "-"
                sheet[f"H{row_idx}"].value = "-"
                sheet[f"O{row_idx}"].value = "-"
                sheet[f"P{row_idx}"].value = "-"
                print(
                    f"[MSC {progress_idx}/{total_rows}] "
                    f"Filled '-' for Excel row {row_idx} due to error"
                )
                # Thử reconnect driver nếu tab bị đóng
                try:
                    driver.current_url
                except:
                    print("Driver lost connection, trying to reconnect...")
                    try:
                        driver = connect_driver()
                    except:
                        print("Reconnect failed, skipping remaining rows")
                        break
                break # Break retry loop on normal error

        try:
            wb.save(file_path)
        except:
            print("Error saving Excel file")

    print("Completed all rows")
    try:
        driver.quit()
    except:
        pass
    sys.exit(0)
