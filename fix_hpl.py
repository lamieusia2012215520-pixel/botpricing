import re

def fix():
    with open(r"c:\He_thong_Bot\bot_hpl.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Clean up the duplicate set_date_hpl if it exists
    pattern_date = r"(window\.HTMLInputElement\.prototype, 'value'\)\.set;)(\s*\n\s*# ===+\n\s*# --- ĐẶT NGÀY ---\n\s*# ===+\ndef set_date_hpl\(\):.*?window\.HTMLInputElement\.prototype, 'value'\)\.set;)"
    if re.search(pattern_date, content, flags=re.DOTALL):
        content = re.sub(pattern_date, r"\1", content, count=1, flags=re.DOTALL)
        print("Fixed duplicate set_date_hpl.")

    # Let's completely rebuild do_search from definition to the print("   [Tab{tab_idx}] 🖱️ Đã Search")
    
    start_str = "def do_search(pol, pod, tab_idx):"
    end_str = 'print(f"   [Tab{tab_idx}] 🖱️ Đã Search")'
    
    if start_str in content and end_str in content:
        start_idx = content.find(start_str)
        end_idx = content.find(end_str) + len(end_str)
        
        correct_do_search = """def do_search(pol, pod, tab_idx):
    global cookie_dismissed
    t = tab_idx - 1  # 0-based index

    print(f"   [Tab{tab_idx}] 📍 Nhập {pol} → {pod}")
    check_security_block(tab_idx)
    pause_if_hpl_manual_check(f"Tab{tab_idx} truoc khi nhap form", tab_idx=tab_idx)
    hpl_raise_if_service_unavailable(tab_idx, "trước khi nhập form")

    # Dismiss cookie/notification banner — chỉ 1 lần
    if not cookie_dismissed:
        try:
            driver.execute_script(\"\"\"
                ['#onetrust-accept-btn-handler', '.onetrust-close-btn-handler',
                 '.hl-notification__close', '.hl-cookie-banner__close'
                ].forEach(function(s){
                    var el = document.querySelector(s); if (el) el.click();
                });
            \"\"\")
            rand_sleep(0.2, 0.3)
            cookie_dismissed = True
        except: pass

    # --- SMART POL: chỉ nhập lại nếu khác row trước ---
    if tab_last_pol[t] == pol:
        # Check xem field có còn giữ giá trị cũ không
        try:
            inp = driver.find_element(By.CSS_SELECTOR, 'input[data-testid="start-input"]')
            cur = driver.execute_script("return arguments[0].value;", inp) or ""
            if hpl_field_matches_port(cur, pol, POL_ALIASES.get(pol.upper())):
                print(f"        -> POL giữ nguyên: {cur.strip()}")
            else:
                select_port_hpl('input[data-testid="start-input"]', pol, aliases=POL_ALIASES.get(pol.upper()))
                rand_sleep(0.3, 0.5)
        except:
            select_port_hpl('input[data-testid="start-input"]', pol, aliases=POL_ALIASES.get(pol.upper()))
            rand_sleep(0.3, 0.5)
    else:
        select_port_hpl('input[data-testid="start-input"]', pol, aliases=POL_ALIASES.get(pol.upper()))
        rand_sleep(0.3, 0.5)
    tab_last_pol[t] = pol

    # --- SMART POD: chỉ nhập lại nếu khác row trước ---
    pod_upper   = pod.upper()
    pod_aliases = POD_ALIASES.get(pod_upper, [pod])
    if tab_last_pod[t] == pod:
        try:
            inp = driver.find_element(By.CSS_SELECTOR, 'input[data-testid="end-input"]')
            cur = driver.execute_script("return arguments[0].value;", inp) or ""
            first_alias = pod_aliases[0]
            if first_alias.upper() in cur.upper() or pod_upper in cur.upper():
                print(f"        -> POD giữ nguyên: {cur.strip()}")
            else:
                select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
        except:
            select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
    else:
        select_port_hpl('input[data-testid="end-input"]', pod_aliases[0], aliases=pod_aliases)
    tab_last_pod[t] = pod

    set_date_hpl()

    try:
        search_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[.//span[text()='Search']]")))
        human_move_and_click(search_btn)
        print(f"   [Tab{tab_idx}] 🖱️ Đã Search")
    except Exception as e:
        if hpl_manual_check_present():
            pause_if_hpl_manual_check(f"Captcha/Security chặn lúc click Search", tab_idx=tab_idx)
        raise e"""

        content = content[:start_idx] + correct_do_search + content[end_idx:]
        print("Fixed do_search block.")

    with open(r"c:\He_thong_Bot\bot_hpl.py", "w", encoding="utf-8") as f:
        f.write(content)

fix()
