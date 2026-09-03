import re

def fix():
    with open(r"c:\He_thong_Bot\bot_whl.py", "r", encoding="utf-8") as f:
        content = f.read()

    # Find the corrupted handle_captcha_if_needed
    start_str = 'def handle_captcha_if_needed():'
    end_str = '# ==================================================================================='
    
    if start_str in content and end_str in content:
        start_idx = content.find(start_str)
        # Find the end_str *after* start_idx
        end_idx = content.find(end_str, start_idx)
        
        correct_logic = """def handle_captcha_if_needed():
    \"\"\"Phát hiện captcha và hú còi báo động cho tới khi giải quyết xong.\"\"\"
    import winsound
    try:
        body_text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
    except Exception:
        body_text = ""
    
    captcha_indicators = ["captcha", "verify code", "verification code", "驗證碼", "验证码", "security check is required", "i am human"]
    has_captcha_text = any(t in body_text for t in captcha_indicators)
    
    captcha_imgs = driver.find_elements(By.XPATH,
        "//img[contains(@src,'captcha') or contains(@src,'CaptchaImg') or contains(@src,'verify') or contains(@src,'imgCode')] | //iframe[contains(@src, 'hcaptcha')]"
    )
    
    if has_captcha_text or captcha_imgs:
        log("   🛑 PHÁT HIỆN CAPTCHA — Vui lòng mở cửa sổ WHL giải quyết Captcha!")
        try:
            winsound.Beep(1500, 500) # Hú còi 1 lần báo hiệu
        except: pass
        log("   [WHL SILENT] Bot sẽ IM LẶNG chờ bạn giải quyết. Xong thì quay lại terminal và nhấn ENTER.")
        try:
            print("   [WHL SILENT] Nhấn ENTER sau khi WHL đã pass captcha...")
            input()
        except EOFError:
            log(f"   [WHL SILENT] Không có stdin, sleep 60s.")
            time.sleep(60)
        except KeyboardInterrupt:
            raise
            
        # Check xem Captcha đã qua chưa (thấy thẻ Select hoặc thẻ Bảng kết quả xuất hiện)
        try:
            if driver.find_elements(By.CSS_SELECTOR, "select#from_nation") or driver.find_elements(By.XPATH, "//table//tr"):
                log("   ✅ Đã qua Captcha, tự động cày tiếp...")
        except:
            pass
"""

        content = content[:start_idx] + correct_logic + "\n" + content[end_idx:]
        print("Fixed handle_captcha_if_needed block.")

    with open(r"c:\He_thong_Bot\bot_whl.py", "w", encoding="utf-8") as f:
        f.write(content)

fix()
