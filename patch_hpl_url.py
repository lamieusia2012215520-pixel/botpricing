import re

with open(r"c:\He_thong_Bot\bot_HPL.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """def do_search(pol, pod, tab_idx):
    global cookie_dismissed
    t = tab_idx - 1  # 0-based index

    print(f"   [Tab{tab_idx}] \U0001f4cd Nhập {pol} \u2192 {pod}")
    check_security_block(tab_idx)"""

replacement = """def do_search(pol, pod, tab_idx):
    global cookie_dismissed
    t = tab_idx - 1  # 0-based index

    # YÊU CẦU: Luôn check URL trước khi nhập liệu
    target_url = "https://www.hapag-lloyd.com/solutions/new-quote/#/simple"
    if target_url not in driver.current_url:
        print(f"   [Tab{tab_idx}] 🔄 URL sai, điều hướng về: {target_url}")
        try:
            driver.get(target_url)
            time.sleep(3)
        except Exception as e:
            print(f"   [Tab{tab_idx}] ⚠️ Lỗi khi load URL: {e}")

    print(f"   [Tab{tab_idx}] 📍 Nhập {pol} → {pod}")
    check_security_block(tab_idx)"""

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_HPL.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_HPL.py with URL check!")
else:
    print("Failed to find target in bot_HPL.py!")
