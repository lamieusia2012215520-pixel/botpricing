import re

with open(r"c:\He_thong_Bot\bot_COSCO.py", "r", encoding="utf-8") as f:
    content = f.read()

target = "for row in danh_sach_dong:"

new_code = """total_valid_rows = 0
for row in danh_sach_dong:
    pol = str(sheet.cell(row=row, column=3).value or "").strip()
    pod = str(sheet.cell(row=row, column=4).value or "").strip()
    carrier = str(sheet.cell(row=row, column=5).value or "").strip().upper()
    if not pol or not pod or carrier != "COSCO":
        continue
    if FILTER_POL and pol.upper() != FILTER_POL:
        continue
    if FILTER_POD and pod.upper() != FILTER_POD:
        continue
    total_valid_rows += 1

print(f"Tổng cộng có {total_valid_rows} dòng cần check.")

for row in danh_sach_dong:"""

if target in content:
    content = content.replace(target, new_code, 1) # Only replace the first occurrence (which is the main loop)
    with open(r"c:\He_thong_Bot\bot_COSCO.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_COSCO.py with pre-scan loop")
else:
    print("Could not find target string in bot_COSCO.py!")
