import re

with open(r"c:\He_thong_Bot\bot_cma.py", "r", encoding="utf-8") as f:
    content = f.read()

target = """    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):"""

new_code = """    total_valid_rows = 0
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if target_row and i != target_row: continue
        pol_excel = str(row[2] or "").strip()
        pod       = str(row[3] or "").strip()
        carrier   = str(row[4] or "").strip().upper()
        if not pol_excel or not pod: continue
        if carrier not in {"CMA", "ANL", "CNC", "APL"}: continue
        if FILTER_POL and pol_excel.upper() != FILTER_POL: continue
        if FILTER_POD and pod.upper() != FILTER_POD: continue
        total_valid_rows += 1
    print(f"Tổng cộng có {total_valid_rows} dòng cần check.")

    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):"""

if target in content:
    content = content.replace(target, new_code)
    with open(r"c:\He_thong_Bot\bot_cma.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Success")
else:
    print("Failed to find target")
