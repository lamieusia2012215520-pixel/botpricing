import re

with open(r"c:\He_thong_Bot\bot_cma.py", "r", encoding="utf-8") as f:
    content = f.read()

target1 = 'print(f"   [OK] Đã lưu dữ liệu dòng {row_index} thành công!")'
replacement1 = 'print(f"   [INFO] Đã trích xuất giá dòng {row_index} xong!")'

target2 = '        print(f"   -> Nghỉ {CMA_ROW_SLEEP_MIN:g}-{CMA_ROW_SLEEP_MAX:g}s để chuyển dòng...")'
replacement2 = '        print(f"   [OK] Đã lưu dữ liệu dòng {i} thành công!")\n        print(f"   -> Nghỉ {CMA_ROW_SLEEP_MIN:g}-{CMA_ROW_SLEEP_MAX:g}s để chuyển dòng...")'

patched = False
if target1 in content:
    content = content.replace(target1, replacement1)
    patched = True
    print("Replaced target 1")
else:
    print("Could not find target 1")

if target2 in content:
    content = content.replace(target2, replacement2)
    patched = True
    print("Replaced target 2")
else:
    print("Could not find target 2")

if patched:
    with open(r"c:\He_thong_Bot\bot_cma.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_cma.py!")
