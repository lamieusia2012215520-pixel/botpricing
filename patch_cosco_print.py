import re

with open(r"c:\He_thong_Bot\bot_COSCO.py", "r", encoding="utf-8") as f:
    content = f.read()

# Add print statement at the end of the loop
target = """        # Ghi Free Time c\u1ed9t N (14)
        ft_value = free_time_dict.get(pod_country.upper(), "") 
        sheet.cell(row=row, column=14).value = ft_value
        
        wb.save(EXCEL_FILE)"""

replacement = """        # Ghi Free Time c\u1ed9t N (14)
        ft_value = free_time_dict.get(pod_country.upper(), "") 
        sheet.cell(row=row, column=14).value = ft_value
        
        wb.save(EXCEL_FILE)
        print(f"   [OK] Đã lưu dữ liệu dòng {row} thành công!")"""

if target in content:
    content = content.replace(target, replacement)
    with open(r"c:\He_thong_Bot\bot_COSCO.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Successfully patched bot_COSCO.py print save")
else:
    print("Failed to find target in bot_COSCO.py")
