import re

with open(r"c:\He_thong_Bot\main.py", "r", encoding="utf-8") as f:
    content = f.read()

old_code = """    name = BOT_DISPLAY.get(bot_file, bot_file.replace(".py", ""))
    if suffix:
        name = f"{name}_{suffix}"
    copy_path = excel_path.replace(".xlsx", f"_{name}.xlsx")"""

new_code = """    name = BOT_DISPLAY.get(bot_file, bot_file.replace(".py", ""))
    if suffix:
        name = f"{name}_{suffix}"
    import re
    safe_name = re.sub(r'[\\\\/*?:"<>|]', '-', name)
    copy_path = excel_path.replace(".xlsx", f"_{safe_name}.xlsx")"""

if old_code in content:
    content = content.replace(old_code, new_code)
    with open(r"c:\He_thong_Bot\main.py", "w", encoding="utf-8") as f:
        f.write(content)
    print("Success")
else:
    print("Failed to find target code block")
